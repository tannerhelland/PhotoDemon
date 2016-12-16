Attribute VB_Name = "Paintbrush"
'***************************************************************************
'Paintbrush tool interface
'Copyright 2016-2016 by Tanner Helland
'Created: 1/November/16
'Last updated: 15/December/16
'Last update: ongoing performance improvements
'
'To simplify the design of the primary canvas, it makes brush-related requests to this module.  This module
' then handles all the messy business of managing the actual background brush data.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Internally, we switch between different brush rendering engines depending on the current brush settings.
' The caller doesn't need to concern themselves with this; it's used only to determine internal rendering paths.
Private Enum BRUSH_ENGINE
    BE_GDIPlus = 0
    BE_PhotoDemon = 1
End Enum

#If False Then
    Private Const BE_GDIPlus = 0, BE_PhotoDemon = 1
#End If

Public Enum BRUSH_SOURCES
    BS_Color = 0
End Enum

#If False Then
    Private Const BS_Color = 0
#End If

Public Enum BRUSH_ATTRIBUTES
    BA_Source = 0
    BA_Size = 1
    BA_Opacity = 2
    BA_BlendMode = 3
    BA_AlphaMode = 4
    BA_Antialiasing = 5
    
    'Source-specific values can be stored here, as relevant
    BA_SourceColor = 1000
End Enum

#If False Then
    Private Const BA_Source = 0, BA_Size = 1, BA_Opacity = 2, BA_BlendMode = 3, BA_AlphaMode = 4, BA_Antialiasing = 5
    Private Const BA_SourceColor = 1000
#End If

'The current brush engine is stored here.  Note that this value is not correct until a call has been made to
' the CreateCurrentBrush() function; this function searches brush attributes and determines which brush engine
' to use.
Private m_BrushEngine As BRUSH_ENGINE
Private m_BrushOutlineImage As pdDIB, m_BrushOutlinePath As pd2DPath

'Brush preview quality.  At present, this is directly exposed on the paintbrush toolpanel.  This may change
' in the future, but for now, it's very helpful for testing.
Private m_BrushPreviewQuality As PD_PERFORMANCE_SETTING

'Brush resources, used only as necessary.  Check for null values before using.
Private m_GDIPPen As pd2DPen

'Brush attributes are stored in these variables
Private m_BrushSource As BRUSH_SOURCES
Private m_BrushSize As Single
Private m_BrushOpacity As Single
Private m_BrushBlendmode As LAYER_BLENDMODE
Private m_BrushAlphamode As LAYER_ALPHAMODE
Private m_BrushAntialiasing As PD_2D_Antialiasing

'Note that some brush attributes only exist for certain brush sources.
Private m_BrushSourceColor As Long

'If brush properties have changed since the last brush creation, this is set to FALSE.  We use this to optimize
' brush creation behavior.
Private m_BrushIsReady As Boolean
Private m_BrushCreatedAtLeastOnce As Boolean

'Current mouse/pen input values.  These are blindly relayed to us by the canvas, and it's up to us to perform any
' special tracking calculations.
Private m_MouseDown As Boolean
Private m_MouseX As Single, m_MouseY As Single

'As brush movements are relayed to us, we keep a running note of the modified area of the scratch layer.
' The compositor can use this information to only regenerate the compositor cache area that's changed since the
' last repaint event.  Note that the m_ModifiedRectF may be cleared between accesses, by design - you'll need to
' keep an eye on your usage of parameters in the GetModifiedUpdateRectF function.
'
'If you want the absolute modified area since the stroke began, you can use m_TotalModifiedRectF, which is not
' cleared until the current stroke is released.
Private m_UnionRectRequired As Boolean
Private m_ModifiedRectF As RECTF, m_TotalModifiedRectF As RECTF

'The number of mouse events in the *current* brush stroke.  This value is reset after every mouse release.
' The compositor uses this to know when to fully regenerate the paint cache from scratch.
Private m_NumOfMouseEvents As Long

'pd2D is used for certain paint features
Private m_Painter As pd2DPainter

'To improve responsiveness, we measure the time delta between viewport refreshes.  If painting is happening fast enough,
' we coalesce screen updates together, as they are (by far) the most time-consuming segment of paint rendering.
Private m_TimeSinceLastRender As Currency

Public Function GetBrushSource() As BRUSH_SOURCES
    GetBrushSource = m_BrushSource
End Function

Public Function GetBrushPreviewQuality() As PD_PERFORMANCE_SETTING
    GetBrushPreviewQuality = m_BrushPreviewQuality
End Function

Public Function GetBrushPreviewQuality_GDIPlus() As GP_InterpolationMode
    If (m_BrushPreviewQuality = PD_PERF_FASTEST) Then
        GetBrushPreviewQuality_GDIPlus = GP_IM_NearestNeighbor
    ElseIf (m_BrushPreviewQuality = PD_PERF_BESTQUALITY) Then
        GetBrushPreviewQuality_GDIPlus = GP_IM_HighQualityBicubic
    Else
        GetBrushPreviewQuality_GDIPlus = GP_IM_Bilinear
    End If
End Function

'Universal brush settings, applicable for all sources
Public Function GetBrushSize() As Single
    GetBrushSize = m_BrushSize
End Function

Public Function GetBrushOpacity() As Single
    GetBrushOpacity = m_BrushOpacity
End Function

Public Function GetBrushBlendMode() As LAYER_BLENDMODE
    GetBrushBlendMode = m_BrushBlendmode
End Function

Public Function GetBrushAlphaMode() As LAYER_ALPHAMODE
    GetBrushAlphaMode = m_BrushAlphamode
End Function

Public Function GetBrushAntialiasing() As PD_2D_Antialiasing
    GetBrushAntialiasing = m_BrushAntialiasing
End Function

'Brush settings that vary by source
Public Function GetBrushSourceColor() As Long
    GetBrushSourceColor = m_BrushSourceColor
End Function

Public Sub SetBrushSource(ByVal newSource As BRUSH_SOURCES)
    If (newSource <> m_BrushSource) Then
        m_BrushSource = newSource
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushPreviewQuality(ByVal newQuality As PD_PERFORMANCE_SETTING)
    If (newQuality <> m_BrushPreviewQuality) Then
        m_BrushPreviewQuality = newQuality
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSize(ByVal newSize As Single)
    If (newSize <> m_BrushSize) Then
        m_BrushSize = newSize
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushOpacity(Optional ByVal newOpacity As Single = 100#)
    If (newOpacity <> m_BrushOpacity) Then
        m_BrushOpacity = newOpacity
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushBlendMode(Optional ByVal newBlendMode As LAYER_BLENDMODE = BL_NORMAL)
    If (newBlendMode <> m_BrushBlendmode) Then
        m_BrushBlendmode = newBlendMode
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushAlphaMode(Optional ByVal newAlphaMode As LAYER_ALPHAMODE = LA_NORMAL)
    If (newAlphaMode <> m_BrushAlphamode) Then
        m_BrushAlphamode = newAlphaMode
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushAntialiasing(Optional ByVal newAntialiasing As PD_2D_Antialiasing = P2_AA_HighQuality)
    If (newAntialiasing <> m_BrushAntialiasing) Then
        m_BrushAntialiasing = newAntialiasing
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSourceColor(Optional ByVal newColor As Long = vbWhite)
    If (newColor <> m_BrushSourceColor) Then
        m_BrushSourceColor = newColor
        m_BrushIsReady = False
    End If
End Sub

Public Function GetBrushProperty(ByVal bProperty As BRUSH_ATTRIBUTES) As Variant
    
    Select Case bProperty
        Case BA_Source
            GetBrushProperty = GetBrushSource()
        Case BA_Size
            GetBrushProperty = GetBrushSize()
        Case BA_Opacity
            GetBrushProperty = GetBrushOpacity()
        Case BA_BlendMode
            GetBrushProperty = GetBrushBlendMode()
        Case BA_AlphaMode
            GetBrushProperty = GetBrushAlphaMode()
        Case BA_Antialiasing
            GetBrushProperty = GetBrushAntialiasing()
        Case BA_SourceColor
            GetBrushProperty = GetBrushSourceColor()
    End Select
    
End Function

Public Sub SetBrushProperty(ByVal bProperty As BRUSH_ATTRIBUTES, ByVal newPropValue As Variant)
    
    Select Case bProperty
        Case BA_Source
            SetBrushSource newPropValue
        Case BA_Size
            SetBrushSize newPropValue
        Case BA_Opacity
            SetBrushOpacity newPropValue
        Case BA_BlendMode
            SetBrushBlendMode newPropValue
        Case BA_AlphaMode
            SetBrushAlphaMode newPropValue
        Case BA_Antialiasing
            SetBrushAntialiasing newPropValue
        Case BA_SourceColor
            SetBrushSourceColor newPropValue
    End Select
    
End Sub

Public Sub CreateCurrentBrush(Optional ByVal alsoCreateBrushOutline As Boolean = True, Optional ByVal forceCreation As Boolean = False)
    
    If ((Not m_BrushIsReady) Or forceCreation Or (Not m_BrushCreatedAtLeastOnce)) Then
    
        'In the future we'll be implementing a full custom brush engine, but for this early testing phase,
        ' I'm restricting things to GDI+ for simplicity's sake.
        m_BrushEngine = BE_GDIPlus
        
        Select Case m_BrushEngine
            
            Case BE_GDIPlus
                'For now, create a circular pen at the current size
                If (m_GDIPPen Is Nothing) Then Set m_GDIPPen = New pd2DPen
                Drawing2D.QuickCreateSolidPen m_GDIPPen
        
        End Select
        
        'Whenever we create a new brush, we should also refresh the current brush outline
        If alsoCreateBrushOutline Then CreateCurrentBrushOutline
        
        m_BrushIsReady = True
        m_BrushCreatedAtLeastOnce = True
        
    End If
    
End Sub

'As part of rendering the current brush, we also need to render a brush outline onto the canvas at the current
' mouse location.  The specific outline technique used varies by brush engine.
Public Sub CreateCurrentBrushOutline()
    
    Select Case m_BrushEngine
    
        'If this is a GDI+ brush, outline creation is pretty easy.  Assume a circular brush and simply
        ' create a path at that same size.  (Note that circles are defined by radius, while brushes are
        ' defined by diameter - hence the "/ 2".)
        Case BE_GDIPlus
        
            Set m_BrushOutlinePath = New pd2DPath
            
            'Single-pixel brushes are treated as a square for cursor purposes.
            If (m_BrushSize > 0#) Then
                If (m_BrushSize = 1) Then
                    m_BrushOutlinePath.AddRectangle_Absolute -0.75, -0.75, 0.75, 0.75
                Else
                    m_BrushOutlinePath.AddCircle 0, 0, m_BrushSize / 2 + 0.5
                End If
            End If
            
    End Select

End Sub

'Notify the brush engine of the current mouse position.  Coordinates should always be in *image* coordinate space,
' not screen space.  (Translation between spaces will be handled internally.)
Public Sub NotifyBrushXY(ByVal mouseButtonDown As Boolean, ByVal srcX As Single, ByVal srcY As Single)
    
    Dim isFirstStroke As Boolean, isLastStroke As Boolean
    isFirstStroke = CBool((Not m_MouseDown) And mouseButtonDown)
    isLastStroke = CBool(m_MouseDown And (Not mouseButtonDown))
    
    'If this is a MouseDown operation, we need to prep the full paint engine.
    ' (TODO: initialize this elsewhere, so there's no "stutter" on first paint.)
    If isFirstStroke Then
        
        'Reset the number of mouse events
        m_NumOfMouseEvents = 1
        
        'Make sure the current scratch layer is properly initialized
        Tool_Support.InitializeToolsDependentOnImage
        pdImages(g_CurrentImage).ScratchLayer.SetLayerOpacity m_BrushOpacity
        pdImages(g_CurrentImage).ScratchLayer.SetLayerBlendMode m_BrushBlendmode
        pdImages(g_CurrentImage).ScratchLayer.SetLayerAlphaMode m_BrushAlphamode
        
        'Reset the "last mouse position" values to match the current ones
        m_MouseX = srcX
        m_MouseY = srcY
    
    Else
        m_NumOfMouseEvents = m_NumOfMouseEvents + 1
    End If
    
    'If the mouse button is down, perform painting between the old and new points.
    ' (All painting occurs in image coordinate space, and is applied to the current image's scratch layer.)
    If mouseButtonDown Then
    
        'Want to profile this function?  Use this line of code (and the matching report line at the bottom of the function).
        Dim startTime As Currency
        VB_Hacks.GetHighResTime startTime
        
        'Calculate new modification rects (which the compositor requires)
        UpdateModifiedRect srcX, srcY, isFirstStroke
        
        'Create required pd2D drawing tools (a painter and surface)
        Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDC cSurface, pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBDC, CBool(m_BrushAntialiasing = P2_AA_HighQuality)
        
        Dim cPen As pd2DPen
        Drawing2D.QuickCreateSolidPen cPen, m_BrushSize, m_BrushSourceColor, , P2_LJ_Round, P2_LC_Round
        
        'Render the line
        If isFirstStroke Then
            'GDI+ refuses to draw a line if the start and end points match; this isn't documented (as far as I know),
            ' but it may exist to provide backwards compatibility with GDI, which deliberately leaves the last point
            ' of a line unplotted, in case you are drawing multiple connected lines.
            m_Painter.DrawLineF cSurface, cPen, srcX, srcY, srcX - 0.01, srcY - 0.01
        Else
            m_Painter.DrawLineF cSurface, cPen, m_MouseX, m_MouseY, srcX, srcY
        End If
        
        Set cSurface = Nothing: Set cPen = Nothing
        
        pdImages(g_CurrentImage).ScratchLayer.NotifyOfDestructiveChanges
        
        Debug.Print "Paint tool render timing: " & Format(CStr(VB_Hacks.GetTimerDifferenceNow(startTime) * 1000), "0000.00") & " ms"
        
    End If
    
    'With all painting tasks complete, update all old state values to match the new state values
    m_MouseDown = mouseButtonDown
    m_MouseX = srcX
    m_MouseY = srcY
    
    'Unlike other drawing tools, the paintbrush engine controls viewport redraws.  This allows us to optimize behavior
    ' if we fall behind, and a long queue of drawing actions builds up.
    '
    '(Note that we only request manual redraws if the mouse is currently down; if the mouse *isn't* down, the canvas
    ' handles this for us.)
    If mouseButtonDown Then
    
        'If this is the first paint stroke, we always want to update the viewport to reflect that.
        Dim updateViewportNow As Boolean
        updateViewportNow = isFirstStroke
        
        'In the background, paint tool rendering is uncapped.  (60+ fps is achievable on most modern PCs, thankfully.)
        ' However, relaying those paint tool updates to the screen is a time-consuming process, as we have to composite
        ' the full image, apply color management, calculate zoom, and a whole bunch of other crap.  Because of this,
        ' it improves the user experience to run background paint calculations and on-screen viewport updates at
        ' different framerates, with an emphasis on making sure the *background* paint tool rendering gets top priority.
        If (Not updateViewportNow) Then
        
            'Limit viewport updates to 15 fps for now; we can revisit this in the future, as necessary
            updateViewportNow = CBool(VB_Hacks.GetTimerDifferenceNow(m_TimeSinceLastRender) * 1000 > 66#)
            
        End If
        
        'If a viewport update is required, composite the full layer stack prior to updating the screen
        If updateViewportNow Then
            VB_Hacks.GetHighResTime m_TimeSinceLastRender
            Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0), , , pdImages(g_CurrentImage).GetActiveLayerIndex
        
        'If not enough time has passed since the last redraw, simply update the cursor
        Else
            Viewport_Engine.Stage5_FlipBufferAndDrawUI pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        End If
        
    End If
    
End Sub

'Whenever we receive notifications of a new mouse (x, y) pair, you need to call this sub to calculate a new "affected area" rect.
Private Sub UpdateModifiedRect(ByVal newX As Single, ByVal newY As Single, ByVal isFirstStroke As Boolean)

    'Start by calculating the affected rect for just this stroke.
    Dim tmpRectF As RECTF
    If (newX < m_MouseX) Then
        tmpRectF.Left = newX
        tmpRectF.Width = m_MouseX - newX
    Else
        tmpRectF.Left = m_MouseX
        tmpRectF.Width = newX - m_MouseX
    End If
    
    If (newY < m_MouseY) Then
        tmpRectF.Top = newY
        tmpRectF.Height = m_MouseY - newY
    Else
        tmpRectF.Top = m_MouseY
        tmpRectF.Height = newY - m_MouseY
    End If
    
    'Inflate the rect calculation by the size of the current brush
    tmpRectF.Left = tmpRectF.Left - m_BrushSize / 2
    tmpRectF.Top = tmpRectF.Top - m_BrushSize / 2
    tmpRectF.Width = tmpRectF.Width + m_BrushSize
    tmpRectF.Height = tmpRectF.Height + m_BrushSize
    
    Dim tmpOldRectF As RECTF
    
    'If this is *not* the first modified rect calculation, union this rect with our previous update rect
    If m_UnionRectRequired And (Not isFirstStroke) Then
        tmpOldRectF = m_ModifiedRectF
        Math_Functions.UnionRectF m_ModifiedRectF, tmpRectF, tmpOldRectF
    Else
        m_UnionRectRequired = True
        m_ModifiedRectF = tmpRectF
    End If
    
    'Always calculate a total combined RectF, for use in the final merge step
    If isFirstStroke Then
        m_TotalModifiedRectF = tmpRectF
    Else
        tmpOldRectF = m_TotalModifiedRectF
        Math_Functions.UnionRectF m_TotalModifiedRectF, tmpRectF, tmpOldRectF
    End If
    
End Sub

'Return the area of the image modified by the current stroke.  By default, the running modified rect is erased after a call to
' this function, but this behavior can be toggled by resetRectAfter.  Also, if you want to get the full modified rect since this
' paint stroke began, you can set the GetModifiedRectSinceStrokeBegan parameter to TRUE.  Note that when
' GetModifiedRectSinceStrokeBegan is TRUE, the resetRectAfter parameter is ignored.
Public Function GetModifiedUpdateRectF(Optional ByVal resetRectAfter As Boolean = True, Optional ByVal GetModifiedRectSinceStrokeBegan As Boolean = False) As RECTF
    If GetModifiedRectSinceStrokeBegan Then
        GetModifiedUpdateRectF = m_TotalModifiedRectF
    Else
        GetModifiedUpdateRectF = m_ModifiedRectF
        If resetRectAfter Then m_UnionRectRequired = False
    End If
End Function

Public Function GetNumOfStrokes() As Long
    GetNumOfStrokes = m_NumOfMouseEvents
End Function

'Want to commit your current brush work?  Call this function to make the brush results permanent.
Public Sub CommitBrushResults()
    
    'Reset the current mouse event counter
    m_NumOfMouseEvents = 0
    
    'Committing brush results is actually pretty easy!
    
    'First, if the layer beneath the paint stroke is a raster layer, we simply want to merge the scratch
    ' layer onto it.
    If pdImages(g_CurrentImage).GetActiveLayer.IsLayerRaster Then
        
        Dim tmpRectF As RECTF
        tmpRectF = m_TotalModifiedRectF
        
        'Clip the modified rect to the paint layer's bounds, as necessary
        With tmpRectF
            If (.Left < 0) Then .Left = 0
            If (.Top < 0) Then .Top = 0
            If (.Width > pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBWidth) Then .Width = pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBWidth
            If (.Height > pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBHeight) Then .Height = pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBHeight
        End With
        
        Dim bottomLayerFullSize As Boolean
        With pdImages(g_CurrentImage).GetActiveLayer
            bottomLayerFullSize = CBool((.GetLayerOffsetX = 0) And (.GetLayerOffsetY = 0) And (.layerDIB.GetDIBWidth = pdImages(g_CurrentImage).Width) And (.layerDIB.GetDIBHeight = pdImages(g_CurrentImage).Height))
        End With
        
        pdImages(g_CurrentImage).MergeTwoLayers pdImages(g_CurrentImage).ScratchLayer, pdImages(g_CurrentImage).GetActiveLayer, bottomLayerFullSize, True, VarPtr(tmpRectF)
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, pdImages(g_CurrentImage).GetActiveLayerIndex
        
        'Ask the central processor to create Undo/Redo data for us
        Processor.Process "Paint stroke", , , UNDO_LAYER, g_CurrentTool
        
        'Reset the scratch layer
        pdImages(g_CurrentImage).ScratchLayer.layerDIB.ResetDIB 0
    
    'If the layer beneath this one is *not* a raster layer, let's add the stroke as a new layer, instead.
    Else
    
        Dim newLayerID As Long
        newLayerID = pdImages(g_CurrentImage).CreateBlankLayer(pdImages(g_CurrentImage).GetActiveLayerIndex)
        
        'Point the new layer index at our scratch layer
        pdImages(g_CurrentImage).PointLayerAtNewObject newLayerID, pdImages(g_CurrentImage).ScratchLayer
        pdImages(g_CurrentImage).GetLayerByID(newLayerID).SetLayerName g_Language.TranslateMessage("Paint layer")
        Set pdImages(g_CurrentImage).ScratchLayer = Nothing
        
        'Activate the new layer
        pdImages(g_CurrentImage).SetActiveLayerByID newLayerID
        
        'Notify the parent image of the new layer
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_IMAGE_VECTORSAFE
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
        'Ask the central processor to create Undo/Redo data for us
        Processor.Process "Paint stroke", , , UNDO_IMAGE_VECTORSAFE, g_CurrentTool
        
        'Create a new scratch layer
        Tool_Support.InitializeToolsDependentOnImage
        
    End If
    
End Sub

'Render the current brush outline to the canvas, using the stored mouse coordinates as the brush's position
Public Sub RenderBrushOutline(ByRef targetCanvas As pdCanvas)
    
    'If a brush outline doesn't exist, create one now
    If (Not m_BrushIsReady) Then CreateCurrentBrush True
    
    'Start by creating a transformation from the image space to the canvas space
    Dim canvasMatrix As pd2DTransform
    Drawing.GetTransformFromImageToCanvas canvasMatrix, targetCanvas, pdImages(g_CurrentImage), m_MouseX, m_MouseY
    
    'We also want to pinpoint the precise cursor position
    Dim cursX As Double, cursY As Double
    Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, pdImages(g_CurrentImage), m_MouseX, m_MouseY, cursX, cursY
    
    'If the on-screen brush size is above a certain threshold, we'll paint a full brush outline.
    ' If it's too small, we'll only paint a cross in the current brush position.
    Dim onScreenSize As Double
    onScreenSize = Drawing.ConvertImageSizeToCanvasSize(m_BrushSize, pdImages(g_CurrentImage))
    
    Dim brushTooSmall As Boolean
    If (onScreenSize < 7#) Then brushTooSmall = True
    
    'Create a pair of UI pens
    Dim innerPen As pd2DPen, outerPen As pd2DPen
    Drawing2D.QuickCreatePairOfUIPens outerPen, innerPen
    
    'Create other required pd2D drawing tools (a surface)
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, targetCanvas.hDC, True
    
    'Paint a target cursor - but *only* if the mouse is not currently down!
    Dim crossLength As Single, outerCrossBorder As Single
    crossLength = 3#
    outerCrossBorder = 0.5
    
    If (Not m_MouseDown) Then
        m_Painter.DrawLineF cSurface, outerPen, cursX, cursY - crossLength - outerCrossBorder, cursX, cursY + crossLength + outerCrossBorder
        m_Painter.DrawLineF cSurface, outerPen, cursX - crossLength - outerCrossBorder, cursY, cursX + crossLength + outerCrossBorder, cursY
        m_Painter.DrawLineF cSurface, innerPen, cursX, cursY - crossLength, cursX, cursY + crossLength
        m_Painter.DrawLineF cSurface, innerPen, cursX - crossLength, cursY, cursX + crossLength, cursY
    End If
    
    'If size allows, render a transformed brush outline onto the canvas as well
    If (Not brushTooSmall) Then
        
        'Get a copy of the current brush outline, transformed into position
        Dim copyOfBrushOutline As pd2DPath
        Set copyOfBrushOutline = New pd2DPath
        
        copyOfBrushOutline.CloneExistingPath m_BrushOutlinePath
        copyOfBrushOutline.ApplyTransformation canvasMatrix
    
        m_Painter.DrawPath cSurface, outerPen, copyOfBrushOutline
        m_Painter.DrawPath cSurface, innerPen, copyOfBrushOutline
    End If
    
    Set cSurface = Nothing
    Set innerPen = Nothing: Set outerPen = Nothing
    
End Sub

'A brush is considered active if the mouse state is currently DOWN, or if it is up but we are still rendering a
' previous stroke.
Public Function IsBrushActive() As Boolean
    IsBrushActive = m_MouseDown
End Function

'Any specialized initialization tasks can be handled here.  This function is called early in the PD load process.
Public Sub InitializeBrushEngine()
    m_BrushPreviewQuality = PD_PERF_BALANCED
    m_BrushAntialiasing = P2_AA_HighQuality
    Drawing2D.QuickCreatePainter m_Painter
End Sub

'Before PD closes, you *must* call this function!  It will free any lingering brush resources (which are cached
' for performance reasons).
Public Sub FreeBrushResources()
    Set m_GDIPPen = Nothing
    Set m_BrushOutlineImage = Nothing
    Set m_BrushOutlinePath = Nothing
    Set m_Painter = Nothing
End Sub
