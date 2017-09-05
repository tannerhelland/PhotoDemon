Attribute VB_Name = "FillTool"
'***************************************************************************
'PhotoDemon Bucket Fill Manager
'Copyright 2017-2017 by Tanner Helland
'Created: 30/August/17
'Last updated: 04/September/17
'Last update: continued work on initial build
'
'This module interfaces between the bucket fill UI and pdFloodFill backend.  Look in the relevant tool panel
' form for more details on how the UI relays relevant fill data here.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'All flood fill behavior is handled by a pdFloodFill instance.  (A matching pd2D path instance holds the
' fill outline, which gives us great flexibility in how the fill area is rendered.)
Private m_FloodFill As pdFloodFill, m_FillOutline As pd2DPath

'Fill operations may require a special copy of the current relevant image data.  (This may be a fully composited
' image copy, or a null-padded version of the current layer.)  We cache this locally to improve fill performance,
' and we use the notification timestamp from the parent image to determine when it's time to update our local copy.
Private m_FillImage As pdDIB, m_FillImageTimestamp As Currency

'Fill rendering is handled via pd2D.  A default painter class is created automatically, so don't worry about
' instantiating your own.
Private m_Painter As pd2DPainter

'Current mouse/pen input values.  These are blindly relayed to us by the canvas, and it's up to us to perform any
' special tracking calculations.
Private m_MouseDown As Boolean
Private m_MouseX As Single, m_MouseY As Single

'Most fill "properties" are forwarded to a pdFloodFill instance, but rendering-specific properties are actually
' cached internally.  (pdFloodFill just generates a flood outline - we have to do the actual filling.)
Private m_FillBrush As String, m_FillBlendMode As PD_BlendMode, m_FillAlphaMode As PD_AlphaMode
Private m_FillSampleMerged As Boolean

'Before attempting to set flood fill properties, call this sub to ensure the m_FloodFill object exists.
' (It returns TRUE if m_FloodFill exists.)
Private Function EnsureFillerExists() As Boolean
    If (m_FloodFill Is Nothing) Then Set m_FloodFill = New pdFloodFill
    EnsureFillerExists = True
End Function

Public Sub NotifyMouseXY(ByVal mouseButtonDown As Boolean, ByVal srcX As Single, ByVal srcY As Single, ByRef srcCanvas As pdCanvas)
    
    Dim oldMouseState As Boolean
    oldMouseState = m_MouseDown
    
    'Update all internal trackers.  (Note that the passed positions are in *image* coordinates, not *screen* coordinates.)
    m_MouseX = srcX
    m_MouseY = srcY
    m_MouseDown = mouseButtonDown
    
    'If this is an initial click, apply the fill.
    If mouseButtonDown And (Not oldMouseState) Then
        
        'Start by grabbing (or producing) the source DIB required for the fill.  (If the user wants us to
        ' sample all layers, we need to generate a composite image.)
        Dim fillImageRefreshRequired: fillImageRefreshRequired = False
        fillImageRefreshRequired = (m_FillImage Is Nothing)
        If (Not fillImageRefreshRequired) Then fillImageRefreshRequired = (m_FillImageTimestamp <> pdImages(g_CurrentImage).GetTimeOfLastChange())
        
        If fillImageRefreshRequired Then
            
            'A new image copy is required.  As much as possible, still try to minimize the work we do here
            If (m_FillImage Is Nothing) Then Set m_FillImage = New pdDIB
            If (m_FillImage.GetDIBWidth <> pdImages(g_CurrentImage).Width) Or (m_FillImage.GetDIBHeight <> pdImages(g_CurrentImage).Height) Then
                m_FillImage.CreateBlank pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, 32, 0, 0
            Else
                m_FillImage.ResetDIB 0
            End If
            
            'Note the timestamp; this may allow us to skip subsequent copy requests
            m_FillImageTimestamp = pdImages(g_CurrentImage).GetTimeOfLastChange()
            
            'Merged image data is far more cumbersome to generate
            If m_FillSampleMerged Then
                pdImages(g_CurrentImage).GetCompositedImage m_FillImage
                
            Else
                
                Dim tmpLayer As pdLayer
                Set tmpLayer = New pdLayer
                tmpLayer.CopyExistingLayer pdImages(g_CurrentImage).GetActiveLayer
                tmpLayer.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
                Set m_FillImage = tmpLayer.layerDIB
                Set tmpLayer = Nothing
                
            End If
            
        End If
        
        'm_FloodFill should exist here, because we need to create it to set any relevant properties!
        If (m_FloodFill Is Nothing) Then Debug.Print "WARNING!  m_FloodFill doesn't exist!"
        
        'Set the initial flood point
        m_FloodFill.SetInitialPoint srcX, srcY
        
        'Apply the flood fill; the result we want is not the returned image, but the returned path outline
        If (m_FillOutline Is Nothing) Then Set m_FillOutline = New pd2DPath Else m_FillOutline.ResetPath
        m_FloodFill.InitiateFloodFill m_FillImage, pdImages(g_CurrentImage).ScratchLayer.layerDIB, m_FillOutline
        
        'Create a brush using the passed brush settings
        Dim tmpBrush As pd2DBrush
        Set tmpBrush = New pd2DBrush
        tmpBrush.SetBrushPropertiesFromXML m_FillBrush
        
        'Gradient brushes require us to set a gradient rect; calculate that now, using the boundary rect of the fill outline
        If (tmpBrush.GetBrushMode = P2_BM_Gradient) Then
            Dim FillRect As RECTF
            FillRect = m_FillOutline.GetPathBoundariesF
            tmpBrush.SetBoundaryRect FillRect
        End If
        
        'Erase the scratch layer, because we are going to render our own fill result using the user's specified brush
        pdImages(g_CurrentImage).ScratchLayer.layerDIB.ResetDIB 0
        
        'Render the final fill
        Dim tmpSurface As pd2DSurface
        Set tmpSurface = New pd2DSurface
        tmpSurface.WrapSurfaceAroundPDDIB pdImages(g_CurrentImage).ScratchLayer.layerDIB
        m_Painter.FillPath tmpSurface, tmpBrush, m_FillOutline
        
        'Free all finished pd2D objects
        Set tmpSurface = Nothing
        Set tmpBrush = Nothing
        
        'Relay the current blend/alpha modes to the scratch layer
        pdImages(g_CurrentImage).ScratchLayer.SetLayerBlendMode m_FillBlendMode
        pdImages(g_CurrentImage).ScratchLayer.SetLayerAlphaMode m_FillAlphaMode
        
        'Commit the fill results
        FillTool.CommitFillResults
        
    End If
    
End Sub


'Want to commit your current fill work?  Call this function to make any pending fill results permanent.
Public Sub CommitFillResults()
    
    'Committing fill results is actually pretty easy!
    
    'Start by grabbing the boundaries of the fill area, and clipping it to the active layer's bounds, as necessary
    Dim tmpRectF As RECTF
    tmpRectF = m_FillOutline.GetPathBoundariesF
    
    With tmpRectF
        If (.Left < 0) Then .Left = 0
        If (.Top < 0) Then .Top = 0
        If (.Width > pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBWidth) Then .Width = pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBWidth
        If (.Height > pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBHeight) Then .Height = pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBHeight
    End With
    
    'First, if the layer being filled is a raster layer, we simply want to merge the scratch layer onto it.
    If pdImages(g_CurrentImage).GetActiveLayer.IsLayerRaster Then
        
        Dim bottomLayerFullSize As Boolean
        With pdImages(g_CurrentImage).GetActiveLayer
            bottomLayerFullSize = CBool((.GetLayerOffsetX = 0) And (.GetLayerOffsetY = 0) And (.layerDIB.GetDIBWidth = pdImages(g_CurrentImage).Width) And (.layerDIB.GetDIBHeight = pdImages(g_CurrentImage).Height))
        End With
        
        pdImages(g_CurrentImage).MergeTwoLayers pdImages(g_CurrentImage).ScratchLayer, pdImages(g_CurrentImage).GetActiveLayer, bottomLayerFullSize, True, VarPtr(tmpRectF)
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, pdImages(g_CurrentImage).GetActiveLayerIndex
        
        'Ask the central processor to create Undo/Redo data for us
        Processor.Process "Fill tool", , , UNDO_LAYER, g_CurrentTool
        
        'Reset the scratch layer
        pdImages(g_CurrentImage).ScratchLayer.layerDIB.ResetDIB 0
    
    'If the layer beneath this one is *not* a raster layer, let's add the fill as a new layer, instead.
    Else
        
        'Before creating the new layer, check for an active selection.  If one exists, we need to preprocess
        ' the fill layer against it.
        If pdImages(g_CurrentImage).IsSelectionActive Then
            
            'A selection is active.  Pre-mask the scratch layer against it.
            Dim cBlender As pdPixelBlender
            Set cBlender = New pdPixelBlender
            cBlender.ApplyMaskToTopDIB pdImages(g_CurrentImage).ScratchLayer.layerDIB, pdImages(g_CurrentImage).MainSelection.GetMaskDIB, VarPtr(tmpRectF)
            
        End If
        
        Dim newLayerID As Long
        newLayerID = pdImages(g_CurrentImage).CreateBlankLayer(pdImages(g_CurrentImage).GetActiveLayerIndex)
        
        'Point the new layer index at our scratch layer
        pdImages(g_CurrentImage).PointLayerAtNewObject newLayerID, pdImages(g_CurrentImage).ScratchLayer
        pdImages(g_CurrentImage).GetLayerByID(newLayerID).SetLayerName g_Language.TranslateMessage("Fill layer")
        Set pdImages(g_CurrentImage).ScratchLayer = Nothing
        
        'Activate the new layer
        pdImages(g_CurrentImage).SetActiveLayerByID newLayerID
        
        'Notify the parent image of the new layer
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_IMAGE_VECTORSAFE
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
        'Ask the central processor to create Undo/Redo data for us
        Processor.Process "Fill tool", , , UNDO_IMAGE_VECTORSAFE, g_CurrentTool
        
        'Create a new scratch layer
        Tools.InitializeToolsDependentOnImage
        
    End If
    
End Sub

'Render a relevant fill cursor outline to the canvas, using the stored mouse coordinates as the cursor's position
Public Sub RenderFillCursor(ByRef targetCanvas As pdCanvas)
    
    'Start by creating a transformation from the image space to the canvas space
    Dim canvasMatrix As pd2DTransform
    Drawing.GetTransformFromImageToCanvas canvasMatrix, targetCanvas, pdImages(g_CurrentImage), m_MouseX, m_MouseY
    
    'We also want to pinpoint the precise cursor position
    Dim cursX As Double, cursY As Double
    Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, pdImages(g_CurrentImage), m_MouseX, m_MouseY, cursX, cursY
    
    'Borrow a pair of UI pens from the main rendering module
    Dim innerPen As pd2DPen, outerPen As pd2DPen
    Drawing.BorrowCachedUIPens outerPen, innerPen
    
    'Create other required pd2D drawing tools (a surface)
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, targetCanvas.hDC, True
    
    'Paint a target cursor
    Dim crossLength As Single, outerCrossBorder As Single
    crossLength = 5#
    outerCrossBorder = 0.5
    
    If (m_Painter Is Nothing) Then Set m_Painter = New pd2DPainter
    m_Painter.DrawLineF cSurface, outerPen, cursX, cursY - crossLength - outerCrossBorder, cursX, cursY + crossLength + outerCrossBorder
    m_Painter.DrawLineF cSurface, outerPen, cursX - crossLength - outerCrossBorder, cursY, cursX + crossLength + outerCrossBorder, cursY
    m_Painter.DrawLineF cSurface, innerPen, cursX, cursY - crossLength, cursX, cursY + crossLength
    m_Painter.DrawLineF cSurface, innerPen, cursX - crossLength, cursY, cursX + crossLength, cursY
    
    Set cSurface = Nothing
    Set innerPen = Nothing: Set outerPen = Nothing
    
End Sub

Public Sub SetFillAA(ByVal newAA As Boolean)
    If EnsureFillerExists Then m_FloodFill.SetAntialiasingMode newAA
End Sub

Public Sub SetFillAlphaMode(ByVal newAlphaMode As PD_AlphaMode)
    m_FillAlphaMode = newAlphaMode
End Sub

Public Sub SetFillBlendMode(ByVal newBlendMode As PD_BlendMode)
    m_FillBlendMode = newBlendMode
End Sub

Public Sub SetFillBrush(ByVal newBrush As String)
    m_FillBrush = newBrush
End Sub

Public Sub SetFillCompareMode(ByVal newCompareMode As PD_FloodCompare)
    If EnsureFillerExists Then m_FloodFill.SetCompareMode newCompareMode
End Sub

Public Sub SetFillSampleMerged(ByVal sampleMerged As Boolean)
    m_FillSampleMerged = sampleMerged
End Sub

Public Sub SetFillSearchMode(ByVal newSearchMode As PD_FloodSearch)
    If EnsureFillerExists Then m_FloodFill.SetSearchMode newSearchMode
End Sub

Public Sub SetFillTolerance(ByVal newTolerance As Single)
    If EnsureFillerExists Then m_FloodFill.SetTolerance newTolerance
End Sub

'Before PD closes, you *must* call this function!  It will free any lingering fill resources (which are cached
' for performance reasons).  You can also call this function any time without penalty, if you need to free
' up memory or GDI resources.  (Freed objects are automatically recreated as-needed.)
Public Sub FreeFillResources()
    If (Not m_FloodFill Is Nothing) Then m_FloodFill.FreeUpResources
    Set m_FillOutline = Nothing
    Set m_FillImage = Nothing
End Sub

