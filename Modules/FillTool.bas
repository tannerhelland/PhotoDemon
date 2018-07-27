Attribute VB_Name = "Tools_Fill"
'***************************************************************************
'PhotoDemon Bucket Fill Manager
'Copyright 2017-2018 by Tanner Helland
'Created: 30/August/17
'Last updated: 04/September/17
'Last update: continued work on initial build
'
'This module interfaces between the bucket fill UI and pdFloodFill backend.  Look in the relevant tool panel
' form for more details on how the UI relays relevant fill data here.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
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

'Current mouse/pen input values.  These are blindly relayed to us by the canvas, and it's up to us to perform any
' special tracking calculations.
Private m_MouseDown As Boolean
Private m_MouseX As Single, m_MouseY As Single

'Most fill "properties" are forwarded to a pdFloodFill instance, but rendering-specific properties are actually
' cached internally.  (pdFloodFill just generates a flood outline - we have to do the actual filling.)
Private m_FillBlendMode As PD_BlendMode, m_FillAlphaMode As PD_AlphaMode
Private m_FillSampleMerged As Boolean

'Fills can be "filled" in two ways: using a standard color + opacity combo, or using a full-on custom brush.
Public Enum PD_FillToolSource
    fts_ColorOpacity = 0
    fts_CustomBrush = 1
End Enum

#If False Then
    Private Const fts_ColorOpacity = 0, fts_CustomBrush = 1
#End If

Private m_FillSource As PD_FillToolSource, m_FillBrush As String, m_FillColor As Long, m_FillOpacity As Single

'Bucket cursor; this is loaded as an on-demand resource, and cached after first use.
Private m_FillCursor As pdDIB

'Before attempting to set flood fill properties, call this sub to ensure the m_FloodFill object exists.
' (It returns TRUE if m_FloodFill exists.)
Private Function EnsureFillerExists() As Boolean
    If (m_FloodFill Is Nothing) Then Set m_FloodFill = New pdFloodFill
    EnsureFillerExists = True
End Function

'Similar to other tools, the fill tool is notified of all mouse actions that occur while it is selected.  At present, however,
' it only triggers a fill when the mouse is actually clicked.  (No action is taken on move events.)
Public Sub NotifyMouseXY(ByVal mouseButtonDown As Boolean, ByVal imgX As Single, ByVal imgY As Single, ByRef srcCanvas As pdCanvas)
    
    Dim oldMouseState As Boolean
    oldMouseState = m_MouseDown
    
    'Update all internal trackers.  (Note that the passed positions are in *image* coordinates, not *screen* coordinates.)
    m_MouseX = imgX
    m_MouseY = imgY
    m_MouseDown = mouseButtonDown
    
    'Different fill modes use different coordinate spaces.  (For example, "sample image" and "sample layer" use different
    ' underlying DIBs to calculate their fill, so we may want to translate the incoming coordinates - which are always in
    ' the *image* coordinate space - to the current layer's coordinate space.)
    Dim fillStartX As Long, fillStartY As Long
    If m_FillSampleMerged Then
        fillStartX = Int(imgX)
        fillStartY = Int(imgY)
    Else
        
        Dim newX As Single, newY As Single
        Drawing.ConvertImageCoordsToLayerCoords_Full pdImages(g_CurrentImage), pdImages(g_CurrentImage).GetActiveLayer, imgX, imgY, newX, newY
        fillStartX = Int(newX)
        fillStartY = Int(newY)
        
    End If
    
    'If this is an initial click, apply the fill.
    If mouseButtonDown And (Not oldMouseState) Then
        
        'Before proceeding, validate the click position.  Unlike paintbrush strokes, fill start points must lie on the
        ' underlying image/layer (depending on the current sampling mode).
        Dim allowedToFill As Boolean
        If m_FillSampleMerged Then
            allowedToFill = PDMath.IsPointInRectF(fillStartX, fillStartY, pdImages(g_CurrentImage).GetBoundaryRectF)
        Else
            
            Dim tmpRectF As RectF
            With tmpRectF
                .Left = 0!
                .Top = 0!
                .Width = pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False)
                .Height = pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False)
            End With
            allowedToFill = PDMath.IsPointInRectF(fillStartX, fillStartY, tmpRectF)
            
        End If
        
        If (Not allowedToFill) Then Exit Sub
        
        'We are allowed to perform a fill.  Notify the central "color history" manager of the color currently
        ' being used (so it can be added to the dynamic color history list).
        If (m_FillSource = fts_ColorOpacity) Then UserControls.PostPDMessage WM_PD_PRIMARY_COLOR_APPLIED, m_FillColor, , True
        
        'Start by grabbing (or producing) the source DIB required for the fill.  (If the user wants us to
        ' sample all layers, we need to generate a composite image.)
        
        'fillSrc is just a thin reference to some other DIB; it is never created as a new object
        Dim fillSrc As pdDIB
        
        'Merged sampling requires us to maintain a local copy of the fully composited image stack.
        If m_FillSampleMerged Then
            
            Dim fillImageRefreshRequired As Boolean
            fillImageRefreshRequired = (m_FillImage Is Nothing)
            If (Not fillImageRefreshRequired) Then fillImageRefreshRequired = (m_FillImageTimestamp <> pdImages(g_CurrentImage).GetTimeOfLastChange())
            
            'A new image copy is required.  As much as possible, still try to minimize the work we do here
            If fillImageRefreshRequired Then
                
                If (m_FillImage Is Nothing) Then Set m_FillImage = New pdDIB
                If (m_FillImage.GetDIBWidth <> pdImages(g_CurrentImage).Width) Or (m_FillImage.GetDIBHeight <> pdImages(g_CurrentImage).Height) Then
                    m_FillImage.CreateBlank pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, 32, 0, 0
                Else
                    m_FillImage.ResetDIB 0
                End If
                
                'Note the timestamp; this may allow us to skip subsequent copy requests
                m_FillImageTimestamp = pdImages(g_CurrentImage).GetTimeOfLastChange()
            
            End If
            
            'Merged image data requires us to obtain a fully composited copy of the current image.
            pdImages(g_CurrentImage).GetCompositedImage m_FillImage
            
            Set fillSrc = m_FillImage
        
        'If the user is only filling the current layer, we can skip the compositing step (yay!) but we also need to convert
        ' the incoming (x, y) coordinates into layer-space coordinates.
        Else
            
            Set fillSrc = pdImages(g_CurrentImage).GetActiveLayer.layerDIB
            
            'To improve performance when variable fill blend and alpha modes are in place, we always render the fill to
            ' a temporary image, then perform a standard merge of that DIB onto the target layer.  This lets the bucket
            ' fill algorithm run at maximum speed, regardless of underlying fill settings.
            
            'Initialize our temporary image now, using the same m_FillImage object we use for other purposes.
            If (m_FillImage Is Nothing) Then Set m_FillImage = New pdDIB
            If (m_FillImage.GetDIBWidth <> fillSrc.GetDIBWidth) Or (m_FillImage.GetDIBHeight <> fillSrc.GetDIBHeight) Then
                m_FillImage.CreateBlank fillSrc.GetDIBWidth, fillSrc.GetDIBHeight, 32, 0, 0
            Else
                m_FillImage.ResetDIB 0
            End If
            
        End If
        
        'Failsafe check for an active m_FloodFill instance (which should always exist here, because create while setting
        ' all relevant fill properties)
        If (m_FloodFill Is Nothing) Then Debug.Print "WARNING!  m_FloodFill doesn't exist!"
        
        'Set the initial flood point
        m_FloodFill.SetInitialPoint fillStartX, fillStartY
        
        'Apply the flood fill.  Note that unlike other places in PD, we *do not* pass a destination DIB.  We are only
        ' interested in the outline of the filled area, in vector form.
        If (m_FillOutline Is Nothing) Then Set m_FillOutline = New pd2DPath Else m_FillOutline.ResetPath
        m_FloodFill.InitiateFloodFill fillSrc, Nothing, m_FillOutline
        Set fillSrc = Nothing
        
        'Create a brush using the previously passed brush settings
        Dim tmpBrush As pd2DBrush
        Set tmpBrush = New pd2DBrush
        
        If (m_FillSource = fts_ColorOpacity) Then
            Drawing2D.QuickCreateSolidBrush tmpBrush, m_FillColor, m_FillOpacity
        Else
            tmpBrush.SetBrushPropertiesFromXML m_FillBrush
        End If
        
        'Gradient brushes require us to set a gradient rect.  Importantly, we need to apply any relevant offsets
        ' (e.g. the offsets required when "filling" a vector layer) to our path *before* calculating gradient boundaries.
        ' (Remember that vector layers are treated differently - to avoid the need to rasterize them, we instead
        ' apply the fill to a *new* layer that sits above the original one.)
        If pdImages(g_CurrentImage).GetActiveLayer.IsLayerVector And (Not m_FillSampleMerged) Then
        
            'Transform the fill outline by the current layer transformation, if any
            Dim cTransform As pd2DTransform
            pdImages(g_CurrentImage).GetActiveLayer.GetCopyOfLayerTransformationMatrix_Full cTransform
            m_FillOutline.ApplyTransformation cTransform
        
        End If
        
        'Calculate the gradient rect (if any), using the boundary rect of the fill outline.
        If (tmpBrush.GetBrushMode = P2_BM_Gradient) Then
            Dim fillBoundaryRect As RectF
            fillBoundaryRect = m_FillOutline.GetPathBoundariesF
            tmpBrush.SetBoundaryRect fillBoundaryRect
        End If
        
        Dim tmpSurface As pd2DSurface
        Set tmpSurface = New pd2DSurface
        
        'Render the final fill.  Note that when we are just filling the current layer, we potentially render the result
        ' differently, if the layer is smaller than the active image (or if it has non-destructive transforms active or
        ' is a vector layer).
        '
        '(At present, selections don't work with this accelerated technique, but that is due to be fixed shortly.)
        Dim useCustomDIB As Boolean
        With pdImages(g_CurrentImage).GetActiveLayer
            
            useCustomDIB = (Not .IsLayerVector)
            
            'If the current layer is the same size as the image (very common during photo editing sessions, as the
            ' user is probably just loading standlone JPEGs), we don't need to use a custom code path.
            If useCustomDIB Then
                If (.GetLayerOffsetX = 0#) And (.GetLayerOffsetY = 0#) And _
                   (.GetLayerWidth(True) = pdImages(g_CurrentImage).Width) And _
                   (.GetLayerHeight(True) = pdImages(g_CurrentImage).Height) And _
                   (Not .AffineTransformsActive) Then
                   
                   useCustomDIB = False
                   
                End If
            End If
            
        End With
        
        If m_FillSampleMerged Or (Not useCustomDIB) Then
            
            'A scratch layer should always be guaranteed to exist, so this exists purely as a paranoid failsafe.
            If (pdImages(g_CurrentImage).ScratchLayer Is Nothing) Then
                PDDebug.LogAction "WARNING!  Tools_Fill.NotifyMouseXY tried to merge into a blank scratch layer!"
                pdImages(g_CurrentImage).ResetScratchLayer True
            End If
            
            'When filling a merged region, we render the result directly onto the current image's scratch layer.
            ' The full scratch layer will then be merged with the layer beneath it.
            pdImages(g_CurrentImage).ScratchLayer.layerDIB.ResetDIB 0
            pdImages(g_CurrentImage).ScratchLayer.layerDIB.SetInitialAlphaPremultiplicationState True
            
            tmpSurface.WrapSurfaceAroundPDDIB pdImages(g_CurrentImage).ScratchLayer.layerDIB
            If m_FloodFill.GetAntialiasingMode Then tmpSurface.SetSurfaceAntialiasing P2_AA_HighQuality Else tmpSurface.SetSurfaceAntialiasing P2_AA_None
            tmpSurface.SetSurfacePixelOffset P2_PO_Half
            
            PD2D.FillPath tmpSurface, tmpBrush, m_FillOutline
            
            'Free all finished pd2D objects
            Set tmpSurface = Nothing: Set tmpBrush = Nothing
            
            'Relay the correct blend and alpha settings to the scratch layer, then permanently commit the results
            pdImages(g_CurrentImage).ScratchLayer.SetLayerBlendMode m_FillBlendMode
            pdImages(g_CurrentImage).ScratchLayer.SetLayerAlphaMode m_FillAlphaMode
            Tools_Fill.CommitFillResults False
            
        Else
        
            'If we are only operating on the currently active layer, and that layer is not the same size as the image,
            ' let's take a smarter approach.  Instead of using the full scratch layer (which is always image-sized),
            ' we'll use a temporary DIB at the same size as the image, and paint to that instead.
            '
            '(Note that such a temporary DIB was already created in a previous step - see earlier in this function.)
            tmpSurface.WrapSurfaceAroundPDDIB m_FillImage
            If m_FloodFill.GetAntialiasingMode Then tmpSurface.SetSurfaceAntialiasing P2_AA_HighQuality Else tmpSurface.SetSurfaceAntialiasing P2_AA_None
            tmpSurface.SetSurfacePixelOffset P2_PO_Half
            
            PD2D.FillPath tmpSurface, tmpBrush, m_FillOutline
            
            'Free all finished pd2D objects
            Set tmpSurface = Nothing: Set tmpBrush = Nothing
            
            'Commit the results permanently
            m_FillImage.SetInitialAlphaPremultiplicationState True
            Tools_Fill.CommitFillResults True, m_FillImage
            
        End If
        
    End If
    
End Sub

'Want to commit your current fill work?  Call this function to make any pending fill results permanent.
Public Sub CommitFillResults(ByVal useCustomDIB As Boolean, Optional ByRef fillDIB As pdDIB = Nothing)
    
    Dim cBlender As pdPixelBlender
    Set cBlender = New pdPixelBlender
                
    'If the caller supplied a custom DIB, skip ahead and use that.  (This provides a performance boost on single-layer
    ' raster images, or filling a single layer in a multi-layer image when it's smaller than the composited image.)
    If useCustomDIB Then
    
        'If a selection is active, we need to mask the fill by the current selection mask
        If pdImages(g_CurrentImage).IsSelectionActive Then
            
            'Next, we need to grab a copy of the current selection mask, mirroring the area where our layer lives
            Dim tmpSelDIB As pdDIB
            Set tmpSelDIB = New pdDIB
            tmpSelDIB.CreateBlank fillDIB.GetDIBWidth, fillDIB.GetDIBHeight, 32, 0, 0
            
            'If no weird affine transforms are active, this step is easy
            If (Not pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)) Then
                GDI.BitBltWrapper tmpSelDIB.GetDIBDC, 0, 0, fillDIB.GetDIBWidth, fillDIB.GetDIBHeight, pdImages(g_CurrentImage).MainSelection.GetMaskDC, pdImages(g_CurrentImage).GetActiveLayer.GetLayerOffsetX, pdImages(g_CurrentImage).GetActiveLayer.GetLayerOffsetY, vbSrcCopy
            
            'Affine transforms make this step noticeably more unpleasant
            Else
            
                'Wrap our temporary selection DIB with a pd2D surface
                Dim cSurface As pd2DSurface
                Set cSurface = New pd2DSurface
                cSurface.WrapSurfaceAroundPDDIB tmpSelDIB
                cSurface.SetSurfaceResizeQuality P2_RQ_Bilinear
                
                'Create a transform matching the target layer
                Dim cTransform As pd2DTransform
                pdImages(g_CurrentImage).GetActiveLayer.GetCopyOfLayerTransformationMatrix_Full cTransform
                
                'Activate the transform for our selection mask copy
                cTransform.InvertTransform
                cSurface.SetSurfaceWorldTransform cTransform
                
                'Paint the selection mask into place, with the transform active
                Dim tmpSrcSurface As pd2DSurface
                Set tmpSrcSurface = New pd2DSurface
                tmpSrcSurface.WrapSurfaceAroundPDDIB pdImages(g_CurrentImage).MainSelection.GetMaskDIB
                PD2D.DrawSurfaceF cSurface, 0, 0, tmpSrcSurface
                
                Set tmpSrcSurface = Nothing
                Set cSurface = Nothing
            
            End If
            
            cBlender.ApplyMaskToTopDIB fillDIB, tmpSelDIB
            Set tmpSelDIB = Nothing
            
        End If
        
        'The fillDIB object already contains a standalone copy of the fill results.  We simply need to merge it
        ' onto the base layer using the appropriate fill settings.
        Dim cCompositor As pdCompositor
        Set cCompositor = New pdCompositor
        cCompositor.QuickMergeTwoDibsOfEqualSize pdImages(g_CurrentImage).GetActiveDIB, fillDIB, m_FillBlendMode, 100#, pdImages(g_CurrentImage).GetActiveLayer.GetLayerAlphaMode, m_FillAlphaMode
        
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_Layer, pdImages(g_CurrentImage).GetActiveLayerIndex
        
        'Ask the central processor to create Undo/Redo data for us
        Processor.Process "Fill tool", , , UNDO_Layer, g_CurrentTool
    
    'If useCustomDIB is FALSE, the current scratch layer contains everything we need for the blend.
    Else
        
        'Start by grabbing the boundaries of the fill area, and clipping it to the image's bounds, as necessary
        Dim tmpRectF As RectF
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
                bottomLayerFullSize = ((.GetLayerOffsetX = 0!) And (.GetLayerOffsetY = 0!) And (.layerDIB.GetDIBWidth = pdImages(g_CurrentImage).Width) And (.layerDIB.GetDIBHeight = pdImages(g_CurrentImage).Height))
            End With
            
            pdImages(g_CurrentImage).MergeTwoLayers pdImages(g_CurrentImage).ScratchLayer, pdImages(g_CurrentImage).GetActiveLayer, bottomLayerFullSize, True   ', VarPtr(tmpRectF)
            pdImages(g_CurrentImage).NotifyImageChanged UNDO_Layer, pdImages(g_CurrentImage).GetActiveLayerIndex
            
            'Before proceeding, trim any empty borders in the resulting layer.  (It will always be the size of the image,
            ' because the scratch layer is always image-sized.)
            pdImages(g_CurrentImage).GetActiveLayer.CropNullPaddedLayer
            
            'Ask the central processor to create Undo/Redo data for us
            Processor.Process "Fill tool", , , UNDO_Layer, g_CurrentTool
            
            'Reset the scratch layer
            pdImages(g_CurrentImage).ScratchLayer.layerDIB.ResetDIB 0
        
        'If the layer beneath this one is *not* a raster layer, let's add the fill as a new layer, instead.
        Else
            
            'Before creating the new layer, check for an active selection.  If one exists, we need to preprocess
            ' the fill layer against it.
            If pdImages(g_CurrentImage).IsSelectionActive Then
                
                'A selection is active.  Pre-mask the scratch layer against it.
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
            
            'Crop any dead space from the scratch layer
            pdImages(g_CurrentImage).GetActiveLayer.CropNullPaddedLayer
            
            'Notify the parent image of the new layer
            pdImages(g_CurrentImage).NotifyImageChanged UNDO_Image_VectorSafe
            
            'Redraw the layer box, and note that thumbnails need to be re-cached
            toolbar_Layers.NotifyLayerChange
            
            'Ask the central processor to create Undo/Redo data for us
            Processor.Process "Fill tool", , , UNDO_Image_VectorSafe, g_CurrentTool
            
            'Create a new scratch layer
            Tools.InitializeToolsDependentOnImage
            
        End If
    
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
    
    PD2D.DrawLineF cSurface, outerPen, cursX, cursY - crossLength - outerCrossBorder, cursX, cursY + crossLength + outerCrossBorder
    PD2D.DrawLineF cSurface, outerPen, cursX - crossLength - outerCrossBorder, cursY, cursX + crossLength + outerCrossBorder, cursY
    PD2D.DrawLineF cSurface, innerPen, cursX, cursY - crossLength, cursX, cursY + crossLength
    PD2D.DrawLineF cSurface, innerPen, cursX - crossLength, cursY, cursX + crossLength, cursY
    
    'If we haven't loaded the fill cursor previously, do so now
    If (m_FillCursor Is Nothing) Then
        IconsAndCursors.LoadResourceToDIB "cursor_bucket", m_FillCursor, IconsAndCursors.GetSystemCursorSizeInPx() * 1.15, IconsAndCursors.GetSystemCursorSizeInPx() * 1.15, , , True
    End If
    
    'Paint the fill icon to the bottom-right of the actual cursor, Photoshop-style
    Dim icoSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDIB icoSurface, m_FillCursor, True
    icoSurface.SetSurfaceResizeQuality P2_RQ_Bilinear
    PD2D.DrawSurfaceF cSurface, cursX + crossLength * 1.4!, cursY + crossLength * 1.4!, icoSurface
    
    Set cSurface = Nothing: Set icoSurface = Nothing
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

Public Sub SetFillBrush(ByRef newBrush As String)
    m_FillBrush = newBrush
End Sub

Public Sub SetFillBrushColor(ByVal newColor As Long)
    m_FillColor = newColor
End Sub

Public Sub SetFillBrushOpacity(ByVal newOpacity As Single)
    m_FillOpacity = newOpacity
End Sub

Public Sub SetFillBrushSource(ByVal newBrushSource As PD_FillToolSource)
    m_FillSource = newBrushSource
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

