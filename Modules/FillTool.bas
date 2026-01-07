Attribute VB_Name = "Tools_Fill"
'***************************************************************************
'PhotoDemon Bucket Fill Manager
'Copyright 2017-2026 by Tanner Helland
'Created: 30/August/17
'Last updated: 22/May/24
'Last update: automatically redirect "sample from layer" to "sample from image" when the cursor is off-layer
'
'This module interfaces between the bucket fill UI and pdFloodFill backend.  Look in the relevant tool panel
' form for more details on how the UI relays relevant fill data here.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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

'Flood-filling vector layers requires use of a temporary fill DIB (as we can't apply the fill directly
' onto the target vector layer).  This value is set during mouse_down, and reset on mouse_up.
Private m_TempDIBInUse As Boolean

'If the user clicks somewhere invalid, this flag will be set to TRUE on _MouseDown.
' On _MouseUp, we check it before committing fill results.
Private m_FillCanceled As Boolean

'If the user clicks off-layer but is using "sample from layer", we silently redirect to "sample from image".
' We need to track this value across both "mouse up" and "mouse down".
Private m_FillSourceOverrideActive As Boolean

'Before attempting to set flood fill properties, call this sub to ensure the m_FloodFill object exists.
' (It returns TRUE if m_FloodFill exists.)
Private Function EnsureFillerExists() As Boolean
    If (m_FloodFill Is Nothing) Then Set m_FloodFill = New pdFloodFill
    EnsureFillerExists = True
End Function

'Similar to other tools, the fill tool is notified of all mouse actions that occur while it is selected.  At present, however,
' it only triggers a fill when the mouse is actually clicked.  (No action is taken on move events.)
Public Sub NotifyMouseXY(ByVal mouseButtonDown As Boolean, ByVal imgX As Single, ByVal imgY As Single, ByRef srcCanvas As pdCanvas)
    
    'm_FillSampleMerged tracks sampling image vs layer.  We may change this in this function (for example, if the user
    ' clicks off the active layer, we can't sample from layer - so assume the user wants to sample merged instead)
    ' so beyond this point, *only use the LOCAL sample merged* tracker
    Dim useSampleMerged As Boolean
    useSampleMerged = m_FillSampleMerged
    
    Dim oldMouseState As Boolean
    oldMouseState = m_MouseDown
    
    'Update all internal trackers.  (Note that the passed positions are in *image* coordinates, not *screen* coordinates.)
    m_MouseX = imgX
    m_MouseY = imgY
    m_MouseDown = mouseButtonDown
    
    'Produce a parameter string that stores all data for the current fill (including coordinates!)
    Dim curFillParams As String
    If ((Macros.GetMacroStatus <> MacroPLAYBACK) And (Macros.GetMacroStatus <> MacroBATCH)) Then
        curFillParams = GetAllFillSettings(imgX, imgY)
    Else
        curFillParams = vbNullString
    End If
    
    'Different fill modes use different coordinate spaces.  (For example, "sample image" and "sample layer" use different
    ' underlying DIBs to calculate their fill, so we may want to translate the incoming coordinates - which are always in
    ' the *image* coordinate space - to the current layer's coordinate space.)
    Dim fillStartX As Long, fillStartY As Long
    If useSampleMerged Then
        fillStartX = Int(imgX)
        fillStartY = Int(imgY)
    Else
        Dim newX As Single, newY As Single
        Drawing.ConvertImageCoordsToLayerCoords_Full PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, imgX, imgY, newX, newY
        fillStartX = Int(newX)
        fillStartY = Int(newY)
    End If
    
    'If the mouse button is down, generate a new fill area
    If mouseButtonDown Then
        
        m_FillSourceOverrideActive = False
        
        'Before proceeding, validate the click position.  Unlike paintbrush strokes, fill start points must lie on the
        ' underlying image/layer (depending on the current sampling mode).
        
        'Sample from image...
        If useSampleMerged Then
            m_FillCanceled = Not PDMath.IsPointInRectF(fillStartX, fillStartY, PDImages.GetActiveImage.GetBoundaryRectF)
        
        'Sample from layer...
        Else
            
            Dim tmpRectF As RectF
            With tmpRectF
                .Left = 0!
                .Top = 0!
                .Width = PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)
                .Height = PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
            End With
            m_FillCanceled = Not PDMath.IsPointInRectF(fillStartX, fillStartY, tmpRectF)
            
            'The user clicked outside the active layer.  If this occurs, silently redirect to the (likely) expected behavior,
            ' which is "sample merged".
            If m_FillCanceled Then
                
                PDDebug.LogAction "Fill layer won't work; attempting fill image instead"
                PDDebug.LogAction fillStartX & ", " & fillStartY & ", " & tmpRectF.Width & ", " & tmpRectF.Height & ", " & imgX & ", " & imgY
                
                'Reset coordinates to *image* coordinate space (not *layer* coordinate space)
                fillStartX = Int(imgX)
                fillStartY = Int(imgY)
                
                'Attempt an image fill, and if it's valid, proceed as if this is a merged fill
                m_FillCanceled = Not PDMath.IsPointInRectF(fillStartX, fillStartY, PDImages.GetActiveImage.GetBoundaryRectF)
                useSampleMerged = (Not m_FillCanceled)
                If useSampleMerged Then PDDebug.LogAction "Proceeding with merged fill instead..."
                m_FillSourceOverrideActive = True
                
                'We also need to update our param string to reflect the new setting
                Dim tmpParams As pdSerialize
                Set tmpParams = New pdSerialize
                tmpParams.SetParamString curFillParams
                tmpParams.UpdateParam "fill-source", useSampleMerged
                
            End If
            
        End If
        
        If m_FillCanceled Then Exit Sub
        
        'We are allowed to perform a fill.  Notify the central "color history" manager of the color currently
        ' being used (so it can be added to the dynamic color history list).
        If (m_FillSource = fts_ColorOpacity) Then UserControls.PostPDMessage WM_PD_PRIMARY_COLOR_APPLIED, m_FillColor, , True
        
        'Start by grabbing (or producing) the source DIB required for the fill.
        ' (If the user wants us to sample all layers, we need to generate a composite image.)
        
        'fillSrc is just a thin reference to some other DIB; it is never created as a new object
        Dim fillSrc As pdDIB
        
        'Merged sampling requires us to maintain a local copy of the fully composited image stack.
        If useSampleMerged Then
            
            Dim fillImageRefreshRequired As Boolean
            fillImageRefreshRequired = (m_FillImage Is Nothing)
            If (Not fillImageRefreshRequired) Then fillImageRefreshRequired = (m_FillImageTimestamp <> PDImages.GetActiveImage.GetTimeOfLastChange())
            
            'A new image copy is required.  As much as possible, still try to minimize the work we do here
            If fillImageRefreshRequired Then
                
                If (m_FillImage Is Nothing) Then Set m_FillImage = New pdDIB
                If (m_FillImage.GetDIBWidth <> PDImages.GetActiveImage.Width) Or (m_FillImage.GetDIBHeight <> PDImages.GetActiveImage.Height) Then
                    m_FillImage.CreateBlank PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 32, 0, 0
                Else
                    m_FillImage.ResetDIB 0
                End If
                
                'Note the timestamp; this may allow us to skip subsequent copy requests
                m_FillImageTimestamp = PDImages.GetActiveImage.GetTimeOfLastChange()
            
            End If
            
            'Merged image data requires us to obtain a fully composited copy of the current image.
            PDImages.GetActiveImage.GetCompositedImage m_FillImage
            
            Set fillSrc = m_FillImage
            
        'If the user is only filling the current layer, we can skip the compositing step (yay!) but we also need to convert
        ' the incoming (x, y) coordinates into layer-space coordinates.
        Else
            
            Set fillSrc = PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB
            
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
        
        'Failsafe check for an active m_FloodFill instance (which should always exist here, because we created it
        ' while setting relevant fill properties)
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
        If PDImages.GetActiveImage.GetActiveLayer.IsLayerVector And (Not useSampleMerged) Then
        
            'Transform the fill outline by the current layer transformation, if any
            Dim cTransform As pd2DTransform
            PDImages.GetActiveImage.GetActiveLayer.GetCopyOfLayerTransformationMatrix_Full cTransform
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
        With PDImages.GetActiveImage.GetActiveLayer
        
            m_TempDIBInUse = (Not .IsLayerVector)
            
            'If the current layer is the same size as the image (very common during photo editing sessions, as the
            ' user is probably just loading standlone JPEGs), we don't need to use a custom code path.
            If m_TempDIBInUse Then
                If (.GetLayerOffsetX = 0#) And (.GetLayerOffsetY = 0#) And _
                   (.GetLayerWidth(True) = PDImages.GetActiveImage.Width) And _
                   (.GetLayerHeight(True) = PDImages.GetActiveImage.Height) And _
                   (Not .AffineTransformsActive) Then
                   
                   m_TempDIBInUse = False
                   
                End If
            End If
            
        End With
        
        If useSampleMerged Or (Not m_TempDIBInUse) Then
            
            'A scratch layer should always be guaranteed to exist, so this exists purely as a paranoid failsafe.
            PDImages.GetActiveImage.ResetScratchLayer True
            
            'When filling a merged region, we render the result directly onto the current image's scratch layer.
            ' The full scratch layer will then be merged with the layer beneath it.
            PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.ResetDIB 0
            PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.SetInitialAlphaPremultiplicationState True
            
            tmpSurface.WrapSurfaceAroundPDDIB PDImages.GetActiveImage.ScratchLayer.GetLayerDIB
            If m_FloodFill.GetAntialiasingMode Then tmpSurface.SetSurfaceAntialiasing P2_AA_HighQuality Else tmpSurface.SetSurfaceAntialiasing P2_AA_None
            tmpSurface.SetSurfacePixelOffset P2_PO_Half
            
            PD2D.FillPath tmpSurface, tmpBrush, m_FillOutline
            
            'Free all finished pd2D objects
            Set tmpSurface = Nothing: Set tmpBrush = Nothing
            
            'Relay the correct blend and alpha settings to the scratch layer
            PDImages.GetActiveImage.ScratchLayer.SetLayerBlendMode m_FillBlendMode
            PDImages.GetActiveImage.ScratchLayer.SetLayerAlphaMode m_FillAlphaMode
            
        Else
        
            'If we are only operating on the currently active layer, and that layer is not the same size as the image,
            ' let's take a smarter approach.  Instead of using the full scratch layer (which is always image-sized),
            ' we'll use a temporary DIB at the same size as the current layer, and paint to that instead.
            '
            '(Note that such a temporary DIB was already created in a previous step - see earlier in this function.)
            tmpSurface.WrapSurfaceAroundPDDIB m_FillImage
            If m_FloodFill.GetAntialiasingMode Then tmpSurface.SetSurfaceAntialiasing P2_AA_HighQuality Else tmpSurface.SetSurfaceAntialiasing P2_AA_None
            tmpSurface.SetSurfacePixelOffset P2_PO_Half
            
            PD2D.FillPath tmpSurface, tmpBrush, m_FillOutline
            
            'Free all finished pd2D objects
            Set tmpSurface = Nothing: Set tmpBrush = Nothing
            
            'Note that the fill image is already premultiplied
            m_FillImage.SetInitialAlphaPremultiplicationState True
            
        End If
        
        'In the future, a "live preview" could be displayed here.  This would be nice as the user could move the mouse
        ' to see how it affects rendering, or perhaps use the mousewheel to dynamically change fill tolerance
        ' (and see immediate results).  However, the rendering nuances of doing this in real-time are complicated
        ' (largely due to things like real-time affine transforms on the target layer), and it would require more
        ' testing than I can perform at present.  Consider this a low-priority TODO.
        
    End If
    
    'If the user is releasing the mouse, commit the fill results permanently
    If (Not mouseButtonDown) And oldMouseState Then
        
        'Make sure the fill event wasn't canceled (this happens if the user clicks off
        ' the active image/layer with matching fill settings)
        If m_FillCanceled Then
            m_FillCanceled = False
            Exit Sub
        End If
        
        'Look for mode overrides
        If m_FillSourceOverrideActive Then
            useSampleMerged = True
        Else
            useSampleMerged = m_FillSampleMerged
        End If
        
        If useSampleMerged Or (Not m_TempDIBInUse) Then
            Tools_Fill.CommitFillResults False, , curFillParams
        Else
            Tools_Fill.CommitFillResults True, m_FillImage, curFillParams
        End If
        
    End If
    
End Sub

'Want to commit your current fill work?  Call this function to make any pending fill results permanent.
Public Sub CommitFillResults(ByVal useCustomDIB As Boolean, Optional ByRef fillDIB As pdDIB = Nothing, Optional ByRef curFillParams As String = vbNullString)
    
    Dim cBlender As pdPixelBlender
    Set cBlender = New pdPixelBlender
                
    'If the caller supplied a custom DIB, skip ahead and use that.  (This provides a performance boost on single-layer
    ' raster images, or filling a single layer in a multi-layer image when it's smaller than the composited image.)
    If useCustomDIB Then
    
        'If a selection is active, we need to mask the fill by the current selection mask
        If PDImages.GetActiveImage.IsSelectionActive Then
            
            'Next, we need to grab a copy of the current selection mask, mirroring the area where our layer lives
            Dim tmpSelDIB As pdDIB
            Set tmpSelDIB = New pdDIB
            tmpSelDIB.CreateBlank fillDIB.GetDIBWidth, fillDIB.GetDIBHeight, 32, 0, 0
            
            'If no weird affine transforms are active, this step is easy
            If (Not PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then
                GDI.BitBltWrapper tmpSelDIB.GetDIBDC, 0, 0, fillDIB.GetDIBWidth, fillDIB.GetDIBHeight, PDImages.GetActiveImage.MainSelection.GetMaskDC, PDImages.GetActiveImage.GetActiveLayer.GetLayerOffsetX, PDImages.GetActiveImage.GetActiveLayer.GetLayerOffsetY, vbSrcCopy
            
            'Affine transforms make this step noticeably more unpleasant
            Else
            
                'Wrap our temporary selection DIB with a pd2D surface
                Dim cSurface As pd2DSurface
                Set cSurface = New pd2DSurface
                cSurface.WrapSurfaceAroundPDDIB tmpSelDIB
                cSurface.SetSurfaceResizeQuality P2_RQ_Bilinear
                
                'Create a transform matching the target layer
                Dim cTransform As pd2DTransform
                PDImages.GetActiveImage.GetActiveLayer.GetCopyOfLayerTransformationMatrix_Full cTransform
                
                'Activate the transform for our selection mask copy
                cTransform.InvertTransform
                cSurface.SetSurfaceWorldTransform cTransform
                
                'Paint the selection mask into place, with the transform active
                Dim tmpSrcSurface As pd2DSurface
                Set tmpSrcSurface = New pd2DSurface
                tmpSrcSurface.WrapSurfaceAroundPDDIB PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB
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
        cCompositor.QuickMergeTwoDibsOfEqualSize PDImages.GetActiveImage.GetActiveDIB, fillDIB, m_FillBlendMode, 100#, PDImages.GetActiveImage.GetActiveLayer.GetLayerAlphaMode, m_FillAlphaMode
        
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
        
        'Ask the central processor to create Undo/Redo data for us
        Processor.Process "Fill tool", , curFillParams, UNDO_Layer, g_CurrentTool
    
    'If useCustomDIB is FALSE, the current scratch layer contains everything we need for the blend.
    Else
        
        'This dummy string only exists to ensure that the processor name gets localized properly
        ' (as that text is used for Undo/Redo descriptions).  PD's translation engine will detect
        ' the TranslateMessage() call and produce a matching translation entry.
        Dim strDummy As String
        strDummy = g_Language.TranslateMessage("Fill tool")
        Layers.CommitScratchLayer "Fill tool", m_FillOutline.GetPathBoundariesF, curFillParams
    
    End If
    
End Sub

'To be used *ONLY* during macro playback!
Public Sub PlayFillFromMacro(ByRef srcParams As String)

    'Failsafe check; the central processor should have verified this
    If ((Macros.GetMacroStatus = MacroPLAYBACK) Or (Macros.GetMacroStatus = MacroBATCH)) Then
    
        'Parse param string and call the appropriate filler
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString srcParams
        
        'Some properties come directly from the fill object
        If EnsureFillerExists Then
        
            With cParams
                Tools_Fill.SetFillAA .GetBool("antialias")
                Tools_Fill.SetFillAlphaMode .GetLong("alpha-mode")
                Tools_Fill.SetFillBlendMode .GetLong("blend-mode")
                Tools_Fill.SetFillBrush .GetString("brush")
                Tools_Fill.SetFillBrushColor .GetLong("brush-color")
                Tools_Fill.SetFillBrushOpacity .GetSingle("brush-opacity")
                Tools_Fill.SetFillBrushSource .GetLong("fill-source")
                Tools_Fill.SetFillCompareMode .GetLong("compare-mode")
                Tools_Fill.SetFillSampleMerged .GetBool("sample-merged")
                Tools_Fill.SetFillSearchMode .GetLong("search-mode")
                Tools_Fill.SetFillTolerance .GetSingle("tolerance")
            End With
            
        End If
        
        'Simulate a mouse click
        m_MouseDown = False
        Tools_Fill.NotifyMouseXY True, cParams.GetSingle("src-x"), cParams.GetSingle("src-y"), FormMain.MainCanvas(0)
        Tools_Fill.NotifyMouseXY False, cParams.GetSingle("src-x"), cParams.GetSingle("src-y"), FormMain.MainCanvas(0)
        
    End If

End Sub

'Render a relevant fill cursor outline to the canvas, using the stored mouse coordinates as the cursor's position
Public Sub RenderFillCursor(ByRef targetCanvas As pdCanvas)
    
    'Convert the current stored mouse coordinates from image coordinate space to viewport coordinate space
    Dim cursX As Double, cursY As Double
    Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, PDImages.GetActiveImage(), m_MouseX, m_MouseY, cursX, cursY
    
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

'Retrieve all current settings as a param string
Private Function GetAllFillSettings(ByVal srcX As Single, ByVal srcY As Single) As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'Some properties come directly from the fill object
    If EnsureFillerExists Then
    
        With cParams
            .AddParam "src-x", srcX, True
            .AddParam "src-y", srcY, True
            .AddParam "antialias", m_FloodFill.GetAntialiasingMode(), True
            .AddParam "alpha-mode", m_FillAlphaMode, True
            .AddParam "blend-mode", m_FillBlendMode, True
            .AddParam "brush", m_FillBrush, True
            .AddParam "brush-color", m_FillColor, True
            .AddParam "brush-opacity", m_FillOpacity, True
            .AddParam "fill-source", m_FillSource, True
            .AddParam "compare-mode", m_FloodFill.GetCompareMode(), True
            .AddParam "sample-merged", m_FillSampleMerged, True
            .AddParam "search-mode", m_FloodFill.GetSearchMode(), True
            .AddParam "tolerance", m_FloodFill.GetTolerance(), True
        End With
        
    End If
        
    GetAllFillSettings = cParams.GetParamString()
    
End Function

'Want to free up memory without completely releasing everything tied to this class?  That's what this function
' is for.  It should (ideally) be called whenever this tool is deactivated.
'
'Importantly, this sub does *not* touch anything that may require the underlying tool engine to be re-initialized.
' It only releases objects that the tool will auto-generate as necessary.
Public Sub ReduceMemoryIfPossible()
    If (Not m_FloodFill Is Nothing) Then m_FloodFill.FreeUpResources
    Set m_FillOutline = Nothing
    Set m_FillImage = Nothing
End Sub

'Before PD closes, you *must* call this function!  It will free any lingering fill resources (which are cached
' for performance reasons).  You can also call this function any time without penalty, if you need to free
' up memory or GDI resources.  (Freed objects are automatically recreated as-needed.)
Public Sub FreeFillResources()
    ReduceMemoryIfPossible
End Sub
