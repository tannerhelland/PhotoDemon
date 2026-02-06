Attribute VB_Name = "Tools_Clone"
'***************************************************************************
'Clone stamp tool interface
'Copyright 2019-2026 by Tanner Helland
'Created: 16/September/19
'Last updated: 29/October/19
'Last update: add support for cloning from layers with active non-destructive transforms
'
'The clone tool is nearly identical to the standard soft brush tool.  The only difference is in how the
' source overlay is calculated (e.g. instead of a solid fill, it samples from a source image/layer).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The current brush engine is stored here.  Note that this value is not correct until a call has been made to
' the CreateCurrentBrush() function; this function searches brush attributes and determines which brush engine
' to use.
Private m_BrushOutlinePath As pd2DPath

'Brush resources, used only as necessary.  Check for null values before using.
Private m_SrcPenDIB As pdDIB
Private m_Mask() As Byte, m_MaskSize As Long

'Brush attributes are stored in these variables
Private m_BrushSize As Single
Private m_BrushOpacity As Single
Private m_BrushBlendmode As PD_BlendMode
Private m_BrushAlphamode As PD_AlphaMode
Private m_BrushHardness As Single
Private m_BrushSpacing As Single
Private m_BrushFlow As Single

'If brush properties have changed since the last brush creation, this is set to FALSE.  We use this to optimize
' brush creation behavior.
Private m_BrushIsReady As Boolean
Private m_BrushCreatedAtLeastOnce As Boolean

'Current mouse/pen input values.  These are blindly relayed to us by the canvas, and it's up to us to perform any
' special tracking calculations.
Private m_MouseX As Single, m_MouseY As Single
Private Const MOUSE_OOB As Single = -9.99999E+14!

'If the shift key is being held down, we draw a different type of preview
Private m_ShiftKeyDown As Boolean

'Brush dynamics are calculated on-the-fly, and they include things like velocity, distance, angle, and more.
Private m_DistPixels As Long, m_BrushSizeInt As Long
Private m_BrushSpacingCheck As Long

'As brush movements are relayed to us, we keep a running note of the modified area of the scratch layer.
' The compositor can use this information to only regenerate the compositor cache area that's changed since the
' last repaint event.  Note that the m_ModifiedRectF may be cleared between accesses, by design - you'll need to
' keep an eye on your usage of parameters in the GetModifiedUpdateRectF function.
'
'If you want the absolute modified area since the stroke began, you can use m_TotalModifiedRectF, which is not
' cleared until the current stroke is released.
Private m_UnionRectRequired As Boolean
Private m_ModifiedRectF As RectF, m_TotalModifiedRectF As RectF

'pd2D is used for certain brush styles
Private m_Surface As pd2DSurface

'A dedicated class produces the actual dab coordinates for us, from mouse events we've forwarded to it
Private m_Paintbrush As pdPaintbrush

'If a source point exists, this will be set to TRUE.  Unlike other paint tools, the clone tool source
' does *not* reset when switching between images, which creates some complicates - namely, we must always
' check that the source image+layer combination still exists!
Private m_SourceExists As Boolean

'The user can clone from a *different* source or *different* layer!  (Note that the layer setting can
' be overridden by the Sample Merged setting)
Private m_SourceImageID As Long, m_SourceLayerID As Long
Private m_SourcePoint As PointFloat, m_SourceSetThisClick As Boolean, m_FirstStroke As Boolean
Private m_SourceOffsetX As Single, m_SourceOffsetY As Single, m_OrigSourceOffsetX As Single, m_OrigSourceOffsetY As Single
Private m_SampleMerged As Boolean, m_SampleMergedCopy As pdDIB
Private m_Sample As pdDIB, m_SampleUntouched As pdDIB
Private m_Aligned As Boolean, m_WrapMode As PD_2D_WrapMode
Private m_CtrlKeyDown As Boolean

'If the source layer is using one or more non-destructive transforms, we need to make a local copy
' of the layer with *all* transforms applied.  (This is much faster to clone.)
Private m_SourceLayerIsTransformed As Boolean, m_SourceLayerTransformed As pdDIB

'Universal brush settings, applicable for most sources.  (I say "most" because some settings can contradict each other;
' for example, a "locked" alpha mode + "erase" blend mode makes little sense, but it is technically possible to set
' those values simultaneously.)
Public Function GetBrushAligned() As Boolean
    GetBrushAligned = m_Aligned
End Function

Public Function GetBrushAlphaMode() As PD_AlphaMode
    GetBrushAlphaMode = m_BrushAlphamode
End Function

Public Function GetBrushBlendMode() As PD_BlendMode
    GetBrushBlendMode = m_BrushBlendmode
End Function

Public Function GetBrushFlow() As Single
    GetBrushFlow = m_BrushFlow
End Function

Public Function GetBrushHardness() As Single
    GetBrushHardness = m_BrushHardness
End Function

Public Function GetBrushOpacity() As Single
    GetBrushOpacity = m_BrushOpacity
End Function

Public Function GetBrushSampleMerged() As Boolean
    GetBrushSampleMerged = m_SampleMerged
End Function

Public Function GetBrushSize() As Single
    GetBrushSize = m_BrushSize
End Function

Public Function GetBrushSpacing() As Single
    GetBrushSpacing = m_BrushSpacing
End Function

Public Function GetBrushWrapMode() As PD_2D_WrapMode
    GetBrushWrapMode = m_WrapMode
End Function

'Property set functions.
Public Sub SetBrushAligned(Optional ByVal newState As Boolean = False)
    m_Aligned = newState
End Sub

Public Sub SetBrushAlphaMode(Optional ByVal newAlphaMode As PD_AlphaMode = AM_Normal)
    If (newAlphaMode <> m_BrushAlphamode) Then
        m_BrushAlphamode = newAlphaMode
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushBlendMode(Optional ByVal newBlendMode As PD_BlendMode = BM_Normal)
    If (newBlendMode <> m_BrushBlendmode) Then
        m_BrushBlendmode = newBlendMode
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushFlow(Optional ByVal newFlow As Single = 100!)
    If (newFlow <> m_BrushFlow) Then
        m_BrushFlow = newFlow
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushHardness(Optional ByVal newHardness As Single = 100!)
    newHardness = newHardness * 0.01
    If (newHardness <> m_BrushHardness) Then
        m_BrushHardness = newHardness
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushOpacity(ByVal newOpacity As Single)
    If (newOpacity <> m_BrushOpacity) Then
        m_BrushOpacity = newOpacity
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSampleMerged(ByVal newState As Boolean)
    If (newState <> m_SampleMerged) Then
        m_SampleMerged = newState
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSize(ByVal newSize As Single)
    If (newSize <> m_BrushSize) Then
        m_BrushSize = newSize
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSpacing(ByVal newSpacing As Single)
    newSpacing = newSpacing * 0.01
    If (newSpacing <> m_BrushSpacing) Then
        m_BrushSpacing = newSpacing
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushWrapMode(ByVal newMode As PD_2D_WrapMode)
    m_WrapMode = newMode
End Sub

Private Sub CreateCurrentBrush(Optional ByVal alsoCreateBrushOutline As Boolean = True, Optional ByVal forceCreation As Boolean = False)
    
    If ((Not m_BrushIsReady) Or forceCreation Or (Not m_BrushCreatedAtLeastOnce)) Then
        
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime

        'Build a new brush reference image that reflects the current brush properties
        m_BrushSizeInt = Int(m_BrushSize + 0.999999)
        CreateSoftBrushReference_PD
        
        'We also need to calculate a brush spacing reference.  A spacing of 1 means that every pixel in
        ' the current stroke is dabbed.  From a performance perspective, this is simply not feasible for
        ' large brushes, so avoid it if possible.
        '
        'The "Automatic" setting (which maps to spacing = 0) automatically calculates spacing based on
        ' the current brush size.  (Basically, we dab every 1/2pi of a radius.)
        Dim tmpBrushSpacing As Single
        tmpBrushSpacing = m_BrushSize / PI_DOUBLE
        
        If (m_BrushSpacing > 0!) Then
            tmpBrushSpacing = (m_BrushSpacing * tmpBrushSpacing)
        End If
        
        'The module-level spacing check is an integer (because we Mod it to test for paint dabs)
        m_BrushSpacingCheck = Int(tmpBrushSpacing + 0.5)
        If (m_BrushSpacingCheck < 1) Then m_BrushSpacingCheck = 1
        
        'Whenever we create a new brush, we should also refresh the current brush outline
        If alsoCreateBrushOutline Then CreateCurrentBrushOutline
        
        m_BrushIsReady = True
        m_BrushCreatedAtLeastOnce = True
        
    End If
    
End Sub

Private Sub CreateSoftBrushReference_PD()
    
    'Initialize our reference DIB as necessary
    If (m_SrcPenDIB Is Nothing) Then Set m_SrcPenDIB = New pdDIB
    If (m_SrcPenDIB.GetDIBWidth < m_BrushSizeInt) Or (m_SrcPenDIB.GetDIBHeight < m_BrushSizeInt) Then
        m_SrcPenDIB.CreateBlank m_BrushSizeInt, m_BrushSizeInt, 32, 0, 0
    Else
        m_SrcPenDIB.ResetDIB 0
    End If
    
    'PD's central brush engine will produce a reference brush shape for us
    Tools_Paint.CreateBrushMask_SolidColor m_SrcPenDIB, vbBlack, m_BrushSize, m_BrushHardness, m_BrushFlow
    
    'We now want to do something unique to the clone brush.  We don't actually need a full image for the
    ' brush source (as we're going to be producing one "on the fly" using the base image's pixels) - so instead
    ' of maintaining a full mask, just copy the relevant alpha bytes into a dedicated byte array.
    
    'Why not just produce a byte array in the first place, you ask?  Because we want this brush to produce
    ' identical border results to a standard brush, which means we want to mimic GDI+ antialiasing precisely -
    ' so we still need to lean on it for conditions like tiny brushes or 100% hardness brushes.
    If (m_MaskSize = 0) Or (m_MaskSize <> m_BrushSizeInt) Then
        m_MaskSize = m_BrushSizeInt
        ReDim m_Mask(0 To m_MaskSize - 1, 0 To m_MaskSize - 1) As Byte
    Else
        FillMemory VarPtr(m_Mask(0, 0)), m_MaskSize * m_MaskSize, 0
    End If
    
    'Finally, strip out the alpha component of the filled brush image
    Dim dstImageDataRGBA() As RGBQuad
    Dim x As Long, y As Long, tmpSA As SafeArray1D
    For y = 0 To m_MaskSize - 1
        m_SrcPenDIB.WrapRGBQuadArrayAroundScanline dstImageDataRGBA, tmpSA, y
    For x = 0 To m_MaskSize - 1
        m_Mask(x, y) = dstImageDataRGBA(x).Alpha
    Next x
    Next y
    
    m_SrcPenDIB.UnwrapRGBQuadArrayFromDIB dstImageDataRGBA

End Sub

'As part of rendering the current brush, we also need to render a brush outline onto the canvas at the current
' mouse location.  The specific outline technique used varies by brush engine.
Private Sub CreateCurrentBrushOutline()

    'TODO!  Right now this is just a copy+paste of the GDI+ outline algorithm; we obviously need a more sophisticated
    ' one in the future.
    Set m_BrushOutlinePath = New pd2DPath
    
    'Single-pixel brushes are treated as a square for cursor purposes.
    If (m_BrushSize > 0!) Then
        If (m_BrushSize = 1) Then
            m_BrushOutlinePath.AddRectangle_Absolute -0.75, -0.75, 0.75, 0.75
        Else
            m_BrushOutlinePath.AddCircle 0, 0, m_BrushSize / 2 + 0.5
        End If
    End If

End Sub

'Notify the brush engine of the current mouse position.  Coordinates should always be in *image* coordinate space,
' not screen space.  (Translation between spaces will be handled internally.)
Public Sub NotifyBrushXY(ByVal mouseButtonDown As Boolean, ByVal Shift As ShiftConstants, ByVal srcX As Single, ByVal srcY As Single, ByVal mouseTimeStamp As Long, ByRef srcCanvas As pdCanvas)
    
    'Couple things - first, determine if the CTRL key is being pressed alongside a mouse button.
    ' If it is, the user is setting a new anchor point.
    If ((Shift And vbCtrlMask) = vbCtrlMask) Then
        
        m_CtrlKeyDown = True
        m_MouseX = srcX
        m_MouseY = srcY
        
        If mouseButtonDown Then
            
            'Mark the current image, layer, and location
            m_SourceExists = True
            m_SourceImageID = PDImages.GetActiveImageID
            m_SourceLayerID = PDImages.GetActiveImage.GetActiveLayerID
            m_SourcePoint.x = srcX
            m_SourcePoint.y = srcY
            
            'Note that the brush is *not* yet ready, then immediately create the brush;
            ' this ensures the brush will be ready before the next stroke
            m_BrushIsReady = False
            CreateCurrentBrush
            
        End If
        
        m_SourceSetThisClick = True
        
        'Nothing else needs to be done here; exit immediately
        Exit Sub
        
    Else
        Message "Ctrl+Click to set clone source", "DONOTLOG"
        m_CtrlKeyDown = False
    End If
    
    'Relay this action to the brush engine; it calculates dab positions for us.
    m_Paintbrush.NotifyBrushXY mouseButtonDown, Shift, srcX, srcY, mouseTimeStamp
    
    'Reset source-set mode
    If m_Paintbrush.IsFirstDab And (Not m_CtrlKeyDown) Then
        If m_SourceSetThisClick Then m_FirstStroke = True
        m_SourceSetThisClick = False
    Else
        m_FirstStroke = False
    End If
    
    'Regardless of mouse button state (up *or* down), cache a local copy of mouse coords; we require these for
    ' rendering a brush outline.
    
    'Perform a failsafe check for brush creation
    If (Not m_BrushIsReady) Then CreateCurrentBrush
    
    'If this is a MouseDown operation, we need to make sure the full paint engine is synchronized
    ' against any property changes that are applied "on-demand" - but for this clone brush, we only
    ' do this if a valid source point has been set!
    If m_Paintbrush.IsFirstDab() And m_SourceExists Then
        
        'Switch the target canvas into high-resolution, non-auto-drop mode.  This basically means the mouse tracker
        ' reconstructs full mouse movement histories via GetMouseMovePointsEx, and it reports every last event to us,
        ' regardless of the delays involved.  (Normally, as mouse events become increasingly delayed, they are
        ' auto-dropped until the processor catches up.  We have other ways of working around that problem in the
        ' brush engine.)
        '
        'IMPORTANT NOTE: VirtualBox returns bad data via GetMouseMovePointsEx, so I now expose this setting to the user
        ' via the Tools > Options menu.  If the user disables high-res input, we will also ignore it.
        srcCanvas.SetMouseInput_HighRes Tools.GetToolSetting_HighResMouse()
        srcCanvas.SetMouseInput_AutoDrop False
        
        'Make sure the current scratch layer is properly initialized
        Tools.InitializeToolsDependentOnImage
        PDImages.GetActiveImage.ScratchLayer.SetLayerOpacity m_BrushOpacity
        PDImages.GetActiveImage.ScratchLayer.SetLayerBlendMode m_BrushBlendmode
        PDImages.GetActiveImage.ScratchLayer.SetLayerAlphaMode m_BrushAlphamode
        PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.SetInitialAlphaPremultiplicationState True
        
        'Reset the "last mouse position" values to match the current ones
        m_MouseX = srcX
        m_MouseY = srcY
        
        'Calculate an offset from the current point to the source point; this is maintained
        ' for the duration of this stroke.
        If m_FirstStroke Or (Not m_Aligned) Then
            m_SourceOffsetX = (m_SourcePoint.x - m_MouseX)
            m_SourceOffsetY = (m_SourcePoint.y - m_MouseY)
            m_OrigSourceOffsetX = m_SourceOffsetX
            m_OrigSourceOffsetY = m_SourceOffsetY
        Else
            m_SourceOffsetX = m_OrigSourceOffsetX
            m_SourceOffsetY = m_OrigSourceOffsetY
        End If
        
        'Initialize any relevant GDI+ objects for the current brush
        Drawing2D.QuickCreateSurfaceFromDC m_Surface, PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.GetDIBDC
        
        'Reset any brush dynamics that are calculated on a per-stroke basis
        m_DistPixels = 0
        
        'If "sample merged" is active, retrieve said merged sample now
        If m_SampleMerged Then
            
            'If the source image is a single-layer image, skip making a copy and instead just point the
            ' object directly at the source layer.  (The copy is *never* modified.)
            Dim mergeShortcutOK As Boolean
            mergeShortcutOK = (PDImages.GetImageByID(m_SourceImageID).GetNumOfLayers = 1)
            If mergeShortcutOK Then mergeShortcutOK = (Not PDImages.GetImageByID(m_SourceImageID).GetLayerByID(m_SourceLayerID).AffineTransformsActive(True))
            If mergeShortcutOK Then
                Set m_SampleMergedCopy = PDImages.GetImageByID(m_SourceImageID).GetLayerByID(m_SourceLayerID).GetLayerDIB
            Else
                If (m_SampleMergedCopy Is Nothing) Then Set m_SampleMergedCopy = New pdDIB
                PDImages.GetImageByID(m_SourceImageID).GetCompositedImage m_SampleMergedCopy, True
            End If
        
        'Similarly, if the source layer has one or more active non-destructive transforms, we want to
        ' cache a transformed copy now.
        Else
            
            m_SourceLayerIsTransformed = PDImages.GetImageByID(m_SourceImageID).GetLayerByID(m_SourceLayerID).AffineTransformsActive(True)
            If m_SourceLayerIsTransformed Then
                
                Dim tmpTransformDIB As pdDIB, tmpX As Long, tmpY As Long
                PDImages.GetImageByID(m_SourceImageID).GetLayerByID(m_SourceLayerID).GetAffineTransformedDIB tmpTransformDIB, tmpX, tmpY
                
                If (m_SourceLayerTransformed Is Nothing) Then Set m_SourceLayerTransformed = New pdDIB
                m_SourceLayerTransformed.CreateBlank PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 32, 0, 0
                
                GDI.BitBltWrapper m_SourceLayerTransformed.GetDIBDC, tmpX, tmpY, tmpTransformDIB.GetDIBWidth, tmpTransformDIB.GetDIBHeight, tmpTransformDIB.GetDIBDC, 0, 0, vbSrcCopy
                Set tmpTransformDIB = Nothing
                
            End If
            
        End If
        
    End If
    
    'Next, determine if the shift key is being pressed.  If it is, and if the user has already committed a
    ' brush stroke to this image (on a previous paint tool event), we want to draw a smooth line between the
    ' last paint point and the current one.  Note that this special condition is stored at module level,
    ' as we render a custom UI on mouse move events if the mouse button is *not* pressed, to help communicate
    ' what the shift key does.
    m_ShiftKeyDown = ((Shift And vbShiftMask) <> 0)
    
    Dim startTime As Currency
    
    'If the mouse button is down, perform painting between the old and new points.
    ' (All painting occurs in image coordinate space, and is applied to the current image's scratch layer.)
    If (mouseButtonDown Or m_Paintbrush.IsLastDab()) And m_SourceExists Then
    
        'Want to profile this function?  Use this line of code (and the matching report line at the bottom of the function).
        VBHacks.GetHighResTime startTime
        
        'See if there are more points in the mouse move queue.  If there are, grab them all and stroke them immediately.
        Dim numPointsRemaining As Long
        numPointsRemaining = srcCanvas.GetNumMouseEventsPending
        
        If (numPointsRemaining > 0) And (Not m_Paintbrush.IsFirstDab()) Then
        
            Dim tmpMMP As MOUSEMOVEPOINT
            Dim imgX As Double, imgY As Double
            
            Do While srcCanvas.GetNextMouseMovePoint(VarPtr(tmpMMP))
                
                'The (x, y) points returned by this request are in the *hWnd's* coordinate space.  We must manually convert them
                ' to the image coordinate space.
                If Drawing.ConvertCanvasCoordsToImageCoords(srcCanvas, PDImages.GetActiveImage(), tmpMMP.x, tmpMMP.y, imgX, imgY) Then
                    
                    'Add these points to the brush engine
                    m_Paintbrush.NotifyBrushXY True, 0, imgX, imgY, tmpMMP.ptTime
                    
                End If
                
            Loop
        
        End If
        
        'Unlike other drawing tools, the paintbrush engine controls viewport redraws.  This allows us to optimize behavior
        ' if we fall behind, and a long queue of drawing actions builds up.
        '
        '(Note that we only request manual redraws if the mouse is currently down; if the mouse *isn't* down, the canvas
        ' handles this for us.)
        Dim tmpPoint As PointFloat, numPointsDrawn As Long
        Do While m_Paintbrush.GetNextPoint(tmpPoint)
            
            'Calculate new modification rects, e.g. the portion of the paintbrush layer affected by this stroke.
            ' (The central compositor requires this information for its optimized paintbrush renderer.)
            UpdateModifiedRect tmpPoint.x, tmpPoint.y, m_Paintbrush.IsFirstDab() And (numPointsDrawn = 0)
            
            'Paint this dab
            ApplyPaintDab tmpPoint.x, tmpPoint.y
                
            'Update the "old" mouse coordinate trackers
            m_MouseX = tmpPoint.x
            m_MouseY = tmpPoint.y
            numPointsDrawn = numPointsDrawn + 1
            
        Loop
        
        'Notify the scratch layer of our updates
        PDImages.GetActiveImage.ScratchLayer.NotifyOfDestructiveChanges
        
        'Report paint tool render times, as relevant
        'Debug.Print "Paint tool render timing: " & Format$(VBHacks.GetTimerDifferenceNow(startTime) * 1000, "0000.00") & " ms"
    
    'The previous x/y coordinate trackers are updated automatically when the mouse is DOWN.  When the mouse is UP, we must manually
    ' modify those values.
    Else
        m_MouseX = srcX
        m_MouseY = srcY
    End If
    
    'If the shift key is down, we're gonna commit the paint results immediately - so don't waste time
    ' updating the screen, as it's about to be overwritten.
    If mouseButtonDown And (Shift = 0) And m_SourceExists Then UpdateViewportWhilePainting startTime, srcCanvas
    
    'If the mouse button has been released, we can also release our internal GDI+ objects.
    ' (Note that the current *brush* resources are *not* released, by design.)
    If m_Paintbrush.IsLastDab() And m_SourceExists Then
        
        Set m_Surface = Nothing
        
        'Reset the target canvas's mouse handling behavior
        srcCanvas.SetMouseInput_HighRes False
        srcCanvas.SetMouseInput_AutoDrop True
        
    End If
    
End Sub

'While painting, we use a (fairly complicated) set of heuristics to decide when to update the primary viewport.
' We don't want to update it on every paint stroke event, as compositing the full viewport can be a very
' time-consuming process (especially for large images and/or images with many layers).
Private Sub UpdateViewportWhilePainting(ByVal strokeStartTime As Currency, ByRef srcCanvas As pdCanvas)
    
    'Ask the paint engine if now is a good time to update the viewport.
    If m_Paintbrush.IsItTimeForScreenUpdate(strokeStartTime) Or m_Paintbrush.IsFirstDab() Then
    
        'Retrieve viewport parameters, then perform a full layer stack merge and repaint the screen
        Dim tmpViewportParams As PD_ViewportParams
        tmpViewportParams = Viewport.GetDefaultParamObject()
        tmpViewportParams.renderScratchLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex()
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), srcCanvas, VarPtr(tmpViewportParams)
    
    'If not enough time has passed since the last redraw, simply update the cursor
    Else
        Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), srcCanvas
    End If
    
    'Notify the paint engine that we refreshed the image; it will add this to its running fps tracker
    m_Paintbrush.NotifyScreenUpdated strokeStartTime
    
End Sub

'Apply a single paint dab to the target position.  Note that dab opacity is currently hard-coded at 100%; flow is controlled
' at brush creation time (instead of on-the-fly).  This may change depending on future brush dynamics implementations.
Private Sub ApplyPaintDab(ByVal srcX As Single, ByVal srcY As Single, Optional ByVal dabOpacity As Single = 1!)
    
    Dim allowedToDab As Boolean: allowedToDab = True
    
    'If brush dynamics are active, we only dab the brush if certain criteria are met.  (For example, if enough pixels have
    ' elapsed since the last dab, as controlled by the Brush Spacing parameter.)
    If (m_BrushSpacingCheck > 1) Then allowedToDab = ((m_DistPixels Mod m_BrushSpacingCheck) = 0)
    
    If allowedToDab Then
        
        Dim srcDIB As pdDIB
        If m_SampleMerged Then
            Set srcDIB = m_SampleMergedCopy
        Else
            If m_SourceLayerIsTransformed Then
                Set srcDIB = m_SourceLayerTransformed
            Else
                Set srcDIB = PDImages.GetImageByID(m_SourceImageID).GetLayerByID(m_SourceLayerID).GetLayerDIB
            End If
        End If
        
        'Prep the sampling DIB.  Note that it is *always* full size, regardless of the actual brush size we're gonna use
        ' (e.g. the brush may shrink because it's beyond the border of the image, but to reduce memory thrashing, we
        ' always maintain the brush DIB at a fixed size).
        If (m_Sample Is Nothing) Then Set m_Sample = New pdDIB
        If (m_Sample.GetDIBWidth <> m_BrushSizeInt) Or (m_Sample.GetDIBHeight <> m_BrushSizeInt) Then
            m_Sample.CreateBlank m_BrushSizeInt, m_BrushSizeInt, 32, 0, 0
        Else
            m_Sample.ResetDIB 0
        End If
        
        'Next, we need to calculate relevant source and destination rectangles.  GDI's AlphaBlend is incredibly picky
        ' about rectangles that don't lie off the edge of a given image, and we also get a performance boost by only
        ' masking a minimal amount of the brush.
        Dim srcRectL As RectL_WH, dstRectL As RectL_WH
        Dim dstX As Long, dstY As Long
        If CalculateSrcDstRects(Int(srcX + 0.5), Int(srcY + 0.5), dstX, dstY, srcRectL, dstRectL, srcDIB, PDImages.GetActiveImage.ScratchLayer.GetLayerDIB) Then
            
            'Retrieve the relevant portion of the source image, then make an *untouched* copy of it
            m_Sample.SetInitialAlphaPremultiplicationState True
            If (m_WrapMode = P2_WM_Clamp) Then
                
                GDI.BitBltWrapper m_Sample.GetDIBDC, dstRectL.Left, dstRectL.Top, dstRectL.Width, dstRectL.Height, srcDIB.GetDIBDC, srcRectL.Left, srcRectL.Top
                
                'The same thing can be accomplished with GDI+, but raw GDI calls tend to be faster,
                ' especially on old hardware:
                'GDI_Plus.GDIPlus_StretchBlt m_Sample, dstRectL.Left, dstRectL.Top, dstRectL.Width, dstRectL.Height, srcDIB, srcRectL.Left, srcRectL.Top, dstRectL.Width, dstRectL.Height, 1!, GP_IM_Bilinear, , , , True
                
            Else
            
                'Create a matching texture brush, using the source image as our texture
                Dim cBrush As pd2DBrush
                Set cBrush = New pd2DBrush
                cBrush.SetBrushMode P2_BM_Texture
                cBrush.SetBrushTextureFromDIB srcDIB
                cBrush.SetBrushTextureWrapMode m_WrapMode
                cBrush.CreateBrush
                
                'Create a transformation matrix that ensure the source texture offset is in the correct
                ' position for this brush stroke.
                Dim cTransform As pd2DTransform
                Set cTransform = New pd2DTransform
                cTransform.ApplyTranslation -1 * srcRectL.Left, -1 * srcRectL.Top
                cBrush.SetBrushTextureTransform cTransform
                
                'Finally, paint the texture onto the brush image
                Dim dstSurface As pd2DSurface
                Drawing2D.QuickCreateSurfaceFromDIB dstSurface, m_Sample, False
                PD2D.FillRectangleF dstSurface, cBrush, dstRectL.Left, dstRectL.Top, dstRectL.Width, dstRectL.Height
                
                Set cBrush = Nothing: Set cTransform = Nothing: Set dstSurface = Nothing
                
            End If
            
            Set srcDIB = Nothing
            
            If (m_SampleUntouched Is Nothing) Then Set m_SampleUntouched = New pdDIB
            m_SampleUntouched.CreateFromExistingDIB m_Sample
            
            'Mask the outline of the current brush over the source image.
            Dim pxSample() As Byte
            Dim sampleSA As SafeArray1D, dstSA As SafeArray1D
            
            Dim x As Long, y As Long, xStride As Long, tmpFloat As Single
            Dim fLookup(0 To 255) As Single
            
            For x = 0 To 255
                fLookup(x) = CSng(x) / 255!
            Next x
            
            Dim xStart As Long, xEnd As Long, yStart As Long, yEnd As Long, srcMaskByte As Byte
            xStart = dstRectL.Left
            xEnd = dstRectL.Left + dstRectL.Width - 1
            yStart = dstRectL.Top
            yEnd = dstRectL.Top + dstRectL.Height - 1
            
            Dim dstPtr As Long, dstStride As Long
            m_Sample.WrapArrayAroundScanline pxSample, sampleSA, 0
            dstPtr = m_Sample.GetDIBPointer
            dstStride = m_Sample.GetDIBStride
            
            For y = yStart To yEnd
                sampleSA.pvData = dstPtr + (y * dstStride)
            For x = xStart To xEnd
            
                xStride = x * 4
                srcMaskByte = m_Mask(x, y)
                
                'Because large chunks of the brush will always be transparent (e.g. outside the brush circle)
                ' or solid (e.g. inside the brush hardness radius), we can shortcut this inner loop by
                ' checking 0/255 values before performing floating-point math.  Profiling showed performance
                ' improvements of ~200% for a brush with hardness 50, and even larger gains as brush
                ' hardness increases.
                If (srcMaskByte < 255) Then
                    If (srcMaskByte = 0) Then
                        pxSample(xStride) = 0
                        pxSample(xStride + 1) = 0
                        pxSample(xStride + 2) = 0
                        pxSample(xStride + 3) = 0
                    Else
                        tmpFloat = fLookup(srcMaskByte)
                        pxSample(xStride) = pxSample(xStride) * tmpFloat
                        pxSample(xStride + 1) = pxSample(xStride + 1) * tmpFloat
                        pxSample(xStride + 2) = pxSample(xStride + 2) * tmpFloat
                        pxSample(xStride + 3) = pxSample(xStride + 3) * tmpFloat
                    End If
                End If
                
            Next x
            Next y
            
            m_Sample.UnwrapArrayFromDIB pxSample
            
            'Apply the dab
            Dim dstDIB As pdDIB
            Set dstDIB = PDImages.GetActiveImage.ScratchLayer.GetLayerDIB
            m_Sample.AlphaBlendToDCEx dstDIB.GetDIBDC, dstX, dstY, dstRectL.Width, dstRectL.Height, dstRectL.Left, dstRectL.Top, dstRectL.Width, dstRectL.Height, dabOpacity * 255
            
            'We now need to do something special for semi-transparent pixels.  These pixels *cannot* be allowed
            ' to become more transparent than the source pixel data (otherwise it wouldn't be a clone operation).
            ' As such, we need to scan the pixels we just painted, and ensure that they do not exceed their
            ' original transparency values.
            
            'Why not just perform the alpha-blend ourselves and do this as we go?  Because for fully opaque clones
            ' (e.g. photos!), a GDI AlphaBlend is hardware-accelerated - much faster than we can possibly blend -
            ' while this secondary step is extremely fast as the CPU can branch-predict it with 100% accuracy.
            ' So nearly no time is lost compared to a regular alpha blend op for the most common use-case.
            
            'Normally we would need to do some messy boundary checks here, but this was already handled by the
            ' CalculateSrcDstRects() function, above.
            Dim pxSampleL() As RGBQuad, pxDstL() As RGBQuad
            Dim refAlpha As Long, testAlpha As Long
            
            Dim xOffset As Long, yOffset As Long
            xOffset = (dstX - dstRectL.Left)
            yOffset = dstY - dstRectL.Top
            
            xStart = dstRectL.Left
            xEnd = (dstRectL.Left + dstRectL.Width - 1)
            yStart = dstRectL.Top
            yEnd = dstRectL.Top + dstRectL.Height - 1
            
            For y = yStart To yEnd
                m_SampleUntouched.WrapRGBQuadArrayAroundScanline pxSampleL, sampleSA, y
                dstDIB.WrapRGBQuadArrayAroundScanline pxDstL, dstSA, yOffset + y
            For x = xStart To xEnd
                
                'Retrieve our "reference" alpha values from the sample
                refAlpha = pxSampleL(x).Alpha
                
                'Ignore opaque pixels
                If (refAlpha < 255) Then
                
                    'Is the destination alpha higher than the original source's alpha?
                    ' If it is, clone the destination pixel in its place.
                    testAlpha = pxDstL(xOffset + x).Alpha
                    If (testAlpha > refAlpha) Then pxDstL(xOffset + x) = pxSampleL(x)
                    
                End If
            
            Next x
            Next y
            
            'Free array references
            m_SampleUntouched.UnwrapRGBQuadArrayFromDIB pxSampleL
            dstDIB.UnwrapRGBQuadArrayFromDIB pxDstL
            
        '/end clone regions are valid
        End If
        
    End If
    
    'Each time we make a new dab, we keep a running tally of how many pixels we've traversed.  Some brush dynamics (e.g. spacing)
    ' rely on this value for correct rendering behavior.
    m_DistPixels = m_DistPixels + 1
    
End Sub

'Determine source+dest blends for the current clone region.  Returns TRUE if the region is non-zero; FALSE otherwise.
' (Do *NOT* waste time rendering a dab if the return value is FALSE.)
Private Function CalculateSrcDstRects(ByVal srcX As Long, ByVal srcY As Long, ByRef dstX As Long, ByRef dstY As Long, ByRef srcRectL As RectL_WH, ByRef dstRectL As RectL_WH, ByRef srcDIB As pdDIB, ByRef scratchLayerDIB As pdDIB) As Boolean

    'Start by populating both rects with default values
    With srcRectL
        .Left = Int(srcX - m_BrushSizeInt \ 2 + m_SourceOffsetX)
        .Top = Int(srcY - m_BrushSizeInt \ 2 + m_SourceOffsetY)
        .Width = m_BrushSizeInt
        .Height = m_BrushSizeInt
    End With
    
    'If sampling directly from a source layer (e.g. sampling merged is NOT set), we need to adjust
    ' our source rectangle to account for the source layer's potential x/y offsets in the image.
    If (Not m_SampleMerged) And (Not m_SourceLayerIsTransformed) Then
        With srcRectL
            .Left = .Left - PDImages.GetImageByID(m_SourceImageID).GetLayerByID(m_SourceLayerID).GetLayerOffsetX
            .Top = .Top - PDImages.GetImageByID(m_SourceImageID).GetLayerByID(m_SourceLayerID).GetLayerOffsetY
        End With
    End If
    
    With dstRectL
        .Left = 0
        .Top = 0
        'dstRectL width and height will *always* be identical to srcRectL width and height; as such,
        ' we don't populate them until the end of this function
    End With
    
    'Next, calculate overlap.  Note that source overlap is *only* calculated if the current wrap mode
    ' is set to NONE.  (Otherwise, we'll wrap the source at boundaries, so we don't want to crop the rect.)
    If (m_WrapMode = P2_WM_Clamp) Then
        
        'Next, perform boundary checks on the source rectangle, and modify *both* rectangles to account for any changes
        Dim tmpOffset As Long
        If (srcRectL.Left < 0) Then
            tmpOffset = srcRectL.Left
            srcRectL.Left = 0
            dstRectL.Left = dstRectL.Left - tmpOffset
            srcRectL.Width = srcRectL.Width + tmpOffset
        End If
        
        If (srcRectL.Top < 0) Then
            tmpOffset = srcRectL.Top
            srcRectL.Top = 0
            dstRectL.Top = dstRectL.Top - tmpOffset
            srcRectL.Height = srcRectL.Height + tmpOffset
        End If
        
        If (srcRectL.Left + srcRectL.Width > srcDIB.GetDIBWidth) Then
            tmpOffset = (srcRectL.Left + srcRectL.Width) - srcDIB.GetDIBWidth
            srcRectL.Width = srcRectL.Width - tmpOffset
        End If
        
        If (srcRectL.Top + srcRectL.Height > srcDIB.GetDIBHeight) Then
            tmpOffset = (srcRectL.Top + srcRectL.Height) - srcDIB.GetDIBHeight
            srcRectL.Height = srcRectL.Height - tmpOffset
        End If
        
    End If
    
    'As a convenience, let's also calculate out-of-bounds destination pixels (which use scratch layer boundaries)
    
    'Determine where an "ideal" dab will be placed
    dstX = Int(srcX - m_BrushSize \ 2) + dstRectL.Left
    dstY = Int(srcY - m_BrushSize \ 2) + dstRectL.Top
        
    If (dstX < 0) Then
        tmpOffset = dstX
        dstX = 0
        dstRectL.Left = dstRectL.Left - tmpOffset
        
        'If we are wrapping the source texture (e.g. treating it like a pattern), we do *not* want to modify
        ' the source left value - it will be automatically handled correctly, according to the current wrap mode.
        If (m_WrapMode = P2_WM_Clamp) Then srcRectL.Left = srcRectL.Left - tmpOffset
        srcRectL.Width = srcRectL.Width + tmpOffset
    End If
    
    If (dstY < 0) Then
        tmpOffset = dstY
        dstY = 0
        dstRectL.Top = dstRectL.Top - tmpOffset
        
        'See the previous wrap mode note for an explanation of this If/Then statement
        If (m_WrapMode = P2_WM_Clamp) Then srcRectL.Top = srcRectL.Top - tmpOffset
        srcRectL.Height = srcRectL.Height + tmpOffset
    End If
    
    If (dstX + srcRectL.Width > scratchLayerDIB.GetDIBWidth) Then
        tmpOffset = (dstX + srcRectL.Width) - scratchLayerDIB.GetDIBWidth
        srcRectL.Width = srcRectL.Width - tmpOffset
    End If
    
    If (dstY + srcRectL.Height > scratchLayerDIB.GetDIBHeight) Then
        tmpOffset = (dstY + srcRectL.Height) - scratchLayerDIB.GetDIBHeight
        srcRectL.Height = srcRectL.Height - tmpOffset
    End If
    
    'Mirror the source width/height to the destination
    dstRectL.Width = srcRectL.Width
    dstRectL.Height = srcRectL.Height
    
    'Rects with sub-zero (or zero) dimensions are invalid, and we can skip painting them entirely
    CalculateSrcDstRects = (srcRectL.Width > 0) And (srcRectL.Height > 0)
    
End Function

'Whenever we receive notifications of a new mouse (x, y) pair, you need to call this sub to calculate a new "affected area" rect.
' The compositor uses this "affected area" rect to minimize the amount of rendering work it needs to perform.
Private Sub UpdateModifiedRect(ByVal newX As Single, ByVal newY As Single, ByVal isFirstStroke As Boolean)

    'Start by calculating the affected rect for just this stroke.
    Dim tmpRectF As RectF
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
    
    'Inflate the rect calculation by the size of the current brush, while accounting for the possibility of antialiasing
    ' (which may extend up to 1.0 pixel outside the calculated boundary area).
    Dim halfBrushSize As Single
    halfBrushSize = m_BrushSize / 2! + 1!
    
    tmpRectF.Left = tmpRectF.Left - halfBrushSize
    tmpRectF.Top = tmpRectF.Top - halfBrushSize
    
    halfBrushSize = halfBrushSize * 2
    tmpRectF.Width = tmpRectF.Width + halfBrushSize
    tmpRectF.Height = tmpRectF.Height + halfBrushSize
    
    Dim tmpOldRectF As RectF
    
    'Normally, we union the current rect against our previous (running) modified rect.
    ' Two circumstances prevent this, however:
    ' 1) This is the first dab in a stroke (so there is no running modification rect)
    ' 2) The compositor just retrieved our running modification rect, and updated the screen accordingly.
    '    This means we can start a new rect instead.
    'If this is *not* the first modified rect calculation, union this rect with our previous update rect
    If m_UnionRectRequired And (Not isFirstStroke) Then
        tmpOldRectF = m_ModifiedRectF
        PDMath.UnionRectF m_ModifiedRectF, tmpRectF, tmpOldRectF
    Else
        m_UnionRectRequired = True
        m_ModifiedRectF = tmpRectF
    End If
    
    'Always calculate a running "total combined RectF", for use in the final merge step
    If isFirstStroke Then
        m_TotalModifiedRectF = tmpRectF
    Else
        tmpOldRectF = m_TotalModifiedRectF
        PDMath.UnionRectF m_TotalModifiedRectF, tmpRectF, tmpOldRectF
    End If
    
End Sub

'If the source image+layer combination doesn't exist, this sub will reset m_SourceExists to FALSE
Private Sub EnsureSourceExists()

    If m_SourceExists Then
    
        If (Not PDImages.IsImageActive(m_SourceImageID)) Then m_SourceExists = False
        If m_SourceExists And (Not m_SampleMerged) Then
            If (PDImages.GetImageByID(m_SourceImageID).GetLayerByID(m_SourceLayerID) Is Nothing) Then m_SourceExists = False
        End If
    
    End If
    
    'This is also a convenient place to reset anything related to the source existing
    If (Not m_SourceExists) Then m_FirstStroke = False
    
End Sub

'When the active image changes, we need to reset certain brush-related parameters
Public Sub NotifyActiveImageChanged()
    
    m_Paintbrush.Reset
    
    m_MouseX = MOUSE_OOB
    m_MouseY = MOUSE_OOB
    
    'Make sure our source point hasn't disappeared (e.g. been unloaded)
    EnsureSourceExists
    
End Sub

'When image size changes (via not just resize but rotation, crop, etc) we need to clear the source point,
' as it may no longer be valid.
Public Sub NotifyImageSizeChanged()
    m_SourceExists = False
    NotifyActiveImageChanged
End Sub

'Return the area of the image modified by the current stroke.
' IMPORTANTLY: the running modified rect is FORCIBLY RESET after a call to this function, by design.
' (After PD's compositor retrieves the modification rect, everything inside that rect will get updated -
'  so we can start our next batch of modifications afresh.)
Public Function GetModifiedUpdateRectF() As RectF
    GetModifiedUpdateRectF = m_ModifiedRectF
    m_UnionRectRequired = False
End Function

Public Function IsFirstDab() As Boolean
    If (m_Paintbrush Is Nothing) Then IsFirstDab = False Else IsFirstDab = m_Paintbrush.IsFirstDab()
End Function

'Want to commit your current brush work?  Call this function to make the brush results permanent.
Public Sub CommitBrushResults()
    
    'Check ctrl key status and skip this step accordingly
    If m_SourceSetThisClick Then Exit Sub
    
    'This dummy string only exists to ensure that the processor name gets localized properly
    ' (as that text is used for Undo/Redo descriptions).  PD's translation engine will detect
    ' the TranslateMessage() call and produce a matching translation entry.
    Dim strDummy As String
    strDummy = g_Language.TranslateMessage("Clone stamp")
    Layers.CommitScratchLayer "Clone stamp", m_TotalModifiedRectF
    
End Sub

'Render the current brush outline to the canvas, using the stored mouse coordinates as the brush's position.
' (As of August 2022, Caps Lock can be used to toggle between precision and outline modes; this mimics Photoshop.
'  See https://github.com/tannerhelland/PhotoDemon/issues/425 for details.)
Public Sub RenderBrushOutline(ByRef targetCanvas As pdCanvas)
    
    'If a brush outline doesn't exist, create one now
    If (Not m_BrushIsReady) Then CreateCurrentBrush True
    
    'Ensure the source point (if any) exists
    EnsureSourceExists
    
    'If the on-screen brush size is above a certain threshold, we'll paint a full brush outline.
    ' If it's too small, we'll only paint a cross in the current brush position.
    Dim onScreenSize As Double
    onScreenSize = Drawing.ConvertImageSizeToCanvasSize(m_BrushSize, PDImages.GetActiveImage())
    
    Dim brushTooSmall As Boolean
    brushTooSmall = (onScreenSize < 7#)
    
    'Like Photoshop, the CAPS LOCK key can be used to toggle between brush outlines and "precision" cursor mode.
    ' In "precision" mode, we only draw a target cursor.
    Dim renderInPrecisionMode As Boolean
    renderInPrecisionMode = brushTooSmall Or OS.IsVirtualKeyDown_Synchronous(VK_CAPITAL, True)
    
    'Borrow a pair of UI pens from the main rendering module
    Dim innerPen As pd2DPen, outerPen As pd2DPen
    Drawing.BorrowCachedUIPens outerPen, innerPen
    
    'Create other required pd2D drawing tools (a surface)
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, targetCanvas.hDC, True
    cSurface.SetSurfacePixelOffset P2_PO_Normal
    
    'Other misc drawing tools
    Dim canvasMatrix As pd2DTransform
    Dim srcX As Double, srcY As Double
    Dim cursX As Double, cursY As Double
    Dim oldX As Double, oldY As Double
    Dim lastPoint As PointFloat
    Dim crossLength As Single, crossDistanceFromCenter As Single, outerCrossBorder As Single
    Dim copyOfBrushOutline As pd2DPath
    Dim backupWidth As Single
    Dim okToProceed As Boolean
    
    'Dash sizes (for the source point outline)
    Dim dashSizes(0 To 1) As Single
    dashSizes(0) = 2.5!
    dashSizes(1) = 2.5!
    
    'If the user is currently holding down the ctrl key, they're trying to set a source point.
    ' Render a special outline just for this occasion.
    If m_CtrlKeyDown Then
        
        srcX = m_MouseX
        srcY = m_MouseY
        
        'Same steps as normal rendering; skip down to see detailed comments on what these lines do
        Set canvasMatrix = Nothing
        Drawing.GetTransformFromImageToCanvas canvasMatrix, targetCanvas, PDImages.GetActiveImage(), srcX, srcY
        Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, PDImages.GetActiveImage(), srcX, srcY, cursX, cursY
        
        If renderInPrecisionMode Then
            
            outerPen.SetPenLineCap P2_LC_Round
            innerPen.SetPenLineCap P2_LC_Round
            
            'Paint a target cursor
            crossLength = 3!
            crossDistanceFromCenter = 4!
            outerCrossBorder = 0.25!
            
            'Four "beneath" shadows
            PD2D.DrawLineF cSurface, outerPen, cursX, cursY - crossDistanceFromCenter + outerCrossBorder, cursX, cursY - crossDistanceFromCenter - crossLength - outerCrossBorder
            PD2D.DrawLineF cSurface, outerPen, cursX, cursY + crossDistanceFromCenter - outerCrossBorder, cursX, cursY + crossDistanceFromCenter + crossLength + outerCrossBorder
            PD2D.DrawLineF cSurface, outerPen, cursX - crossDistanceFromCenter + outerCrossBorder, cursY, cursX - crossDistanceFromCenter - crossLength - outerCrossBorder, cursY
            PD2D.DrawLineF cSurface, outerPen, cursX + crossDistanceFromCenter - outerCrossBorder, cursY, cursX + crossDistanceFromCenter + crossLength + outerCrossBorder, cursY
            
            'Four "above" opaque lines
            PD2D.DrawLineF cSurface, innerPen, cursX, cursY - crossDistanceFromCenter, cursX, cursY - crossDistanceFromCenter - crossLength
            PD2D.DrawLineF cSurface, innerPen, cursX, cursY + crossDistanceFromCenter, cursX, cursY + crossDistanceFromCenter + crossLength
            PD2D.DrawLineF cSurface, innerPen, cursX - crossDistanceFromCenter, cursY, cursX - crossDistanceFromCenter - crossLength, cursY
            PD2D.DrawLineF cSurface, innerPen, cursX + crossDistanceFromCenter, cursY, cursX + crossDistanceFromCenter + crossLength, cursY
            
        'If size and settings allow, render a transformed brush outline onto the canvas
        Else
            
            'Get a copy of the current brush outline, transformed into position
            Set copyOfBrushOutline = New pd2DPath
            copyOfBrushOutline.CloneExistingPath m_BrushOutlinePath
            copyOfBrushOutline.ApplyTransformation canvasMatrix
    
            backupWidth = outerPen.GetPenWidth
            outerPen.SetPenWidth 1.6
            
            innerPen.SetPenStyle P2_DS_Custom
            innerPen.SetPenDashCap P2_DC_Round
            innerPen.SetPenDashes_UNSAFE VarPtr(dashSizes(0)), 2
            
            PD2D.DrawPath cSurface, outerPen, copyOfBrushOutline
            PD2D.DrawPath cSurface, innerPen, copyOfBrushOutline
            
            innerPen.SetPenStyle P2_DS_Solid
            outerPen.SetPenWidth backupWidth
            
        End If
        
    'If the ctrl key is *not* down, rendering is more complicated, as it depends on a bunch of
    ' factors (is a source point set? is a paint stroke active? is shift being used? etc)
    Else
        
        'We now want to (potentially) draw *two* sets of brush outlines - one for the brush location,
        ' and a second one for the source location, if it exists.  We also want the cursor for the active
        ' brush position (i = 0) to "overlay" the indicator for the source position (i = 1), which is
        ' why we draw the points in reverse
        Dim i As Long
        For i = 1 To 0 Step -1
        
            'Skip all steps for the source point if it doesn't exist
            If (i = 1) And (Not m_SourceExists) Then GoTo DrawNextPoint
            If (i = 1) And (m_SourceImageID <> PDImages.GetActiveImageID) Then GoTo DrawNextPoint
            
            'Set the relevant source point (i = 0, use cursor position; i = 1, use source point)
            If (i = 0) Then
                srcX = m_MouseX
                srcY = m_MouseY
            Else
                If m_Paintbrush.IsMouseDown() Then
                    srcX = m_MouseX + m_SourceOffsetX
                    srcY = m_MouseY + m_SourceOffsetY
                Else
                    If m_Aligned And (Not m_SourceSetThisClick) Then
                        srcX = m_MouseX + m_SourceOffsetX
                        srcY = m_MouseY + m_SourceOffsetY
                    Else
                        srcX = m_SourcePoint.x
                        srcY = m_SourcePoint.y
                    End If
                End If
            End If
            
            'Start by creating a transformation from the image space to the canvas space
            Set canvasMatrix = Nothing
            Drawing.GetTransformFromImageToCanvas canvasMatrix, targetCanvas, PDImages.GetActiveImage(), srcX, srcY
            
            'We also want to pinpoint the precise cursor position
            Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, PDImages.GetActiveImage(), srcX, srcY, cursX, cursY
            
            'If the user is holding down the SHIFT key, paint a line between the end of the previous stroke and the current
            ' mouse position.  This helps communicate that shift+clicking will string together separate strokes.
            If (i = 0) Then okToProceed = m_ShiftKeyDown And m_Paintbrush.GetLastAddedPoint(lastPoint)
            If okToProceed Then
                
                outerPen.SetPenLineCap P2_LC_Round
                innerPen.SetPenLineCap P2_LC_Round
                
                If (i = 0) Then
                    Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, PDImages.GetActiveImage(), lastPoint.x, lastPoint.y, oldX, oldY
                    PD2D.DrawLineF cSurface, outerPen, oldX, oldY, cursX, cursY
                    PD2D.DrawLineF cSurface, innerPen, oldX, oldY, cursX, cursY
                Else
                    lastPoint.x = m_SourcePoint.x + (m_MouseX - lastPoint.x)
                    lastPoint.y = m_SourcePoint.y + (m_MouseY - lastPoint.y)
                    Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, PDImages.GetActiveImage(), lastPoint.x, lastPoint.y, oldX, oldY
                    PD2D.DrawLineF cSurface, outerPen, oldX, oldY, cursX, cursY
                    PD2D.DrawLineF cSurface, innerPen, oldX, oldY, cursX, cursY
                End If
                
            Else
                
                'In precision mode (CAPS LOCK down or zoomed out too far to render a proper brush outline),
                ' we want to *always* paint a target cursor at the source location, but we only paint a target
                ' cursor at the current location *if* the mouse is not down.  (This prevents obscuring the
                ' location being painted.)
                crossLength = 3!
                crossDistanceFromCenter = 4!
                outerCrossBorder = 0.25!
                
                If renderInPrecisionMode Then
                    
                    'For the active cursor, only paint a target cross when the mouse is *not* down
                    If (i = 0) Then
                        okToProceed = (Not m_Paintbrush.IsMouseDown())
                    
                    'For the source cursor, *always* paint a target cross
                    Else
                        okToProceed = True
                    End If
                    
                    outerPen.SetPenLineCap P2_LC_Round
                    innerPen.SetPenLineCap P2_LC_Round
                    
                    'Four "beneath" shadows
                    PD2D.DrawLineF cSurface, outerPen, cursX, cursY - crossDistanceFromCenter + outerCrossBorder, cursX, cursY - crossDistanceFromCenter - crossLength - outerCrossBorder
                    PD2D.DrawLineF cSurface, outerPen, cursX, cursY + crossDistanceFromCenter - outerCrossBorder, cursX, cursY + crossDistanceFromCenter + crossLength + outerCrossBorder
                    PD2D.DrawLineF cSurface, outerPen, cursX - crossDistanceFromCenter + outerCrossBorder, cursY, cursX - crossDistanceFromCenter - crossLength - outerCrossBorder, cursY
                    PD2D.DrawLineF cSurface, outerPen, cursX + crossDistanceFromCenter - outerCrossBorder, cursY, cursX + crossDistanceFromCenter + crossLength + outerCrossBorder, cursY
                    
                    'Four "above" opaque lines
                    PD2D.DrawLineF cSurface, innerPen, cursX, cursY - crossDistanceFromCenter, cursX, cursY - crossDistanceFromCenter - crossLength
                    PD2D.DrawLineF cSurface, innerPen, cursX, cursY + crossDistanceFromCenter, cursX, cursY + crossDistanceFromCenter + crossLength
                    PD2D.DrawLineF cSurface, innerPen, cursX - crossDistanceFromCenter, cursY, cursX - crossDistanceFromCenter - crossLength, cursY
                    PD2D.DrawLineF cSurface, innerPen, cursX + crossDistanceFromCenter, cursY, cursX + crossDistanceFromCenter + crossLength, cursY
                    
                End If
                
            End If
            
            'If size allows, render a transformed brush outline onto the canvas as well
            If (Not renderInPrecisionMode) Then
                
                'Get a copy of the current brush outline, transformed into position
                Set copyOfBrushOutline = New pd2DPath
                copyOfBrushOutline.CloneExistingPath m_BrushOutlinePath
                copyOfBrushOutline.ApplyTransformation canvasMatrix
                
                If (i = 1) Then
                    
                    backupWidth = outerPen.GetPenWidth
                    outerPen.SetPenWidth 1.6
                    
                    innerPen.SetPenStyle P2_DS_Custom
                    innerPen.SetPenDashCap P2_DC_Round
                    innerPen.SetPenDashes_UNSAFE VarPtr(dashSizes(0)), 2
                    
                    PD2D.DrawPath cSurface, outerPen, copyOfBrushOutline
                    PD2D.DrawPath cSurface, innerPen, copyOfBrushOutline
                    
                    innerPen.SetPenStyle P2_DS_Solid
                    outerPen.SetPenWidth backupWidth
                    
                Else
                    PD2D.DrawPath cSurface, outerPen, copyOfBrushOutline
                    PD2D.DrawPath cSurface, innerPen, copyOfBrushOutline
                End If
                
            End If
    
DrawNextPoint:
        Next i
        
    End If
    
    Set cSurface = Nothing
    
End Sub

'Any specialized initialization tasks can be handled here.  This function is called early in the PD load process.
Public Sub InitializeBrushEngine()
    
    'Initialize the underlying brush class
    Set m_Paintbrush = New pdPaintbrush
    
    'Reset all coordinates
    m_MouseX = MOUSE_OOB
    m_MouseY = MOUSE_OOB
    
    'Note that the current brush has *not* been created yet!
    m_BrushIsReady = False
    m_BrushCreatedAtLeastOnce = False
    
    'Wrap mode is now available, effectively making this a clone stamp *and* pattern stamp brush!
    m_WrapMode = P2_WM_Clamp
    
    'Flow and spacing are *not* currently available to the user (in the tool UI).  They may be restored
    ' in a future update, and as such, no work is required to integrate them - the values are "ready to go".
    ' For now, we just set them to default values.
    m_BrushSpacing = 0!
    m_BrushFlow = 100!
    
End Sub

'Want to free up memory without completely releasing everything tied to this class?  That's what this function
' is for.  It should (ideally) be called whenever this tool is deactivated.
'
'Importantly, this sub does *not* touch anything that may require the underlying tool engine to be re-initialized.
' It only releases objects that the tool will auto-generate as necessary.
Public Sub ReduceMemoryIfPossible()
    
    Set m_BrushOutlinePath = Nothing
    
    'When freeing the underlying brush, we also need to reset its creation flags
    ' (to ensure it gets re-created correctly)
    m_BrushIsReady = False
    m_BrushCreatedAtLeastOnce = False
    Set m_SrcPenDIB = Nothing
    
    m_MaskSize = 0
    Erase m_Mask
    
    Set m_SampleMergedCopy = Nothing
    Set m_Sample = Nothing
    Set m_SampleUntouched = Nothing
    
    'While we're here, remove the source point
    m_SourceExists = False
    
End Sub

'Before PD closes, you *must* call this function!  It will free any lingering brush resources (which are cached
' for performance reasons).
Public Sub FreeBrushResources()
    ReduceMemoryIfPossible
End Sub
