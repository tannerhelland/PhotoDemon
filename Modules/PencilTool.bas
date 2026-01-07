Attribute VB_Name = "Tools_Pencil"
'***************************************************************************
'Pencil tool interface
'Copyright 2016-2026 by Tanner Helland
'Created: 1/November/16
'Last updated: 29/January/25
'Last update: add "align to pixel grid" setting, and allow the user to toggle at their leisure.
'             (When enabled, all pen strokes are forcibly centered against the pixel grid, for "perfect" precision.)
'
'PD's pencil tool is just a thin wrapper around standard GDI+ pens.  This makes it fast but somewhat quirky
' to code against compared to PD's paintbrush tools (which use a custom brush engine).
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
Private m_GDIPPen As pd2DPen

'Brush attributes are stored in these variables
Private m_BrushSize As Single
Private m_BrushOpacity As Single
Private m_BrushBlendmode As PD_BlendMode
Private m_BrushAlphamode As PD_AlphaMode
Private m_BrushAntialiasing As PD_2D_Antialiasing

'Note that some brush attributes only exist for certain brush sources.
Private m_BrushColor As Long

'In 2025, a new option was added for strictly snapping the brush to pixel centers (see https://github.com/tannerhelland/PhotoDemon/discussions/635).
' When this option is toggled OFF, paint tools behave like Photoshop or Paint.NET.  Turning it ON makes it more acceptable
' for e.g. pixel art, with strict precision (particularly with 1px brush sizes).
Private m_StrictPixelCentering As Boolean

'If brush properties have changed since the last brush creation, this is set to FALSE.
' (We use this to optimize brush creation behavior.)
Private m_BrushIsReady As Boolean

'Current mouse/pen input values.  These are blindly relayed to us by the canvas, and it's up to us to perform any
' special tracking calculations.
Private m_MouseDown As Boolean
Private m_MouseX As Single, m_MouseY As Single
Private m_MouseLastUserX As Single, m_MouseLastUserY As Single
Private Const MOUSE_OOB As Single = -9.99999E+14!
Private m_MouseShiftEventOK As Boolean
Private m_isFirstStroke As Boolean, m_isLastStroke As Boolean

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

'PD's brush engine handles some tasks for us (like calculating FPS for UI updates)
Private m_Paintbrush As pdPaintbrush

'Universal brush settings, applicable for most sources.  (I say "most" because some settings can contradict each other;
' for example, a "locked" alpha mode + "erase" blend mode makes little sense, but it is technically possible to set
' those values simultaneously.)
Public Function GetBrushAlphaMode() As PD_AlphaMode
    GetBrushAlphaMode = m_BrushAlphamode
End Function

Public Function GetBrushAntialiasing() As PD_2D_Antialiasing
    GetBrushAntialiasing = m_BrushAntialiasing
End Function

Public Function GetBrushBlendMode() As PD_BlendMode
    GetBrushBlendMode = m_BrushBlendmode
End Function

Public Function GetBrushOpacity() As Single
    GetBrushOpacity = m_BrushOpacity
End Function

Public Function GetBrushSize() As Single
    GetBrushSize = m_BrushSize
End Function

Public Function GetBrushColor() As Long
    GetBrushColor = m_BrushColor
End Function

Public Function GetStrictPixelAlignment() As Boolean
    GetStrictPixelAlignment = m_StrictPixelCentering
End Function

'Property set functions.  Note that not all brush properties are used by all styles.
' (e.g. "brush hardness" is not used by "pencil" style brushes, etc)
Public Sub SetBrushAlphaMode(Optional ByVal newAlphaMode As PD_AlphaMode = AM_Normal)
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

Public Sub SetBrushBlendMode(Optional ByVal newBlendMode As PD_BlendMode = BM_Normal)
    If (newBlendMode <> m_BrushBlendmode) Then
        m_BrushBlendmode = newBlendMode
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushOpacity(ByVal newOpacity As Single)
    If (newOpacity <> m_BrushOpacity) Then
        m_BrushOpacity = newOpacity
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSize(ByVal newSize As Single)
    If (newSize <> m_BrushSize) Then
        m_BrushSize = newSize
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushColor(Optional ByVal newColor As Long = vbWhite)
    If (newColor <> m_BrushColor) Then
        m_BrushColor = newColor
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetStrictPixelAlignment(ByVal newValue As Boolean)
    If (newValue <> m_StrictPixelCentering) Then
        m_StrictPixelCentering = newValue
        m_BrushIsReady = False
    End If
End Sub

Private Sub CreateCurrentBrush(Optional ByVal alsoCreateBrushOutline As Boolean = True, Optional ByVal forceCreation As Boolean = False)
        
    If ((Not m_BrushIsReady) Or forceCreation) Then
    
        'For now, create a circular pen at the current size
        If (m_GDIPPen Is Nothing) Then Set m_GDIPPen = New pd2DPen
        Drawing2D.QuickCreateSolidPen m_GDIPPen, m_BrushSize, m_BrushColor, 100!, P2_LJ_Round, P2_LC_Round
        
        'Whenever we create a new brush, we should also refresh the current brush outline
        If alsoCreateBrushOutline Then CreateCurrentBrushOutline
        
        m_BrushIsReady = True
        
    End If
    
End Sub

'As part of rendering the current brush, we also need to render a brush outline onto the canvas at the current
' mouse location.  The specific outline technique used varies by brush engine.
Private Sub CreateCurrentBrushOutline()
        
    'If this is a GDI+ brush, outline creation is pretty easy.  Assume a circular brush and simply
    ' create a path at that same size.  (Note that circles are defined by radius, while brushes are
    ' defined by diameter - hence the "/ 2".)
    Set m_BrushOutlinePath = New pd2DPath
    
    'Single-pixel brushes are treated as a square for cursor purposes.
    If (m_BrushSize > 0!) Then
        If (m_BrushSize <= 1!) And m_StrictPixelCentering Then
            m_BrushOutlinePath.AddRectangle_Absolute -0.6, -0.6, 0.6, 0.6
        Else
            m_BrushOutlinePath.AddCircle 0, 0, m_BrushSize / 2! + 0.5!
        End If
    End If
    
End Sub

Public Function IsFirstDab() As Boolean
    IsFirstDab = m_isFirstStroke
End Function

'Notify the brush engine of the current mouse position.  Coordinates should always be in *image* coordinate space,
' not screen space.  (Translation between spaces will be handled internally.)
Public Sub NotifyBrushXY(ByVal mouseButtonDown As Boolean, ByVal Shift As ShiftConstants, ByVal srcX As Single, ByVal srcY As Single, ByVal mouseTimeStamp As Long, ByRef srcCanvas As pdCanvas)
    
    m_isFirstStroke = (Not m_MouseDown) And mouseButtonDown
    m_isLastStroke = m_MouseDown And (Not mouseButtonDown)
    
    'Perform a failsafe check for brush creation
    If (Not m_BrushIsReady) Then CreateCurrentBrush
    
    'A new toggle (as of 2025) now exists for strictly positioning the cursor in the center of the current pixel.
    If m_StrictPixelCentering Then
        srcX = Int(srcX) + 0.5!
        srcY = Int(srcY) + 0.5!
    End If
    
    'If this is a MouseDown operation, we need to make sure the full paint engine is synchronized against any property
    ' changes that are applied "on-demand".
    If m_isFirstStroke Then
        
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
        
        'Reset the "last mouse position" values to match the current ones
        m_MouseX = srcX
        m_MouseY = srcY
        If m_StrictPixelCentering Then
            m_MouseX = Int(m_MouseX) + 0.5!
            m_MouseY = Int(m_MouseY) + 0.5!
        End If
        
        'Notify the central "color history" manager of the color currently being used
        UserControls.PostPDMessage WM_PD_PRIMARY_COLOR_APPLIED, m_BrushColor, , True
        
        'Initialize any relevant GDI+ objects for the current brush
        Drawing2D.QuickCreateSurfaceFromDC m_Surface, PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.GetDIBDC, (m_BrushAntialiasing = P2_AA_HighQuality)
        
        'If we're directly using GDI+ for painting (by calling various GDI+ line commands), we need to explicitly set
        ' half-pixel offsets, so each pixel "coordinate" is treated as the *center* of the pixel instead of the top-left corner.
        ' (PD's paint engine handles this internally.)
        m_Surface.SetSurfacePixelOffset P2_PO_Half
        
    End If
    
    'Next, determine if the shift key is being pressed.  If it is, and if the user has already committed a
    ' brush stroke to this image (on a previous paint tool event), we want to draw a smooth line between the
    ' last paint point and the current one.  Note that this special condition is stored at module level,
    ' as we render a custom UI on mouse move events if the mouse button is *not* pressed, to help communicate
    ' what the shift key does.
    m_MouseShiftEventOK = (Shift = vbShiftMask) And (m_MouseLastUserX <> MOUSE_OOB) And (m_MouseLastUserY <> MOUSE_OOB)
    m_MouseShiftEventOK = m_MouseShiftEventOK And (m_MouseLastUserX <> srcX) And (m_MouseLastUserY <> srcY)
    
    Dim startTime As Currency
    
    'Do not stroke unless a minimum distance threshold is met.
    ' (This helps work around some GDI+ issues with tiny lines.)
    Const MINIMUM_DISTANCE_TO_STROKE As Single = 0.25
    
    'If the mouse button is down, perform painting between the old and new points.
    ' (All painting occurs in image coordinate space, and is applied to the current image's scratch layer.)
    If mouseButtonDown Then
    
        'Want to profile this function?  Use this line of code (and the matching report line at the bottom of the function).
        VBHacks.GetHighResTime startTime
        
        'The user wants us to connect the start of this stroke to the end of the previous stroke
        If m_MouseShiftEventOK Then
            
            'Replace the last rendering x/y with the mouse position of the last paint event
            m_MouseX = m_MouseLastUserX
            m_MouseY = m_MouseLastUserY
            If m_StrictPixelCentering Then
                m_MouseX = Int(m_MouseX) + 0.5!
                m_MouseY = Int(m_MouseY) + 0.5!
            End If
            
            'Initialize the paint stroker at the previous mouse position (but importantly, ask it to
            ' suspend actual graphics operations - this will initialize things like the compositor rect,
            ' without applying paint to the canvas, and we do it so that the connecting point between
            ' the two strokes is not painted twice)
            ApplyPaintLine srcX, srcY, True, True
            
            'Paint all subsequent strokes
            ApplyPaintLine srcX, srcY, False
            
        'This is a normal paint stroke
        Else
            
            'Require a minimum distance threshold in order to paint
            If (PDMath.DistanceTwoPoints(m_MouseX, m_MouseY, srcX, srcY) >= MINIMUM_DISTANCE_TO_STROKE) Or m_isFirstStroke Then
                ApplyPaintLine srcX, srcY, m_isFirstStroke
            End If
            
        End If
        
        'See if there are more points in the mouse move queue.  If there are, grab them all and stroke them immediately.
        Dim numPointsRemaining As Long
        numPointsRemaining = srcCanvas.GetNumMouseEventsPending
        
        If (numPointsRemaining > 0) And (Not m_isFirstStroke) Then
        
            Dim tmpMMP As MOUSEMOVEPOINT
            Dim imgX As Double, imgY As Double
            
            Do While srcCanvas.GetNextMouseMovePoint(VarPtr(tmpMMP))
                
                'The (x, y) points returned by this request are in the *hWnd's* coordinate space.  We must manually convert them
                ' to the image coordinate space.
                If Drawing.ConvertCanvasCoordsToImageCoords(srcCanvas, PDImages.GetActiveImage(), tmpMMP.x, tmpMMP.y, imgX, imgY) Then
                
                    'As noted elsewhere in this function, all coordinates can be forcibly pixel-centered to ensure
                    ' consistent stroke behavior.
                    If m_StrictPixelCentering Then
                        imgX = Int(imgX) + 0.5!
                        imgY = Int(imgY) + 0.5!
                    End If
                    
                    'The paint layer is always full-size, so we don't need to perform a separate "image space to layer space"
                    ' coordinate conversion here.
                    ApplyPaintLine imgX, imgY, False
                    
                End If
                
            Loop
        
        End If
        
        'Notify the scratch layer of our updates
        PDImages.GetActiveImage.ScratchLayer.NotifyOfDestructiveChanges
        
        'Cache the last x/y position retrieved from the queue
        m_MouseLastUserX = srcX
        m_MouseLastUserY = srcY
    
    'The previous x/y coordinate trackers are updated automatically when the mouse is DOWN.  When the mouse is UP, we must manually
    ' modify those values.
    Else
        m_MouseX = srcX
        m_MouseY = srcY
        If m_StrictPixelCentering Then
            m_MouseX = Int(m_MouseX) + 0.5!
            m_MouseY = Int(m_MouseY) + 0.5!
        End If
    End If
    
    'With all painting tasks complete, update all old state values to match the new state values.
    m_MouseDown = mouseButtonDown
    
    'Unlike other drawing tools, the paintbrush engine controls viewport redraws.  This allows us to optimize behavior
    ' if we fall behind, and a long queue of drawing actions builds up.
    '
    '(Note that we only request manual redraws if the mouse is currently down; if the mouse *isn't* down, the canvas
    ' handles this for us.)
    If mouseButtonDown Then UpdateViewportWhilePainting startTime, srcCanvas
    
    'If the mouse button has been released, we can also release our internal GDI+ objects.
    ' (Note that the current *brush* resources are *not* released, by design.)
    If m_isLastStroke Then
        
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
    If m_Paintbrush.IsItTimeForScreenUpdate(strokeStartTime) Or m_isFirstStroke Then
        
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

'Formally render a line between the old mouse (x, y) coordinate pair and this new pair.  Replacement of the old (x, y) pair
' with the new coordinates is handled automatically.
Private Sub ApplyPaintLine(ByVal srcX As Single, ByVal srcY As Single, ByVal isFirstStroke As Boolean, Optional ByVal skipRendering As Boolean = False)
    
    'Calculate new modification rects, e.g. the portion of the paintbrush layer affected by this stroke.
    ' (The central compositor requires this information for its optimized paintbrush renderer.)
    UpdateModifiedRect srcX, srcY, isFirstStroke
    
    'When using the shift-key to link together disparate strokes, we don't want to paint the connecting point twice.
    ' In that rare circumstance, the caller will request that update our compositor rect, but *skip* actual painting
    ' of the initial dab.
    If (Not skipRendering) Then
    
        'GDI+ refuses to draw a line if the start and end points match; this isn't documented (as far as I know),
        ' but it may exist to provide backwards compatibility with GDI, which deliberately leaves the last point
        ' of a line unplotted, in case you are drawing multiple connected lines.  Because of this, we have to
        ' manually render a dab at the initial starting position.
        If isFirstStroke Then
            
            'Manually fill a circle or rectangle, depending on brush size
            Dim tmpBrush As pd2DBrush
            Set tmpBrush = New pd2DBrush
            Drawing2D.QuickCreateSolidBrush tmpBrush, m_BrushColor, m_BrushOpacity
            
            If (m_BrushSize <= 1!) Then
                PD2D.FillRectangleF m_Surface, tmpBrush, m_MouseX - m_BrushSize * 0.5, m_MouseY - m_BrushSize * 0.5, m_BrushSize, m_BrushSize
            Else
                PD2D.FillCircleF m_Surface, tmpBrush, m_MouseX, m_MouseY, m_BrushSize * 0.5!
            End If
            
            Set tmpBrush = Nothing
            
        Else
            PD2D.DrawLineF m_Surface, m_GDIPPen, m_MouseX, m_MouseY, srcX, srcY
        End If
        
        'Update the "old" mouse coordinate trackers
        m_MouseX = srcX
        m_MouseY = srcY
        
    End If
    
End Sub

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
    halfBrushSize = m_BrushSize * 0.5! + 1!
    
    tmpRectF.Left = tmpRectF.Left - halfBrushSize
    tmpRectF.Top = tmpRectF.Top - halfBrushSize
    
    halfBrushSize = halfBrushSize * 2
    tmpRectF.Width = tmpRectF.Width + halfBrushSize
    tmpRectF.Height = tmpRectF.Height + halfBrushSize
    
    Dim tmpOldRectF As RectF
    
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

'When the active image changes, we need to reset certain brush-related parameters
Public Sub NotifyActiveImageChanged()
    m_MouseX = MOUSE_OOB
    m_MouseY = MOUSE_OOB
    m_MouseLastUserX = MOUSE_OOB
    m_MouseLastUserY = MOUSE_OOB
End Sub

'Return the area of the image modified by the current stroke.  By default, the running modified rect is erased after a call to
' this function, but this behavior can be toggled by resetRectAfter.  Also, if you want to get the full modified rect since this
' paint stroke began, you can set the GetModifiedRectSinceStrokeBegan parameter to TRUE.  Note that when
' GetModifiedRectSinceStrokeBegan is TRUE, the resetRectAfter parameter is ignored.
Public Function GetModifiedUpdateRectF(Optional ByVal resetRectAfter As Boolean = True, Optional ByVal GetModifiedRectSinceStrokeBegan As Boolean = False) As RectF
    If GetModifiedRectSinceStrokeBegan Then
        GetModifiedUpdateRectF = m_TotalModifiedRectF
    Else
        GetModifiedUpdateRectF = m_ModifiedRectF
        If resetRectAfter Then m_UnionRectRequired = False
    End If
End Function

'Want to commit your current brush work?  Call this function to make the brush results permanent.
Public Sub CommitBrushResults()
    
    'This dummy string only exists to ensure that the processor name gets localized properly
    ' (as that text is used for Undo/Redo descriptions).  PD's translation engine will detect
    ' the TranslateMessage() call and produce a matching translation entry.
    Dim strDummy As String
    strDummy = g_Language.TranslateMessage("Pencil stroke")
    Layers.CommitScratchLayer "Pencil stroke", m_TotalModifiedRectF
    
End Sub

'Render the current brush outline to the canvas, using the stored mouse coordinates as the brush's position
Public Sub RenderBrushOutline(ByRef targetCanvas As pdCanvas)
    
    'If a brush outline doesn't exist, create one now
    If (Not m_BrushIsReady) Then CreateCurrentBrush True
    
    'Start by creating a transformation from the image space to the canvas space
    Dim canvasMatrix As pd2DTransform
    Drawing.GetTransformFromImageToCanvas canvasMatrix, targetCanvas, PDImages.GetActiveImage(), m_MouseX, m_MouseY
    
    'We also want to pinpoint the precise cursor position
    Dim cursX As Double, cursY As Double
    Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, PDImages.GetActiveImage(), m_MouseX, m_MouseY, cursX, cursY
    
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
    'cSurface.SetSurfacePixelOffset P2_PO_Normal
    cSurface.SetSurfacePixelOffset P2_PO_Normal
    If (m_BrushSize = 1!) Then cSurface.SetSurfacePixelOffset P2_PO_Half
    
    'If the user is holding down the SHIFT key, paint a line between the end of the previous stroke and the current
    ' mouse position.  This helps communicate that shift+clicking will string together separate strokes.
    If m_MouseShiftEventOK Then
        
        outerPen.SetPenLineCap P2_LC_Round
        innerPen.SetPenLineCap P2_LC_Round
        
        Dim oldX As Double, oldY As Double
        Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, PDImages.GetActiveImage(), m_MouseLastUserX, m_MouseLastUserY, oldX, oldY
        PD2D.DrawLineF cSurface, outerPen, oldX, oldY, cursX, cursY
        PD2D.DrawLineF cSurface, innerPen, oldX, oldY, cursX, cursY
        
    Else
        
        'Paint a target cursor - but *only* if the mouse is not currently down!
        Dim crossLength As Single, crossDistanceFromCenter As Single, outerCrossBorder As Single
        crossLength = 3!
        crossDistanceFromCenter = 4!
        outerCrossBorder = 0.25!
        
        If (Not m_MouseDown) And renderInPrecisionMode Then
        
            outerPen.SetPenLineCap P2_LC_Round
            innerPen.SetPenLineCap P2_LC_Round
            
            'Four "beneath" lines
            PD2D.DrawLineF cSurface, outerPen, cursX, cursY - crossDistanceFromCenter + outerCrossBorder, cursX, cursY - crossDistanceFromCenter - crossLength - outerCrossBorder
            PD2D.DrawLineF cSurface, outerPen, cursX, cursY + crossDistanceFromCenter - outerCrossBorder, cursX, cursY + crossDistanceFromCenter + crossLength + outerCrossBorder
            PD2D.DrawLineF cSurface, outerPen, cursX - crossDistanceFromCenter + outerCrossBorder, cursY, cursX - crossDistanceFromCenter - crossLength - outerCrossBorder, cursY
            PD2D.DrawLineF cSurface, outerPen, cursX + crossDistanceFromCenter - outerCrossBorder, cursY, cursX + crossDistanceFromCenter + crossLength + outerCrossBorder, cursY
            
            'Four "above" lines
            PD2D.DrawLineF cSurface, innerPen, cursX, cursY - crossDistanceFromCenter, cursX, cursY - crossDistanceFromCenter - crossLength
            PD2D.DrawLineF cSurface, innerPen, cursX, cursY + crossDistanceFromCenter, cursX, cursY + crossDistanceFromCenter + crossLength
            PD2D.DrawLineF cSurface, innerPen, cursX - crossDistanceFromCenter, cursY, cursX - crossDistanceFromCenter - crossLength, cursY
            PD2D.DrawLineF cSurface, innerPen, cursX + crossDistanceFromCenter, cursY, cursX + crossDistanceFromCenter + crossLength, cursY
            
        End If
    End If
    
    'If size allows, render a transformed brush outline onto the canvas as well
    If (Not renderInPrecisionMode) Then
        
        'Get a copy of the current brush outline, transformed into position
        Dim copyOfBrushOutline As pd2DPath
        Set copyOfBrushOutline = New pd2DPath
        
        copyOfBrushOutline.CloneExistingPath m_BrushOutlinePath
        copyOfBrushOutline.ApplyTransformation canvasMatrix
        PD2D.DrawPath cSurface, outerPen, copyOfBrushOutline
        PD2D.DrawPath cSurface, innerPen, copyOfBrushOutline
        
    End If
    
    Set cSurface = Nothing
    
End Sub

'Any specialized initialization tasks can be handled here.  This function is called early in the PD load process.
Public Sub InitializeBrushEngine()
    
    'Initialize the underlying brush class
    Set m_Paintbrush = New pdPaintbrush
    
    'Reset UI-centric features
    m_BrushAntialiasing = P2_AA_HighQuality
    
    'Reset all coordinates
    m_MouseX = MOUSE_OOB
    m_MouseY = MOUSE_OOB
    m_MouseLastUserX = MOUSE_OOB
    m_MouseLastUserY = MOUSE_OOB
    m_isFirstStroke = False
    m_isLastStroke = False
    
    'Note that the current brush has *not* been created yet!
    m_BrushIsReady = False
    
End Sub

'Before PD closes, you *must* call this function!  It will free any lingering brush resources (which are cached
' for performance reasons).
Public Sub FreeBrushResources()
    Set m_GDIPPen = Nothing
    Set m_BrushOutlinePath = Nothing
End Sub

