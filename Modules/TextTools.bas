Attribute VB_Name = "Tools_Text"
'***************************************************************************
'Text tools (both of 'em) on-canvas interface wrapper
'Copyright 2015-2026 by Tanner Helland
'Created: 14/May/15
'Last updated: 10/December/21
'Last update: migrate various text-related bits out of pdCanvas in preparation for new tools
'
'To simplify the design of the primary canvas, various text layer requests are blindly forwarded here.
' This module then handles the messy business of forwarding correct text layer commands to the underlying
' vector layer engine.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To simplify the process of setting/getting text settings for a given layer, this Enum is used to pass text properties
Public Enum PD_TextProperty
    ptp_Text = 0
    ptp_FontColor = 1
    ptp_FontFace = 2
    ptp_FontSize = 3
    ptp_FontSizeUnit = 4
    ptp_FontBold = 5
    ptp_FontItalic = 6
    ptp_FontUnderline = 7
    ptp_FontStrikeout = 8
    ptp_HorizontalAlignment = 9
    ptp_VerticalAlignment = 10
    ptp_TextAntialiasing = 11
    ptp_TextContrast = 12
    ptp_RenderingEngine = 13
    ptp_TextHinting = 14
    ptp_WordWrap = 15
    ptp_FillActive = 16
    ptp_FillBrush = 17
    ptp_OutlineActive = 18
    ptp_OutlinePen = 19
    ptp_BackgroundActive = 20
    ptp_BackgroundBrush = 21
    ptp_BackBorderActive = 22
    ptp_BackBorderPen = 23
    ptp_LineSpacing = 24
    ptp_MarginLeft = 25
    ptp_MarginTop = 26
    ptp_MarginRight = 27
    ptp_MarginBottom = 28
    ptp_CharRemap = 29
    ptp_CharSpacing = 30
    ptp_CharOrientation = 31
    ptp_CharJitterX = 32
    ptp_CharJitterY = 33
    ptp_CharInflation = 34
    ptp_CharMirror = 35
    ptp_StretchToFit = 36
    ptp_AlignLastLine = 37
    ptp_OutlineAboveFill = 38
End Enum

#If False Then
    Const ptp_Text = 0, ptp_FontColor = 1, ptp_FontFace = 2, ptp_FontSize = 3, ptp_FontSizeUnit = 4, ptp_FontBold = 5, ptp_FontItalic = 6
    Const ptp_FontUnderline = 7, ptp_FontStrikeout = 8, ptp_HorizontalAlignment = 9, ptp_VerticalAlignment = 10, ptp_TextAntialiasing = 11
    Const ptp_TextContrast = 12, ptp_RenderingEngine = 13, ptp_TextHinting = 14, ptp_WordWrap = 15, ptp_FillActive = 16, ptp_FillBrush = 17
    Const ptp_OutlineActive = 18, ptp_OutlinePen = 19, ptp_BackgroundActive = 20, ptp_BackgroundBrush = 21, ptp_BackBorderActive = 22
    Const ptp_BackBorderPen = 23, ptp_LineSpacing = 24, ptp_MarginLeft = 25, ptp_MarginTop = 26, ptp_MarginRight = 27, ptp_MarginBottom = 28
    Const ptp_CharRemap = 29, ptp_CharSpacing = 30, ptp_CharOrientation = 31, ptp_CharJitterX = 32, ptp_CharJitterY = 33, ptp_CharInflation = 34
    Const ptp_CharMirror = 35, ptp_StretchToFit = 36, ptp_AlignLastLine = 37, ptp_OutlineAboveFill = 38
#End If

'PD's internal glyph renderer supports a number of esoteric capabilities
Public Enum PD_TextWordwrap
    tww_None = 0
    tww_Manual = 1
    tww_AutoCharacter = 2
    tww_AutoWord = 3
End Enum

#If False Then
    Private Const tww_None = 0, tww_Manual = 1, tww_AutoCharacter = 2, tww_AutoWord = 3
#End If

Public Enum PD_CharacterMirror
    cm_None = 0
    cm_Horizontal = 1
    cm_Vertical = 2
    cm_Both = 3
End Enum

#If False Then
    Private Const cm_None = 0, cm_Horizontal = 1, cm_Vertical = 2, cm_Both = 3
#End If

'I'm looking at adding additional "stretch-to-fit" text options in future builds, so this is no longer
' a binary setting but an enum.
Public Enum PD_TextStretchToFit
    stf_None = 0
    stf_Box = 1
    stf_Slab = 2
End Enum

#If False Then
    Private Const stf_None = 0, stf_Box = 1, stf_Slab = 2
#End If

Public Sub NotifyMouseDown(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single)

    'One of two things can happen when the mouse is clicked in text mode:
    ' 1) The current layer is a text layer, and the user wants to edit it
    '     (move it around, resize, etc)
    ' 2) The user wants to add a new text layer, which they can do by clicking
    '     anywhere on the image that isn't already occupied by a text layer.
    
    'Let's start by distinguishing between these two states.
    Dim curPOI As Long
    curPOI = PDImages.GetActiveImage.GetActiveLayer.CheckForPointOfInterest(imgX, imgY)
    
    Dim userIsEditingCurrentTextLayer As Boolean
    userIsEditingCurrentTextLayer = PDImages.GetActiveImage.GetActiveLayer.IsLayerText And (curPOI <> poi_Undefined)
    
    'If the user is editing the current text layer, we can switch directly into
    ' layer transform mode.
    If userIsEditingCurrentTextLayer Then
        
        'Initiate the layer transformation engine.  Note that nothing will happen
        ' until the user actually moves the mouse.
        Tools.SetInitialLayerToolValues PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, imgX, imgY, curPOI
        
    'The user is not editing a text layer.  Create a new text layer for them.
    Else
        
        'Create a new text layer directly; note that we *do not* pass this command through the central processor,
        ' as we don't want the delay associated with full Undo/Redo creation.
        If (g_CurrentTool = TEXT_BASIC) Then
            Layers.AddNewLayer PDImages.GetActiveImage.GetActiveLayerIndex, PDL_TextBasic, 0, 0, 0, True, vbNullString, imgX, imgY, True
        ElseIf (g_CurrentTool = TEXT_ADVANCED) Then
            Layers.AddNewLayer PDImages.GetActiveImage.GetActiveLayerIndex, PDL_TextAdvanced, 0, 0, 0, True, vbNullString, imgX, imgY, True
        End If
        
        'Use a special initialization command that basically copies all existing text properties into the newly created layer.
        Tools.SyncCurrentLayerToToolOptionsUI
        
        'Put the newly created layer into transform mode, with the bottom-right corner selected
        Tools.SetInitialLayerToolValues PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, imgX, imgY, poi_CornerSE
        
        'Also, note that we have just created a new text layer.  The MouseUp event needs to know this, so it can initiate a full-image Undo/Redo event.
        Tools.SetCustomToolState PD_TEXT_TOOL_CREATED_NEW_LAYER
        
        'Redraw the viewport immediately
        Dim tmpViewportParams As PD_ViewportParams
        tmpViewportParams = Viewport.GetDefaultParamObject()
        tmpViewportParams.curPOI = poi_CornerSE
        
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0), VarPtr(tmpViewportParams)
        
    End If

End Sub

Public Sub NotifyMouseUp(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single, ByVal numOfMouseMovements As Long, ByVal clickEventAlsoFiring As Boolean)

    'Pass a final transform request to the layer handler.
    ' (This will initiate Undo/Redo creation, among other things.)
    
    '(Note that this function branches according to two states:
    ' 1) whether this click is creating a new text layer (which requires a full image stack Undo/Redo), or...
    ' 2) whether we are simply modifying an existing text layer.
    If (Tools.GetCustomToolState = PD_TEXT_TOOL_CREATED_NEW_LAYER) Then
        
        'Mark the current tool as busy to prevent any unwanted UI syncing
        Tools.SetToolBusyState True
        
        'See if this was just a click (as it might be at creation time).  If it was,
        ' we need to determine initial text layer boundaries for the user.
        Const MINIMUM_INIT_TEXT_LAYER_SIZE_PX As Single = 5!
        If clickEventAlsoFiring Or (numOfMouseMovements <= 2) Or (PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth < MINIMUM_INIT_TEXT_LAYER_SIZE_PX) Or (PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight < MINIMUM_INIT_TEXT_LAYER_SIZE_PX) Then
            
            'Since the user just clicked on the canvas, we can't know what text boundaries they want.
            ' So, let's do our best to "guess".
            
            'Start by retrieving the current image rectangle that is visible in the viewport.
            Dim curImageRectF As RectF
            PDImages.GetActiveImage.ImgViewport.GetIntersectRectImage curImageRectF
            
            'Working backwards from that, determine initial layer size.
            With PDImages.GetActiveImage()
            
                'To start, we should probably put the text layer's top-left corner at
                ' the current mouse position.
                Dim txtLayerLeft As Single, txtLayerTop As Single
                txtLayerLeft = imgX
                txtLayerTop = imgY
                
                'Width and height are trickier, because who knows how wide/tall of a text layer
                ' the user wants?  We currently aim for at least "500 on-screen pixels" (which
                ' means we need to account for zoom to achieve this), with automatic shrinking
                ' if this extends beyond the edge of the image.
                Const DEFAULT_INIT_TEXT_LAYER_WIDTH As Single = 500!
                Const DEFAULT_INIT_TEXT_LAYER_HEIGHT As Single = 250!
                Dim sizeToAttemptX As Single, sizeToAttemptY As Single
                sizeToAttemptX = Drawing.ConvertCanvasSizeToImageSize(DEFAULT_INIT_TEXT_LAYER_WIDTH, PDImages.GetActiveImage())
                sizeToAttemptY = Drawing.ConvertCanvasSizeToImageSize(DEFAULT_INIT_TEXT_LAYER_HEIGHT, PDImages.GetActiveImage())
                
                Dim txtLayerRight As Single, txtLayerBottom As Single
                txtLayerRight = txtLayerLeft + sizeToAttemptX
                txtLayerBottom = txtLayerTop + sizeToAttemptY
                
                'If this places the layer outside image bounds, shrink it accordingly
                If (txtLayerRight > PDImages.GetActiveImage.Width) Then txtLayerRight = PDImages.GetActiveImage.Width
                If (txtLayerBottom > PDImages.GetActiveImage.Height) Then txtLayerBottom = PDImages.GetActiveImage.Height
                
                'Similarly, if this places the layer outside *viewport* bounds, shrink it accordingly.
                If (txtLayerRight > (curImageRectF.Left + curImageRectF.Width)) Then txtLayerRight = (curImageRectF.Left + curImageRectF.Width)
                If (txtLayerBottom > (curImageRectF.Top + curImageRectF.Height)) Then txtLayerBottom = (curImageRectF.Top + curImageRectF.Height)
                
                'Finally, ensure a minimum width/height of some arbitrary value
                ' (see above for constant declaration).
                If (txtLayerRight - txtLayerLeft < MINIMUM_INIT_TEXT_LAYER_SIZE_PX) Then
                    txtLayerLeft = txtLayerRight - MINIMUM_INIT_TEXT_LAYER_SIZE_PX
                    If (txtLayerLeft < 0!) Then txtLayerLeft = 0!
                End If
                If (txtLayerBottom - txtLayerTop < MINIMUM_INIT_TEXT_LAYER_SIZE_PX) Then
                    txtLayerTop = txtLayerBottom - MINIMUM_INIT_TEXT_LAYER_SIZE_PX
                    If (txtLayerTop < 0!) Then txtLayerTop = 0!
                End If
                
                'Apply the final values to the layer
                .GetActiveLayer.SetLayerOffsetX txtLayerLeft
                .GetActiveLayer.SetLayerOffsetY txtLayerTop
                .GetActiveLayer.SetLayerWidth Int(txtLayerRight - txtLayerLeft + 0.5)
                .GetActiveLayer.SetLayerHeight Int(txtLayerBottom - txtLayerTop + 0.5)
                
            End With
            
            'If the current text box is empty, set some new text to orient the user
            If (g_CurrentTool = TEXT_BASIC) Then
                If (LenB(toolpanel_TextBasic.txtTextTool.Text) = 0) Then toolpanel_TextBasic.txtTextTool.Text = g_Language.TranslateMessage("(enter text here)")
            Else
                If (LenB(toolpanel_TextAdvanced.txtTextTool.Text) = 0) Then toolpanel_TextAdvanced.txtTextTool.Text = g_Language.TranslateMessage("(enter text here)")
            End If
            
            'Manually synchronize the new size values against their on-screen UI elements
            Tools.SyncToolOptionsUIToCurrentLayer
            
            'Manually force a viewport redraw
            Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
            
        'If the user already specified a size, use their values to finalize the layer size
        Else
            Tools.TransformCurrentLayer imgX, imgY, PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, FormMain.MainCanvas(0), (Shift And vbShiftMask)
        End If
        
        'As a failsafe, ensure the layer has a proper rotational center point.  (If the user dragged the mouse so that
        ' the text box was 0x0 pixels at some size, the rotational center point math would have failed and become (0, 0)
        ' to match.)
        PDImages.GetActiveImage.GetActiveLayer.SetLayerRotateCenterX 0.5
        PDImages.GetActiveImage.GetActiveLayer.SetLayerRotateCenterY 0.5
        
        'Release the tool engine
        Tools.SetToolBusyState False
        
        'Process the addition of the new layer; this will create proper Undo/Redo data for the entire image (required, as the layer order
        ' has changed due to this new addition).
        With PDImages.GetActiveImage.GetActiveLayer
            Process "New text layer", , BuildParamList("layerheader", .GetLayerHeaderAsXML(), "layerdata", .GetVectorDataAsXML()), UNDO_Image_VectorSafe
        End With
        
        'Manually synchronize menu, layer toolbox, and other UI settings against the newly created layer.
        Interface.SyncInterfaceToCurrentImage
        
        'Finally, set focus to the text layer text entry box
        If (g_CurrentTool = TEXT_BASIC) Then
            toolpanel_TextBasic.NotifyNewLayerCreated
        Else
            toolpanel_TextAdvanced.NotifyNewLayerCreated
        End If
        
    'The user is simply editing an existing layer.
    Else
        
        'As a convenience to the user, ignore clicks that don't actually change layer settings
        If (numOfMouseMovements > 0) Then Tools.TransformCurrentLayer imgX, imgY, PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, FormMain.MainCanvas(0), (Shift And vbShiftMask), True
        
    End If
    
    'Reset the generic tool mouse tracking function
    Tools.TerminateGenericToolTracking

End Sub
