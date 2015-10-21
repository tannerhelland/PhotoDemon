VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdCompositor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Compositor class
'Copyright 2014-2015 by Tanner Helland
'Created: 01/May/14
'Last updated: 16/October/15
'Last update: implement level-of-detail (LOD) caches for compositor requests
'
'Image compositing is a fairly arduous process, especially when dealing with stuff like custom blend modes.
' Previously, the pdImage class handled all compositing on its own, but as PD's compositing needs have
' grown more complex, I thought it prudent to separate compositing code into a dedicated class.
'
'One instance of this compositor class is stored within each pdImage object.  In the future, it may be helpful
' to devise some sort of caching mechanism to reduce the amount of data traded between the two classes, but at
' present the parent pdImage object always passes a self-reference to this class when a compositing operation
' is requested.  (This is simpler than dealing with persistent circular references.)
'
'As of 20 June '14, this class also supports the processing of non-destructive edits for a given layer.  Many
' optimizations have been made to try and improve the performance of non-destructive edits, and those optimizations
' comprise a large portion of the class.
'
'Many thanks to expert contributor Audioglider for his help implementing PhotoDemon's many supported blend modes.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Compositing requires a temporary assembly DIB at the size of the destination image (viewport).  All compositing must be done to this DIB,
' and as the final step, the composited DIB is alpha-blended onto the DIB passed by the caller.  Why not composite directly onto the DIB
' passed by the caller?  The caller is allowed to provide any background (color, checkerboard, etc) based on their intended purpose.
' Non-standard blendmodes use underlying pixels to calculate layer color.  We do not want the caller's DIB to be part of those calculations,
' so we must build our own DIB, then at the end - after all compositing is finished - blend it onto the user's background.
'
'Anyway, as of v6.6, PD now caches this DIB at module-level.  This spares us from recreating the DIB if the viewport size hasn't changed,
' which improves responsiveness during painting.
Private m_dstDIB As pdDIB

'Non-standard blend modes require use of a temporary DIB, because we must do our own compositing.  In an attempt to improve performance,
' we cache this at module level.  If an image only has a single non-standard blend mode layer, we can avoid recreating this DIB between
' composite passes.
Private m_tmpDIB As pdDIB

'Non-destructive effects require use of a temporary DIB.  In the event that only a single layer is being non-destructively edited, we can cache
' the DIB at module level, and perform a simple erase instead of recreation on updates.  This improves non-destructive FX performance.
Private m_NDFXDIB As pdDIB

'If a layer has non-destructive, non-standard transformations active (e.g. rotation, skew), we use a special, parallelogram-based rendering function.
' To cut down on variable declarations, we declare this array once, and initialize it when a compositor instance is initialized.
Private m_PlgPoints() As POINTFLOAT

'For the rect-specific compositor, we also need a copy of the parallelogram points in the *canvas*/viewport coordinate space
Private m_PlgPointsDst() As POINTFLOAT

'Historically, pdCompositor handled its own blending, but as part of wide-ranging compositor optimizations in the 7.0 release, blending was
' moved to a dedicated class.
Private m_Blender As pdPixelBlender

'Composite two DIBs using the requested blend mode.  An offset can be specified for the top DIB, relative to the bottom DIB.
' Per the function name, the *top DIB* is treated as the destination.  After compositing, the caller would need to alpha-blend
' the finished top DIB over the bottom DIB to achieve full composition.
'
'TODO: profile an integer-based implementation.
Friend Sub compositeDIBsWithoutBlending(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal blendMode As LAYER_BLENDMODE, ByVal dstX As Single, ByVal dstY As Single, Optional ByVal alphaMode As LAYER_ALPHAMODE = LA_NORMAL)

    'The pixel blender class handles this for us
    m_Blender.blendDIBs topDIB, bottomDIB, blendMode, dstX, dstY, 1, alphaMode, False

End Sub

'Perform a full composite and blend of two DIBs, using the requested blend mode.
'
'IMPORTANT NOTE: unlike compositeDIBs(), above, this function merges the result of the blend into the BOTTOM DIB, using the supplied
'                 alpha modifier.  This means that the top layer remains UNTOUCHED, but the bottom layer IS MODIFIED.  The end result
'                 of this function is a bottom layer representing a "merge" of the two layers.
'
'This function should not be used outside of getCompositedImage, because that function guarantees proper boundary checks.  This function assumes
' a number of things about the composited DIBs (e.g. that they overlap, that the bottom layer is large enough to contain the overlap, etc), and it
' *will crash* if you call it directly without meeting this criteria.
'
'TODO: profile an integer-based implementation.
Friend Sub compositeAndBlendDIBs(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal blendMode As LAYER_BLENDMODE, ByVal dstX As Single, ByVal dstY As Single, Optional ByVal alphaModifier As Double = 1#, Optional ByVal alphaMode As LAYER_ALPHAMODE = LA_NORMAL)
    
    'The blend class handles this for us
    m_Blender.blendDIBs topDIB, bottomDIB, blendMode, dstX, dstY, alphaModifier, alphaMode, True
    
End Sub

'Merge two layers together.  Note this can be used to merge any two arbitrary layers, with the bottom layer holding the result
' of the merge.  It is up to the caller to deal with any subsequent layer deletions, etc - this sub just performs the merge.
'
'The optional parameter, "bottomLayerIsFullSize", should be set to TRUE if the bottom layer is the size of the image.  This saves
' us some processing time, because we don't have to check for rect intersection.
Public Sub mergeLayers(ByRef topLayer As pdLayer, ByRef bottomLayer As pdLayer, Optional ByVal bottomLayerIsFullSize As Boolean = True)

    Dim xOffset As Double, yOffset As Double
    Dim i As Long
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'If the layers are vector layers, make sure their DIBs are properly synchronized.  (If the layers are not vector layers,
    ' this action does nothing.)
    topLayer.syncInternalDIB
    bottomLayer.syncInternalDIB
    
    'We also need to commit any non-destructive effects; otherwise, the merged layer may look different from its on-screen representation,
    ' as the settings of the bottom layer will be "assumed" for the merged layer as well.
    If bottomLayer.getLayerNonDestructiveFXState Then
        Me.applyNDFXToDIB bottomLayer, bottomLayer.layerDIB
        bottomLayer.deactivateAllNDFX
    End If
    
    If topLayer.getLayerNonDestructiveFXState Then
        Me.applyNDFXToDIB topLayer, topLayer.layerDIB
        topLayer.deactivateAllNDFX
    End If
    
    'Next, we need to find the union rect of the two layers.  This is the smallest rectangle that holds both layers.
    Dim bottomRect As RECTF, topRect As RECTF, finalRect As RECTF
    Layer_Handler.fillRectForLayerF bottomLayer, bottomRect, True, True
    Layer_Handler.fillRectForLayerF topLayer, topRect, True, True
    UnionRectF finalRect, topRect, bottomRect
    
    'finalRect now contains the coordinates of the union rect.  (Note that some settings allow us to bypass this rect entirely, e.g. if the
    ' bottom layer is image-sized and the top layer has no active affine transforms.)
    
    'If the bottom layer is the size of the image itself (e.g. when flattening the image), we can use its DIB as the target DIB instead of
    ' creating a temporary one.  This can provide a nice performance boost depending on the complexity of subsequent layers.
    If bottomLayerIsFullSize Then
    
        xOffset = topLayer.getLayerOffsetX
        yOffset = topLayer.getLayerOffsetY
    
    'The top and bottom layer sizes are totally independent.  This makes our life somewhat unpleasant.
    Else
        
        'Create a blank DIB at the dimensions of the union rect.
        tmpDIB.createBlank finalRect.Width, finalRect.Height, 32, 0
        tmpDIB.setInitialAlphaPremultiplicationState True
        
        'We now need to do a couple of things.  Let's start by copying the bottom DIB into this new temporary DIB.
        xOffset = bottomRect.Left - finalRect.Left
        yOffset = bottomRect.Top - finalRect.Top
        
        'If no non-destructive resizes are active, we can simply copy the layer over as-is
        If Not bottomLayer.affineTransformsActive(True) Then
        
            If bottomLayer.getLayerOpacity = 100 Then
                
                With bottomLayer
                    BitBlt tmpDIB.getDIBDC, xOffset, yOffset, .getLayerWidth(False), .getLayerHeight(False), .layerDIB.getDIBDC, 0, 0, vbSrcCopy
                End With
                
            Else
                With bottomLayer
                    .layerDIB.alphaBlendToDC tmpDIB.getDIBDC, .getLayerOpacity * 2.55, xOffset, yOffset
                End With
            End If
        
        'Active transforms make this step harder.
        Else
            
            'If the only active transform is rescaling, we can use a shortcut StretchBlt function for a performance boost.
            If Not bottomLayer.affineTransformsActive(False) Then
            
                With bottomLayer
                    GDI_Plus.GDIPlus_StretchBlt tmpDIB, xOffset, yOffset, .getLayerWidth(True), .getLayerHeight(True), .layerDIB, 0, 0, .getLayerWidth(False), .getLayerHeight(False), .getLayerOpacity / 100, .getLayerResizeQuality_GDIPlus
                End With
            
            'If rotation or other affine transforms are active, a more complicated solution is required.
            Else
            
                'First, retrieve the layer's coordinates in the default image coordinate space
                bottomLayer.getLayerCornerCoordinates m_PlgPoints
                
                'Next, we need to convert those to the union rect coordinate space
                For i = 0 To 3
                    m_PlgPoints(i).x = (m_PlgPoints(i).x - finalRect.Left)
                    m_PlgPoints(i).y = (m_PlgPoints(i).y - finalRect.Top)
                Next i
                
                'Render the transformed DIB now.
                With bottomLayer
                    GDI_Plus.GDIPlus_PlgBlt tmpDIB, m_PlgPoints, .layerDIB, 0, 0, .getLayerWidth(False), .getLayerHeight(False), .getLayerOpacity / 100, .getLayerResizeQuality_GDIPlus
                End With
            
            End If
            
            'With the bottom layer successfully rendered onto the destination DIB, we can reset all non-destructive size modifiers now.
            bottomLayer.setLayerCanvasXModifier 1
            bottomLayer.setLayerCanvasYModifier 1
            bottomLayer.setLayerAngle 0
            
        End If
        
        'We now need to calculate a new layer offset for this temporary DIB, which will eventually be copied into the bottom layer.
        ' Without this, the main composite won't know where to stick the layer!
        bottomLayer.setLayerOffsetX finalRect.Left
        bottomLayer.setLayerOffsetY finalRect.Top
        
        'Copy the temporary DIB into the bottom layer.  (We can't simply reference it as we're likely to need a temporary DIB for
        ' processing the top layer as well.)
        bottomLayer.layerDIB.createFromExistingDIB tmpDIB
        tmpDIB.eraseDIB
        
        'Reset the layer's opacity, as we have already applied it in the previous steps.
        bottomLayer.setLayerOpacity 100
        
        'Calculate new offsets for the top layer, then carry on with business as usual!
        xOffset = topRect.Left - finalRect.Left
        yOffset = topRect.Top - finalRect.Top
        
    End If
    
    'The bottom layer DIB now represents the target for the final layer merge.  It has had any non-destructive effects processed, its opacity
    ' has been accounted for (and it's tracking value reset to 100), and any non-destructive transformations (resize, scaling, etc) have been
    ' made permanent.
    
    'Now all we need to do is merge the top layer onto it.
    
    'If the top layer has no non-destructive resizing active, this step becomes much easier.
    If Not topLayer.affineTransformsActive(True) Then
        
        'We don't even need a temporary DIB - just composite the current layer as-is
        If (topLayer.getLayerBlendMode = BL_NORMAL) And (topLayer.getLayerAlphaMode = LA_NORMAL) Then
            topLayer.layerDIB.alphaBlendToDC bottomLayer.layerDIB.getDIBDC, topLayer.getLayerOpacity * 2.55, xOffset, yOffset
        Else
            compositeAndBlendDIBs topLayer.layerDIB, bottomLayer.layerDIB, topLayer.getLayerBlendMode, xOffset, yOffset, topLayer.getLayerOpacity / 100, topLayer.getLayerAlphaMode
        End If
    
    'Non-destructive transforms are active, so we must prep a temporary DIB for the top layer (containing a transformed copy of its DIB)
    Else
    
        'Rescale-only transforms can use a shortcut copy method
        If Not topLayer.affineTransformsActive(False) Then
            
            With topLayer
                tmpDIB.createBlank .getLayerWidth(True), .getLayerHeight(True), 32, 0
                tmpDIB.setInitialAlphaPremultiplicationState True
                GDIPlusResizeDIB tmpDIB, 0, 0, .getLayerWidth(True), .getLayerHeight(True), .layerDIB, 0, 0, .getLayerWidth(False), .getLayerHeight(False), .getLayerResizeQuality_GDIPlus
            End With
        
        'Full affine transforms are hairier
        Else
            
            'Ask the layer to give us a copy of its transformed DIB
            Dim xOffsetLong As Long, yOffsetLong As Long
            topLayer.getAffineTransformedDIB tmpDIB, xOffsetLong, yOffsetLong, 0, 0
            
            'Calculate new offsets, using the x/y returned by that function (which may be slightly modified in order to preserve subpixel positioning)
            xOffset = xOffsetLong - finalRect.Left
            yOffset = yOffsetLong - finalRect.Top
            
        End If
        
        'With all non-destructive options accounted for, we can finally calculate composited layer data!
        If (topLayer.getLayerBlendMode = BL_NORMAL) And (topLayer.getLayerAlphaMode = LA_NORMAL) Then
            tmpDIB.alphaBlendToDC bottomLayer.layerDIB.getDIBDC, topLayer.getLayerOpacity * 2.55, xOffset, yOffset
        Else
            compositeAndBlendDIBs tmpDIB, bottomLayer.layerDIB, topLayer.getLayerBlendMode, xOffset, yOffset, topLayer.getLayerOpacity / 100, topLayer.getLayerAlphaMode
        End If
                
    End If
    
    'The two layers have been merged successfully!  Any further actions (like deleting the top layer) must be handled by the caller.
    
End Sub

'Given two DIBs of equal size, perform a fast merge with variable blend mode and opacity.  Previously, PD required you to create temporary layers
' for each DIB, assign them blendmodes, then manually perform a full merge (with checks for non-destructive effects, sizing, etc).  If all you want
' is two merged DIBs, use this function.
'
'IMPORTANT NOTE!  Both the top and bottom DIBs will be modified by this function, per standard merge rules.  The top DIB must be modified because it
'                  contains the intermediate results of the merge.  The bottom DIB must be modified because it contains the final result.  As such,
'                  use temp copies of DIBs you can't afford to lose.
Friend Sub quickMergeTwoDibsOfEqualSize(ByRef bottomDIB As pdDIB, ByRef topDIB As pdDIB, Optional ByVal blendMode As LAYER_BLENDMODE = BL_NORMAL, Optional ByVal topLayerOpacity As Double = 100#, Optional ByVal alphaMode As LAYER_ALPHAMODE = LA_NORMAL)
    Me.compositeAndBlendDIBs topDIB, bottomDIB, blendMode, 0, 0, topLayerOpacity / 100, alphaMode
End Sub

'Returns all layers of the image as a single, composited image (in pdDIB format, of course).  Because of the way VB handles
' object references, we ask the calling function to supply the DIB they want filled.  Optionally, they can also request a
' particular premultiplication status of the composited DIB's alpha values.  (This is helpful for save functions, which
' require non-premultiplied alpha, vs viewport functions, which require premultiplied alpha).
Friend Sub getCompositedImage(ByRef srcImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal premultiplicationStatus As Boolean = True)
    
    'TODO: perform special pre-check for 24bpp single-layer images.  For these, we can simply copy the current layer out,
    '       saving a lot of time!
    
    'Start by creating the destination DIB, as necessary
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    'Next, size the destination DIB to the match the composited image
    If (dstDIB.getDIBWidth <> srcImage.Width) Or (dstDIB.getDIBHeight <> srcImage.Height) Then
        dstDIB.createBlank srcImage.Width, srcImage.Height, 32
        dstDIB.setInitialAlphaPremultiplicationState True
    Else
        dstDIB.resetDIB
    End If
    
    'Subsequent steps are going to use a helper function.  This helper function makes it simple to "insert" paint operations
    ' into the current layer stack, without fucking up the existing layer order.
        
    'If the image has additional layers, proceed to merge the rest of them, starting from the bottom and working our way up.
    ' Note that if a layer is invisible, we simply skip it - this is the most performance-friendly way to handle them.
    Dim i As Long
    For i = 0 To srcImage.getNumOfLayers - 1
        
        'Make sure the layer is visible
        If srcImage.getLayerByIndex(i).getLayerVisibility Then
            
            'Pass this layer to the helper function.  It will take care of the compositing step for us.
            getCompositedImage_LayerHelper dstDIB, srcImage.getLayerByIndex(i), i
        
        'End layer visibility check
        End If
        
    Next i
        
    'If the user requested non-premultiplied alpha, calculate it now.
    ' (By default, this function always returns a premultiplied image, because that's what the compositor functions return.)
    If premultiplicationStatus Then
        dstDIB.setInitialAlphaPremultiplicationState True
    Else
        If dstDIB.getDIBColorDepth = 32 Then dstDIB.setAlphaPremultiplication False
    End If
    
End Sub

'This helper function is only used internally, by the getCompositedImage function.
'
'The purpose of this helper function is to take a source layer, and merge it onto a destination DIB (*not* a destination layer).
' This logic is handled separately so that paint operations can be silently "inserted" into the existing layer stack, without requiring
' any modifications of the parent pdImage object.
'
'For performance reasons, the module-level m_tmpDIB object is used.  This is by design, because images that have only one
' special-treatment layer (e.g. non-destructive FX, resizing, custom blend modes, etc) can avoid the creation of multiple temporary DIBs.
Private Sub getCompositedImage_LayerHelper(ByRef dstDIB As pdDIB, ByRef srcLayer As pdLayer, ByVal srcLayerIndex As Long)
    
    'Hypothetically, we could write special handling for 24-bpp DIBs in the future.  But at present, PD only operates on 32-bpp DIBs.
    
    'Some variables are potentially used in multiple branches of this function
    Dim layerRect As RECTF
    Dim intLayerOffsetX As Long, intLayerOffsetY As Long
    Dim layerBoundariesOkay As Boolean
    
    'If this layer is a vector layer, make sure its DIB is properly synchronized.  (If it's *not* a vector layer, this action does nothing.)
    srcLayer.syncInternalDIB
    
    'Custom blend-modes and alpha-modes are available to all layers but the base one.  If the normal blend mode is active, we can shortcut
    ' a lot of the compositing pipeline.
    Dim standardBlendingActive As Boolean
    
    If srcLayerIndex = 0 Then
        standardBlendingActive = True
    Else
        standardBlendingActive = CBool(srcLayer.getLayerBlendMode = BL_NORMAL) And CBool(srcLayer.getLayerAlphaMode = LA_NORMAL)
    End If
    
    'Start by separating our handling into "simple" and "complex" branches.  Basic layers with no non-destructive effects or much easier
    ' to deal with.
    If Not srcLayer.getLayerNonDestructiveFXState Then
        
        'Non-destructive effects are not active.  Yay!
        
        'Next, check for custom blend modes.  These require an intermediate DIB copy of the image, because PD must internally render the
        ' blend modes (as GDI+ can't do it for us).
        If standardBlendingActive Then
        
            'Custom blend modes are not active.  Yay!
            
            'Check for affine transformations.  If none are present (e.g. the layer is at its original size and orientation),
            ' we can use GDI's AlphaBlend for a performance boost.
            If (Not srcLayer.affineTransformsActive(True)) Then
            
                With srcLayer
                    .layerDIB.alphaBlendToDC dstDIB.getDIBDC, .getLayerOpacity * 2.55, .getLayerOffsetX, .getLayerOffsetY
                End With
            
            Else
                
                'One or more affine transformations are active.  GDI+ must be used to render the image.
                With srcLayer
                
                    'If scaling is the only active affine transformation, we can use a special, optimized render function
                    If (Not srcLayer.affineTransformsActive(False)) Then
                        GDI_Plus.GDIPlus_StretchBlt dstDIB, .getLayerOffsetX, .getLayerOffsetY, .getLayerWidth(True), .getLayerHeight(True), .layerDIB, 0, 0, .getLayerWidth(False), .getLayerHeight(False), .getLayerOpacity / 100, .getLayerResizeQuality_GDIPlus
                    
                    'If non-scaling affine transforms are active, we must perform a full affine transformation.
                    Else
                        
                        'Retrieve the layer's destination position into our POINTFLOAT array
                        .getLayerCornerCoordinates m_PlgPoints, True
                        
                        'Apply a full parallelogram transformation
                        GDI_Plus.GDIPlus_PlgBlt dstDIB, m_PlgPoints, .layerDIB, 0, 0, .getLayerWidth(False), .getLayerHeight(False), .getLayerOpacity / 100, .getLayerResizeQuality_GDIPlus
                        
                    End If
                    
                End With
            
            End If
        
        'This Else statement represents the case "If srcLayer.getLayerBlendMode <> BL_NORMAL..."
        Else
        
            'Custom blend modes are active.  We have no choice but to create a temporary compositing DIB.
            With srcLayer
            
                'If non-destructive transforms are NOT active, pass the source layer to the all-in-one composite + blend function
                If (Not .affineTransformsActive(True)) Then
                
                    'Perform an all-in-one composite and blend using PD's awesome all-in-one function
                    compositeAndBlendDIBs .layerDIB, dstDIB, .getLayerBlendMode, .getLayerOffsetX, .getLayerOffsetY, .getLayerOpacity / 100, .getLayerAlphaMode
                
                'One or more affine transforms are present.  We have no choice but to create a temporary intermediate DIB,
                ' prior to full compositing.
                Else
                
                    'The layer itself provides a highly optimized method for extracting the affine-transformed portion of its DIB
                    layerBoundariesOkay = .getAffineTransformedDIB(m_tmpDIB, intLayerOffsetX, intLayerOffsetY, dstDIB.getDIBWidth, dstDIB.getDIBHeight)
                
                    'Next, perform an all-in-one composite + blend using our awesome new compositeAndBlendDIBs function
                    If layerBoundariesOkay Then compositeAndBlendDIBs m_tmpDIB, dstDIB, .getLayerBlendMode, intLayerOffsetX, intLayerOffsetY, .getLayerOpacity / 100, .getLayerAlphaMode
                    
                End If
                
            End With
            
        End If
        
    'Non-destructive effects are active.  This requires a fairly significant chunk of special processing (ugh).
    Else
    
        'See if non-destructive resizing is active.  If it is, create a copy of the source DIB, *with* resizing applied.  This will be
        ' passed to the non-destructive FX function.
        
        '(Potential future TODO: if non-destructive resizing makes the image larger, apply the NDFX first, then the resize)
        If (Not srcLayer.affineTransformsActive(True)) Then
            layerBoundariesOkay = m_tmpDIB.createFromExistingDIB(srcLayer.layerDIB)
        Else
            
            'The layer itself provides a highly optimized method for extracting the affine-transformed portion of its DIB
            layerBoundariesOkay = srcLayer.getAffineTransformedDIB(m_tmpDIB, intLayerOffsetX, intLayerOffsetY, dstDIB.getDIBWidth, dstDIB.getDIBHeight)
                
        End If
        
        If layerBoundariesOkay Then
        
            'Apply any/all non-destructive effects to the temporary DIB
            applyNDFXToDIB srcLayer, m_tmpDIB
            
            With srcLayer
                
                'Next, check the active blend and alpha modes.  NORMAL blend and NORMAL alpha mode can be handled by GDI's AlphaBlend
                ' for a nice performance boost.
                If standardBlendingActive Then
                
                    If .affineTransformsActive(False) Then
                        m_tmpDIB.alphaBlendToDC dstDIB.getDIBDC, .getLayerOpacity * 2.55, intLayerOffsetX, intLayerOffsetY
                    Else
                        m_tmpDIB.alphaBlendToDC dstDIB.getDIBDC, .getLayerOpacity * 2.55, .getLayerOffsetX, .getLayerOffsetY
                    End If
                
                'Custom blend and/or alpha modes are active.  Perform a single pass composite + blend operation using our awesome
                ' all-in-one function.
                Else
                
                    If .affineTransformsActive(False) Then
                        compositeAndBlendDIBs m_tmpDIB, dstDIB, .getLayerBlendMode, intLayerOffsetX, intLayerOffsetY, .getLayerOpacity / 100, .getLayerAlphaMode
                    Else
                        compositeAndBlendDIBs m_tmpDIB, dstDIB, .getLayerBlendMode, .getLayerOffsetX, .getLayerOffsetY, .getLayerOpacity / 100, .getLayerAlphaMode
                    End If
            
                End If
                
            End With
            
        End If
        
    'End NDFX branching
    End If
        
End Sub

'Returns a subsection of the fully composited image (in pdDIB format, of course).  This is helpful for rendering the main viewport,
' as we only composite the relevant portions of the image.
'
'This function is large, complicated, and unfortunately tricky, as we must resize and composite each layer in turn.  However, it's the secret
' to PD's incredible viewport performance.  (e.g., this function is how we can outperform both GIMP and Paint.NET despite being single-threaded
' and heavily CPU-bound)
Friend Sub getCompositedRect(ByRef srcImage As pdImage, ByRef dstDIB As pdDIB, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal interpolationType As InterpolationMode, Optional ByVal ignoreInternalCaches As Boolean = False, Optional ByVal levelOfDetail As COMPOSITOR_LOD_CLUE = CLC_Generic)
    
    'Before doing anything else, fill two rects with the coordinates we were passed.  We will use these rects to calculate the
    ' boundary regions for all individual layers.
    Dim dstViewportRect As RECTF, srcImageRect As RECTF
    
    With dstViewportRect
        .Left = dstX
        .Top = dstY
        .Width = dstWidth
        .Height = dstHeight
    End With
    
    With srcImageRect
        .Left = srcX
        .Top = srcY
        .Width = srcWidth
        .Height = srcHeight
    End With
        
    'A few other things to note before we begin:
    ' - dstDIB will have already been created by the caller.  It may contain a background checkerboard, so we can't
    '    delete or recreate it.
    ' - The difference between the destination and source sizes can be used to infer a zoom value.  Note that aspect
    '    ratio will *always* be preserved by this function.
    ' - At present, GDI+ is used for all resizing.  Any other function with a StretchBlt-like interface could also
    '    be used, but GDI+ is currently the fastest method we have access to.
    
    'Check for the special case of 100% zoom (e.g. the source rect and dest rect are the same size).  When this happens, we can
    ' shortcut certain parts of the compositing process.
    Dim isRectZoomless As Boolean
    
    If (dstWidth = srcWidth) And (dstHeight = srcHeight) Then
        isRectZoomless = True
    Else
        isRectZoomless = False
    End If
    
    'Start by creating a temporary DIB the size of the destination image (viewport).  All compositing will be done to this DIB,
    ' and as the final step, we will AlphaBlend the finished image onto dstDIB.
    Dim dibWidthCalc As Long, dibHeightCalc As Long
    dibWidthCalc = Int(dstWidth + (dstX - Int(dstX)) + 0.99)
    dibHeightCalc = Int(dstHeight + (dstY - Int(dstY)) + 0.99)
    
    'Note that we try to avoid recreating this DIB unless absolutely necessary
    If (m_dstDIB.getDIBWidth < dibWidthCalc) Or (m_dstDIB.getDIBHeight < dibHeightCalc) Then
        m_dstDIB.createBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
        m_dstDIB.setInitialAlphaPremultiplicationState True
    Else
        m_dstDIB.resetDIB 0
    End If
    
    'Now our work is pretty simple: iterate layers, and pass visible ones off to the helper function (which handles the messy
    ' work of solving all the coordinate math)
    Dim i As Long
    For i = 0 To srcImage.getNumOfLayers - 1
        
        'Only process a layer if it is currently visible.
        If srcImage.getLayerByIndex(i).getLayerVisibility Then
            
            'The helper function handles the actual compositing
            getCompositedRect_LayerHelper m_dstDIB, srcImage, srcImage.getLayerByIndex(i), dstViewportRect, srcImageRect, interpolationType, isRectZoomless, CBool(i = 0), ignoreInternalCaches, levelOfDetail
        
        'END block for layer visibility toggle
        End If
        
    Next i
    
    'With all layers successfully blended onto tmpLayer's DIB, we can now perform a final alphaBlend onto dstDIB.
    ' Because of compositor rewrites in v7.0, we can use GDI's AlphaBlend regardless of zoom, because the module-level DIB was
    ' constructed with subpixel positioning already accounted for.
    '
    'For this release, I have left the original GDI+ paint method, in case I find it necessary to revert back in the future.
    If isRectZoomless Then
        m_dstDIB.alphaBlendToDCEx dstDIB.getDIBDC, dstX, dstY, dstWidth, dstHeight, 0, 0, dstWidth, dstHeight
    Else
        
        Dim bltWidth As Single, bltHeight As Single
        bltWidth = dstWidth + (dstX - Int(dstX))
        bltHeight = dstHeight + (dstY - Int(dstY))
        GDIPlus_StretchBlt dstDIB, Int(dstX), Int(dstY), bltWidth, bltHeight, m_dstDIB, 0, 0, bltWidth, bltHeight, 1, InterpolationModeNearestNeighbor, , True
        
        'While we might be able to get away with GDI's AlphaBlend here, it may cause bleeding on the trailing width/height edges,
        ' which can cause the underlying checkerboard to show through.
        'm_dstDIB.alphaBlendToDCEx dstDIB.getDIBDC, Int(dstX), Int(dstY), dibWidthCalc, dibHeightCalc, 0, 0, dibWidthCalc, dibHeightCalc
        
    End If
    
    'The target DIB will always be premultiplied, by design
    dstDIB.setInitialAlphaPremultiplicationState True
    
End Sub

'This helper function is only used internally, by the getCompositedRect function.
'
'The purpose of this helper function is to take a source layer, and merge it onto a destination DIB (*not* a destination layer).
' This logic is handled separately so that paint operations can be silently "inserted" into the existing layer stack, without requiring
' any modifications of the parent pdImage object.
'
'Extensive caching has been implemented to maximize performance.  Each individual pdLayer object stores its own viewport cache, but note that
' the optional ignoreInternalCaches can be set to TRUE, in which case these caches must be forcibly regenerated.  If the caches do not need to
' be forcibly generated, the optional levelOfDetail parameter should be correctly set, which clues the compositor into which LOD cache to use
' in the target layer, greatly improving performance.
Private Sub getCompositedRect_LayerHelper(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByRef dstViewportRect As RECTF, ByRef srcImageRect As RECTF, ByVal interpolationType As InterpolationMode, Optional ByVal isRectZoomless As Boolean = False, Optional ByVal isBaseLayer As Boolean = False, Optional ByVal ignoreInternalCaches As Boolean = False, Optional ByVal levelOfDetail As COMPOSITOR_LOD_CLUE = CLC_Generic)
    
    'This function requires two non-optional rects as parameters:
    
    ' - dstViewportRect: the destination rect of the fully composited rect, IN VIEWPORT COORDINATES.  Note that its size is constant if the
    '                     viewport size is constant.  (e.g. on an image larger than the viewport, zoom of 100% and 200% will have identical
    '                     values for this rect)
    ' - srcImageRect: the source rect of the image, IN PDIMAGE COORDINATES.  This is the rectangular region of the image represented by the
    '                  current viewport.  (Remember: if zoomed-out, it will represent a *larger region* than the viewport rect.)
    
    'dstViewportRect is required to infer the current viewport zoom, while srcImageRect is used for intersection testing against
    ' the current layers.  (Layers outside the viewport can be safely ignored.)
    
    'To simplify coordinate calculation, we also generate a bunch of our own rect objects.  The objects are as follows:
    
    'Floating-point rect of the FULL SOURCE LAYER, IN IMAGE COORDINATES, *WITHOUT* AFFINE TRANSFORMS APPLIED
    Dim layerRect As RECTF
    
    'Floating-point rect of the FULL SOURCE LAYER, IN IMAGE COORDINATES, WITH AFFINE TRANSFORMS APPLIED
    Dim layerRectAffine As RECTF
    
    'Floating-point rect of the PORTION OF THE SOURCE LAYER RELEVANT TO THE CURRENT VIEWPORT, IN VIEWPORT COORDINATES
    Dim intRectSrc As RECTF
    
    'Floating-point rect of the DESTINATION OF intRectSrc (above)
    Dim dstRect As RECTF
    
    'Finally, a floating-point rect of intRectSrc, IN 0-BASED 1:1 COORDINATES.  This rect is used to crop-out a portion of the
    ' source layer for pre-processing, if non-destructive FX are active.
    Dim srcRect As RECTF
    
    'Note that some of these rects may only be necessary under certain conditions.
    
    'Vector layers are a special case.  Their temporary compositing DIB may pass a hash test (because it matches the size and
    ' position of the current viewport), even though the layer's contents have changed.  If this is the case, we must still
    ' recomposite the layer.
    '
    'Note that raster layers may also trigger this setting, if their raster contents have changed since a previous cache request.
    Dim cachedDIBValid As Boolean
    
    'To optimize rendering, all temporary DIBs are created with their subpixel offsets already accounted for.  This allows us to use
    ' non-subpixel rendering methods when drawing to the screen, for a huge performance boost, while also making sure that normal and
    ' alternate blend modes do not change a layer's on-screen size by fractional amounts.  A downside of this is that fractional
    ' offsets will not be 100% accurate under some circumstances (e.g. limited viewport scrolling without changing any other image
    ' parameters).  This is an acceptable trade-off for this optimized rendering chain; for idealized results, the "best quality"
    ' viewport preference should be toggled, which will enable full subpixel behavior all the time.
    '
    'Anyway, my point in saying all this is that we will use integer offsets after a temp image has been created.
    Dim xOffsetInt As Long, yOffsetInt As Long
    
    'Although the caller can specify a required interpolation type, we will override this under certain circumstances (e.g. when
    ' creating temporary DIBs, we may use an alternate interpolation type for a speed boost).  Because this happens frequently
    ' throughout the function, we declare a single variable instance here.
    Dim tmpResampleMode As InterpolationMode
    
    'Before doing anything else, we want to see if we can completely skip processing this layer.  A layer is skippable if it does not
    ' intersect the current viewport, but note that we must explicitly account for non-destructive transforms (including complex affine
    ' transforms, like rotate/skew) when calculating the intersection!
    Dim layerAndViewportIntersect As Boolean
    
    'TODO: some kind of fancy function that finds the intersection between a rotated/skewed image's boundary polygon instead of its full rect.
    '      The full rect may indicate overlap, but if the overlap is only a transparent corner region, we can still skip this layer.
    If srcLayer.affineTransformsActive(False) Then
        Layer_Handler.fillRectForLayerF srcLayer, layerRectAffine, True, True
        layerAndViewportIntersect = GDI_Plus.IntersectRectF(intRectSrc, srcImageRect, layerRectAffine)
    
    'Layers without affine transforms use a shortcut intersection approach.
    Else
        Layer_Handler.fillRectForLayerF srcLayer, layerRect, True, False
        layerAndViewportIntersect = GDI_Plus.IntersectRectF(intRectSrc, srcImageRect, layerRect)
    End If
    
    'The rest of this function is skippable if the layer does not overlap the current viewport.
    If layerAndViewportIntersect Then
        
        'The passed layer intersects the viewport, so we must deal with it.
        
        'If this layer is a vector layer, make sure its DIB is properly synchronized.  (If it's *not* a vector layer, this action does nothing.)
        ' This check is important because changes to vector data, e.g. modifying text, will not change any rendering-specific settings, which means
        ' the layer hash won't change.  However, the layer *still needs to be redrawn*, obviously.
        cachedDIBValid = Not srcLayer.syncInternalDIB(levelOfDetail)
        
        'As part of 7.0's comprehensive performance overhaul, we're going to use a dedicated level-of-detail (LOD) DIB stored inside the
        ' source pdLayer object.  Rather than accessing this DIB over and over again, cache a single reference to it now.
        Dim lodDIB As pdDIB
        Set lodDIB = srcLayer.tmpLODDIB(levelOfDetail)
        
        'Several conditions require us to create a temporary DIB prior to actually compositing this layer:
        ' 1) Non-destructive effects
        ' 2) Non-standard blend modes (on layers other than the background layer; the background layer blend mode is ignored)
        ' 3) Non-standard alpha modes (on layers other than the background layer; the background layer alpha mode is ignored)
        ' 4) Affine transformations (rotate, skew, etc).  These may cover complicated sub-sections of the base image, so they require
        '     a lot of extra work; in fact, their calculations form a completely separate rendering pipeline.
        ' 5) In the future, features like layer masks and styles will also affect this calculation.
                
        'First, check for non-standard alpha and blend-modes now, because those features affect all subsequent branches.  (We have to use
        ' PD's internal renderers for those features, as they are not supported by any core Windows libraries.)
        Dim nonStandardBlendingActive As Boolean
        
        If isBaseLayer Then
            nonStandardBlendingActive = False
        Else
            nonStandardBlendingActive = CBool(srcLayer.getLayerBlendMode <> BL_NORMAL) Or CBool(srcLayer.getLayerAlphaMode <> LA_NORMAL)
        End If
        
        'To improve performance, this function also calculates and checks a dedicated "viewport hash" of the layer.  If the layer AND viewport
        ' have not changed since our last render, we can just grab the last temporary viewport object we generated, which makes things like
        ' non-destructive effects and transforms incredibly fast once they've been calculated at least once.
        '
        '(Note that the hash value is valid for any combination of non-destructive FX, non-destructive resizing, and weird blend modes,
        ' by design.)
        '
        'We declare the hash here, but note that we don't calculate it until later in the processing chain.  (The hash includes some
        ' viewport-specific values, which vary depending on subsequent compositing branches.)
        Dim testHash As Long
        
        'We will also be creating a temporary compositing DIB for this layer.  Because coordinates may be floating-point, we need integer
        ' versions that we explicitly size outside the floating-point boundaries, to ensure no cropping takes place.
        Dim dibWidthCalc As Long, dibHeightCalc As Long
        
        'We are now going to branch down two processing paths.  Layers with no complicated affine transforms active are much easier to handle,
        ' as we can mathematically determine a rectangular region of the layer to process.  Layers *with* active affine transforms require
        ' much uglier handling, and a fairly significant hit to processing speed.
        If Not srcLayer.affineTransformsActive(False) Then
            
            'For normal, rectangular layers, the composite operation works similar to StretchBlt, where we calculate a destination rectangle
            ' and a source rectangle.  We know that these two rectangles define the same rectangle of the image, which allows us to infer
            ' zoom (without requiring access to the target canvas object).
            
            'The earlier layerAndViewportIntersect test created a new intersection rect, which tells us where this layer overlaps the
            ' destination viewport rect.  However, the intersect rect was calculated in the layer coordinate space, and we also need
            ' to know where it lies in the viewport coordinate space.  (In many cases, the intersect rect will only cover some subsection
            ' of the viewport, meaning our temporary compositing DIB can also be smaller than the viewport, saving time and memory).
            
            'The getDstRectFromSrcRectF() function uses our already calculated eight bare destination and source values (x/y/width/height)
            ' to calculate a canvas-space destination rect for this layer, e.g. defining the destination region of the viewport covered by this layer.
            getDstRectFromSrcRectF dstRect, intRectSrc, dstViewportRect.Left, dstViewportRect.Top, dstViewportRect.Width, dstViewportRect.Height, srcImageRect.Left, srcImageRect.Top, srcImageRect.Width, srcImageRect.Height
            
            'dstRect and intRectSrc now contain StretchBlt-compatible destination and source rectangles RELATIVE TO THE FULL IMAGE.
            
            'Because the current layer may not be the same size as the full image, we must perform one final translation: a source
            ' rect that represents the source area, relative to the current layer's DIB.  This step is important if the layer has
            ' non-destructive resizing active, because this is when we apply that calculation.
            With srcRect
                .Left = (intRectSrc.Left - CDbl(srcLayer.getLayerOffsetX)) * (1 / srcLayer.getLayerCanvasXModifier)
                .Width = intRectSrc.Width * (1 / srcLayer.getLayerCanvasXModifier)
                .Top = (intRectSrc.Top - CDbl(srcLayer.getLayerOffsetY)) * (1 / srcLayer.getLayerCanvasYModifier)
                .Height = intRectSrc.Height * (1 / srcLayer.getLayerCanvasYModifier)
            End With
            
            'Precalculate dimensions of any temporary DIBs we need to create.  These can have fractional offsets on either side, so we need
            ' to make sure the DIB is at least large enough to contain 0.999 offsets on either side.
            dibWidthCalc = Int(dstRect.Width + (dstRect.Left - Int(dstRect.Left)) + 0.999)
            dibHeightCalc = Int(dstRect.Height + (dstRect.Top - Int(dstRect.Top)) + 0.999)
            
            'With all coordinate math complete, we now proceed with the actual compositing!
            
            'While the layer provides its own hash (representing whether internal layer settings have changed), we also append some viewport-specific
            ' bits to the hash.  This way, if the layer OR the viewport changes, we regenerate the viewport-specific cache for this layer.
            testHash = srcLayer.getViewportHash_Theoretical(srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, dibWidthCalc, dibHeightCalc)
            
            'We know affine transforms (aside from scaling) are not active.  Proceed with our highly optimized rendering chain.
            
            'Non-destructive effects, custom blend-modes, and custom alpha-modes require creation of a temporary DIB, because we must perform
            ' our own compositing (as GDI/GDI+ can't do it for us).  Check for those states now.
            If srcLayer.getLayerNonDestructiveFXState Or nonStandardBlendingActive Then
            
                'One way or another, we will be using integer-based offsets to render the final image.  Calculate them now.
                xOffsetInt = Int(dstRect.Left - dstViewportRect.Left)
                yOffsetInt = Int(dstRect.Top - dstViewportRect.Top)
                
                If (srcLayer.getViewportHash(levelOfDetail) = 0) Or (srcLayer.getViewportHash(levelOfDetail) <> testHash) Or (Not cachedDIBValid) Or ignoreInternalCaches Then
                
                    'The layer and/or viewport has changed since our last composite.  Create a new DIB in the cheapest way possible
                    ' (e.g., if the current temporary DIB size is acceptable, just reset it instead of fully redrawing it)
                    If (lodDIB.getDIBWidth < dibWidthCalc) Or (lodDIB.getDIBHeight < dibHeightCalc) Then
                        lodDIB.createBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
                        lodDIB.setInitialAlphaPremultiplicationState True
                    Else
                        lodDIB.resetDIB
                    End If
                    
                    'Check for the special case of 100% zoom; when this happens, and the image has no non-destructive resizing applied,
                    ' we can completely skip the resize step and simply use BitBlt (way faster)
                    If isRectZoomless And (Not srcLayer.affineTransformsActive(True)) Then
                        
                        BitBlt lodDIB.getDIBDC, 0, 0, dibWidthCalc, dibHeightCalc, srcLayer.layerDIB.getDIBDC, srcRect.Left, srcRect.Top, vbSrcCopy
                        
                        'Apply non-destructive FX now
                        If srcLayer.getLayerNonDestructiveFXState Then applyNDFXToDIB srcLayer, lodDIB
                    
                    'Zoom or non-destructive resize/rotate are active.  Extra work is required.
                    Else
                                    
                        'If zoomed-in, we can switch to StretchBlt for performance gains.  (As nearest-neighbor interpolation will
                        ' be used regardless of the caller's requested interpolation type.)
                        If srcImageRect.Width < dstViewportRect.Width Then
                            
                            'We are zoomed-in to the image, e.g. zoom is larger than 100%.  Check for non-destructive FX being present.
                            
                            'Non-destructive FX are active.  To reduce processing time, chop out the source rectangle, apply the non-destructive FX,
                            ' and *then* resize the data.
                            If srcLayer.getLayerNonDestructiveFXState Then
                            
                                'Extract the smaller source rectangle into its own DIB.  Note that no extra coordinate math is involved here; this is
                                ' a direct transfer.  We need a temp copy of the data because the effect is (per its name!) non-destructive, so it is
                                ' only rendered when the image is written to the screen or to file.
                                
                                'Calculate width and height values for the DIB in advance.  These calculations are complicated because floating-point
                                ' offsets and widths (due to subpixel positioning) may cause the relevant area of the DIB to as much as 2px larger than
                                ' its integer-truncated values.  Also, if these calculated width/height values match the size of the current NDFX DIB,
                                ' we don't need to recreate it!
                                dibWidthCalc = Int(srcRect.Width + (srcRect.Left - Int(srcRect.Left)) + 0.99)
                                dibHeightCalc = Int(srcRect.Height + (srcRect.Top - Int(srcRect.Top)) + 0.99)
                                
                                If (m_NDFXDIB.getDIBWidth < dibWidthCalc) Or (m_NDFXDIB.getDIBHeight < dibHeightCalc) Then
                                    m_NDFXDIB.createBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
                                    m_NDFXDIB.setInitialAlphaPremultiplicationState True
                                Else
                                    m_NDFXDIB.resetDIB
                                End If
                                
                                BitBlt m_NDFXDIB.getDIBDC, 0, 0, dibWidthCalc, dibHeightCalc, srcLayer.layerDIB.getDIBDC, Int(srcRect.Left), Int(srcRect.Top), vbSrcCopy
                                
                                'Apply the non-destructive FX to the smaller source DIB; this is faster than applying it after the resize action,
                                ' as we only have to process a small fraction of the pixels
                                applyNDFXToDIB srcLayer, m_NDFXDIB
                                
                                'Copy the processed DIB into tmpDIB, which will be used for compositing later in this function
                                GDI_Plus.GDIPlus_StretchBlt lodDIB, (dstRect.Left - dstViewportRect.Left) - xOffsetInt, (dstRect.Top - dstViewportRect.Top) - yOffsetInt, dstRect.Width, dstRect.Height, m_NDFXDIB, srcRect.Left - Int(srcRect.Left), srcRect.Top - Int(srcRect.Top), srcRect.Width, srcRect.Height, 1, interpolationType, , Not isBaseLayer
                                                                
                            'Non-destructive FX are not active.  Resize the source rectangle immediately.
                            Else
                            
                                'Paint the resized DIB to the layer, at offset (0, 0), with subpixel offsets already accounted for.
                                ' (If non-destructive resizing is active, we'll also use the layer's specified resample mode instead of
                                '  the function-wide resample mode.)
                                If (srcLayer.getLayerCanvasXModifier <> 1) Or (srcLayer.getLayerCanvasYModifier <> 1) Then tmpResampleMode = srcLayer.getLayerResizeQuality_GDIPlus Else tmpResampleMode = InterpolationModeNearestNeighbor
                                GDIPlus_StretchBlt lodDIB, (dstRect.Left - dstViewportRect.Left) - xOffsetInt, (dstRect.Top - dstViewportRect.Top) - yOffsetInt, dstRect.Width, dstRect.Height, srcLayer.layerDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, 1, tmpResampleMode, , Not isBaseLayer
                                
                            End If
                        
                        'We are zoomed-out, so StretchBlt cannot be used (as it will cause artifacting for 32bpp data).  GDI+ is our only
                        ' option at present.
                        Else
                            
                            'Paint the resized DIB to the layer, at offset (0, 0), with subpixel offsets already accounted for.
                            ' (If non-destructive resizing is active, we'll also use the layer's specified resample mode instead of
                            '  the function-wide resample mode.)
                            If (srcLayer.getLayerCanvasXModifier <> 1) Or (srcLayer.getLayerCanvasYModifier <> 1) Then tmpResampleMode = srcLayer.getLayerResizeQuality_GDIPlus Else tmpResampleMode = interpolationType
                            GDIPlus_StretchBlt lodDIB, (dstRect.Left - dstViewportRect.Left) - xOffsetInt, (dstRect.Top - dstViewportRect.Top) - yOffsetInt, dstRect.Width, dstRect.Height, srcLayer.layerDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, 1, tmpResampleMode, , Not isBaseLayer
                                
                            'Apply non-destructive FX now
                            If srcLayer.getLayerNonDestructiveFXState Then applyNDFXToDIB srcLayer, lodDIB
                            
                        End If
                        
                    End If
                    
                    'Update the layer hash, so we can possibly skip this work in the future
                    srcLayer.setViewportHash_FromLong levelOfDetail, testHash
                    
                'End If for regenerating temporary layer compositing level-of-detail DIB (lodDIB)
                End If
                
                'lodDIB now contains the chunk of this layer that appears on the viewport, with all non-destructive edits applied,
                ' and subpixel positioning already accounted for.
                
                ' Time to composite it!
                If nonStandardBlendingActive Then
                    
                    'Composite and blend in one fell swoop
                    compositeAndBlendDIBs lodDIB, m_dstDIB, srcLayer.getLayerBlendMode, xOffsetInt, yOffsetInt, srcLayer.getLayerOpacity / 100, srcLayer.getLayerAlphaMode
                    
                'If no wacky blend modes are in use, we can use GDI's AlphaBlend for a performance boost.
                ' (Note that this is only possible because subpixel offsets have already been accounted for, when creating the temporary DIB.)
                Else
                    lodDIB.alphaBlendToDC m_dstDIB.getDIBDC, srcLayer.getLayerOpacity * 2.55, xOffsetInt, yOffsetInt
                End If
                
            'This is a simple layer, with no non-destructive effects or weird blend modes.  Composite it instantly, without using
            ' a temporary placeholder layer.
            Else
                
                'If the rect is zoomless, it means we don't have to perform subpixel positioning.  GDI's AlphaBlend is thus perfectly
                ' acceptable, and it will be much faster (5-10x) than GDI+.
                If isRectZoomless And (srcLayer.getLayerCanvasXModifier = 1) And (srcLayer.getLayerCanvasYModifier = 1) Then
                    
                    'AlphaBlend is very picky about invalid width/height values.  Make sure the source rect is valid.
                    ' (It may be ever-so-slightly off due to rounding issues during all the floating-point math we perform.)
                    'If srcRect.Width > srcLayer.getLayerWidth(False) Then srcRect.Width = srcLayer.getLayerWidth(False)
                    'If srcRect.Height > srcLayer.getLayerHeight(False) Then srcRect.Height = srcLayer.getLayerHeight(False)
                    'srcLayer.layerDIB.alphaBlendToDCEx m_dstDIB.getDIBDC, dstRect.Left - dstViewportRect.Left, dstRect.Top - dstViewportRect.Top, dstRect.Width, dstRect.Height, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, srcLayer.getLayerOpacity * 2.55
                    
                    'I have temporarily disabled the use of GDI's AlphaBlend function here, as it's causing inexplicable rendering failures
                    ' on certain image/viewport size combinations.
                    GDIPlus_StretchBlt m_dstDIB, dstRect.Left - dstViewportRect.Left, dstRect.Top - dstViewportRect.Top, dstRect.Width, dstRect.Height, srcLayer.layerDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, srcLayer.getLayerOpacity / 100, InterpolationModeNearestNeighbor
                    
                Else
                    
                    If srcRect.Width > srcLayer.getLayerWidth(False) Then srcRect.Width = srcLayer.getLayerWidth(False)
                    If srcRect.Height > srcLayer.getLayerHeight(False) Then srcRect.Height = srcLayer.getLayerHeight(False)
                    
                    'When zoomed in, nearest neighbor is preferred.  This is fast enough that we don't need to use a temporary DIB cache.
                    If srcImageRect.Width < dstViewportRect.Width Then
                        GDIPlus_StretchBlt m_dstDIB, dstRect.Left - dstViewportRect.Left, dstRect.Top - dstViewportRect.Top, dstRect.Width, dstRect.Height, srcLayer.layerDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, srcLayer.getLayerOpacity / 100, InterpolationModeNearestNeighbor
                    
                    'When zoomed out, use the supplied interpolation mode.  The caller typically toggles this according to PD's global
                    ' performance settings.
                    Else
                        
                        'One way or another, we will be using integer-based offsets to render the final image.  Calculate them now.
                        xOffsetInt = Int(dstRect.Left - dstViewportRect.Left)
                        yOffsetInt = Int(dstRect.Top - dstViewportRect.Top)
                        
                        'See if our existing temporary DIB hash is valid; if it is, we can proceed straight to rendering!
                        If (srcLayer.getViewportHash(levelOfDetail) = 0) Or (srcLayer.getViewportHash(levelOfDetail) <> testHash) Or (Not cachedDIBValid) Or ignoreInternalCaches Then
                        
                            'The layer and/or viewport has changed since our last composite.  Create a new DIB in the cheapest way possible
                            ' (e.g., if the current temporary DIB size is acceptable, just reset it instead of fully redrawing it)
                            If (lodDIB.getDIBWidth < dibWidthCalc) Or (lodDIB.getDIBHeight < dibHeightCalc) Then
                                lodDIB.createBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
                                lodDIB.setInitialAlphaPremultiplicationState True
                            Else
                                lodDIB.resetDIB 0
                            End If
                            
                            'Paint the resized DIB to the layer, at offset (0, 0), with subpixel offsets already accounted for.
                            ' (If non-destructive resizing is active, we'll also use the layer's specified resample mode instead of
                            '  the function-wide resample mode.)
                            If (srcLayer.getLayerCanvasXModifier <> 1) Or (srcLayer.getLayerCanvasYModifier <> 1) Then tmpResampleMode = srcLayer.getLayerResizeQuality_GDIPlus Else tmpResampleMode = interpolationType
                            GDIPlus_StretchBlt lodDIB, (dstRect.Left - dstViewportRect.Left) - xOffsetInt, (dstRect.Top - dstViewportRect.Top) - yOffsetInt, dstRect.Width, dstRect.Height, srcLayer.layerDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, 1, tmpResampleMode, , Not isBaseLayer
                            
                            'Update the layer hash, so we can possibly skip this work in the future
                            srcLayer.setViewportHash_FromLong levelOfDetail, testHash
                            
                        End If
                        
                        'Because subpixel offsets have already been accounted for, we can use hardware-accelerated alpha blend here.
                        lodDIB.alphaBlendToDC m_dstDIB.getDIBDC, srcLayer.getLayerOpacity * 2.55, xOffsetInt, yOffsetInt
                        
                    'End zoom in vs zoom out check
                    End If
                                            
                'End zoomless rect check (e.g. toggling GDI's AlphaBlend)
                End If
                
            End If
            
            'End handling of layers with no non-scaling affine transforms
        
        'One or more non-scaling affine transforms are active (e.g. rotation).  Special coordinate calculations are required.
        Else
        
            'For this first test, all affine-transformed layers are going to be forcibly composited to a DIB the size of the current viewport.
            ' We will optimize this heavily in the future, but for now, I just want to get it working.
            
            'Start by retrieving the corner coordinates of the transformed layer.  We need these values to perform a parallelogram-style blt.
            srcLayer.getLayerCornerCoordinates m_PlgPoints, True
            
            'Next, we want to create a copy of those layer corner points, as transformed into the destination coordinate space.
            
            'Start by setting up some conversion factors between the source and destination spaces
            Dim srcCalcWidth As Double, srcCalcHeight As Double
            srcCalcWidth = srcImage.Width
            srcCalcHeight = srcImage.Height
            
            Dim viewportScaleX As Double, viewportScaleY As Double
            viewportScaleX = srcCalcWidth * (dstViewportRect.Width / srcImageRect.Width)
            viewportScaleY = srcCalcHeight * (dstViewportRect.Height / srcImageRect.Height)
            
            'This viewport pipeline operates almost exclusively in the *image* coordinate space - *not* the *layer* coordinate space -
            ' because the layer area is always the same (the full layer).  To that end, we can use a shortcut calculation to solve the
            ' where the image boundaries lie in the canvas coordinate space, because we don't need to account for the usual combination
            ' of source x/y (instead, we can assume 0, 0).  These shift offsets let us know how much to offset the layer parallelogram
            ' corner coordinates prior to rendering.
            Dim viewportShiftX As Double, viewportShiftY As Double
            viewportShiftX = (-srcImageRect.Left / srcImageRect.Width) * dstViewportRect.Width
            viewportShiftY = (-srcImageRect.Top / srcImageRect.Height) * dstViewportRect.Height
            
            'Use all the calculations we just made to create a new copy of the parallelogram coordinates, in the destination coordinate space
            Dim i As Long
            For i = 0 To 3
                With m_PlgPointsDst(i)
                    .x = (m_PlgPoints(i).x / srcCalcWidth) * viewportScaleX + viewportShiftX
                    .y = (m_PlgPoints(i).y / srcCalcHeight) * viewportScaleY + viewportShiftY
                End With
            Next i
            
            'TODO: more complicated calculation of temporary compositing DIB size/position.
            ' For now, make it the size of the viewport.
            '
            'Precalculate dimensions of any temporary DIBs we need to create.  These can have fractional offsets on either side, so we need
            ' to make sure the DIB is at least large enough to contain 0.99 offsets on either side.
            'dibWidthCalc = dstRect.Width + 0.999 + (dstRect.Left - Int(dstRect.Left))
            'dibHeightCalc = dstRect.Height + 0.999 + (dstRect.Top - Int(dstRect.Top))
            dibWidthCalc = Int(dstViewportRect.Width + 0.9999)
            dibHeightCalc = Int(dstViewportRect.Height + 0.9999)
            
            'While the layer provides its own hash (representing whether internal layer settings have changed), we also append some viewport-specific
            ' bits to the hash.  This way, if the layer OR the viewport changes, we regenerate the viewport-specific cache for this layer.  However,
            ' if both stay the same, we can simply re-use our previous rendering, which saves tons of time if layers other than the current one
            ' have active affine transforms (as we don't need to re-process them!).
            testHash = srcLayer.getViewportHash_Theoretical(m_PlgPointsDst(0).x, m_PlgPointsDst(0).y, m_PlgPointsDst(1).x, m_PlgPointsDst(1).y, m_PlgPointsDst(2).x, m_PlgPointsDst(2).y, interpolationType)
            
            'This function differs from the standard pipeline because we *always* create a temporary compositing DIB.  The speed hit from
            ' a complex affine transform is significant, especially on large images, so caching the transform (even for normal blendmode and
            ' no NDFX) is always preferable.
            
            'Check the viewport hash now.  Certain conditions can force us to ignore internal caches (e.g. generating thumbnails).
            If (srcLayer.getViewportHash(levelOfDetail) = 0) Or (srcLayer.getViewportHash(levelOfDetail) <> testHash) Or (Not cachedDIBValid) Or ignoreInternalCaches Then
                
                'The layer and/or viewport has changed since our last composite.  Create a new DIB in the cheapest way possible
                ' (e.g., if the current temporary DIB size is acceptable, just reset it instead of fully redrawing it)
                If (lodDIB.getDIBWidth < dibWidthCalc) Or (lodDIB.getDIBHeight < dibHeightCalc) Then
                    lodDIB.createBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
                    lodDIB.setInitialAlphaPremultiplicationState True
                Else
                    lodDIB.resetDIB
                End If
                
                'Render the affine-transformed DIB to the temporary compositing DIB, at 100% opacity and normal blend mode.
                ' (Variable opacity and blend mode will be handled later.)
                '
                'Also, note that we branch according to the performance preference passed to this function.  By design, this is a different
                ' approach than we use in the normal, rectangular rendering branch.  Because affine transforms can be extremely performance-intensive,
                ' the user needs a way to override the layer's default stretching behavior.  (Note that, as usual, saving to file, copying to
                ' clipboard, etc, *always* use the layer's settings, not the viewport ones.)
                If interpolationType = InterpolationModeNearestNeighbor Then
                    GDIPlus_PlgBlt lodDIB, m_PlgPointsDst, srcLayer.layerDIB, 0, 0, srcLayer.getLayerWidth(False), srcLayer.getLayerHeight(False), 1, InterpolationModeNearestNeighbor
                Else
                    GDIPlus_PlgBlt lodDIB, m_PlgPointsDst, srcLayer.layerDIB, 0, 0, srcLayer.getLayerWidth(False), srcLayer.getLayerHeight(False), 1, srcLayer.getLayerResizeQuality_GDIPlus
                End If
                    
                'If non-destructive FX are active, apply them now
                If srcLayer.getLayerNonDestructiveFXState Then applyNDFXToDIB srcLayer, lodDIB
                
                'Update the layer hash, so we can possibly skip this work in the future
                srcLayer.setViewportHash_FromLong levelOfDetail, testHash
                
            End If
            
            'lodDIB now contains the chunk of this layer that appears on the viewport, with the affine transform and any/all
            ' non-destructive edits applied to it.  The only thing left to do is composite it!
                
            'Non-standard blend modes must be handled by our internal compositor
            If nonStandardBlendingActive Then
                compositeAndBlendDIBs lodDIB, m_dstDIB, srcLayer.getLayerBlendMode, 0, 0, srcLayer.getLayerOpacity / 100, srcLayer.getLayerAlphaMode
                
            'If no wacky blend modes are in use, we can use GDI's AlphaBlend for a performance boost
            Else
                lodDIB.alphaBlendToDC m_dstDIB.getDIBDC, srcLayer.getLayerOpacity * 2.55, 0, 0
            End If
                
        End If
        
    'END block for layer intersecting the current viewport.  (Note that we simply ignore layers that don't intersect the viewport.)
    End If

End Sub

'Given a layer containing one or more non-destructive effect parameters, and a target DIB, apply the layer's specified FX against
' the DIB.  Why supply the DIB separately?  Because during viewport composition, all layers are resized and/or cropped to their
' size relevant to the viewport.  By doing this, we hugely decrease the performance hit from non-destructive edit calculations.
Public Sub applyNDFXToDIB(ByRef parentLayer As pdLayer, ByRef targetDIB As pdDIB)

    'Start by unpremultiplying the layer alpha
    If (targetDIB.getDIBColorDepth = 32) And targetDIB.getAlphaPremultiplication Then targetDIB.setAlphaPremultiplication False

    'Calculate each non-destructive FX in turn
    
    'Brightness, contrast, and clarity are handled simultaneously, to reduce processing time
    If (parentLayer.getLayerNonDestructiveFXValue(NDFX_EXPOSURE) <> 0) Or (parentLayer.getLayerNonDestructiveFXValue(NDFX_CONTRAST) <> 0) Or (parentLayer.getLayerNonDestructiveFXValue(NDFX_CLARITY) <> 0) Then
        fastLayerExposureContrastClarity parentLayer, targetDIB
    End If
    
    'Vibrance, temperature, tint
    If (parentLayer.getLayerNonDestructiveFXValue(NDFX_VIBRANCE) <> 0) Or (parentLayer.getLayerNonDestructiveFXValue(NDFX_TEMPERATURE) <> 0) Or (parentLayer.getLayerNonDestructiveFXValue(NDFX_TINT) <> 0) Then
        fastLayerVibranceTemperatureTint parentLayer, targetDIB
    End If
    
    'Re-premultiply alpha
    If targetDIB.getDIBColorDepth = 32 Then targetDIB.setAlphaPremultiplication True

End Sub

'Super-fast, no-frills adjustment for Exposure, Contrast, and Clarity.
' This function was designed for non-destructive compositing purposes, so it takes some shortcuts that PhotoDemon's dedicated functions
' for these operations do not.  As such, do not expect it to output identical results.
Private Sub fastLayerExposureContrastClarity(ByRef targetLayer As pdLayer, ByRef targetDIB As pdDIB)

    'Create a local array and point it at the target layer's DIB
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, targetDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
    
    'Local loop variables
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = targetDIB.getDIBWidth - 1
    finalY = targetDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = targetDIB.getDIBColorDepth \ 8
    
    'Adjustment values
    Dim exposureAdjustment As Double, contrastAdjustment As Long, clarityAdjustment As Long
    exposureAdjustment = targetLayer.getLayerNonDestructiveFXValue(NDFX_EXPOSURE) / 2
    contrastAdjustment = targetLayer.getLayerNonDestructiveFXValue(NDFX_CONTRAST)
    clarityAdjustment = targetLayer.getLayerNonDestructiveFXValue(NDFX_CLARITY)
    
    'Look-up tables are used to accelerate this function.
    Dim eccLUT() As Byte
    ReDim eccLUT(0 To 255) As Byte
    Dim tmpVal As Double
    
    For x = 0 To 255
        
        'Calculate exposure
        If exposureAdjustment <> 0 Then
            
            tmpVal = x / 255
            tmpVal = tmpVal * 2 ^ (exposureAdjustment)
            tmpVal = tmpVal * 255
            
            'Crop the lookup value to [0, 255] range
            If tmpVal > 255 Then
                tmpVal = 255
            ElseIf tmpVal < 0 Then
                tmpVal = 0
            End If
        
        Else
            tmpVal = x
        End If
                
        'Calculate contrast
        If contrastAdjustment <> 0 Then
            
            tmpVal = tmpVal + (((tmpVal - 127) * contrastAdjustment) \ 100)
        
            'Crop the lookup value to [0, 255] range
            If tmpVal > 255 Then
                tmpVal = 255
            ElseIf tmpVal < 0 Then
                tmpVal = 0
            End If
        
        End If
        
        'Calculate clarity.  Clarity is simply a contrast adjustment limited to midtones.  Values at 127 are processed
        ' most strongly, with a linear decrease as input values approach 0 or 255.  Also, I reduce the strength of the
        ' adjustment by 20% to prevent blowout or gray-washing (for high or low adjustments, respectively).
        If clarityAdjustment <> 0 Then
            
            If x < 127 Then
                tmpVal = tmpVal + (tmpVal / 127) * (((tmpVal - 127) * clarityAdjustment) \ 100) * 0.8
            Else
                tmpVal = tmpVal + ((255 - tmpVal) / 127) * (((tmpVal - 127) * clarityAdjustment) \ 100) * 0.8
            End If
            
            'Crop the lookup value to [0, 255] range
            If tmpVal > 255 Then
                tmpVal = 255
            ElseIf tmpVal < 0 Then
                tmpVal = 0
            End If
        
        End If
        
        eccLUT(x) = tmpVal
        
    Next x
        
    'Now we can loop through each pixel in the image, converting values as we go
    For x = 0 To finalX
        QuickVal = x * qvDepth
    For y = 0 To finalY
            
        'Apply the changes in one fell swoop
        ImageData(QuickVal, y) = eccLUT(ImageData(QuickVal, y))
        ImageData(QuickVal + 1, y) = eccLUT(ImageData(QuickVal + 1, y))
        ImageData(QuickVal + 2, y) = eccLUT(ImageData(QuickVal + 2, y))
        
    Next y
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    
End Sub

'Super-fast, no-frills adjustments for Vibrance, Temperature, and Tint.
' This function was designed for non-destructive compositing purposes, so it takes some shortcuts that PhotoDemon's dedicated functions
' for these operations do not.  As such, do not expect it to output identical results.
Private Sub fastLayerVibranceTemperatureTint(ByRef targetLayer As pdLayer, ByRef targetDIB As pdDIB)

    'Create a local array and point it at the target layer's DIB
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, targetDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
    
    'Local loop variables
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = targetDIB.getDIBWidth - 1
    finalY = targetDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = targetDIB.getDIBColorDepth \ 8
    
    'Adjustment values
    Dim vibranceAdjustment As Double, tempAdjustment As Double, tintAdjustment As Double
    vibranceAdjustment = -0.02 * targetLayer.getLayerNonDestructiveFXValue(NDFX_VIBRANCE)
    tempAdjustment = targetLayer.getLayerNonDestructiveFXValue(NDFX_TEMPERATURE) / 5
    tintAdjustment = targetLayer.getLayerNonDestructiveFXValue(NDFX_TINT) / 5
    
    'Color and related variables
    Dim r As Long, g As Long, b As Long, maxVal As Long
    Dim amtVal As Double, avgVal As Double
    
    'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
    Dim grayLUT() As Long
    ReDim grayLUT(0 To 765) As Long
    For x = 0 To 765
        grayLUT(x) = x \ 3
    Next x
    
    'To handle temperature and tint, we will also be preparing RGB LUT tables
    Dim rLUT() As Byte, gLUT() As Byte, bLUT() As Byte
    ReDim rLUT(0 To 255) As Byte
    ReDim gLUT(0 To 255) As Byte
    ReDim bLUT(0 To 255) As Byte
    
    'Populate the look-up table
    For x = 0 To 255
    
        'Calculate temperature
        If tempAdjustment = 0 Then
            r = x
            g = x
            b = x
        Else
        
            'Temperature affects the red and blue channels
            r = x + tempAdjustment
            g = x
            b = x - tempAdjustment
        
        End If
        
        'Calculate tint.  Note that RGB will have been already set in the previous step, so we work from the already
        ' calculated values (instead of x).
        If tintAdjustment <> 0 Then
        
            'Tint affects the green channel
            g = g + tintAdjustment
        
        End If
        
        'Clip RGB
        If r < 0 Then
            r = 0
        ElseIf r > 255 Then
            r = 255
        End If
        
        If g < 0 Then
            g = 0
        ElseIf g > 255 Then
            g = 255
        End If
        
        If b < 0 Then
            b = 0
        ElseIf b > 255 Then
            b = 255
        End If
        
        'Fill the look-up table
        rLUT(x) = Round(r)
        gLUT(x) = Round(g)
        bLUT(x) = Round(b)
    
    Next x
        
    'Now we can loop through each pixel in the image, converting values as we go
    For x = 0 To finalX
        QuickVal = x * qvDepth
    For y = 0 To finalY
            
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        If vibranceAdjustment <> 0 Then
        
            'Calculate the gray value using the look-up table
            avgVal = grayLUT(r + g + b)
            maxVal = Max3Int(r, g, b)
            
            'Get adjusted average
            amtVal = ((Abs(maxVal - avgVal) / 127) * vibranceAdjustment)
            
            'Apply new vibrance
            If r <> maxVal Then r = r + (maxVal - r) * amtVal
            If g <> maxVal Then g = g + (maxVal - g) * amtVal
            If b <> maxVal Then b = b + (maxVal - b) * amtVal
            
            'Clamp values to [0,255] range
            If r < 0 Then r = 0
            If r > 255 Then r = 255
            If g < 0 Then g = 0
            If g > 255 Then g = 255
            If b < 0 Then b = 0
            If b > 255 Then b = 255
            
        End If
            
        'Apply the new values
        ImageData(QuickVal + 2, y) = rLUT(r)
        ImageData(QuickVal + 1, y) = gLUT(g)
        ImageData(QuickVal, y) = bLUT(b)
        
    Next y
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    
End Sub

'Given StretchBlt-like parameters, create a destination rect that reflects the transformation of a source rect
' from the source coordinate space to the destination coordinate space.
Private Function getDstRectFromSrcRect(ByRef dstRect As RECT, ByRef srcRect As RECT, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long)
    
    'Convert the source rectangle coordinates into a [0, 1] scale relative to the total source area.
    ' Note that we use a RECTF here, which typically uses WIDTH and HEIGHT markers - however, we use them as RIGHT and BOTTOM here,
    ' so ignore the original named values.
    Dim srcRectF As RECTF
    With srcRectF
        .Left = (srcRect.Left - srcX) / srcWidth
        .Width = (srcRect.Right - srcX) / srcWidth
        .Top = (srcRect.Top - srcY) / srcHeight
        .Height = (srcRect.Bottom - srcY) / srcHeight
    End With
    
    'Using those source values, fill the destination rect
    With dstRect
        .Left = dstX + (srcRectF.Left * dstWidth)
        .Right = dstX + (srcRectF.Width * dstWidth)
        .Top = dstY + (srcRectF.Top * dstHeight)
        .Bottom = dstY + (srcRectF.Height * dstHeight)
    End With

End Function

'Given StretchBlt-like parameters, create a destination rect that reflects the transformation of a source rect
' from the source coordinate space to the destination coordinate space.
Private Function getDstRectFromSrcRectF(ByRef dstRect As RECTF, ByRef srcRect As RECTF, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single)
    
    'To make our calculations easier, we're going to convert the source rectangle coordinates to a [0, 1] scale, relative to the total source area.
    Dim tmpRectF As RECTF
    
    If (srcWidth = 0) Or (srcHeight = 0) Then
        Exit Function
    End If
    
    With tmpRectF
        .Left = (srcRect.Left - srcX) / srcWidth
        .Width = srcRect.Width / srcWidth
        .Top = (srcRect.Top - srcY) / srcHeight
        .Height = srcRect.Height / srcHeight
    End With
    
    'Using those source values, fill the destination rect
    With dstRect
        .Left = dstX + (tmpRectF.Left * dstWidth)
        .Width = tmpRectF.Width * dstWidth
        .Top = dstY + (tmpRectF.Top * dstHeight)
        .Height = tmpRectF.Height * dstHeight
    End With

End Function

'In order to render complex images quickly, this class must generate a lot of temporary objects.  If PD is under memory pressure, call this function
' to forcibly release as many temporary objects as possible.  (It should be obvious, but subsequent render calls will result in temporary objects
' being recreated, as necessary.)
Public Sub attemptToFreeMemory()
    
    'Start by clearing our own internal caches
    m_dstDIB.eraseDIB
    m_tmpDIB.eraseDIB
    m_NDFXDIB.eraseDIB
    
    'Ask the pixel blender to release memory too, if it can
    m_Blender.attemptToFreeMemory
    
End Sub

Private Sub Class_Initialize()
    Set m_dstDIB = New pdDIB
    Set m_tmpDIB = New pdDIB
    Set m_NDFXDIB = New pdDIB
    Set m_Blender = New pdPixelBlender
    
    ReDim m_PlgPoints(0 To 3) As POINTFLOAT
    ReDim m_PlgPointsDst(0 To 3) As POINTFLOAT
End Sub

Private Sub Class_Terminate()
    Set m_dstDIB = Nothing
    Set m_tmpDIB = Nothing
    Set m_NDFXDIB = Nothing
    Set m_Blender = Nothing
End Sub

