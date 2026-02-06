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
'Copyright 2014-2026 by Tanner Helland
'Created: 01/May/14
'Last updated: 22/February/19
'Last update: allow outside callers to manage their own scratch layer when invoking _PaintOp functions
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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

'If a layer has non-destructive, non-standard transformations active (e.g. rotation, skew), we use a special, parallelogram-based rendering function.
' To cut down on variable declarations, we declare this array once, and initialize it when a compositor instance is initialized.
Private m_PlgPoints() As PointFloat

'For the rect-specific compositor, we also need a copy of the parallelogram points in the *canvas*/viewport coordinate space
Private m_PlgPointsDst() As PointFloat

'Historically, pdCompositor handled its own blending, but as part of wide-ranging compositor optimizations in the 7.0 release, blending was
' moved to a dedicated class.
Private m_Blender As pdPixelBlender

'To avoid miscounted object references when passing object pointers, we need to manually increment ref counts.
' Thank you to Bonnie West for the original breakdown of this technique, available here (link good as of Feb '19):
' http://www.vbforums.com/showthread.php?707879-VB6-Dereferencing-Pointers-sans-CopyMemory
Private Declare Function ObjSetAddRef Lib "msvbvm60" Alias "__vbaObjSetAddref" (ByRef objDest As Object, ByVal pObject As Long) As Long

'Composite two DIBs using the requested blend mode.  An offset can be specified for the top DIB, relative to the bottom DIB.
' Per the function name, the *top DIB* is treated as the destination.  After compositing, the caller would need to alpha-blend
' the finished top DIB over the bottom DIB to achieve full composition.
'
'This function is currently unused.  I cannot guarantee that it will exist in future releases.
'Friend Sub CompositeDIBsWithoutBlending(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal blendMode As PD_BlendMode, ByVal dstX As Single, ByVal dstY As Single, Optional ByVal alphaMode As PD_AlphaMode = AM_Normal)
'
'    'The pixel blender class handles this for us
'    m_Blender.BlendDIBs topDIB, bottomDIB, blendMode, dstX, dstY, 1, alphaMode, False
'
'End Sub

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
Friend Sub CompositeAndBlendDIBs(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal blendMode As PD_BlendMode, ByVal dstX As Single, ByVal dstY As Single, Optional ByVal alphaModifier As Double = 1#, Optional ByVal bottomAlphaMode As PD_AlphaMode = AM_Normal, Optional ByVal topAlphaMode As PD_AlphaMode = AM_Normal, Optional ByVal ptrToTopLayerAlternateRectF As Long = 0, Optional ByRef topLayerMask As pdDIB = Nothing)
    
    'The blend class handles this for us
    m_Blender.BlendDIBs topDIB, bottomDIB, blendMode, dstX, dstY, alphaModifier, bottomAlphaMode, topAlphaMode, True, ptrToTopLayerAlternateRectF, topLayerMask
    
End Sub

'Merge two layers together.  Note this can be used to merge any two arbitrary layers, with the bottom layer holding the result
' of the merge.  It is up to the caller to deal with any subsequent layer deletions, etc - this sub just performs the merge.
'
'The optional parameter, "bottomLayerIsFullSize", should be set to TRUE if the bottom layer is the size of the image.  This saves
' some processing time, because we don't have to check for rect intersection.
Friend Sub MergeLayers(ByRef topLayer As pdLayer, ByRef bottomLayer As pdLayer)

    Dim xOffset As Double, yOffset As Double
    Dim i As Long
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'If the layers are vector layers, make sure their DIBs are properly synchronized.  (If the layers are not vector layers,
    ' this action does nothing.)
    topLayer.SyncInternalDIB
    bottomLayer.SyncInternalDIB
    
    'Next, we need to find the union rect of the two layers.  This is the smallest rectangle that holds both layers.
    Dim bottomRect As RectF, topRect As RectF, finalRect As RectF
    Layers.FillRectForLayerF bottomLayer, bottomRect, True, True
    Layers.FillRectForLayerF topLayer, topRect, True, True
    UnionRectF finalRect, topRect, bottomRect
    
    'finalRect now contains the coordinates of the union rect.  (Note that some settings allow us to bypass
    ' this rect entirely, e.g. if the bottom layer is image-sized and the top layer has no active affine transforms.)
    
    'Next, solve for the base layer being the precise size of the image.  (This has a high probability
    ' when being called from certain functions - like MergeVisibleLayers() - and we may be able to skip
    ' a bunch of redundant memory allocations by reusing the existing base layer DIB as the merge target
    ' instead of allocating a new DIB at the size of the union rect.)
    Dim bottomLayerIsFullSize As Boolean
    bottomLayerIsFullSize = (bottomLayer.GetLayerOffsetX = 0) And (bottomLayer.GetLayerOffsetY = 0)
    bottomLayerIsFullSize = bottomLayerIsFullSize And (bottomLayer.GetLayerWidth(True) = PDImages.GetActiveImage.Width)
    bottomLayerIsFullSize = bottomLayerIsFullSize And (bottomLayer.GetLayerHeight(True) = PDImages.GetActiveImage.Height)
    bottomLayerIsFullSize = bottomLayerIsFullSize And (Not bottomLayer.AffineTransformsActive(True))
    
    'We now need to perform a failsafe check on the "bottomLayerIsFullSize" parameter we just calculated.
    ' We can't actually use it as a viable shortcut path if the top layer is *larger* than the base layer;
    ' when this happens, we still need to create a new DIB to store the result, because the bottom layer
    ' (even though it is the full size of the image) is too small to hold the merged layer (which will be
    ' *larger* than image itself).
    If bottomLayerIsFullSize Then
        
        Dim bottomLayerTooSmall As Boolean
        bottomLayerTooSmall = (topRect.Left < 0) Or (topRect.Top < 0)
        bottomLayerTooSmall = bottomLayerTooSmall Or ((topRect.Left + topRect.Width) > bottomRect.Width)
        bottomLayerTooSmall = bottomLayerTooSmall Or ((topRect.Top + topRect.Height) > bottomRect.Height)
        
        'If the bottom layer is too small, disable the optimization flag and proceed normally.
        If bottomLayerTooSmall Then bottomLayerIsFullSize = False
        
    End If
    
    'If the bottom layer is the size of the image itself (e.g. when flattening the image), we can use its DIB as the target DIB instead of
    ' creating a temporary one.  This can provide a nice performance boost depending on the complexity of subsequent layers.
    If bottomLayerIsFullSize Then
    
        xOffset = topRect.Left
        yOffset = topRect.Top
        
    'The top and bottom layer sizes are totally independent.  This makes our life somewhat unpleasant.
    Else
        
        'Create a blank DIB at the dimensions of the union rect.
        tmpDIB.CreateBlank finalRect.Width, finalRect.Height, 32, 0
        tmpDIB.SetInitialAlphaPremultiplicationState True
        
        'We now need to do a couple of things.  Let's start by copying the bottom DIB into this new temporary DIB.
        xOffset = bottomRect.Left - finalRect.Left
        yOffset = bottomRect.Top - finalRect.Top
        
        'If no non-destructive resizes are active, we can simply copy the layer over as-is
        If (Not bottomLayer.AffineTransformsActive(True)) Then
        
            If (bottomLayer.GetLayerOpacity = 100) Then
                With bottomLayer
                    GDI.BitBltWrapper tmpDIB.GetDIBDC, xOffset, yOffset, .GetLayerWidth(False), .GetLayerHeight(False), .GetLayerDIB.GetDIBDC, 0, 0, vbSrcCopy
                End With
            Else
                With bottomLayer
                    .GetLayerDIB.AlphaBlendToDC tmpDIB.GetDIBDC, .GetLayerOpacity * 2.55, xOffset, yOffset
                End With
            End If
        
        'Active transforms make this step harder.
        Else
            
            'If the only active transform is rescaling, we can use a shortcut StretchBlt function for a performance boost.
            If (Not bottomLayer.AffineTransformsActive(False)) Then
            
                With bottomLayer
                    GDI_Plus.GDIPlus_StretchBlt tmpDIB, xOffset, yOffset, .GetLayerWidth(True), .GetLayerHeight(True), .GetLayerDIB, 0, 0, .GetLayerWidth(False), .GetLayerHeight(False), .GetLayerOpacity / 100, .GetLayerResizeQuality_GDIPlus
                End With
            
            'If rotation or other affine transforms are active, a more complicated solution is required.
            Else
            
                'First, retrieve the layer's coordinates in the default image coordinate space
                bottomLayer.GetLayerCornerCoordinates m_PlgPoints
                
                'Next, we need to convert those to the union rect coordinate space
                For i = 0 To 3
                    m_PlgPoints(i).x = (m_PlgPoints(i).x - finalRect.Left)
                    m_PlgPoints(i).y = (m_PlgPoints(i).y - finalRect.Top)
                Next i
                
                'Render the transformed DIB now.
                With bottomLayer
                    GDI_Plus.GDIPlus_PlgBlt tmpDIB, m_PlgPoints, .GetLayerDIB, 0, 0, .GetLayerWidth(False), .GetLayerHeight(False), .GetLayerOpacity / 100, .GetLayerResizeQuality_GDIPlus
                End With
            
            End If
            
            'With the bottom layer successfully rendered onto the destination DIB, we can reset all non-destructive size modifiers now.
            bottomLayer.SetLayerCanvasXModifier 1#
            bottomLayer.SetLayerCanvasYModifier 1#
            bottomLayer.SetLayerAngle 0#
            bottomLayer.SetLayerShearX 0#
            bottomLayer.SetLayerShearY 0#
            
        End If
        
        'We now need to calculate a new layer offset for this temporary DIB, which will eventually be copied into the bottom layer.
        ' (Without this, the parent image won't know where to position the layer!)
        bottomLayer.SetLayerOffsetX finalRect.Left
        bottomLayer.SetLayerOffsetY finalRect.Top
        
        'Copy the temporary DIB into the bottom layer.  (We can't simply reference it as we're likely to need a temporary DIB for
        ' processing the top layer as well.)
        bottomLayer.GetLayerDIB.CreateFromExistingDIB tmpDIB
        tmpDIB.EraseDIB
        
        'Reset the layer's opacity, as we have already applied it in the previous steps.
        bottomLayer.SetLayerOpacity 100
        
        'Calculate new offsets for the top layer, then carry on with business as usual!
        xOffset = topRect.Left - finalRect.Left
        yOffset = topRect.Top - finalRect.Top
        
    End If
    
    'The bottom layer DIB now represents the target for the final layer merge.  Any non-destructive effects
    ' have been processed, its opacity has been accounted for (and it's tracking value reset to 100),
    ' and any non-destructive transformations (resize, scaling, etc) have been made permanent.
    
    'Now all we need to do is merge the top layer onto it.
    
    'If the top layer has no non-destructive resizing active, this step is simple.
    If (Not topLayer.AffineTransformsActive(True)) Then
        
        'We don't even need a temporary DIB - just composite the current layer as-is
        If ((topLayer.GetLayerBlendMode = BM_Normal) And (topLayer.GetLayerAlphaMode = AM_Normal)) Then
            topLayer.GetLayerDIB.AlphaBlendToDC bottomLayer.GetLayerDIB.GetDIBDC, topLayer.GetLayerOpacity * 2.55, xOffset, yOffset
        Else
            CompositeAndBlendDIBs topLayer.GetLayerDIB, bottomLayer.GetLayerDIB, topLayer.GetLayerBlendMode, xOffset, yOffset, topLayer.GetLayerOpacity / 100, bottomLayer.GetLayerAlphaMode, topLayer.GetLayerAlphaMode
        End If
    
    'Non-destructive transforms are active, so we must prep a temporary DIB for the top layer (containing a transformed copy of its DIB)
    Else
    
        'Rescale-only transforms can use a shortcut copy method
        If (Not topLayer.AffineTransformsActive(False)) Then
            
            With topLayer
                tmpDIB.CreateBlank .GetLayerWidth(True), .GetLayerHeight(True), 32, 0
                tmpDIB.SetInitialAlphaPremultiplicationState True
                GDIPlusResizeDIB tmpDIB, 0, 0, .GetLayerWidth(True), .GetLayerHeight(True), .GetLayerDIB, 0, 0, .GetLayerWidth(False), .GetLayerHeight(False), .GetLayerResizeQuality_GDIPlus
            End With
        
        'Full affine transforms are hairier
        Else
            
            'Ask the layer to give us a copy of its transformed DIB
            Dim xOffsetLong As Long, yOffsetLong As Long
            topLayer.GetAffineTransformedDIB tmpDIB, xOffsetLong, yOffsetLong
            
            'Calculate new offsets, using the x/y returned by that function (which may be slightly modified
            ' in order to preserve subpixel positioning)
            xOffset = xOffsetLong - finalRect.Left
            yOffset = yOffsetLong - finalRect.Top
            
        End If
        
        'With all non-destructive options accounted for, we can finally calculate composited layer data!
        If (topLayer.GetLayerBlendMode = BM_Normal) And (topLayer.GetLayerAlphaMode = AM_Normal) Then
            tmpDIB.AlphaBlendToDC bottomLayer.GetLayerDIB.GetDIBDC, topLayer.GetLayerOpacity * 2.55, xOffset, yOffset
        Else
            CompositeAndBlendDIBs tmpDIB, bottomLayer.GetLayerDIB, topLayer.GetLayerBlendMode, xOffset, yOffset, topLayer.GetLayerOpacity / 100, bottomLayer.GetLayerAlphaMode, topLayer.GetLayerAlphaMode
        End If
                
    End If
    
    'The two layers have been merged successfully!  Any further actions (like deleting the top layer) must be handled by the caller.
    
End Sub

'Merge two layers together using paintbrush rules.  In paintbrush rules, the bottom layer retains all of its properties (blend mode,
' opacity, etc), with the exception of non-destructive transforms (which are made permanent).  Paint merges are also unique because
' we don't look at the top layer's boundaries for the intersection rect - instead, we use the rect supplied by the paintbrush engine,
' which is the region of the paint layer where painting has occurred.  (In most cases, this will be significantly smaller than the
' full layer rect - and note that it will already be cropped to the image dimensions, as necessary.)
'
'Unlike the regular MergeLayers(), this sub *does not* handle vector layers, by design.  Instead of a merge,
' paint ops on a vector layer are placed into a completely new raster layer.
Friend Sub MergeLayers_PaintStyle(ByRef paintLayer As pdLayer, ByRef bottomLayer As pdLayer, Optional ByVal ptrToPaintbrushRectF As Long = 0, Optional ByRef parentImage As pdImage = Nothing)
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    Dim xOffset As Double, yOffset As Double
    Dim i As Long
    
    Dim tmpDIB As pdDIB
    
    'Start by finding the union rect of the two layers.  This is the smallest rectangle that holds both layers.
    ' (Note that this calculation occurs in the *image's* coordinate space.)
    Dim bottomRect As RectF, topRect As RectF, finalRect As RectF
    Layers.FillRectForLayerF bottomLayer, bottomRect, True, True
    
    If (ptrToPaintbrushRectF = 0) Then
        Layers.FillRectForLayerF paintLayer, topRect, True, True
    Else
        CopyMemoryStrict VarPtr(topRect), ptrToPaintbrushRectF, LenB(topRect)
    End If
    
    PDMath.UnionRectF finalRect, topRect, bottomRect
    
    'finalRect now contains the coordinates of the union rect.  (Note that some settings allow us to bypass this rect entirely,
    ' e.g. if the bottom layer is image-sized and the top layer has no active affine transforms.)
    
    'If the bottom layer is the size of the image itself, we can use its DIB as the target DIB
    ' instead of creating a temporary one.  This provides a nice performance boost.
    Dim bottomLayerIsFullSize As Boolean
    If (Not parentImage Is Nothing) Then
        With bottomLayer
            bottomLayerIsFullSize = ((.GetLayerOffsetX = 0) And (.GetLayerOffsetY = 0) And (.GetLayerDIB.GetDIBWidth = parentImage.Width) And (.GetLayerDIB.GetDIBHeight = parentImage.Height))
        End With
    End If
    
    'An additional caveat is the bottom layer not having any active affine transforms; if it does, we need to commit
    ' those before merging the scratch layer atop it.
    bottomLayerIsFullSize = (bottomLayerIsFullSize And (Not bottomLayer.AffineTransformsActive(True)))
    
    If bottomLayerIsFullSize Then
        xOffset = topRect.Left
        yOffset = topRect.Top
    
    'The bottom layer is not full-sized.  At present, the paint scratch layer is *always* full-sized, so we need to prepare a
    ' new layer DIB for the bottom layer.
    Else
        
        'First, see if the paint strokes lie fully inside the target layer.  If they do, we don't need to create a new DIB.
        ' (Note that this step is skipped if the bottom layer has affine transforms active - if that's the case, we have
        ' to create a new DIB anyway, without all the active transforms.)
        Dim paintStrokesInside As Boolean: paintStrokesInside = False
        
        If ((ptrToPaintbrushRectF <> 0) And (Not bottomLayer.AffineTransformsActive(True))) Then
            If (topRect.Left >= bottomRect.Left) And (topRect.Top >= bottomRect.Top) Then
                If (topRect.Left + topRect.Width) <= (bottomRect.Left + bottomRect.Width) Then
                    If (topRect.Top + topRect.Height) <= (bottomRect.Top + bottomRect.Height) Then
                        paintStrokesInside = True
                    End If
                End If
            End If
        End If
        
        'If the paint strokes fit inside the target layer, we can use it as-is without the need for a temporary merge layer
        If paintStrokesInside Then
            
            'Calculate new offsets for the top layer, then carry on with business as usual!
            xOffset = topRect.Left - bottomRect.Left
            yOffset = topRect.Top - bottomRect.Top
            
        'If the paint strokes do not lie inside the target layer, we have no choice but to create a new layer large
        ' enough to hold the results of the merge.
        Else
            
            Set tmpDIB = New pdDIB
            tmpDIB.CreateBlank finalRect.Width, finalRect.Height, 32, 0
            tmpDIB.SetInitialAlphaPremultiplicationState True
            
            'We now need to do a couple of things.  Let's start by copying the bottom DIB into this new temporary DIB.
            xOffset = bottomRect.Left - finalRect.Left
            yOffset = bottomRect.Top - finalRect.Top
            
            'If no non-destructive resizes are active, we can simply copy the layer over as-is
            If (Not bottomLayer.AffineTransformsActive(True)) Then
            
                With bottomLayer
                    GDI.BitBltWrapper tmpDIB.GetDIBDC, xOffset, yOffset, .GetLayerWidth(False), .GetLayerHeight(False), .GetLayerDIB.GetDIBDC, 0, 0, vbSrcCopy
                End With
                
            'Active transforms must be made permanent prior to a merge.
            Else
                
                'If the only active transform is rescaling, we can use a shortcut StretchBlt function for a performance boost.
                If (Not bottomLayer.AffineTransformsActive(False)) Then
                
                    With bottomLayer
                        GDI_Plus.GDIPlus_StretchBlt tmpDIB, xOffset, yOffset, .GetLayerWidth(True), .GetLayerHeight(True), .GetLayerDIB, 0, 0, .GetLayerWidth(False), .GetLayerHeight(False), 1#, .GetLayerResizeQuality_GDIPlus, , , , True
                    End With
                
                'If rotation or other affine transforms are active, a more complicated solution is required.
                Else
                
                    'First, retrieve the layer's coordinates in the default image coordinate space
                    bottomLayer.GetLayerCornerCoordinates m_PlgPoints
                    
                    'Next, we need to convert those to the union rect coordinate space
                    For i = 0 To 3
                        m_PlgPoints(i).x = (m_PlgPoints(i).x - finalRect.Left)
                        m_PlgPoints(i).y = (m_PlgPoints(i).y - finalRect.Top)
                    Next i
                    
                    'Render the transformed DIB now.
                    With bottomLayer
                        GDI_Plus.GDIPlus_PlgBlt tmpDIB, m_PlgPoints, .GetLayerDIB, 0, 0, .GetLayerWidth(False), .GetLayerHeight(False), 1#, .GetLayerResizeQuality_GDIPlus
                    End With
                
                End If
                
                'With the bottom layer successfully rendered onto the destination DIB, we can reset all non-destructive size modifiers now.
                bottomLayer.SetLayerCanvasXModifier 1#
                bottomLayer.SetLayerCanvasYModifier 1#
                bottomLayer.SetLayerAngle 0#
                bottomLayer.SetLayerShearX 0#
                bottomLayer.SetLayerShearY 0#
                
            End If
                
            'We now need to calculate a new layer offset for this temporary DIB, which will eventually be copied into the bottom layer.
            ' Without this, the main composite won't know where to stick the layer!
            bottomLayer.SetLayerOffsetX finalRect.Left
            bottomLayer.SetLayerOffsetY finalRect.Top
            
            'Assign this temporary DIB as the new bottom layer DIB.
            bottomLayer.SetLayerDIB tmpDIB
            
            'Calculate new offsets for the top layer, then carry on with business as usual!
            xOffset = topRect.Left - finalRect.Left
            yOffset = topRect.Top - finalRect.Top
            
        End If
        
    End If
    
    'The bottom layer DIB now represents the target for the final layer merge.  Any non-destructive transforms have been made permanent,
    ' and it has been resized to account for the union with the paintbrush layer.
    
    'Next, we need to merge the top layer onto it.
    
    'Before merging the top layer, check for an active selection.  If one exists, we need to preprocess the paint layer against it.
    If (Not parentImage Is Nothing) Then
        If parentImage.IsSelectionActive Then
            
            'A selection is active.  Pre-mask the paint scratch layer against it.
            m_Blender.ApplyMaskToTopDIB paintLayer.GetLayerDIB, parentImage.MainSelection.GetCompositeMaskDIB, ptrToPaintbrushRectF
            
        End If
    End If
    
    'Unlike regular layers, we know that the paintbrush layer will never have any active transforms, so we can simply composite the top
    ' layer as-is (without checking for active transforms, like stretching or rotation).
    Dim quickBlendOkay As Boolean
    quickBlendOkay = ((paintLayer.GetLayerBlendMode = BM_Normal) And (paintLayer.GetLayerAlphaMode = AM_Normal) And (Not (bottomLayer.GetLayerAlphaMode = AM_Locked)))
    
    If quickBlendOkay Then
        GDIPlus_StretchBlt bottomLayer.GetLayerDIB, xOffset, yOffset, topRect.Width, topRect.Height, paintLayer.GetLayerDIB, topRect.Left, topRect.Top, topRect.Width, topRect.Height, paintLayer.GetLayerOpacity * 0.01, GP_IM_NearestNeighbor, , True
    Else
        
        'If *either* the paintbrush or the target layer has requested locked alpha, honor that now
        Dim targetAlphaMode As PD_AlphaMode
        If ((paintLayer.GetLayerAlphaMode = AM_Locked) Or (bottomLayer.GetLayerAlphaMode = AM_Locked)) Then
            targetAlphaMode = AM_Locked
        Else
            targetAlphaMode = AM_Normal
        End If
            
        Me.CompositeAndBlendDIBs paintLayer.GetLayerDIB, bottomLayer.GetLayerDIB, paintLayer.GetLayerBlendMode, xOffset, yOffset, paintLayer.GetLayerOpacity * 0.01, targetAlphaMode, paintLayer.GetLayerAlphaMode, VarPtr(topRect)
        
    End If
    
    'The paintbrush layer has been merged successfully!  Any further actions (like resetting the paint layer) must be handled by the caller.
    PDDebug.LogAction "Merged two layers paint-style.  Time taken: " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub

'Given two DIBs of equal size, perform a fast merge with variable blend mode and opacity.
'
'Previously, PD required you to create temporary layers for each DIB, assign them blend and alpha modes, then manually perform
' a full merge (with checks for non-destructive effects, sizing, etc).
'
'Now, if all you want is two merged DIBs, use this function.
'
'IMPORTANT NOTE!  Both the top and bottom DIBs will be modified by this function, per standard merge rules.  The top DIB must be modified because it
'                  contains the intermediate results of the merge.  The bottom DIB must be modified because it contains the final result.  As such,
'                  use temp copies of DIBs you can't afford to lose.
Friend Sub QuickMergeTwoDibsOfEqualSize(ByRef bottomDIB As pdDIB, ByRef topDIB As pdDIB, Optional ByVal blendMode As PD_BlendMode = BM_Normal, Optional ByVal topLayerOpacity As Double = 100#, Optional ByVal bottomAlphaMode As PD_AlphaMode = AM_Normal, Optional ByVal topAlphaMode As PD_AlphaMode = AM_Normal)
    Me.CompositeAndBlendDIBs topDIB, bottomDIB, blendMode, 0, 0, topLayerOpacity / 100, bottomAlphaMode, topAlphaMode
End Sub

'Returns all layers of the image as a single, composited image (in pdDIB format, of course).  Because of the way VB handles
' object references, we ask the calling function to supply the DIB they want filled.  Optionally, they can also request a
' particular premultiplication status of the composited DIB's alpha values.  (This is helpful for save functions, which
' require non-premultiplied alpha, vs viewport functions, which require premultiplied alpha).
'
'As of 7.0, an optional scratch layer index can be specified.  If >= 0, the source image's scratch layer will be rendered
' atop the layer at the requested position.  (For example, if the index is 0, the scratch layer will be rendered atop the
' background layer.)
Friend Sub GetCompositedImage(ByRef srcImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal premultiplicationStatus As Boolean = True, Optional ByVal paintScratchLayerIndex As Long = -1)
    
    'Start by creating the destination DIB, as necessary
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    'Next, size the destination DIB to the match the composited image
    If (dstDIB.GetDIBWidth <> srcImage.Width) Or (dstDIB.GetDIBHeight <> srcImage.Height) Then
        dstDIB.CreateBlank srcImage.Width, srcImage.Height, 32
        dstDIB.SetInitialAlphaPremultiplicationState True
    Else
        dstDIB.ResetDIB
    End If
    
    'Subsequent steps are going to use a helper function.  This helper function makes it simple to "insert" paint operations
    ' into the current layer stack, without fucking up the existing layer order.
    
    'If the image has additional layers, proceed to merge the rest of them, starting from the bottom and working our way up.
    ' Note that if a layer is invisible, we simply skip it - this is the most performance-friendly way to handle them.
    Dim i As Long
    For i = 0 To srcImage.GetNumOfLayers - 1
        
        'Make sure the layer is visible
        If srcImage.GetLayerByIndex(i).GetLayerVisibility Then
            
            'Pass this layer to the helper function.  It will take care of the compositing step for us.
            GetCompositedImage_LayerHelper dstDIB, srcImage.GetLayerByIndex(i), i
        
        'End layer visibility check
        End If
        
        'If a scratch layer index has been specified, and it matches this layer index, paint it now
        If (paintScratchLayerIndex = i) Then
            GetCompositedImage_LayerHelper dstDIB, srcImage.ScratchLayer, -1
        End If
        
    Next i
        
    'If the user requested non-premultiplied alpha, calculate it now.
    ' (By default, this function always returns a premultiplied image, because that's what the compositor functions return.)
    If premultiplicationStatus Then
        dstDIB.SetInitialAlphaPremultiplicationState True
    Else
        If (dstDIB.GetDIBColorDepth = 32) Then dstDIB.SetAlphaPremultiplication False
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
Private Sub GetCompositedImage_LayerHelper(ByRef dstDIB As pdDIB, ByRef srcLayer As pdLayer, ByVal srcLayerIndex As Long)
    
    'If this layer is a vector layer, make sure its DIB is properly synchronized.  (If it's *not* a vector layer, this action does nothing.)
    srcLayer.SyncInternalDIB
    
    'Custom blend-modes and alpha-modes are available to all layers but the base one.  If the normal blend mode is active, we can shortcut
    ' a lot of the compositing pipeline.
    Dim standardBlendingActive As Boolean
    
    If (srcLayerIndex = 0) Then
        standardBlendingActive = True
    Else
        standardBlendingActive = (srcLayer.GetLayerBlendMode = BM_Normal) And (srcLayer.GetLayerAlphaMode = AM_Normal)
    End If
        
    'Next, check for custom blend modes.  These require an intermediate DIB copy of the image, because PD must internally render the
    ' blend modes (as GDI+ can't do it for us).
    If standardBlendingActive Then
    
        'Custom blend modes are not active.  Yay!
        
        'Check for affine transformations.  If none are present (e.g. the layer is at its original size and orientation),
        ' we can use GDI's AlphaBlend for a performance boost.
        If (Not srcLayer.AffineTransformsActive(True)) Then
        
            With srcLayer
                .GetLayerDIB.AlphaBlendToDC dstDIB.GetDIBDC, .GetLayerOpacity * 2.55, .GetLayerOffsetX, .GetLayerOffsetY
            End With
        
        Else
            
            'One or more affine transformations are active.  GDI+ must be used to render the image.
            With srcLayer
            
                'If scaling is the only active affine transformation, we can use a special, optimized render function
                If (Not srcLayer.AffineTransformsActive(False)) Then
                    GDI_Plus.GDIPlus_StretchBlt dstDIB, .GetLayerOffsetX, .GetLayerOffsetY, .GetLayerWidth(True), .GetLayerHeight(True), .GetLayerDIB, 0, 0, .GetLayerWidth(False), .GetLayerHeight(False), .GetLayerOpacity / 100, .GetLayerResizeQuality_GDIPlus
                
                'If non-scaling affine transforms are active, we must perform a full affine transformation.
                Else
                    
                    'Retrieve the layer's destination position into our POINTFLOAT array
                    .GetLayerCornerCoordinates m_PlgPoints
                    
                    'Apply a full parallelogram transformation
                    GDI_Plus.GDIPlus_PlgBlt dstDIB, m_PlgPoints, .GetLayerDIB, 0, 0, .GetLayerWidth(False), .GetLayerHeight(False), .GetLayerOpacity / 100, .GetLayerResizeQuality_GDIPlus
                    
                End If
                
            End With
        
        End If
    
    'This Else statement represents the case "If (srcLayer.GetLayerBlendMode <> BM_Normal)..."
    Else
    
        'Custom blend modes are active.  We have no choice but to create a temporary compositing DIB.
        With srcLayer
        
            'If non-destructive transforms are NOT active, pass the source layer to the all-in-one composite + blend function
            If (Not .AffineTransformsActive(True)) Then
            
                'Perform an all-in-one composite and blend using PD's awesome all-in-one function
                CompositeAndBlendDIBs .GetLayerDIB, dstDIB, .GetLayerBlendMode, .GetLayerOffsetX, .GetLayerOffsetY, .GetLayerOpacity / 100, , .GetLayerAlphaMode
            
            'One or more affine transforms are present.  We have no choice but to create a temporary intermediate DIB,
            ' prior to full compositing.
            Else
            
                'The layer itself provides a highly optimized method for extracting the affine-transformed portion of its DIB
                Dim intLayerOffsetX As Long, intLayerOffsetY As Long
                Dim layerBoundariesOkay As Boolean
                layerBoundariesOkay = .GetAffineTransformedDIB(m_tmpDIB, intLayerOffsetX, intLayerOffsetY, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight)
                
                'Next, perform an all-in-one composite + blend using our awesome new compositeAndBlendDIBs function
                If layerBoundariesOkay Then CompositeAndBlendDIBs m_tmpDIB, dstDIB, .GetLayerBlendMode, intLayerOffsetX, intLayerOffsetY, .GetLayerOpacity / 100, , .GetLayerAlphaMode
                
            End If
            
        End With
        
    End If
        
End Sub

'Returns a subsection of the fully composited image (in pdDIB format, of course).  This is helpful for rendering the main viewport,
' as we only composite the relevant portions of the image.
'
'This function is large, complicated, and unfortunately tricky, as we must resize and composite each layer in turn.  However, it's the secret
' to PD's incredible viewport performance.  (e.g., this function is how we can outperform both GIMP and Paint.NET despite being single-threaded
' and heavily CPU-bound)
'
'As of 7.0, an optional scratch layer index can be specified.  If >= 0, the source image's scratch layer will be rendered atop the layer
' at the requested position.  (For example, if the index is 0, the scratch layer will be rendered atop the background layer.)
Friend Sub GetCompositedRect(ByRef srcImage As pdImage, ByRef dstDIB As pdDIB, ByRef dstViewportRect As RectF, ByRef srcImageRect As RectF, ByVal interpolationType As GP_InterpolationMode, Optional ByVal ignoreInternalCaches As Boolean = False, Optional ByVal levelOfDetail As PD_CompositorLOD = CLC_Generic, Optional ByVal paintScratchLayerIndex As Long = -1, Optional ByVal ptrToAlternateScratch As Long = 0&)
        
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
    isRectZoomless = (dstViewportRect.Width = srcImageRect.Width) And (dstViewportRect.Height = srcImageRect.Height)
    
    'Start by creating a temporary DIB the size of the destination image (viewport).  All compositing will be done to this DIB,
    ' and as the final step, we will AlphaBlend the finished image onto dstDIB.
    Dim dibWidthCalc As Long, dibHeightCalc As Long
    dibWidthCalc = Int(dstViewportRect.Width + PDMath.Frac(dstViewportRect.Left) + 0.99)
    dibHeightCalc = Int(dstViewportRect.Height + PDMath.Frac(dstViewportRect.Top) + 0.99)
    If (dibWidthCalc < 1) Then dibWidthCalc = 1
    If (dibHeightCalc < 1) Then dibHeightCalc = 1
    
    'Note that we try to avoid recreating this DIB unless absolutely necessary
    If (m_dstDIB.GetDIBWidth < dibWidthCalc) Or (m_dstDIB.GetDIBHeight < dibHeightCalc) Then
        m_dstDIB.CreateBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
        m_dstDIB.SetInitialAlphaPremultiplicationState True
    Else
        m_dstDIB.ResetDIB 0
    End If
    
    'If a scratch layer is being composited (e.g. a paint operation is in progress), calculate the correct interpolation mode now.
    ' (To improve responsiveness while painting, the painted layer is assigned a different, "temporary" interpolation op in
    '  certain performance modes.)
    Dim paintLayerInterpMode As GP_InterpolationMode
    If (paintScratchLayerIndex >= 0) Then
        
        'While we're here, initialize the relevant scratch layer's LOD DIB.  This DIB will be used for temporary compositing
        ' of the paint stroke against its base layer, and then that composited result will be blended onto the layer stack
        ' with all of the appropriate layer properties assigned (blend mode, opacity, etc).
        If (srcImage.ScratchLayer.TmpLODDIB(CLC_Painting).GetDIBWidth < dibWidthCalc) Or (srcImage.ScratchLayer.TmpLODDIB(CLC_Painting).GetDIBHeight < dibHeightCalc) Then
            srcImage.ScratchLayer.TmpLODDIB(CLC_Painting).CreateBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
            srcImage.ScratchLayer.TmpLODDIB(CLC_Painting).SetInitialAlphaPremultiplicationState True
        Else
            srcImage.ScratchLayer.TmpLODDIB(CLC_Painting).ResetDIB 0
        End If
    
        If (g_ViewportPerformance = PD_PERF_FASTEST) Then
            paintLayerInterpMode = GP_IM_NearestNeighbor
        ElseIf (g_ViewportPerformance = PD_PERF_BALANCED) Then
            paintLayerInterpMode = GP_IM_Bilinear
        Else
            paintLayerInterpMode = interpolationType
        End If
        
    End If
    
    'Now our work is pretty simple: iterate layers, and pass visible ones to helper functions (which handle the messy
    ' work of solving all releveant coordinate math, and performing the actual painting steps)
    Dim i As Long
    For i = 0 To srcImage.GetNumOfLayers - 1
        
        'Only process a layer if it is currently visible.
        If srcImage.GetLayerByIndex(i).GetLayerVisibility Then
            
            'We now have to check for active paint strokes.  If the user is currently painting, we want to render the paint stroke onto
            ' the current layer before compositing it.  (This is crucial if a non-standard blend or alpha mode is being used, for example,
            ' because we need to composite the stroke and base layer *prior* to applying blend and/or alpha calculations.)
            '
            'Note that we only do this if the current layer is a raster layer.  If it is a vector layer, the paint stroke will simply
            ' be added as a new layer, so we don't need to pre-merge anything.
            If (paintScratchLayerIndex = i) And srcImage.GetLayerByIndex(i).IsLayerRaster Then
                    
                'Pass this paint layer to a dedicated renderer (one specially optimized for real-time performance,
                ' so different trade-offs are made compared to the standard renderer).
                GetCompositedRect_PaintOp m_dstDIB, srcImage, srcImage.GetLayerByIndex(i), dstViewportRect, srcImageRect, interpolationType, isRectZoomless, ptrToAlternateScratch
                
            'Painting is not active; apply a standard composite op
            Else
                GetCompositedRect_LayerHelper m_dstDIB, srcImage, srcImage.GetLayerByIndex(i), dstViewportRect, srcImageRect, interpolationType, isRectZoomless, (i = 0), ignoreInternalCaches, levelOfDetail
            End If
            
        'END block for layer visibility toggle
        End If
        
        'If a scratch layer index has been specified, and the target layer is *not* a raster layer,
        ' composite the paint stroke atop it as a standalone layer.
        If (paintScratchLayerIndex = i) Then
            If (Not srcImage.GetLayerByIndex(i).IsLayerRaster) Then
                
                'Selections require us to pre-mask our selection result
                If srcImage.IsSelectionActive Then
                    GetCompositedRect_LayerHelper m_dstDIB, srcImage, srcImage.ScratchLayer, dstViewportRect, srcImageRect, paintLayerInterpMode, isRectZoomless, False, ignoreInternalCaches, levelOfDetail, srcImage.MainSelection.GetCompositeMaskDIB_ViewportCopy(), ptrToAlternateScratch
                Else
                    GetCompositedRect_LayerHelper m_dstDIB, srcImage, srcImage.ScratchLayer, dstViewportRect, srcImageRect, paintLayerInterpMode, isRectZoomless, False, ignoreInternalCaches, levelOfDetail, Nothing, ptrToAlternateScratch
                End If
                
            End If
        End If
        
    Next i
    
    'With all layers successfully blended onto tmpLayer's DIB, we can now perform a final alphaBlend onto dstDIB.
    ' Because of compositor rewrites in v7.0, we can use GDI's AlphaBlend regardless of zoom, because the module-level DIB was
    ' constructed with subpixel positioning already accounted for.
    '
    'For this release, I have left the original GDI+ paint method, in case I find it necessary to revert back in the future.
    With dstViewportRect
    
        If isRectZoomless Then
            m_dstDIB.AlphaBlendToDCEx dstDIB.GetDIBDC, .Left, .Top, .Width, .Height, 0, 0, .Width, .Height
        Else
            
            Dim bltWidth As Single, bltHeight As Single
            bltWidth = .Width + PDMath.Frac(.Left)
            bltHeight = .Height + PDMath.Frac(.Top)
            GDIPlus_StretchBlt dstDIB, Int(.Left), Int(.Top), bltWidth, bltHeight, m_dstDIB, 0!, 0!, bltWidth, bltHeight, 1!, GP_IM_NearestNeighbor, disableEdgeFix:=True
            
            'While we might be able to get away with GDI's AlphaBlend here, it may cause bleeding on the trailing width/height edges,
            ' which can cause the underlying checkerboard to show through.  (Also, the performance benefit isn't that meaningful.)
            'm_dstDIB.AlphaBlendToDCEx dstDIB.GetDIBDC, Int(.Left), Int(.Top), Int(bltWidth), Int(bltHeight), 0, 0, Int(bltWidth), Int(bltHeight)
            
        End If
    
    End With
    
    'The target DIB will always be premultiplied, by design
    dstDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'This helper function is only used internally, by the GetCompositedRect function.
'
'The purpose of this helper function is to take a source layer, and merge it onto a destination DIB (*not* a destination layer).
' This logic is handled separately so that paint operations can be silently "inserted" into the existing layer stack, without requiring
' any modifications of the parent pdImage object.
'
'Extensive caching has been implemented to maximize performance.  Each individual pdLayer object stores its own viewport cache, but note that
' the optional ignoreInternalCaches can be set to TRUE, in which case these caches must be forcibly regenerated.  If the caches do not need to
' be forcibly generated, the optional levelOfDetail parameter should be correctly set, which clues the compositor into which LOD cache to use
' in the target layer, greatly improving performance.
Private Sub GetCompositedRect_LayerHelper(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByRef dstViewportRect As RectF, ByRef srcImageRect As RectF, ByVal interpolationType As GP_InterpolationMode, Optional ByVal isRectZoomless As Boolean = False, Optional ByVal isBaseLayer As Boolean = False, Optional ByVal ignoreInternalCaches As Boolean = False, Optional ByVal levelOfDetail As PD_CompositorLOD = CLC_Generic, Optional ByRef maskInViewportSpace As pdDIB = Nothing, Optional ByVal ptrToAlternateDIB As Long = 0&)
    
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
    Dim layerRect As RectF
    
    'Floating-point rect of the FULL SOURCE LAYER, IN IMAGE COORDINATES, WITH AFFINE TRANSFORMS APPLIED
    Dim layerRectAffine As RectF
    
    'Floating-point rect of the PORTION OF THE SOURCE LAYER RELEVANT TO THE CURRENT VIEWPORT, IN VIEWPORT COORDINATES
    Dim intRectSrc As RectF
    
    'Floating-point rect of the DESTINATION OF intRectSrc (above)
    Dim dstRect As RectF
    
    'Finally, a floating-point rect of intRectSrc, IN 0-BASED 1:1 COORDINATES.  This rect is used to crop-out a portion of the
    ' source layer for pre-processing, if non-destructive FX are active.
    Dim srcRect As RectF
    
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
    Dim tmpResampleMode As GP_InterpolationMode
    
    'Before doing anything else, we want to see if we can completely skip processing this layer.  A layer is skippable if it does not
    ' intersect the current viewport, but note that we must explicitly account for non-destructive transforms (including complex affine
    ' transforms, like rotate/skew) when calculating the intersection!
    Dim layerAndViewportIntersect As Boolean
    
    'TODO: some kind of fancy function that finds the intersection between a rotated/skewed image's boundary polygon instead of its full rect.
    '      The full rect may indicate overlap, but if the overlap is only a transparent corner region, we can still skip this layer.
    If srcLayer.AffineTransformsActive(False) Then
        Layers.FillRectForLayerF srcLayer, layerRectAffine, True, True
        layerAndViewportIntersect = GDI_Plus.IntersectRectF(intRectSrc, srcImageRect, layerRectAffine)
    
    'Layers without affine transforms use a shortcut intersection approach.
    Else
        Layers.FillRectForLayerF srcLayer, layerRect, True, False
        layerAndViewportIntersect = GDI_Plus.IntersectRectF(intRectSrc, srcImageRect, layerRect)
    End If
    
    'The rest of this function is skippable if the layer does not overlap the current viewport.
    If layerAndViewportIntersect Then
        
        'The passed layer intersects the viewport, so we must deal with it.
        
        'If this layer is a vector layer, make sure its DIB is properly synchronized.  (If it's *not* a vector layer, this action does nothing.)
        ' This check is important because changes to vector data, e.g. modifying text, will not change any rendering-specific settings, which means
        ' the layer hash won't change.  However, the layer *still needs to be redrawn*, obviously.
        cachedDIBValid = Not srcLayer.SyncInternalDIB(levelOfDetail)
        
        'As part of 7.0's comprehensive performance overhaul, we're going to use a dedicated level-of-detail (LOD) DIB stored inside the
        ' source pdLayer object.  Rather than accessing this DIB over and over again, cache a single reference to it now.
        Dim lodDIB As pdDIB
        Set lodDIB = srcLayer.TmpLODDIB(levelOfDetail)
        
        Dim srcDIB As pdDIB
        Set srcDIB = srcLayer.GetLayerDIB
        
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
            nonStandardBlendingActive = (srcLayer.GetLayerBlendMode <> BM_Normal) Or (srcLayer.GetLayerAlphaMode <> AM_Normal)
        End If
        
        'If a mask was supplied, also treat this as a non-standard blend
        If (Not nonStandardBlendingActive) Then nonStandardBlendingActive = (Not maskInViewportSpace Is Nothing)
        
        'Before proceeding further, see if the caller supplied their own alternate scratch surface.  Some tools
        ' (e.g. gradient tools) are capable of rendering their own tool previews in viewport coordinate space.
        ' This provides a large perf boost when a tool is actively applied to a large image while zoomed-out.
        ' If such an alternate layer is provided, we can skip all our expensive internal caching strategies,
        ' and instead, simply apply the source layer as-is.
        If (ptrToAlternateDIB <> 0) Then
            
            'Make sure ref count gets updated
            Dim alternateScratch As pdDIB
            ObjSetAddRef alternateScratch, ptrToAlternateDIB
            
            'Paint using settings supplied by whatever layer was passed into the function
            If nonStandardBlendingActive Then
                CompositeAndBlendDIBs alternateScratch, m_dstDIB, srcLayer.GetLayerBlendMode, 0!, 0!, srcLayer.GetLayerOpacity / 100!, , srcLayer.GetLayerAlphaMode, , maskInViewportSpace
            Else
                alternateScratch.AlphaBlendToDC m_dstDIB.GetDIBDC, srcLayer.GetLayerOpacity * 2.55, 0, 0
            End If
            
            Exit Sub
                
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
        If (Not srcLayer.AffineTransformsActive(False)) Then
            
            'For normal, rectangular layers, the composite operation works similar to StretchBlt, where we calculate a destination rectangle
            ' and a source rectangle.  We know that these two rectangles define the same rectangle of the image, which allows us to infer
            ' zoom (without requiring access to the target canvas object).
            
            'The earlier layerAndViewportIntersect test created a new intersection rect, which tells us where this layer overlaps the
            ' destination viewport rect.  However, the intersect rect was calculated in the layer coordinate space, and we also need
            ' to know where it lies in the viewport coordinate space.  (In many cases, the intersect rect will only cover some subsection
            ' of the viewport, meaning our temporary compositing DIB can also be smaller than the viewport, saving time and memory).
            
            'The getDstRectFromSrcRectF() function uses our already calculated eight bare destination and source values (x/y/width/height)
            ' to calculate a canvas-space destination rect for this layer, e.g. defining the destination region of the viewport covered by this layer.
            GetDstRectFromSrcRectF dstRect, intRectSrc, dstViewportRect.Left, dstViewportRect.Top, dstViewportRect.Width, dstViewportRect.Height, srcImageRect.Left, srcImageRect.Top, srcImageRect.Width, srcImageRect.Height
            
            'dstRect and intRectSrc now contain StretchBlt-compatible destination and source rectangles RELATIVE TO THE FULL IMAGE.
            
            'Because the current layer may not be the same size as the full image, we must perform one final translation: a source
            ' rect that represents the source area, relative to the current layer's DIB.  This step is important if the layer has
            ' non-destructive resizing active, because this is when we apply that calculation.
            Dim srcLayerXDivisor As Double, srcLayerYDivisor As Double
            srcLayerXDivisor = 1# / srcLayer.GetLayerCanvasXModifier
            srcLayerYDivisor = 1# / srcLayer.GetLayerCanvasYModifier
            
            With srcRect
                .Left = (intRectSrc.Left - srcLayer.GetLayerOffsetX) * srcLayerXDivisor
                .Width = intRectSrc.Width * srcLayerXDivisor
                .Top = (intRectSrc.Top - srcLayer.GetLayerOffsetY) * srcLayerYDivisor
                .Height = intRectSrc.Height * srcLayerYDivisor
            End With
            
            'Precalculate dimensions of any temporary DIBs we need to create.  These can have fractional offsets on either side, so we need
            ' to make sure the DIB is at least large enough to contain 0.999 offsets on either side.
            dibWidthCalc = Int(dstRect.Width + (dstRect.Left - Int(dstRect.Left)) + 0.999)
            dibHeightCalc = Int(dstRect.Height + (dstRect.Top - Int(dstRect.Top)) + 0.999)
            If (dibWidthCalc < 1) Then dibWidthCalc = 1
            If (dibHeightCalc < 1) Then dibHeightCalc = 1
            
            'With all coordinate math complete, we now proceed with the actual compositing!
            
            'While the layer provides its own hash (representing whether internal layer settings have changed), we also append some viewport-specific
            ' bits to the hash.  This way, if the layer OR the viewport changes, we regenerate the viewport-specific cache for this layer.
            testHash = srcLayer.GetViewportHash_Theoretical(srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, dibWidthCalc, dibHeightCalc)
            
            'We know affine transforms (aside from scaling) are not active.  Proceed with our highly optimized rendering chain.
            
            'Custom blend-modes and custom alpha-modes require creation of a temporary DIB, because we must perform
            ' our own compositing (as GDI/GDI+ can't do it for us).  Check for those states now.
            If nonStandardBlendingActive Then
            
                'One way or another, we will be using integer-based offsets to render the final image.  Calculate them now.
                xOffsetInt = Int(dstRect.Left - dstViewportRect.Left)
                yOffsetInt = Int(dstRect.Top - dstViewportRect.Top)
                
                If (srcLayer.GetViewportHash(levelOfDetail) = 0) Or (srcLayer.GetViewportHash(levelOfDetail) <> testHash) Or (Not cachedDIBValid) Or ignoreInternalCaches Then
                
                    'The layer and/or viewport has changed since our last composite.  Create a new DIB in the cheapest way possible
                    ' (e.g., if the current temporary DIB size is acceptable, just reset it instead of fully redrawing it)
                    If (lodDIB.GetDIBWidth < dibWidthCalc) Or (lodDIB.GetDIBHeight < dibHeightCalc) Then
                        lodDIB.CreateBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
                        lodDIB.SetInitialAlphaPremultiplicationState True
                    Else
                        lodDIB.ResetDIB
                    End If
                    
                    'Check for the special case of 100% zoom; when this happens, and the image has no non-destructive resizing applied,
                    ' we can completely skip the resize step and simply use BitBlt (way faster)
                    If isRectZoomless And (Not srcLayer.AffineTransformsActive(True)) Then
                        GDI.BitBltWrapper lodDIB.GetDIBDC, 0, 0, dibWidthCalc, dibHeightCalc, srcDIB.GetDIBDC, srcRect.Left, srcRect.Top, vbSrcCopy
                        
                    'Zoom or non-destructive resize/rotate are active.  Extra work is required.
                    Else
                        
                        'If zoomed-in, we can switch to StretchBlt for performance gains.  (As nearest-neighbor interpolation will
                        ' be used regardless of the caller's requested interpolation type.)
                        If (srcImageRect.Width < dstViewportRect.Width) Then
                            
                            'We are zoomed-in to the image, e.g. zoom is larger than 100%.
                            
                            'Paint the resized DIB to the layer, at offset (0, 0), with subpixel offsets already accounted for.
                            ' (If non-destructive resizing is active, we'll also use the layer's specified resample mode instead of
                            '  the function-wide resample mode.)
                            If (srcLayer.GetLayerCanvasXModifier <> 1#) Or (srcLayer.GetLayerCanvasYModifier <> 1#) Then tmpResampleMode = srcLayer.GetLayerResizeQuality_GDIPlus Else tmpResampleMode = GP_IM_NearestNeighbor
                            GDIPlus_StretchBlt lodDIB, (dstRect.Left - dstViewportRect.Left) - xOffsetInt, (dstRect.Top - dstViewportRect.Top) - yOffsetInt, dstRect.Width, dstRect.Height, srcDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, , tmpResampleMode, , Not isBaseLayer, True, True
                            
                        'We are zoomed-out, so StretchBlt cannot be used (as it will cause artifacting for 32bpp data).  GDI+ is our only
                        ' option at present.
                        Else
                            
                            'Paint the resized DIB to the layer, at offset (0, 0), with subpixel offsets already accounted for.
                            ' (If non-destructive resizing is active, we'll also use the layer's specified resample mode instead of
                            '  the function-wide resample mode.)
                            If (srcLayer.GetLayerCanvasXModifier <> 1#) Or (srcLayer.GetLayerCanvasYModifier <> 1#) Then tmpResampleMode = srcLayer.GetLayerResizeQuality_GDIPlus Else tmpResampleMode = interpolationType
                            GDIPlus_StretchBlt lodDIB, (dstRect.Left - dstViewportRect.Left) - xOffsetInt, (dstRect.Top - dstViewportRect.Top) - yOffsetInt, dstRect.Width, dstRect.Height, srcDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, , tmpResampleMode, , Not isBaseLayer, , True
                            
                        End If
                        
                    End If
                    
                    'Update the layer hash, so we can possibly skip this work in the future
                    srcLayer.SetViewportHash_FromLong levelOfDetail, testHash
                    
                'End If for regenerating temporary layer compositing level-of-detail DIB (lodDIB)
                End If
                
                'lodDIB now contains the chunk of this layer that appears on the viewport, with all non-destructive edits applied,
                ' and subpixel positioning already accounted for.
                
                ' Time to composite it!
                If nonStandardBlendingActive Then
                    
                    'Composite and blend in one fell swoop
                    CompositeAndBlendDIBs lodDIB, m_dstDIB, srcLayer.GetLayerBlendMode, xOffsetInt, yOffsetInt, srcLayer.GetLayerOpacity / 100, , srcLayer.GetLayerAlphaMode, , maskInViewportSpace
                    
                'If no wacky blend modes are in use, we can use GDI's AlphaBlend for a performance boost.
                ' (Note that this is only possible because subpixel offsets have already been accounted for, when creating the temporary DIB.)
                Else
                    lodDIB.AlphaBlendToDC m_dstDIB.GetDIBDC, srcLayer.GetLayerOpacity * 2.55, xOffsetInt, yOffsetInt
                End If
                
            'This is a simple layer, with no non-destructive effects or weird blend modes.  Composite it instantly, without using
            ' a temporary placeholder layer.
            Else
                
                'If the rect is zoomless, it means we don't have to perform subpixel positioning.  GDI's AlphaBlend is thus perfectly
                ' acceptable, and it will be much faster (5-10x) than GDI+.
                If isRectZoomless And (srcLayer.GetLayerCanvasXModifier = 1#) And (srcLayer.GetLayerCanvasYModifier = 1#) Then
                    
                    'AlphaBlend is very picky about invalid width/height values.  Make sure the source rect is valid.
                    ' (It may be ever-so-slightly off due to rounding issues during all the floating-point math we perform.)
                    'If srcRect.Width > srcLayer.getLayerWidth(False) Then srcRect.Width = srcLayer.getLayerWidth(False)
                    'If srcRect.Height > srcLayer.getLayerHeight(False) Then srcRect.Height = srcLayer.getLayerHeight(False)
                    'srcDIB.alphaBlendToDCEx m_dstDIB.getDIBDC, dstRect.Left - dstViewportRect.Left, dstRect.Top - dstViewportRect.Top, dstRect.Width, dstRect.Height, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, srcLayer.getLayerOpacity * 2.55
                    
                    'I have temporarily disabled the use of GDI's AlphaBlend function here, as it experiences inexplicable rendering failures
                    ' on certain image/viewport size combinations.
                    GDIPlus_StretchBlt m_dstDIB, dstRect.Left - dstViewportRect.Left, dstRect.Top - dstViewportRect.Top, dstRect.Width, dstRect.Height, srcDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, srcLayer.GetLayerOpacity * 0.01, GP_IM_NearestNeighbor
                    
                Else
                    
                    If (srcRect.Width > srcLayer.GetLayerWidth(False)) Then srcRect.Width = srcLayer.GetLayerWidth(False)
                    If (srcRect.Height > srcLayer.GetLayerHeight(False)) Then srcRect.Height = srcLayer.GetLayerHeight(False)
                    
                    'When zoomed in, nearest neighbor is preferred.  This is fast enough that we don't need to use a temporary DIB cache.
                    If (srcImageRect.Width < dstViewportRect.Width) Then
                        GDIPlus_StretchBlt m_dstDIB, dstRect.Left - dstViewportRect.Left, dstRect.Top - dstViewportRect.Top, dstRect.Width, dstRect.Height, srcDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, srcLayer.GetLayerOpacity * 0.01, GP_IM_NearestNeighbor, , , True
                    
                    'When zoomed out, use the supplied interpolation mode.  The caller typically toggles this according to PD's global
                    ' performance settings.
                    Else
                        
                        'One way or another, we will be using integer-based offsets to render the final image.  Calculate them now.
                        xOffsetInt = Int(dstRect.Left - dstViewportRect.Left)
                        yOffsetInt = Int(dstRect.Top - dstViewportRect.Top)
                        
                        'See if our existing temporary DIB hash is valid; if it is, we can proceed straight to rendering!
                        If (srcLayer.GetViewportHash(levelOfDetail) = 0) Or (srcLayer.GetViewportHash(levelOfDetail) <> testHash) Or (Not cachedDIBValid) Or ignoreInternalCaches Then
                        
                            'The layer and/or viewport has changed since our last composite.  Create a new DIB in the cheapest way possible
                            ' (e.g., if the current temporary DIB size is acceptable, just reset it instead of fully redrawing it)
                            If (lodDIB.GetDIBWidth < dibWidthCalc) Or (lodDIB.GetDIBHeight < dibHeightCalc) Then
                                lodDIB.CreateBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
                                lodDIB.SetInitialAlphaPremultiplicationState True
                            Else
                                lodDIB.ResetDIB 0
                            End If
                            
                            'Paint the resized DIB to the layer, at offset (0, 0), with subpixel offsets already accounted for.
                            ' (If non-destructive resizing is active, we'll also use the layer's specified resample mode instead of
                            '  the function-wide resample mode.)
                            If (srcLayer.GetLayerCanvasXModifier <> 1#) Or (srcLayer.GetLayerCanvasYModifier <> 1#) Then tmpResampleMode = srcLayer.GetLayerResizeQuality_GDIPlus Else tmpResampleMode = interpolationType
                            GDIPlus_StretchBlt lodDIB, (dstRect.Left - dstViewportRect.Left) - xOffsetInt, (dstRect.Top - dstViewportRect.Top) - yOffsetInt, dstRect.Width, dstRect.Height, srcDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, , tmpResampleMode, , Not isBaseLayer, , True
                            
                            'Update the layer hash, so we can possibly skip this work in the future
                            srcLayer.SetViewportHash_FromLong levelOfDetail, testHash
                            
                        End If
                        
                        'Because subpixel offsets have already been accounted for, we can use hardware-accelerated alpha blend here.
                        lodDIB.AlphaBlendToDC m_dstDIB.GetDIBDC, srcLayer.GetLayerOpacity * 2.55, xOffsetInt, yOffsetInt
                        
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
            srcLayer.GetLayerCornerCoordinates m_PlgPoints
            
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
            If (dibWidthCalc < 1) Then dibWidthCalc = 1
            If (dibHeightCalc < 1) Then dibHeightCalc = 1
            
            'While the layer provides its own hash (representing whether internal layer settings have changed), we also append some viewport-specific
            ' bits to the hash.  This way, if the layer OR the viewport changes, we regenerate the viewport-specific cache for this layer.  However,
            ' if both stay the same, we can simply re-use our previous rendering, which saves tons of time if layers other than the current one
            ' have active affine transforms (as we don't need to re-process them!).
            testHash = srcLayer.GetViewportHash_Theoretical(m_PlgPointsDst(0).x, m_PlgPointsDst(0).y, m_PlgPointsDst(1).x, m_PlgPointsDst(1).y, m_PlgPointsDst(2).x, m_PlgPointsDst(2).y, interpolationType)
            
            'This function differs from the standard pipeline because we *always* create a temporary compositing DIB.  The speed hit from
            ' a complex affine transform is significant, especially on large images, so caching the transform (even for normal blendmode)
            ' is always preferable.
            
            'Check the viewport hash now.  Certain conditions can force us to ignore internal caches (e.g. generating thumbnails).
            If (srcLayer.GetViewportHash(levelOfDetail) = 0) Or (srcLayer.GetViewportHash(levelOfDetail) <> testHash) Or (Not cachedDIBValid) Or ignoreInternalCaches Then
                
                'The layer and/or viewport has changed since our last composite.  Create a new DIB in the cheapest way possible
                ' (e.g., if the current temporary DIB size is acceptable, just reset it instead of fully redrawing it)
                If (lodDIB.GetDIBWidth < dibWidthCalc) Or (lodDIB.GetDIBHeight < dibHeightCalc) Then
                    lodDIB.CreateBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
                    lodDIB.SetInitialAlphaPremultiplicationState True
                Else
                    lodDIB.ResetDIB
                End If
                
                'Render the affine-transformed DIB to the temporary compositing DIB, at 100% opacity and normal blend mode.
                ' (Variable opacity and blend mode will be handled later.)
                '
                'Also, note that we branch according to the performance preference passed to this function.  By design, this is a different
                ' approach than we use in the normal, rectangular rendering branch.  Because affine transforms can be extremely performance-intensive,
                ' the user needs a way to override the layer's default stretching behavior.  (Note that, as usual, saving to file, copying to
                ' clipboard, etc, *always* use the layer's settings, not the viewport ones.)
                If (interpolationType = GP_IM_NearestNeighbor) Then
                    GDIPlus_PlgBlt lodDIB, m_PlgPointsDst, srcDIB, 0, 0, srcLayer.GetLayerWidth(False), srcLayer.GetLayerHeight(False), 1, GP_IM_NearestNeighbor
                Else
                    GDIPlus_PlgBlt lodDIB, m_PlgPointsDst, srcDIB, 0, 0, srcLayer.GetLayerWidth(False), srcLayer.GetLayerHeight(False), 1, srcLayer.GetLayerResizeQuality_GDIPlus
                End If
                
                'Update the layer hash, so we can possibly skip this work in the future
                srcLayer.SetViewportHash_FromLong levelOfDetail, testHash
                
            End If
            
            'lodDIB now contains the chunk of this layer that appears on the viewport, with the affine transform and any/all
            ' non-destructive edits applied to it.  The only thing left to do is composite it!
                
            'Non-standard blend modes must be handled by our internal compositor
            If nonStandardBlendingActive Then
                CompositeAndBlendDIBs lodDIB, m_dstDIB, srcLayer.GetLayerBlendMode, 0, 0, srcLayer.GetLayerOpacity / 100, , srcLayer.GetLayerAlphaMode
                
            'If no wacky blend modes are in use, we can use GDI's AlphaBlend for a performance boost
            Else
                lodDIB.AlphaBlendToDC m_dstDIB.GetDIBDC, srcLayer.GetLayerOpacity * 2.55, 0, 0
            End If
                
        End If
        
    'END block for layer intersecting the current viewport.  (Note that we simply ignore layers that don't intersect the viewport.)
    End If

End Sub

'This helper function is only used internally, by the GetCompositedRect function.  It is a specially optimized version of the
' GetCompositedRect_LayerHelper() function, with a number of changes made to improve paint performance.
'
'Paint strokes are messy because we must first composite the paint stroke against its base layer, *then* render the combined
' result onto the compositor stack.  This ensures that the paint stroke assumes any properties of the layer beneath it,
' like a non-standard blend mode or non-100% layer opacity.
'
'The standard compositor is heavily optimized against a purely linear, layer-by-layer rendering approach, so it's easier
' to handle the specialized paint case separately.
'
'IMPORTANT!  This function is only relevant when painting onto a raster layer.  When painting onto a vector layer,
' the paint strokes are automatically assigned to a new layer.
Private Sub GetCompositedRect_PaintOp(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByRef dstViewportRect As RectF, ByRef srcImageRect As RectF, ByVal interpolationType As GP_InterpolationMode, Optional ByVal isRectZoomless As Boolean = False, Optional ByVal ptrToAlternateScratch As Long = 0&)
    
    'Just like the standard GetCompositedRect_LayerHelper() function, this function requires two non-optional rects as parameters:
    
    ' - dstViewportRect: the destination rect of the fully composited rect, IN VIEWPORT COORDINATES.  Note that its size is constant if the
    '                     viewport size is constant.  (e.g. on an image larger than the viewport, zoom of 100% and 200% will have identical
    '                     values for this rect)
    ' - srcImageRect: the source rect of the image, IN IMAGE COORDINATES.  This is the rectangular region of the image represented by the
    '                  current viewport.  (Remember: if zoomed-out, it will represent a *larger region* than the viewport rect.)
    
    'dstViewportRect is required to infer the current viewport zoom, while srcImageRect is used to minimize the amount of custom
    ' rendering we do.  (Portions of the layer lying outside the viewport are ignored.)
    
    'To simplify coordinate calculation, we also generate a bunch of our own rect objects.  The objects are as follows:
    
    'Floating-point rect of the FULL SOURCE LAYER, IN IMAGE COORDINATES, *WITHOUT* AFFINE TRANSFORMS APPLIED
    Dim layerRect As RectF
    
    'Floating-point rect of the FULL SOURCE LAYER, IN IMAGE COORDINATES, WITH AFFINE TRANSFORMS APPLIED
    Dim layerRectAffine As RectF
    
    'Floating-point rect of the PORTION OF THE SOURCE LAYER RELEVANT TO THE CURRENT VIEWPORT, IN VIEWPORT COORDINATES
    Dim intRectSrc As RectF
    
    'Floating-point rect of the DESTINATION OF intRectSrc (above)
    Dim dstRect As RectF
    
    'Finally, a floating-point rect of intRectSrc, IN 0-BASED 1:1 COORDINATES.  This rect is used to crop-out a portion of the
    ' source layer for pre-processing, if non-destructive FX are active.
    Dim srcRect As RectF
    
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
    Dim tmpResampleMode As GP_InterpolationMode
    
    'Before doing anything else, we want to see if we can completely skip processing this layer.  A layer is skippable if it does not
    ' intersect the current viewport, but note that we must explicitly account for non-destructive transforms (including complex affine
    ' transforms, like rotate/skew) when calculating the intersection!
    Dim layerAndViewportIntersect As Boolean
    
    'TODO: some kind of fancy function that finds the intersection between a rotated/skewed image's boundary polygon instead of its full rect.
    '      The full rect may indicate overlap, but if the overlap is only a transparent corner region, we can still skip this layer.
    If srcLayer.AffineTransformsActive(False) Then
        Layers.FillRectForLayerF srcLayer, layerRectAffine, True, True
        layerAndViewportIntersect = GDI_Plus.IntersectRectF(intRectSrc, srcImageRect, layerRectAffine)
    
    'Layers without affine transforms use a shortcut intersection approach.
    Else
        Layers.FillRectForLayerF srcLayer, layerRect, True, False
        layerAndViewportIntersect = GDI_Plus.IntersectRectF(intRectSrc, srcImageRect, layerRect)
    End If
    
    'If the target layer intersects the viewport, we must first copy over the portion of the layer that is visible in the viewport.
    ' (If it is *not* visible, we still need to render the paint stroke, FYI!)
    If layerAndViewportIntersect Then
        
        'The passed layer intersects the viewport, so we must deal with it.
        
        'If this is *not* the first time we're rendering this paint stroke to the target layer, we will have already
        ' cached the relevant portion of the image at the current viewport zoom.  This allows us to skip the (costly)
        ' process of resizing the layer to match current zoom settings.
        cachedDIBValid = Not srcLayer.SyncInternalDIB(CLC_Painting)
        
        'Paint operations use their own specialized level-of-detail (LOD) DIB.
        Dim lodDIB As pdDIB
        Set lodDIB = srcLayer.TmpLODDIB(CLC_Painting)
        
        'Also, point a reference at the target layer's DIB
        Dim srcDIB As pdDIB
        Set srcDIB = srcLayer.GetLayerDIB
        
        'For this first stage, we are going to ignore all specialized layer settings (e.g. blend mode, alpha mode, opacity).
        ' We only want the raw image bytes, resized to match the current viewport.  We will then composite the paint stroke
        ' on top of this result, and *then* we'll composite the result onto the layer stack with all specialized settings applied.
        
        'To improve performance, this function also calculates and checks a dedicated "viewport hash" of the layer.  If the layer
        ' AND viewport have not changed since our last render, we can just grab the last temporary viewport object we generated.
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
        If (Not srcLayer.AffineTransformsActive(False)) Then
            
            'For normal, rectangular layers, the composite operation works similar to StretchBlt, where we calculate a destination rectangle
            ' and a source rectangle.  We know that these two rectangles define the same rectangle of the image, which allows us to infer
            ' zoom (without requiring access to the target canvas object).
            
            'The earlier LayerAndViewportIntersect test created a new intersection rect, which tells us where this layer overlaps the
            ' destination viewport rect.  However, the intersect rect was calculated in the layer coordinate space, and we also need
            ' to know where it lies in the viewport coordinate space.  (In many cases, the intersect rect will only cover some subsection
            ' of the viewport, meaning our temporary compositing DIB can also be smaller than the viewport, saving time and memory).
            
            'The GetDstRectFromSrcRectF() function uses our already calculated eight bare destination and source values (x/y/width/height)
            ' to calculate a canvas-space destination rect for this layer, e.g. defining the destination region of the viewport covered
            ' by this layer.
            GetDstRectFromSrcRectF dstRect, intRectSrc, dstViewportRect.Left, dstViewportRect.Top, dstViewportRect.Width, dstViewportRect.Height, srcImageRect.Left, srcImageRect.Top, srcImageRect.Width, srcImageRect.Height
            
            'dstRect and intRectSrc now contain StretchBlt-compatible destination and source rectangles RELATIVE TO THE FULL IMAGE.
            
            'Because the current layer may not be the same size as the full image, we must perform one final translation: a source
            ' rect that represents the source area, relative to the current layer's DIB.  This step is important if the layer has
            ' non-destructive resizing active, because this is when we apply that calculation.
            With srcRect
                .Left = (intRectSrc.Left - srcLayer.GetLayerOffsetX) * (1# / srcLayer.GetLayerCanvasXModifier)
                .Width = intRectSrc.Width * (1# / srcLayer.GetLayerCanvasXModifier)
                .Top = (intRectSrc.Top - srcLayer.GetLayerOffsetY) * (1# / srcLayer.GetLayerCanvasYModifier)
                .Height = intRectSrc.Height * (1# / srcLayer.GetLayerCanvasYModifier)
            End With
            
            'Precalculate dimensions of any temporary DIBs we need to create.  These can have fractional offsets on either side, so we need
            ' to make sure the DIB is at least large enough to contain 0.999 offsets on either side.
            dibWidthCalc = Int(dstRect.Width + (dstRect.Left - Int(dstRect.Left)) + 0.999)
            dibHeightCalc = Int(dstRect.Height + (dstRect.Top - Int(dstRect.Top)) + 0.999)
            
            'With all coordinate math complete, we now proceed with the actual compositing!
            
            'While the layer provides its own hash (representing whether internal layer settings have changed), we also append some viewport-specific
            ' bits to the hash.  This way, if the layer OR the viewport changes, we regenerate the viewport-specific cache for this layer.
            testHash = srcLayer.GetViewportHash_Theoretical(srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, dibWidthCalc, dibHeightCalc)
            
            'One way or another, we will be using integer-based offsets to render the final image.  Calculate them now.
            xOffsetInt = Int(dstRect.Left - dstViewportRect.Left)
            yOffsetInt = Int(dstRect.Top - dstViewportRect.Top)
                
            'If our current hash is invalid, the layer's contents have changed since our last paint operation (meaning this is likely
            ' the first mouse event of a paint stroke).  Generate a new cached image now.
            If (srcLayer.GetViewportHash(CLC_Painting) = 0) Or (srcLayer.GetViewportHash(CLC_Painting) <> testHash) Or (Not cachedDIBValid) Then
                
                'The layer and/or viewport has changed since our last composite.  Create a new DIB in the cheapest way possible
                ' (e.g., if the current temporary DIB size is acceptable, just reset it instead of fully redrawing it)
                If (lodDIB.GetDIBWidth < dibWidthCalc) Or (lodDIB.GetDIBHeight < dibHeightCalc) Then
                    lodDIB.CreateBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
                    lodDIB.SetInitialAlphaPremultiplicationState True
                Else
                    lodDIB.ResetDIB 0
                End If
                    
                'Check for the special case of 100% zoom; when this happens, and the image has no non-destructive resizing applied,
                ' we can completely skip the resize step and simply use BitBlt (way faster)
                If isRectZoomless And (Not srcLayer.AffineTransformsActive(True)) Then
                    GDI.BitBltWrapper lodDIB.GetDIBDC, 0, 0, dibWidthCalc, dibHeightCalc, srcDIB.GetDIBDC, srcRect.Left, srcRect.Top, vbSrcCopy
                    
                'Zoom or non-destructive resize/rotate are active on this layer.  Extra work is required.
                Else
                                
                    'If zoomed-in, we can switch to StretchBlt for performance gains (because nearest-neighbor interpolation will
                    ' be used regardless of the caller's requested interpolation type.)
                    If (srcImageRect.Width < dstViewportRect.Width) Then
                        
                        'We are zoomed-in to the image, e.g. zoom is larger than 100%.
                        
                        'If non-destructive resizing is active, we'll also use the layer's specified resample mode instead of
                        ' the function-wide resample mode.
                        If (srcLayer.GetLayerCanvasXModifier <> 1#) Or (srcLayer.GetLayerCanvasYModifier <> 1#) Then tmpResampleMode = srcLayer.GetLayerResizeQuality_GDIPlus Else tmpResampleMode = GP_IM_NearestNeighbor
                        
                    'We are zoomed-out, so StretchBlt cannot be used (as it will cause artifacting for 32bpp data).  GDI+ is our only
                    ' option at present.
                    Else
                        
                        'If non-destructive resizing is active, we'll also use the layer's specified resample mode instead of
                        ' the function-wide resample mode.
                        If (srcLayer.GetLayerCanvasXModifier <> 1#) Or (srcLayer.GetLayerCanvasYModifier <> 1#) Then tmpResampleMode = srcLayer.GetLayerResizeQuality_GDIPlus Else tmpResampleMode = interpolationType
                        
                    End If
                    
                    'Paint the resized DIB to the layer, at offset (0, 0), with subpixel offsets already accounted for.
                    GDIPlus_StretchBlt lodDIB, (dstRect.Left - dstViewportRect.Left) - xOffsetInt, (dstRect.Top - dstViewportRect.Top) - yOffsetInt, dstRect.Width, dstRect.Height, srcDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, , tmpResampleMode, , True, (srcImageRect.Width < dstViewportRect.Width), True
                    
                End If
                
                'Update the layer hash, so we can skip this work on our next pass
                srcLayer.SetViewportHash_FromLong CLC_Painting, testHash
                
            'End If for regenerating temporary layer compositing level-of-detail DIB (lodDIB)
            End If
                
            'lodDIB now contains the chunk of this layer that appears on the viewport, with subpixel positioning already accounted for.
            
            'Copy the cached DIB into a temporary compositing DIB (which is used so that we don't screw up our cached copy - we want
            ' that around for future rendering passes, remember!)
            
            '(Also note that this temporary DIB will always exist at a correct size, c/o our parent GetCompositedRect() function.)
            GDI.BitBltWrapper srcImage.ScratchLayer.TmpLODDIB(CLC_Painting).GetDIBDC, xOffsetInt, yOffsetInt, lodDIB.GetDIBWidth, lodDIB.GetDIBHeight, lodDIB.GetDIBDC, 0, 0, vbSrcCopy
            
            'End handling of layers with no non-scaling affine transforms
        
        'One or more non-scaling affine transforms are active (e.g. rotation).  Special coordinate calculations are required.
        Else
            
            'Start by retrieving the corner coordinates of the transformed layer.  We need these values to perform a parallelogram-style blt.
            srcLayer.GetLayerCornerCoordinates m_PlgPoints
            
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
            
            'While the layer provides its own hash (representing whether internal layer settings have changed), we also append some
            ' viewport-specific bits to the hash.  This way, if the layer OR the viewport changes, we regenerate the
            ' viewport-specific cache for this layer.  However, if both stay the same, we can simply re-use our previous rendering,
            ' which saves tons of time if layers other than the current one have active affine transforms (as we don't need to re-process them!).
            testHash = srcLayer.GetViewportHash_Theoretical(m_PlgPointsDst(0).x, m_PlgPointsDst(0).y, m_PlgPointsDst(1).x, m_PlgPointsDst(1).y, m_PlgPointsDst(2).x, m_PlgPointsDst(2).y, interpolationType)
            
            'Check the viewport hash now.  Certain conditions can force us to ignore internal caches (e.g. generating thumbnails).
            If (srcLayer.GetViewportHash(CLC_Painting) = 0) Or (srcLayer.GetViewportHash(CLC_Painting) <> testHash) Or (Not cachedDIBValid) Then
                
                'The layer and/or viewport has changed since our last composite.  Create a new DIB in the cheapest way possible
                ' (e.g., if the current temporary DIB size is acceptable, just reset it instead of fully redrawing it)
                If (lodDIB.GetDIBWidth < dibWidthCalc) Or (lodDIB.GetDIBHeight < dibHeightCalc) Then
                    lodDIB.CreateBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
                    lodDIB.SetInitialAlphaPremultiplicationState True
                Else
                    lodDIB.ResetDIB
                End If
                
                'Render the affine-transformed DIB to the temporary compositing DIB, at 100% opacity and normal blend mode.
                ' (Variable opacity and blend mode will be handled later.)
                '
                'Also, note that we branch according to the performance preference passed to this function.  By design, this is a different
                ' approach than we use in the normal, rectangular rendering branch.  Because affine transforms can be extremely
                ' performance-intensive, the user needs a way to override the layer's default stretching behavior.  (Note that, as usual,
                ' saving to file, copying to clipboard, etc, *always* use the layer's settings, not the viewport ones.)
                If (interpolationType = GP_IM_NearestNeighbor) Then
                    GDIPlus_PlgBlt lodDIB, m_PlgPointsDst, srcDIB, 0, 0, srcLayer.GetLayerWidth(False), srcLayer.GetLayerHeight(False), 1, GP_IM_NearestNeighbor
                Else
                    GDIPlus_PlgBlt lodDIB, m_PlgPointsDst, srcDIB, 0, 0, srcLayer.GetLayerWidth(False), srcLayer.GetLayerHeight(False), 1, srcLayer.GetLayerResizeQuality_GDIPlus
                End If
                
                'Update the layer hash, so we can possibly skip this work in the future
                srcLayer.SetViewportHash_FromLong CLC_Painting, testHash
                
            End If
                 
            'lodDIB now contains the chunk of this layer that appears on the viewport, with the affine transform and any/all
            ' subpixel positioning already accounted for.
            
            'Copy the cached DIB into a temporary compositing DIB (which is used so that we don't screw up our cached copy - we want
            ' that around for future rendering passes, remember!)
            
            '(Also note that this temporary DIB will always exist at a correct size, c/o our parent GetCompositedRect() function.)
            GDI.BitBltWrapper srcImage.ScratchLayer.TmpLODDIB(CLC_Painting).GetDIBDC, 0, 0, lodDIB.GetDIBWidth, lodDIB.GetDIBHeight, lodDIB.GetDIBDC, 0, 0, vbSrcCopy
            
            'End handling of layers with active affine transforms
                
        End If
        
    'END block for layer intersecting the current viewport.  (Note that we simply ignore layers that don't intersect the viewport.)
    End If
    
    'srcImage.ScratchLayer.TmpLODDIB(CLC_Painting) now contains a copy of the layer onto which we are painting.  The layer has had
    ' any/all transforms applied, but it has not had *any* blend mode, non-destructive FX, opacity, or alpha mode calculations
    ' applied, by design.
    
    'Next, we want to merge the paint stroke image onto it.  The paint stroke may have its own combination of blend and/or
    ' alpha modes, and we must handle those just like any other layer.
    
    'Start by calculating the intersection rect between the paintbrush image (which is always the full size of the base image),
    ' and the current viewport
    Layers.FillRectForLayerF srcImage.ScratchLayer, layerRect, False, False
    layerAndViewportIntersect = GDI_Plus.IntersectRectF(intRectSrc, srcImageRect, layerRect)
    
    Dim nonStandardBlendingActive As Boolean
    
    'We generally assume that the current paintbrush overlaps the viewport, but for the sake of good coding, let's check
    ' just in case.
    If layerAndViewportIntersect Then
    
        'The paintbrush layer intersects the viewport.  No surprise there!
        
        'Normally, we would next test the DIB cache for this LOD setting, and re-use a previously cached image if available.
        ' Paint strokes work a bit differently, however, as we always know that *some* portion of the cache must be regenerated.
        ' If this is the first event in the current stroke, we need to generate the entire cache for the paint layer.  If this
        ' is *not* the first event, however, we can simply repaint the portion of the cache that's changed since the last event.
        Set srcDIB = srcImage.ScratchLayer.GetLayerDIB
        
        'Because non-standard blend and/or alpha modes require us to use our internal compositor, we may need to use a temporary
        ' LOD DIB to hold our resized results.  Cache that reference now
        Set lodDIB = srcImage.ScratchLayer.TmpLODDIB(CLC_Viewport)
        
        'Check for non-standard blend or alpha modes for this paint stroke.  As expected, they require a custom processing pipeline.
        nonStandardBlendingActive = (srcImage.ScratchLayer.GetLayerBlendMode <> BM_Normal) Or (srcImage.ScratchLayer.GetLayerAlphaMode <> AM_Normal)
        
        'We also need to check if the destination layer (*not* the paint layer) has locked alpha values.  If it does, we will
        ' pre-process the paint stroke alpha to match.
        If (Not nonStandardBlendingActive) Then nonStandardBlendingActive = (srcLayer.GetLayerAlphaMode = AM_Locked)
        
        'Finally, if a selection mask is active, we need to pre-process the paint stroke against the selection mask,
        ' and remove any stroke details that fall outside the selection boundary.
        Dim srcMask As pdDIB
        If srcImage.IsSelectionActive Then
            nonStandardBlendingActive = True
            Set srcMask = srcImage.MainSelection.GetCompositeMaskDIB_ViewportCopy()
        End If
        
        'Just like the base layer, we now need to perform a composite operation similar to StretchBlt, where we calculate a
        ' destination rectangle and a source rectangle.  We know that these two rectangles define the same rectangle of the
        ' image, which allows us to infer zoom (without requiring access to the target canvas object).
        
        'The earlier LayerAndViewportIntersect test created a new intersection rect, which tells us where this layer overlaps the
        ' destination viewport rect.  However, the intersect rect was calculated in the layer coordinate space, and we also need
        ' to know where it lies in the viewport coordinate space.  (In many cases, the intersect rect will only cover some subsection
        ' of the viewport, meaning our temporary compositing DIB can also be smaller than the viewport, saving time and memory).
        
        'The GetDstRectFromSrcRectF() function uses our already calculated eight bare destination and source values (x/y/width/height)
        ' to calculate a canvas-space destination rect for this layer, e.g. defining the destination region of the viewport covered by this layer.
        
        'Note that we calculate this differently depending on whether this is the first paint event in a stroke, or a subsequent one
        Dim isFirstPaintEvent As Boolean
        If (g_CurrentTool = PAINT_PENCIL) Then
            isFirstPaintEvent = Tools_Pencil.IsFirstDab()
        ElseIf (g_CurrentTool = PAINT_SOFTBRUSH) Or (g_CurrentTool = PAINT_ERASER) Then
            isFirstPaintEvent = Tools_Paint.IsFirstDab()
        ElseIf (g_CurrentTool = PAINT_CLONE) Then
            isFirstPaintEvent = Tools_Clone.IsFirstDab()
        End If
        
        'Note: by forcing isFirstPaintEvent to TRUE, you can forcibly regenerate the full cached image on each stroke.
        'isFirstPaintEvent = True
        
        'If this is...
        ' 1) Our first paint event for this stroke, or...
        ' 2) A 100% zoom viewport, or...
        ' 3) A zoomed-in viewport...
        ' Rebuild the entire cache.  (This is really only strictly necessary for point (1), but the compositor needs to be updated below
        '  to allow (2) and (3) to use partial viewport composites.)
        Dim fullCacheCreationRequired As Boolean
        fullCacheCreationRequired = isFirstPaintEvent Or isRectZoomless Or (srcImageRect.Width < dstViewportRect.Width)
        
        GetDstRectFromSrcRectF dstRect, intRectSrc, dstViewportRect.Left, dstViewportRect.Top, dstViewportRect.Width, dstViewportRect.Height, srcImageRect.Left, srcImageRect.Top, srcImageRect.Width, srcImageRect.Height
        
        'If a full cache creation is *not* required, we're only going to update the paintbrush's compositor rect in the region
        ' that has been modified since the last viewport render.  This can save us a lot of time, especially if the source brush
        ' is large or if a bunch of weird blend/alpha modes are active.
        Dim intRectSrcStroke As RectF, intRectDstStroke As RectF
        If (Not fullCacheCreationRequired) Then
            If (g_CurrentTool = PAINT_PENCIL) Then
                GDI_Plus.IntersectRectF intRectSrcStroke, srcImageRect, Tools_Pencil.GetModifiedUpdateRectF()
            ElseIf (g_CurrentTool = PAINT_SOFTBRUSH) Or (g_CurrentTool = PAINT_ERASER) Then
                GDI_Plus.IntersectRectF intRectSrcStroke, srcImageRect, Tools_Paint.GetModifiedUpdateRectF()
            ElseIf (g_CurrentTool = PAINT_CLONE) Then
                GDI_Plus.IntersectRectF intRectSrcStroke, srcImageRect, Tools_Clone.GetModifiedUpdateRectF()
            End If
            GetDstRectFromSrcRectF intRectDstStroke, intRectSrcStroke, dstViewportRect.Left, dstViewportRect.Top, dstViewportRect.Width, dstViewportRect.Height, srcImageRect.Left, srcImageRect.Top, srcImageRect.Width, srcImageRect.Height
        End If
        
        'dstRect and intRectSrc now contain StretchBlt-compatible destination and source rectangles RELATIVE TO THE FULL IMAGE.
        
        'Next, we would normally apply any layer modifications (offset, x/y stretching, etc), but the paintbrush layer is not allowed
        ' to use these.  Copy over the intersection rect as-is.
        If fullCacheCreationRequired Then
            srcRect = intRectSrc
        Else
            srcRect = intRectSrcStroke
        End If
            
        'Precalculate dimensions of any temporary DIBs we need to create.  These can have fractional offsets on either side, so we need
        ' to make sure the DIB is at least large enough to contain 0.999 offsets on either side.
        dibWidthCalc = Int(dstRect.Width + (dstRect.Left - Int(dstRect.Left)) + 0.999)
        dibHeightCalc = Int(dstRect.Height + (dstRect.Top - Int(dstRect.Top)) + 0.999)
        
        'With all coordinate math complete, we now proceed with compositing the paintbrush layer!
        
        'One way or another, we will be using integer-based offsets to render the final image.  Calculate them now.
        xOffsetInt = Int(dstRect.Left - dstViewportRect.Left)
        yOffsetInt = Int(dstRect.Top - dstViewportRect.Top)
        
        'Create a new DIB in the cheapest way possible (e.g., if the current temporary DIB size is acceptable,
        ' just reset it instead of fully redrawing it)
        If (lodDIB.GetDIBWidth <> dibWidthCalc) Or (lodDIB.GetDIBHeight <> dibHeightCalc) Then
            
            'We should only ever need to recreate the DIB on the first paint stroke, but just in case, reset the event tracker to match
            isFirstPaintEvent = True
            
            lodDIB.CreateBlank dibWidthCalc, dibHeightCalc, 32, 0, 0
            lodDIB.SetInitialAlphaPremultiplicationState True
            
        Else
            'If this is *not* the first paint event, we don't need to reset anything, as we'll only modify the changed region
            If isFirstPaintEvent Then lodDIB.ResetDIB 0
        End If
        
        'Some paint tools are capable of doing their own "generate preview in viewport coordinate space" work.
        ' When they do this, they'll pass in their own viewport-ready rect, which we can simply plug into the
        ' rendering chain, without further modification.  (At present, this is primarily used by the on-canvas
        ' Gradient tools, as they can render very accurate viewport-space previews.)
        If (ptrToAlternateScratch <> 0) Then
            
            'Make sure ref count gets updated
            Dim alternateScratch As pdDIB
            ObjSetAddRef alternateScratch, ptrToAlternateScratch
            
            'Fast copy the tool's render into our local buffer
            GDI.BitBltWrapper lodDIB.GetDIBDC, 0, 0, lodDIB.GetDIBWidth, lodDIB.GetDIBHeight, alternateScratch.GetDIBDC, 0, 0, vbSrcCopy
            
        'Most of the time, we'll be responsible for generating a viewport-space preview of the full scratch layer.
        Else
        
            'Check for the special case of 100% zoom; when this happens, we can completely skip the resize step and
            ' simply use BitBlt (way faster).
            If isRectZoomless Then
                GDI.BitBltWrapper lodDIB.GetDIBDC, 0, 0, dibWidthCalc, dibHeightCalc, srcDIB.GetDIBDC, srcRect.Left, srcRect.Top, vbSrcCopy
                
            'The viewport is zoomed, so extra work is required.
            Else
                
                'If zoomed-in, we can switch to StretchBlt for performance gains.  (Nearest-neighbor interpolation will
                ' be used regardless of the caller's requested interpolation type.)
                If (srcImageRect.Width < dstViewportRect.Width) Then
                    
                    'We are zoomed-in to the image, e.g. zoom is larger than 100%.
                    
                    'Paint the resized DIB to the temporary DIB, at offset (0, 0), with subpixel offsets already accounted for.
                    ' (If non-destructive resizing is active, we'll also use the layer's specified resample mode instead of
                    '  the function-wide resample mode.)
                    GDIPlus_StretchBlt lodDIB, (dstRect.Left - dstViewportRect.Left) - xOffsetInt, (dstRect.Top - dstViewportRect.Top) - yOffsetInt, dstRect.Width, dstRect.Height, srcDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, 1#, GP_IM_NearestNeighbor, , True, True, True
                    
                'We are zoomed-out, so StretchBlt cannot be used (as it will cause artifacting for 32bpp data).  GDI+ is our only
                ' option at present.
                Else
                    
                    'Paint the resized DIB to the layer, at offset (0, 0), with subpixel offsets already accounted for.
                    
                    ' NOTE!  I have no idea why I'm forced to specify "isZoomedIn" as TRUE here (because the image is *definitely* not zoomed-in,
                    ' it's zoomed-out), but if I don't set that to TRUE, previewed paint strokes are very obviously offset by 0.5 pixels,
                    ' as noticed when releasing the mouse and the final layer merge operation occurs (causing the on-screen preview to shift).
                    ' This requires further study, but it will probably wait until post-7.0 when some other compositor projects are planned.
                    ' TODO 8.2: deal with comment above
                    Dim intMode As GP_InterpolationMode
                    intMode = Tools_Paint.GetBrushPreviewQuality_GDIPlus()
                    
                    If fullCacheCreationRequired Then
                        GDIPlus_StretchBlt lodDIB, (dstRect.Left - dstViewportRect.Left) - xOffsetInt, (dstRect.Top - dstViewportRect.Top) - yOffsetInt, dstRect.Width, dstRect.Height, srcDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, , intMode, , True, True, True
                    Else
                        GDIPlus_StretchBlt lodDIB, (intRectDstStroke.Left - dstViewportRect.Left) - xOffsetInt, (intRectDstStroke.Top - dstViewportRect.Top) - yOffsetInt, intRectDstStroke.Width, intRectDstStroke.Height, srcDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, , intMode, , True, True, True
                    End If
                    
                End If
                
            End If
        
        '/End ptrToAlternateScratch <> 0
        End If
        
        'lodDIB now contains the chunk of the paintbrush layer that appears on the viewport.  Our next step is to composite it onto our
        ' temporary copy of the target layer DIB!
        If nonStandardBlendingActive Then
            
            'If *either* the paintbrush or the target layer has requested locked alpha, honor that now
            Dim targetAlphaMode As PD_AlphaMode
            If (srcImage.ScratchLayer.GetLayerAlphaMode = AM_Locked) Or (srcLayer.GetLayerAlphaMode = AM_Locked) Then
                targetAlphaMode = AM_Locked
            Else
                targetAlphaMode = AM_Normal
            End If
            
            'Composite and blend in one fell swoop
            CompositeAndBlendDIBs lodDIB, srcImage.ScratchLayer.TmpLODDIB(CLC_Painting), srcImage.ScratchLayer.GetLayerBlendMode, xOffsetInt, yOffsetInt, srcImage.ScratchLayer.GetLayerOpacity * 0.01, targetAlphaMode, srcImage.ScratchLayer.GetLayerAlphaMode, , srcMask
        
        'If no wacky blend modes are in use, we can use GDI's AlphaBlend for a performance boost.
        ' (Note that this is only possible because subpixel offsets have already been accounted for, when creating the temporary DIB.)
        Else
            lodDIB.AlphaBlendToDC srcImage.ScratchLayer.TmpLODDIB(CLC_Painting).GetDIBDC, srcImage.ScratchLayer.GetLayerOpacity * 2.55, xOffsetInt, yOffsetInt
        End If
        
    End If
    
    'At this point, srcImage.ScratchLayer.TmpLODDIB(CLC_Painting) now contains the fully composited version of both the
    ' base layer AND the current paint stroke.
    Set lodDIB = srcImage.ScratchLayer.TmpLODDIB(CLC_Painting)
    
    'Merge the final, fully blended result onto the rest of the compositor stack
    nonStandardBlendingActive = (srcLayer.GetLayerBlendMode <> BM_Normal) Or (srcLayer.GetLayerAlphaMode <> AM_Normal)
    
    If nonStandardBlendingActive Then
        CompositeAndBlendDIBs lodDIB, m_dstDIB, srcLayer.GetLayerBlendMode, 0, 0, srcLayer.GetLayerOpacity * 0.01, , srcLayer.GetLayerAlphaMode
        
    'If no wacky blend modes are in use, we can use GDI's AlphaBlend for a performance boost
    Else
        lodDIB.AlphaBlendToDC m_dstDIB.GetDIBDC, srcLayer.GetLayerOpacity * 2.55, 0, 0
    End If
    
End Sub

'Given StretchBlt-like parameters, create a destination rect that reflects the transformation of a source rect
' from the source coordinate space to the destination coordinate space.
Private Sub GetDstRectFromSrcRectF(ByRef dstRect As RectF, ByRef srcRect As RectF, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single)
    
    If (srcWidth <> 0#) And (srcHeight <> 0#) Then
        
        Dim wRatio As Single, hRatio As Single
        wRatio = dstWidth / srcWidth
        hRatio = dstHeight / srcHeight
        
        With dstRect
            .Left = dstX + (srcRect.Left - srcX) * wRatio
            .Width = srcRect.Width * wRatio
            .Top = dstY + (srcRect.Top - srcY) * hRatio
            .Height = srcRect.Height * hRatio
        End With
        
    End If

End Sub

'In order to render complex images quickly, this class must generate a lot of temporary objects.  If PD is under memory pressure, call this function
' to forcibly release as many temporary objects as possible.  (It should be obvious, but subsequent render calls will result in temporary objects
' being recreated, as necessary.)
Friend Sub AttemptToFreeMemory()
    
    'Start by clearing our own internal caches
    m_dstDIB.EraseDIB
    m_tmpDIB.EraseDIB
    
    'Ask the pixel blender to release memory too, if it can
    m_Blender.AttemptToFreeMemory
    
End Sub

Private Sub Class_Initialize()
    Set m_dstDIB = New pdDIB
    Set m_tmpDIB = New pdDIB
    Set m_Blender = New pdPixelBlender
    
    ReDim m_PlgPoints(0 To 3) As PointFloat
    ReDim m_PlgPointsDst(0 To 3) As PointFloat
End Sub

Private Sub Class_Terminate()
    Set m_dstDIB = Nothing
    Set m_tmpDIB = Nothing
    Set m_Blender = Nothing
End Sub
