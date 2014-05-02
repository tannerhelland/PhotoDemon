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
'Copyright �2013-2014 by Tanner Helland
'Created: 01/May/14
'Last updated: 01/May/14
'Last update: add blend modes!
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Floating-point rect
Private Type RECTF
    Left        As Single
    Top         As Single
    Right       As Single
    Bottom      As Single
End Type
    
'Some API functions are used to simplify the process of things like merging layers
Private Declare Function UnionRect Lib "user32" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long
Private Declare Function IntersectRect Lib "user32" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long

'Composite two DIBs using the requested blend mode.  An offset can be specified for the top DIB, relative to the bottom DIB.
' Note that the top DIB is treated as the destination, because after all compositing is done, the top DIB still needs
' to be alpha-blended with the bottom DIB (to account for net opacity).
'
'TODO: profile an integer-based implementation.
Friend Sub compositeDIBs(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal blendMode As LAYER_BLENDMODE, ByVal dstX As Single, ByVal dstY As Single)

    'We will be doing a lot of int/float conversions.  To speed things up, use a persistent look-up table.
    Dim IntToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        IntToFloat(i) = i / 255
    Next i
    
    'Calculate bounds for the loop.  These are relative to the top layer, which may lay outside the bounds
    ' of the bottom layer.
    Dim initX As Long, initY As Long
    Dim finalX As Long, finalY As Long
    
    If dstX < 0 Then initX = Abs(dstX)
    If dstY < 0 Then initY = Abs(dstY)
    
    finalX = topDIB.getDIBWidth - 1
    If dstX + finalX > bottomDIB.getDIBWidth - 1 Then finalX = (bottomDIB.getDIBWidth - dstX) - 1
    
    finalY = topDIB.getDIBHeight - 1
    If dstY + finalY > bottomDIB.getDIBHeight - 1 Then finalY = (bottomDIB.getDIBHeight - dstY) - 1
    
    'Make sure our bounds are valid
    If initX > finalX Then Exit Sub
    If initY > finalY Then Exit Sub
    
    'Next, we need to point VB arrays at both the top and bottom DIBs.  Note that the top DIB is treated as
    ' the "destination", because it will receive the full composite results.
    Dim bImageData() As Byte
    Dim bottomSA As SAFEARRAY2D
    prepSafeArray bottomSA, bottomDIB
    CopyMemory ByVal VarPtrArray(bImageData()), VarPtr(bottomSA), 4

    Dim tImageData() As Byte
    Dim topSA As SAFEARRAY2D
    prepSafeArray topSA, topDIB
    CopyMemory ByVal VarPtrArray(tImageData()), VarPtr(topSA), 4
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, QuickX As Long, quickBottomX As Long, quickBottomY As Long
    
    Dim topR As Single, topG As Single, topB As Single, topA As Single
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single
    Dim newR As Single, newG As Single, newB As Single, newA As Single
    
    For x = initX To finalX
        QuickX = x * 4
        quickBottomX = (dstX + x) * 4
    For y = initY To finalY
    
        quickBottomY = dstY + y
        
        'Retrieve top layer RGBA values
        topA = IntToFloat(tImageData(QuickX + 3, y))
        
        'Ignore transparent pixels
        If topA > 0 Then
        
            topR = IntToFloat(tImageData(QuickX + 2, y))
            topG = IntToFloat(tImageData(QuickX + 1, y))
            topB = IntToFloat(tImageData(QuickX, y))
            
            'Retrieve bottom layer RGB values
            bottomA = IntToFloat(bImageData(quickBottomX + 3, quickBottomY))
            
            'Again, ignore transparent pixels
            If bottomA > 0 Then
            
                bottomR = IntToFloat(bImageData(quickBottomX + 2, quickBottomY))
                bottomG = IntToFloat(bImageData(quickBottomX + 1, quickBottomY))
                bottomB = IntToFloat(bImageData(quickBottomX, quickBottomY))
                
                'Un-premultiply all RGB values
                topR = topR / topA
                topG = topG / topA
                topB = topB / topA
                
                bottomR = bottomR / bottomA
                bottomG = bottomG / bottomA
                bottomB = bottomB / bottomA
                
                'Apply the requested blend mode
                Select Case blendMode
                
                    Case BL_DARKEN
                        If bottomR < topR Then newR = bottomR Else newR = topR
                        If bottomG < topG Then newG = bottomG Else newG = topG
                        If bottomB < topB Then newB = bottomB Else newB = topB
                
                    Case BL_MULTIPLY
                        newR = topR * bottomR
                        newG = topG * bottomG
                        newB = topB * bottomB
                        
                    Case BL_COLORBURN
                        If topR = 0 Then newR = 0 Else newR = 1 - (1 - bottomR) / topR
                        If topG = 0 Then newG = 0 Else newG = 1 - (1 - bottomG) / topG
                        If topB = 0 Then newB = 0 Else newB = 1 - (1 - bottomB) / topB
                    
                    Case BL_LINEARBURN
                        newR = topR + bottomR - 1
                        newG = topG + bottomG - 1
                        newB = topB + bottomB - 1
                        
                    Case BL_LIGHTEN
                        If bottomR > topR Then newR = bottomR Else newR = topR
                        If bottomG > topG Then newG = bottomG Else newG = topG
                        If bottomB > topB Then newB = bottomB Else newB = topB
                
                    Case BL_SCREEN
                        newR = 1 - (1 - bottomR) * (1 - topR)
                        newG = 1 - (1 - bottomG) * (1 - topG)
                        newB = 1 - (1 - bottomB) * (1 - topB)
                    
                    Case BL_COLORDODGE
                        If topR = 1 Then newR = 1 Else newR = bottomR / (1 - topR)
                        If topG = 1 Then newG = 1 Else newG = bottomG / (1 - topG)
                        If topB = 1 Then newB = 1 Else newB = bottomB / (1 - topB)
                        
                    Case BL_LINEARDODGE
                        newR = topR + bottomR
                        newG = topG + bottomG
                        newB = topB + bottomB
                    
                    Case BL_OVERLAY
                        If bottomR < 0.5 Then newR = 2 * bottomR * topR Else newR = 1 - 2 * (1 - bottomR) * (1 - topR)
                        If bottomG < 0.5 Then newG = 2 * bottomG * topG Else newG = 1 - 2 * (1 - bottomG) * (1 - topG)
                        If bottomB < 0.5 Then newB = 2 * bottomB * topB Else newB = 1 - 2 * (1 - bottomB) * (1 - topB)
                    
                    Case BL_SOFTLIGHT
                        newR = (1 - 2 * topR) * bottomR * bottomR + 2 * bottomR * topR
                        newG = (1 - 2 * topG) * bottomG * bottomG + 2 * bottomG * topG
                        newB = (1 - 2 * topB) * bottomB * bottomB + 2 * bottomB * topB
                    
                    Case BL_HARDLIGHT
                        If topR < 0.5 Then newR = 2 * bottomR * topR Else newR = 1 - 2 * (1 - bottomR) * (1 - topR)
                        If topG < 0.5 Then newG = 2 * bottomG * topG Else newG = 1 - 2 * (1 - bottomG) * (1 - topG)
                        If topB < 0.5 Then newB = 2 * bottomB * topB Else newB = 1 - 2 * (1 - bottomB) * (1 - topB)
                    
                    Case BL_VIVIDLIGHT
                        If topR > 0.5 Then newR = bottomR + 2 * (topR - 0.5) Else newR = bottomR + 2 * topR - 1
                        If topB > 0.5 Then newG = bottomG + 2 * (topG - 0.5) Else newG = bottomG + 2 * topG - 1
                        If topG > 0.5 Then newB = bottomB + 2 * (topB - 0.5) Else newB = bottomB + 2 * topB - 1
                    
                    Case BL_LINEARLIGHT
                        newR = bottomR + (2 * topR) - 1
                        newG = bottomG + (2 * topG) - 1
                        newB = bottomB + (2 * topB) - 1
                        
                    Case BL_PINLIGHT
                        If topR > 0.5 Then newR = Max(bottomR, 2 * (topR - 0.5)) Else newR = Min(bottomR, 2 * topR)
                        If topG > 0.5 Then newG = Max(bottomG, 2 * (topG - 0.5)) Else newG = Min(bottomG, 2 * topG)
                        If topB > 0.5 Then newB = Max(bottomB, 2 * (topB - 0.5)) Else newB = Min(bottomB, 2 * topB)
                        
                    Case BL_HARDMIX
                        If topR < 1 - bottomR Then newR = 0 Else newR = 1
                        If topG < 1 - bottomG Then newG = 0 Else newG = 1
                        If topB < 1 - bottomB Then newB = 0 Else newB = 1
                    
                    Case BL_DIFFERENCE
                        newR = Abs(bottomR - topR)
                        newG = Abs(bottomG - topG)
                        newB = Abs(bottomB - topB)
                        
                    Case BL_EXCLUSION
                        newR = bottomR + topR - 2 * (bottomR * topR)
                        newG = bottomG + topG - 2 * (bottomG * topG)
                        newB = bottomB + topB - 2 * (bottomB * topB)
                        
                    Case BL_SUBTRACT
                        newR = bottomR - topR
                        newG = bottomG - topG
                        newB = bottomB - topB
                    
                    Case BL_DIVIDE
                        If topR = 0 Then newR = 0 Else newR = bottomR / topR
                        If topG = 0 Then newG = 0 Else newG = bottomG / topG
                        If topB = 0 Then newB = 0 Else newB = bottomB / topB
                    
                    Case Else
                        newR = topR
                        newG = topG
                        newB = topB
                
                End Select
                
                If newR < 0 Then newR = 0
                If newG < 0 Then newG = 0
                If newB < 0 Then newB = 0
                        
                If newR > 1 Then newR = 1
                If newG > 1 Then newG = 1
                If newB > 1 Then newB = 1
                
                'Alpha blend the original color with the new color, per the lower image's alpha.  (This way,
                ' if the bottom image is transparent, the top image is unchanged.)
                newR = (newR * bottomA) + (topR * (1 - bottomA))
                newG = (newG * bottomA) + (topG * (1 - bottomA))
                newB = (newB * bottomA) + (topB * (1 - bottomA))
                
                'Copy the new values into the bottom layer, and premultiply them
                tImageData(QuickX + 2, y) = (newR * 255) * topA
                tImageData(QuickX + 1, y) = (newG * 255) * topA
                tImageData(QuickX, y) = (newB * 255) * topA
                
            End If
            
        End If
    
    Next y
    Next x
    
    'Clear all array references
    CopyMemory ByVal VarPtrArray(bImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(tImageData), 0&, 4
    
End Sub

'Merge two layers together.  Note this can be used to merge any two arbitrary layers, with the bottom layer holding the result
' of the merge.  It is up to the caller to deal with any subsequent layer deletions, etc - this sub just performs the merge.
'
'The optional parameter, "bottomLayerIsFullSize", should be set to TRUE if the bottom layer is the size of the image.  This saves
' us some processing time, because we don't have to check for rect intersection.
Public Sub mergeLayers(ByRef topLayer As pdLayer, ByRef bottomLayer As pdLayer, Optional ByVal bottomLayerIsFullSize As Boolean = True)

    Dim xOffset As Double, yOffset As Double
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'If the bottom layer is the size of the image itself, we can use the existing layer offsets in our calculation.  (Nice!)
    If bottomLayerIsFullSize Then
    
        xOffset = topLayer.getLayerOffsetX
        yOffset = topLayer.getLayerOffsetY
    
    'The top and bottom layer sizes are totally independent.  This makes our life somewhat unpleasant.
    Else
    
        'First, we need to find the union rect of the two layers.  This is the smallest rectangle that holds both layers.
        Dim topRect As RECT, bottomRect As RECT, finalRect As RECT
        fillRectForLayer bottomLayer, bottomRect
        fillRectForLayer topLayer, topRect
        UnionRect finalRect, topRect, bottomRect
        
        'finalRect now contains the coordinates of the resulting rect.  Create a blank DIB at those dimensions.
        tmpDIB.createBlank finalRect.Right - finalRect.Left, finalRect.Bottom - finalRect.Top, 32, 0
        
        'We now need to do a couple of things.  Let's start by copying the bottom DIB into this new temporary DIB.
        xOffset = bottomRect.Left - finalRect.Left
        yOffset = bottomRect.Top - finalRect.Top
        BitBlt tmpDIB.getDIBDC, xOffset, yOffset, bottomLayer.layerDIB.getDIBWidth, bottomLayer.layerDIB.getDIBHeight, bottomLayer.layerDIB.getDIBDC, 0, 0, vbSrcCopy
        
        'We now need to calculate a new layer offset for this temporary DIB, which will eventually be copied into the bottom layer.
        ' Without this, the main composite won't know where to stick the layer!
        bottomLayer.setLayerOffsetX finalRect.Left
        bottomLayer.setLayerOffsetY finalRect.Top
        
        'Copy the temporary DIB into the bottom layer
        bottomLayer.layerDIB.createFromExistingDIB tmpDIB
        
        'Calculate new offsets for the top layer, then carry on with business as usual!
        xOffset = topRect.Left - finalRect.Left
        yOffset = topRect.Top - finalRect.Top
        
    End If
    
    'Now that the two layers are ready to be condensed into one, the merge process is actually very simple.
    
    'For convenience reasons, separate layers into 24bpp and 32bpp categories, and handle each specially
            
    '24bpp case
    If topLayer.layerDIB.getDIBColorDepth = 24 Then
    
        'Possible future project - provide specialized, performance-friendly handling for 24bpp layers??
        With topLayer
            .layerDIB.alphaBlendToDC bottomLayer.layerDIB.getDIBDC, 255, xOffset, yOffset
        End With
        
    '32bpp case
    Else
    
        'Make a copy of the current layer
        tmpDIB.createFromExistingDIB topLayer.layerDIB
    
        'tmpDIB now contains the chunk of this layer that appears on the viewport.  Time to composite it!
        If topLayer.getLayerBlendMode <> BL_NORMAL Then
            Debug.Print xOffset, yOffset
            compositeDIBs tmpDIB, bottomLayer.layerDIB, topLayer.getLayerBlendMode, xOffset, yOffset
        End If
    
        'Always use the AlphaBlend API for 32bpp layers
        With topLayer
            tmpDIB.alphaBlendToDC bottomLayer.layerDIB.getDIBDC, .getLayerOpacity * 2.55, xOffset, yOffset
        End With
    
    End If
    
    'The two layers have been merged successfully!  Any further actions (like deleting the top layer) must be handled
    ' by the caller.
    
End Sub

'Returns all layers of the image as a single, composited image (in pdDIB format, of course).  Because of the way VB handles
' object references, we ask the calling function to supply the DIB they want filled.  Optionally, they can also request a
' particular premultiplication status of the composited DIB's alpha values.  (This is helpful for save functions, which
' require non-premultiplied alpha, vs viewport functions, which require premultiplied alpha).
Friend Sub getCompositedImage(ByRef srcImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal premultiplicationStatus As Boolean = True)
    
    'TODO: perform special checks for 24bpp single-layer images.  For these, we can simply copy the current layer out,
    '       saving a lot of time!
    
    'Start by resizing the DIB to be the size of the current image, and placing it inside a temporary pdLayer container.
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    
    dstDIB.createBlank srcImage.Width, srcImage.Height, 32
    Set tmpLayer.layerDIB = dstDIB
    
    'We will be accessing all layers of the source image in turn.  Rather than re-requesting those layers through
    ' functions, we will simply use a local reference.
    Dim tmpLayerRef As pdLayer
    
    'Always start by rendering the bottom layer onto the (presently empty) destination DIB.  The bottom layer's blend mode
    ' does not matter, so this is a straightforward operation.
    Set tmpLayerRef = srcImage.getLayerByIndex(0)
    If tmpLayerRef.getLayerVisibility Then
    
        With tmpLayerRef
            .layerDIB.alphaBlendToDC dstDIB.getDIBDC, .getLayerOpacity * 2.55, .getLayerOffsetX, .getLayerOffsetY
        End With
    
    End If
    
    'If the image has additional layers, proceed to merge the rest of them, starting from the bottom and working our way up.
    ' Note that if a layer is invisible, we simply skip it - this is the most performance-friendly way to handle them.
    If srcImage.getNumOfLayers > 1 Then
    
        Dim i As Long
        For i = 1 To srcImage.getNumOfLayers - 1
            Set tmpLayerRef = srcImage.getLayerByIndex(i)
            If tmpLayerRef.getLayerVisibility Then mergeLayers tmpLayerRef, tmpLayer, True
        Next i
    
    End If
    
    'Release the temp layer's hold on the destination DIB
    Set tmpLayer.layerDIB = Nothing
    
    'If the user requested non-premultiplied alpha, calculate it now.
    'TODO: how we handle premultiplication will depend on how I implement compositing, above, so this needs to be revisited
    '       after actual compositing has been written.
    If premultiplicationStatus Then
        
    Else
        If dstDIB.getDIBColorDepth = 32 Then dstDIB.fixPremultipliedAlpha False
    End If
    
End Sub

'Returns a subsection of the fully composited image (in pdDIB format, of course).  This is helpful for rendering the main viewport,
' as we only composite the relevant portions of the image.
'
'This function is large, complicated, and unfortunately tricky, as we must resize and composite each layer in turn.  Optimizing this
' function will be a huge project going forward, but we must optimize carefully if we want the viewport to work smoothly.
Friend Sub getCompositedRect(ByRef srcImage As pdImage, ByRef dstDIB As pdDIB, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcDIB As pdDIB, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal interpolationType As InterpolationMode)
    
    'Before doing anything else, fill two rects with the coordinates we were passed.  We will use these rects to calculate the
    ' boundary regions for all individual layers.
    Dim dstViewportRect As RECT, srcImageRect As RECT
    
    With dstViewportRect
        .Left = dstX
        .Top = dstY
        .Right = dstX + dstWidth
        .Bottom = dstY + dstHeight
    End With
    
    With srcImageRect
        .Left = srcX
        .Top = srcY
        .Right = srcX + srcWidth
        .Bottom = srcY + srcHeight
    End With
    
    'A few other things to note before we begin:
    ' - dstDIB will have already been created by the caller.  It may contain a background checkerboard, so we can't
    '    delete or recreate it.
    ' - The difference between the destination and source sizes can be used to infer a zoom value.  Note that aspect
    '    ratio will *always* be preserved by this function.
    ' - At present, GDI+ is used for all resizing.  Any other function with a StretchBlt-like interface could also
    '    be used, but GDI+ is currently the fastest method we have.
    
    'Start by creating a temporary DIB the size of the destination image (viewport).  All compositing will be done to this DIB,
    ' and as the final step, we will AlphaBlend the finished image onto dstDIB.
    Dim dstLayer As pdLayer
    Set dstLayer = New pdLayer
    dstLayer.layerDIB.createBlank dstWidth, dstHeight, 32
    
    'We also need a separate temporary DIB.  This DIB will be used to hold the resized chunk of each layer that fits within the
    ' specified viewport.  By minimizing the size of these DIBs, we can more quickly blend them, regardless of blend mode.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'The base layer will get blended onto the destination area first.  Its blend mode does not matter.  (Blend mode of the
    ' base layer is always ignored, because there are no pixels beneath it with which to blend.)
    
    'As part of the compositing process, we will only be working with the relevant on-screen chunks of each layer.  If a layer
    ' lies off the current viewport, it will be completely ignored.
    Dim xOffset As Double, yOffset As Double
    
    'Start by seeing if the base layer intersects with the viewport at all.  We solve this by finding the intersection rect of
    ' the current source rectangle, and the base layer.  If successful, this will return their intersection, as well as a
    ' non-zero function return value.
    Dim layerRect As RECT, intRectSrc As RECT, dstRect As RECT, srcRect As RECT
    
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long
    For i = 0 To srcImage.getNumOfLayers - 1
        
        Set tmpLayerRef = srcImage.getLayerByIndex(i)
        
        'Only process a layer if it is currently visible.
        If tmpLayerRef.getLayerVisibility Then
            
            'Retrieve a rect for this layer, relative to the full-size composite image
            fillRectForLayer tmpLayerRef, layerRect
    
            If IntersectRect(intRectSrc, srcImageRect, layerRect) <> 0 Then
            
                'The base layer intersects with the viewport rect, and the intersecting area is now contained within intRectSrc.
                ' We need to convert that into a relevant dstRect, so that we have both source and destination values for our
                ' resize function.
                getDstRectFromSrcRect dstRect, intRectSrc, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight
                
                'dstRect and intRectSrc now contain StretchBlt-compatible destination and source rectangles RELATIVE TO THE FULL IMAGE.
                ' Because the layer is unlikely to be the same size as the full image, we must come up with a final source rect that
                ' represents the source area, relative to the layer's DIB (which is probably smaller than the full image).
                With srcRect
                    .Left = intRectSrc.Left - tmpLayerRef.getLayerOffsetX
                    .Right = intRectSrc.Right - tmpLayerRef.getLayerOffsetX
                    .Top = intRectSrc.Top - tmpLayerRef.getLayerOffsetY
                    .Bottom = intRectSrc.Bottom - tmpLayerRef.getLayerOffsetY
                End With
                
                'We now have everything we need to create a resized version of this layer!  Grab it now.
                tmpDIB.createBlank dstRect.Right - dstRect.Left, dstRect.Bottom - dstRect.Top, 32, 0
                GDIPlusResizeDIB tmpDIB, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpLayerRef.layerDIB, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top, interpolationType
                                
                'tmpDIB now contains the chunk of this layer that appears on the viewport.  Time to composite it!
                If tmpLayerRef.getLayerBlendMode <> BL_NORMAL Then
                    compositeDIBs tmpDIB, dstLayer.layerDIB, tmpLayerRef.getLayerBlendMode, dstRect.Left - dstX, dstRect.Top - dstY
                End If
                
                tmpDIB.alphaBlendToDC dstLayer.layerDIB.getDIBDC, tmpLayerRef.getLayerOpacity * 2.55, dstRect.Left - dstX, dstRect.Top - dstY
            
            End If
            
        End If
        
    Next i
    
    'With all layers successfully blended onto tmpLayer's DIB, we can now perform a final alphaBlend onto dstDIB
    dstLayer.layerDIB.alphaBlendToDC dstDIB.getDIBDC, 255, dstX, dstY
    
    'Release the temp layer's hold on the destination DIB
    Set dstLayer.layerDIB = Nothing
    Set tmpDIB = Nothing
    
End Sub

'Given a layer, populate a rect with its coordinates (relative to the main image coordinates, always)
Friend Sub fillRectForLayer(ByRef srcLayer As pdLayer, ByRef dstRect As RECT)

    With srcLayer
        dstRect.Left = .getLayerOffsetX
        dstRect.Right = .getLayerOffsetX + .layerDIB.getDIBWidth
        dstRect.Top = .getLayerOffsetY
        dstRect.Bottom = .getLayerOffsetY + .layerDIB.getDIBHeight
    End With

End Sub

'Given StretchBlt-like parameters, create a destination rect that reflects the transformation of a source rect
' from the source coordinate space to the destination coordinate space.
Private Function getDstRectFromSrcRect(ByRef dstRect As RECT, ByRef srcRect As RECT, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long)
    
    'Convert the source rectangle coordinates into a [0, 1] scale relative to the total source area
    Dim srcRectF As RECTF
    With srcRectF
        .Left = (srcRect.Left - srcX) / srcWidth
        .Right = (srcRect.Right - srcX) / srcWidth
        .Top = (srcRect.Top - srcY) / srcHeight
        .Bottom = (srcRect.Bottom - srcY) / srcHeight
    End With
    
    'Using those source values, fill the destination rect
    With dstRect
        .Left = dstX + (srcRectF.Left * dstWidth)
        .Right = dstX + (srcRectF.Right * dstWidth)
        .Top = dstY + (srcRectF.Top * dstHeight)
        .Bottom = dstY + (srcRectF.Bottom * dstHeight)
    End With

End Function

Private Function Min(x As Single, y As Single) As Single
If x > y Then
      Min = y
 Else
      Min = x
End If
End Function

Private Function Max(x As Single, y As Single) As Single
If x < y Then
    Max = y
 Else
    Max = x
End If
End Function
