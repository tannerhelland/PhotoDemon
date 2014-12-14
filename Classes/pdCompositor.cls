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
'Copyright ©2013-2014 by Tanner Helland
'Created: 01/May/14
'Last updated: 20/June/14
'Last update: add non-destructive edit support!
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
'As of 20 June '14, this class also supports the processing of non-destructive edits for a given layer.  At present,
' PD's stock layer edit functions are used, but it may be worthwhile to write ultra-fast versions specific to this
' class.  I'll investigate further as more non-destructive effects are implemented.
'
'Many thanks to expert contributor Audioglider for his help implementing PhotoDemon's many supported blend modes.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit
    
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
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single
    Dim newR As Single, newG As Single, newB As Single, newA As Single
    
    'Extra variables for transforms that require HSL conversions.
    Dim topH As Single, topS As Single, topV As Single
    Dim bottomH As Single, bottomS As Single, bottomV As Single
    
    For x = initX To finalX
        QuickX = x * 4
        quickBottomX = (Int(dstX) + x) * 4
    For y = initY To finalY
    
        quickBottomY = Int(dstY) + y
        
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
                        If newR < 0 Then newR = 0
                        If newG < 0 Then newG = 0
                        If newB < 0 Then newB = 0
                    
                    Case BL_LINEARBURN
                        newR = topR + bottomR - 1
                        newG = topG + bottomG - 1
                        newB = topB + bottomB - 1
                        If newR < 0 Then newR = 0
                        If newG < 0 Then newG = 0
                        If newB < 0 Then newB = 0
                        
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
                        If newR > 1 Then newR = 1
                        If newG > 1 Then newG = 1
                        If newB > 1 Then newB = 1
                        
                    Case BL_LINEARDODGE
                        newR = topR + bottomR
                        newG = topG + bottomG
                        newB = topB + bottomB
                        If newR > 1 Then newR = 1
                        If newG > 1 Then newG = 1
                        If newB > 1 Then newB = 1
                    
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
                        If newR > 1 Then newR = 1
                        If newG > 1 Then newG = 1
                        If newB > 1 Then newB = 1
                        If newR < 0 Then newR = 0
                        If newG < 0 Then newG = 0
                        If newB < 0 Then newB = 0
                    
                    Case BL_LINEARLIGHT
                        newR = bottomR + (2 * topR) - 1
                        newG = bottomG + (2 * topG) - 1
                        newB = bottomB + (2 * topB) - 1
                        If newR > 1 Then newR = 1
                        If newG > 1 Then newG = 1
                        If newB > 1 Then newB = 1
                        If newR < 0 Then newR = 0
                        If newG < 0 Then newG = 0
                        If newB < 0 Then newB = 0
                        
                    Case BL_PINLIGHT
                        If topR > 0.5 Then newR = fMax(bottomR, 2 * (topR - 0.5)) Else newR = fMin(bottomR, 2 * topR)
                        If topG > 0.5 Then newG = fMax(bottomG, 2 * (topG - 0.5)) Else newG = fMin(bottomG, 2 * topG)
                        If topB > 0.5 Then newB = fMax(bottomB, 2 * (topB - 0.5)) Else newB = fMin(bottomB, 2 * topB)
                        
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
                        If newR < 0 Then newR = 0
                        If newG < 0 Then newG = 0
                        If newB < 0 Then newB = 0
                    
                    Case BL_DIVIDE
                        If topR = 0 Then newR = 1 Else newR = bottomR / topR
                        If topG = 0 Then newG = 1 Else newG = bottomG / topG
                        If topB = 0 Then newB = 1 Else newB = bottomB / topB
                        If newR > 1 Then newR = 1
                        If newG > 1 Then newG = 1
                        If newB > 1 Then newB = 1
                        
                    Case BL_HUE
                        fastRGBtoHSV topR, topG, topB, topH, topS, topV
                        fastRGBtoHSV bottomR, bottomG, bottomB, bottomH, bottomS, bottomV
                        fHSVtoRGB topH, bottomS, bottomV, newR, newG, newB
                    
                    Case BL_SATURATION
                        fastRGBtoHSV topR, topG, topB, topH, topS, topV
                        fastRGBtoHSV bottomR, bottomG, bottomB, bottomH, bottomS, bottomV
                        fHSVtoRGB bottomH, topS, bottomV, newR, newG, newB
                    
                    Case BL_COLOR
                        fastRGBtoHSV topR, topG, topB, topH, topS, topV
                        bottomV = fMax3(bottomR, bottomG, bottomB)
                        fHSVtoRGB topH, topS, bottomV, newR, newG, newB
                    
                    Case BL_LUMINOSITY
                        topV = fMax3(topR, topG, topB)
                        fastRGBtoHSV bottomR, bottomG, bottomB, bottomH, bottomS, bottomV
                        fHSVtoRGB bottomH, bottomS, topV, newR, newG, newB
                    
                    Case Else
                        newR = topR
                        newG = topG
                        newB = topB
                
                End Select
                                
                If topA = 1 Then
                    
                    'Top layer is opaque, so premultiplication is irrelevant
                    
                    'TODO!  Fix compositing so these checks aren't necessary
                    If newR > 1 Then newR = 1
                    If newG > 1 Then newG = 1
                    If newB > 1 Then newB = 1
                    If newR < 0 Then newR = 0
                    If newG < 0 Then newG = 0
                    If newB < 0 Then newB = 0
                    
                    tImageData(QuickX + 2, y) = (newR * 255)
                    tImageData(QuickX + 1, y) = (newG * 255)
                    tImageData(QuickX, y) = (newB * 255)
                
                Else
                
                    'Alpha blend the original color with the new color, per the lower image's alpha.  (This way,
                    ' if the bottom image is transparent, the top image is unchanged.)
                    newR = (newR * bottomA) + (topR * (1 - bottomA))
                    newG = (newG * bottomA) + (topG * (1 - bottomA))
                    newB = (newB * bottomA) + (topB * (1 - bottomA))
                    
                    'TODO!  Fix compositing so these checks aren't necessary
                    If newR > 1 Then newR = 1
                    If newG > 1 Then newG = 1
                    If newB > 1 Then newB = 1
                    
                    'Copy the new values into the bottom layer, and premultiply them
                    tImageData(QuickX + 2, y) = (newR * 255) * topA
                    tImageData(QuickX + 1, y) = (newG * 255) * topA
                    tImageData(QuickX, y) = (newB * 255) * topA
                    
                End If
                
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
        Layer_Handler.fillRectForLayer bottomLayer, bottomRect, True
        Layer_Handler.fillRectForLayer topLayer, topRect, True
        UnionRect finalRect, topRect, bottomRect
        
        'finalRect now contains the coordinates of the union rect.  Create a blank DIB at those dimensions.
        tmpDIB.createBlank finalRect.Right - finalRect.Left, finalRect.Bottom - finalRect.Top, 32, 0
        
        'We now need to do a couple of things.  Let's start by copying the bottom DIB into this new temporary DIB.
        xOffset = bottomRect.Left - finalRect.Left
        yOffset = bottomRect.Top - finalRect.Top
        
        If (bottomLayer.getLayerCanvasXModifier = 1) And (bottomLayer.getLayerCanvasYModifier = 1) Then
            BitBlt tmpDIB.getDIBDC, xOffset, yOffset, bottomLayer.layerDIB.getDIBWidth, bottomLayer.layerDIB.getDIBHeight, bottomLayer.layerDIB.getDIBDC, 0, 0, vbSrcCopy
        Else
            GDIPlusResizeDIB tmpDIB, xOffset, yOffset, bottomLayer.getLayerCanvasWidthModified, bottomLayer.getLayerCanvasHeightModified, bottomLayer.layerDIB, 0, 0, bottomLayer.layerDIB.getDIBWidth, bottomLayer.layerDIB.getDIBHeight, InterpolationModeHighQualityBicubic
            
            'Also, reset the canvas modifiers for the bottom layer now, as we have effectively nullified them by this resize
            bottomLayer.setLayerCanvasXModifier 1
            bottomLayer.setLayerCanvasYModifier 1
        End If
        
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
    
        'Make a copy of the current layer, with any width/height modifications applied
        'Before proceeding further, we must check the top layer for any size modifications; if found, apply them now.
        If (topLayer.getLayerCanvasXModifier = 1) And (topLayer.getLayerCanvasYModifier = 1) Then
            tmpDIB.createFromExistingDIB topLayer.layerDIB
        Else
            tmpDIB.createBlank topLayer.getLayerCanvasWidthModified, topLayer.getLayerCanvasHeightModified, 32, 0
            GDIPlusResizeDIB tmpDIB, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, topLayer.layerDIB, 0, 0, topLayer.layerDIB.getDIBWidth, topLayer.layerDIB.getDIBHeight, InterpolationModeHighQualityBicubic
        End If
        
        'tmpDIB now contains the chunk of this layer that appears on the viewport.  Before we can composite it against the base layer,
        ' we need to check for any non-destructive FX that may be active.
        ' (TODO: check layer size vs non-destructive size, and apply the non-destructive effect to whichever version is smaller!
        '        That could give us a minor speed boost.)
        If topLayer.getLayerNonDestructiveFXState Then applyNDFXToDIB topLayer, tmpDIB
                
        'With all non-destructive options accounted for, we can finally calculate composited layer data!
        If topLayer.getLayerBlendMode <> BL_NORMAL Then
            compositeDIBs tmpDIB, bottomLayer.layerDIB, topLayer.getLayerBlendMode, xOffset, yOffset
        End If
    
        'With compositing complete, apply the final step, which is alpha-blending the layer downward while accounting for layer opacity.
        tmpDIB.alphaBlendToDC bottomLayer.layerDIB.getDIBDC, topLayer.getLayerOpacity * 2.55, xOffset, yOffset
        
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
    
    'We may need a temporary DIB to handle dynamic layer resizing
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'We will be accessing all layers of the source image in turn.  Rather than re-requesting those layers through
    ' functions, we will simply use a local reference.
    Dim tmpLayerRef As pdLayer
    
    'Always start by rendering the bottom layer onto the (presently empty) destination DIB.  The bottom layer's blend mode
    ' does not matter, so this is a straightforward operation.
    Set tmpLayerRef = srcImage.getLayerByIndex(0)
    If tmpLayerRef.getLayerVisibility Then
    
        'Check to see if the base layer has any size modifications; if it does, apply them now
        If tmpLayerRef.getLayerCanvasXModifier = 1 And tmpLayerRef.getLayerCanvasYModifier = 1 Then
            
            'No size modifications are active; copy the layer as-is
            tmpDIB.createFromExistingDIB tmpLayerRef.layerDIB
            
            'Apply any non-destructive effects now
            If tmpLayerRef.getLayerNonDestructiveFXState Then applyNDFXToDIB tmpLayerRef, tmpDIB
            
        Else
            
            With tmpLayerRef
                
                tmpDIB.createBlank .getLayerCanvasWidthModified, .getLayerCanvasHeightModified, 32, 0
                GDIPlusResizeDIB tmpDIB, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, .layerDIB, 0, 0, .layerDIB.getDIBWidth, .layerDIB.getDIBHeight, InterpolationModeHighQualityBicubic
                
                'Before we can composite this temporary DIB against other DIBs, we need to check for any non-destructive FX that may be active.
                ' (TODO: check layer size vs non-destructive size, and apply the non-destructive effect to whichever version is smaller!
                '        That could give us a minor speed boost.)
                If tmpLayerRef.getLayerNonDestructiveFXState Then applyNDFXToDIB tmpLayerRef, tmpDIB
                
            End With
        
        End If
        
        'Blend the bottom layer onto the destination DIB
        tmpDIB.alphaBlendToDC dstDIB.getDIBDC, tmpLayerRef.getLayerOpacity * 2.55, tmpLayerRef.getLayerOffsetX, tmpLayerRef.getLayerOffsetY
    
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
Friend Sub getCompositedRect(ByRef srcImage As pdImage, ByRef dstDIB As pdDIB, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal interpolationType As InterpolationMode)
    
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
    ' TODO: make this temporary layer static, so we don't have to constantly recreate it.
    Dim dstLayer As pdLayer
    Set dstLayer = New pdLayer
    dstLayer.layerDIB.createBlank Int(dstWidth + 0.999), Int(dstHeight + 0.999), 32
    
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
    ' Boolean indicating whether the layers intersect.
    Dim layerRect As RECTF, intRectSrc As RECTF, dstRect As RECTF, srcRect As RECTF
    
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long
    For i = 0 To srcImage.getNumOfLayers - 1
        
        Set tmpLayerRef = srcImage.getLayerByIndex(i)
        
        'Only process a layer if it is currently visible.
        If tmpLayerRef.getLayerVisibility Then
            
            'Retrieve a rect for this layer, relative to the full-size composite image
            Layer_Handler.fillRectForLayerF tmpLayerRef, layerRect, True
            
            'Check for an intersection between the target rect (passed to this function) and the base layer rect.
            If GDI_Plus.IntersectRectF(intRectSrc, srcImageRect, layerRect) Then
                
                'The composite operation works similar to BitBlt, by calculating a destination rectangle and a source rectangle.
                ' We know that these two rectangles define the same rectangle of the image, which allows us to infer zoom as well.
                
                'Now that we know the region of this layer that corresponds to the requested destination rect, we will use the
                ' getDstRectFromSrcRectF() function to translate our already calculated eight bare destination and source values
                ' (x/y/width/height) and retrieve a usable destination rect for this layer, defining the destination region of
                ' the target covered by this layer.
                getDstRectFromSrcRectF dstRect, intRectSrc, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight
                
                'dstRect and intRectSrc now contain StretchBlt-compatible destination and source rectangles RELATIVE TO THE FULL IMAGE.
                
                'Because the current layer may not be the same size as the full image, we must perform one final translation: a source
                ' rect that represents the source area, relative to the current layer's DIB.  This step is important if the layer has
                ' had non-destructive resizing applied, because this is when we apply that calculation.
                With srcRect
                    .Left = (intRectSrc.Left - tmpLayerRef.getLayerOffsetX) * (1 / tmpLayerRef.getLayerCanvasXModifier)
                    .Width = intRectSrc.Width * (1 / tmpLayerRef.getLayerCanvasXModifier)
                    .Top = (intRectSrc.Top - tmpLayerRef.getLayerOffsetY) * (1 / tmpLayerRef.getLayerCanvasYModifier)
                    .Height = intRectSrc.Height * (1 / tmpLayerRef.getLayerCanvasYModifier)
                End With
                
                'With all coordinate math complete, we now proceed with the actual compositing!
                
                'Two conditions require us to create a temporary DIB prior to actually compositing this layer:
                ' 1) Non-destructive effects
                ' 2) Non-standard blend modes
                
                'These conditions are unpleasant, because additional per-pixel processing is required, instead of a basic alpha-blending
                ' that we can let GDI+ automatically handle.
                
                'Check for such conditions now.
                If tmpLayerRef.getLayerNonDestructiveFXState Or (tmpLayerRef.getLayerBlendMode <> BL_NORMAL) Then
                
                    'This layer has one or more "special cases" active.  Create a temporary DIB to reference the image's bits.
                    ' (TODO: skip the temporary DIB step if non-destructive effects are not being used.)
                    tmpDIB.createBlank dstRect.Width + 0.999 + (dstRect.Left - Int(dstRect.Left)), dstRect.Height + 0.999 + (dstRect.Top - Int(dstRect.Top)), 32, 0
                    
                    'Check for the special case of 100% zoom; when this happens, and the image has no non-destructive resizing applied,
                    ' we can completely skip the resize step and simply use BitBlt (way faster)
                    If isRectZoomless And (tmpLayerRef.getLayerCanvasXModifier = 1) And (tmpLayerRef.getLayerCanvasYModifier = 1) Then
                        
                        BitBlt tmpDIB.getDIBDC, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpLayerRef.layerDIB.getDIBDC, srcRect.Left, srcRect.Top, vbSrcCopy
                        
                        'Apply non-destructive FX now
                        If tmpLayerRef.getLayerNonDestructiveFXState Then applyNDFXToDIB tmpLayerRef, tmpDIB
                    
                    Else
                    
                        'If zoomed-in, we can switch to StretchBlt for performance gains.  (As nearest-neighbor interpolation will
                        ' be used regardless of the caller's requested interpolation type.)
                        If srcWidth < dstWidth Then
                            
                            'We are zoomed-in to the image, e.g. zoom is larger than 100%.  Check for non-destructive FX being present.
                            
                            'Non-destructive FX are active.  To reduce processing time, chop out the source rectangle, apply the non-destructive FX,
                            ' and *then* resize the data.
                            If tmpLayerRef.getLayerNonDestructiveFXState Then
                            
                                'Extract the smaller source rectangle into its own DIB
                                Dim tmpNDFXDIB As pdDIB
                                Set tmpNDFXDIB = New pdDIB
                                
                                tmpNDFXDIB.createBlank srcRect.Width + 0.999 + (srcRect.Left - Int(srcRect.Left)), srcRect.Height + 0.999 + (srcRect.Top - Int(srcRect.Top)), 32, 0
                                BitBlt tmpNDFXDIB.getDIBDC, 0, 0, tmpNDFXDIB.getDIBWidth, tmpNDFXDIB.getDIBHeight, tmpLayerRef.layerDIB.getDIBDC, srcRect.Left, srcRect.Top, vbSrcCopy
                                
                                'Apply the non-destructive FX to the smaller source DIB; this is faster than applying it after the resize action,
                                ' as we only have to process a small fraction of the pixels
                                applyNDFXToDIB tmpLayerRef, tmpNDFXDIB
                                
                                'Copy the processed DIB into tmpDIB, which will be used for compositing later in this function
                                If (tmpLayerRef.getLayerCanvasXModifier = 1) And (tmpLayerRef.getLayerCanvasYModifier = 1) Then
                                    StretchBlt tmpDIB.getDIBDC, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpNDFXDIB.getDIBDC, 0, 0, tmpNDFXDIB.getDIBWidth, tmpNDFXDIB.getDIBHeight, vbSrcCopy
                                Else
                                    GDIPlusResizeDIB tmpDIB, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpNDFXDIB, 0, 0, tmpNDFXDIB.getDIBWidth, tmpNDFXDIB.getDIBHeight, InterpolationModeNearestNeighbor
                                End If
                                
                                'Erase the NDFX DIB
                                Set tmpNDFXDIB = Nothing
                                
                            'Non-destructive FX are not active.  Resize the source rectangle immediately.
                            Else
                            
                                'Check for non-destructive resizing.  If present, switch to GDI+; otherwise, StretchBlt is sufficient.
                                If (tmpLayerRef.getLayerCanvasXModifier = 1) And (tmpLayerRef.getLayerCanvasYModifier = 1) Then
                                    StretchBlt tmpDIB.getDIBDC, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpLayerRef.layerDIB.getDIBDC, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, vbSrcCopy
                                Else
                                    GDIPlusResizeDIB tmpDIB, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpLayerRef.layerDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, InterpolationModeNearestNeighbor
                                End If
                                
                            End If
                        
                        'We are zoomed-out, so StretchBlt cannot be used (as it will cause artifacting for 32bpp data).  GDI+ is our only
                        ' option at present.
                        Else
                            
                            'If the user has requested MAX PERFORMANCE rendering, override the requested interpolation type with nearest neighbor
                            GDIPlusResizeDIB tmpDIB, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpLayerRef.layerDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, interpolationType
                            
                            'Apply non-destructive FX now
                            If tmpLayerRef.getLayerNonDestructiveFXState Then applyNDFXToDIB tmpLayerRef, tmpDIB
                            
                        End If
                        
                    End If
                    
                    'tmpDIB now contains the chunk of this layer that appears on the viewport, with all non-destructive edits applied.
                    ' Time to composite it!
                    If tmpLayerRef.getLayerBlendMode <> BL_NORMAL Then
                        compositeDIBs tmpDIB, dstLayer.layerDIB, tmpLayerRef.getLayerBlendMode, dstRect.Left - dstX, dstRect.Top - dstY
                    End If
                    
                    tmpDIB.alphaBlendToDC dstLayer.layerDIB.getDIBDC, tmpLayerRef.getLayerOpacity * 2.55, dstRect.Left - dstX, dstRect.Top - dstY
                
                'This is a simple layer, with no non-destructive effects or weird blend modes.  Composite it instantly, without using
                ' a temporary placeholder layer.
                Else
                
                    'GDIPlusResizeDIB tmpDIB, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpLayerRef.layerDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, InterpolationModeNearestNeighbor
                    GDIPlus_StretchBlt dstLayer.layerDIB, dstRect.Left - dstX, dstRect.Top - dstY, dstRect.Width, dstRect.Height, tmpLayerRef.layerDIB, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height, tmpLayerRef.getLayerOpacity / 100
                    
                End If
                
            'END block for intersecting viewports
            End If
        
        'END block for layer visibility toggle
        End If
        
    Next i
    
    'With all layers successfully blended onto tmpLayer's DIB, we can now perform a final alphaBlend onto dstDIB
    'dstLayer.layerDIB.alphaBlendToDC dstDIB.getDIBDC, 255, dstX, dstY
    GDIPlus_StretchBlt dstDIB, dstX, dstY, dstWidth, dstHeight, dstLayer.layerDIB, dstX - Int(dstX), dstY - Int(dstY), dstWidth, dstHeight, 1, InterpolationModeHighQualityBicubic
    Debug.Print "Final GDI+ calc: ", dstX, dstY, dstWidth, dstHeight
    
    'Release the temp layer's hold on the destination DIB
    Set dstLayer.layerDIB = Nothing
    Set tmpDIB = Nothing
    
End Sub

'Given a layer containing one or more non-destructive effect parameters, and a target DIB, apply the layer's specified FX against
' the DIB.  Why supply the DIB separately?  Because during viewport composition, all layers are resized and/or cropped to their
' size relevant to the viewport.  By doing this, we hugely decrease the performance hit from non-destructive edit calculations.
Public Sub applyNDFXToDIB(ByRef parentLayer As pdLayer, ByRef targetDIB As pdDIB)

    'Start by unpremultiplying the layer alpha
    If targetDIB.getDIBColorDepth = 32 Then targetDIB.fixPremultipliedAlpha False

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
    If targetDIB.getDIBColorDepth = 32 Then targetDIB.fixPremultipliedAlpha True

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
    Dim grayLookUp() As Long
    ReDim grayLookUp(0 To 765) As Long
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'To handle temperature and tint, we will also be preparing RGB lookup tables
    Dim rLookup() As Byte, gLookUp() As Byte, bLookup() As Byte
    ReDim rLookup(0 To 255) As Byte
    ReDim gLookUp(0 To 255) As Byte
    ReDim bLookup(0 To 255) As Byte
    
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
        rLookup(x) = Round(r)
        gLookUp(x) = Round(g)
        bLookup(x) = Round(b)
    
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
            avgVal = grayLookUp(r + g + b)
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
        ImageData(QuickVal + 2, y) = rLookup(r)
        ImageData(QuickVal + 1, y) = gLookUp(g)
        ImageData(QuickVal, y) = bLookup(b)
        
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

'A heavily modified RGB to HSV transform, courtesy of http://lolengine.net/blog/2013/01/13/fast-rgb-to-hsv.
' Note that the code assumes RGB values already in the [0, 1] range, and it will return HSV values in the [0, 1] range.
Private Sub fastRGBtoHSV(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByRef h As Single, ByRef s As Single, ByRef v As Single)

    Dim k As Single, tmpSwap As Single, chroma As Single
    
    If (g < b) Then
        tmpSwap = b
        b = g
        g = tmpSwap
        k = -1
    End If
    
    If (r < g) Then
        tmpSwap = g
        g = r
        r = tmpSwap
        k = -(2 / 6) - k
    End If
    
    chroma = r - fMin(g, b)
    h = Abs(k + (g - b) / (6 * chroma + 0.0000001))
    s = chroma / (r + 0.00000001)
    v = r
    
End Sub

'Convert [0,1] HSV values to [0,255] RGB values, with thanks to easyrgb.com for the conversion math
Private Sub fHSVtoRGB(ByRef h As Single, ByRef s As Single, ByRef v As Single, ByRef r As Single, ByRef g As Single, ByRef b As Single)

    'If saturation is 0, RGB are calculated identically
    If s = 0 Then
        r = v
        g = v
        b = v
        Exit Sub
    
    'If saturation is not 0, we have to calculate RGB independently
    Else
       
        Dim var_H As Single
        var_H = h * 6
        
        'To keep our math simple, limit hue to [0, 5.9999999]
        If var_H >= 6 Then var_H = 0
        
        Dim var_I As Long
        var_I = Int(var_H)
        
        Dim var_1 As Single, var_2 As Single, var_3 As Single
        var_1 = v * (1 - s)
        var_2 = v * (1 - s * (var_H - var_I))
        var_3 = v * (1 - s * (1 - (var_H - var_I)))
        
        Select Case var_I
        
            Case 0
                r = v
                g = var_3
                b = var_1
                
            Case 1
                r = var_2
                g = v
                b = var_1
                
            Case 2
                r = var_1
                g = v
                b = var_3
                
            Case 3
                r = var_1
                g = var_2
                b = v
            
            Case 4
                r = var_3
                g = var_1
                b = v
                
            Case Else
                r = v
                g = var_1
                b = var_2
                
        End Select
                
    End If

End Sub

'Return the minimum of two floating-point values
Private Function fMin(x As Single, y As Single) As Single
    If x > y Then fMin = y Else fMin = x
End Function

'Return the maximum of two floating-point values
Private Function fMax(x As Single, y As Single) As Single
    If x < y Then fMax = y Else fMax = x
End Function

'Return the maximum of three floating point values
Private Function fMax3(rR As Single, rG As Single, rB As Single) As Single
   If (rR > rG) Then
      If (rR > rB) Then
         fMax3 = rR
      Else
         fMax3 = rB
      End If
   Else
      If (rB > rG) Then
         fMax3 = rB
      Else
         fMax3 = rG
      End If
   End If
End Function

'Return the minimum of three floating point values
Private Function fMin3(rR As Single, rG As Single, rB As Single) As Single
   If (rR < rG) Then
      If (rR < rB) Then
         fMin3 = rR
      Else
         fMin3 = rB
      End If
   Else
      If (rB < rG) Then
         fMin3 = rB
      Else
         fMin3 = rG
      End If
   End If
End Function

