VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPixelBlender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Pixel Blender class
'Copyright 2014-2015 by Tanner Helland
'Created: 01/May/14
'Last updated: 14/October/15
'Last update: implement alpha inheritance
'
'This class exists for a single purpose: blending colors from two DIBs, using any of the blend and/or alpha modes
' supported by PD.
'
'Blending was previously handled inside pdCompositor, but as part of broad compositor optimizations in PD 7.0,
' it was split off into this dedicated class.  This allows us to cover a lot more edge cases with specially
' optimized code, without cluttering up the compositor (which already deals with way too much).
'
'Because this class is a WIP, not all blend modes have their own dedicated, specially optimized function.  For these,
' use the generic function, which is slowly being chopped up into smaller parts.
'
'Many thanks to expert contributor Audioglider for his help implementing PhotoDemon's many supported blend modes.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Temporary blending DIB.  We try to minimize how often we create this, but it allows us for much better organization throughout the class,
' because we aren't forced to modify the source DIBs we are passed if the caller doesn't want us to.
Private m_topDIBCopy As pdDIB

'Given two DIBs, blend them together and place the result inside the bottom DIB.
' IMPORTANT NOTES FOR THIS FUNCTION:
' - The top DIB will not be modified in any way.
' - A full blend (including alpha handling) will be applied.
' - The bottom DIB must already be large enough to hold the result of the composition.
' - The two DIBs should overlap at least partially.  If they don't, correctness of the function is not guaranteed.
Friend Sub blendDIBs(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal blendMode As LAYER_BLENDMODE, ByVal dstX As Single, ByVal dstY As Single, Optional ByVal alphaModifier As Double = 1#, Optional ByVal alphaMode As LAYER_ALPHAMODE = LA_NORMAL, Optional ByVal applyFinalBlend As Boolean = True)
    
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long, xOffsetBottom As Long, yOffsetBottom As Long
    
    'Calculate the overlap between the top and bottom DIBs.  The overlap function will automatically determine loop bounds for us,
    ' including safely rounding all floating-point coordinates to loop-friendly integers.
    If calculateDIBOverlap(topDIB, bottomDIB, dstX, dstY, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom) Then
        
        'With loop bounds safely calculated, we can now proceed with blending.
        
        'Because we are potentially doing a *lot* of calculations here, we need to make a temporary copy of the top DIB.  (Our copy will
        ' be a beaten, mangled mess by the end of this function, and we aren't allowed to make changes to the caller's copy.)
        If (m_topDIBCopy.getDIBWidth < topDIB.getDIBWidth) Or (m_topDIBCopy.getDIBHeight < topDIB.getDIBHeight) Then
            m_topDIBCopy.createBlank topDIB.getDIBWidth, topDIB.getDIBHeight, 32, 0, 0
        End If
        BitBlt m_topDIBCopy.getDIBDC, 0, 0, topDIB.getDIBWidth, topDIB.getDIBHeight, topDIB.getDIBDC, 0, 0, vbSrcCopy
        
        'To keep our individual blend functions as simple as possible, we apply a pre-processing alpha pass that handles several things:
        ' 1) Apply layer masks (if any)
        ' 2) Apply alpha inheritance (if any)
        
        'Note that we can skip these steps if none of these special alpha modes are active for the top DIB.
        
        If (alphaMode <> LA_NORMAL) Then
            preProcessAlphaEffects m_topDIBCopy, bottomDIB, alphaMode, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom, False
        End If
        
        'With alpha inheritance successfully applied to non-overlapping regions of the bottom layer, we can now forward further handling to
        ' one of several dedicated blend-mode compositor subs.  These subs apply the actual per-pixel blending loops, using the bounds we have
        ' determined for them.  They will also manually handle alpha inheritance for the region where the top and bottom layers overlap.
        Select Case blendMode
            
            'If this function is only being called for alpha handling reasons, but the blend mode is normal, we don't have to do anything here.
            Case BL_NORMAL
            
            'Blend modes with already-optimized functions are listed here
            Case BL_DARKEN
                applyBlendMode_Darken m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BL_MULTIPLY
                applyBlendMode_Multiply m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BL_LINEARDODGE
                applyBlendMode_LinearDodge m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BL_HARDLIGHT
                applyBlendMode_HardLight m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            'Some blendmodes don't have dedicated, optimized subs just yet.  This catch-all function handles their code for now.
            Case Else
                applyBlendMode_Generic m_topDIBCopy, bottomDIB, blendMode, alphaMode, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
        
        End Select
        
        'Composited results will have been placed inside the temporary top DIB copy.  Apply the final blend now!
        If applyFinalBlend Then m_topDIBCopy.alphaBlendToDCEx bottomDIB.getDIBDC, xOffsetBottom, yOffsetBottom, finalX - initX, finalY - initY, 0, 0, finalX - initX, finalY - initY, alphaModifier * 255
        
        'That's all there is to it!

    End If

End Sub

'Prior to blending two DIBs, their overlap region must be calculated.  (Only the intersection of the two DIBs needs to be composited;
' everything else can be ignored, saving us valuable time.)
'
'Besides top and bottom DIB references, the destination (x, y) of the top DIB (relative to the bottom DIB) must be supplied.
'
'Returns: TRUE and filled destination coordinates if overlap occurs; FALSE otherwise.
Private Function calculateDIBOverlap(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal dstX As Single, ByVal dstY As Single, ByRef initX As Long, ByRef initY As Long, ByRef finalX As Long, ByRef finalY As Long, ByRef xOffsetBottom As Long, ByRef yOffsetBottom As Long) As Boolean
    
    'Start by calculating bounds for the main per-pixel compositor loop.  These bounds are relative to the *top layer*, as we do
    ' not perform compositing on areas of the bottom layer that *are not overlapped* by the top layer.  (We just treat them as if
    ' the top layer is transparent in those regions, e.g. we ignore them.)
    
    'Start by finding the intersection between the top DIB rect and the bottom DIB rect.  To ensure edges of subpixel positioning aren't missed,
    ' we use RECTFs for the check.
    Dim topRect As RECTF, bottomRect As RECTF, IntersectRect As RECTF
    With topRect
        .Left = 0
        .Top = 0
        .Width = topDIB.getDIBWidth
        .Height = topDIB.getDIBHeight
    End With
    
    With bottomRect
        .Left = -dstX
        .Top = -dstY
        .Width = bottomDIB.getDIBWidth
        .Height = bottomDIB.getDIBHeight
    End With
    
    Dim dibsOverlap As Boolean
    dibsOverlap = IntersectRectF(IntersectRect, topRect, bottomRect)
    
    'If the DIBs do not overlap, our work here is done!
    ' (This should not technically be possible, but better safe than sorry...)
    If dibsOverlap Then
        
        'From the intersection rect, create integer-only bounds relative to the *top DIB*.
        initX = Int(IntersectRect.Left)
        initY = Int(IntersectRect.Top)
        finalX = Int(IntersectRect.Left + IntersectRect.Width + 0.9999)
        finalY = Int(IntersectRect.Top + IntersectRect.Height + 0.9999)
        If finalX > (topDIB.getDIBWidth - 1) Then finalX = topDIB.getDIBWidth - 1
        If finalY > (topDIB.getDIBHeight - 1) Then finalY = topDIB.getDIBHeight - 1
        
        'Individual blend functions need to know where the offset into the bottom DIB is, so calculate those last.
        xOffsetBottom = Fix(dstX)
        yOffsetBottom = Fix(dstY)
        
        'As a final failsafe, make sure the loop bounds never travel outside the *bottom* DIB's boundaries.
        If (xOffsetBottom + finalX) > (bottomDIB.getDIBWidth - 1) Then finalX = (bottomDIB.getDIBWidth - 1) - xOffsetBottom
        If (yOffsetBottom + finalY) > (bottomDIB.getDIBHeight - 1) Then finalY = (bottomDIB.getDIBHeight - 1) - yOffsetBottom
        
    End If
    
    calculateDIBOverlap = dibsOverlap
    
End Function

'Prior to blending two DIBs, call this function preprocess any alpha-specific effects (masks, inheritance, etc).  This greatly simplifies
' the actual blending code inside each blend function, because special alpha handling is never required.
Private Sub preProcessAlphaEffects(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal alphaMode As LAYER_ALPHAMODE, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long, Optional ByVal returnResultsUnpremultiplied As Boolean = False)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = i / 255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.getDIBArrayWidth:  bottomDIBPointer = bottomDIB.getActualDIBBits
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.getDIBArrayWidth:  topDIBPointer = topDIB.getActualDIBBits
    
    Dim bottomPixels() As Byte, bottomSA As SAFEARRAY1D
    prepInternalSafeArray1D bottomSA, bottomDIB
    
    Dim topPixels() As Byte, topSA As SAFEARRAY1D
    prepInternalSafeArray1D topSA, topDIB
    
    'Note: we don't actually copy the SafeArray structs into place until the loops start.
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomA As Single, bottomAInt As Byte
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        CopyMemory ByVal VarPtrArray(bottomPixels()), VarPtr(bottomSA), 4
        
        topSA.pvData = topDIBPointer + topScanlineSize * y
        CopyMemory ByVal VarPtrArray(topPixels()), VarPtr(topSA), 4
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely; they go untouched.
        If topAInt <> 0 Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'Alpha inheritance can only make pixels *more* transparent, not *less* transparent.  This means we can ignore opaque pixels,
            ' as they have no effect on the top layer.
            If bottomAInt < 255 Then
            
                'If the bottom pixel is transparent, we can simply erase the top pixel in this location
                If bottomAInt = 0 Then
                    topPixels(x) = 0
                    topPixels(x + 1) = 0
                    topPixels(x + 2) = 0
                    topPixels(x + 3) = 0
                    
                'If the bottom pixel is *not* transparent, our work is more involved
                Else
                    
                    'Retrieve the top layer's color and alpha values and convert them to floating-point
                    topA = intToFloat(topAInt)
                    topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                    
                    'We don't need to premultiply the values here; all we need to do is multiply everything by the bottom alpha value.
                    bottomA = intToFloat(bottomAInt)
                    topA = topA * bottomA
                    topB = topB * bottomA
                    topG = topG * bottomA
                    topR = topR * bottomA
                    
                    'Copy the final RGB values into the top layer, and premultiply them.
                    topPixels(x) = topB * 255: topPixels(x + 1) = topG * 255: topPixels(x + 2) = topR * 255: topPixels(x + 3) = topA * 255
                    
                End If
            
            'End "If bottomAInt < 255..."
            End If
        
        'End "If topAInt <> 0..."
        End If
    
    Next x
    Next y
    
    'Clear all temporary array references
    CopyMemory ByVal VarPtrArray(bottomPixels), 0&, 4
    CopyMemory ByVal VarPtrArray(topPixels), 0&, 4
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.setInitialAlphaPremultiplicationState True
    
End Sub

'Current catch-all for unoptimized blend mode calculations.  The goal is to move as many of these as possible to dedicated subs, where we can make
' things like dedicated lookup tables, contingent on the actual blend mode being applied.
Private Sub applyBlendMode_Generic(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal blendMode As LAYER_BLENDMODE, ByVal alphaMode As LAYER_ALPHAMODE, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)

    'In this generic blend sub, we will potentially be doing a lot of int/float conversions.  To speed things up, use a persistent look-up table.
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = i / 255
    Next i
    
    'In the inner loop, we'll be doing a bunch of compositing relative to newly calculated alpha value.  Storing some intermediate terms
    ' allows this to go much faster.
    Dim bottomAlphaIntermediate As Double
    
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
        quickBottomX = (xOffset + x) * 4
    For y = initY To finalY
    
        quickBottomY = yOffset + y
        
        'Retrieve top layer RGBA values
        topA = tImageData(QuickX + 3, y)
        
        'Ignore transparent pixels
        If topA <> 0 Then
        
            topR = tImageData(QuickX + 2, y)
            topG = tImageData(QuickX + 1, y)
            topB = tImageData(QuickX, y)
            
            'Retrieve bottom layer alpha
            bottomA = bImageData(quickBottomX + 3, quickBottomY)
            
            'If the bottom layer is transparent, we can simply copy over the top layer as-is
            If bottomA = 0 Then
                bImageData(quickBottomX, quickBottomY) = topB
                bImageData(quickBottomX + 1, quickBottomY) = topG
                bImageData(quickBottomX + 2, quickBottomY) = topR
                bImageData(quickBottomX + 3, quickBottomY) = topA
            Else
                
                'Convert to floating-point
                topR = intToFloat(topR)
                topG = intToFloat(topG)
                topB = intToFloat(topB)
                topA = intToFloat(topA)
                
                'Retrieve bottom layer values
                bottomA = intToFloat(bottomA)
                bottomR = intToFloat(bImageData(quickBottomX + 2, quickBottomY))
                bottomG = intToFloat(bImageData(quickBottomX + 1, quickBottomY))
                bottomB = intToFloat(bImageData(quickBottomX, quickBottomY))
                
                'Un-premultiply all RGB values
                If topA <> 1 Then
                    topR = topR / topA
                    topG = topG / topA
                    topB = topB / topA
                End If
                
                If bottomA <> 1 Then
                    bottomR = bottomR / bottomA
                    bottomG = bottomG / bottomA
                    bottomB = bottomB / bottomA
                End If
                
                'Apply the requested blend mode
                Select Case blendMode
                
                    'PD's rendering pipeline will not generally pass normal-mode layers to this function, as we can simply composite them
                    ' as-is using WAPI and/or GDI+.  But in the future, specialty features like alpha inheritance *will* require normal layers
                    ' to pass through this function, so I've added this branch now in preparation.
                    Case BL_NORMAL
                        newR = topR
                        newG = topG
                        newB = topB
                        
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
                        
                    Case BL_OVERLAY
                        If bottomR < 0.5 Then newR = 2 * bottomR * topR Else newR = 1 - 2 * (1 - bottomR) * (1 - topR)
                        If bottomG < 0.5 Then newG = 2 * bottomG * topG Else newG = 1 - 2 * (1 - bottomG) * (1 - topG)
                        If bottomB < 0.5 Then newB = 2 * bottomB * topB Else newB = 1 - 2 * (1 - bottomB) * (1 - topB)
                    
                    Case BL_SOFTLIGHT
                        newR = (1 - 2 * topR) * bottomR * bottomR + 2 * bottomR * topR
                        newG = (1 - 2 * topG) * bottomG * bottomG + 2 * bottomG * topG
                        newB = (1 - 2 * topB) * bottomB * bottomB + 2 * bottomB * topB
                    
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
                        
                    Case BL_GRAINEXTRACT
                        newR = bottomR - topR + 0.5
                        newG = bottomG - topG + 0.5
                        newB = bottomB - topB + 0.5
                        If newR > 1 Then newR = 1
                        If newG > 1 Then newG = 1
                        If newB > 1 Then newB = 1
                        If newR < 0 Then newR = 0
                        If newG < 0 Then newG = 0
                        If newB < 0 Then newB = 0
                    
                    Case BL_GRAINMERGE
                        newR = bottomR + topR - 0.5
                        newG = bottomG + topG - 0.5
                        newB = bottomB + topB - 0.5
                        If newR > 1 Then newR = 1
                        If newG > 1 Then newG = 1
                        If newB > 1 Then newB = 1
                        If newR < 0 Then newR = 0
                        If newG < 0 Then newG = 0
                        If newB < 0 Then newB = 0
                    
                    '****************************************************************************************************
                    ' IMPORTANT NOTE!
                    '
                    ' Blend modes that have already been optimized are included beneath this comment as a temporary
                    ' measure only.  They should not actually be applied via this function.  Use their dedicated functions
                    ' instead, as they are significantly faster.
                    '****************************************************************************************************
                    
                    Case BL_DARKEN
                        If bottomR < topR Then newR = bottomR Else newR = topR
                        If bottomG < topG Then newG = bottomG Else newG = topG
                        If bottomB < topB Then newB = bottomB Else newB = topB
                
                    Case BL_MULTIPLY
                        newR = topR * bottomR
                        newG = topG * bottomG
                        newB = topB * bottomB
                    
                    Case BL_LINEARDODGE
                        newR = topR + bottomR
                        newG = topG + bottomG
                        newB = topB + bottomB
                        If newR > 1 Then newR = 1
                        If newG > 1 Then newG = 1
                        If newB > 1 Then newB = 1
                    
                    Case BL_HARDLIGHT
                        If topR < 0.5 Then newR = 2 * bottomR * topR Else newR = 1 - 2 * (1 - bottomR) * (1 - topR)
                        If topG < 0.5 Then newG = 2 * bottomG * topG Else newG = 1 - 2 * (1 - bottomG) * (1 - topG)
                        If topB < 0.5 Then newB = 2 * bottomB * topB Else newB = 1 - 2 * (1 - bottomB) * (1 - topB)
                        
                    Case Else
                        newR = topR
                        newG = topG
                        newB = topB
                
                End Select
                
                'Alpha blend the original color with the new color, per the lower image's alpha.  (This way,
                ' if the bottom image is transparent, the top image is unchanged, regardless of blend mode.)
                '
                'Note that GIMP actually factors alpha into its compositing order, which I find odd and unintuitive.  However, it might be
                ' nice to someday provide a toggle for that behavior; I'm waiting to see if anyone actually requests it before
                ' embarking, however.
                If bottomA <> 1 Then
                    newR = (newR * bottomA) + (topR * (1 - bottomA))
                    newG = (newG * bottomA) + (topG * (1 - bottomA))
                    newB = (newB * bottomA) + (topB * (1 - bottomA))
                End If
                
                'Copy the final values into the top layer, and premultiply them.  Note that alpha remains unchanged; the final blend
                ' is handled by a separate function.
                topA = topA * 255
                tImageData(QuickX + 2, y) = newR * topA
                tImageData(QuickX + 1, y) = newG * topA
                tImageData(QuickX, y) = newB * topA
                    
            End If
            
        End If
    
    Next y
    Next x
    
    'Clear all array references
    CopyMemory ByVal VarPtrArray(bImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(tImageData), 0&, 4
    
    'DIBs must be premultiplied for this function to work, and they'll always leave the function in a premultiplied state.
    topDIB.setInitialAlphaPremultiplicationState True

End Sub

'Optimized "darken" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub applyBlendMode_Darken(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = i / 255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.getDIBArrayWidth:  bottomDIBPointer = bottomDIB.getActualDIBBits
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.getDIBArrayWidth:  topDIBPointer = topDIB.getActualDIBBits
    
    Dim bottomPixels() As Byte, bottomSA As SAFEARRAY1D
    prepInternalSafeArray1D bottomSA, bottomDIB
    
    Dim topPixels() As Byte, topSA As SAFEARRAY1D
    prepInternalSafeArray1D topSA, topDIB
    
    'Note: we don't actually copy the SafeArray structs into place until the loops start.
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single, newA As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        CopyMemory ByVal VarPtrArray(bottomPixels()), VarPtr(bottomSA), 4
        
        topSA.pvData = topDIBPointer + topScanlineSize * y
        CopyMemory ByVal VarPtrArray(topPixels()), VarPtr(topSA), 4
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If topAInt <> 0 Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'If the bottom pixel is transparent, we can simply copy over the top pixel as-is
            If bottomAInt = 0 Then
                bottomPixels(bottomX) = topPixels(x)
                bottomPixels(bottomX + 1) = topPixels(x + 1)
                bottomPixels(bottomX + 2) = topPixels(x + 2)
                bottomPixels(bottomX + 3) = topAInt
                
            'If the bottom pixel is *not* transparent, our work is more involved
            Else
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If topA <> 1 Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If bottomA <> 1 Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Darken" blend mode is very simple
                If bottomR < topR Then newR = bottomR Else newR = topR
                If bottomG < topG Then newG = bottomG Else newG = topG
                If bottomB < topB Then newB = bottomB Else newB = topB
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If bottomA <> 1 Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all temporary array references
    CopyMemory ByVal VarPtrArray(bottomPixels), 0&, 4
    CopyMemory ByVal VarPtrArray(topPixels), 0&, 4
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.setInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "multiply" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub applyBlendMode_Multiply(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = i / 255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.getDIBArrayWidth:  bottomDIBPointer = bottomDIB.getActualDIBBits
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.getDIBArrayWidth:  topDIBPointer = topDIB.getActualDIBBits
    
    Dim bottomPixels() As Byte, bottomSA As SAFEARRAY1D
    prepInternalSafeArray1D bottomSA, bottomDIB
    
    Dim topPixels() As Byte, topSA As SAFEARRAY1D
    prepInternalSafeArray1D topSA, topDIB
    
    'Note: we don't actually copy the SafeArray structs into place until the loops start.
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single, newA As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        CopyMemory ByVal VarPtrArray(bottomPixels()), VarPtr(bottomSA), 4
        
        topSA.pvData = topDIBPointer + topScanlineSize * y
        CopyMemory ByVal VarPtrArray(topPixels()), VarPtr(topSA), 4
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If topAInt <> 0 Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'If the bottom pixel is transparent, we can simply copy over the top pixel as-is
            If bottomAInt = 0 Then
                bottomPixels(bottomX) = topPixels(x)
                bottomPixels(bottomX + 1) = topPixels(x + 1)
                bottomPixels(bottomX + 2) = topPixels(x + 2)
                bottomPixels(bottomX + 3) = topAInt
                
            'If the bottom pixel is *not* transparent, our work is more involved
            Else
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If topA <> 1 Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If bottomA <> 1 Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Multiply" blend mode is very simple
                newR = topR * bottomR
                newG = topG * bottomG
                newB = topB * bottomB
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If bottomA <> 1 Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all temporary array references
    CopyMemory ByVal VarPtrArray(bottomPixels), 0&, 4
    CopyMemory ByVal VarPtrArray(topPixels), 0&, 4
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.setInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "linear dodge" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub applyBlendMode_LinearDodge(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = i / 255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.getDIBArrayWidth:  bottomDIBPointer = bottomDIB.getActualDIBBits
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.getDIBArrayWidth:  topDIBPointer = topDIB.getActualDIBBits
    
    Dim bottomPixels() As Byte, bottomSA As SAFEARRAY1D
    prepInternalSafeArray1D bottomSA, bottomDIB
    
    Dim topPixels() As Byte, topSA As SAFEARRAY1D
    prepInternalSafeArray1D topSA, topDIB
    
    'Note: we don't actually copy the SafeArray structs into place until the loops start.
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single, newA As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        CopyMemory ByVal VarPtrArray(bottomPixels()), VarPtr(bottomSA), 4
        
        topSA.pvData = topDIBPointer + topScanlineSize * y
        CopyMemory ByVal VarPtrArray(topPixels()), VarPtr(topSA), 4
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If topAInt <> 0 Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'If the bottom pixel is transparent, we can simply copy over the top pixel as-is
            If bottomAInt = 0 Then
                bottomPixels(bottomX) = topPixels(x)
                bottomPixels(bottomX + 1) = topPixels(x + 1)
                bottomPixels(bottomX + 2) = topPixels(x + 2)
                bottomPixels(bottomX + 3) = topAInt
                
            'If the bottom pixel is *not* transparent, our work is more involved
            Else
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If topA <> 1 Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If bottomA <> 1 Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Linear dodge" blend mode is very simple
                newR = topR + bottomR
                newG = topG + bottomG
                newB = topB + bottomB
                If newR > 1 Then newR = 1
                If newG > 1 Then newG = 1
                If newB > 1 Then newB = 1
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If bottomA <> 1 Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all temporary array references
    CopyMemory ByVal VarPtrArray(bottomPixels), 0&, 4
    CopyMemory ByVal VarPtrArray(topPixels), 0&, 4
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.setInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "hard light" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub applyBlendMode_HardLight(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = i / 255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.getDIBArrayWidth:  bottomDIBPointer = bottomDIB.getActualDIBBits
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.getDIBArrayWidth:  topDIBPointer = topDIB.getActualDIBBits
    
    Dim bottomPixels() As Byte, bottomSA As SAFEARRAY1D
    prepInternalSafeArray1D bottomSA, bottomDIB
    
    Dim topPixels() As Byte, topSA As SAFEARRAY1D
    prepInternalSafeArray1D topSA, topDIB
    
    'Note: we don't actually copy the SafeArray structs into place until the loops start.
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single, newA As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        CopyMemory ByVal VarPtrArray(bottomPixels()), VarPtr(bottomSA), 4
        
        topSA.pvData = topDIBPointer + topScanlineSize * y
        CopyMemory ByVal VarPtrArray(topPixels()), VarPtr(topSA), 4
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If topAInt <> 0 Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'If the bottom pixel is transparent, we can simply copy over the top pixel as-is
            If bottomAInt = 0 Then
                bottomPixels(bottomX) = topPixels(x)
                bottomPixels(bottomX + 1) = topPixels(x + 1)
                bottomPixels(bottomX + 2) = topPixels(x + 2)
                bottomPixels(bottomX + 3) = topAInt
                
            'If the bottom pixel is *not* transparent, our work is more involved
            Else
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If topA <> 1 Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If bottomA <> 1 Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Hard light" blend mode is very simple
                If topR < 0.5 Then newR = 2 * bottomR * topR Else newR = 1 - 2 * (1 - bottomR) * (1 - topR)
                If topG < 0.5 Then newG = 2 * bottomG * topG Else newG = 1 - 2 * (1 - bottomG) * (1 - topG)
                If topB < 0.5 Then newB = 2 * bottomB * topB Else newB = 1 - 2 * (1 - bottomB) * (1 - topB)
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If bottomA <> 1 Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all temporary array references
    CopyMemory ByVal VarPtrArray(bottomPixels), 0&, 4
    CopyMemory ByVal VarPtrArray(topPixels), 0&, 4
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.setInitialAlphaPremultiplicationState True
    
End Sub

'Given a 1D SafeArray struct and a target DIB, point the 1D array at the first scanline of the DIB, and initialize all relevant
' SafeArray parameters to match.
Private Sub prepInternalSafeArray1D(ByRef dstSafeArray As SAFEARRAY1D, ByRef srcDIB As pdDIB)

    With dstSafeArray
        .cbElements = 1
        .cDims = 1
        .lBound = 0
        .cElements = srcDIB.getDIBArrayWidth
        .pvData = srcDIB.getActualDIBBits
    End With
    
End Sub

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

'In order to render complex images quickly, this class must generate a lot of temporary objects.  If PD is under memory pressure, call this function
' to forcibly release as many temporary objects as possible.  (It should be obvious, but subsequent render calls will result in temporary objects
' being recreated, as necessary.)
Public Sub attemptToFreeMemory()
    
    'Erase any temporary DIBs we may have created
    m_topDIBCopy.eraseDIB
    
End Sub

Private Sub Class_Initialize()
    Set m_topDIBCopy = New pdDIB
End Sub

Private Sub Class_Terminate()
    Set m_topDIBCopy = Nothing
End Sub
