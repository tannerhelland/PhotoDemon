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
'Last updated: 12/October/15
'Last update: split blend-specific code away from pdCompositor, and into this class
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

'Given two DIBs, blend them together and place the result inside the top DIB.
' IMPORTANT NOTES FOR THIS FUNCTION:
' - Because the top DIB is receiving the result, *no alpha-blending will occur*.  Only blend mode and alpha inheritance
'   calculations will be applied to the top DIB.  The caller will need to perform a subsequent AlphaBlend (or equivalent).
' - Because the bottom DIB is untouched, regions not covered by the top DIB must be handled separately (TODO!)
' - The two DIBs should overlap at least partially.  If they don't, correctness of the function is not guaranteed.
Friend Sub blendDIBs_TopResult(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal blendMode As LAYER_BLENDMODE, ByVal dstX As Single, ByVal dstY As Single, Optional ByVal alphaMode As LAYER_ALPHAMODE = LA_NORMAL)
    
    'We will be doing a lot of int/float conversions.  To speed things up, use a persistent look-up table.
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = i / 255
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
        topA = intToFloat(tImageData(QuickX + 3, y))
        
        'Ignore transparent pixels
        If topA <> 0 Then
        
            topR = intToFloat(tImageData(QuickX + 2, y))
            topG = intToFloat(tImageData(QuickX + 1, y))
            topB = intToFloat(tImageData(QuickX, y))
            
            'Retrieve bottom layer RGB values
            bottomA = intToFloat(bImageData(quickBottomX + 3, quickBottomY))
            
            'Again, ignore transparent pixels
            If bottomA <> 0 Then
            
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
                
                'NOTE!  I've tried to write the compositor so that additional bounds checks are not needed.
                ' If errors ever occur in the future, you can start debugging by enabling the checks below.
                'If newR > 1 Then newR = 1
                'If newG > 1 Then newG = 1
                'If newB > 1 Then newB = 1
                
                'Copy the new values into the bottom layer, and premultiply them
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
    
End Sub

'Given two DIBs, blend them together and place the result inside the bottom DIB.
' IMPORTANT NOTES FOR THIS FUNCTION:
' - The top DIB will not be modified in any way.
' - A full blend (including alpha handling) will be applied.
' - The bottom DIB must already be large enough to hold the result of the composition.
' - The two DIBs should overlap at least partially.  If they don't, correctness of the function is not guaranteed.
Friend Sub blendDIBs_BottomResult(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal blendMode As LAYER_BLENDMODE, ByVal dstX As Single, ByVal dstY As Single, Optional ByVal alphaModifier As Double = 1#, Optional ByVal alphaMode As LAYER_ALPHAMODE = LA_NORMAL)
    
    'Start by calculating bounds for the main per-pixel compositor loop.  These bounds are relative to the *top layer*, as we do
    ' not perform compositing on areas of the bottom layer that *are not overlapped* by the top layer.  (We just treat them as if
    ' the top layer is transparent in those regions, e.g. we ignore them.)
    
    'Start by finding the intersection between the top DIB rect and the bottom DIB rect.
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
    
    'If the DIBs do not overlap, there's not much we can do here!
    ' (This should not technically be possible, but better safe than sorry)
    If dibsOverlap Then
        
        'From the intersection rect, create integer-only bounds relative to the top DIB.
        Dim initX As Long, initY As Long
        Dim finalX As Long, finalY As Long
        initX = Int(IntersectRect.Left)
        initY = Int(IntersectRect.Top)
        finalX = Int(IntersectRect.Left + IntersectRect.Width + 0.9999)
        finalY = Int(IntersectRect.Top + IntersectRect.Height + 0.9999)
        If finalX > (topDIB.getDIBWidth - 1) Then finalX = topDIB.getDIBWidth - 1
        If finalY > (topDIB.getDIBHeight - 1) Then finalY = topDIB.getDIBHeight - 1
        
        'Individual blend functions need to know where the offset into the bottom DIB is.
        Dim xOffsetBottom As Long, yOffsetBottom As Long
        xOffsetBottom = Fix(dstX)
        yOffsetBottom = Fix(dstY)
        
        'As a failsafe, make sure the new offsets don't lie outside the *bottom* DIB
        If (xOffsetBottom + finalX) > (bottomDIB.getDIBWidth - 1) Then finalX = (bottomDIB.getDIBWidth - 1) - xOffsetBottom
        If (yOffsetBottom + finalY) > (bottomDIB.getDIBHeight - 1) Then finalY = (bottomDIB.getDIBHeight - 1) - yOffsetBottom
        
        'With loop bounds correctly calculated, we can now proceed with blending.
        
        'Because we are potentially doing a *lot* of calculations here, we need to make a temporary copy of the top DIB.  (Our copy will
        ' be a beaten, mangled mess by the end of this function, and we aren't allowed to make changes to the caller's copy.)
        If (m_topDIBCopy.getDIBWidth < topDIB.getDIBWidth) Or (m_topDIBCopy.getDIBHeight < topDIB.getDIBHeight) Then
            m_topDIBCopy.createBlank topDIB.getDIBWidth, topDIB.getDIBHeight, 32, 0, 0
        End If
        BitBlt m_topDIBCopy.getDIBDC, 0, 0, topDIB.getDIBWidth, topDIB.getDIBHeight, topDIB.getDIBDC, 0, 0, vbSrcCopy
        
        'To keep our individual blend functions as simple as possible, we apply a pre-processing alpha pass that handles several things:
        ' 1) Apply an alpha modifier (e.g. custom layer opacity, if any)
        ' 2) Apply layer masks (if any)
        ' 2) Apply alpha inheritance (if any)
        
        'Note that we can skip these steps if none of these steps are relevant.
        
        'TODO: actually write the alpha preprocessor!
        
        'With alpha inheritance successfully applied to non-overlapping regions of the bottom layer, we can now forward further handling to
        ' one of several dedicated blend-mode compositor subs.  These subs apply the actual per-pixel blending loops, using the bounds we have
        ' determined for them.  They will also manually handle alpha inheritance for the region where the top and bottom layers overlap.
        Select Case blendMode
            
            'Blend modes with already-optimized functions are listed here
            Case BL_DARKEN
                applyBlendMode_Darken m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            'Some blendmodes don't have dedicated, optimized subs just yet.  This catch-all function handles their code for now.
            Case Else
                applyBlendMode_Generic m_topDIBCopy, bottomDIB, blendMode, alphaMode, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
        
        End Select
        
        'Composited results will have been stored in the top DIB.  Apply the final blend now!
        m_topDIBCopy.alphaBlendToDC bottomDIB.getDIBDC, alphaModifier * 255, xOffsetBottom, yOffsetBottom
        
        'That's all there is to it!

    End If

End Sub

'Prior to blending two DIBs, call this function preprocess any alpha-specific effects (masks, inheritance, etc).  This greatly simplifies
' the actual blending code inside each blend function.
Private Sub preProcessAlphaEffects(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal alphaMode As LAYER_ALPHAMODE, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long, Optional ByVal returnResultsUnpremultiplied As Boolean = False)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
        
    'Alpha requires a lot of floating-point math.  To speed things up, we use a persistent look-up table.
    ' TODO: profile this to see if it actually buys us any extra time.
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = i / 255
    Next i
        
    'In the inner loop, we'll be doing a bunch of compositing relative to newly calculated alpha value.  Storing some intermediate terms
    ' allows this to go much faster.
    Dim bottomAlphaIntermediate As Double
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, topScanlineSize As Long
    bottomScanlineSize = bottomDIB.getDIBArrayWidth
    topScanlineSize = topDIB.getDIBArrayWidth
    
    Dim bottomDIBPointer As Long, topDIBPointer As Long
    bottomDIBPointer = bottomDIB.getActualDIBBits
    topDIBPointer = topDIB.getActualDIBBits
    
    Dim bottomPixels() As Byte
    Dim bottomSA As SAFEARRAY1D
    With bottomSA
        .cbElements = 1
        .cDims = 1
        .lBound = 0
        .cElements = bottomScanlineSize
        .pvData = bottomDIBPointer
    End With
        
    Dim topPixels() As Byte
    Dim topSA As SAFEARRAY1D
    With topSA
        .cbElements = 1
        .cDims = 1
        .lBound = 0
        .cElements = topScanlineSize
        .pvData = topDIBPointer
    End With
    
    'We won't actually copy the SafeArray structs into place until the loops start
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single
    Dim newR As Single, newG As Single, newB As Single, newA As Single
    
    For y = initY To finalY
        
        'Point a 1D pixel array at the proper scanline of the top image
        topSA.pvData = topDIBPointer + topScanlineSize * y
        CopyMemory ByVal VarPtrArray(topPixels()), VarPtr(topSA), 4
        
        'We can completely ignore the bottom image during standard alpha processing
        If alphaMode <> LA_NORMAL Then
        
            'The main (x, y) variables point into the top image's coordinate space.  We need to calculate corresponding offsets
            ' into the bottom DIB's coordinate space.
            bottomY = yOffset + y
            bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
            CopyMemory ByVal VarPtrArray(bottomPixels()), VarPtr(bottomSA), 4
            
        End If
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topA = topPixels(x + 3)
            
        'Ignore transparent pixels.
        ' (NOTE: if alpha inheritance is active, we can't ignote transparent pixels; we may want to address that ability separately...)
        If topA <> 0 Then
                
            'Retrieve color and alpha from the *top* image, apply the passed alpha modifier (if any), and convert to floating-point
            topB = intToFloat(topPixels(x))
            topG = intToFloat(topPixels(x + 1))
            topR = intToFloat(topPixels(x + 2))
            topA = intToFloat(topPixels(x + 3))
            
            'If the caller wants our results un-premultiplied, unpremultiply alpha now.
            If returnResultsUnpremultiplied Then
                If topA <> 1 Then
                    topB = topB / topA
                    topG = topG / topA
                    topB = topB / topA
                End If
            End If
            
            'Store the newly modified values into the top DIB.
            topPixels(x) = topB * 255
            topPixels(x + 1) = topG * 255
            topPixels(x + 2) = topB * 255
            topPixels(x + 3) = topA * 255
                
        End If
        
    Next x
    Next y
    
    'Clear all array references
    CopyMemory ByVal VarPtrArray(bottomPixels), 0&, 4
    CopyMemory ByVal VarPtrArray(topPixels), 0&, 4
    
    'At present, this function only ever modifies the top DIB.
    topDIB.setInitialAlphaPremultiplicationState (Not returnResultsUnpremultiplied)
    
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

'Optimized "darken" blend mode function.  Supports storage of the end result in either the top DIB (allowing the caller to apply the final
' alpha blend in some later step), or the bottom DIB (and this function will handle blending for you).
'
'32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub applyBlendMode_Darken(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'In this generic blend sub, we will potentially be doing a lot of int/float conversions.  To speed things up, use a persistent look-up table.
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = i / 255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, topScanlineSize As Long
    bottomScanlineSize = bottomDIB.getDIBArrayWidth
    topScanlineSize = topDIB.getDIBArrayWidth
    
    Dim bottomDIBPointer As Long, topDIBPointer As Long
    bottomDIBPointer = bottomDIB.getActualDIBBits
    topDIBPointer = topDIB.getActualDIBBits
    
    Dim bottomPixels() As Byte
    Dim bottomSA As SAFEARRAY1D
    With bottomSA
        .cbElements = 1
        .cDims = 1
        .lBound = 0
        .cElements = bottomScanlineSize
        .pvData = bottomDIBPointer
    End With
        
    Dim topPixels() As Byte
    Dim topSA As SAFEARRAY1D
    With topSA
        .cbElements = 1
        .cDims = 1
        .lBound = 0
        .cElements = topScanlineSize
        .pvData = topDIBPointer
    End With
    
    'Note: we don't actually copy the SafeArray structs into place until the loops start.
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single
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
        topA = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If topA <> 0 Then
        
            topR = topPixels(x + 2)
            topG = topPixels(x + 1)
            topB = topPixels(x)
            
            'Retrieve bottom layer alpha
            bottomA = bottomPixels(bottomX + 3)
            
            'If the bottom pixel is transparent, we can simply copy over the top pixel as-is
            If bottomA = 0 Then
                bottomPixels(bottomX) = topB
                bottomPixels(bottomX + 1) = topG
                bottomPixels(bottomX + 2) = topR
                bottomPixels(bottomX + 3) = topA
                
            'If the bottom pixel is *not* transparent, our work is more involved
            Else
                
                'Convert to floating-point
                topR = intToFloat(topR)
                topG = intToFloat(topG)
                topB = intToFloat(topB)
                topA = intToFloat(topA)
                
                'Retrieve bottom layer values
                bottomA = intToFloat(bottomA)
                bottomR = intToFloat(bottomPixels(bottomX + 2))
                bottomG = intToFloat(bottomPixels(bottomX + 1))
                bottomB = intToFloat(bottomPixels(bottomX))
                
                'Un-premultiply all RGB values.  Note that alpha values cannot physically be zero, as we've already checked for that.
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
                
                '"Darken" blend mode is pretty damn simple
                If bottomR < topR Then newR = bottomR Else newR = topR
                If bottomG < topG Then newG = bottomG Else newG = topG
                If bottomB < topB Then newB = bottomB Else newB = topB
                    
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
                
                'Copy the final values into the top layer, and premultiply them.  Note that alpha remains unchanged; the final alpha blend
                ' is handled by a separate function.
                topA = topA * 255
                topPixels(x + 2) = newR * topA
                topPixels(x + 1) = newG * topA
                topPixels(x) = newB * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all array references
    CopyMemory ByVal VarPtrArray(bottomPixels), 0&, 4
    CopyMemory ByVal VarPtrArray(topPixels), 0&, 4
    
    'DIBs must be premultiplied for this function to work, and the target DIB generated by this function will always be
    ' premultiplied as well.
    topDIB.setInitialAlphaPremultiplicationState True
    
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
