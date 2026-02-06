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
'Copyright 2014-2026 by Tanner Helland
'Created: 01/May/14
'Last updated: 03/March/25
'Last update: implement "Overwrite/Replace" blend mode
'
'This class exists for a single purpose: blending colors from two DIBs, using any of the
' blend and/or alpha modes supported by PhotoDemon.
'
'Blending was previously handled inside pdCompositor, but as part of broad compositor
' optimizations in PD 7.0, it was split off into this dedicated class.  This allows us to
' cover many more edge cases with specially optimized code, without cluttering up the
' compositor (which already deals with way too many tasks).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Temporary blending DIB.  We try to minimize how often we create this, but it allows us for much better organization throughout
' the class, because we aren't forced to modify the source DIBs we are passed if the caller doesn't want us to.
Private m_topDIBCopy As pdDIB

'If we are asked to blend two DIBs but the bottom one has its alpha channel locked, we make a temporary copy of its alpha values.
' After the blend, we restore the original values (regardless of any changes caused by the blend).
Private m_AlphaCopy() As Byte

Private Const ONE_DIV_255 As Single = 1! / 255!

'Given two DIBs, blend them together and place the result inside the bottom DIB.
' IMPORTANT NOTES FOR THIS FUNCTION:
' - The top DIB will not be modified in any way.
' - A full blend (including alpha handling) will be applied.
' - The bottom DIB must already be large enough to hold the result of the composition.
' - The two DIBs should overlap at least partially.  If they don't, correctness of the function is not guaranteed.
Friend Sub BlendDIBs(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal blendMode As PD_BlendMode, ByVal dstX As Single, ByVal dstY As Single, Optional ByVal alphaModifier As Single = 1!, Optional ByVal bottomAlphaMode As PD_AlphaMode = AM_Normal, Optional ByVal topAlphaMode As PD_AlphaMode = AM_Normal, Optional ByVal applyFinalBlend As Boolean = True, Optional ByVal ptrToTopLayerAlternateRectF As Long = 0, Optional ByRef topLayerMask As pdDIB = Nothing)
    
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long, xOffsetBottom As Long, yOffsetBottom As Long
    
    'Calculate the overlap between the top and bottom DIBs.  The overlap function will automatically determine loop bounds for us,
    ' including safely rounding all floating-point coordinates to loop-friendly integers.
    If CalculateDIBOverlap(topDIB, bottomDIB, dstX, dstY, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom, ptrToTopLayerAlternateRectF) Then
        
        'With loop bounds safely calculated, we can now proceed with blending.
        
        'Because we are potentially doing a *lot* of calculations here, we need to make a temporary copy of the top DIB.
        ' (Our copy will be a beaten, mangled mess by the end of this function, and we aren't allowed to make changes
        ' to the caller's copy.)
        If (m_topDIBCopy.GetDIBWidth <> topDIB.GetDIBWidth) Or (m_topDIBCopy.GetDIBHeight <> topDIB.GetDIBHeight) Then
            m_topDIBCopy.CreateBlank topDIB.GetDIBWidth, topDIB.GetDIBHeight, 32, 0, 0
        End If
        
        GDI.BitBltWrapper m_topDIBCopy.GetDIBDC, 0, 0, topDIB.GetDIBWidth, topDIB.GetDIBHeight, topDIB.GetDIBDC, 0, 0, vbSrcCopy
        m_topDIBCopy.SetInitialAlphaPremultiplicationState topDIB.GetAlphaPremultiplication
        
        'To keep our individual blend functions as simple as possible, we apply a pre-processing alpha pass
        ' that handles several things:
        ' 1) Apply layer masks (if any)
        ' 2) Apply alpha inheritance (if any)
        ' 3) Overwrite blend mode on paintbrush (identified by presence of a top layer mask)
        
        'Note that we can skip these steps if none of these special alpha modes are active for the top DIB.
        If (topAlphaMode <> AM_Normal) Or (bottomAlphaMode <> AM_Normal) Or (Not topLayerMask Is Nothing) Then
            PreProcessAlphaEffects m_topDIBCopy, bottomDIB, bottomAlphaMode, topAlphaMode, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom, False, ptrToTopLayerAlternateRectF, topLayerMask
        End If
        
        'With alpha inheritance successfully applied to non-overlapping regions of the bottom layer, we can now forward further handling to
        ' one of several dedicated blend-mode compositor subs.  These subs apply the actual per-pixel blending loops, using the bounds we have
        ' determined for them.  They will also manually handle alpha inheritance for the region where the top and bottom layers overlap.
        Select Case blendMode
            
            'If this function is only being called for alpha handling reasons, but the blend mode is normal, we don't have to do anything here.
            Case BM_Normal
            
            'Blend modes with already-optimized functions are listed here
            Case BM_Color
                ApplyBlendMode_Color m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_ColorBurn
                ApplyBlendMode_ColorBurn m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_ColorDodge
                ApplyBlendMode_ColorDodge m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_Darken
                ApplyBlendMode_Darken m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_Difference
                ApplyBlendMode_Difference m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_Divide
                ApplyBlendMode_Divide m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_Exclusion
                ApplyBlendMode_Exclusion m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_GrainExtract
                ApplyBlendMode_GrainExtract m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_GrainMerge
                ApplyBlendMode_GrainMerge m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_Hue
                ApplyBlendMode_Hue m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_HardLight
                ApplyBlendMode_HardLight m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_HardMix
                ApplyBlendMode_HardMix m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_Lighten
                ApplyBlendMode_Lighten m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_LinearBurn
                ApplyBlendMode_LinearBurn m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_LinearDodge
                ApplyBlendMode_LinearDodge m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_LinearLight
                ApplyBlendMode_LinearLight m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_Luminosity
                ApplyBlendMode_Luminosity m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_Multiply
                ApplyBlendMode_Multiply m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_Overlay
                ApplyBlendMode_Overlay m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_PinLight
                ApplyBlendMode_PinLight m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_Saturation
                ApplyBlendMode_Saturation m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_Screen
                ApplyBlendMode_Screen m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_SoftLight
                ApplyBlendMode_SoftLight m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
            
            Case BM_Subtract
                ApplyBlendMode_Subtract m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_VividLight
                ApplyBlendMode_VividLight m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom
                
            Case BM_Erase
                'Erase mode is totally pointless if alpha is locked, so just skip it entirely
                If (bottomAlphaMode <> AM_Locked) Then
                    ApplyBlendMode_Erase m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom, alphaModifier
                End If
            
            Case BM_Behind
                '"Behind" mode is handled by the final blend step (see below)
                
            Case BM_Overwrite
                'Note that overwrite mode is *NOT* always available.  Like other photo editors,
                ' it exists primarily for adjustment and effect tools.  It does not work on paint tools,
                ' because a separate mask would be required (for the brush outline).
                '
                'As such, this function may not even be accessible from some calling paths, by design.
                ApplyBlendMode_Overwrite m_topDIBCopy, bottomDIB, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom, alphaModifier
            
            'Some blendmodes don't have dedicated, optimized subs just yet.  This catch-all function handles their code for now.
            Case Else
                PDDebug.LogAction "WARNING: UNKNOWN BLEND MODE REQUESTED IN pdPixelBlender.BlendDIBs!"
                
        End Select
        
        'Composited results will have been placed inside the temporary top DIB copy.  Apply the final blend now!
        If applyFinalBlend Then
            
            'In "behind" blend mode, this step is different.
            If (blendMode = BM_Behind) Then
                
                'In "Behind" blend mode, we actually blend the bottom DIB "atop" the top DIB.  Instead of making a
                ' temporary copy of the top DIB, we actually need to make a copy of the *bottom* one.
                If (m_topDIBCopy.GetDIBWidth <> bottomDIB.GetDIBWidth) Or (m_topDIBCopy.GetDIBHeight <> bottomDIB.GetDIBHeight) Then
                    m_topDIBCopy.CreateBlank bottomDIB.GetDIBWidth, bottomDIB.GetDIBHeight, 32, 0, 0
                End If
                
                GDI.BitBltWrapper m_topDIBCopy.GetDIBDC, 0, 0, bottomDIB.GetDIBWidth, bottomDIB.GetDIBHeight, bottomDIB.GetDIBDC, 0, 0, vbSrcCopy
                m_topDIBCopy.SetInitialAlphaPremultiplicationState bottomDIB.GetAlphaPremultiplication
                
                'm_topDIBCopy now contains a copy of the *bottom* DIB's data.  Next we want to "cut out" the
                ' merge region in the *bottom* DIB, and replace it with the top DIB's pixels.  (We also need
                ' to blank out the target region first, so that the layer's opacity - if any - is reflected
                ' correctly in the blend.)
                Dim tmpFillRectF As RectF
                With tmpFillRectF
                    .Left = xOffsetBottom + initX
                    .Top = yOffsetBottom + initY
                    .Width = finalX - initX + 1
                    .Height = finalY - initY + 1
                End With
                
                bottomDIB.FillRectWithColor tmpFillRectF, 0, 0
                topDIB.AlphaBlendToDCEx bottomDIB.GetDIBDC, xOffsetBottom + initX, yOffsetBottom + initY, finalX - initX + 1, finalY - initY + 1, initX, initY, finalX - initX + 1, finalY - initY + 1, Int(alphaModifier * 255! + 0.5!)
                
                'Now, blend the *copy* of the bottom layer onto the space where we just "chopped out"
                ' the bottom layer's pixels and replaced them with the top layer's
                m_topDIBCopy.AlphaBlendToDCEx bottomDIB.GetDIBDC, xOffsetBottom + initX, yOffsetBottom + initY, finalX - initX + 1, finalY - initY + 1, xOffsetBottom + initX, yOffsetBottom + initY, finalX - initX + 1, finalY - initY + 1, 255!
            
            'All other blend modes behave normally
            Else
            
                'Some blend modes ignore this step (Erase and Replace, at present)
                If (blendMode <> BM_Erase) And (blendMode <> BM_Overwrite) Then
                    m_topDIBCopy.AlphaBlendToDCEx bottomDIB.GetDIBDC, xOffsetBottom + initX, yOffsetBottom + initY, finalX - initX + 1, finalY - initY + 1, initX, initY, finalX - initX + 1, finalY - initY + 1, Int(alphaModifier * 255! + 0.5!)
                End If
                'Debug.Print "Blending this many pixels: " & CStr((finalX - initX + 1) * (finalY - initY + 1))
            
            End If
            
            'Also, locked alpha requires us to restore the original bottom layer alpha values now
            If (bottomAlphaMode = AM_Locked) Then PostProcessAlphaEffects m_topDIBCopy, bottomDIB, bottomAlphaMode, topAlphaMode, initX, initY, finalX, finalY, xOffsetBottom, yOffsetBottom, False, ptrToTopLayerAlternateRectF
            
        End If
        
        'That's all there is to it!

    End If

End Sub

'Prior to blending two DIBs, their overlap region must be calculated.  (Only the intersection of the two DIBs needs to be composited;
' everything else can be ignored, saving us valuable time.)
'
'Besides top and bottom DIB references, the destination (x, y) of the top DIB (relative to the bottom DIB) must be supplied.
'
'Returns: TRUE and filled destination coordinates if overlap occurs; FALSE otherwise.
Private Function CalculateDIBOverlap(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal dstX As Single, ByVal dstY As Single, ByRef initX As Long, ByRef initY As Long, ByRef finalX As Long, ByRef finalY As Long, ByRef xOffsetBottom As Long, ByRef yOffsetBottom As Long, Optional ByVal ptrToTopRectF As Long = 0) As Boolean
    
    'Start by calculating bounds for the main per-pixel compositor loop.  These bounds are relative to the *top layer*, as we do
    ' not perform compositing on areas of the bottom layer that *are not overlapped* by the top layer.  (We just treat them as if
    ' the top layer is transparent in those regions, e.g. we ignore them.)
    
    'Start by finding the intersection between the top DIB rect and the bottom DIB rect.  To ensure edges of subpixel positioning aren't missed,
    ' we use RECTFs for the check.
    Dim topRect As RectF, bottomRect As RectF, IntersectRect As RectF
    
    If (ptrToTopRectF = 0) Then
        With topRect
            .Left = 0
            .Top = 0
            .Width = topDIB.GetDIBWidth
            .Height = topDIB.GetDIBHeight
        End With
    Else
        CopyMemoryStrict VarPtr(topRect), ptrToTopRectF, LenB(topRect)
    End If
    
    With bottomRect
        If (ptrToTopRectF = 0) Then
            .Left = -dstX
            .Top = -dstY
        Else
            .Left = topRect.Left - dstX
            .Top = topRect.Top - dstY
        End If
        .Width = bottomDIB.GetDIBWidth
        .Height = bottomDIB.GetDIBHeight
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
        If (finalX > (topDIB.GetDIBWidth - 1)) Then finalX = topDIB.GetDIBWidth - 1
        If (finalY > (topDIB.GetDIBHeight - 1)) Then finalY = topDIB.GetDIBHeight - 1
        
        'Individual blend functions need to know where the offset into the bottom DIB is, so calculate those last.
        xOffsetBottom = Fix(dstX - topRect.Left)
        yOffsetBottom = Fix(dstY - topRect.Top)
        
        'As a final failsafe, make sure the loop bounds never travel outside the *bottom* DIB's boundaries.
        If (xOffsetBottom + finalX) > (bottomDIB.GetDIBWidth - 1) Then finalX = (bottomDIB.GetDIBWidth - 1) - xOffsetBottom
        If (yOffsetBottom + finalY) > (bottomDIB.GetDIBHeight - 1) Then finalY = (bottomDIB.GetDIBHeight - 1) - yOffsetBottom
        
    End If
    
    CalculateDIBOverlap = dibsOverlap
    
End Function

'Prior to applying a pixel blend, you can use this function to pre-mask the top DIB in a blend.  At present, PD uses this
' to apply a selection mask (if any) to a paint stroke prior to merging it onto its base layer.  Because of this limitation,
' this function currently assumes that the two layers are guaranteed to be the same size.
'
'An optional RectF pointer can be passed; this defines the area of the target DIB that contains useful information.
' It is assumed that any pixels *not* inside this rect have already been forced to RGBA(0, 0, 0, 0).
'
'dstDIB itself is used to store the result of the masking, so create a copy prior to masking if you need a backup.
Friend Sub ApplyMaskToTopDIB(ByRef dstDIB As pdDIB, ByRef maskDIB As pdDIB, Optional ByVal ptrToPaintbrushRectF As Long = 0)

    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    
    'If no rect is provided, process the entire image
    If (ptrToPaintbrushRectF = 0) Then
        initX = 0
        finalX = dstDIB.GetDIBStride - 1
        initY = 0
        finalY = dstDIB.GetDIBHeight - 1
    
    'If a rect *is* provided, constrain our loop boundaries accordingly
    Else
        Dim tmpRectF As RectF
        CopyMemoryStrict VarPtr(tmpRectF), ptrToPaintbrushRectF, LenB(tmpRectF)
        initX = Int(tmpRectF.Left) * 4
        initY = Int(tmpRectF.Top)
        finalX = Int(tmpRectF.Left + tmpRectF.Width + 0.9999) - 1
        finalX = finalX * 4
        If (finalX > dstDIB.GetDIBStride - 1) Then finalX = dstDIB.GetDIBStride - 1
        finalY = Int(tmpRectF.Top + tmpRectF.Height + 0.9999) - 1
        If (finalY > dstDIB.GetDIBHeight - 1) Then finalY = dstDIB.GetDIBHeight - 1
    End If
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access),
    ' so we cheat and use 1D arrays, which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = dstDIB.GetDIBStride:  bottomDIBPointer = dstDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = maskDIB.GetDIBStride:  topDIBPointer = maskDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    dstDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    maskDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long
    
    'New and old RGB values
    Dim topA As Single, topAInt As Byte
    Dim bottomAInt As Byte
    
    For y = initY To finalY
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * y
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        'Retrieve mask layer value.  (We don't care about mask layer colors at present.)
        topAInt = topPixels(x)
        
        'Ignore opaque mask pixels
        If (topAInt <> 255) Then
            
            'Retrieve target layer alpha
            bottomAInt = bottomPixels(x + 3)
            
            'If the target pixel is transparent, ignore it
            If (bottomAInt <> 0) Then
                
                'The destination pixel is *not* transparent, meaning we must process it
                topA = intToFloat(topAInt)
                
                'Pixels are already premultiplied, so we can simply multiply by the mask value!
                bottomPixels(x) = bottomPixels(x) * topA
                bottomPixels(x + 1) = bottomPixels(x + 1) * topA
                bottomPixels(x + 2) = bottomPixels(x + 2) * topA
                bottomPixels(x + 3) = bottomAInt * topA
                
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all temporary array references
    dstDIB.UnwrapArrayFromDIB bottomPixels
    maskDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    dstDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Prior to blending two DIBs, call this function preprocess any alpha-specific effects (masks, inheritance, etc).
' This greatly simplifies the actual blending code inside each blend function, because special alpha handling
' is never required.
Private Sub PreProcessAlphaEffects(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal bottomAlphaMode As PD_AlphaMode, ByVal topAlphaMode As PD_AlphaMode, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long, Optional ByVal returnResultsUnpremultiplied As Boolean = False, Optional ByVal ptrToTopLayerAlternateRectF As Long = 0, Optional ByRef topLayerMask As pdDIB = Nothing)
    
    'Only some settings require us to use this function.
    Dim needToStay As Boolean
    needToStay = (bottomAlphaMode <> AM_Normal) Or (topAlphaMode <> AM_Normal) Or (Not topLayerMask Is Nothing)
    
    If needToStay Then
        
        Dim x As Long, y As Long, bottomX As Long, bottomY As Long
        Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
        Dim bottomA As Single, bottomAInt As Byte
        
        Dim bottomScanlineSize As Long, bottomDIBPointer As Long
        Dim bottomPixels() As Byte, bottomSA As SafeArray1D
        Dim topPixels() As Byte, topSA As SafeArray1D
        Dim topScanlineSize As Long, topDIBPointer As Long
        
        'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
        ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
        Dim intToFloat(0 To 255) As Single
        Dim i As Long
        For i = 0 To 255
            intToFloat(i) = CSng(i) * ONE_DIV_255
        Next i
        
        'Before going any further, the first alpha task we must handle is masking.  (Masking takes priority over any other
        ' alpha settings.)
        If (Not topLayerMask Is Nothing) Then
            
            Dim tmpInitX As Long, tmpInitY As Long, tmpFinalX As Long, tmpFinalY As Long
            
            'Perform failsafe checks against the mask boundaries
            tmpInitX = initX * 4
            tmpInitY = initY
            tmpFinalX = finalX * 4
            tmpFinalY = finalY
            If (tmpFinalX > topLayerMask.GetDIBStride - 1) Then tmpFinalX = topLayerMask.GetDIBStride - 1
            If (tmpFinalY > topLayerMask.GetDIBHeight - 1) Then tmpFinalY = topLayerMask.GetDIBHeight - 1
            
            '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access),
            ' so we cheat and use 1D arrays, which we reset between scanlines.
            bottomScanlineSize = topDIB.GetDIBStride:  bottomDIBPointer = topDIB.GetDIBPointer
            topScanlineSize = topLayerMask.GetDIBStride:  topDIBPointer = topLayerMask.GetDIBPointer
            
            topDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
            topLayerMask.WrapArrayAroundDIB_1D topPixels, topSA
            
            For y = tmpInitY To tmpFinalY
                
                'Point each 1D pixel array at the proper scanline
                bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * y
                topSA.pvData = topDIBPointer + topScanlineSize * y
                
            For x = tmpInitX To tmpFinalX Step 4
                
                'Retrieve mask layer alpha.  (We don't care about mask layer colors at present.)
                'NOTE: this was changed to check color on 09 August 2017.  Exterior and bordered selections
                ' were not working correctly, and this fixes it; I don't know if it has future repercussions.
                topAInt = topPixels(x)  'topPixels(x + 3)
                
                'Ignore opaque mask pixels
                If (topAInt <> 255) Then
                    
                    'Retrieve target layer alpha
                    bottomAInt = bottomPixels(x + 3)
                    
                    'If the target pixel is transparent, ignore it
                    If (bottomAInt <> 0) Then
                        
                        'The destination pixel is *not* transparent, meaning we must process it
                        topA = intToFloat(topAInt)
                        
                        'Pixels are already premultiplied, so we can simply multiply by the mask value!
                        bottomPixels(x) = bottomPixels(x) * topA
                        bottomPixels(x + 1) = bottomPixels(x + 1) * topA
                        bottomPixels(x + 2) = bottomPixels(x + 2) * topA
                        bottomPixels(x + 3) = bottomAInt * topA
                        
                    End If
                    
                End If
            
            Next x
            Next y
            
            'Clear all temporary array references
            topDIB.UnwrapArrayFromDIB bottomPixels
            topLayerMask.UnwrapArrayFromDIB topPixels
            
            'All blend mode functions return premultiplied DIBs, by design.
            topDIB.SetInitialAlphaPremultiplicationState True
        
        End If
        
        'Next comes locked alpha.  If alpha is locked for the destination, we don't care about further processing;
        ' instead, we cache alpha values and bail immediately.
        If (bottomAlphaMode = AM_Locked) Then
            
            'Locked alpha is a little weird, because we're copying out the alpha values of the *bottom* layer.  Because the rect
            ' we've been passed is relative to the *top* layer, we need to calculate a transform now.
            If (ptrToTopLayerAlternateRectF <> 0) Then
                Dim tmpRectF As RectF, tmpRectFSrc As RectF
                CopyMemoryStrict VarPtr(tmpRectFSrc), ptrToTopLayerAlternateRectF, LenB(tmpRectF)
                With tmpRectF
                    .Left = tmpRectFSrc.Left + xOffset
                    .Top = tmpRectFSrc.Top + yOffset
                    .Width = tmpRectFSrc.Width
                    .Height = tmpRectFSrc.Height
                End With
                ptrToTopLayerAlternateRectF = VarPtr(tmpRectF)
            End If
            
            DIBs.RetrieveTransparencyTable bottomDIB, m_AlphaCopy, ptrToTopLayerAlternateRectF
            Exit Sub
            
        End If
    
        If (topAlphaMode = AM_Inherit) Then
        
            'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
            ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
            initX = initX * 4
            finalX = finalX * 4
            xOffset = xOffset * 4
            
            '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
            ' which we reset between scanlines.
            bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
            topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
            
            bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
            topDIB.WrapArrayAroundDIB_1D topPixels, topSA
            
            For y = initY To finalY
                
                'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
                ' into the bottom DIB's coordinate space.
                bottomY = yOffset + y
                
                'Point each 1D pixel array at the proper scanline
                bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
                topSA.pvData = topDIBPointer + topScanlineSize * y
                
            For x = initX To finalX Step 4
                
                bottomX = (xOffset + x)
                
                'Retrieve top layer RGBA values
                topAInt = topPixels(x + 3)
                
                'Ignore transparent pixels completely; they go untouched.
                If (topAInt <> 0) Then
                    
                    'Retrieve bottom layer alpha
                    bottomAInt = bottomPixels(bottomX + 3)
                    
                    'Alpha inheritance can only make pixels *more* transparent, not *less* transparent.  This means we can ignore opaque pixels,
                    ' as they have no effect on the top layer.
                    If (bottomAInt < 255) Then
                    
                        'If the bottom pixel is transparent, we can simply erase the top pixel in this location
                        If (bottomAInt = 0) Then
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
            bottomDIB.UnwrapArrayFromDIB bottomPixels
            topDIB.UnwrapArrayFromDIB topPixels
            
            'All blend mode functions return premultiplied DIBs, by design.
            topDIB.SetInitialAlphaPremultiplicationState True
        
        End If
        
    End If
    
End Sub

'After blending two DIBs, call this function to postprocess any alpha-specific effects (masks, inheritance, etc).  Note that
' not all alpha modes require this, but there is no harm in calling it regardless.
Private Sub PostProcessAlphaEffects(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal bottomAlphaMode As PD_AlphaMode, ByVal topAlphaMode As PD_AlphaMode, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long, Optional ByVal returnResultsUnpremultiplied As Boolean = False, Optional ByVal ptrToTopLayerAlternateRectF As Long = 0)
    
    If (bottomAlphaMode = AM_Locked) Then
        
        'Locked alpha is a little weird, because we're copying in the alpha values of the *bottom* layer.  Because the rect
        ' we've been passed is relative to the *top* layer, we need to calculate a transform now.
        If (ptrToTopLayerAlternateRectF <> 0) Then
            Dim tmpRectF As RectF, tmpRectFSrc As RectF
            CopyMemoryStrict VarPtr(tmpRectFSrc), ptrToTopLayerAlternateRectF, LenB(tmpRectF)
            With tmpRectF
                .Left = tmpRectFSrc.Left + xOffset
                .Top = tmpRectFSrc.Top + yOffset
                .Width = tmpRectFSrc.Width
                .Height = tmpRectFSrc.Height
            End With
            ptrToTopLayerAlternateRectF = VarPtr(tmpRectF)
        End If
            
        DIBs.ApplyTransparencyTable bottomDIB, m_AlphaCopy, ptrToTopLayerAlternateRectF
    End If
    
End Sub

'Optimized "color" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Color(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim topH As Single, topS As Single, topV As Single, bottomV As Single
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend step is very simple
                FastRGBtoHSV topR, topG, topB, topH, topS, topV
                'bottomV = fMax3(bottomR, bottomG, bottomB)
                bottomV = 0.3! * bottomR + 0.59! * bottomG + 0.11! * bottomB
                fHSVtoRGB topH, topS, bottomV, newR, newG, newB
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255!
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "color burn" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_ColorBurn(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                If (topR <> 0!) Then
                    newR = 1! - (1! - bottomR) / topR
                    If (newR < 0!) Then newR = 0!
                Else
                    newR = 0!
                End If
                
                If (topG <> 0!) Then
                    newG = 1! - (1! - bottomG) / topG
                    If (newG < 0!) Then newG = 0!
                Else
                    newG = 0!
                End If
                    
                If (topB <> 0!) Then
                    newB = 1! - (1! - bottomB) / topB
                    If (newB < 0!) Then newB = 0!
                Else
                    newB = 0!
                End If
                
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255!
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "color dodge" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_ColorDodge(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                If (topR <> 1!) Then
                    newR = bottomR / (1! - topR)
                    If (newR > 1!) Then newR = 1!
                Else
                    newR = 1!
                End If
                
                If (topG <> 1!) Then
                    newG = bottomG / (1! - topG)
                    If (newG > 1!) Then newG = 1!
                Else
                    newG = 1!
                End If
                
                If (topB <> 1!) Then
                    newB = bottomB / (1! - topB)
                    If (newB > 1!) Then newB = 1!
                Else
                    newB = 1!
                End If
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "darken" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Darken(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Darken" blend mode is very simple
                If (bottomR < topR) Then newR = bottomR Else newR = topR
                If (bottomG < topG) Then newG = bottomG Else newG = topG
                If (bottomB < topB) Then newB = bottomB Else newB = topB
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "difference" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results
' of the blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Difference(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Difference" blend mode is very simple
                newR = (bottomR - topR)
                If (newR < 0!) Then newR = -newR
                newG = (bottomG - topG)
                If (newG < 0!) Then newG = -newG
                newB = (bottomB - topB)
                If (newB < 0!) Then newB = -newB
                
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "divide" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results
' of the blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Divide(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                If (topR <> 0!) Then
                    newR = bottomR / topR
                    If (newR > 1!) Then newR = 1!
                Else
                    newR = 1!
                End If
                If (topG <> 0!) Then
                    newG = bottomG / topG
                    If (newG > 1!) Then newG = 1!
                Else
                    newG = 1!
                End If
                If (topB <> 0!) Then
                    newB = bottomB / topB
                    If (newB > 1!) Then newB = 1!
                Else
                    newB = 1!
                End If
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "erase" blend mode function.  Unlike other blend modes, this blend mode *does* modify the bottom DIB (as required by
' erase functionality).  *Do not* apply a separate alpha blend after using this function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Erase(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long, Optional ByVal alphaModifier As Single = 1!)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve the top layer's alpha value
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels
        If (topAInt <> 0) Then
        
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'If the bottom pixel is already transparent, we can ignore it
            If (bottomAInt <> 0) Then
                
                'Convert the top-layer alpha to its floating-point equivalent
                topA = intToFloat(255 - (topAInt * alphaModifier))
                
                'Retrieve the bottom-layer color and alpha values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply the bottom-layer RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'Paste the modified alpha values directly into the bottom layer
                bottomA = (topA * bottomA) * 255
                bottomPixels(bottomX) = bottomB * bottomA: bottomPixels(bottomX + 1) = bottomG * bottomA: bottomPixels(bottomX + 2) = bottomR * bottomA
                bottomPixels(bottomX + 3) = bottomA
                    
            End If
            
        End If
        
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "exclusion" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results
' of the blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Exclusion(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                newR = bottomR + topR - 2 * (bottomR * topR)
                newG = bottomG + topG - 2 * (bottomG * topG)
                newB = bottomB + topB - 2 * (bottomB * topB)
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "grain extract" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_GrainExtract(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                newR = bottomR - topR + 0.5
                newG = bottomG - topG + 0.5
                newB = bottomB - topB + 0.5
                If (newR > 1!) Then
                    newR = 1!
                ElseIf (newR < 0!) Then
                    newR = 0!
                End If
                If (newG > 1!) Then
                    newG = 1!
                ElseIf (newG < 0!) Then
                    newG = 0!
                End If
                If (newB > 1!) Then
                    newB = 1!
                ElseIf (newB < 0!) Then
                    newB = 0!
                End If
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "grain merge" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_GrainMerge(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                newR = bottomR + topR - 0.5
                newG = bottomG + topG - 0.5
                newB = bottomB + topB - 0.5
                If (newR > 1!) Then
                    newR = 1!
                ElseIf (newR < 0!) Then
                    newR = 0!
                End If
                If (newG > 1!) Then
                    newG = 1!
                ElseIf (newG < 0!) Then
                    newG = 0!
                End If
                If (newB > 1!) Then
                    newB = 1!
                ElseIf (newB < 0!) Then
                    newB = 0!
                End If
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "hard light" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_HardLight(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Hard light" blend mode is very simple
                If (topR < 0.5) Then newR = 2 * bottomR * topR Else newR = 1! - 2 * (1! - bottomR) * (1! - topR)
                If (topG < 0.5) Then newG = 2 * bottomG * topG Else newG = 1! - 2 * (1! - bottomG) * (1! - topG)
                If (topB < 0.5) Then newB = 2 * bottomB * topB Else newB = 1! - 2 * (1! - bottomB) * (1! - topB)
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "hard mix" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_HardMix(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                If (topR < 1 - bottomR) Then newR = 0! Else newR = 1!
                If (topG < 1 - bottomG) Then newG = 0! Else newG = 1!
                If (topB < 1 - bottomB) Then newB = 0! Else newB = 1!
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "hue" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Hue(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim topH As Single, bottomH As Single, bottomS As Single, bottomV As Single
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend step is very simple
                FastRGBtoHOnly topR, topG, topB, topH
                FastRGBtoHSV bottomR, bottomG, bottomB, bottomH, bottomS, bottomV
                fHSVtoRGB topH, bottomS, bottomV, newR, newG, newB
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "lighten" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Lighten(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Lighten" blend mode is very simple
                If bottomR > topR Then newR = bottomR Else newR = topR
                If bottomG > topG Then newG = bottomG Else newG = topG
                If bottomB > topB Then newB = bottomB Else newB = topB
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "linear burn" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_LinearBurn(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                newR = topR + bottomR - 1
                If (newR < 0!) Then newR = 0!
                newG = topG + bottomG - 1
                If (newG < 0!) Then newG = 0!
                newB = topB + bottomB - 1
                If (newB < 0!) Then newB = 0!
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "linear dodge" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_LinearDodge(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Linear dodge" blend mode is very simple
                newR = topR + bottomR
                newG = topG + bottomG
                newB = topB + bottomB
                If newR > 1 Then newR = 1!
                If newG > 1 Then newG = 1!
                If newB > 1 Then newB = 1!
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "linear light" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_LinearLight(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                newR = bottomR + (2 * topR) - 1
                If (newR > 1!) Then
                    newR = 1!
                ElseIf (newR < 0!) Then
                    newR = 0!
                End If
                
                newG = bottomG + (2 * topG) - 1
                If (newG > 1!) Then
                    newG = 1!
                ElseIf (newG < 0!) Then
                    newG = 0!
                End If
                
                newB = bottomB + (2 * topB) - 1
                If (newB > 1!) Then
                    newB = 1!
                ElseIf (newB < 0!) Then
                    newB = 0!
                End If
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub


'Optimized "luminosity" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Luminosity(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim topV As Single, bottomH As Single, bottomS As Single, bottomV As Single
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend step is very simple
                'topV = fMax3(topR, topG, topB)
                topV = 0.3 * topR + 0.59 * topB + 0.11 * topG
                FastRGBtoHSV bottomR, bottomG, bottomB, bottomH, bottomS, bottomV
                fHSVtoRGB bottomH, bottomS, topV, newR, newG, newB
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "multiply" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Multiply(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Multiply" blend mode is very simple
                newR = topR * bottomR
                newG = topG * bottomG
                newB = topB * bottomB
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "overlay" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Overlay(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Overlay" blend mode is very simple
                If bottomR < 0.5 Then newR = 2 * bottomR * topR Else newR = 1! - 2 * (1! - bottomR) * (1! - topR)
                If bottomG < 0.5 Then newG = 2 * bottomG * topG Else newG = 1! - 2 * (1! - bottomG) * (1! - topG)
                If bottomB < 0.5 Then newB = 2 * bottomB * topB Else newB = 1! - 2 * (1! - bottomB) * (1! - topB)
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "overwrite" or "replace" blend mode function.
' - Unlike other blend modes, *the bottom DIB is changed by this function*!  Do NOT apply a separate alpha-blend
'   after this function returns.
' - Unlike other blend modes, *the top DIB is returned unchanged*!
' - Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Overwrite(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long, Optional ByVal alphaModifier As Single = 1!)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    Dim invAlphaModifier As Single
    invAlphaModifier = 1! - alphaModifier
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
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
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
        'If alpha isn't changing, this is a simple copy op, replacing the old pixels with the new ones
        If (alphaModifier = 1!) Then
            VBHacks.CopyMemoryStrict VarPtr(bottomPixels(xOffset)), VarPtr(topPixels(0)), finalX - initX
        Else
                
            For x = initX To finalX Step 4
                
                bottomX = (xOffset + x)
                
                'Simple as can be, replace the bottom DIB values with the top DIB values.
                ' (Note that this manual path is no longer needed; see the above branch for details!)
                If (alphaModifier = 1!) Then
                    bottomPixels(bottomX) = topPixels(x)
                    bottomPixels(bottomX + 1) = topPixels(x + 1)
                    bottomPixels(bottomX + 2) = topPixels(x + 2)
                    bottomPixels(bottomX + 3) = topPixels(x + 3)
                
                'At lower opacities, "fade" the replacement with the original pixels
                Else
                    
                    bottomB = intToFloat(bottomPixels(bottomX))
                    bottomG = intToFloat(bottomPixels(bottomX + 1))
                    bottomR = intToFloat(bottomPixels(bottomX + 2))
                    bottomA = intToFloat(bottomPixels(bottomX + 3))
                    
                    topB = intToFloat(topPixels(x))
                    topG = intToFloat(topPixels(x + 1))
                    topR = intToFloat(topPixels(x + 2))
                    topA = intToFloat(topPixels(x + 3))
                    
                    newB = topB * alphaModifier + bottomB * invAlphaModifier
                    newG = topG * alphaModifier + bottomG * invAlphaModifier
                    newR = topR * alphaModifier + bottomR * invAlphaModifier
                    newA = topA * alphaModifier + bottomA * invAlphaModifier
                    
                    bottomPixels(x) = Int(newB * 255#)
                    bottomPixels(x + 1) = Int(newG * 255#)
                    bottomPixels(x + 2) = Int(newR * 255#)
                    bottomPixels(x + 3) = Int(newA * 255#)
                    
                End If
                
            Next x
            
        End If
        
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "pin light" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_PinLight(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                If (topR > 0.5) Then newR = fMax(bottomR, 2 * (topR - 0.5)) Else newR = fMin(bottomR, 2 * topR)
                If (topG > 0.5) Then newG = fMax(bottomG, 2 * (topG - 0.5)) Else newG = fMin(bottomG, 2 * topG)
                If (topB > 0.5) Then newB = fMax(bottomB, 2 * (topB - 0.5)) Else newB = fMin(bottomB, 2 * topB)
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "saturation" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blend mode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Saturation(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim topS As Single, bottomH As Single, bottomS As Single, bottomV As Single
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend step is very simple
                FastRGBtoSOnly topR, topG, topB, topS
                FastRGBtoHSV bottomR, bottomG, bottomB, bottomH, bottomS, bottomV
                fHSVtoRGB bottomH, topS, bottomV, newR, newG, newB
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "screen" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Screen(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Screen" blend mode is very simple.  A full expansion looks like this:
                'newR = 1.0 - (1.0 - bottomR) * (1.0 - topR)
                'newG = 1.0 - (1.0 - bottomG) * (1.0 - topG)
                'newB = 1.0 - (1.0 - bottomB) * (1.0 - topB)
                
                '...but algebraically, we can reduce it to this:
                newR = bottomR + topR - (bottomR * topR)
                newG = bottomG + topG - (bottomG * topG)
                newB = bottomB + topB - (bottomB * topB)
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255!
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "soft light" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_SoftLight(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                '"Soft light" blend mode is very simple
                newR = (1# - 2# * topR) * bottomR * bottomR + 2# * bottomR * topR
                newG = (1# - 2# * topG) * bottomG * bottomG + 2# * bottomG * topG
                newB = (1# - 2# * topB) * bottomB * bottomB + 2# * bottomB * topB
                
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "subtract" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_Subtract(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                newR = bottomR - topR
                If (newR < 0!) Then newR = 0!
                newG = bottomG - topG
                If (newG < 0!) Then newG = 0!
                newB = bottomB - topB
                If (newB < 0!) Then newB = 0!
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'Optimized "vivid light" blend mode function.  The bottom DIB will be returned unchanged.  The top DIB will contain the results of the
' blendmode calculations, which should then be alpha-blended onto the target via some other function.
' Note that 32-bpp DIBs are required.  Other color depths (including 24-bpp) will crash.
Private Sub ApplyBlendMode_VividLight(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByVal xOffset As Long, ByVal yOffset As Long)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    initX = initX * 4
    finalX = finalX * 4
    xOffset = xOffset * 4
    
    'Blending requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table for converting
    ' single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        intToFloat(i) = CSng(i) * ONE_DIV_255
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use 1D arrays,
    ' which we reset between scanlines.
    Dim bottomScanlineSize As Long, bottomDIBPointer As Long
    bottomScanlineSize = bottomDIB.GetDIBStride:  bottomDIBPointer = bottomDIB.GetDIBPointer
    
    Dim topScanlineSize As Long, topDIBPointer As Long
    topScanlineSize = topDIB.GetDIBStride:  topDIBPointer = topDIB.GetDIBPointer
    
    Dim bottomPixels() As Byte, bottomSA As SafeArray1D, topPixels() As Byte, topSA As SafeArray1D
    bottomDIB.WrapArrayAroundDIB_1D bottomPixels, bottomSA
    topDIB.WrapArrayAroundDIB_1D topPixels, topSA
    
    'Loop through the relevant portion of the top layer, compositing its pixels onto the bottom layer as we go.
    Dim x As Long, y As Long, bottomX As Long, bottomY As Long
    
    'New and old RGB values
    Dim topR As Single, topG As Single, topB As Single, topA As Single, topAInt As Byte
    Dim bottomR As Single, bottomG As Single, bottomB As Single, bottomA As Single, bottomAInt As Byte
    Dim newR As Single, newG As Single, newB As Single
    
    For y = initY To finalY
        
        'These core (x, y) parameters reference the top image's coordinate space.  We also need to calculate corresponding offsets
        ' into the bottom DIB's coordinate space.
        bottomY = yOffset + y
        
        'Point each 1D pixel array at the proper scanline
        bottomSA.pvData = bottomDIBPointer + bottomScanlineSize * bottomY
        topSA.pvData = topDIBPointer + topScanlineSize * y
        
    For x = initX To finalX Step 4
        
        bottomX = (xOffset + x)
        
        'Retrieve top layer RGBA values
        topAInt = topPixels(x + 3)
        
        'Ignore transparent pixels completely
        If (topAInt <> 0) Then
            
            'Retrieve bottom layer alpha
            bottomAInt = bottomPixels(bottomX + 3)
            
            'We only need to process pixels where the bottom is *not* transparent.  (Transparent pixels are handled
            ' automatically by the final alpha-blend in our parent function.)
            If (bottomAInt <> 0) Then
                
                'Retrieve the top-layer values and convert them to floating-point
                topA = intToFloat(topAInt)
                topB = intToFloat(topPixels(x)): topG = intToFloat(topPixels(x + 1)): topR = intToFloat(topPixels(x + 2))
                
                'Do the same for the bottom layer values
                bottomA = intToFloat(bottomAInt)
                bottomB = intToFloat(bottomPixels(bottomX)): bottomG = intToFloat(bottomPixels(bottomX + 1)): bottomR = intToFloat(bottomPixels(bottomX + 2))
                
                'Un-premultiply both top and bottom RGB values.
                ' (Note that these divisions are safe; alpha values cannot physically be zero, as we've already checked for that case.)
                If (topA <> 1!) Then
                    topR = topR / topA: topG = topG / topA: topB = topB / topA
                End If
                
                If (bottomA <> 1!) Then
                    bottomR = bottomR / bottomA: bottomG = bottomG / bottomA: bottomB = bottomB / bottomA
                End If
                
                'The actual blend is very simple
                If (topR > 0.5) Then newR = bottomR + 2 * (topR - 0.5) Else newR = bottomR + 2 * topR - 1
                If (newR > 1!) Then
                    newR = 1!
                ElseIf (newR < 0!) Then
                    newR = 0!
                End If
                
                If (topG > 0.5) Then newG = bottomG + 2 * (topG - 0.5) Else newG = bottomG + 2 * topG - 1
                If (newG > 1!) Then
                    newG = 1!
                ElseIf (newG < 0!) Then
                    newG = 0!
                End If
                
                If (topB > 0.5) Then newB = bottomB + 2 * (topB - 0.5) Else newB = bottomB + 2 * topB - 1
                If (newB > 1!) Then
                    newB = 1!
                ElseIf (newB < 0!) Then
                    newB = 0!
                End If
                    
                'If the bottom layer contains transparency, mix the newly calculated RGB values against the original top layer
                ' RGB values.  This reduces the strength of the blend mode result, proportional to the bottom layer's alpha.
                If (bottomA <> 1!) Then
                    newR = topR + bottomA * (newR - topR)
                    newG = topG + bottomA * (newG - topG)
                    newB = topB + bottomA * (newB - topB)
                End If
                
                'Copy the final RGB values into the top layer, and premultiply them.  Note that top layer's alpha remains unchanged!
                ' (A final, blended alpha value will be handled by a subsequent function.)
                topA = topA * 255#
                topPixels(x) = newB * topA: topPixels(x + 1) = newG * topA: topPixels(x + 2) = newR * topA
                    
            End If
            
        End If
    
    Next x
    Next y
    
    'Clear all unsafe array references
    bottomDIB.UnwrapArrayFromDIB bottomPixels
    topDIB.UnwrapArrayFromDIB topPixels
    
    'All blend mode functions return premultiplied DIBs, by design.
    topDIB.SetInitialAlphaPremultiplicationState True
    
End Sub

'A heavily modified RGB to HSV transform, courtesy of http://lolengine.net/blog/2013/01/13/fast-rgb-to-hsv.
' Note that the code assumes RGB values already in the [0, 1] range, and it will return HSV values in the [0, 1] range.
Private Sub FastRGBtoHSV(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByRef h As Single, ByRef s As Single, ByRef v As Single)

    Dim k As Single, tmpSwap As Single, chroma As Single
    
    If (g < b) Then
        tmpSwap = b
        b = g
        g = tmpSwap
        k = -1!
    End If
    
    If (r < g) Then
        tmpSwap = g
        g = r
        r = tmpSwap
        k = -0.333333333333333 - k
    End If
    
    chroma = r - fMin(g, b)
    h = k + (g - b) / (6# * chroma + 0.0000001)
    If (h < 0!) Then h = -h
    s = chroma / (r + 0.00000001)
    v = r
    
End Sub

Private Sub FastRGBtoHOnly(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByRef h As Single)

    Dim k As Single, tmpSwap As Single
    
    If (g < b) Then
        tmpSwap = b
        b = g
        g = tmpSwap
        k = -1!
    End If
    
    If (r < g) Then
        tmpSwap = g
        g = r
        r = tmpSwap
        k = -0.333333333333333 - k
    End If
    
    h = k + (g - b) / (6# * (r - fMin(g, b)) + 0.0000001)
    If (h < 0!) Then h = -h
    
End Sub

Private Sub FastRGBtoSOnly(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByRef s As Single)

    Dim k As Single, tmpSwap As Single
    
    If (g < b) Then
        tmpSwap = b
        b = g
        g = tmpSwap
        k = -1!
    End If
    
    If (r < g) Then
        tmpSwap = g
        g = r
        r = tmpSwap
        k = -0.333333333333333 - k
    End If
    
    s = (r - fMin(g, b)) / (r + 0.00000001)
    
End Sub

'Convert [0,1] HSV values to [0,255] RGB values, with thanks to easyrgb.com for the conversion math
Private Sub fHSVtoRGB(ByRef h As Single, ByRef s As Single, ByRef v As Single, ByRef r As Single, ByRef g As Single, ByRef b As Single)

    'If saturation is 0, RGB are calculated identically
    If (s = 0!) Then
        r = v
        g = v
        b = v
        
    'If saturation is not 0, we have to calculate RGB independently
    Else
       
        Dim var_H As Single
        var_H = h * 6#
        
        'To keep our math simple, limit hue to [0, 5.9999999]
        If (var_H >= 6#) Then var_H = 0!
        
        Dim var_I As Long
        var_I = Int(var_H)
        
        Dim var_1 As Single, var_2 As Single, var_3 As Single
        var_1 = v * (1! - s)
        var_2 = v * (1! - s * (var_H - var_I))
        var_3 = v * (1! - s * (1! - (var_H - var_I)))
        
        If (var_I = 0) Then
            r = v
            g = var_3
            b = var_1
                
        ElseIf (var_I = 1) Then
            r = var_2
            g = v
            b = var_1
                
        ElseIf (var_I = 2) Then
            r = var_1
            g = v
            b = var_3
                
        ElseIf (var_I = 3) Then
            r = var_1
            g = var_2
            b = v
            
        ElseIf (var_I = 4) Then
            r = var_3
            g = var_1
            b = v
                
        Else
            r = v
            g = var_1
            b = var_2
        End If
                
    End If

End Sub

'Return the minimum of two floating-point values
Private Function fMin(x As Single, y As Single) As Single
    If (x > y) Then fMin = y Else fMin = x
End Function

'Return the maximum of two floating-point values
Private Function fMax(x As Single, y As Single) As Single
    If (x < y) Then fMax = y Else fMax = x
End Function

'In order to render complex images quickly, this class must generate a lot of temporary objects.  If PD is under memory pressure, call this function
' to forcibly release as many temporary objects as possible.  (It should be obvious, but subsequent render calls will result in temporary objects
' being recreated, as necessary.)
Friend Sub AttemptToFreeMemory()
    m_topDIBCopy.EraseDIB
End Sub

Private Sub Class_Initialize()
    Set m_topDIBCopy = New pdDIB
End Sub

Private Sub Class_Terminate()
    Set m_topDIBCopy = Nothing
End Sub
