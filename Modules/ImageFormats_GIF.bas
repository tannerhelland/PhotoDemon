Attribute VB_Name = "ImageFormats_GIF"
'***************************************************************************
'Additional support functions for GIF support
'Copyright 2001-2021 by Tanner Helland
'Created: 4/15/01
'Last updated: 14/October/21
'Last update: pull actual AGIF encoding stages into their own function; this will make it easier to
'             switch encoders in the future
'
'Most image exporters exist in the ImageExporter module.  GIF is a weird exception because animated GIFs
' require a ton of preprocessing (to optimize animation frames), so I've moved them to their own home.
'
'PhotoDemon automatically optimizes saved GIFs to produce the smallest possible files.  A variety of
' optimizations are used, and the encoder tests various strategies to try and choose the "best"
' (smallest) solution on each frame.
'
'Note that the optimizer is specifically written in an export-library-agnostic way.  PD internally
' stores the results of all optimizations, then just hands the optimized frames off to an encoder
' at the end of the process.  Currently this encoder is FreeImage.  FreeImage has many quirks and
' produces unnecessarily large files, however, so I am actively investigating alternatives.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The animated GIF exporter builds a collection of frame data during export.
Private Type PD_GifFrame
    usesGlobalPalette As Boolean        'GIFs allow for both global and local palettes.  PD optimizes against both.
    frameIsDuplicateOrEmpty As Boolean  'PD automatically drops duplicate and/or empty frames
    frameNeedsTransparency As Boolean   'PD may require transparency as part of optimizing a given frame
    frameTime As Long                   'GIF frame time is in centiseconds (uuuuuuuugh what a terrible decision)
    frameDisposal As FREE_IMAGE_FRAME_DISPOSAL_METHODS  'GIF and APNG disposal methods are roughly identical
    rectOfInterest As RectF             'Frames are auto-cropped to their relevant regions
    frameDIB As pdDIB                   'Only used temporarily, during optimization; ultimately palettized to produce...
    pixelData() As Byte                 '...this bytestream (and associated palette) instead.
    palNumColors As Long
    framePalette() As RGBQuad
End Type

'Optimized GIF frames will be stored here.  This array is auto-cleared after a successful dump to file.
Private m_allFrames() As PD_GifFrame

'PD always writes a global palette, and it attempts to use it on as many frames as possible.
' (Local palettes will automatically be generated too, as necessary.)
Private m_globalPalette() As RGBQuad, m_numColorsInGP As Long, m_GlobalTrnsIndex As Long

'Low-level GIF export interface
Public Function ExportGIF_LL(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportGIFError
    
    ExportGIF_LL = False
    Dim sFileType As String: sFileType = "GIF"
    
    'Parse all relevant GIF parameters.  (See the GIF export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    'Only two parameters are mandatory; the others are used on an as-needed basis
    Dim gifColorMode As String, gifAlphaMode As String
    gifColorMode = cParams.GetString("gif-color-mode", "auto")
    gifAlphaMode = cParams.GetString("gif-alpha-mode", "auto")
    
    Dim gifAlphaCutoff As Long, gifColorCount As Long, gifBackgroundColor As Long, gifAlphaColor As Long
    gifAlphaCutoff = cParams.GetLong("gif-alpha-cutoff", 64)
    gifColorCount = cParams.GetLong("gif-color-count", 256)
    gifBackgroundColor = cParams.GetLong("gif-backcolor", vbWhite)
    gifAlphaColor = cParams.GetLong("gif-alpha-color", RGB(255, 0, 255))
    
    'Some combinations of parameters invalidate other parameters.  Calculate any overrides now.
    Dim gifForceGrayscale As Boolean
    gifForceGrayscale = Strings.StringsEqual(gifColorMode, "gray", True)
    If Strings.StringsEqual(gifColorMode, "auto", True) Then gifColorCount = 256
    
    Dim desiredAlphaStatus As PD_ALPHA_STATUS
    desiredAlphaStatus = PDAS_BinaryAlpha
    If Strings.StringsEqual(gifAlphaMode, "none", True) Then desiredAlphaStatus = PDAS_NoAlpha
    If Strings.StringsEqual(gifAlphaMode, "by-color", True) Then
        desiredAlphaStatus = PDAS_NewAlphaFromColor
        gifAlphaCutoff = gifAlphaColor
    End If
    
    'Generate a composited image copy, with alpha automatically un-premultiplied
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, False
        
    'FreeImage provides the most comprehensive GIF encoder, so we prefer it whenever possible
    If ImageFormats.IsFreeImageEnabled Then
            
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, 8, desiredAlphaStatus, PDAS_ComplicatedAlpha, gifAlphaCutoff, gifBackgroundColor, gifForceGrayscale, gifColorCount)
        
        'Finally, prepare some GIF save flags.  If the user has requested RLE encoding, and this image is <= 8bpp,
        ' request RLE encoding from FreeImage.
        Dim GIFflags As Long: GIFflags = GIF_DEFAULT
        
        'Use that handle to save the image to GIF format, with required color conversion based on the outgoing color depth
        If (fi_DIB <> 0) Then
            ExportGIF_LL = FreeImage_SaveEx(fi_DIB, dstFile, PDIF_GIF, GIFflags, FICD_8BPP, , , , , True)
            If ExportGIF_LL Then
                ExportDebugMsg "Export to " & sFileType & " appears successful."
            Else
                Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            End If
        Else
            Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
            ExportGIF_LL = False
        End If
    
    'If FreeImage is unavailable, fall back to GDI+
    Else
        ExportGIF_LL = GDIPlusSavePicture(srcPDImage, dstFile, P2_FFE_GIF, 8)
    End If
    
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportGIF_LL = False
    
End Function

'Low-level animated GIF export.  Currently relies on FreeImage for export, but it's designed so that any
' capable encoder can be easily dropped-in.  (Frame optimization happens locally, using PD data structures,
' so the encoder doesn't need to support it at all.)
Public Function ExportGIF_Animated_LL(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportGIFError
    
    ExportGIF_Animated_LL = False
    Dim sFileType As String: sFileType = "GIF"
    
    'Initialize a progress bar
    ProgressBars.SetProgBarMax srcPDImage.GetNumOfLayers
    
    'Parse all relevant GIF parameters.  (See the GIF export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    Dim useFixedFrameDelay As Boolean, frameDelayDefault As Long
    useFixedFrameDelay = cParams.GetBool("use-fixed-frame-delay", False)
    frameDelayDefault = cParams.GetLong("frame-delay-default", 100)
    
    Dim gifAlphaCutoff As Long, gifMatteColor As Long
    gifAlphaCutoff = cParams.GetLong("alpha-cutoff", 64)
    gifMatteColor = cParams.GetLong("matte-color", vbWhite)
    
    Dim autoDither As Boolean, useDithering As Boolean, ditherText As String
    ditherText = cParams.GetString("dither", "auto")
    autoDither = Strings.StringsEqual(ditherText, "auto", True)
    If (Not autoDither) Then useDithering = Strings.StringsEqual(ditherText, "on", True)
    
    'FreeImage is currently required for animated GIF export.  I'd really like to change this
    ' in the future because FreeImage is slow and does not write optimal GIFs (it always writes
    ' 256-color palettes, for example, regardless of actual palette count), but for now it's
    ' a dependency that I'm already shipping and it works all the way back to XP.
    '
    'Anyway, FreeImage isn't actually involved until near the end of the function (search for
    ' the text "Finalizing image" to see where we create our first FI handle), but I'd prefer
    ' to check its existence up front so we don't waste anyone's time.
    If (Not ImageFormats.IsFreeImageEnabled) Then
        PDDebug.LogAction "Animated GIF export failed; FreeImage missing."
        ExportGIF_Animated_LL = False
        Exit Function
    End If
    
    'We now begin a long phase of "optimizing" the exported animation.  This involves comparing
    ' neighboring frames against each other, cropping out identical regions (and possibly
    ' blanking out shared overlapping pixels), figuring out optimal frame disposal strategies,
    ' and possibly generating unique palettes for each frame and/or using a mix of local and
    ' global palettes.
    '
    'At the end of this phase, we'll have an array of optimized GIF frames (and all associated
    ' parameters) which we can then hand off to any capable GIF encoder.
    Dim imgPalette() As RGBQuad, palSize As Long
    Dim tmpLayer As pdLayer
    
    'GIF files support a "global palette".  This is a shared palette that any frame can choose to
    ' use (in place of a "local palette").
    
    'PhotoDemon always writes a global palette, because even if just the first frame uses it,
    ' there is no increase in file size (as the first frame will simply not provide a local palette).
    ' If, however, the first frame does *not* require a full 256-color palette, we will merge colors
    ' from subsequent frames into the global palette, until we arrive at 256 colors (or until all
    ' colors in all frames have been assembled).
    
    'This global palette starts out on the range [0, 255] but may be shrunk if fewer colors are used.
    ReDim m_globalPalette(0 To 255) As RGBQuad
    
    'Color trackers exist for both global and local palettes (PD may use one or both)
    m_numColorsInGP = 0
    Dim numColorsInLP As Long
    
    Dim idxPalette As Long
    
    'Frames that only use the global palette are tagged accordingly; they will skip embedding
    ' a local palette if they can.
    Dim frameUsesGP As Boolean: frameUsesGP = False
    
    'Because of the way GIFs are encoded (reliance on a global palette, as just mentioned),
    ' we can produce smaller files if we perform all optimizations "up front" instead of
    ' writing the GIF as-we-go. (Writing as-we-go would prevent things like global palette
    ' optimization, because it has to be provided at the front of the file).  As such, we
    ' build an optimized frame collection in advance, before writing anything to disk -
    ' which comes with the nice perk of making the actual GIF encoding step "encoder
    ' agnostic", e.g. any GIF encoder can be dropped-in at the end, because we perform
    ' all optimizations internally.
    ReDim m_allFrames(0 To srcPDImage.GetNumOfLayers - 1) As PD_GifFrame
        
    'GIFs are obnoxious because each frame specifies a "frame disposal" requirement; this is
    ' what to do with the screen buffer *after* the current frame is displayed.  We calculate
    ' this using data from the next frame in line (because its transparency requirements
    ' are ultimately what determine the frame disposal requirements of the *previous* frame).
    '
    'Frames are generally cleared by default; subsequent analyses will set this value on a
    ' per-frame basis, after testing different optimization strategies.
    Dim i As Long
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        m_allFrames(i).frameDisposal = FIFD_GIF_DISPOSAL_BACKGROUND
    Next i
    
    'As part of optimizing frames, we need to keep a running copy several different frames:
    ' 1) what the current frame looks like right now
    ' 2) what the previous frame looked like
    ' 3) what our current frame looked like before we started fucking with it
    '
    'These are used as part of exploring different optimization strategies.
    Dim prevFrame As pdDIB, bufferFrame As pdDIB, curFrameBackup As pdDIB
    Set prevFrame = New pdDIB
    Set bufferFrame = New pdDIB
    Set curFrameBackup = New pdDIB
    
    'We also use a soft reference that we can point at whatever frame DIB we want;
    ' its contents are not maintained between frames...
    Dim refFrame As pdDIB
    
    '...as well as a temporary "test" DIB, which is used to test optimizations that may
    ' not pan out (because they actually harm compression ratio in the end).
    Dim testDIB As pdDIB
    Set testDIB = New pdDIB
    
    'Some parts of the optimization process are not guaranteed to improve file size
    ' (e.g. blanking out duplicate pixels between frames can actually hurt compression
    ' if the current frame is very noisy). To ensure we only apply beneficial optimizations,
    ' we test-compress the current frame after potentially problematic optimizations to
    ' double-check that our compression ratio improved.  (If it didn't, we'll roll back the
    ' changes we made - see the multiple DIB copies above!)
    '
    '(Note that for performance reasons, we use libdeflate instead of an actual GIF encoder.
    ' I assume that the best-case result with libdeflate correlates strongly with the best-case
    ' result for a GIF-style LZW compressor, since LZ77 and LZ78 compression share most
    ' critical aspects.  It may be beneficial to explore actual LZ78 compression in the future.)
    '
    'To reduce memory churn, we initialize a single worst-case-size buffer in advance,
    ' then reuse it for all compression test runs.
    Dim cmpTestBuffer() As Byte, cmpTestBufferSize As Long
    cmpTestBufferSize = Compression.GetWorstCaseSize(srcPDImage.Width * srcPDImage.Height * 4, cf_Zlib, 1)
    ReDim cmpTestBuffer(0 To cmpTestBufferSize - 1) As Byte
    
    'We also want to know if the source image is non-paletted (e.g. "full color").
    ' If it isn't (meaning if it's already <= 256 colors per frame), the source pixel data probably
    ' came from an existing animated GIF file, and we'll want to optimize the data differently.
    '
    'Also, if auto-dithering is enabled, we dither frames *only* when the source data is full-color.
    Dim sourceIsFullColor As Boolean
    Set tmpLayer = New pdLayer
    tmpLayer.CopyExistingLayer srcPDImage.GetLayerByIndex(0)
    tmpLayer.ConvertToNullPaddedLayer srcPDImage.Width, srcPDImage.Height, True
    
    sourceIsFullColor = (Palettes.GetDIBColorCount(tmpLayer.layerDIB, True) > 256)
    If autoDither Then useDithering = sourceIsFullColor
    
    'If we detect two identical back-to-back frames (surprisingly common in GIFs "in the wild"),
    ' we will simply merge their frame times into a single value and remove the duplicate frame.
    ' This reduces file size "for free", but it requires more complicated tracking as the number
    ' of frames may decrease as optimization proceeds.
    Dim numGoodFrames As Long, lastGoodFrame As Long
    numGoodFrames = 0
    lastGoodFrame = 0
    
    'We are now going to iterate through all layers in the image TWICE.
    Dim finalFrameTime As Long
    
    'On this first pass, we will analyze each layer, produce optimized global and
    ' local palettes, extract frame times from layer names, and determine regions
    ' of interest in each frame.  Then we will palettize each layer and cache the
    ' palettized pixels in a simple 1D array.
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        
        'Optimizing frames can take some time.  Keep the user apprised of our progress.
        ProgressBars.SetProgBarVal i
        Message "Saving animation frame %1 of %2...", i + 1, srcPDImage.GetNumOfLayers()
        
        With m_allFrames(i)
        
            'Before dealing with pixel data, attempt to retrieve a frame time from the
            ' source layer's name. (If the layer name does not provide a frame time,
            ' or if the user has specified a fixed frame time, this value will be
            ' overwritten with their requsted value.)
            finalFrameTime = GetFrameTimeFromLayerName(srcPDImage.GetLayerByIndex(i).GetLayerName, 0)
            If (useFixedFrameDelay Or (finalFrameTime = 0)) Then finalFrameTime = frameDelayDefault
            .frameTime = finalFrameTime
            
            'Remaining parameters are contingent on optimization passes; for now,
            ' populate with safe default parameters (e.g parameters that produce a
            ' functional GIF even if optimization fails). Subsequent optimization
            ' rounds will modify these settings if it produces a smaller file.
            .frameDisposal = FIFD_GIF_DISPOSAL_LEAVE
            .frameIsDuplicateOrEmpty = False
            .rectOfInterest.Left = 0
            .rectOfInterest.Top = 0
            .rectOfInterest.Width = srcPDImage.Width
            .rectOfInterest.Height = srcPDImage.Height
            
        End With
        
        'Make sure this layer is the same size as the parent image, and apply any
        ' non-destructive transforms.  (Note that we *don't* do this for the first frame,
        ' because we already performed that step above as part of whole-image heuristics)
        If (i > 0) Then
            Set tmpLayer = New pdLayer
            tmpLayer.CopyExistingLayer srcPDImage.GetLayerByIndex(i)
            tmpLayer.ConvertToNullPaddedLayer srcPDImage.Width, srcPDImage.Height, True
        End If
        
        'Ensure we have a target DIB to operate on; the final, optimized frame will be stored here.
        If (curFrameBackup Is Nothing) Then Set curFrameBackup = New pdDIB
        curFrameBackup.CreateFromExistingDIB tmpLayer.layerDIB
        Set m_allFrames(i).frameDIB = New pdDIB
        
        'Force alpha to 0 or 255 only (this is a GIF requirement). It simplifies
        ' subsequent steps to handle this up-front.
        Dim trnsTable() As Byte
        DIBs.ApplyAlphaCutoff_Ex tmpLayer.layerDIB, trnsTable, gifAlphaCutoff
        DIBs.ApplyBinaryTransparencyTable tmpLayer.layerDIB, trnsTable, gifMatteColor
        
        'The first frame in the file must always be full-size, per the spec.
        ' We are not really allowed to optimize it.  (Technically, you can write
        ' GIFs whose first frame is smaller than the image as a whole.  PD does not
        ' do this because it's terrible practice, and the file size savings are not
        ' meaningfully better.)
        If (i = 0) Then
        
            'Cache the temporary layer DIB as-is; it serves as both the first frame in
            ' the animation, and the fallback "static" image for decoders that don't
            ' understand animated GIFs.
            m_allFrames(i).frameDIB.CreateFromExistingDIB tmpLayer.layerDIB
            
            'Initialize the frame buffer to be the same as the first frame...
            bufferFrame.CreateFromExistingDIB tmpLayer.layerDIB
            
            '...and initialize the "previous" frame buffer to pure transparency (this is complicated
            ' for GIFs because the recommendation is to avoid this disposal method entirely - due to
            ' the complications it requires in the decoder - but they also suggest that the decoder
            ' can ignore these instructions and use the background color of the GIF "if they have to".
            ' In PNGs, the spec says to use transparent black but GIFs may not support transparency
            ' at all... so I'm not sure what to do.  For now, I'm using the same strategy as PNGs
            ' and will revisit if problems arise.
            prevFrame.CreateBlank tmpLayer.layerDIB.GetDIBWidth, tmpLayer.layerDIB.GetDIBHeight, 32, 0, 0
            
            'Finally, mark this as the "last good frame" (in case subsequent frames are duplicates
            ' of this one, we'll need to refer back to the "last good frame" index)
            lastGoodFrame = i
            numGoodFrames = i + 1
        
        'If this is *not* the first frame, there are many ways we can optimize frame contents.
        Else
        
            'First, we want to figure out how much of this frame needs to be used at all.
            ' If this frame reuses regions from the previous frame (an extremely common
            ' occurrence in animation), we can simply crop out any frame borders that are
            ' identical to the previous frame.  (Said another way: only the first frame really
            ' needs to be the full size of the image - subsequent frames can be *any* size we
            ' want, so long as they remain within the image's boundaries.)
            
            '(Note also that if this check fails, it means that this frame is 100% identical
            ' to the frame that came before it.)
            Dim dupArea As RectF
            If DIBs.GetRectOfInterest_Overlay(tmpLayer.layerDIB, bufferFrame, dupArea) Then
                
                'This frame contains at least one unique pixel, so it needs to be added to the file.
                
                'Before proceeding further, let's compare this frame to one other buffer -
                ' specifically, the frame buffer as it appeared *before* the previous frame
                ' was painted.  GIFs define a "previous" disposal mode, which tells the previous frame
                ' to "undo" its rendering when it's done.  On certain frames and/or animation styles,
                ' this may allow for better compression if this frame is more identical to a frame
                ' further back in the animation sequence.
                Dim prevFrameArea As RectF
                
                'As before, ensure that the previous check succeeded.  If it fails, it means this frame
                ' is 100% identical the the frame that preceded the previous frame.  Rather than encode
                ' this frame at all, we can simply store a lone transparent pixel and paint it "over"
                ' the corresponding frame buffer - maximum file size savings!  (This works very well on
                ' blinking-style animations, for example.)
                If DIBs.GetRectOfInterest_Overlay(tmpLayer.layerDIB, prevFrame, prevFrameArea) Then
                
                    'With an overlap rectangle calculated for both cases, determine a "winner"
                    ' (where "winner" equals "least number of pixels"), store its frame rectangle,
                    ' and then mark the *previous* frame's disposal op accordingly.  (That's important -
                    ' disposal ops describe what you do *after* the current frame is painted, so if
                    ' we want a certain frame buffer state *before* rendering this frame, we set it
                    ' via the disposal op of the *previous* frame).
                    
                    'If the frame *before* the frame *before* this one is smallest...
                    If (prevFrameArea.Width * prevFrameArea.Height) < (dupArea.Width * dupArea.Height) Then
                        Set refFrame = prevFrame
                        m_allFrames(i).rectOfInterest = prevFrameArea
                        m_allFrames(lastGoodFrame).frameDisposal = FIFD_GIF_DISPOSAL_PREVIOUS
                        
                    'or if the frame immediately preceding this one is smallest...
                    Else
                        Set refFrame = bufferFrame
                        m_allFrames(i).rectOfInterest = dupArea
                        m_allFrames(lastGoodFrame).frameDisposal = FIFD_GIF_DISPOSAL_LEAVE
                    End If
                    
                    'We now have the smallest possible rectangle that defines this frame,
                    ' while accounting for both DISPOSAL_LEAVE and DISPOSAL_PREVIOUS.
                    
                    'We have one more potential crop operation we can do, and it involves the third
                    ' disposal op (DISPOSAL_BACKGROUND).  This disposal op asks the previous frame
                    ' to erase itself completely after rendering.  For animations with large
                    ' transparent borders, it may actually be best to crop the current frame
                    ' according to its transparent borders, then use the "erase" disposal op before
                    ' displaying it, thus forgoing any connection whatsoever to preceding frames.
                    Dim trnsRect As RectF
                    If DIBs.GetRectOfInterest(tmpLayer.layerDIB, trnsRect) Then
                    
                        'If this frame is smaller than the previous "winner", switch to using this op instead.
                        If (trnsRect.Width * trnsRect.Height) < (m_allFrames(i).rectOfInterest.Width * m_allFrames(i).rectOfInterest.Height) Then
                            m_allFrames(lastGoodFrame).frameDisposal = FIFD_GIF_DISPOSAL_BACKGROUND
                            m_allFrames(i).rectOfInterest = trnsRect
                        End If
                        
                        'Crop the "winning" region into a separate DIB, and store it as the formal pixel buffer
                        ' for this frame.
                        With m_allFrames(i).rectOfInterest
                            m_allFrames(i).frameDIB.CreateBlank Int(.Width), Int(.Height), 32, 0, 0
                            GDI.BitBltWrapper m_allFrames(i).frameDIB.GetDIBDC, 0, 0, Int(.Width), Int(.Height), tmpLayer.layerDIB.GetDIBDC, Int(.Left), Int(.Top), vbSrcCopy
                        End With
                    
                    'This weird (but valid) branch means that the current frame is 100% transparent.  For this
                    ' special case, request that the previous frame erase the running buffer, then store a 1 px
                    ' transparent pixel.
                    Else
                        
                        m_allFrames(i).frameDIB.CreateBlank 1, 1, 32, 0, 0
                        With m_allFrames(i).rectOfInterest
                            .Left = 0
                            .Top = 0
                            .Width = 1
                            .Height = 1
                        End With
                        m_allFrames(i).frameNeedsTransparency = True
                        m_allFrames(lastGoodFrame).frameDisposal = FIFD_GIF_DISPOSAL_LEAVE
                        
                    End If
                    
                    'Because the current frame came from a premultiplied source, we can safely
                    ' mark it as premultiplied as well.
                    If (Not m_allFrames(i).frameDIB Is Nothing) Then m_allFrames(i).frameDIB.SetInitialAlphaPremultiplicationState True
                    
                    'If the previous frame is not being blanked, we have additional optimization
                    ' strategies to attempt.  (If, however, the previous frame *is* being blanked,
                    ' we are done with preprocessing because we have no "previous" data to work with.)
                    If (m_allFrames(lastGoodFrame).frameDisposal <> FIFD_GIF_DISPOSAL_BACKGROUND) Then
                    
                        'The next optimization we want to attempt is duplicate pixel blanking,
                        ' which takes pixels in the current frame that are identical to the previous frame
                        ' and makes them  transparent, allowing the previous frame to "show through"
                        ' in those regions (instead of storing all that pixel data again in the current frame).
                        
                        'The more pixels we can turn transparent, the better the resulting buffer will compress,
                        ' but note that there are two major caveats to this optimization.  Specifically:
                        
                        '1) The previous frame must use DisposeOp_Leave or DisposeOp_Previous.  If it erases
                        '   the frame (DisposeOp_Background), the previous frame's pixels aren't around for us
                        '   to use.  (We caught this possibility with the If statement above, FYI.)
                        '2) Because this approach requires us to alphablend the current frame "over" the previous
                        '   one (instead of simply *replacing* the previous frame's contents with this one), we
                        '   need to ensure there are no transparency mismatches.  Specifically, if this frame has
                        '   transparency where the previous frame DOES NOT, we can't use this frame blanking
                        '   strategy (as this frame's transparent regions will allow the previous frame to
                        '   "show through" where it shouldn't).
                        
                        'Because (1) has already been taken care of by the frame disposal If/Then statement above,
                        ' we now want to address case (2).
                        If DIBs.RetrieveTransparencyTable(m_allFrames(i).frameDIB, trnsTable) Then
                        If DIBs.CheckAlpha_DuplicatePixels(refFrame, m_allFrames(i).frameDIB, trnsTable, Int(m_allFrames(i).rectOfInterest.Left), Int(m_allFrames(i).rectOfInterest.Top)) Then
                            
                            'This frame contains transparency where the previous frame does not.
                            ' This means the previous frame *must* be blanked.
                            ' Skip any remaining frame differential optimizations entirely.
                            m_allFrames(lastGoodFrame).frameDisposal = FIFD_GIF_DISPOSAL_BACKGROUND
                            
                            'As a consequence of "blanking" the previous frame, we need to render the current
                            ' frame in its entirety (except for transparent borders, if they exist).
                            With m_allFrames(i).rectOfInterest
                                
                                If DIBs.GetRectOfInterest(tmpLayer.layerDIB, trnsRect) Then
                                    .Left = trnsRect.Left
                                    .Top = trnsRect.Top
                                    .Width = trnsRect.Width
                                    .Height = trnsRect.Height
                                
                                'Failsafe only; this state was dealt with in a previous step.
                                Else
                                    .Left = 0
                                    .Top = 0
                                    .Width = srcPDImage.Width
                                    .Height = srcPDImage.Height
                                End If
                                
                                'Crop the "winning" region into a separate DIB, and store it as the formal
                                ' pixel buffer for this frame.
                                m_allFrames(i).frameDIB.CreateBlank Int(.Width), Int(.Height), 32, 0, 0
                                GDI.BitBltWrapper m_allFrames(i).frameDIB.GetDIBDC, 0, 0, Int(.Width), Int(.Height), tmpLayer.layerDIB.GetDIBDC, Int(.Left), Int(.Top), vbSrcCopy
                            
                            End With
                            
                        'The current frame can be safely alpha-blended "over the top" of the previous one
                        Else
                            
                            'This frame is a candidate for frame differentials!
                            m_allFrames(i).frameNeedsTransparency = True
                            
                            'A transparent index must be provided, so we have a mechanism for "erasing" parts
                            ' of this frame to allow the previous frame can show through.
                            
                            'Before proceeding, let's get a rough idea of the current frame's entropy.  If
                            ' subsequent optimizations increase entropy (and thus decrease compression ratio),
                            ' we'll just revert to the current frame as-is.
                            
                            'An easy (and fast) way to estimate entropy is to just compress the current frame!
                            ' Better compression ratios correlate with lower source entropy, and this is
                            ' faster than a formal entropy calculation (in VB6 code, anyway).
                            
                            '(This is also where we use our persistent compression buffer; note that we
                            ' don't need to clear or prep it in any way - the compression engine will
                            ' overwrite whatever it needs to.)
                            Dim initSize As Long, initCmpSize As Long
                            initCmpSize = cmpTestBufferSize
                            initSize = m_allFrames(i).frameDIB.GetDIBStride * m_allFrames(i).frameDIB.GetDIBHeight
                            Plugin_libdeflate.CompressPtrToPtr VarPtr(cmpTestBuffer(0)), initCmpSize, m_allFrames(i).frameDIB.GetDIBPointer, initSize, 1, cf_Zlib
                            
                            'With current entropy established (well... "estimated"), we're next going to try
                            ' blanking out any pixels that are identical between this frame and the current frame
                            ' buffer (as calculated in a previous step).  On many animations, this will create
                            ' large patches of pure transparency that compress *brilliantly* - but note that very
                            ' noisy images - like full-color images originally converted w/dithering - this may
                            ' produce equally noisy results, which can actually *harm* compression ratio. This is
                            ' why we estimated entropy for the untouched frame data (above), and why we're gonna
                            ' perform all our tests on a temporary frame copy (in case we have to throw the copy away).
                            testDIB.CreateFromExistingDIB m_allFrames(i).frameDIB
                            
                            If DIBs.RetrieveTransparencyTable(testDIB, trnsTable) Then
                            If DIBs.ApplyAlpha_DuplicatePixels(testDIB, refFrame, trnsTable, m_allFrames(i).rectOfInterest.Left, m_allFrames(i).rectOfInterest.Top) Then
                            If DIBs.ApplyTransparencyTable(testDIB, trnsTable) Then
                        
                                'The frame differential was produced successfully.  See if it compresses better
                                ' than the original, untouched frame did.
                                Dim testSize As Long
                                testSize = cmpTestBufferSize
                                Plugin_libdeflate.CompressPtrToPtr VarPtr(cmpTestBuffer(0)), testSize, testDIB.GetDIBPointer, initSize, 1, cf_Zlib
                                
                                'This frame compressed better!  Use it instead of the original frame.
                                If (testSize < initCmpSize) Then
                                    m_allFrames(i).frameDIB.CreateFromExistingDIB testDIB
                                    
                                'If this frame compressed worse (or identically), we can simply leave current
                                ' frame settings as they are.
                                End If
                            
                            '/end failsafe "retrieve and apply duplicate pixel test" checks
                            End If
                            End If
                            End If
                            
                        '/end "previous frame is opaque where this frame is transparent"
                        End If
                        End If
                    
                    '/end "previous frame is NOT being blanked, so we can attempt to optimize frame diffs"
                    End If
                    
                'This (rare) branch means that the current frame is identical to the animation as it appeared
                ' *before* the previous frame was rendered.  (This is not unheard of, especially for blinking-
                ' or spinning-style animations.)  A great, cheap optimization is to just ask the previous
                ' frame to dispose of itself using the DISPOSE_OP_PREVIOUS method (which restores the frame
                ' buffer to whatever it was *before* the previous frame was rendered), then store this frame
                ' as a transparent 1-px GIF.  This effectively gives us a copy of the frame two-frames-previous
                ' "for free".
                Else
                    
                    'Ensure transparency is available
                    m_allFrames(i).frameNeedsTransparency = True
                    m_allFrames(i).frameDIB.CreateBlank 1, 1, 32, 0, 0
                    With m_allFrames(i).rectOfInterest
                        .Left = 0
                        .Top = 0
                        .Width = 1
                        .Height = 1
                    End With
                    m_allFrames(lastGoodFrame).frameDisposal = FIFD_GIF_DISPOSAL_PREVIOUS
                
                End If
                
            'If the GetRectOfInterest() check failed, it means this frame is 100% identical to the
            ' frame that preceded it.  Rather than optimize this frame, let's just delete it from
            ' the animation and merge its frame time into the previous frame.
            Else
                m_allFrames(i).frameIsDuplicateOrEmpty = True
                m_allFrames(lastGoodFrame).frameTime = m_allFrames(lastGoodFrame).frameTime + m_allFrames(i).frameTime
                Set m_allFrames(i).frameDIB = Nothing
            End If
            
            'This frame is now optimized as well as we can possibly optimize it.
            
            'Before moving to the next frame, create backup copies of the buffer frames
            ' *we* were handed.  The next frame can request that we reset our state to this
            ' frame, which may be closer to their frame's contents (and thus compress better).
            If (m_allFrames(lastGoodFrame).frameDisposal = FIFD_GIF_DISPOSAL_LEAVE) Then
                prevFrame.CreateFromExistingDIB bufferFrame
            ElseIf (m_allFrames(lastGoodFrame).frameDisposal = FIFD_GIF_DISPOSAL_BACKGROUND) Then
                prevFrame.ResetDIB 0
            
            'We don't have to cover the case of DISPOSE_OP_PREVIOUS, as that's the state the prevFrame
            ' DIB is already in!
            'Else

            End If
            
            'Overwrite the *current* frame buffer with an (untouched) copy of this frame, as it appeared
            ' before we applied optimizations to it.
            bufferFrame.CreateFromExistingDIB curFrameBackup
            
            'If this frame is valid (e.g. not a duplicate of the previous frame), increment our current
            ' "good frame" count, and mark this frame as the "last good" index.
            If (Not m_allFrames(i).frameIsDuplicateOrEmpty) Then
                numGoodFrames = numGoodFrames + 1
                lastGoodFrame = i
            End If
        
        'i !/= 0 branch
        End If
        
        'With optimizations accounted for, it is now time to palettize this layer.
        
        'Generate an optimal 256-color palette for the image.  (TODO: move this to our neural-network quantizer.)
        Palettes.GetOptimizedPaletteIncAlpha m_allFrames(i).frameDIB, imgPalette, 256, pdqs_Variance, True
        numColorsInLP = UBound(imgPalette) + 1
        
        'Ensure that in the course of producing an optimal palette, the optimizer didn't change
        ' any transparent values to number other than 0 or 255.
        Dim pEntry As Long
        For pEntry = LBound(imgPalette) To UBound(imgPalette)
            If (imgPalette(pEntry).Alpha < 127) Then
                imgPalette(pEntry).Alpha = 0
            Else
                imgPalette(pEntry).Alpha = 255
            End If
        Next pEntry
        
        'If the current frame requires transparency, ensure transparency exists in the palette.
        If m_allFrames(i).frameNeedsTransparency Then
            
            Dim trnsFound As Boolean: trnsFound = False
            For pEntry = 0 To UBound(imgPalette)
                If (imgPalette(pEntry).Alpha = 0) Then
                    trnsFound = True
                    Exit For
                End If
            Next pEntry
            
            'If transparency *wasn't* found, add it manually (if there's room), or generate a new
            ' 255-color palette and stick transparency at the end.
            If (Not trnsFound) Then
                
                If (numColorsInLP = 256) Then
                    numColorsInLP = 255
                    Palettes.GetOptimizedPaletteIncAlpha m_allFrames(i).frameDIB, imgPalette, 255, pdqs_Variance, True
                End If
                
                ReDim Preserve imgPalette(0 To numColorsInLP) As RGBQuad
                imgPalette(numColorsInLP).Blue = 0
                imgPalette(numColorsInLP).Green = 0
                imgPalette(numColorsInLP).Red = 0
                imgPalette(numColorsInLP).Alpha = 0
                numColorsInLP = numColorsInLP + 1
                
            End If
            
        End If
        
        'Frames that need transparency are now guaranteed to have it.
        
        'If this is the *first* frame, we will use it as the basis of our global palette.
        If (i = 0) Then
        
            'Simply copy over the palette as-is into our running global palette tracker
            m_numColorsInGP = numColorsInLP
            ReDim m_globalPalette(0 To m_numColorsInGP - 1) As RGBQuad
            
            For idxPalette = 0 To m_numColorsInGP - 1
                m_globalPalette(idxPalette) = imgPalette(idxPalette)
            Next idxPalette
            
            'Sort the palette by popularity (with a few tweaks), which can eke out slightly
            ' better compression ratios.
            Palettes.SortPaletteForCompression_IncAlpha m_allFrames(i).frameDIB, m_globalPalette, True, True
            
            frameUsesGP = True
        
        'If this is *not* the first frame, and we have yet to write a global palette, append as many
        ' unique colors from this palette as we can into the global palette.
        Else
            
            'If there's still room in the global palette, append this palette to it.
            If (m_numColorsInGP < 256) Then
                
                m_numColorsInGP = Palettes.MergePalettes(m_globalPalette, m_numColorsInGP, imgPalette, numColorsInLP)
                
                'Enforce a strict 256-color limit; colors past the end will simply be discarded, and this frame
                ' will use a local palette instead.
                If (m_numColorsInGP > 256) Then
                    m_numColorsInGP = 256
                    ReDim Preserve m_globalPalette(0 To 255) As RGBQuad
                End If
                
            End If
            
            'Next, we need to see if all colors in this frame appear in the global palette.
            ' If they do, we can simply use the global palette to write this frame.
            frameUsesGP = Palettes.DoesPaletteContainPalette(m_globalPalette, m_numColorsInGP, imgPalette, numColorsInLP)
            
        End If
        
        m_allFrames(i).usesGlobalPalette = frameUsesGP
        
        'With all optimizations applied, we are finally ready to palettize this layer.
        
        'If this frame requires a local palette, sort the local palette (to optimize compression ratios),
        ' then cache a copy of the palette before proceeding.
        If (Not m_allFrames(i).usesGlobalPalette) Then
            
            'Sort the palette prior to saving it; this can improve compression ratios
            Palettes.SortPaletteForCompression_IncAlpha m_allFrames(i).frameDIB, imgPalette, True, True
            
            m_allFrames(i).palNumColors = UBound(imgPalette) + 1
            ReDim m_allFrames(i).framePalette(0 To UBound(imgPalette))
            
            For idxPalette = 0 To UBound(imgPalette)
                m_allFrames(i).framePalette(idxPalette) = imgPalette(idxPalette)
            Next idxPalette
            
        End If
        
        'If this frame is a duplicate of the previous frame, we don't need to perform any more
        ' optimizations on its pixel data, because we will simply reuse the previous frame in
        ' its place.
        If (Not m_allFrames(i).frameIsDuplicateOrEmpty) Then
            
            'Using either the local or global palette (whichever matches this image),
            ' create an 8-bit version of the source image.  (TODO: switch to neural network quantizer)
            If frameUsesGP Then
                palSize = m_numColorsInGP
                If useDithering Then
                    Palettes.GetPalettizedImage_Dithered_IncAlpha m_allFrames(i).frameDIB, m_globalPalette, m_allFrames(i).pixelData, PDDM_SierraLite, 0.67, True
                Else
                    DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(i).frameDIB, m_globalPalette, m_allFrames(i).pixelData
                End If
            Else
                palSize = numColorsInLP
                If useDithering Then
                    Palettes.GetPalettizedImage_Dithered_IncAlpha m_allFrames(i).frameDIB, imgPalette, m_allFrames(i).pixelData, PDDM_SierraLite, 0.67, True
                Else
                    DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(i).frameDIB, imgPalette, m_allFrames(i).pixelData
                End If
            End If
            
        End If
        
        'We can now free the (temporary) frame DIB copy, because it is either unnecessary
        ' (because this frame isn't being encoded) or because it has been palettized and
        ' stored inside m_allFrames(i).pixelData.
        Set m_allFrames(i).frameDIB = Nothing
        
    'Next frame
    Next i
    
    'Clear all optimization-related objects, as they are no longer needed
    Set prevFrame = Nothing
    Set bufferFrame = Nothing
    Set curFrameBackup = Nothing
    Set refFrame = Nothing
    Set testDIB = Nothing
    Erase cmpTestBuffer
    
    'Before generating a GIF file, let's get our global palette in order.
    
    ' The GIF spec requires global palette color count to be a power of 2.  (It does this because
    ' the compression table will only use n bits for each of 2 ^ n colors.)
    If (m_numColorsInGP < 2) Then
        m_numColorsInGP = 2
    ElseIf (m_numColorsInGP < 4) Then
        m_numColorsInGP = 4
    ElseIf (m_numColorsInGP < 8) Then
        m_numColorsInGP = 8
    ElseIf (m_numColorsInGP < 16) Then
        m_numColorsInGP = 16
    ElseIf (m_numColorsInGP < 32) Then
        m_numColorsInGP = 32
    ElseIf (m_numColorsInGP < 64) Then
        m_numColorsInGP = 64
    ElseIf (m_numColorsInGP < 128) Then
        m_numColorsInGP = 128
    Else
        m_numColorsInGP = 256
    End If
    
    'Since we need to CopyMemory the palette into our encoder,
    ' make sure we've allocated enough bytes to match the final color count.
    If (UBound(m_globalPalette) <> m_numColorsInGP - 1) Then ReDim Preserve m_globalPalette(0 To m_numColorsInGP - 1) As RGBQuad
    
    'If the global palette has a transparent index, locate it in advance
    m_GlobalTrnsIndex = -1
    For i = 0 To m_numColorsInGP - 1
        If (m_globalPalette(i).Alpha = 0) Then
            m_GlobalTrnsIndex = i
            Exit For
        End If
    Next i
    
    'We've successfully optimized all GIF frames.  Now it's time to involve our encoder.
    Message "Finalizing image..."
    
    ExportGIF_Animated_LL = WriteOptimizedAGIFToFile(srcPDImage, dstFile, formatParams, metadataParams)
    
    ProgressBars.SetProgBarVal 0
    ProgressBars.ReleaseProgressBar
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportGIF_Animated_LL = False
    
End Function

'After optimizing GIF frames (which you must do to generate the structures used by this function),
' call this function to actually write GIF data out to file.  In this build, FreeImage is used.
' This may change pending further testing.
Private Function WriteOptimizedAGIFToFile(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean

    On Error GoTo ExportGIFError
    
    WriteOptimizedAGIFToFile = False
    Dim sFileType As String: sFileType = "GIF"
    
    'Parse all relevant GIF parameters.  (See the GIF export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to
    ' a new file, and if it's saved successfully, overwrite the original file - this way,
    ' if an error occurs mid-save, the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Dim cRandom As pdRandomize
        Set cRandom = New pdRandomize
        cRandom.SetSeed_AutomaticAndRandom
        tmpFilename = dstFile & Hex$(cRandom.GetRandomInt_WH()) & ".pdtmp"
    Else
        tmpFilename = dstFile
    End If
    
    Dim tmpTag As FREE_IMAGE_TAG
    
    'Create a blank multipage FreeImage object.  All GIF frames will be "added" to this object.
    Dim fi_MasterHandle As Long
    fi_MasterHandle = FreeImage_OpenMultiBitmap(FIF_GIF, tmpFilename, True, False, True)
    If (fi_MasterHandle <> 0) Then
    
        'We are now ready to write the GIF file
        Dim i As Long
        For i = 0 To srcPDImage.GetNumOfLayers - 1
            
            If (Not m_allFrames(i).frameIsDuplicateOrEmpty) Then
                
                'Allocate an 8-bpp FreeImage DIB at the same size as the source layer, and populate it with our
                ' palette and pixel data.  (Note that we don't actually use the local palette for frames that use
                ' the global palette - but we have to supply *something* in order to construct the FI image.)
                Dim fi_DIB As Long
                With m_allFrames(i)
                    
                    If .usesGlobalPalette Then
                        fi_DIB = Plugin_FreeImage.GetFIDIB_8Bit(Int(.rectOfInterest.Width), Int(.rectOfInterest.Height), VarPtr(.pixelData(0, 0)), VarPtr(m_globalPalette(0)), m_numColorsInGP)
                    Else
                        fi_DIB = Plugin_FreeImage.GetFIDIB_8Bit(Int(.rectOfInterest.Width), Int(.rectOfInterest.Height), VarPtr(.pixelData(0, 0)), VarPtr(.framePalette(0)), .palNumColors)
                    End If
                    
                    'Pixel data is now unnecessary; free it!
                    Erase .pixelData
                    
                End With
                
                'If the FI object was created successfully, append any required animation metadata,
                ' then append the finished FI object to the parent multipage object
                If (fi_DIB <> 0) Then
                    
                    'If this is the first page in the file, write any parameters that affect the image as a whole
                    If (i = 0) Then
                    
                        'Loop count
                        Dim loopCount As Long
                        loopCount = cParams.GetLong("animation-loop-count", 1)
                        If (loopCount > 65536) Then loopCount = 65536
                        tmpTag = Outside_FreeImageV3.FreeImage_CreateTagEx(FIMD_ANIMATION, "Loop", FIDT_LONG, loopCount, 1, &H4&)
                        If (Not Outside_FreeImageV3.FreeImage_SetMetadataEx(fi_DIB, tmpTag)) Then PDDebug.LogAction "WARNING! ImageExporter.ExportGIF_Animated_LL failed to set a tag"
                        
                        'Global palette
                        If (Not FreeImage_CreateTagTanner(fi_DIB, FIMD_ANIMATION, "GlobalPalette", FIDT_PALETTE, VarPtr(m_globalPalette(0)), m_numColorsInGP, m_numColorsInGP * 4, &H3)) Then PDDebug.LogAction "WARNING! ImageExporter.ExportGIF_Animated_LL failed to set a tag"
                        
                    End If
                    
                    'GIFs store frame time in centiseconds - I know, a bizarre amount that makes it impossible
                    ' to achieve proper 30 or 60 fps display.  To improve output, round the specified msec amount
                    ' to the nearest csec equivalent.  Note also that most browsers enforce a minimum display rate
                    ' of their own, independent of this value (20 msec is prevalent as of 2019).
                    Dim finalFrameTime As Long
                    finalFrameTime = Int((m_allFrames(i).frameTime + 5) \ 10) * 10
                    tmpTag = Outside_FreeImageV3.FreeImage_CreateTagEx(FIMD_ANIMATION, "FrameTime", FIDT_LONG, finalFrameTime, 1, &H1005&)
                    If (Not Outside_FreeImageV3.FreeImage_SetMetadataEx(fi_DIB, tmpTag)) Then PDDebug.LogAction "WARNING! ImageExporter.ExportGIF_Animated_LL failed to set a tag"
                    
                    'Specify frame left/top for all but the first frame (which is always specified
                    ' as starting at [0, 0])
                    If (i > 0) Then
                        tmpTag = Outside_FreeImageV3.FreeImage_CreateTagEx(FIMD_ANIMATION, "FrameLeft", FIDT_SHORT, CLng(Int(m_allFrames(i).rectOfInterest.Left)), 1, &H1001&)
                        If (Not Outside_FreeImageV3.FreeImage_SetMetadataEx(fi_DIB, tmpTag)) Then PDDebug.LogAction "WARNING! ImageExporter.ExportGIF_Animated_LL failed to set a tag"
                        tmpTag = Outside_FreeImageV3.FreeImage_CreateTagEx(FIMD_ANIMATION, "FrameTop", FIDT_SHORT, CLng(Int(m_allFrames(i).rectOfInterest.Top)), 1, &H1002&)
                        If (Not Outside_FreeImageV3.FreeImage_SetMetadataEx(fi_DIB, tmpTag)) Then PDDebug.LogAction "WARNING! ImageExporter.ExportGIF_Animated_LL failed to set a tag"
                    End If
                    
                    'If we use the global palette, flag it now, including the transparent index
                    If m_allFrames(i).usesGlobalPalette Then
                        tmpTag = Outside_FreeImageV3.FreeImage_CreateTagEx(FIMD_ANIMATION, "NoLocalPalette", FIDT_BYTE, 1, 1, &H1003&)
                        If (Not Outside_FreeImageV3.FreeImage_SetMetadataEx(fi_DIB, tmpTag)) Then PDDebug.LogAction "WARNING! ImageExporter.ExportGIF_Animated_LL failed to set a tag"
                        If (m_GlobalTrnsIndex >= 0) Then FreeImage_SetTransparentIndex fi_DIB, m_GlobalTrnsIndex
                    Else
                        
                        'Note that PD prefers that the transparency index - if one exists - is always the
                        ' *first* palette index.  This improves compatibility with old GIF decoders (some of
                        ' which make this exact assumption).
                        If (m_allFrames(i).framePalette(0).Alpha = 0) Then
                            FreeImage_SetTransparentIndex fi_DIB, 0
                        
                        'If PD finds transparency in a non-ideal location, it will still write it correctly,
                        ' but you risk old GIF decoders not displaying the frames properly.
                        Else
                            Dim idxPal As Long
                            For idxPal = 0 To m_allFrames(i).palNumColors - 1
                                If (m_allFrames(i).framePalette(idxPal).Alpha = 0) Then
                                    FreeImage_SetTransparentIndex fi_DIB, idxPal
                                    PDDebug.LogAction "palette transparency in suboptimal location (" & idxPal & "); consider fixing!"
                                    Exit For
                                End If
                            Next idxPal
                        End If
                        
                    End If
                    
                    'Set this frame to either erase to background (transparent black) or retain data
                    ' from the previous frame.
                    tmpTag = Outside_FreeImageV3.FreeImage_CreateTagEx(FIMD_ANIMATION, "DisposalMethod", FIDT_BYTE, m_allFrames(i).frameDisposal, 1, &H1006&)
                    If (Not Outside_FreeImageV3.FreeImage_SetMetadataEx(fi_DIB, tmpTag)) Then PDDebug.LogAction "WARNING! ImageExporter.ExportGIF_Animated_LL failed to set a tag"
                    
                    'Append the finished frame
                    FreeImage_AppendPage fi_MasterHandle, fi_DIB
                    
                    'Make a copy of the current frame handle, as Release our local copy of the current frame (FI has copied it internally)
                    FreeImage_Unload fi_DIB
                    
                Else
                    PDDebug.LogAction "failed to produce FI DIB for frame # " & CStr(i)
                End If
                
            End If
                
        Next i
        
        'With all frames added, we can now finalize a few things.
        ProgressBars.SetProgBarVal ProgressBars.GetProgBarMax()
        
        'Finally, we can close the multipage handle "once and for all"; FreeImage handles the rest from here
        WriteOptimizedAGIFToFile = FreeImage_CloseMultiBitmap(fi_MasterHandle)
        
        'If we wrote our data to a temp file, attempt to replace the original file
        If Strings.StringsNotEqual(dstFile, tmpFilename) Then
            
            WriteOptimizedAGIFToFile = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
            
            If (Not WriteOptimizedAGIFToFile) Then
                Files.FileDelete tmpFilename
                PDDebug.LogAction "WARNING!  ImageExporter could not overwrite GIF file; original file is likely open elsewhere."
            End If
            
        End If
        
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        WriteOptimizedAGIFToFile = False
    End If
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    WriteOptimizedAGIFToFile = False
    
End Function

Private Sub ExportDebugMsg(ByRef debugMsg As String)
    PDDebug.LogAction debugMsg
End Sub
