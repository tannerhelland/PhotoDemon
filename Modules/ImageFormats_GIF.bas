Attribute VB_Name = "ImageFormats_GIF"
'***************************************************************************
'Additional support functions for GIF support
'Copyright 2001-2021 by Tanner Helland
'Created: 4/15/01
'Last updated: 24/October/21
'Last update: switch static GIF encoder to new homebrew GIF encoder; FreeImage is no longer used for any GIF features!
'
'Most image exporters exist in the ImageExporter module.  GIF is a weird exception because animated GIFs
' require a ton of preprocessing (to optimize animation frames), so I've moved them to their own home.
'
'PhotoDemon automatically optimizes saved GIFs to produce the smallest possible files.  A variety of
' optimizations are used, and the encoder tests various strategies to try and choose the "best"
' (smallest) solution on each frame.  As you can see from the size of this module, many many many
' different optimizations are attempted.  Despite this, the optimization pre-pass is reasonably quick,
' and the GIFs produced this way are often an order of magnitude (or more) smaller than a naive
' GIF encoder would produce.
'
'Note that the optimization steps are specifically written in an export-library-agnostic way.
' PD internally stores the results of all optimizations, then just hands the optimized frames off
' to an encoder at the end of the process.  Historically PD used FreeImage for animated GIF encoding,
' but FreeImage has a number of shortcomings (including woeful performance and writing larger GIFs
' than is necessary), so in 2021 we moved to an in-house LZW encoder based off the classic UNIX
' "compress" tool.  The LZW encoder lives in a separate module (ImageFormats_GIF_LZW).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Const GIF_FILE_EXTENSION As String = "gif"

Private Enum PD_GifDisposal
    gd_Unknown = 0      'Do not use
    gd_Leave = 1        'Do nothing after rendering
    gd_Background = 2   'Restore background color
    gd_Previous = 3     'Undo current frame's rendering
End Enum

#If False Then
    Private Const gd_Unknown = 0, gd_Leave = 1, gd_Background = 2, gd_Previous = 3
#End If

'The animated GIF exporter builds a collection of frame data during export.
Private Type PD_GifFrame
    usesGlobalPalette As Boolean        'GIFs allow for both global and local palettes.  PD optimizes against both, and will
                                        ' automatically use the best palette for each frame.
    frameIsDuplicateOrEmpty As Boolean  'PD automatically drops duplicate and/or empty frames
    frameNeedsTransparency As Boolean   'PD may require transparency as part of optimizing a given frame (pixel blanking).
                                        ' If the final palette ends up without transparency, we will roll back this
                                        ' optimization step as necessary.
    frameWasBlanked As Boolean          'TRUE if pixel-blanking produced a (likely) smaller frame; may need to be rolled
                                        ' back if the final palette doesn't contain (or have room for adding) transparency.
                                        ' Note also that the frame may not be *completely* blanked; instead, each scanline is
                                        ' conditionally blanked based on whether it reduces overall entropy or not.
    frameTime As Long                   'GIF frame time is in centiseconds (uuuuuuuugh); we auto-translate from ms
    frameDisposal As PD_GifDisposal     'GIF and APNG disposal methods are roughly identical.  PD may use any/all of them
                                        ' as part of optimizing each frame.
    rectOfInterest As RectF             'Frames are auto-cropped to their relevant minimal regions-of-change
    backupFrameDIB As pdDIB             'If a frame gets pixel-blanked, we'll save its original version here.  If the
                                        ' final palette can't fit transparency (global palettes are infamous for this),
                                        ' we'll revert to this original, non-blanked version of the frame.
    frameDIB As pdDIB                   'Only used temporarily, during optimization; ultimately palettized to produce...
    pixelData() As Byte                 '...this bytestream (and associated palette) instead.
    palNumColors As Long                'Stores the local palette count and color table, if one exists (it may not -
    framePalette() As RGBQuad           ' check the usesGlobalPalette bool before accessing)
End Type

'Optimized GIF frames will be stored here.  This array is auto-cleared after a successful dump to file.
Private m_allFrames() As PD_GifFrame

'PD always writes a global palette, and it attempts to use it on as many frames as possible.
' (Local palettes will automatically be generated too, as necessary.)
Private m_globalPalette() As RGBQuad, m_numColorsInGP As Long, m_GlobalTrnsIndex As Long

'Low-level GIF export interface.  As of 2021, image pre-processing (including palettization) and GIf encoding
' is all performed using homebrew code.
Public Function ExportGIF_LL(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportGIFError
    
    ExportGIF_LL = False
    
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
    
    'As always, pdStream handles actual writing duties.  (Memory mapping is used for ideal performance.)
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, tmpFilename, optimizeAccess:=OptimizeSequentialAccess) Then
        
        'A pdGIF instance handles the actual encoding
        Dim cGIF As pdGIF
        Set cGIF = New pdGIF
        If cGIF.SaveGIF_ToStream_Static(srcPDImage, cStream, formatParams, metadataParams) Then
            
            'Close the stream, then release the pdGIF instance
            cStream.StopStream
            Set cGIF = Nothing
            
            'If we wrote our data to a temp file, attempt to replace the original file
            If Strings.StringsNotEqual(dstFile, tmpFilename) Then
                
                ExportGIF_LL = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
                
                If (Not ExportGIF_LL) Then
                    Files.FileDelete tmpFilename
                    PDDebug.LogAction "WARNING!  ImageExporter could not overwrite GIF file; original file is likely open elsewhere."
                End If
            
            'Encode is already done!
            Else
                ExportGIF_LL = True
            End If
            
        Else
            PDDebug.LogAction "WARNING! pdGIF failed to save GIF"
        End If
        
        ProgressBars.SetProgBarVal 0
        ProgressBars.ReleaseProgressBar
        
    Else
        PDDebug.LogAction "WARNING!  Couldn't initialize stream against " & dstFile
    End If
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error encountered in " & GIF_FILE_EXTENSION & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportGIF_LL = False
    
End Function

'Low-level animated GIF export.  As of 2021, frame optimization and GIF encoding is all done with homebrew code.
Public Function ExportGIF_Animated_LL(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportGIFError
    
    ExportGIF_Animated_LL = False
    
    'Initialize a progress bar
    ProgressBars.SetProgBarMax srcPDImage.GetNumOfLayers * 2
    
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
    
    'We now begin a long phase of "optimizing" the exported animation.  This involves comparing
    ' neighboring frames against each other, cropping out identical regions (and possibly
    ' blanking out shared overlapping pixels), figuring out optimal frame disposal strategies,
    ' and possibly generating unique palettes for each frame and/or using a mix of local and
    ' global palettes.
    '
    'At the end of this phase, we'll have an array of optimized GIF frames (and all associated
    ' parameters) which we can then hand off to any capable GIF encoder.
    Dim imgPalette() As RGBQuad
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
    
    'We cheat and use lz4 as a fast test-run analyzer for various optimization strategies.
    ' If it shows bad results, it's likely GIF's inferior LZW scheme will struggle too; this allows
    ' us to roll back optimizations that don't appear to help the current image.
    Dim netPixels As Long, initSize As Long, initCmpSize As Long, testSize As Long
    
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
    Dim i As Long, j As Long
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        m_allFrames(i).frameDisposal = gd_Background
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
    'Note that for performance reasons, we use lz4 instead of our native GIF encoder.
    ' lz4 is a hell of a lot faster, and I assume that the best-case result with lz4
    ' correlates strongly with the best-case result for a GIF-style LZW compressor,
    ' since LZ77 and LZ78 compression share most critical aspects.
    '
    'To reduce memory churn, we initialize a single worst-case-size buffer in advance,
    ' then reuse it for all compression test runs.
    Dim cmpTestBuffer() As Byte, cmpTestBufferSize As Long
    cmpTestBufferSize = Compression.GetWorstCaseSize(srcPDImage.Width * srcPDImage.Height * 4, cf_Lz4)
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
    
    Dim tmpOriginal() As Byte, tmpBlanked() As Byte, tmpMerged() As Byte
    
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
        Message "Optimizing animation frame %1 of %2...", i + 1, srcPDImage.GetNumOfLayers + 1, "DONOTLOG"
        
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
            .frameDisposal = gd_Leave
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
            m_allFrames(i).frameDIB.SuspendDIB cf_Lz4, False
            
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
                    If Int(prevFrameArea.Width * prevFrameArea.Height) < Int(dupArea.Width * dupArea.Height) Then
                        Set refFrame = prevFrame
                        m_allFrames(i).rectOfInterest = prevFrameArea
                        m_allFrames(lastGoodFrame).frameDisposal = gd_Previous
                        
                    'or if the frame immediately preceding this one is smallest...
                    Else
                        Set refFrame = bufferFrame
                        m_allFrames(i).rectOfInterest = dupArea
                        m_allFrames(lastGoodFrame).frameDisposal = gd_Leave
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
                            m_allFrames(lastGoodFrame).frameDisposal = gd_Background
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
                        m_allFrames(lastGoodFrame).frameDisposal = gd_Leave
                        
                    End If
                    
                    'Because the current frame came from a premultiplied source, we can safely
                    ' mark it as premultiplied as well.
                    If (Not m_allFrames(i).frameDIB Is Nothing) Then m_allFrames(i).frameDIB.SetInitialAlphaPremultiplicationState True
                    
                    'If the previous frame is not being blanked, we have additional optimization
                    ' strategies to attempt.  (If, however, the previous frame *is* being blanked,
                    ' we are done with preprocessing because we have no "previous" data to work with.)
                    If (m_allFrames(lastGoodFrame).frameDisposal <> gd_Background) Then
                        
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
                            m_allFrames(lastGoodFrame).frameDisposal = gd_Background
                            
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
                            
                            'This frame is a candidate for frame differentials!  (If you're curious,
                            ' you can skip this optimization using the constant below; I did this frequently
                            ' while devising an "optimal" optimization strategy.)
                            Const ATTEMPT_PIXEL_BLANKING As Boolean = True
                            If (Not ATTEMPT_PIXEL_BLANKING) Then GoTo SkipPixelBlanking
                        
                            'Now, a quick note before we test pixel blanking on this frame. At this point,
                            ' we are still working in 32-bpp color mode (by design).  We won't waste energy
                            ' palettizing this image until we've successfully reduced it to its minimal
                            ' 32-bpp form, because otherwise we risk doing things like palettizing pixels
                            ' that will end up transparent anyway (which would waste palette entries on
                            ' pixels that don't even appear in the final image!).
                            '
                            'For PNGs, this strategy works very well because we can guarantee access to
                            ' transparency in the final image, however we decide to generate it.  GIFs are
                            ' different.  We may *not* have access to transparency in the final image
                            ' (if, for example, all frames use a single 256-color global palette - the only
                            ' way to make transparency "available" would be to delete an entry from the
                            ' global palette, or waste energy creating local palettes with a transparent
                            ' index).
                            '
                            'So what we will now do is calculate a pixel-blanked version of this frame,
                            ' and we'll store the results if they're better - BUT we'll also cache the
                            ' original, non-pixel-blanked version.  When it comes time to palettize this
                            ' frame, we'll palettize it and determine which palette to use (global or
                            ' local).  Then we'll check the target palette to see if it supports
                            ' transparency.  If it does, we'll use the pixel-blanked version; if it
                            ' doesn't, we'll revert to the original non-blanked copy.  (Which we'll have
                            ' to palettize separately, since we can't use the palette for the blanked
                            ' version.)  Cool?  Cool.
                            
                            testDIB.CreateFromExistingDIB m_allFrames(i).frameDIB
                            If DIBs.RetrieveTransparencyTable(testDIB, trnsTable) Then
                            If DIBs.ApplyAlpha_DuplicatePixels(testDIB, refFrame, trnsTable, m_allFrames(i).rectOfInterest.Left, m_allFrames(i).rectOfInterest.Top, True) Then
                            If DIBs.ApplyTransparencyTable(testDIB, trnsTable) Then
                            
                                'The frame differential was produced successfully.  We won't actually use it here;
                                ' instead, we'll use it later, after palettization (so we can test its entropy
                                ' against the original, untouched version).
                                    
                                'Back up the existing frame copy (in case we decide to use the global palette
                                ' and it lacks transparency)
                                Set m_allFrames(i).backupFrameDIB = New pdDIB
                                m_allFrames(i).backupFrameDIB.CreateFromExistingDIB m_allFrames(i).frameDIB
                                
                                'Copy the current DIB, and set all corresponding flags
                                m_allFrames(i).frameDIB.CreateFromExistingDIB testDIB
                                m_allFrames(i).frameWasBlanked = True
                                m_allFrames(i).frameNeedsTransparency = True
                                
                            '/end failsafe "retrieve and apply duplicate pixel test" checks
                            End If
                            End If
                            End If
SkipPixelBlanking:
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
                ' as a 1-px GIF that matches the top-left pixel color of the previous frame.  This effectively
                ' gives us a copy of the frame two-frames-previous "for free".  (Note that we can't just merge
                ' this frame with the previous one, because this frame doesn't *look* like frame n-1 - it looks
                ' like frame n-2, so we *must* still provide a frame here... but a 1-px one works fine!)
                Else
                    
                    'Create a 1-px DIB and set a corresponding frame rect
                    m_allFrames(i).frameNeedsTransparency = False
                    m_allFrames(i).frameDIB.CreateBlank 1, 1, 32, 0, 0
                    With m_allFrames(i).rectOfInterest
                        .Left = 0
                        .Top = 0
                        .Width = 1
                        .Height = 1
                    End With
                    m_allFrames(lastGoodFrame).frameDisposal = gd_Previous
                    
                    'Set the pixel color to match the original frame.
                    Dim tmpQuad As RGBQuad
                    If prevFrame.GetPixelRGBQuad(0, 0, tmpQuad) Then m_allFrames(i).frameDIB.SetPixelRGBQuad 0, 0, tmpQuad
                    
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
            If (m_allFrames(lastGoodFrame).frameDisposal = gd_Leave) Then
                prevFrame.CreateFromExistingDIB bufferFrame
            ElseIf (m_allFrames(lastGoodFrame).frameDisposal = gd_Background) Then
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
        ' (If this frame is a duplicate of the previous frame, we don't need to perform any more
        ' optimizations on its pixel data, because we will simply reuse the previous frame in
        ' its place.)
        If (Not m_allFrames(i).frameIsDuplicateOrEmpty) Then
            
            'Generate an optimal 256-color palette for the image.  (TODO: move this to our neural-network quantizer.)
            ' Note that this function will return an exact palette for the frame if the frame contains < 256 colors.
            Palettes.GetOptimizedPaletteIncAlpha m_allFrames(i).frameDIB, imgPalette, 256, pdqs_Variance, True
            numColorsInLP = UBound(imgPalette) + 1
            
            'If (i = 4) Then Saving.QuickSaveDIBAsPNG "C:\tanner-dev\test.png", m_allFrames(i).frameDIB
            'Ensure that in the course of producing an optimal palette, the optimizer didn't change
            ' any transparent values to number other than 0 or 255.  (Neural network quantization can be fuzzy
            ' this way - sometimes values shift minute amounts due to the way neighboring colors affect
            ' each other.)
            Palettes.EnsureBinaryAlphaPalette imgPalette
                
            'Frames that need transparency are now guaranteed to have it in their *local* palette.
            
            'If this is the *first* frame, we will use it as the basis of our *global* palette.
            If (i = 0) Then
            
                'Simply copy over the palette as-is into our running global palette tracker
                m_numColorsInGP = numColorsInLP
                ReDim m_globalPalette(0 To m_numColorsInGP - 1) As RGBQuad
                
                For idxPalette = 0 To m_numColorsInGP - 1
                    m_globalPalette(idxPalette) = imgPalette(idxPalette)
                Next idxPalette
                
                'Sort the palette by popularity (with a few tweaks), which can eke out slightly
                ' better compression ratios.  (Obviously we only have popularity data for the first
                ' frame, but in real-world usage this is a useful analog for the "average" frame
                ' that encodes using the global palette.)
                Palettes.SortPaletteForCompression_IncAlpha m_allFrames(i).frameDIB, m_globalPalette, True, False
                
                'The first frame always uses the global palette
                frameUsesGP = True
            
            'If this is *not* the first frame, and we have yet to write a global palette, append as many
            ' unique colors from this palette as we can into the global palette.
            Else
                
                'If there's still room in the global palette, append this palette to it.
                Dim gpTooBig As Boolean: gpTooBig = False
                If (m_numColorsInGP < 256) Then
                    
                    m_numColorsInGP = Palettes.MergePalettes(m_globalPalette, m_numColorsInGP, imgPalette, numColorsInLP)
                    
                    'Enforce a strict 256-color limit; colors past the end will simply be discarded, and this frame
                    ' will use a local palette instead.
                    If (m_numColorsInGP > 256) Then
                        gpTooBig = True
                        m_numColorsInGP = 256
                        ReDim Preserve m_globalPalette(0 To 255) As RGBQuad
                    End If
                    
                End If
                
                'Next, we need to see if all colors in this frame appear in the global palette.
                ' If they do, we can simply use the global palette to write this frame.
                ' (Note that we can automatically skip this step if the previous merge produced
                ' a too-big palette.)
                If (Not gpTooBig) Then
                    frameUsesGP = Palettes.DoesPaletteContainPalette(m_globalPalette, m_numColorsInGP, imgPalette, numColorsInLP)
                Else
                    frameUsesGP = False
                End If
                
            End If
            
            m_allFrames(i).usesGlobalPalette = frameUsesGP
            
            'Frames that use the global palette will be handled later, in a separate pass.
            ' Local palettes can be processed immediately, however, and we can free some
            ' of their larger structs (like their 32-bpp frame copy) immediately after.
            If m_allFrames(i).usesGlobalPalette Then
                
                'Suspend the DIB (e.g. compress it to a smaller memory stream) to reduce memory constraints
                m_allFrames(i).frameDIB.SuspendDIB cf_Lz4, False
                
            Else
                
                'If the current frame requires transparency, and it's using a local palette,
                ' ensure transparency exists in the palette.  (Global palettes will be handled
                ' in a separate loop, later.  We do this because the global palette may not
                ' have a transparent entry *now*, but because GIF color tables have to be
                ' padded to the nearest power-of-two, we may get transparency "for free" when
                ' we finalize the table.)
                Dim pEntry As Long
                If m_allFrames(i).frameNeedsTransparency Then
                    
                    Dim trnsFound As Boolean: trnsFound = False
                    For pEntry = 0 To UBound(imgPalette)
                        If (imgPalette(pEntry).Alpha = 0) Then
                            trnsFound = True
                            Exit For
                        End If
                    Next pEntry
                    
                    'If transparency *wasn't* found, add it manually (if there's room).
                    If (Not trnsFound) Then
                        
                        'There's room for another color in the palette!  Create a transparent
                        ' color and expand the palette accordingly.  (TODO: technically we may
                        ' not want to do this if there's already a power-of-two number of colors
                        ' in the palette.  The reason for this is that we'd have to bump-up
                        ' the entire color count to the *next* power-of-two, which may negate
                        ' any size gains we get from having access to a transparent pixel, argh.)
                        If (numColorsInLP < 256) Then
                            ReDim Preserve imgPalette(0 To numColorsInLP) As RGBQuad
                            imgPalette(numColorsInLP).Blue = 0
                            imgPalette(numColorsInLP).Green = 0
                            imgPalette(numColorsInLP).Red = 0
                            imgPalette(numColorsInLP).Alpha = 0
                            numColorsInLP = numColorsInLP + 1
                            
                        'Damn, the local palette is full, meaning we can't add a transparent color
                        ' without erasing an existing color.  Because this affects our ability to
                        ' losslessly save existing GIFs, dump our pixel-blank-optimized copy of
                        ' the frame and revert to the original, untouched version.
                        Else
                            Set m_allFrames(i).frameDIB = m_allFrames(i).backupFrameDIB
                            Set m_allFrames(i).backupFrameDIB = Nothing
                            m_allFrames(i).frameNeedsTransparency = False
                            m_allFrames(i).frameWasBlanked = False
                        End If
                        
                    End If
                    
                '/end frame needs transparency
                End If
                    
                'With all optimizations applied, we are finally ready to palettize this layer
                ' - again *IF* it uses a local palette.
                
                'In LZ77 encoding (e.g. DEFLATE), reordering palette can improve encoding efficiency
                ' because the sliding-window approach favors recent matches over past ones.  LZ78 is
                ' different this way because it's deterministic and code-agnostic, due to the way it
                ' precisely matches data against a fixed table (which is fully discarded and rebuilt
                ' when the table fills).  As such, we won't do a full sort - instead, we'll just do
                ' an "alpha" sort, which moves the transparent pixel (if any) to the front of the
                ' color table.
                Palettes.SortPaletteForCompression_IncAlpha m_allFrames(i).frameDIB, imgPalette, True, True
                
                'Transfer the final palette into the frame collection
                m_allFrames(i).palNumColors = numColorsInLP
                ReDim m_allFrames(i).framePalette(0 To numColorsInLP - 1)
                CopyMemoryStrict VarPtr(m_allFrames(i).framePalette(0)), VarPtr(imgPalette(0)), numColorsInLP * 4
                
                'One last optimization: if this frame received a pixel-blanking optimization pass,
                ' we want to compare compressibility of the blanked frame against the original frame.
                ' We can't be quite as aggressive with this pass (e.g. merging results from the two
                ' possible frames) because the two images may use wildly different palettes, unlike
                ' a global-palette frame where we are guaranteed that this frame consists only of
                ' shared colors.
                If m_allFrames(i).frameWasBlanked And (Not m_allFrames(i).backupFrameDIB Is Nothing) Then
                    
                    'Palettize the blanked frame (without dithering; the extra noise it introduces is
                    ' not helpful) into a temporary array, then generate a temporary palette and
                    ' palettize the backup frame.
                    DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(i).frameDIB, imgPalette, tmpBlanked
                    
                    Dim tmpFramePalette() As RGBQuad, tmpFramePaletteCount As Long
                    Palettes.GetOptimizedPaletteIncAlpha m_allFrames(i).backupFrameDIB, tmpFramePalette, 256, pdqs_Variance, True
                    tmpFramePaletteCount = UBound(tmpFramePalette) + 1
                    Palettes.EnsureBinaryAlphaPalette tmpFramePalette
                    Palettes.SortPaletteForCompression_IncAlpha m_allFrames(i).backupFrameDIB, tmpFramePalette, True, True
                    DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(i).backupFrameDIB, tmpFramePalette, tmpOriginal
                    
                    'We now have three choices of pixel streams:
                    ' 1) the original (untouched) frame
                    ' 2) a frame with duplicate pixels between (1) and the previous frame blanked out
                    
                    'We're now going to do a quick lz4 check of each stream and take whichever one
                    ' compresses the best.  This provides a very fast, reasonably good estimate of which
                    ' frame will compress the best under GIF's primitive LZW strategy.
                    netPixels = m_allFrames(i).frameDIB.GetDIBWidth * m_allFrames(i).frameDIB.GetDIBHeight
                    ReDim m_allFrames(i).pixelData(0 To m_allFrames(i).rectOfInterest.Width - 1, 0 To m_allFrames(i).rectOfInterest.Height - 1) As Byte
                    
                    initCmpSize = cmpTestBufferSize
                    initSize = netPixels
                    Compression.CompressPtrToPtr VarPtr(cmpTestBuffer(0)), initCmpSize, VarPtr(tmpOriginal(0, 0)), initSize, cf_Lz4
                    
                    testSize = cmpTestBufferSize
                    Compression.CompressPtrToPtr VarPtr(cmpTestBuffer(0)), testSize, VarPtr(tmpBlanked(0, 0)), initSize, cf_Lz4
                    
                    'If the frame-blanked copy compressed better...
                    If (testSize < initCmpSize) Then
                        
                        'Copy its pixel stream into the frame collection
                        CopyMemoryStrict VarPtr(m_allFrames(i).pixelData(0, 0)), VarPtr(tmpBlanked(0, 0)), netPixels
                        
                    'The original frame compressed better...
                    Else
                        
                        'Copy pixel *and* palette data into the frame collection
                        CopyMemoryStrict VarPtr(m_allFrames(i).pixelData(0, 0)), VarPtr(tmpOriginal(0, 0)), netPixels
                        m_allFrames(i).palNumColors = tmpFramePaletteCount
                        ReDim m_allFrames(i).framePalette(0 To tmpFramePaletteCount - 1)
                        CopyMemoryStrict VarPtr(m_allFrames(i).framePalette(0)), VarPtr(tmpFramePalette(0)), tmpFramePaletteCount * 4
                        
                    End If
                    
                    Erase tmpOriginal
                    Erase tmpBlanked
                    Erase tmpFramePalette
                    Set m_allFrames(i).backupFrameDIB = Nothing
                    
                Else
                    
                    'Palettize the image and cache the result in the frame collection.
                    ' (TODO: switch to neural-network quantizer.)
                    If useDithering Then
                        Palettes.GetPalettizedImage_Dithered_IncAlpha m_allFrames(i).frameDIB, imgPalette, m_allFrames(i).pixelData, PDDM_SierraLite, 0.67, True
                    Else
                        DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(i).frameDIB, imgPalette, m_allFrames(i).pixelData
                    End If
                    
                End If
                
                'We can now free the (temporary) frame DIB copy, because it is either unnecessary
                ' (because this frame isn't being encoded) or because it has been palettized and
                ' stored inside m_allFrames(i).pixelData.
                Set m_allFrames(i).frameDIB = Nothing
                
            '/end frame uses local palette
            End If
        
        '/end frame is duplicate or empty
        End If
        
    'Next frame
    Next i
    
    'Clear all optimization-related objects, as they are no longer needed
    Set prevFrame = Nothing
    Set bufferFrame = Nothing
    Set curFrameBackup = Nothing
    Set refFrame = Nothing
    Set testDIB = Nothing
    
    'Note: at this point, frames that rely on the global palette have *not* been optimized yet.
    ' This is because they may be able to use transparency, which wasn't guaranteed present
    ' at the time of their original processing (because the global palette hadn't filled up yet.)
    
    'So our next job is to get the global palette in order.
    
    'The GIF spec requires all palette color counts to be a power of 2.  (It does this because
    ' palette color count is stored in 3-bits, ugh.)  Any unused entries are ignored, but by
    ' convention are usually left as black; we do the same here.
    m_numColorsInGP = 2 ^ Pow2FromColorCount(m_numColorsInGP)
    If (UBound(m_globalPalette) <> m_numColorsInGP - 1) Then ReDim Preserve m_globalPalette(0 To m_numColorsInGP - 1) As RGBQuad
    
    'If the global palette has a transparent index, locate it and ensure it is in position 0.
    ' (While not required by the spec, this *is* required by the PNG spec, and it generally
    ' improves compression to set it early in the table, given where it's likely to be
    ' encountered in real-world images.)
    m_GlobalTrnsIndex = -1
    For i = 0 To m_numColorsInGP - 1
        If (m_globalPalette(i).Alpha = 0) Then
            
            If (i > 0) Then
                
                'Shift all previous colors backward, then plug this transparent pixel into
                ' the *first* palette position .
                Dim tmpColor As RGBQuad
                tmpColor = m_globalPalette(i)
                For j = i To 1 Step -1
                    m_globalPalette(j) = m_globalPalette(j - 1)
                Next j
                
                m_globalPalette(0) = tmpColor
                
            End If
            
            'Once a single transparent color has been located, we can quit searching.  (There may
            ' be more transparent pixels, on account of "filler" entries we had to add to pad
            ' out the color table to a power-of-two, but GIFs don't actually encode alpha data -
            ' they only allow a single flag for marking a transparent index, so any remaining
            ' transparent pixels will just end up as opaque black in the final color table.)
            m_GlobalTrnsIndex = 0
            Exit For
            
        End If
    Next i
    
    'With the global palette finalized, we can now do a final loop through all frames to palettize
    ' any frames that rely on the global palette.
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        
        'Optimizing frames can take some time.  Keep the user apprised of our progress.
        ProgressBars.SetProgBarVal srcPDImage.GetNumOfLayers + i
        Message "Saving animation frame %1 of %2...", i + 1, srcPDImage.GetNumOfLayers + 1, "DONOTLOG"
        
        'The only frames we care about in this pass are non-empty, non-duplicate frames that rely
        ' on the global color table.
        If (Not m_allFrames(i).frameIsDuplicateOrEmpty) And m_allFrames(i).usesGlobalPalette Then
            
            'Basically, repeat the same steps we did with local palette frames, above.
            
            'If this frame requires transparency, see if the global palette provides such a thing.
            If m_allFrames(i).frameNeedsTransparency Then
            
                'Damn, global palette does *not* have transparency support.  Roll back to the non-blanked
                ' version of this frame.
                If (m_GlobalTrnsIndex < 0) Then
                    Set m_allFrames(i).frameDIB = m_allFrames(i).backupFrameDIB
                    Set m_allFrames(i).backupFrameDIB = Nothing
                    m_allFrames(i).frameNeedsTransparency = False
                    m_allFrames(i).frameWasBlanked = False
                End If
                
            End If
            
            'One last optimization: if this frame received a pixel-blanking optimization pass,
            ' we now want to generate a "merged" frame that combines the most-compressible
            ' scanlines from both the blanked frame and the original frame.  This produces a
            ' "best of both worlds" result that compresses better than either frame alone.
            If m_allFrames(i).frameWasBlanked And (Not m_allFrames(i).backupFrameDIB Is Nothing) Then
            
                'Palettize both frames (without dithering; the extra noise it introduces is
                ' not helpful) into temporary arrays.
                DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(i).frameDIB, m_globalPalette, tmpBlanked
                DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(i).backupFrameDIB, m_globalPalette, tmpOriginal
                
                'From these, build a new palettized image that uses the most compressible
                ' scanlines from each.
                DIBs.MakeMinimalEntropyScanlines tmpOriginal, tmpBlanked, m_allFrames(i).frameDIB.GetDIBWidth, m_allFrames(i).frameDIB.GetDIBHeight, tmpMerged
                
                'We now have three choices of pixel streams:
                ' 1) the original (untouched) frame
                ' 2) a frame with duplicate pixels between (1) and the previous frame blanked out
                ' 3) a frame that attempts to combine the best scanlines from (1) and (2) into a single stream.
                
                'We're now going to do a quick compression check of each stream and take whichever one
                ' compresses the best.  This is the closest thing we have to a "foolproof" strategy.
                netPixels = m_allFrames(i).frameDIB.GetDIBWidth * m_allFrames(i).frameDIB.GetDIBHeight
                ReDim m_allFrames(i).pixelData(0 To m_allFrames(i).frameDIB.GetDIBWidth - 1, 0 To m_allFrames(i).frameDIB.GetDIBHeight - 1) As Byte
                
                initCmpSize = cmpTestBufferSize
                initSize = netPixels
                Compression.CompressPtrToPtr VarPtr(cmpTestBuffer(0)), initCmpSize, VarPtr(tmpOriginal(0, 0)), initSize, cf_Lz4
                
                testSize = cmpTestBufferSize
                Compression.CompressPtrToPtr VarPtr(cmpTestBuffer(0)), testSize, VarPtr(tmpBlanked(0, 0)), initSize, cf_Lz4
                
                'Compare the smaller of the previous test to our minimal-entropy attempt
                If (initCmpSize < testSize) Then

                    testSize = cmpTestBufferSize
                    Compression.CompressPtrToPtr VarPtr(cmpTestBuffer(0)), testSize, VarPtr(tmpMerged(0, 0)), initSize, cf_Lz4

                    'Copy the result into the frame collection, then free all DIBs
                    If (initCmpSize < testSize) Then
                        CopyMemoryStrict VarPtr(m_allFrames(i).pixelData(0, 0)), VarPtr(tmpOriginal(0, 0)), netPixels
                    Else
                        CopyMemoryStrict VarPtr(m_allFrames(i).pixelData(0, 0)), VarPtr(tmpMerged(0, 0)), netPixels
                    End If

                'Same thing, but comparing the other array
                Else

                    initCmpSize = cmpTestBufferSize
                    Compression.CompressPtrToPtr VarPtr(cmpTestBuffer(0)), initCmpSize, VarPtr(tmpMerged(0, 0)), initSize, cf_Lz4

                    If (initCmpSize < testSize) Then
                        CopyMemoryStrict VarPtr(m_allFrames(i).pixelData(0, 0)), VarPtr(tmpMerged(0, 0)), netPixels
                    Else
                        CopyMemoryStrict VarPtr(m_allFrames(i).pixelData(0, 0)), VarPtr(tmpBlanked(0, 0)), netPixels
                    End If

                End If
                
                Erase tmpOriginal
                Erase tmpBlanked
                Set m_allFrames(i).frameDIB = Nothing
                Set m_allFrames(i).backupFrameDIB = Nothing
                    
            Else
                
                'Transparency has been dealt with (and rolled back, as necessary).  All that's left to do
                ' is palettize this frame against the finished global palette!  TODO: switch to neural network quantizer.
                If useDithering Then
                    Palettes.GetPalettizedImage_Dithered_IncAlpha m_allFrames(i).frameDIB, m_globalPalette, m_allFrames(i).pixelData, PDDM_SierraLite, 0.67, True
                Else
                    DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(i).frameDIB, m_globalPalette, m_allFrames(i).pixelData
                End If
                
            End If
            
            'Free the 32-bpp frame DIB (it's no longer required)
            Set m_allFrames(i).frameDIB = Nothing
            
        End If
        
    Next i
    
    'Compression test buffer is no longer required
    Erase cmpTestBuffer
    
    Message "Finalizing image..."
    ProgressBars.SetProgBarVal ProgressBars.GetProgBarMax()
    
    'We've successfully optimized all GIF frames.  Now it's time to involve the encoder.
    ' PD can use FreeImage or our own native-VB6 encoder.  Our internal encoder uses less memory
    ' and is slightly faster, while also constructing smaller GIFs (FreeImage always writes 8-bit
    ' color tables even if the palette could use fewer bits), but our internal engine obviously
    ' hasn't received the same widespread testing as FreeImage's.  For now, I consider the
    ' advantages of our own encoder worthwhile, with the caveat that additional testing may
    ' change my decision.
    Const USE_FREEIMAGE_FOR_GIFS As Boolean = False
    If (USE_FREEIMAGE_FOR_GIFS And ImageFormats.IsFreeImageEnabled()) Then
        ExportGIF_Animated_LL = WriteOptimizedAGIFToFile_FI(srcPDImage, dstFile, formatParams, metadataParams)
    Else
        ExportGIF_Animated_LL = WriteOptimizedAGIFToFile_Internal(srcPDImage, dstFile, formatParams, metadataParams)
    End If
    
    'Manually erase any module-level containers (don't want them to live on after this!)
    Erase m_globalPalette
    Erase m_allFrames
    
    ProgressBars.SetProgBarVal 0
    ProgressBars.ReleaseProgressBar
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error encountered in " & GIF_FILE_EXTENSION & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportGIF_Animated_LL = False
    
End Function

'After optimizing GIF frames (which you must do to generate the structures used by this function),
' call this function to actually write GIF data out to file.  Native VB6 functions will be used.
Private Function WriteOptimizedAGIFToFile_Internal(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean

    On Error GoTo ExportGIFError
    
    WriteOptimizedAGIFToFile_Internal = False
    
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
    
    'As always, pdStream handles actual writing duties.  (Memory mapping is used for ideal performance.)
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, tmpFilename, optimizeAccess:=OptimizeSequentialAccess) Then
    
        'For detailed GIF format info, see http://giflib.sourceforge.net/whatsinagif/bits_and_bytes.html
        ' PD doesn't attempt to support every esoteric GIF option (for example, we deliberately omit support
        ' for interlaced GIFs because they're beyond pointless in the 21st century).  Instead, we focus on
        ' minimum filesize and maximum encoding efficiency.
        
        'GIF header is fixed, 3-bytes for "GIF" ID, 3-bytes for version (always "89a" for animation)
        cStream.WriteString_ASCII "GIF89a"
        
        'Next, the "logical screen descriptor".  This is always 7 bytes long:
        ' 4 bytes - unsigned short width + height
        cStream.WriteIntU srcPDImage.Width
        cStream.WriteIntU srcPDImage.Height
        
        'Now, an unpleasant packed 8-bit field
        ' 1 bit - global color table GCT exists (always TRUE in PD)
        ' 3 bits - GCT size N (describing 2 ^ n-1 colors in the palette)
        ' 1 bit - palette is sorted by importance (no longer used, always 0 from PD even though PD produces sorted palettes just fine)
        ' 3 bits - GCT size N again (technically the first field is bit-depth, but they're the same when using a global palette)
        Dim tmpBitField As Byte
        tmpBitField = &H80 'global palette exists
        
        Dim pow2forGP As Long
        pow2forGP = Pow2FromColorCount(m_numColorsInGP) - 1
        tmpBitField = tmpBitField Or (pow2forGP * &H10) Or pow2forGP
        cStream.WriteByte tmpBitField
        
        'Background color index, always 0 by PD
        cStream.WriteByte 0
        
        'Aspect ratio using a bizarre old formula, always 0 by PD
        cStream.WriteByte 0
        
        'Next comes the global color table/palette, in RGB order
        Dim i As Long, j As Long
        For i = 0 To m_numColorsInGP - 1
            With m_globalPalette(i)
                cStream.WriteByte .Red
                cStream.WriteByte .Green
                cStream.WriteByte .Blue
            End With
        Next i
        
        'The image header is now complete.
        
        'For animated images, we now need to place a custom application extension (first used by Netscape)
        ' to declare the loop behavior of the GIF. See https://en.wikipedia.org/wiki/GIF#Animated_GIF for details.
        ' (There will be various magic numbers used here; the wiki page shares more details on them.)
        ' Note that un-looping GIFs can skip this entirely.
        Dim loopCount As Long
        loopCount = cParams.GetLong("animation-loop-count", 1)
        If (loopCount > 65536) Then loopCount = 65536
                        
        If (loopCount <> 1) Then
            
            'Application extension ID
            cStream.WriteByte &H21
            cStream.WriteByte &HFF
            
            'Size of block (always 11 for this block)
            cStream.WriteByte 11
            
            'Application name + 3 verification bytes
            cStream.WriteString_ASCII "NETSCAPE2.0"
            
            'Number of bytes in the following sub-block (always 3)
            cStream.WriteByte 3
            
            'Sub-block index (always 1)
            cStream.WriteByte 1
            
            'Number of repetitions - 1
            If (loopCount < 1) Then loopCount = 1
            loopCount = loopCount - 1
            cStream.WriteIntU loopCount
            
            'End of the sub-block chain
            cStream.WriteByte 0
            
        End If
        
        'Time to iterate and store frames.
        For i = 0 To srcPDImage.GetNumOfLayers - 1
        
            'Skip duplicate or empty frames
            If (Not m_allFrames(i).frameIsDuplicateOrEmpty) Then
            
                'All frames are preceded by a "Graphics Control Extension".
                ' This is a fixed-size struct describing things like frame delay, transparency presence, etc.
                
                'First three bytes are fixed ("introducer", "label", size)
                cStream.WriteByte &H21
                cStream.WriteByte &HF9
                cStream.WriteByte &H4
                
                'Next is an annoying packed field:
                ' - 3 bits reserved (0)
                ' - 3 bits disposal method
                ' - 1 bit user-input flag (ignored)
                ' - 1 bit transparent color flag
                tmpBitField = 0
                tmpBitField = tmpBitField Or (m_allFrames(i).frameDisposal * &H4&)
                
                Dim frameUsesAlpha As Boolean
                If m_allFrames(i).usesGlobalPalette Then
                    frameUsesAlpha = (m_GlobalTrnsIndex >= 0)
                Else
                    frameUsesAlpha = (m_allFrames(i).framePalette(0).Alpha = 0)
                End If
                If frameUsesAlpha Then tmpBitField = tmpBitField Or 1
                
                cStream.WriteByte tmpBitField
                
                'Next is 2-byte delay time, in centiseconds
                Dim finalFrameTime As Long
                finalFrameTime = Int((m_allFrames(i).frameTime + 5) \ 10)
                If (finalFrameTime > 65535) Then finalFrameTime = 65535
                cStream.WriteIntU finalFrameTime
                
                'Next is 1-byte transparent color index (always 0 in PD, but I've seen a convention in other software
                ' to write this as 0xff when the frame isn't using alpha at all... not sure if that matters, but why not?)
                If frameUsesAlpha Then
                    If m_allFrames(i).usesGlobalPalette Then
                        cStream.WriteByte m_GlobalTrnsIndex
                    Else
                        cStream.WriteByte 0
                    End If
                Else
                    cStream.WriteByte &HFF
                End If
                
                'Next is 1-byte block terminator (always 0)
                cStream.WriteByte 0
                
                'Graphics Control Extension is done
                
                'Next up is an "image descriptor", basically a frame header
                
                '1-byte image separator (always 2C)
                cStream.WriteByte &H2C
                
                'Frame dimensions as unsigned shorts, in left/top/width/height order
                cStream.WriteIntU Int(m_allFrames(i).rectOfInterest.Left)
                cStream.WriteIntU Int(m_allFrames(i).rectOfInterest.Top)
                cStream.WriteIntU Int(m_allFrames(i).rectOfInterest.Width)
                cStream.WriteIntU Int(m_allFrames(i).rectOfInterest.Height)
                
                'And my favorite, another packed bit-field!  (uuuuugh)
                tmpBitField = 0
                ' - 1 bit local palette used (varies by frame)
                ' - 1 bit interlaced (PD never interlaces frames)
                ' - 1 bit sort flag (same as global table, PD can - and may - do this, but always writes 0 per giflib convention)
                ' - 2 bits reserved
                ' - 3 bits size of local color table N (describing 2 ^ n-1 colors in the palette)
                Dim pow2forLP As Long
                If (Not m_allFrames(i).usesGlobalPalette) Then
                    tmpBitField = tmpBitField Or &H80
                    pow2forLP = Pow2FromColorCount(m_allFrames(i).palNumColors) - 1
                    tmpBitField = tmpBitField Or pow2forLP
                End If
                cStream.WriteByte tmpBitField
                
                'There is no terminator here.  Instead, if a local palette is in use, we immediately write it
                If (Not m_allFrames(i).usesGlobalPalette) Then
                    
                    'Ensure the local palette is a fixed power-of-2 size (we can reuse the calculation
                    ' from the bit-field, above)
                    Dim newPaletteSize As Long
                    newPaletteSize = 2 ^ (pow2forLP + 1)
                    If (m_allFrames(i).palNumColors < newPaletteSize) Then
                        m_allFrames(i).palNumColors = newPaletteSize
                        ReDim Preserve m_allFrames(i).framePalette(0 To newPaletteSize - 1) As RGBQuad
                    End If
                    
                    'Dump the palette to file, while swizzling to RGB order
                    For j = 0 To newPaletteSize - 1
                        With m_allFrames(i).framePalette(j)
                            cStream.WriteByte .Red
                            cStream.WriteByte .Green
                            cStream.WriteByte .Blue
                        End With
                    Next j
                    
                End If
                
                'All that's left are the pixel bits.  These are prefaced by a byte describing the
                ' minimum LZW code size.  This is a minimum of 2, a maximum of the power-of-2 size
                ' of the frame's palette (global or local).
                Dim lzwCodeSize As Long
                If m_allFrames(i).usesGlobalPalette Then
                    lzwCodeSize = pow2forGP + 1
                Else
                    lzwCodeSize = pow2forLP + 1
                End If
                If (lzwCodeSize < 2) Then lzwCodeSize = 2
                cStream.WriteByte lzwCodeSize
                
                'Next is the image bitstream!  Encoding happens elsewhere; we just pass the stream to them
                ' and let them encode away.
                ImageFormats_GIF_LZW.CompressLZW cStream, VarPtr(m_allFrames(i).pixelData(0, 0)), m_allFrames(i).rectOfInterest.Width * m_allFrames(i).rectOfInterest.Height, lzwCodeSize + 1
                
                'All that's left for this frame is to explicitly terminate the black
                cStream.WriteByte 0
                
            '/end duplicate/empty frames
            End If
            
        'Continue with the next frame!
        Next i
        
        'With all frames written, we can write the trailer and exit!
        ' (This is a magic number from the spec: https://www.w3.org/Graphics/GIF/spec-gif89a.txt)
        cStream.WriteByte &H3B
        
        cStream.StopStream True
        
        'Work complete
        WriteOptimizedAGIFToFile_Internal = True
        
    Else
        WriteOptimizedAGIFToFile_Internal = False
        ExportDebugMsg "failed to open pdStream"
    End If
    
    'If we wrote our data to a temp file, attempt to replace the original file
    If Strings.StringsNotEqual(dstFile, tmpFilename) Then
        
        WriteOptimizedAGIFToFile_Internal = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
        
        If (Not WriteOptimizedAGIFToFile_Internal) Then
            Files.FileDelete tmpFilename
            PDDebug.LogAction "WARNING!  ImageExporter could not overwrite GIF file; original file is likely open elsewhere."
        End If
        
    End If
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error encountered in " & GIF_FILE_EXTENSION & " routine.  Err #" & Err.Number & ", " & Err.Description
    WriteOptimizedAGIFToFile_Internal = False
    
End Function

'After optimizing GIF frames (which you must do to generate the structures used by this function),
' call this function to actually write GIF data out to file.  In this build, FreeImage is used.
' This may change pending further testing.
Private Function WriteOptimizedAGIFToFile_FI(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportGIFError
    
    WriteOptimizedAGIFToFile_FI = False
    
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
        
        'Finally, we can close the multipage handle "once and for all"; FreeImage handles the rest from here
        WriteOptimizedAGIFToFile_FI = FreeImage_CloseMultiBitmap(fi_MasterHandle)
        
        'If we wrote our data to a temp file, attempt to replace the original file
        If Strings.StringsNotEqual(dstFile, tmpFilename) Then
            
            WriteOptimizedAGIFToFile_FI = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
            
            If (Not WriteOptimizedAGIFToFile_FI) Then
                Files.FileDelete tmpFilename
                PDDebug.LogAction "WARNING!  ImageExporter could not overwrite GIF file; original file is likely open elsewhere."
            End If
            
        End If
        
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", GIF_FILE_EXTENSION
        WriteOptimizedAGIFToFile_FI = False
    End If
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error encountered in " & GIF_FILE_EXTENSION & " routine.  Err #" & Err.Number & ", " & Err.Description
    WriteOptimizedAGIFToFile_FI = False
    
End Function

Private Function Pow2FromColorCount(ByVal cCount As Long) As Long
    Pow2FromColorCount = 1
    Do While ((2 ^ Pow2FromColorCount) < cCount)
        Pow2FromColorCount = Pow2FromColorCount + 1
    Loop
End Function

Private Sub ExportDebugMsg(ByRef debugMsg As String)
    PDDebug.LogAction debugMsg
End Sub
