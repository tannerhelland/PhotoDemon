VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'GIF encoding library
'Copyright 2001-2026 by Tanner Helland
'Created: 4/15/01
'Last updated: 27/March/22
'Last update: add optional palette sorter
'
'Most image exporters exist in the ImageExporter module.  GIF is a weird exception
' because animated GIFs require a ton of preprocessing (to optimize animation frames),
' so I've moved them to their own home.
'
'PhotoDemon automatically optimizes saved GIFs to produce the smallest possible files.
' A variety of optimizations are used, and the encoder tests various strategies to try
' and choose the "best" (smallest) solution on each frame.  As you can see from the
' size of this module, many many many different optimizations are attempted.
'
'Despite this, the optimization pre-pass is reasonably quick, and the animated GIFs
' produced this way are often many times smaller than GIFs produced by a naive encoder.
'
'Note that the optimization steps are specifically written in an export library
' agnostic way.  PD internally stores the results of all optimizations, then just hands
' the optimized frames off to an encoder at the end of the process.  Historically PD used
' FreeImage for the actual GIF encoding step, but FreeImage has a number of shortcomings
' (including woeful performance and writing larger GIFs than is necessary), so in 2021
' I moved to an in-house LZW encoder based off the classic UNIX "compress" tool.
' This new LZW encoder has different licensing considerations, so I've kept it in a
' separate file (ImageFormats_GIF_LZW) - look there for additional credits and licensing
' details for that portion of the library.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Compile-time triggers:

'Sort global palette by "importance".  If sorted nicely, we can check the "sort flag" in the global palette header.
' This is mostly a historical curiosity, but I did write a sort algorithm that works well, so you can activate
' if curious.  Note that this imposes an encoding-time perf penalty, but hypothetically decoding performance will
' improve (due to better color locality).
Private Const SORT_PALETTE_BY_IMPORTANCE As Boolean = False

'Static GIF values:

'Copy of the 8-bpp (palettized) array representing the original image.  May be reduced further
' during LZW compression.
Private m_ImgBytes() As Byte

'Final image palette.  Alpha element is important, as there may be a 0-alpha entry for GIF transparency.
Private m_ImgPalette() As RGBQuad

'Animated GIF values:

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

'PD performs compression tests on potential frames to try and see which ones produce the smallest
' compressed size.  LZW is fast-ish but lz4 is much faster, so we use it as a rough predictor of
' compressibility.  (Compression tests use a shared buffer for improved performance.)
Private m_cmpTestBuffer() As Byte, m_cmpTestBufferSize As Long

'Given a param string generated by the Export GIF dialog, apply any GIF pre-processing steps.
' Works for both export preview and actual export prep steps (depending on the value of usePreviewMode).
' - In preview mode, the palette will be applied to the source DIB so you can see it.
' - In non-preview mode, the palette (and palettized image bytes) will be cached in module-level
'   variables so the encoder can use them as-is.
Friend Function GetGifReadyImage(ByRef srcDIB As pdDIB, Optional ByVal formatParams As String = vbNullString, Optional ByVal usePreviewMode As Boolean = False) As Boolean

    'Parse all relevant GIF parameters.
    ' (The GIF export dialog provides more details on how these parameters are generated.)
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
    
    'We now need to produce an image that meets GIF "criteria" - e.g. 8-bit colors with binary transparency.
    ' Start by matting the GIF against the supplied background color, using a strategy appropriate to
    ' whatever transparency method they requested.
    Dim trnsValues() As Byte
    
    'No alpha in the final image
    If (desiredAlphaStatus = PDAS_NoAlpha) Then
        srcDIB.CompositeBackgroundColor Colors.ExtractRed(gifBackgroundColor), Colors.ExtractGreen(gifBackgroundColor), Colors.ExtractBlue(gifBackgroundColor)
    
    'Make the chosen color transparent
    ElseIf (desiredAlphaStatus = PDAS_NewAlphaFromColor) Then
        DIBs.MakeColorTransparent_Ex srcDIB, trnsValues, gifAlphaCutoff
        DIBs.ApplyAlphaCutoff_Gif srcDIB, 127, gifBackgroundColor
        
    'Normal GIF behavior (threshold alpha into "fully transparent" or "fully opaque")
    Else
        DIBs.ApplyAlphaCutoff_Gif srcDIB, gifAlphaCutoff, gifBackgroundColor
    End If
    
    'Alpha is now guaranteed to be only values of 0 or 255.
    If (Not usePreviewMode) Then ProgressBars.SetProgBarVal 2
    
    'If the caller requested grayscale, apply that now.
    If gifForceGrayscale Then DIBs.MakeDIBGrayscale srcDIB, gifColorCount, False
    
    'All that's left to do is palettize the image!  For full-color images, let's use a fast algorithm.
    ' For smaller color counts, a neural network will produce a much better selection of colors
    ' (at a potentially significant cost to performance).
    Dim curColorCount As Long
    curColorCount = Palettes.GetDIBColorCount_FastAbort(srcDIB, m_ImgPalette)
    
    'In preview mode, always use the fast algorithm
    If usePreviewMode Then
        Palettes.GetOptimizedPaletteIncAlpha srcDIB, m_ImgPalette, gifColorCount, pdqs_Variance, True
    Else
        
        'In regular mode, we have a different choice to make.  First, see if the palette is already
        ' a useable size.  (This is likely for e.g. a loaded GIF being saved back out to GIF.)
        If (curColorCount > gifColorCount) Then
            
            'This image has too many colors and needs to be palettized.  For 256-colors, use the
            ' fastest available algorithm (modified median cut).
            If (gifColorCount = 256) Then
                Palettes.GetOptimizedPaletteIncAlpha srcDIB, m_ImgPalette, gifColorCount, pdqs_Variance, True
            
            'For lower color counts, use our modified Neuquant for much better quality.
            Else
                Palettes.GetNeuquantPalette_RGBA srcDIB, m_ImgPalette, gifColorCount, True
            End If
            
        '/no Else required; the palette returned by the color count function is useable as-is!
        End If
    
    End If
    
    If (Not usePreviewMode) Then ProgressBars.SetProgBarVal 3
    
    'We now have an optimized palette for this image.  If this is for export purposes,
    ' produce an 8-bpp array for export to file.  If this is a preview, apply the palette
    ' to the source DIB so the user can review it.
    If usePreviewMode Then
        Palettes.ApplyPaletteToImage_KDTree srcDIB, m_ImgPalette, True
    
    'During preview, palette order doesn't matter, but at export-time we want to sort the palette so
    ' that the transparent index appears in slot 0.
    Else
        Palettes.SortPaletteForCompression_IncAlpha srcDIB, m_ImgPalette, True, True
        DIBs.GetDIBAs8bpp_RGBA_SrcPalette srcDIB, m_ImgPalette, m_ImgBytes
        ProgressBars.SetProgBarVal 4
    End If
    
    GetGifReadyImage = True
    
End Function

'Save a static (non-animated) pdImage object to a pdStream.  This allows for saving to file, memory, etc.
Friend Function SaveGIF_ToStream_Static(ByRef srcPDImage As pdImage, ByRef dstStream As pdStream, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    Const FUNC_NAME As String = "SaveGIF_ToStream_Static"
    
    SaveGIF_ToStream_Static = False
    
    'Failsafe checks for input params
    If (srcPDImage Is Nothing) Or (dstStream Is Nothing) Then
        InternalError FUNC_NAME, "null inputs"
        Exit Function
    End If
    
    'Parameters are available for parsing, although it's expected that most parameters will
    ' only be useful to the pre-processor.
    ' (The GIF export dialog provides more details on how these parameters are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    'Raise a progress bar
    ProgressBars.SetProgBarMax 6
    ProgressBars.SetProgBarVal 0
    
    'Generate a composited image copy, with alpha automatically un-premultiplied
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, False
    
    ProgressBars.SetProgBarVal 1
    
    'Hand the image off to GetGifReadyImage(), which will pre-process the image according to
    ' whatever settings the user supplied in the export dialog.
    If (Not GetGifReadyImage(tmpImageCopy, formatParams, False)) Then
        InternalError FUNC_NAME, "pre-processing failed"
    End If
    
    'We no longer need the 32-bpp copy of the image; free it to conserve memory
    Set tmpImageCopy = Nothing
    
    'Cache the palette size so we don't have to keep querying UBound of the palette array
    Dim numColorsInPalette As Long
    numColorsInPalette = UBound(m_ImgPalette) + 1
    
    'Filsafe check only
    If (numColorsInPalette < 1) Then numColorsInPalette = 1
    If (numColorsInPalette > 256) Then numColorsInPalette = 256
    
    'Background color index; if the first entry in the palette is transparency, we can override its
    ' color with the background color (because that color will basically be ignored as it's transparent).
    ' If the image does *not* contain transparency, however, we'll attempt to find the nearest match to
    ' the matte color in the existing palette.  We do this because there's no guarantee the matte color
    ' will be used in the image - it may have just been used to composite semi-transparent pixels -
    ' but this at least gives us something "close".
    Dim bkgdIndex As Long, bkgdColor As Long
    If (m_ImgPalette(0).Alpha = 0) Then
        
        'Overwrite the first palette entry's color with the background color
        bkgdColor = cParams.GetLong("gif-backcolor", vbWhite)
        With m_ImgPalette(0)
            .Red = Colors.ExtractRed(bkgdColor)
            .Green = Colors.ExtractGreen(bkgdColor)
            .Blue = Colors.ExtractBlue(bkgdColor)
        End With
        
    Else
        
        'Stuff the color at the end of the palette if there's room
        If (numColorsInPalette < (2 ^ Pow2FromColorCount(numColorsInPalette))) Then
            
            'See if the color already exists in the palette
            bkgdColor = cParams.GetLong("gif-backcolor", vbWhite)
            bkgdIndex = Palettes.GetNearestIndexRGB(m_ImgPalette, bkgdColor)
            If (RGB(m_ImgPalette(bkgdIndex).Red, m_ImgPalette(bkgdIndex).Green, m_ImgPalette(bkgdIndex).Blue) <> bkgdColor) Then
            
                'Background color doesn't exist in the palette.  Add it.
                If (numColorsInPalette > UBound(m_ImgPalette)) Then ReDim Preserve m_ImgPalette(0 To numColorsInPalette) As RGBQuad
                With m_ImgPalette(numColorsInPalette)
                    .Red = Colors.ExtractRed(bkgdColor)
                    .Green = Colors.ExtractGreen(bkgdColor)
                    .Blue = Colors.ExtractBlue(bkgdColor)
                End With
                
                'Set the background index to the new entry, and increment the total color count
                bkgdIndex = numColorsInPalette
                numColorsInPalette = numColorsInPalette + 1
            
            '/no Else required - the background color already exists in the palette
            End If
            
        'There's no room; find the nearest color instead and use *that* index
        Else
            bkgdIndex = Palettes.GetNearestIndexRGB(m_ImgPalette, cParams.GetLong("gif-backcolor", vbWhite))
        End If
        
    End If
    
    'For detailed GIF format info, see http://giflib.sourceforge.net/whatsinagif/bits_and_bytes.html
    
    'Start with the file header
    SaveGIF_ToStream_Static = WriteGIF_FileHeader(dstStream, srcPDImage.Width, srcPDImage.Height, numColorsInPalette, m_ImgPalette, bkgdIndex)
    
    'The image header is now complete.
    
    'If the image contains transparency, we now need to write an optional block to flag the
    ' transparent palette index.
    If (m_ImgPalette(0).Alpha = 0) Then
        
        'The transparent color may have been overridden in a previous step.  Replace it with transparent black now,
        ' to ensure future color-matching works correctly.
        m_ImgPalette(0).Red = 0
        m_ImgPalette(0).Green = 0
        m_ImgPalette(0).Blue = 0
        
        'A separate "Graphics Control Extension" writer handles this step for us.
        SaveGIF_ToStream_Static = WriteGIF_GCE(dstStream, True, trnsIndex:=0)
        
    End If
    
    'Next up is an "image descriptor", basically a frame header.
    
    '1-byte image separator (always 2C)
    dstStream.WriteByte &H2C
    
    'Frame dimensions as unsigned shorts, in left/top/width/height order
    dstStream.WriteIntU 0
    dstStream.WriteIntU 0
    dstStream.WriteIntU srcPDImage.Width
    dstStream.WriteIntU srcPDImage.Height
    
    'And my favorite, another packed bit-field!  (uuuuugh)
    ' - 1 bit local palette used (always 0 for static images)
    ' - 1 bit interlaced (always 0, PD never interlaces frames)
    ' - 1 bit sort flag (same as global table, PD can - and may - do this, but always writes 0 per giflib convention)
    ' - 2 bits reserved
    ' - 3 bits size of local color table N (describing 2 ^ n-1 colors in the palette, always 0 for PD)
    dstStream.WriteByte 0
    
    'All that's left are the pixel bits.  These are prefaced by a byte describing the
    ' minimum LZW code size.  This is a minimum of 2, a maximum of the power-of-2 size
    ' of the frame's palette (global or local).
    Dim lzwCodeSize As Long
    lzwCodeSize = Pow2FromColorCount(numColorsInPalette)
    If (lzwCodeSize < 2) Then lzwCodeSize = 2
    dstStream.WriteByte lzwCodeSize
    
    'Next is the image bitstream!  Encoding happens elsewhere; we just pass the stream to them
    ' and let them encode away.
    ProgressBars.SetProgBarVal 5
    ImageFormats_GIF_LZW.CompressLZW dstStream, VarPtr(m_ImgBytes(0, 0)), srcPDImage.Width * srcPDImage.Height, lzwCodeSize + 1, True
    ProgressBars.SetProgBarVal 6
    
    'All that's left for this frame is to explicitly terminate the block
    dstStream.WriteByte 0
    
    'With all frames written, we can write the trailer and exit!
    ' (This is a magic number from the spec: https://www.w3.org/Graphics/GIF/spec-gif89a.txt)
    dstStream.WriteByte &H3B
    
    'Normally we would stop the stream here, but we leave it to the caller instead as they may
    ' also want to free memory (if this write went out to disk).
    SaveGIF_ToStream_Static = True

End Function

'Save an animated (multi-layer) pdImage object to a pdStream.  This allows for saving to file, memory, etc.
Friend Function SaveGIF_ToStream_Animated(ByRef srcPDImage As pdImage, ByRef dstStream As pdStream, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    Const FUNC_NAME As String = "SaveGIF_ToStream_Animated"
    
    SaveGIF_ToStream_Animated = False
    
    'Failsafe checks for input params
    If (srcPDImage Is Nothing) Or (dstStream Is Nothing) Then
        InternalError FUNC_NAME, "null inputs"
        Exit Function
    End If
    
    'Parameters are available for parsing, although it's expected that most parameters will
    ' only be useful to the pre-processor.
    '
    '(The GIF export dialog provides more details on how these parameters are generated.)
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
    
    'We also use a soft reference that we can point at whatever frame DIB we want
    ' (its contents are not maintained between frames).
    Dim refFrame As pdDIB
    
    'Some parts of the optimization process are not guaranteed to improve file size
    ' (e.g. blanking out duplicate pixels between frames can actually hurt compression
    ' if the current frame is very noisy). To ensure we only apply beneficial optimizations,
    ' we test-compress the current frame after potentially problematic optimizations to
    ' double-check that our compression ratio improved.  (If it didn't, we'll roll back the
    ' changes we made - see the multiple DIB copies above!)
    '
    'Note that for performance reasons, we use lz4 instead of our native GIF encoder.
    ' lz4 is a hell of a lot faster, and I assume that the best-case result with lz4
    ' correlates reasonably well with the best-case result for a GIF-style LZW
    ' compressor, since LZ77 and LZ78 compression share most critical aspects on
    ' image-style data.
    '
    'To reduce memory churn, we initialize a single worst-case-size buffer in advance,
    ' then reuse it for all compression test runs.
    m_cmpTestBufferSize = Compression.GetWorstCaseSize(srcPDImage.Width * srcPDImage.Height * 4, cf_Lz4)
    ReDim m_cmpTestBuffer(0 To m_cmpTestBufferSize - 1) As Byte
    
    'We also want to know if the source image is non-paletted (e.g. "full color").
    ' If it isn't (meaning if it's already <= 256 colors per frame), the source pixel data probably
    ' came from an existing animated GIF file, and we'll want to optimize the data differently.
    '
    'Also, if auto-dithering is enabled, we dither frames *only* when the source data is full-color.
    Dim sourceIsFullColor As Boolean
    Set tmpLayer = New pdLayer
    tmpLayer.CopyExistingLayer srcPDImage.GetLayerByIndex(0)
    tmpLayer.ConvertToNullPaddedLayer srcPDImage.Width, srcPDImage.Height, True
    
    sourceIsFullColor = (Palettes.GetDIBColorCount_FastAbort(tmpLayer.GetLayerDIB, imgPalette) > 256)
    If autoDither Then useDithering = sourceIsFullColor
    
    'If we detect two identical back-to-back frames (surprisingly common in GIFs "in the wild"),
    ' we will simply merge their frame times into a single value and remove the duplicate frame.
    ' This reduces file size "for free", but it requires more complicated tracking as the number
    ' of frames may decrease as optimization proceeds.
    Dim numGoodFrames As Long, lastGoodFrame As Long
    numGoodFrames = 0
    lastGoodFrame = 0
    
    'We are now going to iterate through all layers in the image TWICE.
    
    'On this first pass, we will analyze each layer, produce optimized global and
    ' local palettes, extract frame times from layer names, and determine regions
    ' of interest in each frame.  Then we will palettize each layer and cache the
    ' palettized pixels in a simple 1D array.
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        
        'Optimizing frames can take some time.  Keep the user apprised of our progress.
        ProgressBars.SetProgBarVal i
        Message "Optimizing animation frame %1 of %2...", i + 1, srcPDImage.GetNumOfLayers + 1, "DONOTLOG"
        
        'Initialize default frame settings
        SaveGIF_ToStream_Animated = AnimatedGIF_InitFrame(i, useFixedFrameDelay, frameDelayDefault, srcPDImage)
        
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
        curFrameBackup.CreateFromExistingDIB tmpLayer.GetLayerDIB
        Set m_allFrames(i).frameDIB = New pdDIB
        
        'Force alpha to 0 or 255 only (this is a GIF requirement). It simplifies
        ' subsequent steps to handle this up-front.
        DIBs.ApplyAlphaCutoff_Gif tmpLayer.GetLayerDIB, gifAlphaCutoff, gifMatteColor
        
        'The first frame in the file should always be full-size, which limits what we can do
        ' to optimize it.  (Technically, you *could* write GIFs whose first frame is smaller
        ' than the image as a whole.  PD does not do this because it's terrible practice,
        ' and the file size savings are not meaningfully better.)
        If (i = 0) Then
        
            'Cache the temporary layer DIB as-is; it serves as both the first frame in
            ' the animation, and the fallback "static" image for decoders that don't
            ' understand animated GIFs.  (Also, after creating it, we immediately suspend
            ' the DIB to a compressed array to reduce memory constraints.)
            m_allFrames(i).frameDIB.CreateFromExistingDIB tmpLayer.GetLayerDIB
            m_allFrames(i).frameDIB.SuspendDIB cf_Lz4, False
            
            'Initialize the frame buffer to be the same as the first frame...
            bufferFrame.CreateFromExistingDIB tmpLayer.GetLayerDIB
            
            '...and initialize the "previous" frame buffer to pure transparency (this is complicated
            ' for GIFs because the recommendation is to avoid this disposal method entirely - due to
            ' the complications it requires in the decoder - but they also suggest that the decoder
            ' can ignore these instructions and use the background color of the GIF "if they have to".
            ' In PNGs, the spec says to use transparent black but GIFs may not support transparency
            ' at all... so I'm not sure what to do.  For now, I'm using the same strategy as PNGs
            ' and will revisit if problems arise.
            prevFrame.CreateBlank tmpLayer.GetLayerDIB.GetDIBWidth, tmpLayer.GetLayerDIB.GetDIBHeight, 32, 0, 0
            
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
            If DIBs.GetRectOfInterest_Overlay(tmpLayer.GetLayerDIB, bufferFrame, dupArea) Then
                
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
                If DIBs.GetRectOfInterest_Overlay(tmpLayer.GetLayerDIB, prevFrame, prevFrameArea) Then
                
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
                    '
                    '(This is handled by a separate function; all frame flags will be set accordingly
                    ' if the function succeeds.)
                    AnimatedGIF_CheckDisposalBkgd tmpLayer, i, lastGoodFrame
                    
                    'Because the current frame came from a premultiplied source, we can safely
                    ' mark it as premultiplied as well.
                    If (Not m_allFrames(i).frameDIB Is Nothing) Then m_allFrames(i).frameDIB.SetInitialAlphaPremultiplicationState True
                    
                    'If the previous frame is not being blanked, we have additional optimization
                    ' strategies to attempt.  (If, however, the previous frame *is* being blanked,
                    ' we are done with preprocessing because we have no "previous" data to work with.)
                    If (m_allFrames(lastGoodFrame).frameDisposal <> gd_Background) Then
                        
                        'Pixel-blanking is handled by a dedicated function
                        AnimatedGIF_CheckPixelBlanking tmpLayer, refFrame, srcPDImage, i, lastGoodFrame
                        
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
                        If (UBound(m_globalPalette) > 255) Then ReDim Preserve m_globalPalette(0 To 255) As RGBQuad
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
                    AnimatedGIF_CheckLocalPaletteCompressibility i, imgPalette
                
                'Pixel blanking wasn't available for this frame (likely due to this frame containing transparency
                ' where the prior frame did not).  Palettize it as-is.
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
    
    'Note: at this point, frames that rely on the global palette have *not* been optimized yet.
    ' This is because they may be able to use transparency, which wasn't guaranteed present
    ' at the time of their original processing (because the global palette hadn't filled up yet.)
    
    'So our next job is to get the global palette in order.
    
    'The GIF spec requires all palette color counts to be a power of 2.  (It does this because
    ' palette color count is stored in 3-bits, ugh.)  Any unused entries are ignored, but by
    ' convention are usually left as black; we do the same here.
    m_numColorsInGP = 2 ^ Pow2FromColorCount(m_numColorsInGP)
    If (UBound(m_globalPalette) <> m_numColorsInGP - 1) Then ReDim Preserve m_globalPalette(0 To m_numColorsInGP - 1) As RGBQuad
    
    'Move the transparent index, if any, to the front of the global palette
    AnimatedGIF_SortGlobalPaletteAlpha
    
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
                AnimatedGIF_CheckGlobalPaletteCompressibility i
                
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
    
    'Compression test buffer is no longer required; free it to regain some memory
    Erase m_cmpTestBuffer
    
    'With all frames generated, we could now do an interesting legacy optimization: sorting the
    ' palette by "importance", which would allow us to check the "Sort flag" in the GIF header.
    ' Because this is fairly intensive (we have to analyze all frames that use the global palette),
    ' I have disabled this feature in production code, but you can reactivate it if curious.
    If SORT_PALETTE_BY_IMPORTANCE Then AnimatedGIF_CheckGlobalPaletteSort srcPDImage
    
    Message "Finalizing image..."
    ProgressBars.SetProgBarVal ProgressBars.GetProgBarMax()
    
    'We've successfully optimized all GIF frames.  Now it's time to involve the encoder.
    '
    'In 2021, PD switched from the 3rd-party FreeImage library to a homebrew solution for GIF encoding.
    ' Our pure-VB6 encoder uses less memory and is slightly faster, while also constructing smaller GIFs
    ' (FreeImage always writes 8-bit color tables even if the palette could use fewer bits).  If you
    ' encounter any issues with the new encoder, please file an issue on GitHub so I can investigate!
    SaveGIF_ToStream_Animated = WriteOptimizedAGIFToFile_Internal(srcPDImage, dstStream, formatParams, metadataParams)
    
    'Manually erase any module-level containers (don't need them to live on after this!)
    Erase m_globalPalette
    Erase m_allFrames
    
    Exit Function
    
ExportGIFError:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ", " & Err.Description
    SaveGIF_ToStream_Animated = False
    
End Function

'From the comments in the main animated GIF export function:
' "We have one more potential crop operation we can do, and it involves the third
' disposal op (DISPOSAL_BACKGROUND).  This disposal op asks the previous frame
' to erase itself completely after rendering.  For animations with large
' transparent borders, it may actually be best to crop the current frame
' according to its transparent borders, then use the "erase" disposal op before
' displaying it, thus forgoing any connection whatsoever to preceding frames."
Private Function AnimatedGIF_CheckDisposalBkgd(ByRef tmpLayer As pdLayer, ByVal frameIndex As Long, ByVal idxLastGoodFrame As Long) As Boolean
    
    Dim trnsRect As RectF
    If DIBs.GetRectOfInterest(tmpLayer.GetLayerDIB, trnsRect) Then
    
        'If this frame is smaller than the previous "winner", switch to using this op instead.
        If (trnsRect.Width * trnsRect.Height) < (m_allFrames(frameIndex).rectOfInterest.Width * m_allFrames(frameIndex).rectOfInterest.Height) Then
            m_allFrames(idxLastGoodFrame).frameDisposal = gd_Background
            m_allFrames(frameIndex).rectOfInterest = trnsRect
        End If
        
        'Crop the "winning" region into a separate DIB, and store it as the formal pixel buffer
        ' for this frame.
        With m_allFrames(frameIndex).rectOfInterest
            m_allFrames(frameIndex).frameDIB.CreateBlank Int(.Width), Int(.Height), 32, 0, 0
            GDI.BitBltWrapper m_allFrames(frameIndex).frameDIB.GetDIBDC, 0, 0, Int(.Width), Int(.Height), tmpLayer.GetLayerDIB.GetDIBDC, Int(.Left), Int(.Top), vbSrcCopy
        End With
    
    'This weird (but valid) branch means that the current frame is 100% transparent.  For this
    ' special case, request that the previous frame erase the running buffer, then store a 1 px
    ' transparent pixel.
    Else
        
        m_allFrames(frameIndex).frameDIB.CreateBlank 1, 1, 32, 0, 0
        With m_allFrames(frameIndex).rectOfInterest
            .Left = 0
            .Top = 0
            .Width = 1
            .Height = 1
        End With
        m_allFrames(frameIndex).frameNeedsTransparency = True
        m_allFrames(idxLastGoodFrame).frameDisposal = gd_Leave
        
    End If
    
    AnimatedGIF_CheckDisposalBkgd = True
    
End Function

'When a frame uses the global palette, we have additional optimzation options (vs local palettes).
' We can freely mix-and-match valid duplicate-pixel-blanked and original pixels in whatever combination
' produces the most compressible output.  PD attempts this in fixed "runs" of pixels, in an attempt to
' produce the smallest possible output.
Private Function AnimatedGIF_CheckGlobalPaletteCompressibility(ByVal frameIndex As Long) As Boolean
    
    'Palettize both frames (without dithering; the extra noise it introduces is
    ' not helpful) into temporary arrays.
    Dim tmpBlanked() As Byte, tmpOriginal() As Byte
    DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(frameIndex).frameDIB, m_globalPalette, tmpBlanked
    DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(frameIndex).backupFrameDIB, m_globalPalette, tmpOriginal
    
    'From these, build a new palettized image that uses the most compressible
    ' scanlines from each.
    Dim tmpMerged() As Byte
    DIBs.MakeMinimalEntropyScanlines tmpOriginal, tmpBlanked, m_allFrames(frameIndex).frameDIB.GetDIBWidth, m_allFrames(frameIndex).frameDIB.GetDIBHeight, tmpMerged
    
    'We now have three choices of pixel streams:
    ' 1) the original (untouched) frame
    ' 2) a frame with duplicate pixels between (1) and the previous frame blanked out
    ' 3) a frame that attempts to combine the "best" scanlines from (1) and (2)
    '
    'These frames all have the same width and height, but they differ in internal contents.
    Dim netPixels As Long
    netPixels = m_allFrames(frameIndex).frameDIB.GetDIBWidth * m_allFrames(frameIndex).frameDIB.GetDIBHeight
    ReDim m_allFrames(frameIndex).pixelData(0 To m_allFrames(frameIndex).frameDIB.GetDIBWidth - 1, 0 To m_allFrames(frameIndex).frameDIB.GetDIBHeight - 1) As Byte
    
    'We're now going to do a quick compression check of each stream and take whichever one
    ' compresses the best.  This is the closest thing we have to a "foolproof" strategy.
    Dim initCmpSize As Long, initSize As Long
    initCmpSize = m_cmpTestBufferSize
    initSize = netPixels
    Compression.CompressPtrToPtr VarPtr(m_cmpTestBuffer(0)), initCmpSize, VarPtr(tmpOriginal(0, 0)), initSize, cf_Lz4
    
    Dim testSize As Long
    testSize = m_cmpTestBufferSize
    Compression.CompressPtrToPtr VarPtr(m_cmpTestBuffer(0)), testSize, VarPtr(tmpBlanked(0, 0)), initSize, cf_Lz4
    
    'Compare the smaller of the previous test to our minimal-entropy attempt
    If (initCmpSize < testSize) Then
        
        testSize = m_cmpTestBufferSize
        Compression.CompressPtrToPtr VarPtr(m_cmpTestBuffer(0)), testSize, VarPtr(tmpMerged(0, 0)), initSize, cf_Lz4
        
        'Copy the result into the frame collection, then free all DIBs
        If (initCmpSize < testSize) Then
            CopyMemoryStrict VarPtr(m_allFrames(frameIndex).pixelData(0, 0)), VarPtr(tmpOriginal(0, 0)), netPixels
        Else
            CopyMemoryStrict VarPtr(m_allFrames(frameIndex).pixelData(0, 0)), VarPtr(tmpMerged(0, 0)), netPixels
        End If
        
    'Same thing, but comparing the other array
    Else
        
        initCmpSize = m_cmpTestBufferSize
        Compression.CompressPtrToPtr VarPtr(m_cmpTestBuffer(0)), initCmpSize, VarPtr(tmpMerged(0, 0)), initSize, cf_Lz4
        
        If (initCmpSize < testSize) Then
            CopyMemoryStrict VarPtr(m_allFrames(frameIndex).pixelData(0, 0)), VarPtr(tmpMerged(0, 0)), netPixels
        Else
            CopyMemoryStrict VarPtr(m_allFrames(frameIndex).pixelData(0, 0)), VarPtr(tmpBlanked(0, 0)), netPixels
        End If
    
    End If
    
    'Free all DIB copies for this frame; they are no longer required
    Set m_allFrames(frameIndex).frameDIB = Nothing
    Set m_allFrames(frameIndex).backupFrameDIB = Nothing
    
    AnimatedGIF_CheckGlobalPaletteCompressibility = True
    
End Function

'The GIF spec has this interesting (historically speaking) flag, called the "Sort Flag".
' From the spec:
' v) Sort Flag - Indicates whether the Global Color Table is sorted.
'            If the flag is set, the Global Color Table is sorted, in order of
'            decreasing importance. Typically, the order would be decreasing
'            frequency, with most frequent color first. This assists a decoder,
'            with fewer available colors, in choosing the best subset of colors;
'            the decoder may use an initial segment of the table to render the
'            graphic.
'
'            Values :    0 -   Not ordered.
'                        1 -   Ordered by decreasing importance, most important color first.
'
'This sort value won't matter with modern decoders, but I was curious to see how I'd implement
' such a feature.  Truth be told, I actually went a slightly different direction, by picking the
' most popular color first and *then* sorting subsequent values by how many instances follow
' the previous color in the palette.  This (hypothetically) produces the most perf-optimized
' palette for decompression, but benchmarking has shown the final result to be inconclusively
' useful, at least on the few decoders I tried.
'
'Anyway, this function only works on the global palette (which is the only palette with the "sort flag")
' and I'm leaving it here because it's interesting!
Private Sub AnimatedGIF_CheckGlobalPaletteSort(ByRef srcPDImage As pdImage)
    
    'Start by seeing how many frames in this image use the global palette.  If more than 10% do, let's sort
    ' the global palette accordingly.  (If fewer frames use it, this image is clearly local-palette-dominated
    ' and the global palette doesn't matter.)
    Dim candidateForGlobalPaletteSort As Boolean
    Dim numFramesGlobalPalette As Long, totalCount As Long
    
    Dim i As Long
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        If (Not m_allFrames(i).frameIsDuplicateOrEmpty) Then
            totalCount = totalCount + 1
            If m_allFrames(i).usesGlobalPalette Then numFramesGlobalPalette = numFramesGlobalPalette + 1
        End If
    Next i
    
    If (totalCount > 0) Then
        candidateForGlobalPaletteSort = ((numFramesGlobalPalette / totalCount) >= 0.1)
    Else
        candidateForGlobalPaletteSort = False
    End If
    
    If (Not candidateForGlobalPaletteSort) Then Exit Sub
    
    'If we're still here, it's time to sort the global palette.
    
    'Unfortunately, sorting the global palette requires us to analyze every frame that uses the global palette...
    ' so this can be kind of an intensive task!
    
    'Start by building a "next-palette-index" lookup table.  We want to know not just palette index popularity,
    ' but which palette entries are most likely to follow each other.  To do this, we'll keep a simple running
    ' histogram that tracks how many times each palette entry follows each other palette entry.
    Dim palPopularity() As Long
    ReDim palPopularity(0 To m_numColorsInGP - 1) As Long
    Dim palFollows() As Long
    ReDim palFollows(0 To m_numColorsInGP - 1, 0 To m_numColorsInGP - 1) As Long
    
    Dim x As Long, curIndex As Long, nextIndex As Long
    Dim px1D() As Byte, pxSA As SafeArray1D
    
    'Scan all frames that use the global palette and populate both histograms accordingly.
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        If (Not m_allFrames(i).frameIsDuplicateOrEmpty) And m_allFrames(i).usesGlobalPalette Then
            
            With m_allFrames(i)
                
                totalCount = .rectOfInterest.Width * .rectOfInterest.Height - 1
                VBHacks.WrapArrayAroundPtr_Byte px1D, pxSA, VarPtr(.pixelData(0, 0)), totalCount + 1
                
                'Deliberately skip the last pixel in the image (we need to access "next pixels" and the
                ' last pixel in the image doesn't have one)
                For x = 0 To totalCount - 1
                    curIndex = px1D(x)
                    nextIndex = px1D(x + 1)
                    palPopularity(curIndex) = palPopularity(curIndex) + 1
                    palFollows(curIndex, nextIndex) = palFollows(curIndex, nextIndex) + 1
                Next x
                
                VBHacks.UnwrapArrayFromPtr_Byte px1D
                
            End With
            
        End If
    Next i
    
    'Histograms are now built.  We now want to sort the global palette by popularity,
    ' but we're going to use kind of a cool criteria:
    ' 1) Transparent index, if any, still gets primo spot 0 (this improves compatibility with old GIF decoders)
    ' 2) The *most popular* color gets index 1
    ' 3) All subsequent colors are placed according to how frequently they follow the previous color in the palette
    '
    'Obviously, we also need to build a remap table that notes where each palette index was, and where it is
    ' in the new palette order.
    Dim remapTable() As Byte
    ReDim remapTable(0 To m_numColorsInGP - 1) As Byte
    
    'Start with the transparent index, if any
    Dim idxRemap As Long
    If (m_GlobalTrnsIndex >= 0) Then
        remapTable(0) = 0
        curIndex = 1
        idxRemap = 1
        
        'Erase popularity data for position 0
        palPopularity(0) = -1
        For i = 0 To m_numColorsInGP - 1
            palFollows(i, 0) = -1
        Next i
        
    Else
        curIndex = 0
        idxRemap = 0
    End If
    
    'Next, find the most popular color in the palette
    Dim idxCurMax As Long, maxPopularity As Long, prevIndex As Long
    idxCurMax = curIndex
    maxPopularity = palPopularity(idxCurMax)
    curIndex = curIndex + 1
    
    Do While (curIndex < m_numColorsInGP)
        If (palPopularity(curIndex) > maxPopularity) Then
            maxPopularity = palPopularity(curIndex)
            idxCurMax = curIndex
        End If
        curIndex = curIndex + 1
    Loop
    
    'idxCurMax now points to the "most popular" color across all frames that use the global palette.
    ' Assign its place in the remap table, then erase its popularity data.
    remapTable(idxCurMax) = idxRemap
    idxRemap = idxRemap + 1
    
    palPopularity(idxCurMax) = -1
    For i = 0 To m_numColorsInGP - 1
        palFollows(i, idxCurMax) = -1
    Next i
    
    'For all remaining colors, we now want to pick the "most popular color" that follows the previous color
    ' in the table.  It is possible that *no* colors will follow the current color (just due to the way the
    ' search algorithm works), and when that happens, we'll revert back to using the most-popular, currently
    ' un-remapped color.
    Do While (idxRemap < m_numColorsInGP)
        
        maxPopularity = -1
        prevIndex = idxCurMax
        
        'Find the most popular color that follows the *previous* color in the table
        For i = 0 To m_numColorsInGP - 1
            If (palFollows(prevIndex, i) >= 0) Then
                If (palFollows(prevIndex, i) > maxPopularity) Then
                    maxPopularity = palFollows(prevIndex, i)
                    idxCurMax = i
                End If
            End If
        Next i
        
        'If the maximum popularity is 0, revert back to "most popular color"
        If (maxPopularity < 0) Then
            For i = 0 To m_numColorsInGP - 1
                If (palPopularity(i) >= maxPopularity) Then
                    maxPopularity = palPopularity(i)
                    idxCurMax = i
                End If
            Next i
        End If
        
        'We have our next remap candidate.  Place it in the remap table, then erase all popularity data
        ' for this pixel (so we don't select it again).
        remapTable(idxCurMax) = idxRemap
        palPopularity(idxCurMax) = -1
        For i = 0 To m_numColorsInGP - 1
            palFollows(i, idxCurMax) = -1
        Next i
        
        idxRemap = idxRemap + 1
        
    Loop
    
    'We now have our final remap table!  All that's left to do is...
    ' 1) Construct a new, remapped global palette
    ' 2) Remap all pixel streams accordingly
    
    'Start by remapping the global palette
    Dim newPalette() As RGBQuad
    ReDim newPalette(0 To m_numColorsInGP - 1) As RGBQuad
    For i = 0 To m_numColorsInGP - 1
        newPalette(remapTable(i)) = m_globalPalette(i)
    Next i
    
    VBHacks.CopyMemoryStrict VarPtr(m_globalPalette(0)), VarPtr(newPalette(0)), m_numColorsInGP * 4
    
    'Next, remap all pixel streams that rely on the global palette
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        If (Not m_allFrames(i).frameIsDuplicateOrEmpty) And m_allFrames(i).usesGlobalPalette Then
            
            With m_allFrames(i)
                
                totalCount = .rectOfInterest.Width * .rectOfInterest.Height - 1
                VBHacks.WrapArrayAroundPtr_Byte px1D, pxSA, VarPtr(.pixelData(0, 0)), totalCount + 1
                
                'Deliberately skip the last pixel in the image (we need to access "next pixels" and the
                ' last pixel in the image doesn't have one)
                For x = 0 To totalCount
                    px1D(x) = remapTable(px1D(x))
                Next x
                
                VBHacks.UnwrapArrayFromPtr_Byte px1D
                
            End With
            
        End If
    Next i
    
End Sub

'Different options are available for frames that use local palettes instead of global ones.
' Global palette optimization is actually easier because we can blank out pixels freely,
' knowing that back-to-back frames share identical colors.  (So if we intermix results
' from both, we can produce the most compressible scanlines from either.)  But if the
' current frame uses a local palette, we *have* to either palettize the whole thing as-is,
' or we need to blank-out all duplicate pixels as-is.  Mixing and matching is not doable
' because the previous frame may not share colors identically - so we either need to suck
' it up and treat this as an independent frame, or we need to go all-in on reusing data
' from the previous frame.
'
'Anyway, compare this function to the similar one for global palette optimization if
' you're curious about how the two approaches differ.  (Note also that originally they
' shared all code, but using those optimizations on local palette frames was breaking
' some esoteric GIFs that use multiple animation frames to produce a final result with
' more than 256 colors.)
Private Function AnimatedGIF_CheckLocalPaletteCompressibility(ByVal frameIndex As Long, ByRef imgPalette() As RGBQuad) As Boolean
    
    'Palettize the blanked frame into a temporary array (without dithering; the extra noise
    ' it introduces is not helpful), then generate a temporary palette and palettize the
    ' backup frame too.
    Dim tmpBlanked() As Byte
    DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(frameIndex).frameDIB, imgPalette, tmpBlanked
    
    'Backup frame.  Note that the palette produced here is important - we may keep it as the frame's
    ' final palette *if* the backup frame compresses better than the pixel-blanked copy.
    Dim tmpFramePalette() As RGBQuad, tmpFramePaletteCount As Long
    Palettes.GetOptimizedPaletteIncAlpha m_allFrames(frameIndex).backupFrameDIB, tmpFramePalette, 256, pdqs_Variance, True
    tmpFramePaletteCount = UBound(tmpFramePalette) + 1
    Palettes.EnsureBinaryAlphaPalette tmpFramePalette
    Palettes.SortPaletteForCompression_IncAlpha m_allFrames(frameIndex).backupFrameDIB, tmpFramePalette, True, True
    
    Dim tmpOriginal() As Byte
    DIBs.GetDIBAs8bpp_RGBA_SrcPalette m_allFrames(frameIndex).backupFrameDIB, tmpFramePalette, tmpOriginal
    
    'We now have two choices of pixel streams:
    ' 1) the original (untouched) frame, in tmpOriginal()
    ' 2) a frame with duplicate pixels between (1) and the previous frame blanked out, in tmpBlanked()
    '
    'These two frames have the same dimensions (width x height), but they may compress differently
    ' due to their pixel contents.
    Dim netPixels As Long
    netPixels = m_allFrames(frameIndex).frameDIB.GetDIBWidth * m_allFrames(frameIndex).frameDIB.GetDIBHeight
    
    'We're now going to do a quick lz4 check of each stream and keep whichever one compresses best.
    ' This provides a very fast, reasonably good estimate of which frame will compress the best
    ' under GIF's far more primitive LZW strategy.
    Dim initCmpSize As Long, initSize As Long
    initCmpSize = m_cmpTestBufferSize
    initSize = netPixels
    Compression.CompressPtrToPtr VarPtr(m_cmpTestBuffer(0)), initCmpSize, VarPtr(tmpOriginal(0, 0)), initSize, cf_Lz4
    
    Dim testSize As Long
    testSize = m_cmpTestBufferSize
    Compression.CompressPtrToPtr VarPtr(m_cmpTestBuffer(0)), testSize, VarPtr(tmpBlanked(0, 0)), initSize, cf_Lz4
    
    'Prep a destination array for the pixel data, inside the current frame object.  (This will be
    ' the final palettized copy of the image that gets embedded in the GIF file.)
    ReDim m_allFrames(frameIndex).pixelData(0 To Int(m_allFrames(frameIndex).rectOfInterest.Width) - 1, 0 To Int(m_allFrames(frameIndex).rectOfInterest.Height) - 1) As Byte
    
    'If the pixel-blanked frame compressed better...
    If (testSize < initCmpSize) Then
        
        'Copy its pixel stream into the frame collection
        CopyMemoryStrict VarPtr(m_allFrames(frameIndex).pixelData(0, 0)), VarPtr(tmpBlanked(0, 0)), netPixels
        
    'The original frame compressed better...
    Else
        
        'Copy pixel *and* palette data into the frame collection, overwriting the pixel-blanked
        ' frame we generated (and its associated palette)
        CopyMemoryStrict VarPtr(m_allFrames(frameIndex).pixelData(0, 0)), VarPtr(tmpOriginal(0, 0)), netPixels
        m_allFrames(frameIndex).palNumColors = tmpFramePaletteCount
        ReDim m_allFrames(frameIndex).framePalette(0 To tmpFramePaletteCount - 1)
        CopyMemoryStrict VarPtr(m_allFrames(frameIndex).framePalette(0)), VarPtr(tmpFramePalette(0)), tmpFramePaletteCount * 4
        
    End If
    
    'Free the backup DIB while we're here - it's no longer required
    Set m_allFrames(frameIndex).backupFrameDIB = Nothing
    
    AnimatedGIF_CheckLocalPaletteCompressibility = True
    
End Function

'If the previous frame in an animated GIF is not being blanked (e.g. not gd_Background),
' we can attempt to "blank" pixels in the current frame whose color hasn't changed.
' On some GIF frames, this can produce large compression improvements - but note that there
' are many edge-cases that need to be resolved, including states like "a pixel in the current
' frame is transparent, but not in the previous frame", which negate this strategy.
' Blanked frames may also become extremely noisy if the original frames are dithered or
' captured from film/photo sources, which also hurts compression.  Subsequent functions in
' PD's animated GIF encoder will test these various states to ensure compression is optimal.
' This function just produces the blanked frame copy (if disposal modes allow).
Private Function AnimatedGIF_CheckPixelBlanking(ByRef tmpLayer As pdLayer, ByRef refFrame As pdDIB, ByRef srcImage As pdImage, ByVal frameIndex As Long, ByVal idxLastGoodFrame As Long) As Boolean

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
    Dim trnsTable() As Byte
    If DIBs.RetrieveTransparencyTable(m_allFrames(frameIndex).frameDIB, trnsTable) Then
    If DIBs.CheckAlpha_DuplicatePixels(refFrame, m_allFrames(frameIndex).frameDIB, trnsTable, Int(m_allFrames(frameIndex).rectOfInterest.Left), Int(m_allFrames(frameIndex).rectOfInterest.Top)) Then
        
        'This frame contains transparency where the previous frame does not.
        ' This means the previous frame *must* be blanked.
        ' Skip any remaining frame differential optimizations entirely.
        m_allFrames(idxLastGoodFrame).frameDisposal = gd_Background
        
        'As a consequence of "blanking" the previous frame, we need to render the current
        ' frame in its entirety (except for transparent borders, if they exist).
        With m_allFrames(frameIndex).rectOfInterest
            
            Dim trnsRect As RectF
            If DIBs.GetRectOfInterest(tmpLayer.GetLayerDIB, trnsRect) Then
                .Left = trnsRect.Left
                .Top = trnsRect.Top
                .Width = trnsRect.Width
                .Height = trnsRect.Height
            
            'Failsafe only; this state was dealt with in a previous step.
            Else
                .Left = 0
                .Top = 0
                .Width = srcImage.Width
                .Height = srcImage.Height
            End If
            
            'Crop the "winning" region into a separate DIB, and store it as the formal
            ' pixel buffer for this frame.
            m_allFrames(frameIndex).frameDIB.CreateBlank Int(.Width), Int(.Height), 32, 0, 0
            GDI.BitBltWrapper m_allFrames(frameIndex).frameDIB.GetDIBDC, 0, 0, Int(.Width), Int(.Height), tmpLayer.GetLayerDIB.GetDIBDC, Int(.Left), Int(.Top), vbSrcCopy
        
        End With
        
    'The current frame can be safely alpha-blended "over the top" of the previous one
    Else
        
        'This frame is a candidate for frame differentials!
        
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
        Dim testDIB As pdDIB
        Set testDIB = New pdDIB
        testDIB.CreateFromExistingDIB m_allFrames(frameIndex).frameDIB
        If DIBs.RetrieveTransparencyTable(testDIB, trnsTable) Then
        If DIBs.ApplyAlpha_DuplicatePixels(testDIB, refFrame, trnsTable, m_allFrames(frameIndex).rectOfInterest.Left, m_allFrames(frameIndex).rectOfInterest.Top, True) Then
        If DIBs.ApplyTransparencyTable(testDIB, trnsTable) Then
        
            'The frame differential was produced successfully.  We won't actually use it here;
            ' instead, we'll use it later, after palettization (so we can test its entropy
            ' against the original, untouched version).
            
            'Back up the existing frame copy (in case we decide to use the global palette
            ' and it lacks transparency)
            Set m_allFrames(frameIndex).backupFrameDIB = m_allFrames(frameIndex).frameDIB
            
            'Copy the current DIB, and set all corresponding flags
            Set m_allFrames(frameIndex).frameDIB = testDIB
            m_allFrames(frameIndex).frameWasBlanked = True
            m_allFrames(frameIndex).frameNeedsTransparency = True
            
        '/end failsafe "retrieve and apply duplicate pixel test" checks
        End If
        End If
        End If
        
    '/end "previous frame is opaque where this frame is transparent"
    End If
    End If
    
    AnimatedGIF_CheckPixelBlanking = True

End Function

'PD places the transparent color index - if any - in the first position in the global palette.
' (This is more of a convention than a requirement by the GIF spec, but it has potential knock-on
' effects for compressibility, so there is potential gain and zero harm in ordering the palette
' this way.)
Private Function AnimatedGIF_SortGlobalPaletteAlpha() As Boolean

    'If the global palette has a transparent index, locate it and ensure it is in position 0.
    ' (While not required by the spec, this *is* required by the PNG spec, and it generally
    ' improves compression to set it early in the GIF color table too, given where it's likely
    ' to appear in real-world images.)
    m_GlobalTrnsIndex = -1
    
    Dim i As Long
    
    'Next, count how many non-transparent colors are in this palette.  If it is an even
    ' power-of-two, and the first frame does *not* contain transparency, we are going
    ' to cancel transparency for this GIF.
    Dim numNonTransparentColors As Long
    For i = 0 To m_numColorsInGP - 1
        If (m_globalPalette(i).Alpha <> 0) Then numNonTransparentColors = numNonTransparentColors + 1
    Next i
    
    If (numNonTransparentColors = 2 ^ (Pow2FromColorCount(numNonTransparentColors))) Then
        
        'If we're already at 256 colors, don't worry about this step.
        If (numNonTransparentColors < 256) Then
            
            'The number of non-transparent colors is an exact power of two.  Some GIF optimizers
            ' work by cutting the image palette down to 128 or 64 colors, which is fine - but we
            ' need to check the first frame of the image to see if it uses transparency.  If it
            ' does, we have no choice but to reserve transparency in the final palette, but if it
            ' *doesn't* use transparency, that means PD's palette optimizer added transparency
            ' as part of frame optimizations.  We now need to roll-back that decision to avoid
            ' kicking the LZW compressor up to the next bit-count (which will disproportionately
            ' increase final file size).
            If (Not DIBs.IsDIBTransparent(m_allFrames(0).frameDIB)) Then
                
                'The first frame does not use transparency.
                
                'Condense all non-transparent pixels into the base [n] colors of the palette,
                ' then mark this image as non-transparent and exit.  Note that this doesn't
                ' create any problems further down the line - PD knows to validate this case,
                ' and roll-back to a non-pixel-blanked version of each affected frame.
                m_numColorsInGP = Palettes.GetPalette_OpaqueColorsOnly(m_globalPalette)
                m_GlobalTrnsIndex = -1
                
                AnimatedGIF_SortGlobalPaletteAlpha = True
                Exit Function
                
            End If
            
        End If
        
    End If
    
    'If we're still here, the palette is OK as-is.  Proceed with shifting the transparent index
    ' to the front of the palette.
    For i = 0 To m_numColorsInGP - 1
        If (m_globalPalette(i).Alpha = 0) Then
            
            If (i > 0) Then
                
                'Shift all previous colors backward, then plug this transparent pixel into
                ' the *first* palette position .
                Dim tmpColor As RGBQuad
                tmpColor = m_globalPalette(i)
                Dim j As Long
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
    
    AnimatedGIF_SortGlobalPaletteAlpha = True
    
End Function

Private Function AnimatedGIF_InitFrame(ByVal frameIndex As Long, ByVal useFixedFrameDelay As Boolean, ByVal frameDelayDefault As Long, ByRef srcImage As pdImage) As Boolean

    With m_allFrames(frameIndex)
    
        'Before dealing with pixel data, attempt to retrieve a frame time from the
        ' source layer's name. (If the layer name does not provide a frame time,
        ' or if the user has specified a fixed frame time, this value will be
        ' overwritten with their requsted value.)
        Dim finalFrameTime As Long
        finalFrameTime = GetFrameTimeFromLayerName(srcImage.GetLayerByIndex(frameIndex).GetLayerName, 0)
        If (useFixedFrameDelay Or (finalFrameTime = 0)) Then finalFrameTime = frameDelayDefault
        .frameTime = finalFrameTime
        
        'Remaining parameters are contingent on optimization passes; for now, populate with
        ' safe default parameters (e.g parameters that produce a functional GIF even if
        ' optimization fails). Subsequent optimization rounds will modify these settings if
        ' it produces a smaller file.
        .frameDisposal = gd_Leave
        .frameIsDuplicateOrEmpty = False
        .rectOfInterest.Left = 0
        .rectOfInterest.Top = 0
        .rectOfInterest.Width = srcImage.Width
        .rectOfInterest.Height = srcImage.Height
        
    End With

    AnimatedGIF_InitFrame = True
    
End Function

'After optimizing GIF frames (which you must do to generate the structures used by this function),
' call this function to actually write GIF data out to file.  Native VB6 functions will be used.
Private Function WriteOptimizedAGIFToFile_Internal(ByRef srcPDImage As pdImage, ByRef dstStream As pdStream, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    Const FUNC_NAME As String = "WriteOptimizedAGIFToFile_Internal"
    
    On Error GoTo ExportGIFError
    
    WriteOptimizedAGIFToFile_Internal = False
    
    'Parse all relevant GIF parameters.  (See the GIF export dialog for details on how these are generated.)
    ' Note that in this function, we only need the parameter for animation loop count - all other settings
    ' were dealt with in previous steps.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    'Cache the nearest power-of-two for the global palette; it will be used for LZW encoding
    Dim pow2forGP As Long
    pow2forGP = Pow2FromColorCount(m_numColorsInGP) - 1
    
    'For detailed GIF format info, see http://giflib.sourceforge.net/whatsinagif/bits_and_bytes.html
    '
    'This class provides helper functions (shared between static and animated GIFs) that write various
    ' chunks of the GIF file for us.
    WriteOptimizedAGIFToFile_Internal = WriteGIF_FileHeader(dstStream, srcPDImage.Width, srcPDImage.Height, m_numColorsInGP, m_globalPalette, 0)
    
    'The image header is now complete.
    
    'For animated images, we now need to place a custom application extension (first used by Netscape)
    ' to declare the loop behavior of the GIF. See https://en.wikipedia.org/wiki/GIF#Animated_GIF for details.
    ' (There will be various magic numbers used here; the wiki page shares more details on them.)
    ' Note that non-looping GIFs can skip this block entirely.
    Dim loopCount As Long
    loopCount = cParams.GetLong("animation-loop-count", 1)
    If (loopCount > 65536) Then loopCount = 65536
                    
    If (loopCount <> 1) Then
        
        'Application extension ID
        dstStream.WriteByte &H21
        dstStream.WriteByte &HFF
        
        'Size of block (always 11 for this block)
        dstStream.WriteByte 11
        
        'Application name + 3 verification bytes
        dstStream.WriteString_ASCII "NETSCAPE2.0"
        
        'Number of bytes in the following sub-block (always 3)
        dstStream.WriteByte 3
        
        'Sub-block index (always 1)
        dstStream.WriteByte 1
        
        'Number of repetitions - 1
        If (loopCount < 1) Then loopCount = 1
        loopCount = loopCount - 1
        dstStream.WriteIntU loopCount
        
        'End of the sub-block chain
        dstStream.WriteByte 0
        
    End If
    
    'Time to iterate and store frames.
    Dim i As Long
    For i = 0 To srcPDImage.GetNumOfLayers - 1
    
        'Skip duplicate or empty frames
        If (Not m_allFrames(i).frameIsDuplicateOrEmpty) Then
        
            'All frames are preceded by a "Graphics Control Extension".
            ' This is a fixed-size struct describing things like frame delay, transparency presence, etc.
            
            'Precalculate some values before passing them off to the dedicated GCE writer
            Dim frameUsesAlpha As Boolean, frameTrnsIndex As Long
            If m_allFrames(i).usesGlobalPalette Then
                frameUsesAlpha = (m_GlobalTrnsIndex >= 0)
                If frameUsesAlpha Then frameTrnsIndex = m_GlobalTrnsIndex
            Else
                frameUsesAlpha = (m_allFrames(i).framePalette(0).Alpha = 0)
                frameTrnsIndex = 0
            End If
            
            'giflib has an interesting convention where images without transparency always write the
            ' transparent index as 255.  The GIF spec says nothing about this, but sort of warns against
            ' it by stating "The index is present if and only if the Transparency Flag is set to 1."
            ' I assume "present" means "non-zero" which makes giflib's behavior even stranger,
            ' but regardless, PD just writes out a zero if the current frame doesn't use transparency.
            
            'PD stores frame times as ms.  Convert to centiseconds as required by the GIF spec.
            Dim finalFrameTime As Long
            finalFrameTime = Int((m_allFrames(i).frameTime + 5) \ 10)
            If (finalFrameTime > 65535) Then finalFrameTime = 65535
            
            'Pass final values to the writer
            WriteOptimizedAGIFToFile_Internal = WriteGIF_GCE(dstStream, frameUsesAlpha, m_allFrames(i).frameDisposal, frameTrnsIndex, finalFrameTime)
            
            'Graphics Control Extension is done
            
            'Next up is an "image descriptor", basically a frame header
            
            '1-byte image separator (always 2C)
            dstStream.WriteByte &H2C
            
            'Frame dimensions as unsigned shorts, in left/top/width/height order
            dstStream.WriteIntU Int(m_allFrames(i).rectOfInterest.Left)
            dstStream.WriteIntU Int(m_allFrames(i).rectOfInterest.Top)
            dstStream.WriteIntU Int(m_allFrames(i).rectOfInterest.Width)
            dstStream.WriteIntU Int(m_allFrames(i).rectOfInterest.Height)
            
            'And my favorite, another packed bit-field!  (uuuuugh)
            ' - 1 bit local palette used (varies by frame)
            ' - 1 bit interlaced (PD never interlaces frames)
            ' - 1 bit sort flag (same as global table, PD can - and may - do this, but always writes 0 per giflib convention)
            ' - 2 bits reserved
            ' - 3 bits size of local color table N (describing 2 ^ n-1 colors in the palette)
            Dim tmpBitField As Long
            tmpBitField = 0
            
            Dim pow2forLP As Long
            If (Not m_allFrames(i).usesGlobalPalette) Then
                tmpBitField = tmpBitField Or &H80
                pow2forLP = Pow2FromColorCount(m_allFrames(i).palNumColors) - 1
                tmpBitField = tmpBitField Or pow2forLP
            End If
            dstStream.WriteByte tmpBitField
            
            'There is no terminator here.  Instead, if a local palette is in use, we immediately write it.
            If (Not m_allFrames(i).usesGlobalPalette) Then
                WriteOptimizedAGIFToFile_Internal = WriteGIF_PaletteColors(dstStream, m_allFrames(i).palNumColors, m_allFrames(i).framePalette)
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
            dstStream.WriteByte lzwCodeSize
            
            'Next is the image bitstream!  Encoding happens elsewhere; we just pass the stream to them
            ' and let them encode away.
            ImageFormats_GIF_LZW.CompressLZW dstStream, VarPtr(m_allFrames(i).pixelData(0, 0)), m_allFrames(i).rectOfInterest.Width * m_allFrames(i).rectOfInterest.Height, lzwCodeSize + 1, True
            
            'All that's left for this frame is to explicitly terminate the black
            dstStream.WriteByte 0
            
        '/end duplicate/empty frames
        End If
        
    'Continue with the next frame!
    Next i
    
    'With all frames written, we can write the trailer and exit!
    ' (This is a magic number from the spec: https://www.w3.org/Graphics/GIF/spec-gif89a.txt)
    dstStream.WriteByte &H3B
    
    'Work complete!
    WriteOptimizedAGIFToFile_Internal = True
    
    Exit Function
    
ExportGIFError:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ", " & Err.Description
    WriteOptimizedAGIFToFile_Internal = False
    
End Function

Private Function WriteGIF_FileHeader(ByRef dstStream As pdStream, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal gpColorCount As Long, ByRef gpColors() As RGBQuad, Optional ByVal gpBackgroundIndex As Long = 0) As Boolean

    'For detailed GIF format info, see http://giflib.sourceforge.net/whatsinagif/bits_and_bytes.html
    ' PD doesn't attempt to support every esoteric GIF option (for example, we deliberately omit support
    ' for interlaced GIFs because they're pointless in the 21st century).  Instead, we focus on two things:
    ' *minimum filesize* and *maximum encoding efficiency*.
    
    'GIF header is fixed, 3-bytes for "GIF" ID, 3-bytes for version (always "89a" for animation)
    dstStream.WriteString_ASCII "GIF89a"
    
    'Next, the "logical screen descriptor".  This is always 7 bytes long:
    ' 4 bytes - unsigned short width + height
    dstStream.WriteIntU imgWidth
    dstStream.WriteIntU imgHeight
    
    'Now, an unpleasant packed 8-bit field
    ' 1 bit - global color table GCT exists (always TRUE in PD, for both static and animated images)
    ' 3 bits - GCT size N (describing 2 ^ n-1 colors in the palette)
    ' 1 bit - palette is sorted by importance (no longer used, always 0 from PD even though PD produces sorted palettes just fine)
    ' 3 bits - GCT size N again (technically the first field is bit-depth, but they're the same when using a global palette)
    Dim tmpBitField As Byte
    tmpBitField = &H80 'global palette exists
    
    Dim pow2forGP As Long
    pow2forGP = Pow2FromColorCount(gpColorCount) - 1
    tmpBitField = tmpBitField Or (pow2forGP * &H10) Or pow2forGP
    dstStream.WriteByte tmpBitField
    
    'Background color index (typically 0 by PD)
    dstStream.WriteByte gpBackgroundIndex
    
    'Aspect ratio using a bizarre old formula that hasn't been relevant for 30+ years (always 0 by PD)
    dstStream.WriteByte 0
    
    'Next comes the global color table/palette, in RGB order
    WriteGIF_PaletteColors dstStream, gpColorCount, gpColors
    
    'The image header is now complete.
    WriteGIF_FileHeader = True

End Function

'Write a GIF palette to a pdStream object.
'
'IMPORTANT NOTE: this function will automatically pad the palette to the nearest power-of-two,
' as required by the GIF spec.  Padded RGB entries will be (0, 0, 0) per convention.
Private Function WriteGIF_PaletteColors(ByRef dstStream As pdStream, ByVal numColors As Long, ByRef palColors() As RGBQuad) As Boolean
    
    'To improve performance, we pre-swizzle palette colors into a byte array, then dump the whole array to file at once.
    ' Besides a perf boost, this also lets us pre-allocate an array padded to the proper nearest power-of-two.
    Dim numColorsPadded As Long
    numColorsPadded = 2 ^ Pow2FromColorCount(numColors)
    
    'Use raw bytes (instead of an RGB UDT) to avoid padding considerations
    Dim rgbColors() As Byte
    ReDim rgbColors(0 To numColorsPadded * 3 - 1) As Byte
    
    'Iterate over all colors in the source palette (NOT the padded color count, which may be larger)
    ' and swizzle them into the temp array.
    Dim i As Long
    For i = 0 To numColors - 1
        rgbColors(i * 3) = palColors(i).Red
        rgbColors(i * 3 + 1) = palColors(i).Green
        rgbColors(i * 3 + 2) = palColors(i).Blue
    Next i
    
    'Swizzle complete!  Dump the result to stream.
    dstStream.WriteBytesFromPointer VarPtr(rgbColors(0)), numColorsPadded * 3
    WriteGIF_PaletteColors = True
    
End Function

'Write an (optional) graphics control extension block.
Private Function WriteGIF_GCE(ByRef dstStream As pdStream, ByVal transparencyActive As Boolean, Optional ByVal disposalOp As PD_GifDisposal = gd_Unknown, Optional ByVal trnsIndex As Long = 0, Optional ByVal delayTimeInCS As Long = 0) As Boolean

    'This function writes an optional "Graphics Control Extension" block.
    ' It is a fixed-size struct; we only need it for describing transparency or animation delay time.
    ' If *NEITHER* animation delay time nor transparency are required, you do not need to write
    ' this block.  (Skip it and save a few bytes!)
    
    'First three bytes are fixed ("introducer", "label", block size)
    dstStream.WriteByte &H21
    dstStream.WriteByte &HF9
    dstStream.WriteByte 4
    
    'Next is an annoying packed field:
    ' - 3 bits reserved (0)
    ' - 3 bits disposal method (unused for static GIFs)
    ' - 1 bit user-input flag (ignored)
    ' - 1 bit transparent color flag
    
    'Place disposal op bits (note that this is just a value [0, 3] matching the corresponding GIF constants)
    Dim bitField As Long
    bitField = disposalOp * &H4&
    
    '...and merge transparency too
    If transparencyActive Then bitField = bitField Or 1
    dstStream.WriteByte bitField
    
    'Next is 2-byte delay time in centiseconds (uuuuugh, also this is obviously ignored by static GIFs)
    dstStream.WriteIntU delayTimeInCS
    
    'Next is 1-byte transparent color index (almost always 0 in PD, but caller can specify otherwise)
    dstStream.WriteByte trnsIndex
    
    'Next is 1-byte block terminator (always 0)
    dstStream.WriteByte 0
    
    'This Graphics Control Extension block is complete.
    WriteGIF_GCE = True
    
End Function

'GIFs implement a variety of settings based on the nearest power-of-two to the
' image/frame palette's color count.
Private Function Pow2FromColorCount(ByVal cCount As Long) As Long
    Pow2FromColorCount = 1
    Do While ((2 ^ Pow2FromColorCount) < cCount)
        Pow2FromColorCount = Pow2FromColorCount + 1
    Loop
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String)
    PDDebug.LogAction "WARNING! pdGIF." & funcName & "() error: " & errDescription
End Sub
