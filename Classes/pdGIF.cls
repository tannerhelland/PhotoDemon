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
'Copyright 2001-2021 by Tanner Helland
'Created: 4/15/01
'Last updated: 23/October/21
'Last update: wrap up work on homebrew static GIF exporter
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
' produced this way are often an order of magnitude (or more) smaller than GIFs
' produced by a naive encoder.
'
'Note that the optimization steps are specifically written in an export library
' agnostic way.  PD internally stores the results of all optimizations, then just hands
' the optimized frames off to an encoder at the end of the process.  Historically PD used
' FreeImage for animated GIF encoding, but FreeImage has a number of shortcomings
' (including woeful performance and writing larger GIFs than is necessary), so in 2021
' I moved to an in-house LZW encoder based off the classic UNIX "compress" tool.
' The LZW encoder has different licensing considerations, so I've kept it in a separate
' file (ImageFormats_GIF_LZW).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Copy of the 8-bpp (palettized) array representing the original image.  May be reduced further
' during LZW compression.
Private m_ImgBytes() As Byte

'Final image palette.  Alpha element is important, as there may be a 0-alpha entry for GIF transparency.
Private m_ImgPalette() As RGBQuad

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
    
    'For detailed GIF format info, see http://giflib.sourceforge.net/whatsinagif/bits_and_bytes.html
    
    'GIF header is fixed, 3-bytes for "GIF" ID, 3-bytes for version (always "89a" for PD)
    dstStream.WriteString_ASCII "GIF89a"
    
    'Next, the "logical screen descriptor".  This is always 7 bytes long:
    ' 4 bytes - unsigned short width + height
    dstStream.WriteIntU srcPDImage.Width
    dstStream.WriteIntU srcPDImage.Height
    
    'Now, an unpleasant packed 8-bit field
    ' 1 bit - global color table GCT exists (always TRUE in PD)
    ' 3 bits - GCT size N (describing 2 ^ n-1 colors in the palette)
    ' 1 bit - palette is sorted by importance (no longer used, always 0 from PD even though PD produces sorted palettes just fine)
    ' 3 bits - GCT size N again (technically the first field is bit-depth, but they're the same when using a global palette)
    Dim tmpBitField As Byte
    tmpBitField = &H80 'global palette exists
    
    Dim pow2forGP As Long
    pow2forGP = Pow2FromColorCount(numColorsInPalette) - 1
    tmpBitField = tmpBitField Or (pow2forGP * &H10) Or pow2forGP
    dstStream.WriteByte tmpBitField
    
    'Background color index; basically we attempt to find the nearest match to the matte color, if any.
    ' (There's no guarantee the matte color will be used in the image - it may have just been used to
    ' composite semi-transparent pixels - but this at least gives us something "close".)
    Dim bkgdIndex As Long
    bkgdIndex = Palettes.GetNearestIndexRGB(m_ImgPalette, cParams.GetLong("gif-backcolor", vbWhite))
    dstStream.WriteByte bkgdIndex
    
    'Aspect ratio using a bizarre old formula, always 0 by PD
    dstStream.WriteByte 0
    
    'Next comes the global color table/palette, in RGB order.  Importantly, note that the GIF spec
    ' requires all palette color counts to be a power of 2.  (It does this because palette color
    ' count is stored in 3-bits, ugh.)  Any unused entries are ignored, but by convention are
    ' usually left as black; we do the same here.
    numColorsInPalette = 2 ^ Pow2FromColorCount(numColorsInPalette)
    If (UBound(m_ImgPalette) <> numColorsInPalette - 1) Then ReDim Preserve m_ImgPalette(0 To numColorsInPalette - 1) As RGBQuad
    
    Dim i As Long
    For i = 0 To numColorsInPalette - 1
        With m_ImgPalette(i)
            dstStream.WriteByte .Red
            dstStream.WriteByte .Green
            dstStream.WriteByte .Blue
        End With
    Next i
    
    'The image header is now complete.
    
    'If the image contains transparency, we now need to write an optional block to flag the
    ' transparent palette index.
    If (m_ImgPalette(0).Alpha = 0) Then
        
        'This block is the "Graphics Control Extension".
        ' It is a fixed-size struct; we only need it for describing transparency.
        
        'First three bytes are fixed ("introducer", "label", size)
        dstStream.WriteByte &H21
        dstStream.WriteByte &HF9
        dstStream.WriteByte &H4
        
        'Next is an annoying packed field:
        ' - 3 bits reserved (0)
        ' - 3 bits disposal method (unused for static GIFs)
        ' - 1 bit user-input flag (ignored)
        ' - 1 bit transparent color flag
        '
        'For static GIFs, we only need the transparency flag
        dstStream.WriteByte 1
        
        'Next is 2-byte delay time (unused in static GIFs)
        dstStream.WriteIntU 0
        
        'Next is 1-byte transparent color index, which is always 0 in PD
        dstStream.WriteByte 0
        
        'Next is 1-byte block terminator (always 0)
        dstStream.WriteByte 0
        
        'Graphics Control Extension is done
    
    End If
    
    'Next up is an "image descriptor", basically a frame header
    
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
    lzwCodeSize = pow2forGP + 1
    If (lzwCodeSize < 2) Then lzwCodeSize = 2
    dstStream.WriteByte lzwCodeSize
    
    'Next is the image bitstream!  Encoding happens elsewhere; we just pass the stream to them
    ' and let them encode away.
    ProgressBars.SetProgBarVal 5
    ImageFormats_GIF_LZW.CompressLZW dstStream, VarPtr(m_ImgBytes(0, 0)), srcPDImage.Width * srcPDImage.Height, lzwCodeSize + 1
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

'Given a param string generated by the Export GIF dialog, apply any GIF pre-processing steps.
' Works for both export preview and actual export prep steps (depending on the value of usePreviewMode).
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
        DIBs.ApplyAlphaCutoff_Ex srcDIB, trnsValues, 127
        DIBs.ApplyBinaryTransparencyTable srcDIB, trnsValues, gifBackgroundColor
    
    'Normal GIF behavior (threshold alpha into "fully transparent" or "fully opaque")
    Else
        DIBs.ApplyAlphaCutoff_Ex srcDIB, trnsValues, gifAlphaCutoff
        DIBs.ApplyBinaryTransparencyTable srcDIB, trnsValues, gifBackgroundColor
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
