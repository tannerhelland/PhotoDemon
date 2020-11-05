VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdMBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon MBM (Psion, Symbian) Image I/O
'Copyright 2020-2020 by Tanner Helland
'Created: 03/November/20
'Last updated: 05/November/20
'Last update: wrapped up provisional support for non-compressed images of every possible color-depth and
'             color model combination.  I *think* they all work but since testing images are hard to
'             come by, I'm disproportionately relying on a large collection of testing images from this
'             GitHub repo: https://github.com/SymbianSource/oss.FCL.sf.os.mmimaging
'
'This class is a WIP with no guarantee of ever reaching "release" state.
'
'For details, see https://github.com/tannerhelland/PhotoDemon/issues/341
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Enum MBM_Compression
    mbmc_None = 0
    mbmc_RLE_8bit = 1
    mbmc_RLE_12bit = 2
    mbmc_RLE_16bit = 3
    mbmc_RLE_24bit = 4
End Enum

#If False Then
    Private Const mbmc_None = 0, mbmc_RLE_8bit = 1, mbmc_RLE_12bit = 2, mbmc_RLE_16bit = 3, mbmc_RLE_24bit = 4
#End If

Private Type MBMFrame
    mbm_Offset As Long
    mbm_Length As Long
    mbm_HeaderLength As Long
    mbm_Width As Long
    mbm_Height As Long
    mbm_WidthTwips As Long
    mbm_HeightTwips As Long
    mbm_BPP As Long
    mbm_IsColor As Boolean
    mbm_FrameOK As Boolean  'Internal value; set to TRUE if this frame was successfully parsed
    mbm_PaletteSize As Long
    mbm_CompressionType As MBM_Compression
    mbm_DIB As pdDIB
End Type

'Frame collection is assembled as the underlying file is parsed
Private m_FrameCount As Long
Private m_Frames() As MBMFrame

'All parsing duties are handled by pdStream
Private m_Stream As pdStream

'Validate and load a candidate MBM file
Friend Function LoadMBM_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    LoadMBM_FromFile = False
    
    'Wrap a stream around the file
    Set m_Stream = New pdStream
    If (Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile)) Then
        InternalError "LoadMBM_FromFile", "couldn't start pdStream"
        Exit Function
    End If
    
    'Validate the first 12 bytes (which have fixed values)
    Dim isMBM As Boolean
    isMBM = (m_Stream.ReadLong_BE() = &H37000010)
    If isMBM Then isMBM = (m_Stream.ReadLong_BE() = &H42000010)
    If isMBM Then isMBM = (m_Stream.ReadLong_BE() = 0&)
    If isMBM Then isMBM = (m_Stream.ReadLong_BE() = &H39643947)
    
    'If the file has validated, continue parsing.  Next offset is trailer position.
    Dim trailerOffset As Long
    trailerOffset = m_Stream.ReadLong()
    
    'The trailer tells us how many images are in this file, and what their offsets are.
    ' First value is number of frames
    Dim origPosition As Long
    origPosition = m_Stream.GetPosition()
    
    m_Stream.SetPosition trailerOffset, FILE_BEGIN
    m_FrameCount = m_Stream.ReadLong()
    Debug.Print "frames: " & m_FrameCount
    
    If (m_FrameCount <= 0) Then
        InternalError "LoadMBM_FromFile", "bad frame count"
        Exit Function
    End If
    
    'Retrieve all frame offsets
    ReDim m_Frames(0 To m_FrameCount - 1) As MBMFrame
    Dim i As Long
    For i = 0 To m_FrameCount - 1
        m_Frames(i).mbm_Offset = m_Stream.ReadLong()
    Next i
    
    'The first offset should always be 20 (0x14)
    If (m_Frames(0).mbm_Offset <> 20) Then
        InternalError "LoadMBM_FromFile", "bad first offset"
        Exit Function
    End If
    
    'Start loading all frames.  (Failures will simply skip ahead to the next frame; because of this,
    ' you need to double-check that frame data is good before
    For i = 0 To m_FrameCount - 1
        
        With m_Frames(i)
            
            'Start by pointing the stream at this frame's offset
            m_Stream.SetPosition .mbm_Offset, FILE_BEGIN
            
            'Retrieve total frame length
            .mbm_Length = m_Stream.ReadLong()
            If (.mbm_Length <= 40) Then GoTo NextFrame
            Debug.Print m_Frames(i).mbm_Length
            
            'Retrieve header size (should always be 40; consider validating this?)
            .mbm_HeaderLength = m_Stream.ReadLong()
            If (.mbm_HeaderLength < 40) Then GoTo NextFrame
            Debug.Print m_Frames(i).mbm_HeaderLength
            
            'Retrieve x/y dimensions in both pixels and twips
            .mbm_Width = m_Stream.ReadLong()
            .mbm_Height = m_Stream.ReadLong()
            .mbm_WidthTwips = m_Stream.ReadLong()   'In the wild, these appear to frequently (always?) be 0
            .mbm_HeightTwips = m_Stream.ReadLong()
            If (.mbm_Width <= 0) Or (.mbm_Height <= 0) Then GoTo NextFrame
            Debug.Print .mbm_Width, .mbm_Height, .mbm_WidthTwips, .mbm_WidthTwips
            
            'Retrieve color space data
            .mbm_BPP = m_Stream.ReadLong()
            If (.mbm_BPP <= 0) Then GoTo NextFrame
            .mbm_IsColor = (m_Stream.ReadLong <> 0)
            .mbm_PaletteSize = m_Stream.ReadLong()
            Debug.Print .mbm_BPP, .mbm_IsColor, .mbm_PaletteSize
            
            'Retrieve compression
            .mbm_CompressionType = m_Stream.ReadLong()
            Debug.Print "Compression", .mbm_CompressionType
            
            'Hypothetically we should be pointing at pixel data now
            If (m_Stream.GetPosition <> .mbm_Offset + .mbm_HeaderLength) Then m_Stream.SetPosition .mbm_Offset + .mbm_HeaderLength
            
            'Prep our target DIB and make it opaque (most images in this format will be < 32-bpp)
            Set .mbm_DIB = New pdDIB
            .mbm_DIB.CreateBlank .mbm_Width, .mbm_Height, 32, vbWhite, 255
            
            'Build palette?
            Dim srcPalette(0 To 255) As RGBQuad
            
            'Palettes appear to be hard-coded depending on color depth?
            Dim j As Long, r As Long, g As Long, b As Long
            If .mbm_IsColor Then
                
                'Fixed palettes have been reverse-engineered for 4- and 8-bpp; other bit-depths are TBD
                If (.mbm_BPP = 4) Then
                    FillQuadFromRGB srcPalette(0), 0, 0, 0
                    FillQuadFromRGB srcPalette(1), 85, 85, 85
                    FillQuadFromRGB srcPalette(2), 128, 0, 0
                    FillQuadFromRGB srcPalette(3), 128, 128, 0
                    FillQuadFromRGB srcPalette(4), 0, 128, 0
                    FillQuadFromRGB srcPalette(5), 255, 0, 0
                    FillQuadFromRGB srcPalette(6), 255, 255, 0
                    FillQuadFromRGB srcPalette(7), 0, 255, 0
                    FillQuadFromRGB srcPalette(8), 255, 0, 255
                    FillQuadFromRGB srcPalette(9), 0, 0, 255
                    FillQuadFromRGB srcPalette(10), 0, 255, 255
                    FillQuadFromRGB srcPalette(11), 128, 0, 128
                    FillQuadFromRGB srcPalette(12), 0, 0, 128
                    FillQuadFromRGB srcPalette(13), 0, 128, 128
                    FillQuadFromRGB srcPalette(14), 170, 170, 170
                    FillQuadFromRGB srcPalette(15), 255, 255, 255
                ElseIf (.mbm_BPP = 8) Then
                    Dim palIndex As Long
                    
                    For b = 0 To 5
                    For g = 0 To 5
                    For r = 0 To 5
                        
                        palIndex = b * 36 + g * 6 + r
                        
                        'For inexplicable reasons, the middle of their palette is filled with a bunch
                        ' of custom values, while the expected pattern colors continue 40 indices later.
                        If (palIndex <= 107) Then
                            FillQuadFromRGB srcPalette(palIndex), r * 51, g * 51, b * 51
                        Else
                            FillQuadFromRGB srcPalette(palIndex + 40), r * 51, g * 51, b * 51
                        End If
                        
                    Next r
                    Next g
                    Next b
                    
                    'We now need to fill a bunch of custom values in the middle of this palette because
                    ' fuck whoever designed this format! :p
                    FillQuadFromRGBHex srcPalette(108), &H111111
                    FillQuadFromRGBHex srcPalette(109), &H222222
                    FillQuadFromRGBHex srcPalette(110), &H444444
                    FillQuadFromRGBHex srcPalette(111), &H555555
                    FillQuadFromRGBHex srcPalette(112), &H777777
                    FillQuadFromRGBHex srcPalette(113), &H110000
                    FillQuadFromRGBHex srcPalette(114), &H220000
                    FillQuadFromRGBHex srcPalette(115), &H440000
                    FillQuadFromRGBHex srcPalette(116), &H550000
                    FillQuadFromRGBHex srcPalette(117), &H770000
                    FillQuadFromRGBHex srcPalette(118), &H1100&
                    FillQuadFromRGBHex srcPalette(119), &H2200&
                    FillQuadFromRGBHex srcPalette(120), &H4400&
                    FillQuadFromRGBHex srcPalette(121), &H5500&
                    FillQuadFromRGBHex srcPalette(122), &H7700&
                    FillQuadFromRGBHex srcPalette(123), &H11&
                    FillQuadFromRGBHex srcPalette(124), &H22&
                    FillQuadFromRGBHex srcPalette(125), &H44&
                    FillQuadFromRGBHex srcPalette(126), &H55&
                    FillQuadFromRGBHex srcPalette(127), &H77&
                    FillQuadFromRGBHex srcPalette(128), &H88&
                    FillQuadFromRGBHex srcPalette(129), &HAA&
                    FillQuadFromRGBHex srcPalette(130), &HBB&
                    FillQuadFromRGBHex srcPalette(131), &HDD&
                    FillQuadFromRGBHex srcPalette(132), &HEE&
                    FillQuadFromRGBHex srcPalette(133), &H8800&
                    FillQuadFromRGBHex srcPalette(134), &HAA00&
                    FillQuadFromRGBHex srcPalette(135), &HBB00&
                    FillQuadFromRGBHex srcPalette(136), &HDD00&
                    FillQuadFromRGBHex srcPalette(137), &HEE00&
                    FillQuadFromRGBHex srcPalette(138), &H880000
                    FillQuadFromRGBHex srcPalette(139), &HAA0000
                    FillQuadFromRGBHex srcPalette(140), &HBB0000
                    FillQuadFromRGBHex srcPalette(141), &HDD0000
                    FillQuadFromRGBHex srcPalette(142), &HEE0000
                    FillQuadFromRGBHex srcPalette(143), &H888888
                    FillQuadFromRGBHex srcPalette(144), &HAAAAAA
                    FillQuadFromRGBHex srcPalette(145), &HBBBBBB
                    FillQuadFromRGBHex srcPalette(146), &HDDDDDD
                    FillQuadFromRGBHex srcPalette(147), &HEEEEEE
                End If
            
            'Grayscale palettes follow standard rules
            Else
                Dim numShades As Long, scaleFactor As Long
                numShades = 2 ^ .mbm_BPP
                scaleFactor = 255 / (numShades - 1)
                For j = 0 To numShades - 1
                    g = j * scaleFactor
                    srcPalette(j).Red = g
                    srcPalette(j).Green = g
                    srcPalette(j).Blue = g
                    srcPalette(j).Alpha = 255
                Next j
            End If
            
            'How we retrieve pixels depends on color-depth, obviously.  We're going to do this
            ' in two passes to simplify the process of handling messy compression and color-depth
            ' complications.
            Dim pxWidth As Long, xFinal As Long, pxBitCount As Long
            pxWidth = .mbm_Width
            xFinal = pxWidth - 1
            pxBitCount = .mbm_BPP
            
            Dim pxScanline() As Byte, scanlineSize As Long
            If (pxBitCount = 1) Then
                scanlineSize = (pxWidth + 7) \ 8
            ElseIf (pxBitCount = 2) Then
                scanlineSize = (pxWidth + 3) \ 4
            ElseIf (pxBitCount = 4) Then
                scanlineSize = (pxWidth + 1) \ 2
            ElseIf (pxBitCount = 8) Then
                scanlineSize = pxWidth
            ElseIf (pxBitCount = 12) Then
                scanlineSize = pxWidth * 2  '(pxWidth * 3) \ 2
            ElseIf (pxBitCount = 16) Then
                scanlineSize = pxWidth * 2
            ElseIf (pxBitCount = 24) Then
                scanlineSize = pxWidth * 3
            ElseIf (pxBitCount = 32) Then
                scanlineSize = pxWidth * 4
            Else
                InternalError "LoadMBM_FromFile", "bad bitcount: " & pxBitCount
                GoTo NextFrame
            End If
            
            If (scanlineSize <= 0) Then
                InternalError "LoadMBM_FromFile", "bad scanline size: " & scanlineSize
                GoTo NextFrame
            End If
            
            'Next, I've encountered some messy behavior with line alignment.  For uncompressed
            ' files with bit-depths < 24, line alignment appears to be on 4-byte boundaries.
            ' For higher bit-depths, the results are much weirder.
            If (.mbm_BPP < 24) Then
                scanlineSize = (scanlineSize + 3) And &HFFFFFFFC
            
            'This one makes *zero* sense to me, but it produces useable images from symbian's
            ' massive .mbm collection... so who am I to doubt
            ElseIf (.mbm_BPP = 24) Then
                scanlineSize = ((scanlineSize + 11) \ 12) * 12
            
            '32-bpp is a mess; on some images, it produces useable images with 4-byte alignment;
            ' on others, it appears to require 8-bpp boundaries?  I do not know how to differentiate these,
            ' or if the results are even meant to be useable (and not just test images)
            ElseIf (.mbm_BPP > 24) Then
                scanlineSize = (scanlineSize + 3) And &HFFFFFFFC
                'scanlineSize = (scanlineSize + 7) And &HFFFFFFF8
            End If
            
            If (.mbm_CompressionType = mbmc_None) Then
                
                
            End If
            'scanlineSize = ((scanlineSize + (LINE_SIZE - 1)) \ LINE_SIZE) * LINE_SIZE
            'scanlineSize = (scanlineSize + 3) And &HFFFFFFFC
            'scanlineSize = ((scanlineSize + 7) \ 8) * 8
            
            'scanlineSize = (scanlineSize + 3) And &HFFFFFFFC
            'Debug.Print pxWidth, scanlineSize
            
            'The icon's size appears to be valid.  Initialize the destination DIB and a temporary
            ' array for holding raw scanline data (before it's proceed to 32-bpp).
            ReDim pxScanline(0 To scanlineSize - 1) As Byte
            
            'Some bit-depths are easier to handle with lookup tables.  (In effect, we pre-convert
            ' each scanline to 8-bpp.)
            Dim preConvert() As Byte, bitFlags() As Byte
            If (pxBitCount < 8) Then
            
                ReDim preConvert(0 To pxWidth - 1) As Byte
                
                If (pxBitCount = 1) Then
                    ReDim bitFlags(0 To 7) As Byte
                    bitFlags(0) = 2 ^ 7
                    bitFlags(1) = 2 ^ 6
                    bitFlags(2) = 2 ^ 5
                    bitFlags(3) = 2 ^ 4
                    bitFlags(4) = 2 ^ 3
                    bitFlags(5) = 2 ^ 2
                    bitFlags(6) = 2 ^ 1
                    bitFlags(7) = 1
                ElseIf (pxBitCount = 2) Then
                    ReDim bitFlags(0 To 3) As Byte
                    bitFlags(0) = 2 ^ 6
                    bitFlags(1) = 2 ^ 4
                    bitFlags(2) = 2 ^ 2
                    bitFlags(3) = 1
                End If
            
            End If
            
            'Process each scanline in turn
            Dim x As Long, y As Long, alphaFound As Boolean
            alphaFound = False
            
            Dim tmpSA1D As SafeArray1D, dstPixels() As RGBQuad
            Dim srcByte As Byte, numPixelsProcessed As Long
            
            Dim numScanlines As Long
            numScanlines = .mbm_Height - 1
            
            For y = 0 To numScanlines
                
                'TODO: decompress and copy contents into scanline buffer
                If (.mbm_CompressionType = mbmc_None) Then
                    m_Stream.ReadBytesToBarePointer VarPtr(pxScanline(0)), scanlineSize
                Else
                    Debug.Print "compression isn't ready"
                End If
                
                'With the line decompressed, we can now convert it to RGBA
                
                'For low bit-depth images, immediately upsample to 8-bpp
                If (pxBitCount < 8) Then
                    
                    numPixelsProcessed = 0
                    If (pxBitCount = 1) Then
                        
                        For x = 0 To scanlineSize - 1
                            
                            srcByte = pxScanline(x)
                            
                            'Ignore empty bytes at the end of each scanline
                            For j = 0 To 7
                                If (numPixelsProcessed <= xFinal) Then
                                    If (bitFlags(7 - j) = (srcByte And bitFlags(7 - j))) Then preConvert(numPixelsProcessed) = 1 Else preConvert(numPixelsProcessed) = 0
                                    numPixelsProcessed = numPixelsProcessed + 1
                                End If
                            Next j
                            
                        Next x
                    
                    ElseIf (pxBitCount = 2) Then
                    
                        For x = 0 To scanlineSize - 1
                            srcByte = pxScanline(x)
                            For j = 0 To 3
                                If (numPixelsProcessed <= xFinal) Then
                                    preConvert(numPixelsProcessed) = (srcByte \ bitFlags(3 - j)) And &H3
                                    numPixelsProcessed = numPixelsProcessed + 1
                                End If
                            Next j
                        Next x
                    
                    ElseIf (pxBitCount = 4) Then
                    
                        For x = 0 To scanlineSize - 1
                            
                            srcByte = pxScanline(x)
                            preConvert(numPixelsProcessed) = srcByte And &HF
                            numPixelsProcessed = numPixelsProcessed + 1
                            
                            If (numPixelsProcessed <= xFinal) Then
                                preConvert(numPixelsProcessed) = (srcByte \ 16) And &HF
                                numPixelsProcessed = numPixelsProcessed + 1
                            End If
                            
                        Next x
                    
                    End If
                
                '/end pre-processing of < 8-bpp images
                End If
                
                'Point a destination array at the target DIB
                .mbm_DIB.WrapRGBQuadArrayAroundScanline dstPixels, tmpSA1D, y
                
                'Used on 12, 16-bpp images to avoid byte overflow issues when masking
                Dim tmpInteger As Long
                
                'Process each pixel in turn
                For x = 0 To xFinal
                
                    Select Case pxBitCount
                    
                        Case 1, 2, 4
                            dstPixels(x) = srcPalette(preConvert(x))
                            
                        Case 8
                            dstPixels(x) = srcPalette(pxScanline(x))
                        
                        '12-bpp uses 0-4-4-4 masking (network byte order) on 16-bit boundaries
                        Case 12
                            tmpInteger = CLng(pxScanline(x * 2)) + CLng(pxScanline(x * 2 + 1)) * 256
                            dstPixels(x).Red = ((tmpInteger And &HF00&) \ (2 ^ 8)) * 16
                            dstPixels(x).Green = ((tmpInteger And &HF0&) \ (2 ^ 4)) * 16
                            dstPixels(x).Blue = (tmpInteger And &HF&) * 16
                            
                        '16-bpp appears to use 5-6-5 masking (network byte order)
                        Case 16
                            tmpInteger = CLng(pxScanline(x * 2)) + CLng(pxScanline(x * 2 + 1)) * 256
                            dstPixels(x).Red = ((tmpInteger And &HF800&) \ (2 ^ 11)) * 8
                            dstPixels(x).Green = ((tmpInteger And &H7E0&) \ (2 ^ 5)) * 4
                            dstPixels(x).Blue = (tmpInteger And &H1F&) * 8
                            dstPixels(x).Alpha = 255
                            
                        Case 24
                            dstPixels(x).Blue = pxScanline(x * 3)
                            dstPixels(x).Green = pxScanline(x * 3 + 1)
                            dstPixels(x).Red = pxScanline(x * 3 + 2)
                            dstPixels(x).Alpha = 255
                            
                        Case 32
                            GetMem4_Ptr VarPtr(pxScanline(x * 4)), VarPtr(dstPixels(x))
                            If (dstPixels(x).Alpha > 0) Then alphaFound = True
                    
                    End Select
                
                Next x
            
            Next y
            
            'Release our unsafe DIB array wrapper
            .mbm_DIB.UnwrapRGBQuadArrayFromDIB dstPixels
            
            'Premultiply our finished alpha channel
            .mbm_DIB.SetAlphaPremultiplication True
            
            'Mark this frame as successful!
            .mbm_FrameOK = True
            
        End With
    
NextFrame:
    Next i
    
    'With all frames parsed, we now need to construct a new layer for each frame in the
    ' destination pdImage object.
    
    'Start by finding the largest frame in the file; we'll use this for our master image dimensions.
    Dim maxWidth As Long, maxHeight As Long
    For i = 0 To m_FrameCount - 1
        If m_Frames(i).mbm_FrameOK Then
            If (Not m_Frames(i).mbm_DIB Is Nothing) Then
                maxWidth = PDMath.Max2Int(maxWidth, m_Frames(i).mbm_DIB.GetDIBWidth())
                maxHeight = PDMath.Max2Int(maxHeight, m_Frames(i).mbm_DIB.GetDIBHeight())
            End If
        End If
    Next i
    
    'Ensure both width and height are non-zero
    If (maxWidth > 0) And (maxHeight > 0) Then
        
        'We have enough data to produce a usable image.  Start by initializing basic pdImage attributes.
        'dstImage.SetOriginalFileFormat PDIF_ICO
        dstImage.Width = maxWidth
        dstImage.Height = maxHeight
        dstImage.SetDPI 96#, 96#
        
        'Next, we want to figure out which layer to activate + make visible.  This should be the...
        ' 1) largest image in the file...
        ' 2) ...that also has the highest bit-depth
        Dim activeLayerIndex As Long, highestBitDepth As Long
        For i = 0 To m_FrameCount - 1
            If m_Frames(i).mbm_FrameOK And (Not m_Frames(i).mbm_DIB Is Nothing) Then
                If (m_Frames(i).mbm_DIB.GetDIBWidth = maxWidth) And (m_Frames(i).mbm_DIB.GetDIBHeight = maxHeight) Then
                
                    'This layer matches the largest layer size we have so far.  If it *also* has the
                    ' highest bit-depth, flag it as the new active index.
                    If (m_Frames(i).mbm_BPP > highestBitDepth) Then
                        highestBitDepth = m_Frames(i).mbm_BPP
                        activeLayerIndex = i
                    End If
                
                End If
            End If
        Next i
        
        'Next, we want to produce a pdLayer object for each valid frame
        Dim tmpLayer As pdLayer, newLayerID As Long
        
        For i = 0 To m_FrameCount - 1
            
            'Skip frames that didn't validate during loading
            If m_Frames(i).mbm_FrameOK And (Not m_Frames(i).mbm_DIB Is Nothing) Then
                
                'Ensure alpha is premultiplied
                If (Not m_Frames(i).mbm_DIB.GetAlphaPremultiplication()) Then m_Frames(i).mbm_DIB.SetAlphaPremultiplication True
                
                'Prep a new layer object and initialize it with the image bits we've retrieved
                newLayerID = dstImage.CreateBlankLayer()
                Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                tmpLayer.InitializeNewLayer PDL_Image, g_Language.TranslateMessage("Layer %1", i + 1), m_Frames(i).mbm_DIB
                
                'If this layer's dimensions match the largest layer, make this layer visible.
                ' (All other layers will be hidden, by default.)
                tmpLayer.SetLayerVisibility (i = activeLayerIndex)
                If tmpLayer.GetLayerVisibility Then dstImage.SetActiveLayerByID newLayerID
                
                'Notify the layer of new changes, so it knows to regenerate internal caches on next access
                tmpLayer.NotifyOfDestructiveChanges
                
            End If
        
        Next i
        
        'Notify the image of destructive changes, so it can rebuild internal caches
        dstImage.NotifyImageChanged UNDO_Everything
        dstImage.SetActiveLayerByIndex activeLayerIndex
        
        'Return success
        LoadMBM_FromFile = True
        
    Else
        LoadMBM_FromFile = False
        InternalError "LoadMBM_FromFile", "no frames with non-zero width/height"
        Exit Function
    End If
    
End Function

Private Sub FillQuadFromRGB(ByRef dstQuad As RGBQuad, ByVal r As Long, ByVal g As Long, ByVal b As Long)
    dstQuad.Red = r
    dstQuad.Green = g
    dstQuad.Blue = b
    dstQuad.Alpha = 255
End Sub

Private Sub FillQuadFromRGBHex(ByRef dstQuad As RGBQuad, ByVal hexValue As Long)
    dstQuad.Red = Colors.ExtractBlue(hexValue)
    dstQuad.Green = Colors.ExtractGreen(hexValue)
    dstQuad.Blue = Colors.ExtractRed(hexValue)
    dstQuad.Alpha = 255
End Sub

Private Sub InternalError(ByRef fncName As String, ByRef errDetails As String)
    PDDebug.LogAction "WARNING!  Error in pdMBM." & fncName & ": " & errDetails
End Sub
