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
'PhotoDemon MBM (Psion, Symbian) Image Importer (export not currently supported)
'Copyright 2020-2026 by Tanner Helland
'Created: 03/November/20
'Last updated: 30/December/20
'Last update: add support for AIF files, including embedded masks
'
'This class imports legacy Symbian "MBM" files.  All well-defined color-depths and compression types
' are supported.  (I say "well-defined" because there are many MBM files in the wild with non-standard
' color-depth, grayscale/color, and compression values, and PD is only guaranteed to load files
' produced by official SDKs.)
'
'For details on how this class came to be, see https://github.com/tannerhelland/PhotoDemon/issues/341
'
'Currently this class only supports importing MBM images.  Export support could be added with
' minimal investment code-wise, but there are significant UI implications (especially localization)
' for a matching export dialog, so I do not currently have plans to add this.
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
    mbmc_RLE_32bit = 5  'IMPORTANT: this compression type is not documented anywhere online, but it exists in
                        ' 32-bpp files from the Symbian test suite.  It's behavior seems self-explanatory, so
                        ' I've gone ahead and added support for it.
                        
                        'ALSO IMPORTANT: symbian's image test suite has images with a compression value of 6;
                        ' this is not documented anywhere, so PD treats other values as errors.
End Enum

#If False Then
    Private Const mbmc_None = 0, mbmc_RLE_8bit = 1, mbmc_RLE_12bit = 2, mbmc_RLE_16bit = 3, mbmc_RLE_24bit = 4, mbmc_RLE_32bit = 5
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
    mbm_FrameOK As Boolean  'Internal value; set to TRUE if this frame was successfully parsed.  (FALSE tells us to discard the frame because something about it is broken.)
    mbm_PaletteSize As Long
    mbm_CompressionType As MBM_Compression
    mbm_DIB As pdDIB
End Type

'Frame collection is assembled as the underlying file is parsed
Private m_FrameCount As Long
Private m_Frames() As MBMFrame

'AIF files are a minor MBM variant; they require slightly different parsing behavior
Private m_IsAIF As Boolean

'All parsing duties are handled by pdStream
Private m_Stream As pdStream

'Only valid *after* an image has been loaded.  Because an image can contain multiple frames
' (of varying color depth), this will return the largest-found color depth, by design.
Friend Function GetColorDepth() As Long
    If (m_FrameCount > 0) Then
        Dim i As Long
        For i = 0 To m_FrameCount - 1
            If (m_Frames(i).mbm_BPP > GetColorDepth) Then GetColorDepth = m_Frames(i).mbm_BPP
        Next i
    End If
End Function

'See if a file is MBM.  File extension is *not* relevant.
Friend Function IsFileMBM(ByRef srcFile As String, Optional ByVal calledInternally As Boolean = False) As Boolean

    IsFileMBM = False
    
    'Wrap a stream around the file, as necessary
    If (Not calledInternally) Then Set m_Stream = New pdStream
    
    Dim okToProceed As Boolean
    okToProceed = True
    If (Not calledInternally) Then okToProceed = m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile)
    
    If okToProceed Then
        
        'Validate the first 12 bytes (which have fixed values)
        IsFileMBM = (m_Stream.ReadLong_BE() = &H37000010)
        If IsFileMBM Then IsFileMBM = (m_Stream.ReadLong_BE() = &H42000010)
        If IsFileMBM Then IsFileMBM = (m_Stream.ReadLong_BE() = 0&)
        If IsFileMBM Then IsFileMBM = (m_Stream.ReadLong_BE() = &H39643947)
        
        'If the file is *not* MBM, reset the stream pointer and check for an AIF signature instead
        If (Not IsFileMBM) Then
        
            m_Stream.SetPosition 0, FILE_BEGIN
            IsFileMBM = (m_Stream.ReadLong_BE() = &H37000010)
            If IsFileMBM Then IsFileMBM = (m_Stream.ReadLong_BE() = &H6A000010)
            
            'The next 8 bytes are App ID a checksum; we can skip them
            m_Stream.SetPosition 8, FILE_CURRENT
            
            'If the previous checks succeeded, this is an AIF file with potential MBM contents
            m_IsAIF = IsFileMBM
            
        End If
        
    End If
    
    'Free the stream before exiting
    If (Not calledInternally) Then Set m_Stream = Nothing
    
End Function

'Only valid *after* an image has been loaded, and returns data for the first frame *only*
Friend Function IsMBMGrayscale() As Boolean
    If (m_FrameCount > 0) Then IsMBMGrayscale = (Not m_Frames(0).mbm_IsColor)
End Function

'Validate and load a candidate MBM file
Friend Function LoadMBM_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    LoadMBM_FromFile = False
    
    'Wrap a stream around the file
    Set m_Stream = New pdStream
    If (Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile)) Then
        InternalError "LoadMBM_FromFile", "couldn't start pdStream"
        Exit Function
    End If
    
    'Validate the file
    If (Not Me.IsFileMBM(srcFile, True)) Then
        Set m_Stream = Nothing
        LoadMBM_FromFile = False
        Exit Function
    End If
    
    'Parsing this format requires nested loops in various places
    Dim i As Long, j As Long, k As Long
    
    'If the file has validated, continue parsing.  Next offset is trailer position.
    Dim trailerOffset As Long
    trailerOffset = m_Stream.ReadLong()
    
    'Note our current location, then jump to the start of the trailer
    Dim origPosition As Long
    origPosition = m_Stream.GetPosition()
    m_Stream.SetPosition trailerOffset, FILE_BEGIN
    
    'Trailer contents vary between traditional MBM and AIF files.  MBMs are simpler.
    'Regardless of file type, what we ultimately want to determine here is how many images
    ' are in this file, and what their offsets are.
    If m_IsAIF Then
        
        'Number of captions * 2.  Each caption has a 4-byte offset and a 2-byte language code;
        ' we simply want to skip all these bytes
        Dim numCaptions As Long
        numCaptions = m_Stream.ReadByte() \ 2
        m_Stream.SetPosition numCaptions * 6, FILE_CURRENT
        
        '*Now* we are pointing at the image directory.  One byte describes number of icons
        ' in the file * 2.
        m_FrameCount = m_Stream.ReadByte() \ 2
        
    Else
    
        'First trailer value is number of frames
        m_FrameCount = m_Stream.ReadLong()
        
    End If
    
    'I don't think there's an upper limit on frames, but some of the files in symbian's test suite
    ' contain extremely large values (e.g. 0x0fffffff), so it seems necessary to enforce some kind
    ' of limit here.  I can revisit as needed.
    If (m_FrameCount <= 0) Or (m_FrameCount >= 256) Then
        InternalError "LoadMBM_FromFile", "unexpected frame count: " & m_FrameCount
        Exit Function
    End If
    
    'Retrieve all frame offsets
    ReDim m_Frames(0 To m_FrameCount - 1) As MBMFrame
    
    If m_IsAIF Then
    
        'For each pair, retrieve a 4-byte offset, and a 2-byte width+height value.
        ' (This width/height value is effectively ignored; each MBM entry will repeat this
        '  information, and we want to use the embedded width/height values over this top-level
        '  OS cue.)
        For i = 0 To m_FrameCount - 1
            m_Frames(i).mbm_Offset = m_Stream.ReadLong()
            m_Stream.ReadInt    'Skip reported width/height
        Next i
        
    Else
    
        For i = 0 To m_FrameCount - 1
            m_Frames(i).mbm_Offset = m_Stream.ReadLong()
        Next i
        
    End If
    
    'In a normal MBM file, the first offset should always be 20 (0x14).  If it isn't, abandon ship.
    If (Not m_IsAIF) And (m_Frames(0).mbm_Offset <> 20) Then
        InternalError "LoadMBM_FromFile", "bad first offset"
        Exit Function
    End If
    
    'In AIF files, each frame is followed by a transparency mask.  We want to parse each mask and
    ' "merge" it with the preceding image to create a 32-bpp entry.
    Dim inMaskMode As Boolean
    inMaskMode = False
    
    Dim maskDIB As pdDIB
    If m_IsAIF Then Set maskDIB = New pdDIB
    
    'On AIF files, every-other-frame switches between targeting actual pixel values, and mask values.
    ' We'll use this reference to switch between the two.
    Dim targetDIB As pdDIB
    
    'Start loading all frames.  (Failures will simply skip ahead to the next frame; because of this,
    ' you need to double-check that frame data is good before
    For i = 0 To m_FrameCount - 1
        
        With m_Frames(i)
            
            'Start by pointing the stream at this frame's offset
            m_Stream.SetPosition .mbm_Offset, FILE_BEGIN
            
StartOfFrame:
                        
            'IMPORTANT NOTE ON AIF FILES.
            
            'AIF files behave a little differently from plain MBM files.  They contain additional
            ' frames with mask data for the associated application icons.  These frames are parsed
            ' correctly by this reader, but each transparency mask header will overwrite the header
            ' of the frame it is associated with (by design).  This is 100% okay because the frame
            ' itself no longer requires its header once its associated bitmap has been created.
            '
            'I'm just noting that here because the index i will *always* point to the index of
            ' the underlying bitmap data, but depending on the value of "inMaskMode", we may
            ' actually be parsing a mask instead of a a bitmap.
                        
            'Retrieve total frame length.  (It needs to be at least 40 bytes long.)
            .mbm_Length = m_Stream.ReadLong()
            If (.mbm_Length <= 40) Then GoTo NextFrame
            
            'Failsafe validation against end-of-stream failure
            If (.mbm_Length + m_Stream.GetPosition() > m_Stream.GetStreamSize()) Then
                InternalError "LoadMBM_FromFile", "frame extends beyond end of file"
                Exit Function
            End If
            
            'Retrieve header size (should always be 40)
            .mbm_HeaderLength = m_Stream.ReadLong()
            If (.mbm_HeaderLength < 40) Then GoTo NextFrame
            
            'Retrieve x/y dimensions in both pixels and twips
            .mbm_Width = m_Stream.ReadLong()
            .mbm_Height = m_Stream.ReadLong()
            .mbm_WidthTwips = m_Stream.ReadLong()   'In the wild, these appear to always (?) be 0
            .mbm_HeightTwips = m_Stream.ReadLong()
            If (.mbm_Width <= 0) Or (.mbm_Height <= 0) Then GoTo NextFrame
            
            'Retrieve color space data
            .mbm_BPP = m_Stream.ReadLong()
            If (.mbm_BPP <= 0) Then GoTo NextFrame
            .mbm_IsColor = (m_Stream.ReadLong <> 0)
            .mbm_PaletteSize = m_Stream.ReadLong()
            
            'Retrieve compression, with a sanity check for correct values
            .mbm_CompressionType = m_Stream.ReadLong()
            
            'Compression values >= 5 (e.g. undocumented compression values) show up in symbian's
            ' test suite with some regularity.  This compression model isn't described anywhere,
            ' and using the same rules for 8/16/24-bit compression, just expanded to 32-bpp,
            ' don't produce useable results.  It's possible *someone* online knows how to handle
            ' these images, but I don't.  (Further complicating matters is that there is some
            ' repetition in the data, but it tends to occur every *6* bytes which is just weird.
            ' Also, there is no obvious marker at the start of the stream for zip or gz compression,
            ' and the size of the data is so large that those models seem unlikely anyway... idk.)
            '
            'For now, I just throw a 32-bit RLE decompressor at them and let them run until they
            ' hit an error (typically a bad RLE marker that extends beyond the end of the file).
            If (.mbm_CompressionType > mbmc_RLE_32bit) Then
                InternalError "LoadMBM_FromFile", "bad compression type: " & .mbm_CompressionType
                .mbm_CompressionType = .mbm_CompressionType And &H3&
            End If
            
            'Hypothetically we should be pointing at pixel data now.  (This failsafe is helpful for malformed
            ' files - which are prevalent "in the wild" - but it is only relevant for non-AIF files, because AIF
            ' files contain an additional mask without an explicit offset.)
            If (Not inMaskMode) Then
                If (m_Stream.GetPosition <> .mbm_Offset + .mbm_HeaderLength) Then m_Stream.SetPosition .mbm_Offset + .mbm_HeaderLength
            End If
            
            'Prep our target DIB and make it opaque (most images in this format will be < 32-bpp)
            If inMaskMode Then
                If (maskDIB Is Nothing) Then Set maskDIB = New pdDIB
                maskDIB.CreateBlank .mbm_Width, .mbm_Height, 32, vbWhite, 255
                Set targetDIB = maskDIB
            Else
                Set .mbm_DIB = New pdDIB
                .mbm_DIB.CreateBlank .mbm_Width, .mbm_Height, 32, vbWhite, 255
                Set targetDIB = .mbm_DIB
            End If
            
            'Build palette?
            Dim srcPalette() As RGBQuad
            ReDim srcPalette(0 To 255) As RGBQuad
            
            'Palettes appear to be hard-coded depending on color depth?
            Dim r As Long, g As Long
            If .mbm_IsColor Then
                
                'Fixed palettes have been reverse-engineered for 4- and 8-bpp; bit-depths below two
                ' can exist but aren't well-defined; for these, we currently default to the 4-bpp palette.
                If (.mbm_BPP <= 4) Then
                    BuildPalette_4 srcPalette
                ElseIf (.mbm_BPP = 8) Then
                    BuildPalette_8 srcPalette
                End If
            
            'Grayscale palettes follow standard rules
            Else
                
                'Failsafe check for too-large depth values (symbian's test suite has images that
                ' are listed as grayscale and 16-bpp??)
                If (.mbm_BPP <= 8) Then
                    
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
                    
                'For bad color-depths, force to non-grayscale mode
                Else
                    .mbm_IsColor = True
                End If
                
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
                scanlineSize = pxWidth * 2  '16-bit alignment; 4-bits go unused
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
            
            'Based on a sampling of images from the Symbian test suite on GitHub, 4-byte alignment
            ' on 32-bpp data produces good images 99+% of the time (and the remaining 1% are potentially
            ' test-specific images for fuzzing), so I'm inclined to commit to it here, despite informal
            ' specs being unclear on the matter.
            ElseIf (.mbm_BPP > 24) Then
                scanlineSize = (scanlineSize + 3) And &HFFFFFFFC
            End If
            
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
            
            'RLE compression requires a second buffer for decompression; individual scanlines do
            ' not have their own markers (like e.g. PSD's PackBits), so it's significantly easier
            ' to just decode the full image in advance, since padding bytes will *also* be encoded
            ' inside the RLE data, and we need to account for those to get proper pixel offsets.
            ' (XNView does not do this on 24-bpp MBM files, for example, and the resulting images
            ' become obviously skewed.)
            Dim numScanlines As Long
            numScanlines = .mbm_Height - 1
            
            Dim srcByte As Byte, srcLong As Long, numPixelsProcessed As Long
            Dim numSafeBytes As Long, curOffset As Long
            
            'Perform RLE decompression first
            Dim pxPostRLE() As Byte
            If (.mbm_CompressionType <> mbmc_None) Then
                
                'Add a small margin of safety to the RLE decompression buffer; old file formats like
                ' this have a non-zero risk of being malformed, and while I do perform safety checks
                ' in the inner loop, this provides an added safeguard against surprises.
                numSafeBytes = scanlineSize * .mbm_Height
                ReDim pxPostRLE(0 To numSafeBytes + scanlineSize - 1) As Byte
                
                'RLE markers define how many pixel-sized values we copy and/or repeat (based on the
                ' value of the RLE marker byte).  This obviously varies by RLE compression strategy.
                ' One interesting thing to note is that I don't know if the RLE compression used must
                ' match the color-depth of the image - e.g. could you use 16-bit compression on 8-bit
                ' data, for potentially better results?  I haven't seen any such mismatches in my
                ' collection of 1000+ test images from Symbian, but if that possibility exists, I may
                ' need to rework the code below to account for it.
                Dim copyPxSize As Long, pxTemp(0 To 3) As Byte
                If (.mbm_CompressionType = mbmc_RLE_8bit) Then
                    copyPxSize = 1
                ElseIf (.mbm_CompressionType = mbmc_RLE_16bit) Then
                    copyPxSize = 2
                ElseIf (.mbm_CompressionType = mbmc_RLE_24bit) Then
                    copyPxSize = 3
                ElseIf (.mbm_CompressionType = mbmc_RLE_32bit) Then
                    copyPxSize = 4
                End If
                
                'Ignore scanlines and just treat the data as an arbitrary source of bytes.
                ' Note that this function is *not* guaranteed to be robust against malformed data;
                ' e.g. it may crash if fed deliberately corrupted RLE data that produces runs
                ' extending beyond the end of the image.  (I've tried to account for this as best I
                ' can, but I don't have an easy way to aggressively fuzz the decompressor!)
                curOffset = 0
                Do While curOffset < numSafeBytes
                    
                    '12-bit RLE uses special rules (only a nibble is used for RLE flags)
                    If (.mbm_CompressionType = mbmc_RLE_12bit) Then
                        
                        'Retrieve both bytes
                        r = m_Stream.ReadByte()
                        g = m_Stream.ReadByte()
                        
                        'Mask off the RLE bits
                        srcLong = (g And &HF0&) \ 16
                        
                        'If the value is larger than 0, repeat this integer (n - 1) times
                        For j = 0 To srcLong
                            pxPostRLE(curOffset + j * 2) = r
                            pxPostRLE(curOffset + j * 2 + 1) = g
                        Next j
                        
                        curOffset = curOffset + (srcLong + 1) * 2
                        
                    'Other compression models behave similarly; the only difference is how
                    ' many bytes we repeat for each RLE indicator (1, 2, or 3 for 8, 16, 24-bpp respectively)
                    Else
                        
                        'Get the RLE byte
                        srcLong = m_Stream.ReadByte()
                        
                        'Use MSB to determine RLE meaning
                        If (srcLong < &H80&) Then
                            
                            'This is a run.  Fill the byte value manually (n+1 times)
                            srcLong = srcLong + 1
                            
                            'Overflow safeguard
                            If (curOffset + srcLong * copyPxSize > numSafeBytes) Then
                                InternalError "LoadMBM_FromFile", "bad RLE run: " & CStr(curOffset + srcLong * copyPxSize) & " vs " & CStr(numSafeBytes) & ", " & CStr(trailerOffset - m_Stream.GetPosition()) & " bytes remaining in stream"
                                srcLong = (numSafeBytes - curOffset) \ copyPxSize
                            End If
                            
                            '8-bit can be handled more quickly, c/o FillMemory
                            If (.mbm_CompressionType = mbmc_RLE_8bit) Then
                                
                                srcByte = m_Stream.ReadByte()
                                
                                VBHacks.FillMemory VarPtr(pxPostRLE(curOffset)), srcLong, srcByte
                                curOffset = curOffset + srcLong
                            
                            '16/24/32 bpp require different behavior, owing to the custom number of bytes
                            ' we must retrieve and then copy into place
                            Else
                                
                                'Retrieve the pixel bytes that must be copied
                                For k = 0 To copyPxSize - 1
                                    pxTemp(k) = m_Stream.ReadByte()
                                Next k
                                
                                'Perform the copy
                                For j = 0 To srcLong - 1
                                    For k = 0 To copyPxSize - 1
                                        pxPostRLE(curOffset + j * copyPxSize + k) = pxTemp(k)
                                    Next k
                                Next j
                                
                                curOffset = curOffset + srcLong * copyPxSize
                                
                            End If
                            
                        Else
                        
                            'This is a segment (size 0x100 - value) of uncompressed bytes.  Read the bytes
                            ' directly into the target buffer.
                            srcLong = 256 - srcLong
                            
                            'All pixel widths can be handled identically, since it's just a memcpy size
                            ' that needs to be modified.
                            If (curOffset + srcLong * copyPxSize > numSafeBytes) Then
                                InternalError "LoadMBM_FromFile", "bad RLE chunk: " & CStr(curOffset + srcLong * copyPxSize) & " vs " & CStr(numSafeBytes) & ", " & CStr(trailerOffset - m_Stream.GetPosition()) & " bytes remaining in stream"
                                srcLong = (numSafeBytes - curOffset) \ copyPxSize
                            End If
                            
                            m_Stream.ReadBytesToBarePointer VarPtr(pxPostRLE(curOffset)), srcLong * copyPxSize
                            curOffset = curOffset + srcLong * copyPxSize
                            
                        End If
                        
                    End If
                    
                Loop
                
            End If
            
            'Process each scanline in turn
            Dim x As Long, y As Long, alphaFound As Boolean
            alphaFound = False
            
            Dim tmpSA1D As SafeArray1D, dstPixels() As RGBQuad
            Dim tmpSA1DSrc As SafeArray1D, srcPixels() As RGBQuad
            
            For y = 0 To numScanlines
                
                'Where we copy the bytes from varies by compression type (RLE uses a separate buffer)
                If (.mbm_CompressionType = mbmc_None) Then
                    m_Stream.ReadBytesToBarePointer VarPtr(pxScanline(0)), scanlineSize
                Else
                    CopyMemoryStrict VarPtr(pxScanline(0)), VarPtr(pxPostRLE(0)) + y * scanlineSize, scanlineSize
                End If
                
                'With the line decompressed, we can now convert it to RGB/A
                
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
                            
                            'Weird alignment requirements mean that scanlines can extend quite far
                            ' beyond pixel boundaries at lower bit-depths - so we must always check
                            ' to ensure we're still inside pixel boundaries, even though we're safely
                            ' inside the bounds of the upsampled scanline in the file.
                            If (numPixelsProcessed <= xFinal) Then
                                srcByte = pxScanline(x)
                                preConvert(numPixelsProcessed) = srcByte And &HF
                                numPixelsProcessed = numPixelsProcessed + 1
                            End If
                            
                            If (numPixelsProcessed <= xFinal) Then
                                preConvert(numPixelsProcessed) = (srcByte \ 16) And &HF
                                numPixelsProcessed = numPixelsProcessed + 1
                            End If
                            
                        Next x
                    
                    End If
                
                '/end pre-processing of < 8-bpp images
                End If
                
                'Point a destination array at the target DIB
                targetDIB.WrapRGBQuadArrayAroundScanline dstPixels, tmpSA1D, y
                
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
            targetDIB.UnwrapRGBQuadArrayFromDIB dstPixels
            
            'Premultiply our finished alpha channel
            targetDIB.SetAlphaPremultiplication True
            
            'If this is plain bitmap data, mark this frame as successful!
            If (Not inMaskMode) Then
                .mbm_FrameOK = True
                
            'If we have just parsed a mask, however, we now need to merge the mask's data with the
            ' associated bitmap frame.  Masks make no guarantees about bit-depth or color-model
            ' (e.g. they can be 24-bpp RGB format!), and they also are not guaranteed to be the
            ' same size as their associated frame, so all we can do is try and copy over relevant
            ' bytes as best we can.
            Else
            
                'Start by finding the smallest width/height value between the mask and associated image.
                Dim minWidth As Long, minHeight As Long
                minWidth = PDMath.Min2Int(maskDIB.GetDIBWidth, .mbm_DIB.GetDIBWidth)
                minHeight = PDMath.Min2Int(maskDIB.GetDIBHeight, .mbm_DIB.GetDIBHeight)
                If (minWidth <= 0) Or (minHeight <= 0) Then
                    InternalError "LoadMBM_FromFile", "mask or frame has invalid dimensions; abandoning merge"
                    inMaskMode = False
                    GoTo NextFrame
                End If
                
                'Unpremultiply the target DIB
                .mbm_DIB.SetAlphaPremultiplication False
                
                'Iterate through mask pixels and replace destination image pixels with their value
                ' (where black equals opaque, white equals transparent, and grayscale values
                ' in-between are treated as variable opacity).
                Dim newOpacity As Long
                For y = 0 To minHeight - 1
                    maskDIB.WrapRGBQuadArrayAroundScanline srcPixels, tmpSA1DSrc, y
                    .mbm_DIB.WrapRGBQuadArrayAroundScanline dstPixels, tmpSA1D, y
                For x = 0 To minWidth - 1
                
                    newOpacity = Colors.GetLuminance(srcPixels(x).Red, srcPixels(x).Green, srcPixels(x).Blue)
                    dstPixels(x).Alpha = (255 - newOpacity)
                
                Next x
                Next y
                
                'Free both wrappers before exiting, obviously
                maskDIB.UnwrapRGBQuadArrayFromDIB srcPixels
                .mbm_DIB.UnwrapRGBQuadArrayFromDIB dstPixels
                
                'Re-apply premultiplication
                .mbm_DIB.SetAlphaPremultiplication True
                
            End If
            
            'If this is an AIF file, each frame will be followed by a transparency mask.
            ' We want to parse that mask now.
            If m_IsAIF Then
                
                'Every-other-frame toggles between image and mask
                inMaskMode = Not inMaskMode
                
                'If we are about to parse a mask, forcibly set the frame offset to the end-of-frame.
                ' (Masks do not specify an offset, oddly, so the only way to orient ourselves is to
                '  forcibly add frame length to original offset for non-mask-frames.)
                If inMaskMode Then
                    m_Stream.SetPosition (.mbm_Offset + .mbm_Length), FILE_BEGIN
                    GoTo StartOfFrame
                End If
                
            End If
            
        End With
    
NextFrame:
    Next i
    
    'If a temporary mask DIB exists, free it
    Set targetDIB = Nothing
    Set maskDIB = Nothing
    
    'With all frames parsed, we now need to construct a new layer for each frame in the
    ' destination pdImage object.
    
    'Start by finding the largest frame in the file; we'll use this for the parent image dimensions.
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
        dstImage.SetOriginalFileFormat PDIF_MBM
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

Private Sub BuildPalette_4(ByRef srcPalette() As RGBQuad)
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
End Sub

Private Sub BuildPalette_8(ByRef srcPalette() As RGBQuad)
    
    Dim palIndex As Long
    Dim r As Long, g As Long, b As Long
                    
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
    
End Sub

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
