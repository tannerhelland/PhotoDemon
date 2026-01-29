VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdHDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon HDR/RGBE (Radiance) Decoder
'Copyright 2026-2026 by Tanner Helland
'Created: 27/January/26
'Last updated: 29/January/26
'Last update: wrap up initial build
'
'The HDR file format is one of the oldest HDR-compatible image file formats.  It combines a human-readable
' header and simple uncompressed or RLE-encoded data streams with a novel 32-bpp floating-point storage
' mechanism.  Both RGBE and XYZE color formats are allowed, but due to a lack of XYZE test data, PhotoDemon
' only supports RGBE (for now).
'
'PhotoDemon previously used the 3rd-party FreeImage library to parse HDR files, but in 2026, I wrote this
' novel VB6 implementation from scratch.  FreeImage's implementation makes erroneous assumptions about
' color data and header limitations (for example, it limits header lines to 256-characters which causes it
' to fail on valid "real world" stacked HDR images - see https://github.com/tannerhelland/PhotoDemon/issues/739)
' which necessitated this changeover.
'
'The "official" Radiance spec is the definitive source of information on the HDR format.  I used this PDF,
' as linked by Wikipedia (https://en.wikipedia.org/wiki/RGBE_image_format):
' https://radsite.lbl.gov/radiance/refer/filefmts.pdf
'
'I have not encountered any deviations from the spec in "real-world" test images, so I think that PDF is
' a comprehensive resource for other developers looking to support HDR data.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps additional diagnostic information
' to PD's primary debug log.
Private Const HDR_DEBUG_VERBOSE As Boolean = False

'Image width/height, in pixels.
Private m_Width As Long, m_Height As Long

'Byte-by-byte access is provided, as always, by a pdStream instance.
Private m_Stream As pdStream

'List of human-readable strings in the header, stored in sequential order.
Private m_HeaderLines As pdStringStack

'Relevant header data only
Private Type PD_RadianceHeader
    rh_Width As Long
    rh_Height As Long
    rh_Orientation As GP_RotateFlip
    rh_isXYZE As Boolean
    rh_Exposure As Double
    rh_Primaries(0 To 7) As Single
End Type

Private m_Header As PD_RadianceHeader

'HDR files always start with a human-readable header, separated line-by-line.
'
'In this function, all we want to do is read and verify the *first* line of the file (if one exists)
' and look for the proper magic word.
Friend Function IsFileHDR(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True) As Boolean
    
    Const FUNC_NAME As String = "IsFileHDR"
    
    'Ensure target file exists
    Dim potentialMatch As Boolean
    potentialMatch = Files.FileExists(srcFilename)
    If potentialMatch And requireValidFileExtension Then
        
        'Very, very old Radiance files might use the ".pic" file extension.  Modern files exclusively
        ' use ".hdr".  The "rgbe" and "xyze" extensions are only used to mirror Photoshop's behavior
        ' for this format (I have not found any "in the wild" images that use these extensions).
        potentialMatch = Strings.StringsEqualAny(Files.FileGetExtension(srcFilename), True, "hdr", "rgbe", "xyze", "pic")
        
    End If
    
    If potentialMatch Then
        
        potentialMatch = False
        
        'Open a stream on the target file.
        Set m_Stream = New pdStream
        If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename, optimizeAccess:=OptimizeSequentialAccess) Then
            
            'Pull the first line from the file.  (Radiance files only use LF not CRLF line-endings.)
            Const ASCII_LF As Long = 10
            Dim posLineEnding As Long
            posLineEnding = m_Stream.FindByte(ASCII_LF, 0, False, 11)
            If (posLineEnding > 0) Then
                
                Dim sFirstLine As String
                sFirstLine = m_Stream.ReadString_ASCII(posLineEnding)
                
                Const HDR_MAGIC_HEADER As String = "#?RADIANCE"
                If Strings.StringsEqual(sFirstLine, HDR_MAGIC_HEADER, True) Then
                    If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "Radiance header found; HDR format likely."
                    potentialMatch = True
                
                'Some 3rd-party HDR parsers imply that "#?RGBE" is also a valid header; see for example
                ' https://github.com/plepers/hdr2png/blob/master/hdrloader.cpp
                ElseIf Strings.StringsEqual(Left$(sFirstLine, 6), "#?RGBE", True) Then
                    If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "HDRE header found; HDR format likely."
                    potentialMatch = True
                    
                Else
                    potentialMatch = False
                End If
                
                'TODO: do any alternate headers exist?  Studying this is TBD.
                
            End If
            
            m_Stream.StopStream True
            
        End If
        
    End If
    
    IsFileHDR = potentialMatch
    
End Function

'HDR files always start with a human-readable header, separated line-by-line.
'
'In this function, all we want to do is read and verify the *first* line of the file (if one exists)
' and look for the proper magic word.
Friend Function LoadHDR_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "LoadHDR_FromFile"
    LoadHDR_FromFile = False
    
    'Validate the file
    If Me.IsFileHDR(srcFile, False) Then
        
        'This appears to be a radiance-compatible HDR file.
        
        'Next, we want to parse the file header.
        ' (The file header is a series of human-readable lines.  The header-reader simply grabs all these lines,
        '  stores them in a string stack, and extracts useable information into a simple header object.)
        If (Not ParseHeader(srcFile)) Then
            InternalError FUNC_NAME, "header parse failed"
            Exit Function
        End If
        
        'Still here?  The header must have validated OK and the stream now points at pixel data.
        ' Go ahead and parse it!
        
        'Prep the destination buffer and set it to white opaque.  (HDR files do not support alpha channels.)
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank m_Header.rh_Width, m_Header.rh_Height, 32, vbWhite, 255
        
        'Decode pixel data directly into the destination DIB
        LoadHDR_FromFile = ParsePixels(dstDIB)
        If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "LoadHDR_FromFile: " & LoadHDR_FromFile
        
        'The returned data is always premultiplied
        If LoadHDR_FromFile Then dstDIB.SetInitialAlphaPremultiplicationState True
    
    End If
    
End Function

'Read the header from the target HDR file.  The existing class-level stream (if any) will be closed and re-opened
' on the target file.
Private Function ParseHeader(ByRef srcFile As String) As Boolean
    
    Const FUNC_NAME As String = "ParseHeader"
    ParseHeader = False
    
    'Individual header lines will be stored in this stack object
    Set m_HeaderLines = New pdStringStack
    
    'Open a stream on the target file.
    Set m_Stream = New pdStream
    If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
        
        'Next, we want to iterate header lines *until* we encounter an empty line.
        ' (This marks the end of the header.)
        ' There might be one line in the header or 100.  No guarantees!
        
        'Max line width is not defined in the spec, but we need some arbitrary upper value as a safety guarantee.
        ' (FreeImage uses 256 chars here, but that fails on some real-world files; we'll use a much larger value
        '  which provides an ~order-of-magnitude safety net.)
        Const MAX_LINE_IN_CHARS As Long = 2048
        
        'Start pulling lines from the file.  (Radiance files only use LF not CR/CRLF line-endings.)
        Const ASCII_LF As Long = 10
        Dim posLineStart As Long
        posLineStart = 0
        
        Dim posLineEnd As Long
        posLineEnd = m_Stream.FindByte(ASCII_LF, 0, False, MAX_LINE_IN_CHARS)
        
        Do While (posLineEnd > posLineStart + 1)
            
            'Store this line to the stack
            Dim tmpString As String
            tmpString = m_Stream.ReadString_ASCII((posLineEnd - posLineStart))
            m_HeaderLines.AddString tmpString
            
            'Advance the pointer by 1 (to skip the line-end character)
            m_Stream.SetPosition posLineEnd + 1, FILE_BEGIN
            posLineStart = m_Stream.GetPosition()
            
            'Look for the next lf char
            posLineEnd = m_Stream.FindByte(ASCII_LF, posLineStart, False, MAX_LINE_IN_CHARS)
            
        Loop
        
        'Ensure we ended on an empty line
        ParseHeader = (posLineEnd = posLineStart + 1)
        
        'DEBUG ONLY: print out all found lines
        If HDR_DEBUG_VERBOSE Then
            PDDebug.LogAction "Radiance header follows: "
            Dim i As Long
            If (m_HeaderLines.GetNumOfStrings > 0) Then
                For i = 0 To m_HeaderLines.GetNumOfStrings - 1
                    PDDebug.LogAction CStr(i + 1) & ": " & m_HeaderLines.GetString(i)
                Next i
            End If
        End If
        
        'TODO: parse header strings for usable information (e.g. exposure)
        
        'Immediately following the empty line marking the end of the header is a (so-called) RESOLUTION string.
        ' This defines image dimensions and orientation.
        m_Stream.SetPosition 1, FILE_CURRENT
        posLineEnd = m_Stream.FindByte(ASCII_LF, -1, False, MAX_LINE_IN_CHARS)
        
        Dim sResolution As String
        If (posLineEnd > 0) Then
            sResolution = m_Stream.ReadString_ASCII(posLineEnd - m_Stream.GetPosition())
            If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "Resolution: " & sResolution
        
            'Advance the stream pointer by 1 (to skip the LF char at the end of the resolution).
            m_Stream.SetPosition 1, FILE_CURRENT
            ParseHeader = True
        Else
            ParseHeader = False
        End If
        
        'The stream now points at the start of the pixel data.
        
        'If parsing looks OK so far, we next want to pull out any useable data from the recovered strings
        If ParseHeader Then
        
            'Right now, my primary concern is the resolution string.
            ' (Exposure, white point, etc can be solved later.)
            
            'The resolution string is always in a format similar to "-Y N +X M", but the x and y can appear
            ' in any order (which corresponds to different orientations).
            
            'Resolution unit should always be upper-case, but better safe than sorry
            sResolution = UCase$(sResolution)
            
            Dim sResParts() As String
            sResParts = Split(sResolution, " ")
            If (UBound(sResParts) < 3) Then
                PDDebug.LogAction "Bad resolution string: " & UBound(sResParts)
                ParseHeader = False
                Exit Function
            End If
            
            'Solve according to the order of x/y
            If (sResParts(0) = "-Y") Then
                m_Header.rh_Width = CLng(sResParts(3))
                m_Header.rh_Height = CLng(sResParts(1))
                If (sResParts(2) = "+X") Then
                    m_Header.rh_Orientation = GP_RF_NoneFlipNone
                Else
                    m_Header.rh_Orientation = GP_RF_NoneFlipX
                End If
            ElseIf (sResParts(0) = "+Y") Then
                m_Header.rh_Width = CLng(sResParts(3))
                m_Header.rh_Height = CLng(sResParts(1))
                If (sResParts(2) = "+X") Then
                    m_Header.rh_Orientation = GP_RF_NoneFlipY
                Else
                    m_Header.rh_Orientation = GP_RF_NoneFlipXY
                End If
            ElseIf (sResParts(0) = "+X") Then
                m_Header.rh_Width = CLng(sResParts(3))
                m_Header.rh_Height = CLng(sResParts(1))
                If (sResParts(2) = "+Y") Then
                    m_Header.rh_Orientation = GP_RF_270FlipNone
                Else
                    m_Header.rh_Orientation = GP_RF_90FlipY
                End If
            ElseIf (sResParts(0) = "-X") Then
                m_Header.rh_Width = CLng(sResParts(3))
                m_Header.rh_Height = CLng(sResParts(1))
                If (sResParts(2) = "+Y") Then
                    m_Header.rh_Orientation = GP_RF_270FlipY
                Else
                    m_Header.rh_Orientation = GP_RF_90FlipNone
                End If
            End If
            
            'Ensure width/height are >= 0
            If (m_Header.rh_Width <= 0) Or (m_Header.rh_Height <= 0) Then
                ParseHeader = False
            Else
            
                'HDR files can be quite large; make sure we have a RAM safety net before loading.
                ' (500 MB, decompressed, is a good safety net when considering the amount of extra
                ' RAM needed for various internal operations and/or the rest of the decode process.)
                If ((CSng(m_Header.rh_Width) * CSng(m_Header.rh_Height) * 4!) < 5E+08!) Then
                    ParseHeader = True
                Else
                    InternalError FUNC_NAME, "source image is huge; rejecting due to memory pressure"
                    ParseHeader = False
                End If
                
            End If
            
            If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "ParseHeader: " & ParseHeader
            
        End If
        
    End If

End Function

'Assuming the header validated OK, go ahead and use its data to parse the pixel stream.
' 1) DO NOT call this function IF the header didn't validate (it'll crash)
' 2) YOU MUST call this function IMMEDIATELY following the header parser; only then is stream pointer
'    alignment guaranteed correct against the start of the pixel data.
Private Function ParsePixels(ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "ParsePixels"
    ParsePixels = False
    
    'HDR files use a simple RLE compression scheme.  Two RLE varieties are used (called "old" and "new"),
    ' and uncompressed data is also allowed.
    '
    'I have not validated this, but hypothetically individual scanlines can independently use different
    ' encoding schemes.  (Flags at the beginning of each line determine what scheme is used on that line.)
    Dim dstWidth As Long, dstHeight As Long
    dstWidth = m_Header.rh_Width
    dstHeight = m_Header.rh_Height
    If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "Decoding image with dimensions: " & dstWidth & "x" & dstHeight
    
    'Pull the entire remaining file bytes into memory.
    ' (This is significantly faster than walking the file byte-by-byte with stream I/O calls.)
    Dim srcBytes() As Byte, srcByteLen As Long
    If (m_Stream.ReadBytes(srcBytes) = 0) Then
        InternalError FUNC_NAME, "Rest of stream read failed??"
        Exit Function
    End If
    
    srcByteLen = UBound(srcBytes) + 1
    m_Stream.StopStream True
    
    'DO NOT TOUCH THE STREAM OBJECT PAST THIS POINT; IT IS CLOSED.
    
    'Index into the source (compressed) array.  Because the source data is (potentially) RLE-compressed,
    ' this does not map cleanly to positions in the destination array.  Both must be tracked independently.
    Dim idxSrc As Long: idxSrc = 0
    Dim nextQuad As RGBQuad
    
    'Width of a scanline in the *destination* data
    Dim scanlineWidth As Long
    scanlineWidth = dstWidth * 4
    
    'Holds the uncompressed source bytes, still in RGBE format (but with RLE removed).
    Dim srcBytesUncompressed() As Byte
    ReDim srcBytesUncompressed(0 To scanlineWidth - 1) As Byte
    
    Dim dstPixels() As Byte, dstSA As SafeArray1D
    
    'Start iterating image scanlines.
    Dim x As Long, y As Long
    For y = 0 To dstHeight - 1
        
        'The second scanline width check here is because the first bit of a 2-byte flag is used to mark
        ' RLE compression.  Images larger than this (which don't really affect PD as a 32-bit app) cannot
        ' use RLE compression because they will overwrite this flag.
        If ((scanlineWidth < 8) Or (scanlineWidth > &H7FFF&)) Then
            
            'Use the "no-compression" loader for this line.
            If (Not ReadScanline_NoRLE(srcBytes, idxSrc, srcByteLen, scanlineWidth, srcBytesUncompressed)) Then
                Exit Function
            End If
        
        'Scanline size is RLE-compatible
        Else
        
            'Peek the first quad in this scanline.
            CopyMemoryStrict VarPtr(nextQuad), VarPtr(srcBytes(idxSrc)), 4
            
            'Per the spec, RLE compression is marked by the first two bytes each being 2, the high bit
            ' of the next two bytes being 0, and the remaining 15-bits equaling image width.
            If (nextQuad.Blue <> 2) Or (nextQuad.Green <> 2) Or ((nextQuad.Red And &H80&) <> 0) Then
                
                'This line doesn't use modern RLE.  Look for traditional RLE instead, and failing that,
                ' consider the line uncompressed.
                
                If HDR_DEBUG_VERBOSE Then
                    PDDebug.LogAction "WARNING: this image uses old-style RLE.  It cannot be decompressed yet."
                End If
                
                'TODO: implement old RLE!
                If (Not ReadScanline_NoRLE(srcBytes, idxSrc, srcByteLen, scanlineWidth, srcBytesUncompressed)) Then
                    Exit Function
                End If
            
            'This line is RLE-compressed using the "better" (planar) RLE scheme.
            Else
                
                'Check the last-two bytes, which must equal the image's width (so image width must be <= 32768).
                Dim slCheck As Long
                slCheck = CLng(nextQuad.Red) * 256 + CLng(nextQuad.Alpha)
                
                'PD checks both width
                If (slCheck <> dstWidth) Then
                    
                    'If the image is rotated, it *is* OK to match image height here
                    If (slCheck <> dstHeight) And (m_Header.rh_Orientation = GP_RF_NoneFlipNone) Then
                        InternalError FUNC_NAME, "bad scanline width in RLE packet: " & slCheck
                        Exit Function
                    End If
                    
                End If
                
                'All RLE flags check out!  Advance the source pointer past these flags to the start
                ' of the actual pixel data.
                idxSrc = idxSrc + 4
                
                'In this scheme, channels are separated into planar (not interleaved) format,
                ' and RLE-compressed in order, starting with the three primaries and then the exponent.
                ' This obviously produces better compression than interleaved data.
                
                'Prep a temporary buffer for the planar-format bytes.
                Dim srcBytesPlanar() As Byte, srcBytesPlanarSize As Long
                If (srcBytesPlanarSize = 0) Then
                    srcBytesPlanarSize = dstWidth * 4
                    ReDim srcBytesPlanar(0 To srcBytesPlanarSize - 1) As Byte
                End If
                
                'We can simply RLE-decompress everything into a planar buffer, then manually re-assemble that into
                ' a standard interleaved buffer.  (This is slower then decoding straight from planar, but it's also
                ' a hell of a lot less code.)
                Dim idxDst As Long, nextByte As Byte
                idxDst = 0
                
                'Iterate bytes until we reach end-of-scanline
                Do
                    
                    'Read the next byte from the source data
                    nextByte = srcBytes(idxSrc)
                    idxSrc = idxSrc + 1
                    
                    'High-bit = 1 means REPEAT the next byte [n] times
                    If (nextByte > 128) Then
                        
                        Dim runSize As Long
                        runSize = nextByte - 128
                        
                        nextByte = srcBytes(idxSrc)
                        idxSrc = idxSrc + 1
                        
                        For x = 0 To runSize - 1
                            If (idxDst < srcBytesPlanarSize) Then
                                srcBytesPlanar(idxDst) = nextByte
                                idxDst = idxDst + 1
                            Else
                                Debug.Print "bad idxDst vs srcBytesPlanarSize: " & idxDst & " vs " & srcBytesPlanarSize
                            End If
                        Next x
                        
                    'High-bit = 0 means COPY [n] bytes directly from source
                    Else
                        
                        'Copy must have > 0 bytes specified, and not exceed scanline size
                        If (nextByte > 0) And (idxDst + nextByte <= srcBytesPlanarSize) Then
                            CopyMemoryStrict VarPtr(srcBytesPlanar(idxDst)), VarPtr(srcBytes(idxSrc)), nextByte
                            idxDst = idxDst + nextByte
                            idxSrc = idxSrc + nextByte
                        
                        '0-byte copies are pointless but this value *does* appear in images "in the wild",
                        ' so continue parsing despite it.
                        Else
                            If (nextByte <> 0) Then
                                InternalError FUNC_NAME, "bad non-run: " & nextByte & ", " & idxDst & " vs " & srcBytesPlanarSize
                                Exit Function
                            End If
                        End If
                    End If
                
                'As long as the destination buffer has room, keep reading bytes from the source
                Loop While idxDst < srcBytesPlanarSize
                
                'We now have all bytes decompressed into planar format.  Swizzle them into interleaved format.
                For x = 0 To dstWidth - 1
                    srcBytesUncompressed(x * 4) = srcBytesPlanar(x)
                    srcBytesUncompressed(x * 4 + 1) = srcBytesPlanar(dstWidth + x)
                    srcBytesUncompressed(x * 4 + 2) = srcBytesPlanar(dstWidth * 2 + x)
                    srcBytesUncompressed(x * 4 + 3) = srcBytesPlanar(dstWidth * 3 + x)
                Next x
                
            End If
            
        End If
        
        'srcBytesUncompressed() now contains standard-format RGBE/XYZE data.  We must translate it into BGRA.
        
        'Build a LUT for powers of 2.  (This is much faster than exponent calculations on-the-fly,
        ' particularly in VB6 which relies on old x87 float calls.)
        Dim lutPow2(0 To 255) As Single
        For x = 0 To 255
            lutPow2(x) = (2# ^ CDbl(x - 128))
        Next x
        
        'Color values will be translated into 32-bit floats first, then downsampled to 8-bit RGBA
        ' (with color management, obviously)
        Dim dstPixelsF() As Single
        ReDim dstPixelsF(0 To dstWidth * 4 - 1) As Single
        
        Dim xOffset As Long
        Dim r As Long, g As Long, b As Long, e As Long
            
        'Translate RGBE/XYZE format to BGRA.
        ' (TODO: XYZE; sample files needed!)
        For x = 0 To dstWidth - 1
            
            xOffset = x * 4
            
            r = srcBytesUncompressed(xOffset)
            g = srcBytesUncompressed(xOffset + 1)
            b = srcBytesUncompressed(xOffset + 2)
            e = srcBytesUncompressed(xOffset + 3)
            
            'Null exponent requires null channel values
            If (e = 0) Then
                dstPixelsF(xOffset) = 0!
                dstPixelsF(xOffset + 1) = 0!
                dstPixelsF(xOffset + 2) = 0!
            
            'Transform mantissas by the shared exponent
            Else
                dstPixelsF(xOffset) = CSng(r) * lutPow2(e)
                dstPixelsF(xOffset + 1) = CSng(g) * lutPow2(e)
                dstPixelsF(xOffset + 2) = CSng(b) * lutPow2(e)
            End If
            
            'Alpha is irrelevant; we do not read or write it.
            'dstPixelsF(xOffset + 3) = 1!
            
        Next x
        
        'Wrap a pointer around the target line in the destination pixel array.
        dstDIB.WrapArrayAroundScanline dstPixels, dstSA, y
        
        'TODO: additional color-management/tone-mapping here, based on embedded color space data
        
        Dim fR As Single, fG As Single, fB As Single
        
        'HDR files are always non-gamma-corrected; to match behavior in other software,
        ' we must gamma-correct at 2.2 by default.  (Files can specify their own values but they are *not*
        ' guaranteed to work; for example, a gamma of "0.4545" or "1" both theoretically mean "use normal 1/2.2".
        Const GAMMA_CORRECTION_FACTOR As Double = 1# / 2.2
        Dim lutGamma(0 To 255) As Byte
        For x = 0 To 255
            fR = CSng(x)
            fR = (fR / 255!) ^ GAMMA_CORRECTION_FACTOR * 255!
            If (fR > 255!) Then fR = 255!
            If (fR < 0!) Then fR = 0!
            lutGamma(x) = Int(fR + 0.5!)
        Next x
        
        'Iterate pixels and convert as-we-go
        For x = 0 To dstWidth - 1
        
            xOffset = x * 4
            
            fR = dstPixelsF(xOffset)
            fG = dstPixelsF(xOffset + 1)
            fB = dstPixelsF(xOffset + 2)
            
            r = Int(fR)
            g = Int(fG)
            b = Int(fB)
            
            'Failsafe only
            If (r < 0) Then r = 0
            If (r > 255) Then r = 255
            If (g < 0) Then g = 0
            If (g > 255) Then g = 255
            If (b < 0) Then b = 0
            If (b > 255) Then b = 255
            
            'Apply gamma-correction now.  (It would be higher-quality to do this while the pixel data is
            ' still in floats, but the perf impact of that is *huge* due to powf performance.)
            dstPixels(xOffset) = lutGamma(b)
            dstPixels(xOffset + 1) = lutGamma(g)
            dstPixels(xOffset + 2) = lutGamma(r)
            
            'Alpha does not need to be forced to 255 here because we initialized the target surface to opaque white
            'dstPixels(xOffset + 3) = 255
            
        Next x
        
        'Unwrap the line pointer before continuing
        dstDIB.UnwrapArrayFromDIB dstPixels
        
    Next y
    
    ParsePixels = True
    
End Function

'Copy a line of bytes from the source buffer into the temporary uncompressed buffer, without considering RLE.
' idxSrc will be incremented by scanlineWidth.
' This function will crash if the source or destination arrays are not initialized to correct sizes.
Private Function ReadScanline_NoRLE(ByRef srcBytes() As Byte, ByRef idxSrc As Long, ByVal srcByteLen As Long, ByVal scanlineWidth As Long, ByRef dstBytes() As Byte) As Boolean
    
    Const FUNC_NAME As String = "ReadScanline_NoRLE"
    
    If (idxSrc + scanlineWidth <= srcByteLen) Then
        CopyMemoryStrict VarPtr(dstBytes(0)), VarPtr(srcBytes(idxSrc)), scanlineWidth
        idxSrc = idxSrc + scanlineWidth
        ReadScanline_NoRLE = True
    Else
        InternalError FUNC_NAME, "bad index+scanline"
        ReadScanline_NoRLE = False
    End If
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdHDR." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdHDR." & funcName & "() reported an error: " & errDescription
    End If
End Sub
