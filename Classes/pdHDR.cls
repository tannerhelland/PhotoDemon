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
'Last updated: 01/February/26
'Last update: finish implementing export compression; I now consider this class feature-complete
'
'The HDR file format is one of the oldest HDR-compatible image file formats.  It combines a human-readable
' header and simple RLE-encoded data streams with a novel 32-bpp floating-point storage mechanism.
' Both RGBE and XYZE color formats are allowed but the format is otherwise quite strict (only 3-channel
' images are allowed, and only at the single 32-bpp floating-point depth).  This makes support very
' straightforward compared to other "old" image formats.
'
'PhotoDemon previously used the 3rd-party FreeImage library to handle HDR files, but in 2026, I wrote this
' novel VB6 implementation from scratch.  FreeImage's implementation makes erroneous assumptions about
' color data and header limitations (for example, it limits header lines to 256-characters which causes it
' to fail on valid "real world" stacked HDR images - see https://github.com/tannerhelland/PhotoDemon/issues/739)
' which necessitated this changeover.  FreeImage also produced buggy HDR files for many years, which prevented
' them from being read in e.g. GIMP, due to FreeImage manually writing out an "EXPOSURE=0" tag which caused the
' entire image to appear black in software that respects exposure data.
'
'The "official" Radiance spec is the definitive source of information on the HDR format.  I used this PDF,
' as linked by Wikipedia (https://en.wikipedia.org/wiki/RGBE_image_format):
' https://radsite.lbl.gov/radiance/refer/filefmts.pdf
'
'I have not encountered any deviations from the spec in "real-world" test images, so I believe that PDF is
' everything other developers need if they want to also support the HDR format.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps additional diagnostic information
' to PD's primary debug log.
Private Const HDR_DEBUG_VERBOSE As Boolean = False

'Byte-by-byte access is provided, as always, by a pdStream instance.
Private m_Stream As pdStream

'List of human-readable strings in the header, stored in sequential order.
Private m_HeaderLines As pdStringStack

'This struct collects header data that's relevant to PD.  (Other header data may exist - for example,
' "gamma" is a non-standard tag written by some editors, but because its meaning is inconsistent
' between apps, PD avoids using it.)
Private Type PD_RadianceHeader
    rh_Width As Long
    rh_Height As Long
    rh_Orientation As GP_RotateFlip
    rh_Exposure As Double
    rh_isXYZE As Boolean
    rh_PrimariesPresent As Boolean
    rh_Primaries(0 To 7) As Single
End Type

Private m_Header As PD_RadianceHeader

'HDR files always start with a human-readable header, separated line-by-line.
'
'In this function, all we want to do is read and verify the *first* line of the file (if one exists)
' and look for the proper magic word.
Friend Function IsFileHDR(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True) As Boolean
    
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
            posLineEnding = m_Stream.FindByte(ASCII_LF, 0, False, 1024)
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
                
                'If no valid matches of the "magic ID" were found, there's one last possibility.  If this is
                ' a very old .pic Radiance file, it may lack a magic number entirely - but we can still look
                ' for an unambiguous "FORMAT" key+value pair that identifies this file.
                ElseIf Strings.StringsEqual(Files.FileGetExtension(srcFilename), "pic", True) Then
                    
                    'This is a very old .pic file.  We'll allow the load process to continue, and future lines
                    ' will look for an unambiguous "FORMAT" line.
                    potentialMatch = True
                    If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "No header found, but .pic extension means we'll try to find a header anyway..."
                    
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
        LoadHDR_FromFile = ParsePixels(dstDIB, dstImage)
        If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "LoadHDR_FromFile: " & LoadHDR_FromFile
        
        'The returned data is always premultiplied (HDR does not support alpha)
        If LoadHDR_FromFile Then dstDIB.SetInitialAlphaPremultiplicationState True
        
        'If the image specifies a rotation, apply it now.
        ' (NOTE: I don't have any real-world images to test this behavior with, so it's possible I'm
        '  handling rotation backwards (e.g. rotating 270 when I should be rotating 90).  I'll revisit
        '  this as soon as I find compatible test images.  Surprisingly, no other photo editors (in my
        '  testing) handle orientation at *all*, so who knows if I should do anything here!  I'm just
        '  trying to follow the spec to a T...
        If (m_Header.rh_Orientation <> GP_RF_NoneFlipNone) Then
            If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "Image requires orientation adjustment; applying now..."
            Dim tmpDIB As pdDIB
            If GDI_Plus.GDIPlusRotateFlipDIB(dstDIB, tmpDIB, m_Header.rh_Orientation) Then
                Set dstDIB = tmpDIB
            End If
        End If
    
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
        ' There might be one text line in the header or 100.  No guarantees!
        
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
        Dim i As Long
        
        If HDR_DEBUG_VERBOSE Then
            PDDebug.LogAction "Radiance header follows: "
            If (m_HeaderLines.GetNumOfStrings > 0) Then
                For i = 0 To m_HeaderLines.GetNumOfStrings - 1
                    PDDebug.LogAction CStr(i + 1) & ": " & m_HeaderLines.GetString(i)
                Next i
            End If
        End If
        
        'Set up default header values.  Some of these can accumulate (e.g. there can be multiple
        ' exposure strings, compounding) so it's critical to start with e.g. 1.0 or 0.0 as appropriate.
        With m_Header
            .rh_Orientation = GP_RF_NoneFlipNone
            .rh_Exposure = 1#
            .rh_isXYZE = False
            .rh_PrimariesPresent = False
        End With
        
        'Some apps write gamma, but they do so inconsistently (e.g. 0.454545 and 1.0 typically mean
        ' the same correction, assuming either technical vs colloquial definitions of 2.2 gamma!).
        ' Because gamma isn't part of the formal spec and most apps ignore it if present, I also
        ' ignore it in PD.
        ResetHeaderPrimaries
        
        'If we don't find an unambiguous FORMAT line, we'll purposely fail the image load
        Dim formatOK As Boolean
        formatOK = False
        
        'Now it's time to iterate header strings and extract information that's useful to us.
        Dim hLine As String
        For i = 0 To m_HeaderLines.GetNumOfStrings - 1
            
            'We have to do a lot of string-to-numeric conversions in this function, which always represent
            ' a non-zero error potential.  At present, PD simply skips malformed lines.
            On Error GoTo BadLineParse
            
            hLine = Trim$(m_HeaderLines.GetString(i))
            
            'Ignore comments
            Const HDR_COMMENT_MARKER As String = "#"
            If (Left$(hLine, 1) <> HDR_COMMENT_MARKER) Then
                
                'Try to separate the line by "=" to obtain a key+value pair
                 Const HDR_KEYVALUE_SPLIT As String = "="
                 Dim hSplit() As String
                 hSplit = Split(hLine, HDR_KEYVALUE_SPLIT)
                 
                 Dim dValue As Double, sValue As String, keyValue As String
                 
                 'Ignore lines that don't cleanly separate into key+value pairs.
                 ' (Some HDR writers write many lines of comments without a prefixed "#";
                 ' we want to handle these gracefully.)
                 If (UBound(hSplit) = 1) Then
                    
                    keyValue = UCase$(Trim$(hSplit(0)))
                    sValue = Trim$(hSplit(1))
                    
                    Select Case keyValue
                        
                        'Per the spec, exposure is cumulative.  Specifically:
                        ' "EXPOSURE values are cumulative, so the original pixel values (i.e., radiances in
                        '  watts/steradian/m^2) must be derived by taking the values in the file and dividing
                        '  by all the EXPOSURE settings multiplied together."
                        Case "EXPOSURE"
                            dValue = Val(sValue)
                            If (dValue > 0#) Then
                                m_Header.rh_Exposure = m_Header.rh_Exposure * dValue
                            Else
                                If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "WARNING: bad exposure in file: " & sValue
                            End If
                        
                        'Per the spec there are only two valid FORMAT values.  Specifically:
                        ' "At most one FORMAT line is allowed, and it must be assigned a value of either
                        '  "32-bit_rle_rgbe" or "32-bit_rle_xyze" to be a valid Radiance picture."
                        Case "FORMAT"
                            
                            'Allow case deviations (unclear if this matters in the real-world, but it doesn't seem to hurt)
                            If Strings.StringsEqual(sValue, "32-bit_rle_rgbe", True) Then
                                m_Header.rh_isXYZE = False
                                formatOK = True
                            ElseIf Strings.StringsEqual(sValue, "32-bit_rle_xyze", True) Then
                                m_Header.rh_isXYZE = True
                                formatOK = True
                            Else
                                If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "WARNING: bad format in file: " & sValue
                                formatOK = False
                            End If
                        
                        'Per the spec:
                        ' "The CIE (x,y) chromaticity coordinates of the three (RGB) primaries and the
                        ' white point used to standardize the picture's Color System."
                        Case "PRIMARIES"
                            
                            'There should be *exactly* eight primaries (2 each for r/g/b/white point).
                            Dim sPrimaries() As String
                            sPrimaries = Split(sValue, " ")
                            If (UBound(sPrimaries) = 7) Then
                                
                                m_Header.rh_Primaries(0) = Val(sPrimaries(0))
                                m_Header.rh_Primaries(1) = Val(sPrimaries(1))
                                m_Header.rh_Primaries(2) = Val(sPrimaries(2))
                                m_Header.rh_Primaries(3) = Val(sPrimaries(3))
                                m_Header.rh_Primaries(4) = Val(sPrimaries(4))
                                m_Header.rh_Primaries(5) = Val(sPrimaries(5))
                                m_Header.rh_Primaries(6) = Val(sPrimaries(6))
                                m_Header.rh_Primaries(7) = Val(sPrimaries(7))
                                
                                'Do a quick check to ensure that the values look OK.  Some bad software
                                ' ("HDR Shop") writes all-zeroes which will crash if we try to apply it
                                ' (due to DBZ).
                                With m_Header
                                    If (.rh_Primaries(0) = 0!) And (.rh_Primaries(1) = 0!) And (.rh_Primaries(2) = 0!) And _
                                       (.rh_Primaries(3) = 0!) And (.rh_Primaries(4) = 0!) And (.rh_Primaries(5) = 0!) And _
                                       (.rh_Primaries(6) = 0!) And (.rh_Primaries(7) = 0!) Then
                                        If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "WARNING: bad primary values: " & sValue
                                        .rh_PrimariesPresent = False
                                        ResetHeaderPrimaries
                                    Else
                                        .rh_PrimariesPresent = True
                                    End If
                                End With
                                
                            End If
                            
                    End Select
                    
                 End If
                
            End If
            
BadLineParse:
        Next i
        
        'Reset error handling
        On Error GoTo 0
        
        'If we failed to find a format ID, bail now
        If (Not formatOK) Then Exit Function
        
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
        
        'If parsing looks OK so far, we next want to pull out any useable data from the resolution string
        If ParseHeader Then
            
            'The resolution string is always in a format similar to "-Y N +X M", but the x and y can appear
            ' in any order (which corresponds to different orientations).
            
            'Resolution units should always be upper-case, but better safe than sorry
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
Private Function ParsePixels(ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As Boolean
    
    Const FUNC_NAME As String = "ParsePixels"
    ParsePixels = False
    
    'HDR files use a simple RLE compression scheme.  Only a single RLE variety is used (there is supposedly
    ' an "old" RLE scheme in existence, but even ancient .pic files do not use it, so PD doesn't implement
    ' it at present; I'll revisit if I can actually find an ancient file that uses this!).  Uncompressed lines
    ' are also allowed (and are supported in all other readers, so we also support them here.)
    '
    'Importantly, each individual scanline defines its own encoding scheme.  (Specifically, a 4-byte flag at
    ' the beginning of each line determine what scheme is used on that line.  Flags can vary by line.)
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
    
    'Start iterating scanlines.
    Dim x As Long, y As Long
    For y = 0 To dstHeight - 1
        
        'The second scanline width check here is because the first bit of a 2-byte flag is used to mark
        ' RLE compression.  Images larger than this (which don't really affect PD as a 32-bit app) cannot
        ' use RLE compression because they will overwrite this flag when marking scanline width.
        If ((scanlineWidth < 8) Or (scanlineWidth > &H7FFF&)) Then
            
            'Use the "no-compression" loader for this line.
            If (Not ReadScanline_NoRLE(srcBytes, idxSrc, srcByteLen, scanlineWidth, srcBytesUncompressed)) Then
                Exit Function
            End If
        
        'Scanline size is RLE-compatible, so we need to check flags to see how to decompress this line.
        Else
        
            'Peek the first quad in this scanline.
            CopyMemoryStrict VarPtr(nextQuad), VarPtr(srcBytes(idxSrc)), 4
            
            'Per the spec, RLE compression is marked by the first two bytes each being 2, the high bit
            ' of the 3rd byte being 0, and the remaining 15-bits equaling image width.
            If (nextQuad.Blue <> 2) Or (nextQuad.Green <> 2) Or ((nextQuad.Red And &H80&) <> 0) Then
                
                'This line doesn't use modern RLE.  Look for traditional RLE instead, and failing that,
                ' consider the line uncompressed.
                
                '"Old-style" RLE requires per-pixel checks, which are cumbersome and greatly impact speed.
                ' I do not currently plan to support this scheme as no files found "in the wild" use it.
                ' Even extremely ancient government test files (from 1992! https://radsite.lbl.gov/radiance/pub/pics/index.html)
                ' use the "modern" RLE scheme, so without test images, I see no reason to support this scheme.
                'If HDR_DEBUG_VERBOSE Then
                '    PDDebug.LogAction "WARNING: this image uses old-style RLE.  It cannot be decompressed yet."
                'End If
                
                'Load this scanline as bare float data, no decompression.
                If (Not ReadScanline_NoRLE(srcBytes, idxSrc, srcByteLen, scanlineWidth, srcBytesUncompressed)) Then
                    Exit Function
                End If
            
            'This line is RLE-compressed using the "modern" (planar) RLE scheme.
            Else
                
                'Check the last-two bytes, which must equal the image's width (so image width must be <= 32768).
                Dim slCheck As Long
                slCheck = CLng(nextQuad.Red) * 256 + CLng(nextQuad.Alpha)
                
                'If the width doesn't match, check the height - images can be stored rotated but PD flags this
                ' state and swaps width/height before loading to simplify the process.
                If (slCheck <> dstWidth) Then
                    
                    'If the image is rotated, it *is* OK to match image height here
                    If (slCheck <> dstHeight) And (m_Header.rh_Orientation = GP_RF_NoneFlipNone) Then
                        InternalError FUNC_NAME, "bad scanline width in RLE packet: " & slCheck
                        Exit Function
                    End If
                    
                End If
                
                'All RLE flags check out!  Advance the source pointer past the 4-byte flag to the start
                ' of the actual pixel data.
                idxSrc = idxSrc + 4
                
                'In this compression scheme, channels are separated into planar (not interleaved) format
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
                            
                            'I used this branch to study my own decompression results, but it isn't necessary
                            ' in production builds - just skip bad runs to salvage the load process.
                            Else
                                Exit For
                                'Debug.Print "bad idxDst vs srcBytesPlanarSize: " & idxDst & " vs " & srcBytesPlanarSize
                            End If
                        Next x
                        
                    'High-bit = 0 means COPY [n] bytes directly from source
                    Else
                        
                        'Only copy if > 0 bytes, and if the copy doesn't exceed scanline size
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
                
                'We now have all bytes decompressed into planar format.
                ' Do another pass to swizzle into interleaved.
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
        '
        'Note that despite stack size implications, we need to use Double due to the ranges involved (2^127).
        Dim lutPow2(0 To 255) As Double
        For x = 0 To 255
            lutPow2(x) = (2# ^ CDbl(x - 128))
        Next x
        
        'If the file embeds exposure data, use it by building it directly into the LUT.
        ' (In PhotoDemon, this behavior can be toggled by the
        '  `Tools > Options > Color management > use format-specific color-management data` option.)
        If ColorManagement.UseEmbeddedLegacyProfiles() Then
            If (m_Header.rh_Exposure <> 1#) And (m_Header.rh_Exposure <> 0#) Then
                For x = 0 To 255
                    lutPow2(x) = lutPow2(x) * m_Header.rh_Exposure
                Next x
            End If
        End If
        
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
            
            'For all other values of e, transform mantissas by the shared exponent
            Else
                dstPixelsF(xOffset) = CDbl(r) * lutPow2(e)
                dstPixelsF(xOffset + 1) = CDbl(g) * lutPow2(e)
                dstPixelsF(xOffset + 2) = CDbl(b) * lutPow2(e)
            End If
            
            'Alpha is irrelevant; we do not read or write it.
            'dstPixelsF(xOffset + 3) = 1!
            
        Next x
        
        'If the files are in XYZE format, we need to do a second pass for XYZ to RGB conversion.
        Dim fR As Single, fG As Single, fB As Single
        Dim fX As Single, fY As Single, fZ As Single
        
        If m_Header.rh_isXYZE Then
            
            For x = 0 To dstWidth - 1
                xOffset = x * 4
                fX = dstPixelsF(xOffset)
                fY = dstPixelsF(xOffset + 1)
                fZ = dstPixelsF(xOffset + 2)
                Colors.XYZtosRGB_Float fX, fY, fZ, dstPixelsF(xOffset), dstPixelsF(xOffset + 1), dstPixelsF(xOffset + 2)
            Next x
            
        End If
        
        'Wrap a pointer around the target line in the destination pixel array.
        dstDIB.WrapArrayAroundScanline dstPixels, dstSA, y
        
        'HDR files are always non-gamma-corrected; to match behavior in other software,
        ' we must gamma-correct at 2.2 by default.  (Files can specify their own values but they are *not*
        ' guaranteed to work; for example, a gamma of "0.4545" or "1" could both theoretically mean
        ' "use normal 1/2.2" - we have no way to distinguish between those interpretations.)
        Const GAMMA_CORRECTION_FACTOR As Double = 1# / 2.2
        Dim lutGamma(0 To 255) As Byte
        For x = 0 To 255
            fR = CSng(x)
            fR = (fR / 255!) ^ GAMMA_CORRECTION_FACTOR * 255!
            If (fR > 255!) Then fR = 255!
            If (fR < 0!) Then fR = 0!
            lutGamma(x) = Int(fR + 0.5!)
        Next x
        
        'Iterate pixels and color-convert as-we-go
        For x = 0 To dstWidth - 1
        
            xOffset = x * 4
            r = Int(dstPixelsF(xOffset))
            g = Int(dstPixelsF(xOffset + 1))
            b = Int(dstPixelsF(xOffset + 2))
            
            'Failsafe only (to ensure valid lookups into the LUT)
            If (r < 0) Then r = 0
            If (r > 255) Then r = 255
            If (g < 0) Then g = 0
            If (g > 255) Then g = 255
            If (b < 0) Then b = 0
            If (b > 255) Then b = 255
            
            'Apply gamma-correction now.  (It would be higher-quality to do this while the pixel data is
            ' still in floats, but the perf impact of that is *huge* due to powf performance in VB6.)
            dstPixels(xOffset) = lutGamma(b)
            dstPixels(xOffset + 1) = lutGamma(g)
            dstPixels(xOffset + 2) = lutGamma(r)
            
            'Alpha does not need to be forced to 255 here because we initialized the target surface to opaque white
            'dstPixels(xOffset + 3) = 255
            
        Next x
        
        'Unwrap the line pointer before continuing
        dstDIB.UnwrapArrayFromDIB dstPixels
        
    Next y
    
    'The last bit of post-processing we need to do is apply embedded color corrections, if any.
    If ColorManagement.UseEmbeddedLegacyProfiles() And m_Header.rh_PrimariesPresent Then
        
        If HDR_DEBUG_VERBOSE Then PDDebug.LogAction "Embedded chromaticity data found; applying now..."
        
        'The following code is adopted from pdPNG; PNG's cHRM chunk works nearly identically.
        
        'Extract the embedded primaries and assign them to standard correction matrices
        Dim wpValues() As Double, rgbValues() As Double
        ReDim wpValues(0 To 2) As Double
        ReDim rgbValues(0 To 8) As Double
        wpValues(0) = m_Header.rh_Primaries(6)
        wpValues(1) = m_Header.rh_Primaries(7)
        wpValues(2) = 1#
        
        rgbValues(0) = m_Header.rh_Primaries(0)
        rgbValues(1) = m_Header.rh_Primaries(1)
        rgbValues(2) = 1#
        rgbValues(3) = m_Header.rh_Primaries(2)
        rgbValues(4) = m_Header.rh_Primaries(3)
        rgbValues(5) = 1#
        rgbValues(6) = m_Header.rh_Primaries(4)
        rgbValues(7) = m_Header.rh_Primaries(5)
        rgbValues(8) = 1#
        
        'You might be saying: wait, aren't you double-applying gamma here?  And you'd be correct, we are.
        ' The reason for this is embedded data always requires a 1/2.2 gamma correction to restore it
        ' back to linear, but if an app writes color-correction data, my testing shows that the underlying
        ' data is *always* already linear.  So we must effectively "undo" the previous gamma step here.
        ' (There is probably some minor precision loss involved in this step, but I'm not overly concerned
        ' with it as PSNR remains above 99%.)
        Dim gammaFloat As Double
        gammaFloat = 2.2
        
        'Our next goal is to assemble an ICC profile that matches the embedded chromaticity coordinates.
        ' With this, we can use our normal color-management pipeline to color-correct this DIB.
        Dim srcProfile As pdLCMSProfile
        Set srcProfile = New pdLCMSProfile
        
        'Make sure the chromaticity data passes internal LCMS validation
        Dim chrmOK As Boolean
        chrmOK = srcProfile.CreateCustomRGBProfile(VarPtr(wpValues(0)), VarPtr(rgbValues(0)), gammaFloat)
        If chrmOK Then
            
            Dim dstProfile As pdLCMSProfile
            Set dstProfile = New pdLCMSProfile
            dstProfile.CreateSRGBProfile
            
            'Apply the profile in-place to the current RGB data
            Dim dstTransform As pdLCMSTransform
            Set dstTransform = New pdLCMSTransform
            If dstTransform.CreateInPlaceTransformForDIB(dstDIB, srcProfile, dstProfile, INTENT_PERCEPTUAL, cmsFLAGS_COPY_ALPHA) Then
                
                dstTransform.ApplyTransformToPDDib dstDIB
                Set dstTransform = Nothing
                
                'If the transform was successful, let's embed this ICC profile into the destination image.
                ' This ensures reasonable persistence for the data described by the original chromaticity values.
                Dim profHash As String
                profHash = ColorManagement.AddLCMSProfileToCache(srcProfile, True)
                dstDIB.SetColorProfileHash profHash
                dstDIB.SetColorManagementState cms_ProfileConverted
                
                If (Not dstImage Is Nothing) Then dstImage.SetColorProfile_Original profHash
                
            End If
            
            'Free the destination LCMS profile before exiting
            Set dstProfile = Nothing
            
        End If
        
        'Free the source profile handle, if any
        Set srcProfile = Nothing
        
    End If
    
    ParsePixels = True
    
End Function

'Copy a line of bytes from the source buffer into the temporary uncompressed buffer, without considering RLE.
' idxSrc will be incremented by scanlineWidth.
' NOTE: this function will crash if the source or destination arrays are not initialized to correct sizes.
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

'Default values come from the official spec (https://radsite.lbl.gov/radiance/refer/filefmts.pdf).
' The use of E illuminant (instead of say, D65 ) probably stems from Radiance's history as 3D modeling software,
' vs a capture device (like a camera).
Private Sub ResetHeaderPrimaries()
    With m_Header
        .rh_Primaries(0) = 0.64
        .rh_Primaries(1) = 0.33
        .rh_Primaries(2) = 0.29
        .rh_Primaries(3) = 0.6
        .rh_Primaries(4) = 0.15
        .rh_Primaries(5) = 0.06
        .rh_Primaries(6) = 0.333
        .rh_Primaries(7) = 0.333
    End With
End Sub

'Export the target pdImage's pixel data to an HDR-format file.  PD does not display an export UI for this format;
' however, some toggles may be available via compile-time settings (for testing purposes only).
Friend Function SaveHDR_ToFile(ByRef srcImage As pdImage, ByRef dstFile As String) As Boolean

    Const FUNC_NAME As String = "SaveHDR_ToFile"
    SaveHDR_ToFile = False
    
    'This exists only as a failsafe; potential encoding errors are explicitly handled via other means
    On Error GoTo SaveFailed
    
    'HDR doesn't support multiple frames.
    ' As such, grab a composited copy of the current image.
    Dim compositeDIB As pdDIB
    srcImage.GetCompositedImage compositeDIB, True
    compositeDIB.CompositeBackgroundColor 255, 255, 255
    
    'Open a stream (memory-mapped) on the target file
    Set m_Stream = New pdStream
    If (Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFile, compositeDIB.GetDIBWidth * compositeDIB.GetDIBHeight * 2, 0, OptimizeSequentialAccess)) Then
        InternalError FUNC_NAME, "couldn't start file stream: " & dstFile
        Exit Function
    End If
    
    'HDR export is very well-optimized so this is gonna be quick, but let's initialize a progress bar anyway.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax compositeDIB.GetDIBHeight()
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Radiance headers are plain-text, which simplifies the first chunk of export work.
    Dim cString As pdString
    Set cString = New pdString
    
    'Start with a magic number
    cString.Append "#?RADIANCE"
    cString.AppendLineBreakLf
    
    'Software comes next (so users know who to blame when export didn't work, ha)
    cString.Append "SOFTWARE=" & Updates.GetPhotoDemonNameAndVersion()
    cString.AppendLineBreakLf
    
    'PD uses the standard exposure value to improve compatibility.  (Specifically, importers can ignore
    ' this value and they'll still obtain a correct image.)
    cString.Append "EXPOSURE=1.0"
    cString.AppendLineBreakLf
    
    'Only RGBE export is used (I have the code to export in XYZE, but XYZE support in decoders is rare)
    cString.Append "FORMAT=32-bit_rle_rgbe"
    cString.AppendLineBreakLf
    
    'A null-line marks the end of the header
    cString.AppendLineBreakLf
    
    'One single line of text follows, containing image orientation and dimensions
    cString.Append "-Y " & compositeDIB.GetDIBHeight() & " +X " & compositeDIB.GetDIBWidth()
    cString.AppendLineBreakLf
    
    'Pixel data follows, so dump the header out now; we are finished with it
    m_Stream.WriteString_ASCII cString.ToString()
    Set cString = Nothing
    
    'Boilerplate for pixel export data
    Dim x As Long, y As Long
    Dim srcWidth As Long, srcHeight As Long, srcWidthScanline As Long
    srcWidth = compositeDIB.GetDIBWidth()
    srcWidthScanline = srcWidth * 4
    srcHeight = compositeDIB.GetDIBHeight()
    
    'Fast ubyte > float conversion, with 2.2 gamma pre-calculated.
    ' (Note: an sRGB gamma ramp would technically be more accurate here given that we're coming from sRGB,
    '  but in my testing, other HDR loaders *always* assume 2.2 gamma and apply it irreversibly at load,
    '  regardless of the target color space you specify.  Thus this conversion is most likely to produce
    '  expected behavior in other software.)
    Dim btof(0 To 255) As Single
    For x = 0 To 255
        btof(x) = (CSng(x) / 255!) ^ 2.2!
    Next x
    
    'Floating-point HDR values need to be decomposed into a shared power-of-two exponent
    ' (used by all three color components) and a per-color 8-byte mantissa.
    ' VB lacks the necessary low-level calls to do these intrinsically, so we'll need to calculate
    ' them manually.  Because we are (currently) only ever exporting from 8-bit sources, we can
    ' precalculate LUTs for various values and *hugely* improve export performance.
    Dim exponent As Long
    Dim lutExponent(0 To 255) As Byte    '(technically -128 to 127, but VB lacks signed chars)
    Dim lutMantissa(0 To 255) As Single
    
    'These LUTs can skip entry 0 as it causes a DBZ error.  (0 also represents pure black -
    ' because the LUT is indexed on the MAX channel value for each pixel - and index 0 will
    ' never trigger on the inner loop because we explicitly branch on that state.)
    For x = 1 To 255
        
        'In C, for a base-2 calculation this is as simple as calling std::frexp.
        ' In VB6, we don't have frexp so we need to manually calculate it (sigh).
        Const LOG_2 As Double = 0.693147180559945
        exponent = Int(Log(btof(x)) / LOG_2) + 1
        
        'Pre-scale the exponent to byte range (to work around a lack of signed char)
        lutExponent(x) = exponent + 128
        
        'Similarly, pre-scale the mantissa to byte range because that's how it is ultimately encoded
        lutMantissa(x) = (1! / (2 ^ exponent)) * 255!
        
    Next x
    
    'We have everything we need to start writing pixel data.  HDR files use a novel 4-bytes-per-3-channels
    ' floating-point representation, where the mantissa of each pixel channel is stored, followed by a
    ' shared power-of-2 exponent for those 3 mantissas.
    
    'First, we'll translate the current BGRA line into the required RGB floating-point representation
    Dim dstBytes() As Byte
    ReDim dstBytes(0 To srcWidth * 4 - 1) As Byte
    
    'RLE compression requires additional buffers for temp storage
    Dim dstBytesPlanar() As Byte
    ReDim dstBytesPlanar(0 To srcWidth * 4 - 1) As Byte
    
    Dim dstBytesRLE() As Byte
    ReDim dstBytesRLE(0 To srcWidth * 6 - 1) As Byte    'Worst-case scenario is 1 extra byte for every 127 bytes in,
                                                        ' so this sizing is way beyond even the worst-case scenario.
    
    'Because HDR images tend to be large, we can greatly improve performance by batching scanline writes
    ' before sending them out to file.  As such, we're going to use a temporary in-memory stream to hold
    ' our results, and then we'll flush those results out to file every [n] lines.  This achieves a nice
    ' balance between perf and memory usage, particularly when writing to external storage.
    '
    '(This value must always be of the form (2^n)-1 because it's &&'d and compared to 0 on the inner loop)
    Const NUM_LINES_TO_BUFFER As Long = 127
    
    Dim bufStream As pdStream
    Set bufStream = New pdStream
    If (Not bufStream.StartStream(PD_SM_MemoryBacked, PD_SA_ReadWrite, startingBufferSize:=srcWidth * 4 * NUM_LINES_TO_BUFFER)) Then
        InternalError FUNC_NAME, "couldn't create output buffer"
        Exit Function
    End If
    
    'RLE encoding is limited to images with widths on the range [2, 32767]
    Dim useRLE As Boolean
    useRLE = (srcWidth >= 2) And (srcWidth <= 32767)
    
    'HDR data is encoded one scanline at a time (each scanline starts with a mandatory 4-byte header)
    For y = 0 To srcHeight - 1
        
        'Point an array at the next scanline of source pixels
        Dim srcPixels() As Byte, srcSA As SafeArray1D
        compositeDIB.WrapArrayAroundScanline srcPixels, srcSA, y
        
        'Iterate pixels, converting as we go
        For x = 0 To srcWidth - 1
            
            'BGRA data means each source pixel occupies four neighboring bytes
            Dim xOffset As Long
            xOffset = x * 4
            
            'Find the maximum channel for this pixel; we'll use its nearest power-of-two value
            ' as the exponent for all channels in this pixel.
            Dim b As Long, g As Long, r As Long, bMax As Long
            b = srcPixels(xOffset)
            g = srcPixels(xOffset + 1)
            r = srcPixels(xOffset + 2)
            If (g > b) Then bMax = g Else bMax = b
            If (r > bMax) Then bMax = r
            
            'If the pixel is pure black, write it as-is
            If (bMax = 0) Then
                dstBytes(xOffset) = 0
                dstBytes(xOffset + 1) = 0
                dstBytes(xOffset + 2) = 0
                dstBytes(xOffset + 3) = 0
                
            'Not-black
            Else
                
                'We now need to decompose the pixel color triplet into a shared exponent + 3 mantissas.
                ' Lucky for us, we precalculated all required values before entering this loop!
                '
                'Retrieve the mantissa multiplier for the calculated power of two.
                ' (Remember: this value encompasses a lot of pre-calculated data, like gamma correction,
                '  power-of-two basing, final scaling, etc.)
                Dim fixed_pow_2 As Single
                fixed_pow_2 = lutMantissa(bMax)
                
                'Convert the color channels and write them out, followed by the shared exponent (normalized to [0, 255])
                dstBytes(xOffset) = Int(btof(r) * fixed_pow_2)
                dstBytes(xOffset + 1) = Int(btof(g) * fixed_pow_2)
                dstBytes(xOffset + 2) = Int(btof(b) * fixed_pow_2)
                dstBytes(xOffset + 3) = lutExponent(bMax)
                
            End If
            
        Next x
        
        'Release the unsafe pixel wrapper from the source pixels.
        compositeDIB.UnwrapArrayFromDIB srcPixels
        
        'We now need to RLE-encode this line before dumping it to file.
        ' (Conversely, the other branch of this statement dumps data in raw, uncompressed form.)
        If useRLE Then
            
            'Before doing *anything* with this scanline, we first need to convert it from
            ' interleaved (RGBRGBRGB...) to planar (RRRGGGBBB...)
            For x = 0 To srcWidth - 1
                xOffset = x * 4
                dstBytesPlanar(x) = dstBytes(xOffset)
                dstBytesPlanar(srcWidth + x) = dstBytes(xOffset + 1)
                dstBytesPlanar(srcWidth * 2 + x) = dstBytes(xOffset + 2)
                dstBytesPlanar(srcWidth * 3 + x) = dstBytes(xOffset + 3)
            Next x
            
            'Each scanline has to start with a specific 4-byte marker.  These magic numbers come from the spec.
            Dim idxDst As Long: idxDst = 0
            dstBytesRLE(0) = 2
            dstBytesRLE(1) = 2
            
            'Width of the image, encoded in two bytes (msb first)
            dstBytesRLE(2) = srcWidth \ 256
            dstBytesRLE(3) = srcWidth Mod 256
            
            'Actual RLE contents start at offset 4 (byte 5)
            idxDst = 4
            
            'RLE is pretty simple from here on out.  There's a one-byte preface; if >= 128 it indicates
            ' an [n-128] length run of the next byte.  If < 128 it indicates a copy of the next [n] bytes.
            ' (Yes, worthless lengths of 0 are allowed by the spec, and do appear in real-world images.)
            Dim idxSrc As Long: idxSrc = 0
            
            '(Also, normally, you'd need to provide special coverage here for 1-pixel images, but we've
            ' already checked for images with width < 2 and branched above.)
            
            'Run and literal lengths are 7-bits long and again, 0 is a valid value (pointless, but valid),
            ' so we can only encode runs/literals up to 127.  The extra high bit distinguishes run vs literal.
            Const MAXIMUM_RUN_LENGTH As Long = 127
            Const MAXIMUM_LITERAL_LENGTH As Long = 127
            
            Dim curRunLength As Long: curRunLength = 0
            Dim targetByte As Byte
            
            'We now need to compress each channel separately.
            ' (Runs cannot extend between channel boundaries, even though flag chunks only appear
            ' at the start of each *scanline*, NOT each *channel* - I tested this and while it compresses better,
            ' it breaks a number of 3rd-party decoders, incluing Photoshop.)
            Dim channelOffset As Long: channelOffset = 0
            
            Dim idxChannel As Long
            For idxChannel = 0 To 3
                
                idxSrc = 0
                
                'Iterate source pixels until we reach end-of-line
                Do
                    
                    'Previous iterations, if any, will have advanced the source pointer to a new byte, one that does
                    ' not belong to the previous segment (repeats *or* literals).  Our job is to analyze this byte
                    ' and its immediate neighbor to determine if we should start a new repeat run or literal run.
                    
                    'Start with the special case of "only one byte left in the stream".  This will only occur if
                    ' we just ended a repeat run, and the last byte in the line is either...
                    ' 1) another repeat, but the previous run hit the maximum repeat length, or...
                    ' 2) a unique value that requires its own segment.
                    '
                    'To avoid slow EOF checks in the inner loop, we want to cover this as a special case
                    ' (so the branch predictor can predictably ignore it).
                    If (idxSrc = srcWidth - 1) Then
                    
                        'Write out the final byte as a 1-byte literal
                        dstBytesRLE(idxDst) = 1
                        dstBytesRLE(idxDst + 1) = dstBytesPlanar(channelOffset + idxSrc)
                        idxDst = idxDst + 2
                        idxSrc = idxSrc + 1
                        
                    'We have room to check neighboring pixels for repeats, without worrying about buffer overflow.
                    Else
                        
                        'Compare this byte to its immediate neighbor.  If they match, start a new *repeat* run.
                        targetByte = dstBytesPlanar(channelOffset + idxSrc)
                        If (targetByte = dstBytesPlanar(channelOffset + idxSrc + 1)) Then
                            
                            'We already know that this byte and its neighbor match, so start the count at 2.
                            curRunLength = 2
                            
                            'Make sure we don't run past the end of the channel.
                            Dim maxRunLength As Long
                            If (srcWidth - idxSrc < MAXIMUM_RUN_LENGTH) Then maxRunLength = srcWidth - idxSrc Else maxRunLength = MAXIMUM_RUN_LENGTH
                            
                            'Count the number of repeats.
                            Do While (curRunLength < maxRunLength)
                                If (dstBytesPlanar(channelOffset + idxSrc + curRunLength) = targetByte) Then
                                    curRunLength = curRunLength + 1
                                Else
                                    Exit Do
                                End If
                            Loop
                            
                            'curRunLength equals the actual size of the run.
                            ' Encode it in the destination and advance both pointers (src and dst) to match.
                            dstBytesRLE(idxDst) = 128 + curRunLength
                            dstBytesRLE(idxDst + 1) = targetByte
                            idxDst = idxDst + 2
                            idxSrc = idxSrc + curRunLength
                        
                        'This byte and its neighbor do *not* match.  Start a new literal run.
                        Else
                            
                            'We already know the next two bytes do *not* match, so start the count at 2.
                            curRunLength = 2
                            
                            'Make sure we don't run past the end of the channel
                            If (srcWidth - idxSrc < MAXIMUM_LITERAL_LENGTH) Then maxRunLength = srcWidth - idxSrc Else maxRunLength = MAXIMUM_LITERAL_LENGTH
                            
                            'Count the number of non-repeats that follow these two non-repeating bytes
                            Do While (curRunLength < maxRunLength)
                                
                                'The next byte is also novel
                                If (dstBytesPlanar(channelOffset + idxSrc + curRunLength - 1) <> dstBytesPlanar(channelOffset + idxSrc + curRunLength)) Then
                                    curRunLength = curRunLength + 1
                                
                                'The next two bytes are identical, which raises the question:
                                ' should we start a new RLE packet, or simply continue this stretch of literals?
                                Else
                                    
                                    'See if this new repeat run is at least three bytes long.  If it isn't, we usually
                                    ' *improve* compression by keeping those two matching bytes inside this literal
                                    ' run, because...
                                    ' 1) we've already eaten the "cost" of the flag byte for this run, and...
                                    ' 2) since byte 3 is novel and would just start a new literal run, it would
                                    '    incur a new flag byte cost - which we can avoid by just keeping the current
                                    '    run intact!
                                    '
                                    ' (We can only do this check if there's space remaining in the channel, however.)
                                    If (curRunLength + 1 < maxRunLength) Then
                                        
                                        If (dstBytesPlanar(channelOffset + idxSrc + curRunLength) = dstBytesPlanar(channelOffset + idxSrc + curRunLength + 1)) Then
                                            
                                            'Byte three *does* match, so we should use an RLE packet to encode it.
                                            Exit Do
                                        
                                        'These two bytes are immediately followed by a different byte, so leave them
                                        ' inside the current literal run.
                                        Else
                                            curRunLength = curRunLength + 1
                                        End If
                                    
                                    'We're already at maximum length for this run, so we couldn't extend it even if
                                    ' we wanted to.
                                    Else
                                        Exit Do
                                    End If
                                    
                                End If
                                
                            Loop
                            
                            'curRunLength now represents the size of this run of non-repeating characters.
                            ' Encode it in the destination and advance all pointers to match.
                            dstBytesRLE(idxDst) = curRunLength
                            CopyMemoryStrict VarPtr(dstBytesRLE(idxDst + 1)), VarPtr(dstBytesPlanar(channelOffset + idxSrc)), curRunLength
                            idxDst = idxDst + curRunLength + 1  'The +1 is for the flag byte (marking run length)
                            idxSrc = idxSrc + curRunLength
                            
                        End If
                    
                    End If
                    
                Loop While (idxSrc < srcWidth)
                
                'Offset the channel offset into the source by (srcWidth) bytes
                channelOffset = channelOffset + srcWidth
                
            Next idxChannel
            
            'Dump the RLE-compressed bytes into the buffer stream
            bufStream.WriteBytesFromPointer VarPtr(dstBytesRLE(0)), idxDst
        
        'This branch writes uncompressed data.  This does produce a valid HDR file but the file will also be huge.
        Else
            
            'Dump the whole line as-is into the buffer stream
            bufStream.WriteBytesFromPointer VarPtr(dstBytes(0)), srcWidthScanline
            
        End If
        
        'Periodically dump the memory buffer out to file, resetting the buffer's stream pointer as necessary.
        If (y And NUM_LINES_TO_BUFFER = 0) Then
            bufStream.SetPosition 0, FILE_BEGIN
            m_Stream.WriteBytesFromPointer bufStream.ReadBytes_PointerOnly(bufStream.GetStreamSize()), bufStream.GetStreamSize()
            bufStream.SetSizeExternally 0
        End If
        
        'Periodically update the progress bar (see top of class for how this increment is determined)
        If (y And progBarCheck) = 0 Then ProgressBars.SetProgBarVal y
        
    Next y
    
    'Dump the final buffer bytes (if any) out to disk
    If (bufStream.GetPosition > 0) Then
        bufStream.SetPosition 0, FILE_BEGIN
        m_Stream.WriteBytesFromPointer bufStream.ReadBytes_PointerOnly(bufStream.GetStreamSize()), bufStream.GetStreamSize()
        Set bufStream = Nothing
    End If
    
    '*Don't* free the progress bar here.  (Freeing is left to the caller; this prevents flicker if additional
    ' export steps occur, like copying the file over an already-named instance.)
    'ProgressBars.ReleaseProgressBar
    
    'No trailing bytes are required after pixel data!  Close the stream and exit.
    m_Stream.StopStream True
    SaveHDR_ToFile = True
    
    Exit Function
    
SaveFailed:
    InternalError FUNC_NAME, "abandoned save due to critical error"
    SaveHDR_ToFile = False
    
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen Then m_Stream.StopStream True
    End If
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdHDR." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdHDR." & funcName & "() reported an error: " & errDescription
    End If
End Sub
