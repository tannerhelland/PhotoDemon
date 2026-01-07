VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdQOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Quite OK Image" (QOI) Decoder/Encoder
'Copyright 2022-2026 by Tanner Helland
'Created: 24/February/22
'Last updated: 25/February/22
'Last update: wrap up initial build
'
'The QOI image format was developed by Dominic Szablewski.  It is designed as a compact, lossless
' image format with much greater simplicity than e.g. PNG but much smaller filesizes than e.g. BMP.
' A final 1.0 spec was released in 2022 and PhotoDemon uses that spec as the basis for its
' custom VB6 implementation.  You can find the QOI spec here (link good as of Feb 2022):
'
' https://qoiformat.org/qoi-specification.pdf
'
'Decoding and encoding in native VB6 code (for better or worse) are both implemented here.
'
'I referenced the QOI reference implementation (qoi.h) while developing this implementation.
' qoi.h is MIT-licensed and Copyright 2021 Dominic Szablewski.  License text of qoi.h follows:
'
'   QOI - The "Quite OK Image" format for fast, lossless image compression
'   Dominic Szablewski - https://phoboslab.org
'
'   -- LICENSE: The MIT License (MIT)
'
'   Copyright(c) 2021 Dominic Szablewski
'
'   Permission is hereby granted, free of charge, to any person obtaining a copy of
'   this software and associated documentation files(the "Software"), to deal in
'   the Software without restriction, including without limitation the rights to
'   use, copy, modify, merge, publish, distribute, sublicense, and / or sell copies
'   of the Software, and to permit persons to whom the Software is furnished to do
'   so, subject to the following conditions :
'
'   The above copyright notice and this permission notice shall be included in all
'   copies or substantial portions of the Software.
'
'   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
'   AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'   LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'   OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'   SOFTWARE.
'
'Thank you to Dominic for this neat concept!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps a variety of
' diagnostic information to the debug log.
Private Const QOI_DEBUG_VERBOSE As Boolean = False

'Literals taken from qoi.h (https://github.com/phoboslab/qoi/blob/master/qoi.h)
Private Const QOI_OP_INDEX  As Byte = &H0   '/* 00xxxxxx */
Private Const QOI_OP_DIFF   As Byte = &H40  '/* 01xxxxxx */
Private Const QOI_OP_LUMA   As Byte = &H80  '/* 10xxxxxx */
Private Const QOI_OP_RUN    As Byte = &HC0  '/* 11xxxxxx */
Private Const QOI_OP_RGB    As Byte = &HFE  '/* 11111110 */
Private Const QOI_OP_RGBA   As Byte = &HFF  '/* 11111111 */

'Masks for 2-bit flags followed by 6-bits of info
Private Const QOI_MASK_2    As Byte = &HC0  '/* 11000000 */
Private Const QOI_MASK_L6   As Byte = &H3F  '/* 00111111 */

'QOI files only support a few relevant items in the header:

'Image width/height, in pixels (big-endian)
Private m_qoiWidth As Long, m_qoiHeight As Long

'Channel count (3 = RGB, 4 = RGBA)
Private m_qoiChannelCount As Long

'Color space (1 = sRGB with linear alpha, 2 = all channels linear)
Private Enum QOI_ColorSpace
    cs_sRGB = 0
    cs_Linear = 1
End Enum

#If False Then
    Private Const cs_sRGB = 0, cs_Linear = 1
#End If

Private m_qoiColorSpace As QOI_ColorSpace

'Other relevant notes from the spec (https://qoiformat.org/qoi-specification.pdf):
' - Colorspace and channel fields are purely informative. They do not change the way data chunks are encoded.
' - Images are encoded row by row, left to right, top to bottom.
' - Decoder and encoder start with {r: 0, g: 0, b: 0, a: 255} as the previous pixel value.
' - An image is complete when all pixels specified by width * height have been covered.
' - Color channels are assumed to not be premultiplied with the alpha channel (“un-premultiplied alpha”).

'Byte-by-byte access is provided, as always, by a pdStream instance - but note that this class dumps
' the core data stream to local arrays before certain parsing duties (for perf reasons)
Private m_Stream As pdStream

'Only returns valid information after a successful call to LoadQOI
Friend Function GetOriginalChannelCount() As Long
    GetOriginalChannelCount = m_qoiChannelCount
End Function

'Validate a source filename as QOI format.  Validation *does* touch the file - we must validate
' a "magic number" in the header.
Friend Function IsFileQOI(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True, Optional ByVal onSuccessLeaveStreamOpen As Boolean = False) As Boolean
    
    Dim potentiallyQOI As Boolean
    potentiallyQOI = Files.FileExists(srcFilename)
    If potentiallyQOI Then potentiallyQOI = (Files.FileLenW(srcFilename) > 22)
    
    'Check extension up front, if requested.
    If (potentiallyQOI And requireValidFileExtension) Then
        potentiallyQOI = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "qoi", True)
    End If
    
    'Proceed with deeper validation as necessary
    If potentiallyQOI Then
        
        'Attempt to load the file
        Set m_Stream = New pdStream
        If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename) Then
            
            'The first 4 bytes of a QOI file must be the ASCII values "qoif"
            potentiallyQOI = (m_Stream.ReadString_ASCII(4) = "qoif")
            If (potentiallyQOI And QOI_DEBUG_VERBOSE) Then PDDebug.LogAction "Valid QOI file found"
            
        End If
        
    End If
    
    IsFileQOI = potentiallyQOI
    If (Not IsFileQOI) Or (Not onSuccessLeaveStreamOpen) Then Set m_Stream = Nothing
    
End Function

'Validate and load a candidate QOI file
Friend Function LoadQOI_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "LoadQOI_FromFile"
    LoadQOI_FromFile = False
    
    'Validate the file
    If Me.IsFileQOI(srcFile, False, True) Then
        
        'If validation passed, m_Stream has already validated the first 4 bytes and now points
        ' at the second member of the header struct:
        ' qoi_header {
        '    char magic[4]; // magic bytes "qoif"
        '    uint32_t width; // image width in pixels (BE)
        '    uint32_t height; // image height in pixels (BE)
        '    uint8_t channels; // 3 = RGB, 4 = RGBA
        '    uint8_t colorspace; // 0 = sRGB with linear alpha, 1 = all channels linear
        ' };
        m_qoiWidth = m_Stream.ReadLong_BE()
        m_qoiHeight = m_Stream.ReadLong_BE()
        m_qoiChannelCount = m_Stream.ReadByte()
        m_qoiColorSpace = m_Stream.ReadByte()
        
        'Quick validation on key members
        If (m_qoiWidth <= 0) Or (m_qoiHeight <= 0) Then
            InternalError FUNC_NAME, "bad dimensions: " & m_qoiWidth & "x" & m_qoiHeight
            Set m_Stream = Nothing
            Exit Function
        End If
        
        If (m_qoiChannelCount < 3) Or (m_qoiChannelCount > 4) Then
            InternalError FUNC_NAME, "bad channel count: " & m_qoiChannelCount
            Set m_Stream = Nothing
            Exit Function
        End If
        
        If (m_qoiColorSpace < 0) Or (m_qoiColorSpace > 1) Then InternalError FUNC_NAME, "unknown color space (will attempt to load anyway): " & m_qoiColorSpace
        
        'Debug data is only relevant under certain circumstances
        If QOI_DEBUG_VERBOSE Then
            PDDebug.LogAction "QOI dimensions: " & m_qoiWidth & "x" & m_qoiHeight
            PDDebug.LogAction "QOI channels and color space: " & m_qoiChannelCount & ", " & m_qoiColorSpace
        End If
        
        'Prep image buffer
        Set dstDIB = New pdDIB
        If dstDIB.CreateBlank(m_qoiWidth, m_qoiHeight, 32, vbWhite, 255) Then
            
            'Hand the rest of the load process off to the parser
            LoadQOI_FromFile = LoadQOI_ParsePixels(dstDIB)
            
            'If successful, premultiply alpha (as required by PD)
            If LoadQOI_FromFile Then dstDIB.SetAlphaPremultiplication True
            
            'Regardless of outcome, free the underlying stream
            Set m_Stream = Nothing
            
        Else
            InternalError FUNC_NAME, "out of memory"
            Set m_Stream = Nothing
            Exit Function
        End If
        
    End If
    
End Function

'DO NOT CALL outside of LoadQOI_FromFile; for perf reasons, this function assumes correct stream alignment,
' initialization of all header members, and prepped destination buffer.
'
'Note that this VB6 implementation is noticeably more verbose than the original C version.  This is necessary
' because VB doesn't support bitshifts, and a lack of unsigned ints means "wraparound" over/underflow must
' be implemented manually.
Private Function LoadQOI_ParsePixels(ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "LoadQOI_ParsePixels"
    LoadQOI_ParsePixels = False
    
    On Error GoTo BadDataStream
    
    'Reset previous pixel trackers
    Dim prevR As Long, prevG As Long, prevB As Long, prevA As Long
    prevR = 0
    prevG = 0
    prevB = 0
    prevA = 255
    
    'LUT of previous pixel values (size fixed per spec)
    Dim prevPixels(0 To 63) As RGBQuad
    Dim indexPos As Long
    
    'Wrap a 1D array around the destination buffer
    Dim dstPixels() As RGBQuad, dstSA As SafeArray1D
    dstDIB.WrapRGBQuadArrayAroundDIB_1D dstPixels, dstSA
    
    'Pixel offsets are tracked to avoid OOB on bad files
    Dim dstOffset As Long: dstOffset = 0
    Dim numPixels As Long: numPixels = m_qoiWidth * m_qoiHeight
    Dim nextByte As Long, curRun As Long
    
    'For perf reasons, we don't want to tap the stream object for single-byte reads.
    ' Instead, grab the entire run of bytes and store it locally.
    Dim srcBytes() As Byte, numSrcBytes As Long
    
    'Run size is (length of the file) - 14-byte header - 8-byte trailer
    Const QOI_FIXED_HEADER_SIZE As Long = 14
    Const QOI_FIXED_TRAILER_SIZE As Long = 8
    numSrcBytes = m_Stream.GetStreamSize() - (QOI_FIXED_HEADER_SIZE + QOI_FIXED_TRAILER_SIZE)
    If (numSrcBytes <> m_Stream.ReadBytes(srcBytes, numSrcBytes)) Then
        InternalError FUNC_NAME, "not enough source data"
        LoadQOI_ParsePixels = False
        Exit Function
    End If
    
    'The source stream object is no longer required.
    m_Stream.StopStream True
    
    Dim srcOffset As Long: srcOffset = 0
    
    'Parse until we fill all destination pixels.  Note that we do not explicitly check
    ' safety on source pixel accesses, so this function may error on malformed files
    ' with truncated source bytes.  I'm okay with this - the reference decoder works the
    ' same way, and there are large perf repercussions for safety checks on every source
    ' access.
    Do
        
        'Handle runs as a special case
        If (curRun > 0) Then
            curRun = curRun - 1
        
        'Anything else requires us to pull 1+ bytes from the source stream
        Else
            
            'Retrieve the next byte
            nextByte = srcBytes(srcOffset)
            srcOffset = srcOffset + 1
            
            'Look for special 8-bit flags first
            
            'Explicit RGBA values follow
            If (nextByte = QOI_OP_RGBA) Then
                prevR = srcBytes(srcOffset)
                prevG = srcBytes(srcOffset + 1)
                prevB = srcBytes(srcOffset + 2)
                prevA = srcBytes(srcOffset + 3)
                srcOffset = srcOffset + 4
            
            'Explicit RGB values follow
            ElseIf (nextByte = QOI_OP_RGB) Then
            
                'Explicit RGB values follow.  Read directly into the previous pixel holder (which conveniently
                ' reuses its alpha value, which is required behavior per the spec) but don't forget to swizzle R/B.
                prevR = srcBytes(srcOffset)
                prevG = srcBytes(srcOffset + 1)
                prevB = srcBytes(srcOffset + 2)
                srcOffset = srcOffset + 3
                
            'Remaining options are 2-bit flags (2 MSBs) followed by 6-bit indicators.
            Else
                
                Dim high2b As Long, low6b As Long
                high2b = (nextByte And QOI_MASK_2)
                low6b = (nextByte And QOI_MASK_L6)
                
                'Many of these flags require us to calculate diffs between pixels
                Dim rDiff As Long, gDiff As Long, bDiff As Long
                
                Select Case high2b
                    
                    'Low 6-bits are an index into the previous color array [0-63]
                    Case QOI_OP_INDEX
                        With prevPixels(low6b)
                            prevR = .Red
                            prevG = .Green
                            prevB = .Blue
                            prevA = .Alpha
                        End With
                        
                    'Low 6-bits are 2-bit differences (wraparound) between current RGB and previous RGB values,
                    ' so the range [0, 3] must be translated to the range [-2, 1]
                    Case QOI_OP_DIFF
                        prevR = prevR + ((low6b \ 16) And &H3) - 2
                        prevG = prevG + ((low6b \ 4) And &H3) - 2
                        prevB = prevB + (low6b And &H3) - 2
                        
                        'Manually handle wraparound
                        If (prevR >= 0) Then prevR = prevR And &HFF Else prevR = prevR + 256
                        If (prevG >= 0) Then prevG = prevG And &HFF Else prevG = prevG + 256
                        If (prevB >= 0) Then prevB = prevB And &HFF Else prevB = prevB + 256
                        
                    'Low 6-bits are a 6-bit difference (wraparound) between current G and previous G values,
                    ' (translated from [0, 63] to [-32, 31]).  Another byte must be read for difference
                    ' in R, B values (4-bits translated from [0, 15] to [-8, 7] and added to G diff).
                    Case QOI_OP_LUMA
                        gDiff = low6b - 32
                        nextByte = srcBytes(srcOffset)
                        srcOffset = srcOffset + 1
                        rDiff = (nextByte \ 16) - 8 + gDiff
                        bDiff = (nextByte And &HF) - 8 + gDiff
                        
                        'Calculate new values, with manual coverage for wraparound
                        prevR = prevR + rDiff
                        prevG = prevG + gDiff
                        prevB = prevB + bDiff
                        
                        If (prevR >= 0) Then prevR = prevR And &HFF Else prevR = prevR + 256
                        If (prevG >= 0) Then prevG = prevG And &HFF Else prevG = prevG + 256
                        If (prevB >= 0) Then prevB = prevB And &HFF Else prevB = prevB + 256
                    
                    'Simple RLE encoding does not require explicit handling.  Simply set the "run"
                    ' flag to allow automatic use of prevR/G/B.  (A branch at the top of the Do Loop
                    ' checks this flag and decrements it on each iteration.)
                    Case QOI_OP_RUN
                        curRun = low6b
                    
                End Select
                
            End If
            
            'The contents of prevPixel will now reflect the color we just read.
            ' It must always be added to the recent color table.
            
            'Generate a simple hash into the lookup table using the formula from the spec:
            ' index_position = (r * 3 + g * 5 + b * 7 + a * 11) % 64
            indexPos = prevR * 3 + prevG * 5 + prevB * 7 + prevA * 11
            With prevPixels(indexPos And QOI_MASK_L6)
                .Blue = prevB
                .Green = prevG
                .Red = prevR
                .Alpha = prevA
            End With
                
        End If
        
        'Write the pixel value into the destination image
        With dstPixels(dstOffset)
            .Blue = prevB
            .Green = prevG
            .Red = prevR
            .Alpha = prevA
        End With
        
        'Increment pixel offset into the destination image and continue
        dstOffset = dstOffset + 1
        
    Loop While (dstOffset < numPixels)
    
    'Safely unwrap array wrapper and return
    dstDIB.UnwrapRGBQuadArrayFromDIB dstPixels
    LoadQOI_ParsePixels = (dstOffset = numPixels)
    
    'In debug mode, require an extra validation against the source data length
    If QOI_DEBUG_VERBOSE Then LoadQOI_ParsePixels = LoadQOI_ParsePixels And (srcOffset = numSrcBytes)
    
    'Remaining code must *only* be error-handler related
    Exit Function
    
BadDataStream:
    InternalError FUNC_NAME, "internal VB error #" & Err.Number & ": " & Err.Description
    LoadQOI_ParsePixels = False
    
End Function

'Save an arbitrary DIB to a standalone QOI file.  An optional parent pdImage object can also be passed,
' but it is not currently required (or even used).
Friend Function SaveQOI_ToFile(ByRef dstFile As String, ByRef srcImage As pdImage, ByRef srcDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "SaveQOI_ToFile"
    SaveQOI_ToFile = False
    
    On Error GoTo SaveFailed
    
    'QOI files are very easy to save.  There are no user-controlled parameters, and saving is deterministic.
    
    'The only things we really need to know about the source image are its...
    ' 1) Size (width, height)
    ' 2) Color-depth (alpha presence)
    ' 3) Color-management state (linear or sRGB)
    
    'PhotoDemon guarantees that (2) will always be TRUE and (3) will always be sRGB, which makes this very simple!
    
    'Start by opening a pdStream instance on the target file, then writing a fixed-size header.
    Set m_Stream = New pdStream
    If (Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFile, optimizeAccess:=OptimizeSequentialAccess)) Then
        InternalError FUNC_NAME, "couldn't initialize stream on target file"
        SaveQOI_ToFile = False
        Exit Function
    End If
    
    'Before writing the header, scan the source image to see if it contains any non-opaque pixels.
    ' (If it does NOT, we can omit the alpha channel entirely.)
    Dim srcHasAlpha As Boolean
    srcHasAlpha = DIBs.IsDIBTransparent(srcDIB)
    
    ' qoi_header {
    '    char magic[4]; // magic bytes "qoif"
    '    uint32_t width; // image width in pixels (BE)
    '    uint32_t height; // image height in pixels (BE)
    '    uint8_t channels; // 3 = RGB, 4 = RGBA
    '    uint8_t colorspace; // 0 = sRGB with linear alpha, 1 = all channels linear
    ' };
    With m_Stream
        .WriteString_ASCII "qoif"
        .WriteLong_BE srcDIB.GetDIBWidth
        .WriteLong_BE srcDIB.GetDIBHeight
        .WriteByte IIf(srcHasAlpha, 4, 3)
        .WriteByte cs_sRGB
    End With
    
    'Per the spec, the decoder and encoder must keep a running 64-color hash table of the last
    ' 64 colors "encountered" (e.g. colors with unique hashes).
    Dim prevPixels(0 To 63) As RGBQuad
    Dim indexPos As Long
    
    'Track the immediately previous color with both an RGB int (for fast compares)...
    Dim prevRGB As Long
    
    '...and individual ints (because we need to do wraparound math which is cumbersome with bytes in VB6)
    Dim prevR As Long, prevG As Long, prevB As Long, prevA As Long
    prevR = 0
    prevG = 0
    prevB = 0
    prevA = 255
    
    Dim curPixel As RGBQuad, curRGB As Long
    Dim curR As Long, curG As Long, curB As Long, curA As Long
    
    'QOI treats the image as a constant pixel stream (scanlines do not matter)
    Dim srcPixels() As RGBQuad, srcSA As SafeArray1D
    srcDIB.WrapRGBQuadArrayAroundDIB_1D srcPixels, srcSA
    
    'Simple pixel offsets will be tracked to measure progress
    Dim srcOffset As Long: srcOffset = 0
    Dim numPixels As Long: numPixels = srcDIB.GetDIBWidth * srcDIB.GetDIBHeight
    
    'Instead of writing bytes directly to the stream, we'll write them to a local array
    ' and flush the array out to file in arbitrary length chunks (I'm using 64k here,
    ' a number that was plucked out of thin air).
    Const FLUSH_BYTE_LIMIT As Long = 65536
    Dim dstBytes() As Byte
    ReDim dstBytes(0 To 65552) As Byte  'Add a few extra bytes for safety, since we flush after adding new data (max 6 bytes in a single pass)
    Dim dstOffset As Long: dstOffset = 0
    
    'RLE runs will be automatically detected and used preferentially
    Dim curRun As Long: curRun = 0
    
    'Start encoding!
    Do
        
        curPixel = srcPixels(srcOffset)
        
        'Always start by checking for duplicates of the previous pixel
        curRGB = RGB(curPixel.Red, curPixel.Green, curPixel.Blue)
        If (curRGB = prevRGB) And (curPixel.Alpha = prevA) Then
            
            curRun = curRun + 1
            
            'Check for two states:
            ' 1) Max size of possible RLE (62 per the spec)
            ' 2) EOF
            'Either of these indicate an end of run, which means we must write this run
            ' to file and reset the run counter.
            If (curRun = 62) Or (srcOffset = (numPixels - 1)) Then
                dstBytes(dstOffset) = QOI_OP_RUN Or (curRun - 1)
                dstOffset = dstOffset + 1
                curRun = 0
            End If
            
        'This pixel does *not* match the previous pixel
        Else
        
            'If this pixel breaks a running RLE segment, write the in-progress RLE segment
            ' before continuing to process this pixel.
            If (curRun > 0) Then
                dstBytes(dstOffset) = QOI_OP_RUN Or (curRun - 1)
                dstOffset = dstOffset + 1
                curRun = 0
            End If
            
            'Grab int versions of each channel (we need them for anything involving math)
            With curPixel
                curB = .Blue
                curG = .Green
                curR = .Red
                curA = .Alpha
            End With
            
            'Next, check to see if this pixel matches its partner in the hash table
            indexPos = (curR * 3 + curG * 5 + curB * 7 + curA * 11) And QOI_MASK_L6
            If (curRGB = RGB(prevPixels(indexPos).Red, prevPixels(indexPos).Green, prevPixels(indexPos).Blue)) And (curA = prevPixels(indexPos).Alpha) Then
                dstBytes(dstOffset) = QOI_OP_INDEX Or indexPos
                dstOffset = dstOffset + 1
            Else
                
                'Hash table failed.  Update its table entry with this pixel.
                prevPixels(indexPos) = curPixel
                
                'Next, look for mismatched alpha.  All diff flags operate on color channels.
                ' Alpha changes require an explicit RGBA entry.
                If (curA <> prevA) Then
                    dstBytes(dstOffset) = QOI_OP_RGBA
                    dstBytes(dstOffset + 1) = curR
                    dstBytes(dstOffset + 2) = curG
                    dstBytes(dstOffset + 3) = curB
                    dstBytes(dstOffset + 4) = curA
                    dstOffset = dstOffset + 5
                Else
                    
                    'Alpha matches, which means we may be able to get away with a diff calculation
                    ' instead of an explicit color entry.
                    
                    'Calculate RGB diffs now
                    Dim rDiff As Long, gDiff As Long, bDiff As Long
                    rDiff = curR - prevR
                    gDiff = curG - prevG
                    bDiff = curB - prevB
                    
                    'Calculate two more offsets based on the g channel (one diff formula requires these)
                    Dim rgDiff As Long, bgDiff As Long
                    rgDiff = rDiff - gDiff
                    bgDiff = bDiff - gDiff
                    
                    'See if we can get away with a 1-byte diff indicator
                    If ((rDiff > -3) And (rDiff < 2) And (gDiff > -3) And (gDiff < 2) And (bDiff > -3) And (bDiff < 2)) Then
                        dstBytes(dstOffset) = QOI_OP_DIFF Or ((rDiff + 2) * 16) Or ((gDiff + 2) * 4) Or (bDiff + 2)
                        dstOffset = dstOffset + 1
                        
                    'Next, see if we can get away with a 2-byte diff indicator
                    ElseIf ((rgDiff > -9) And (rgDiff < 8) And (gDiff > -33) And (gDiff < 32) And (bgDiff > -9) And (bgDiff < 8)) Then
                        dstBytes(dstOffset) = QOI_OP_LUMA Or (gDiff + 32)
                        dstBytes(dstOffset + 1) = ((rgDiff + 8) * 16) Or (bgDiff + 8)
                        dstOffset = dstOffset + 2
                    
                    'The difference is too large for a diff indicator; write an explicit RGB triplet
                    Else
                        dstBytes(dstOffset) = QOI_OP_RGB
                        dstBytes(dstOffset + 1) = curR
                        dstBytes(dstOffset + 2) = curG
                        dstBytes(dstOffset + 3) = curB
                        dstOffset = dstOffset + 4
                    End If
                
                '/end curA <> prevA
                End If
            
            '/end curPixel = prevPixels(indexPos)
            End If
            
            'Update previous pixel data before continuing
            prevRGB = curRGB
            prevR = curR
            prevG = curG
            prevB = curB
            prevA = curA
            
        End If
        
        'Increment source pixel counter (destination pixel counter is not used)
        srcOffset = srcOffset + 1
        
        'If we have [arbitrary amount] worth of data, flush it to the stream and reset our
        ' intermediary buffer offset.
        If (dstOffset > FLUSH_BYTE_LIMIT) Then
            m_Stream.WriteBytesFromPointer VarPtr(dstBytes(0)), dstOffset
            dstOffset = 0
        End If
        
    Loop While (srcOffset < numPixels)
    
    'If any bytes remain, flush them immediately
    If (dstOffset > 0) Then
        m_Stream.WriteBytesFromPointer VarPtr(dstBytes(0)), dstOffset
        dstOffset = 0
    End If
        
    'Free unsafe array reference
    srcDIB.UnwrapRGBQuadArrayFromDIB srcPixels
    
    'End the file with a mandatory 8-byte trailer
    Dim i As Long
    For i = 0 To 6
        m_Stream.WriteByte 0
    Next i
    m_Stream.WriteByte 1
    
    'Finished!
    m_Stream.StopStream
    SaveQOI_ToFile = True
    Exit Function
    
SaveFailed:
    SaveQOI_ToFile = False
    InternalError FUNC_NAME, "internal VB error #" & Err.Number & ": " & Err.Description
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdQOI." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdQOI." & funcName & "() reported an error: " & errDescription
    End If
End Sub
