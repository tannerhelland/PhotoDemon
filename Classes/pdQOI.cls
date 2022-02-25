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
'PhotoDemon "QUITE OK IMAGE FORMAT" (QOI) Container and Parser
'Copyright 2022-2022 by Tanner Helland
'Created: 24/February/22
'Last updated: 24/February/22
'Last update: start work on initial build
'
'The QOI image format was developed by Dominic Szablewski.  It is meant to be a compact, lossless
' image format with much greater simplicity than e.g. PNG but much smaller filesizes than e.g. BMP.
' A final 1.0 spec was released in 2022 and PhotoDemon uses that spec as the basis for its
' custom VB6 implementation.  You can find the QOI spec here (link good as of Feb 2022):
'
' https://qoiformat.org/qoi-specification.pdf
'
'Only decoding is currently planned.  Encoding may be added in the future.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps all kinds of
' diagnostic information to the debug log.
Private Const QOI_DEBUG_VERBOSE As Boolean = True

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

'LUT of previous pixel values (size fixed per spec)
Private m_PrevPixels(0 To 63) As RGBQuad

'Byte-by-byte access is provided, as always, by a pdStream instance
Private m_Stream As pdStream

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
            
            'QOI header is 14-bytes, with the first 4 bytes as ASCII values "qoif"
            potentiallyQOI = (m_Stream.ReadString_ASCII(4) = "qoif")
            If (potentiallyQOI And QOI_DEBUG_VERBOSE) Then PDDebug.LogAction "Valid QOI file found!"
            
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
        
        'Prep image buffer
        Set dstDIB = New pdDIB
        If dstDIB.CreateBlank(m_qoiWidth, m_qoiHeight, 32, 255, 255) Then
            
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
    
    'Literals taken from qoi.h (https://github.com/phoboslab/qoi/blob/master/qoi.h)
    Const QOI_OP_INDEX  As Byte = &H0   '/* 00xxxxxx */
    Const QOI_OP_DIFF   As Byte = &H40  '/* 01xxxxxx */
    Const QOI_OP_LUMA   As Byte = &H80  '/* 10xxxxxx */
    Const QOI_OP_RUN    As Byte = &HC0  '/* 11xxxxxx */
    Const QOI_OP_RGB    As Byte = &HFE  '/* 11111110 */
    Const QOI_OP_RGBA   As Byte = &HFF  '/* 11111111 */
    
    'Masks for 2-bit flags followed by 6-bits of info
    Const QOI_MASK_2    As Byte = &HC0  '/* 11000000 */
    Const QOI_MASK_L6   As Byte = &H3F  '/* 00111111 */
    
    'Reset previous pixel trackers
    Dim prevPixel As RGBQuad
    With prevPixel
        .Red = 0: .Green = 0: .Blue = 0: .Alpha = 255
    End With
    
    VBHacks.ZeroMemory VarPtr(m_PrevPixels(0)), 64 * 4
        
    'Wrap a 1D array around the destination buffer
    Dim dstPixels() As RGBQuad, dstSA As SafeArray1D
    dstDIB.WrapRGBQuadArrayAroundDIB_1D dstPixels, dstSA
    
    'Pixel offsets will be tracked to avoid OOB on bad files
    Dim pxOffset As Long: pxOffset = 0
    Dim numPixels As Long: numPixels = m_qoiWidth * m_qoiHeight
    
    'A few different read values are used, in part because we need to swizzle from RGBA to BGRA order
    Dim curPixel As RGBQuad
    Dim nextByte As Byte
    Dim tmpColor As Long
    Dim curRun As Long
    
    'Parse until we fill all pixels
    Do
        
        'Handle runs as a special case
        If (curRun > 0) Then
            curRun = curRun - 1
        
        'Anything else requires us to pull 1+ bytes from the source stream
        Else
            
            'Retrieve the next byte
            nextByte = m_Stream.ReadByte()
            
            'Look for special 8-bit flags first
            If (nextByte = QOI_OP_RGBA) Then
                
                'Explicit RGBA values follow.  Read directly into the previous pixel holder so we can swizzle.
                m_Stream.ReadBytesToBarePointer VarPtr(prevPixel), 4
                tmpColor = prevPixel.Red
                prevPixel.Red = prevPixel.Blue
                prevPixel.Blue = tmpColor
                
            ElseIf (nextByte = QOI_OP_RGB) Then
            
                'Explicit RGB values follow.  Read directly into the previous pixel holder (which conveniently
                ' reuses its alpha value, which is required behavior per the spec) but don't forget to swizzle R/B.
                m_Stream.ReadBytesToBarePointer VarPtr(prevPixel), 3
                tmpColor = prevPixel.Red
                prevPixel.Red = prevPixel.Blue
                prevPixel.Blue = tmpColor
                
            'Remaining options are 2-bit flags (2 MSBs) followed by 6-bit indicators.
            Else
                
                Dim high2b As Byte, low6b As Byte
                high2b = (nextByte And QOI_MASK_2)
                low6b = (nextByte And QOI_MASK_L6)
                
                'Many of these flags require us to calculate diffs between pixels
                Dim rDiff As Long, gDiff As Long, bDiff As Long
                Dim rNew As Long, gNew As Long, bNew As Long
                
                Select Case high2b
                    
                    'Low 6-bits are an index into the previous color array [0-63]
                    Case QOI_OP_INDEX
                        prevPixel = m_PrevPixels(low6b)
                        
                    'Low 6-bits are 2-bit differences (wraparound) between current RGB and previous RGB values,
                    ' so the range [0, 3] must be translated to the range [-2, 1]
                    Case QOI_OP_DIFF
                        rDiff = CLng((low6b \ 16) And &H3) - 2
                        gDiff = CLng((low6b \ 4) And &H3) - 2
                        bDiff = CLng(low6b And &H3) - 2
                        
                        'Calculate new values, with manual coverage for wraparound
                        rNew = CLng(prevPixel.Red) + rDiff
                        If (rNew < 0) Then rNew = rNew + 256 Else rNew = rNew And &HFF
                        gNew = CLng(prevPixel.Green) + gDiff
                        If (gNew < 0) Then gNew = gNew + 256 Else gNew = gNew And &HFF
                        bNew = CLng(prevPixel.Blue) + bDiff
                        If (bNew < 0) Then bNew = bNew + 256 Else bNew = bNew And &HFF
                        
                        prevPixel.Red = rNew
                        prevPixel.Green = gNew
                        prevPixel.Blue = bNew
                        
                    'Low 6-bits are a 6-bit difference (wraparound) between current G and previous G values,
                    ' (translated from [0, 63] to [-32, 31]) and another byte must be read for difference
                    ' in R, B values (4-bits translated from [0, 15] to [-8, 7] and added to G diff).
                    Case QOI_OP_LUMA
                        gDiff = CLng(low6b) - 32
                        nextByte = m_Stream.ReadByte()
                        rDiff = CLng(nextByte \ 16) - 8 + gDiff
                        bDiff = CLng(nextByte And &HF) - 8 + gDiff
                        
                        'Calculate new values, with manual coverage for wraparound
                        rNew = CLng(prevPixel.Red) + rDiff
                        If (rNew < 0) Then rNew = rNew + 256 Else rNew = rNew And &HFF
                        gNew = CLng(prevPixel.Green) + gDiff
                        If (gNew < 0) Then gNew = gNew + 256 Else gNew = gNew And &HFF
                        bNew = CLng(prevPixel.Blue) + bDiff
                        If (bNew < 0) Then bNew = bNew + 256 Else bNew = bNew And &HFF
                        
                        prevPixel.Red = rNew
                        prevPixel.Green = gNew
                        prevPixel.Blue = bNew
                        
                    Case QOI_OP_RUN
                        curRun = low6b
                    
                End Select
                
            End If
            
            'The contents of prevPixel will now reflect the color we just read.
            ' It must always be added to the recent color table.
            UpdateRecentColorsLUT prevPixel
            
        End If
        
        'Write the pixel value into the destination image
        dstPixels(pxOffset) = prevPixel
        
        'Increment pixel offset into the destination image and continue
        pxOffset = pxOffset + 1
        
    Loop While (pxOffset < numPixels)
    
    'Safely unwrap array wrapper and return
    dstDIB.UnwrapRGBQuadArrayFromDIB dstPixels
    LoadQOI_ParsePixels = True
    
End Function

Private Sub UpdateRecentColorsLUT(ByRef newColor As RGBQuad)
    
    'Generate a simple hash into the lookup table using this formula from the spec:
    ' index_position = (r * 3 + g * 5 + b * 7 + a * 11) % 64
    Dim indexPos As Long
    indexPos = CLng(newColor.Red) * 3 + CLng(newColor.Green) * 5 + CLng(newColor.Blue) * 7 + CLng(newColor.Alpha) * 11
    m_PrevPixels(indexPos And &H3F&) = newColor
    
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdQOI." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdQOI." & funcName & "() reported an error: " & errDescription
    End If
End Sub
