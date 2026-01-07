VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdWBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Wireless Application Protocol Bitmap (WBMP) Decoder
'Copyright 2024-2026 by Tanner Helland
'Created: 19/March/24
'Last updated: 19/March/24
'Last update: initial build, heavily derived from PD's just-written XBM parser
'
'The WBMP file format has been around since the late 90s.  You can find a copy of the latest spec revision
' here: https://www.wapforum.org/what/technical/SPEC-WAESpec-19990524.pdf
'
'PhotoDemon originally handed WBMP support off to the 3rd-party FreeImage library, but ongoing issues with
' FreeImage reliability eventually prompted me to write my own WBMP decoder.  (This was done shortly after
' writing a similar decoder for the XBM format, which is another ancient monochrome-only image format;
' this decoder is based roughly on the code for that one.)
'
'PD's WBMP decoder is small and lightweight, covering all known variants of the format.  Encoding support
' is *not* planned at present, but this may be revisited (as always) if users complain.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps additional diagnostic information
' to PD's primary debug log.
Private Const WBMP_DEBUG_VERBOSE As Boolean = True

'Image width/height, in pixels, after a file has been validated.
' (Calling "IsFileWBMP" is enough to populate these.)
Private m_Width As Long, m_Height As Long

'Byte-by-byte access is provided, as always, by a pdStream instance.
Private m_Stream As pdStream

'The last filename loaded.  The contents of the filename will be stored in m_FileAsString, below.
Private m_OpenFilename As String

'WBMPs use VLQs ("variable-length quantities"), which are 128-bits of data with a 1-bit flag indicating
' "more data follows" if set to 1.  To simplify the process of encoding these in VB6, I use a fixed
' 5-byte array (which can store up to 4 individual bytes of actual data encoded as a VLQ).
Private m_VLQ(0 To 4) As Byte

'WBMP files use a very small header with no magic numbers, so we check a few key fields (including width/height)
' when attempting to validate.
Friend Function IsFileWBMP(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True, Optional ByVal closeStreamWhenDone As Boolean = True) As Boolean
        
    Const FUNC_NAME As String = "IsFileWBMP"
        
    IsFileWBMP = False
    On Error GoTo BadWBMPFile
    
    m_OpenFilename = vbNullString
    m_Width = 0: m_Height = 0
    
    Dim potentialMatch As Boolean
    potentialMatch = Files.FileExists(srcFilename)
    
    'Check extension too, as requested.
    If (potentialMatch And requireValidFileExtension) Then
        potentialMatch = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "wbmp", True) Or Strings.StringsEqual(Files.FileGetExtension(srcFilename), "wbm", True)
    End If
    
    'If any of the failsafe checks failed, exit immediately
    If (Not potentialMatch) Then
        Exit Function
    ElseIf WBMP_DEBUG_VERBOSE Then
        PDDebug.LogAction "File might be WBMP.  Opening image and attempting to validate..."
    End If
    
    'WBMP files use an extremely simple header comprised of a series of VLQs
    ' (Wiki description of VLQs here: https://en.wikipedia.org/wiki/Variable-length_quantity)
    '
    'We want to read enough of the header to validate width/height.  (It's only a few bytes!)
    
    'Open a stream on the target file
    If (m_Stream Is Nothing) Then Set m_Stream = New pdStream
    If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename, optimizeAccess:=OptimizeSequentialAccess) Then
        
        m_OpenFilename = srcFilename
        
        'First byte is technically a VLQ, but per the spec, the only valid definition (from the working group)
        ' is 0.
        If (m_Stream.ReadByte() <> 0) Then Exit Function
        
        'This is followed by a fixed header field and a variable number of optional extension header fields...
        ' ...but again, the only valid value the fixed header is 0.
        If (m_Stream.ReadByte() <> 0) Then Exit Function
        
        'Finally we arrive at width and height, which are both encoded as VLQs
        m_Width = ReadVLQ(m_Stream)
        m_Height = ReadVLQ(m_Stream)
        
    '/Couldn't start stream; file may be locked or inaccessible
    Else
        GoTo BadWBMPFile
    End If
    
    'After all that, ensure we have valid, non-zero width/height
    If (m_Width > 0) And (m_Height > 0) Then
        IsFileWBMP = True
        If WBMP_DEBUG_VERBOSE Then PDDebug.LogAction "File validates as WBMP!"
    End If
    
    'Close the file stream before exiting
    If (Not m_Stream Is Nothing) And closeStreamWhenDone Then m_Stream.StopStream True
    
    Exit Function
    
'On any parse error, this function jumps to this branch and simply closes the underlying file, then exits
BadWBMPFile:
    
    Set m_Stream = Nothing
    InternalError FUNC_NAME, "critical parse failure"
    IsFileWBMP = False
    
End Function

'Validate and load a candidate WBMP file
Friend Function LoadWBMP_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    On Error GoTo BadWBMPFile
    
    Const FUNC_NAME As String = "LoadWBMP_FromFile"
    LoadWBMP_FromFile = False
    
    'Validate the file and leave it open if validated successfully
    If Me.IsFileWBMP(srcFile, False, False) Then
        
        'If validation passed, the width and height (if any) will be stored in m_width and m_height.
        If (m_Width <= 0) Or (m_Height <= 0) Then
            InternalError FUNC_NAME, "bad dimensions: " & m_Width & "x" & m_Height
            Set m_Stream = Nothing
            Exit Function
        End If
        
        If WBMP_DEBUG_VERBOSE Then
            PDDebug.LogAction "WBMP dimensions: " & m_Width & "x" & m_Height
            PDDebug.LogAction "Starting to parse pixel data..."
        End If
        
        'We left the stream open after validation, so the stream pointer already points at pixel data!
        
        'Prep image buffer; we'll dump intensity values straight into it.
        Set dstDIB = New pdDIB
        If dstDIB.CreateBlank(m_Width, m_Height, 32, vbWhite, 255) Then
            
            'There's not a performance-friendly way to mask flags in VB, so let's just use a byte array for clarity
            Dim bitFlags(0 To 7) As Byte
            bitFlags(0) = 1
            bitFlags(1) = 2
            bitFlags(2) = 4
            bitFlags(3) = 8
            bitFlags(4) = 16
            bitFlags(5) = 32
            bitFlags(6) = 64
            bitFlags(7) = 128
            
            'To avoid having to pull pixels out one-by-one (which is slow),
            ' grab the whole source data chunk as a local byte array.
            Dim srcPixels() As Byte, idxSrcByte As Long
            m_Stream.ReadBytes srcPixels, -1, True
            m_Stream.StopStream True
            
            'We know how many pixels we should be addressing (based on the underlying width/height)
            Dim totalNumPixels As Long, numPixelsScanline As Long, curByte As Byte
            totalNumPixels = m_Width * m_Height
            
            'Scanlines must be null-padded to end on byte boundaries.
            Dim maxPixelsScanline As Long
            maxPixelsScanline = Int((m_Width + 7) \ 8) * 8
            
            'Valiate the width/height against the size of the pixel data
            If (((maxPixelsScanline \ 8) * m_Height) > UBound(srcPixels)) Then
                InternalError FUNC_NAME, "not enough pixel data", True
                Exit Function
            End If
            
            'We'll need to track an index into the source data; it'll be updated "as we go"
            idxSrcByte = 0
            
            'Wrap an array around the destination DIB.  (This is unsafe, and must be manually
            ' freed before this function xits.)
            Dim imgPixels() As Long, imgSA As SafeArray1D
            dstDIB.WrapLongArrayAroundDIB_1D imgPixels, imgSA
            
            'Start iterating bytes and converting bits to monochrome colors!
            Dim x As Long
            Do While (x < totalNumPixels)
                
                'We don't need to validate these reads, as we already asserted source size in a previous step
                curByte = srcPixels(idxSrcByte)
                
                'Parse each bit in turn
                Dim i As Long
                For i = 7 To 0 Step -1
                    
                    'Ignore empty bytes at the end of the image
                    If (x < totalNumPixels) Then
                        
                        'Ignore empty bytes at the end of each scanline
                        If (numPixelsScanline < m_Width) Then
                            
                            'Draw black bits into the image
                            If ((curByte And bitFlags(i)) = 0) Then imgPixels(x) = &HFF000000
                            
                            'Increment pixel pointer
                            x = x + 1
                            
                        End If
                        
                    End If
                    
                    'Keep track of how many pixels we've copied *on this scanline*
                    numPixelsScanline = numPixelsScanline + 1
                    
                Next i
                
                'If we've reached the end of a scanline, reset the scanline pixel counter
                If (numPixelsScanline >= maxPixelsScanline) Then numPixelsScanline = 0
                
                'Advance to the next byte
                idxSrcByte = idxSrcByte + 1
                
            Loop
            
            'Release our unsafe array wrapper
            dstDIB.UnwrapLongArrayFromDIB imgPixels
            
            'File was loaded successfully!
            LoadWBMP_FromFile = True
            
            'The returned data is always premultiplied
            If LoadWBMP_FromFile Then dstDIB.SetInitialAlphaPremultiplicationState True
            
        Else
            InternalError FUNC_NAME, "out of memory"
            Set m_Stream = Nothing
            Exit Function
        End If
        
    '/File is not WBMP; silently ignore it
    End If
    
    Exit Function
    
BadWBMPFile:
    InternalError FUNC_NAME, "abandoned load due to critical error"
    LoadWBMP_FromFile = False
    
End Function

'Decode a "variable-length quantity" entry inside the target stream, starting at the current stream location.
' VLQs are a way to encode 128-bit entries in a variable-length stream, by using the first bit of each byte
' to indicate that additional bytes follow.  You then keep reading bytes one-at-a-time, shifting each one
' 7-bits left, until you arrive at a byte with a 0 in the first bit space.
Private Function ReadVLQ(ByRef srcStream As pdStream) As Long
    
    Const FUNC_NAME As String = "ReadVLQ"
    
    On Error GoTo BadVLQ
    
    Dim numBytesPulled As Long
    ReadVLQ = 0: numBytesPulled = 0
    
    'Pull the next byte in, but stop if...
    ' 1) the end of stream is reached (otherwise we could loop eternally!)
    ' 2) too many bytes have been read
    Do While (srcStream.GetPosition < srcStream.GetStreamSize()) And (numBytesPulled < 5)
    
        Dim nextByte As Byte
        nextByte = srcStream.ReadByte()
        
        'Add this byte to the running VLQ entry (but make sure to strip the flag bit)
        ReadVLQ = ReadVLQ Or (nextByte And &H7F)
    
        'Check the flag bit and stop looping if this is the last entry
        If ((nextByte And &H80) = &H0) Then Exit Do
        
        'Shift existing numerical data by 2^7 (remember - we're reading 128 bits at a time)
        ReadVLQ = ReadVLQ * 128
        
        'Track how many bytes we've pulled out, because this function can only hold up to 4-bytes at a time
        numBytesPulled = numBytesPulled + 1
        
    Loop
    
    'Auto-return whatever we've calculated so far
    Exit Function
    
BadVLQ:
    ReadVLQ = 0
    InternalError FUNC_NAME, "bad VLQ entry"
    
End Function

'Encode a "variable-length quantity" entry inside the target stream, starting at the current stream location.
' VLQs are a way to encode 128-bit entries in a variable-length stream, by using the first bit of each byte
' to indicate that additional bytes follow.  (A 0 in the MSB position indicates no additional bits follow.)
Private Sub WriteVLQ(ByRef dstStream As pdStream, ByVal valToWrite As Long)
    
    Dim numBytesEncoded As Long, tmpEncode As Byte
    numBytesEncoded = 0
    
    'It's easier to calculate VLQs from right to left, then "push" the encoded
    ' bytes out in the reverse order we calculated them.
    Do
        
        'Add the least-significant 7-bits to the running list
        m_VLQ(numBytesEncoded) = (valToWrite And &H7F&)
        numBytesEncoded = numBytesEncoded + 1
        
        'Shift the original value 7-bits to the left
        valToWrite = valToWrite \ 128
        
    Loop While (valToWrite > 0)
    
    'Push the finished byte stream out to file (in *reverse* order), and add a preceding 1-bit to all bits
    ' but the final one, to indicate that additional bits follow.
    Dim i As Long
    For i = numBytesEncoded - 1 To 0 Step -1
        tmpEncode = m_VLQ(i)
        If (i > 0) Then tmpEncode = tmpEncode Or &H80&
        dstStream.WriteByte tmpEncode
    Next i
    
End Sub

'Save an arbitrary DIB to a standalone WBMP file.  An optional parent pdImage object can also be passed,
' but it is not currently required (or even used).
Friend Function SaveWBMP_ToFile(ByRef dstFile As String, ByRef srcImage As pdImage, ByRef srcDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "SaveWBMP_ToFile"
    SaveWBMP_ToFile = False
    
    On Error GoTo SaveFailed
    
    'WBMP files are very easy to save.  There are no user-controlled parameters, and saving is deterministic.
    
    'WBMP files are always saved as 1-bit monochrome, so the only thing we need to embed in the file
    ' (besides pixel contents, obviously) are the image's dimensions!
    
    'Start by opening a pdStream instance on the target file, then writing a fixed-size header.
    Set m_Stream = New pdStream
    If (Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFile, optimizeAccess:=OptimizeSequentialAccess)) Then
        InternalError FUNC_NAME, "couldn't initialize stream on target file"
        SaveWBMP_ToFile = False
        Exit Function
    End If
    
    'Set up progress bar reports
    ProgressBars.SetProgBarMax srcDIB.GetDIBHeight
    
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Headers for WBMP files are fixed, and only one type of WBMP image is defined by the spec.  Easy-peasy!
    m_Stream.WriteByte 0
    m_Stream.WriteByte 0
    WriteVLQ m_Stream, srcDIB.GetDIBWidth
    WriteVLQ m_Stream, srcDIB.GetDIBHeight
    
    'We expect the caller to already have retrieved the composited image and composited it against
    ' a background (into srcDIB).  Now all we have to do is retrieve each pixel and write it out as
    ' a monochrome bit stream.
    
    'Convert the source image into a grayscale byte array
    Dim imgGray() As Byte
    DIBs.GetDIBGrayscaleMap srcDIB, imgGray, False
    
    'Use dithering to further convert that byte array to a list of pure 0 or 255 values
    Filters_ByteArray.ThresholdByteArray imgGray, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 127
    
    'There's not a performance-friendly way to mask flags in VB, so let's just use a byte array for clarity
    Dim bitFlags(0 To 7) As Byte
    bitFlags(0) = 128
    bitFlags(1) = 64
    bitFlags(2) = 32
    bitFlags(3) = 16
    bitFlags(4) = 8
    bitFlags(5) = 4
    bitFlags(6) = 2
    bitFlags(7) = 1
    
    'Iterate scanlines (encoded pixels must null-pad bits beyond the current scanline) and dump
    ' encoded bytes to file as we go!
    Dim x As Long, y As Long, i As Long, encodedByte As Byte
    
    Dim xLimit As Long
    xLimit = (srcDIB.GetDIBWidth - 1)
    
    'It's slow to write individual bytes out to stream (via function), so instead, cache each
    ' scanline into an array and dump them out in groups.
    Dim xEncoded() As Byte, maxBytesScanline As Long, idxDst As Long
    maxBytesScanline = Int((srcDIB.GetDIBWidth + 7) \ 8)
    ReDim xEncoded(0 To maxBytesScanline - 1) As Byte
    
    For y = 0 To srcDIB.GetDIBHeight - 1
        
        'Reset the x-trackers
        x = 0
        idxDst = 0
        
        Do
            
            encodedByte = 0
            
            'Copy the next 8 pixels into a single byte
            For i = 0 To 7
                
                'Because the image's converted grayscale values are guaranteed to be 0 or 255,
                ' we can skip an expensive conditional (note the commented out code, which I've
                ' left for clarity) and simply && the gray value as-is
                'If (imgGray(x, y) > 0) Then encodedByte = encodedByte Or bitFlags(i)
                encodedByte = encodedByte Or (bitFlags(i) And imgGray(x, y))
                x = x + 1
                If (x > xLimit) Then Exit For
                
            Next i
            
            'Cache this byte locally.
            'm_Stream.WriteByte encodedByte
            xEncoded(idxDst) = encodedByte
            idxDst = idxDst + 1
            
        Loop While (x < xLimit)
        
        'Dump the encoded array out to file
        m_Stream.WriteBytesFromPointer VarPtr(xEncoded(0)), maxBytesScanline
        
        'Relay progress as we go
        If (y And progBarCheck) = 0 Then ProgressBars.SetProgBarVal y
        
    Next y
    
    'Finished!
    m_Stream.StopStream
    ProgressBars.ReleaseProgressBar
    
    SaveWBMP_ToFile = True
    Exit Function
    
SaveFailed:
    SaveWBMP_ToFile = False
    InternalError FUNC_NAME, "internal VB error #" & Err.Number & ": " & Err.Description
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdWBMP." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdWBMP." & funcName & "() reported an error: " & errDescription
    End If
End Sub

'The underlying stream would auto-free naturally, but I like being tidy
Private Sub Class_Terminate()
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen Then m_Stream.StopStream True
    End If
End Sub
