VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPNGChunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PNG Chunk Manager (companion to the pdPNG class)
'Copyright 2018-2018 by Tanner Helland
'Created: 12/April/18
'Last updated: 12/April/18
'Last update: initial build
'
'pdPNG leans on this class to handle the actual storage (and deciphering) of individual PNG chunks.
' This simplifies the decode process because we only deal with the chunks that matter to PNG, and it
' gives us an convenient way to concatenate IDAT chunks post-load.  (PNG allows image data to be split
' over multiple IDAT chunks.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Chunk-specific enums
Private Enum PD_PNGFilter
    png_None = 0
    png_Sub = 1
    png_Up = 2
    png_Average = 3
    png_Paeth = 4
End Enum

#If False Then
    Private Const png_None = 0, png_Sub = 1, png_Up = 2, png_Average = 3, png_Paeth = 4
#End If

'Core PNG data
Private m_Type As String
Private m_Data As pdStream, m_Size As Long
Private m_CRCRuntime As Long, m_CRCEmbedded As Long

'Internal PD flag; we use this for different things at different times, so it's meaning varies
Private m_Flagged As Boolean

Friend Function GetType() As String
    GetType = m_Type
End Function

Friend Function GetFlag() As Boolean
    GetFlag = m_Flagged
End Function

Friend Sub SetFlag(ByVal flagState As Boolean)
    m_Flagged = flagState
End Sub

'I deliberately call this function "BorrowData" instead of "GetData" as a reminder that the stream is
' *not* duplicated - it is the same stream, so any position manipulations need to be considered, because this
' object won't reset them after-the-fact!
Friend Function BorrowData() As pdStream
    Set BorrowData = m_Data
End Function

Friend Function GetDataSize() As Long
    GetDataSize = m_Size
End Function

'To create a new chunk, notify this class of the chunk type and length.  Note that you don't actually load
' the chunk data in this step - that happens separately.
Friend Sub CreateChunk(ByRef chunkType As String, ByVal chunkSize As Long)
    
    m_Type = chunkType
    m_Size = chunkSize
    
    'Even if the chunk size is zero, we still want to start a stream object (in case we merge additional
    ' chunks into this one).
    Set m_Data = New pdStream
    m_Data.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, , m_Size
    
End Sub

'After creating the chunk, load its data via this step.  IMPORTANTLY, the source stream needs to
' *already* be pointing at the correct offset in the PNG data!  ALSO IMPORTANTLY, *do not call this
' function if the source data is length zero*.  We treat that as an error condition.
Friend Function LoadChunkData(ByRef srcStream As pdStream) As Boolean

    If (m_Size > 0) Then
        LoadChunkData = (m_Size = srcStream.ReadBytesToBarePointer(m_Data.Peek_PointerOnly(0), m_Size))
        If LoadChunkData Then
            m_Data.SetSizeExternally m_Size
            LoadChunkData = m_Data.SetPosition(m_Size, FILE_BEGIN)
        End If
    Else
        LoadChunkData = False
    End If

End Function

Friend Sub NotifyCRCs(ByVal runtimeCRC As Long, ByVal embeddedCRC As Long)
    m_CRCRuntime = runtimeCRC
    m_CRCEmbedded = embeddedCRC
End Sub

'Pixel data can be spread across multiple IDAT chunks.  To simplify the load process, we merge these
' separate instances into a single contiguous stream.  (As part of this process, the source chunk will
' be automatically freed.)
Friend Sub MergeOtherChunk(ByRef srcChunk As pdPNGChunk)
    
    If (Not srcChunk Is Nothing) Then
    
        'Merge the stream data into our stream
        Dim tmpStream As pdStream
        Set tmpStream = srcChunk.BorrowData()
        m_Size = m_Size + srcChunk.GetDataSize()
        m_Data.WriteBytesFromPointer tmpStream.Peek_PointerOnly(0), srcChunk.GetDataSize()
        
        'Free the source chunk
        Set tmpStream = Nothing
        Set srcChunk = Nothing
        
    End If
    
End Sub

'If this chunk contains compressed data, call this function to decompress it.  The decompressed results will
' automatically overwrite the original, compressed chunks by design, and our internal "size" tracker will be
' updated to match.
'
'For IDAT chunks specifically, you *must* pass a valid inflateSize parameter.  This tells us how large our
' decompression buffer must be (since zLib streams don't independently store that data).  Other compressed
' chunk types do not require that information; they need to use more specialized techniques for determining
' decompression size.
'
'RETURNS: TRUE if decompression was successful, or if the chunk wasn't compressed in the first place.
'         FALSE means zLib reported an error - you'll need to query it for additional details.
Friend Function DecompressChunk(ByRef warningStack As pdStringStack, Optional ByVal inflateSize As Long = 0) As Boolean
    
    DecompressChunk = True
    Dim tmpBuffer As pdStream
    
    If (m_Type = "IDAT") Then
        
        'Failsafe checks
        If (inflateSize = 0) Then
            warningStack.AddString "pdPNGChunk.DecompressChunk was passed a null inflate size - decompression abandoned!"
            DecompressChunk = False
            Exit Function
        End If
        
        'Create a new buffer at the required size
        Set tmpBuffer = New pdStream
        tmpBuffer.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, , inflateSize
        
        'Use zLib to perform a direction decompression, and validate that our decompressed size is
        ' exactly what we anticipated.
        Dim initSize As Long
        initSize = inflateSize
        DecompressChunk = Plugin_zLib.ZlibDecompress_UnsafePtr(tmpBuffer.Peek_PointerOnly(0), inflateSize, m_Data.Peek_PointerOnly(0), m_Size)
        
        If DecompressChunk Then
            DecompressChunk = (initSize = inflateSize)
            If (Not DecompressChunk) Then warningStack.AddString "pdPNGChunk.DecompressChunk received mismatched sizes from zLib: " & initSize & ", " & inflateSize
        Else
            warningStack.AddString "pdPNGChunk.DecompressChunk received a failure state from zLib; FYI sizes were init: " & initSize & ", " & inflateSize
        End If
        
        'If decompression worked, swap streams and update our internal size tracker
        If DecompressChunk Then
            Set m_Data = tmpBuffer
            m_Data.SetSizeExternally inflateSize
            m_Size = inflateSize
        End If
    
    'ICC profiles are also compressed, but they're more finicky to decompress
    ElseIf (m_Type = "iCCP") Then
    
        'Here we encounter yet another asinine implementation decision.  The iCCP chunk contains an ICC profile.
        ' The chunk layout is as follows:
        ' Profile name          1-79 bytes (character string)
        ' Null separator        1 byte (null character)
        ' Compression method    1 byte
        ' Compressed profile    (n) bytes
        
        'You'll notice that nowhere in this layout is the original, uncompressed size of the ICC profile.
        ' Astute readers will know that zLib streams - by design - do not store the original size of the stream;
        ' that's up to the user.  So there's an obvious problem: how the fuck do we size our inflate buffer?
        ' Do the PNG authors really expect us to start with an arbitrary buffer size and increase it until zLib
        ' is satisfied?  Yes, they actually do, because they do the exact same thing with iTXT and zTXT chunks.
        ' (With decision-making skills like this, no wonder libPNG is a constant source of overflow vulnerabilities.)
        
        'Instead of that shitty approach, we try to do something smarter.  ICC profile headers actually contain
        ' their own length as one of the members.  (The first member in their header, in fact!)  Because ICC headers
        ' are 128-bytes, we can perform a partial decompression, then use the ICC header's data to compute a valid
        ' size for the *full* ICC extraction.
        
        'First things first, however: we need to find out where the hell the profile actually lies, which means
        ' parsing the stream looking for the first null-byte.
        Dim nullLoc As Long
        m_Data.SetPosition 0, FILE_BEGIN
        nullLoc = m_Data.FindByte(0, , False)
        
        'Hypothetically, there *must* be a name, so we really shouldn't continue if a null byte is found at
        ' position 0.  Similarly, the spec imposes an arbitrary 79-character limit on profile names.
        ' (We could also check the resulting string for valid chars, as required by the spec, but we really
        ' don't care about that at present.)
        If (nullLoc > 0) And (nullLoc < 79) Then
        
            'We don't currently cache the ICC profile name (maybe that would be useful in the future? IDK),
            ' but let's at least retrieve it to make sure our implementation works.  This will also move the
            ' file pointer into the correct position.
            Dim profName As String
            profName = m_Data.ReadString_ASCII(nullLoc)
            Debug.Print "Found ICC profile: (" & profName & ")"
            
            'The stream now points at the null-terminator of the name string.  Advance it one byte to point at
            ' the "compression method" byte, then validate that byte (it must be 0).
            m_Data.SetPosition 1, FILE_CURRENT
            If (m_Data.ReadByte() = 0) Then
                
                'We now need to grab the ICC header and parse it for a profile length
                Set tmpBuffer = New pdStream
                tmpBuffer.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, , 128
                
                Dim curOffset As Long
                curOffset = m_Data.GetPosition()
                
                'ZLib is going to fail - we know that - but it's okay.  We know 128 bytes isn't enough to extract the
                ' entire profile.  (We do, however, need to fail specifically because the buffer is too small, and not
                ' because of some other problem state.)
                Dim writeSize As Long: writeSize = 128
                Const ZLIB_BUFF_TOO_SMALL As Long = -5
                
                Dim zLibResult As Long
                zLibResult = Plugin_zLib.ZlibDecompress_UnsafePtrEx(tmpBuffer.Peek_PointerOnly(0), writeSize, m_Data.Peek_PointerOnly(curOffset), m_Size - curOffset)
                If (zLibResult = ZLIB_BUFF_TOO_SMALL) Then
                    
                    'The ICC header has now been dumped into our temporary stream.  Retrieve the profile header size.
                    tmpBuffer.SetSizeExternally 128
                    tmpBuffer.SetPosition 0, FILE_BEGIN
                    
                    Dim profSize As Long
                    profSize = tmpBuffer.ReadLong_BE()
                    
                    'We can now reset our temporary buffer to that size, and extract the *entire* ICC profile
                    tmpBuffer.StopStream True
                    tmpBuffer.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, , profSize
                    DecompressChunk = Plugin_zLib.ZlibDecompress_UnsafePtr(tmpBuffer.Peek_PointerOnly(0), profSize, m_Data.Peek_PointerOnly(curOffset), m_Size - curOffset)
                    
                    'Compression was successful!  Overwrite our stream with the contents of the ICC profile, then exit.
                    If DecompressChunk Then
                        m_Data.StopStream True
                        Set m_Data = tmpBuffer
                        m_Size = profSize
                    Else
                        warningStack.AddString "pdPNGChunk.DecompressChunk couldn't decompress the ICC chunk; ICC retrieval was *not* successful."
                    End If
                    
                Else
                    warningStack.AddString "pdPNGChunk.DecompressChunk couldn't extract a valid ICC header (" & CStr(zLibResult) & "); ICC retrieval was *not* successful."
                End If
                
            Else
                warningStack.AddString "pdPNGChunk.DecompressChunk found an iCCP chunk with unknown compression; ICC retrieval was *not* successful."
                DecompressChunk = False
            End If
            
        Else
            warningStack.AddString "pdPNGChunk.DecompressChunk found a malformed iCCP chunk; ICC retrieval was *not* successful."
            DecompressChunk = False
        End If
        
    'Anything other than IDAT and iCCP is still TODO!
    Else
        DecompressChunk = True
    End If

End Function

'If this chunk contains filtered IDAT data, call this function to un-filter it.  This will transform the PNG IDAT stream
' (which contains things other than pixel data, like PNG filter indicators) into a raw stream of unadulterated pixel data.
'
'RETURNS: TRUE if unfiltering was successful.  FALSE is a critical error; you must suspend processing after receiving it.
Friend Function UnfilterChunk(ByRef warningStack As pdStringStack, ByRef srcHeader As PD_PNGHeader, ByRef xPixelCount() As Long, ByRef yPixelCount() As Long, ByRef xByteCount() As Long) As Boolean
    
    UnfilterChunk = True
    
    'Un-filtering only applies to IDAT chunks
    If (m_Type <> "IDAT") Then
        UnfilterChunk = True
        Exit Function
    End If
    
    'Because member access is slower in VB, transfer some values required in the inner-loop to standalone variables
    Dim imgWidth As Long, imgHeight As Long, imgIsInterlaced As Boolean, imgColorType As PD_PNGColorType, imgBitDepth As Long
    imgWidth = srcHeader.Width
    imgHeight = srcHeader.Height
    imgBitDepth = srcHeader.BitDepth
    imgColorType = srcHeader.ColorType
    imgIsInterlaced = srcHeader.Interlaced
    
    Dim xFinal As Long, yFinal As Long
    xFinal = imgWidth - 1
    yFinal = imgHeight - 1
    
    'Multi-channel images (or images with multi-byte channel values) require offsets during the path phase.
    Dim pxOffset As Long
    If (imgColorType = png_Greyscale) Then
        pxOffset = imgBitDepth \ 8
    ElseIf (imgColorType = png_GreyscaleAlpha) Then
        pxOffset = (imgBitDepth * 2) \ 8
    ElseIf (imgColorType = png_Truecolor) Then
        pxOffset = (imgBitDepth * 3) \ 8
    ElseIf (imgColorType = png_TruecolorAlpha) Then
        pxOffset = (imgBitDepth * 4) \ 8
    End If
    If (pxOffset < 1) Then pxOffset = 1
    
    Debug.Print imgBitDepth, imgColorType, pxOffset
    
    'A note throughout this function: per the PNG spec:
    ' - "Unsigned arithmetic modulo 256 is used, so that both the inputs and outputs fit into bytes."
    ' - "Filters are applied to each byte regardless of bit depth."
    
    'As usual, interlaced images are their own beasts.  For performance reasons, we split handling for interlaced
    ' images into their own function (as they must be processed in a more time-consuming way).
    If imgIsInterlaced Then
    
    Else
    
        'To conserve memory, we're going to do an "in-place" unfiltering of the IDAT stream.  The most efficient way
        ' to do this is to wrap a fake array around the stream's data (rather than copying it around).
        Dim srcBytes() As Byte, tmpSA As SafeArray1D
        m_Data.WrapArrayAroundMemoryStream srcBytes, tmpSA, 0
        
        'Each scanline is prefaced by a "filter type" that tells us what kind of filter was used on this line.
        ' We read that preceding byte to know how to filter the bytes in question.
        Dim xBound As Long, yBound As Long, xStride As Long
        xBound = xPixelCount(0)
        yBound = yPixelCount(0)
        xStride = xByteCount(0)
        
        Dim curValue As Long, prevValue As Long, prevValue2 As Long, prevValue3 As Long
        
        Dim x As Long, y As Long, i As Long, yOffset As Long
        For y = 0 To yFinal
        
            yOffset = y * xStride
            
            'Filtering changes based on the "filter byte" at this position
            Select Case srcBytes(yOffset)
                
                'Regardless of image properties, unfiltered bytes don't require additional processing
                Case png_None
                
                '"Sub" filtering filters according to the previous pixel
                Case png_Sub
                
                    yOffset = yOffset + 1
                    For x = 0 To xStride - 2
                        If (x >= pxOffset) Then prevValue = srcBytes(yOffset + x - pxOffset) Else prevValue = 0
                        curValue = srcBytes(yOffset + x)
                        srcBytes(yOffset + x) = (curValue + prevValue) And 255
                    Next x
                    
                '"Up" filtering filters according to the pixel above
                Case png_Up
                
                    yOffset = yOffset + 1
                    For x = 0 To xStride - 2
                        If (y > 0) Then prevValue = srcBytes(yOffset + x - xStride) Else prevValue = 0
                        curValue = srcBytes(yOffset + x)
                        srcBytes(yOffset + x) = (curValue + prevValue) And 255
                    Next x
                    
                '"Average" filtering filters according to the average between the previous pixel and the pixel above
                Case png_Average
                
                    yOffset = yOffset + 1
                    For x = 0 To xStride - 2
                        If (x >= pxOffset) Then prevValue = srcBytes(yOffset + x - pxOffset) Else prevValue = 0
                        If (y > 0) Then prevValue2 = srcBytes(yOffset + x - xStride) Else prevValue2 = 0
                        curValue = srcBytes(yOffset + x)
                        srcBytes(yOffset + x) = (curValue + ((prevValue + prevValue2) \ 2)) And 255
                    Next x
                
                '"Paeth" uses an external formula
                Case png_Paeth
                    yOffset = yOffset + 1
                    For x = 0 To xStride - 2
                        If (x >= pxOffset) Then prevValue = srcBytes(yOffset + x - pxOffset) Else prevValue = 0
                        If (y > 0) Then
                            prevValue2 = srcBytes(yOffset + x - xStride)
                            If (x >= pxOffset) Then prevValue3 = srcBytes(yOffset + x - xStride - pxOffset) Else prevValue3 = 0
                        Else
                            prevValue2 = 0
                            prevValue3 = 0
                        End If
                        curValue = srcBytes(yOffset + x)
                        srcBytes(yOffset + x) = (curValue + PaethPredictor(prevValue, prevValue2, prevValue3)) And 255
                    Next x
            
            End Select
            
        Next y
        
        m_Data.UnwrapArrayFromMemoryStream srcBytes
        
    End If
    
End Function

'Paeth filtering requires a whole stupid formula
Private Function PaethPredictor(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    
    Dim p As Long, pa As Long, pb As Long, pc As Long
    
    p = a + b - c
    pa = Abs(p - a)
    pb = Abs(p - b)
    pc = Abs(p - c)
    
    If (pa <= pb) And (pa <= pc) Then
        PaethPredictor = a
    Else
        If (pb <= pc) Then PaethPredictor = b Else PaethPredictor = c
    End If
    
End Function

Friend Function ConstructImage(ByRef dstDIB As pdDIB, ByRef srcHeader As PD_PNGHeader, ByRef srcPalette() As RGBQuad, ByVal srcPaletteSize As Long, ByVal srcTrnsRed As Long, ByVal srcTrnsGreen As Long, ByVal srcTrnsBlue As Long, ByRef xPixelCount() As Long, ByRef yPixelCount() As Long, ByRef xByteCount() As Long) As PD_PNGResult
    
    ConstructImage = True
    
    'Image construction only applies to IDAT chunks
    If (m_Type <> "IDAT") Then
        ConstructImage = True
        Exit Function
    End If
    
    'As usual, interlaced and non-interlaced present different challenges
    If srcHeader.Interlaced Then
    
        'Interlaced images are a pain in the ass; we'll deal with them later
    
    Else
        
        'Prep all the required loop values
        Dim imgWidth As Long, imgHeight As Long, xFinal As Long, yFinal As Long
        Dim x As Long, y As Long, i As Long, xStride As Long, xDIB As Long, srcOffset As Long, dstOffset As Long
        imgWidth = srcHeader.Width
        imgHeight = srcHeader.Height
        xFinal = imgWidth - 1
        yFinal = imgHeight - 1
        xStride = xByteCount(0)
        
        'To simplify our code, we now want to upsample any color-depths below 8-bpp.  (They present a unique challenge
        ' within VB, since we lack the bit-shift operators necessary to handle them elegantly.)  Note that the upsampler
        ' will reset our xStride value to match 8-bpp data it produces.
        If (srcHeader.BitDepth < 8) Then UpsampleLowBitDepthData srcHeader, xStride
        
        'If the image is 8-bpp grayscale, we now want to construct an artifical palette.  (This lets us process it
        ' in an identical manner to indexed color images.)
        If (srcHeader.ColorType = png_Greyscale) Then ConstructGrayPalette srcHeader, srcPalette, srcPaletteSize, srcTrnsRed
        
        'Certain bit-depths require us to handle transparency manually (but note that 8-bpp grayscale has already
        ' been covered by the ConstructGrayPalette() function, above!)
        Dim trnsHandlingRequired As Boolean, trnsCheck As Long
        trnsHandlingRequired = (srcHeader.ColorType = png_Truecolor) And (srcTrnsRed >= 0) And (srcTrnsGreen >= 0) And (srcTrnsGreen >= 0)
        If trnsHandlingRequired Then
            If (srcHeader.BitDepth = 8) Then trnsCheck = RGB(srcTrnsRed, srcTrnsGreen, srcTrnsBlue)
        End If
        
        'We can now assign actual pixel bits!
        Dim dstBytes() As Byte, dstSA As SafeArray2D
        dstDIB.WrapArrayAroundDIB dstBytes, dstSA
        
        'To improve performance, we use unsafe direct access to the chunk's contents
        Dim srcBytes() As Byte, srcSA As SafeArray1D
        m_Data.WrapArrayAroundMemoryStream srcBytes, srcSA, 0
        
        '16-bit-depths can be grayscale or truecolor, with or without an alpha channel
        If (srcHeader.BitDepth > 8) Then
        
        '8-bit (or less) images are handled universally, because we upsampled lower depths in a previous step
        Else
            
            Dim r As Long, g As Long, b As Long, a As Long
            
            Select Case srcHeader.ColorType
            
                Case png_Greyscale, png_Indexed
                    
                    Dim curIndex As Long
                    
                    For y = 0 To yFinal
                        srcOffset = y * xStride + 1
                    For x = 0 To xFinal
                        
                        curIndex = srcBytes(srcOffset + x)
                        
                        With srcPalette(curIndex)
                            xDIB = x * 4
                            dstBytes(xDIB, y) = .Blue
                            dstBytes(xDIB + 1, y) = .Green
                            dstBytes(xDIB + 2, y) = .Red
                            dstBytes(xDIB + 3, y) = .Alpha
                        End With
                        
                    Next x
                    Next y
                    
                    ConstructImage = png_Success
                
                Case png_Truecolor
                
                    For y = 0 To yFinal
                        srcOffset = y * xStride + 1
                    For x = 0 To xFinal
                        
                        'Byte order has to be swapped to BGRA
                        r = srcBytes(srcOffset + x * 3)
                        g = srcBytes(srcOffset + x * 3 + 1)
                        b = srcBytes(srcOffset + x * 3 + 2)
                        xDIB = x * 4
                        dstBytes(xDIB, y) = b
                        dstBytes(xDIB + 1, y) = g
                        dstBytes(xDIB + 2, y) = r
                        If trnsHandlingRequired Then
                            If (RGB(r, g, b) = trnsCheck) Then dstBytes(xDIB + 3, y) = 0
                        End If
                        
                    Next x
                    Next y
                    
                    ConstructImage = png_Success
                
                Case png_TruecolorAlpha
                
                    For y = 0 To yFinal
                        srcOffset = y * xStride + 1
                    For x = 0 To xFinal
                        
                        'Byte order has to be swapped to BGRA
                        xDIB = x * 4
                        dstBytes(xDIB, y) = srcBytes(srcOffset + x * 4 + 2)
                        dstBytes(xDIB + 1, y) = srcBytes(srcOffset + x * 4 + 1)
                        dstBytes(xDIB + 2, y) = srcBytes(srcOffset + x * 4)
                        dstBytes(xDIB + 3, y) = srcBytes(srcOffset + x * 4 + 3)
                        
                    Next x
                    Next y
                
                Case png_GreyscaleAlpha
                
                    For y = 0 To yFinal
                        srcOffset = y * xStride + 1
                    For x = 0 To xFinal
                        
                        'Byte order has to be swapped to BGRA
                        g = srcBytes(srcOffset + x * 2)
                        a = srcBytes(srcOffset + x * 2 + 1)
                        xDIB = x * 4
                        dstBytes(xDIB, y) = g
                        dstBytes(xDIB + 1, y) = g
                        dstBytes(xDIB + 2, y) = g
                        dstBytes(xDIB + 3, y) = a
                        
                    Next x
                    Next y
                
            End Select
        
        End If
        
        dstDIB.UnwrapArrayFromDIB dstBytes
        m_Data.UnwrapArrayFromMemoryStream srcBytes
        
    End If

End Function

'1, 2, and 4-bit-per-pixel data must be upsampled to be easily handled in VB.  We do this prior to constructing a
' final image, as it allows us to use identical rendering code for 1/2/4/8-bpp indexed images.
Private Sub UpsampleLowBitDepthData(ByRef srcHeader As PD_PNGHeader, ByRef xStride As Long)
    
    If (srcHeader.BitDepth >= 8) Then Exit Sub
    
    'Prep all the required loop values
    Dim xFinal As Long, yFinal As Long
    Dim x As Long, y As Long, i As Long
    xFinal = srcHeader.Width - 1
    yFinal = srcHeader.Height - 1
        
    Dim curByte As Byte
    
    'We'll use pdStream to generate a new, 8-bpp version of the image; then we can use our standard 8-bpp
    ' decode functions on 1, 2, 4-bpp data.
    Dim tmpStream As pdStream
    Set tmpStream = New pdStream
    tmpStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
    m_Data.SetPosition 0, FILE_BEGIN
    
    Dim numPixelsProcessed As Long
    
    Select Case srcHeader.BitDepth
        
        'There's not a performance-friendly way to mask flags in VB, so we just use a byte array for clarity
        Case 1
        
            Dim bitFlags() As Byte
            ReDim bitFlags(0 To 7) As Byte
            bitFlags(0) = 128
            bitFlags(1) = 64
            bitFlags(2) = 32
            bitFlags(3) = 16
            bitFlags(4) = 8
            bitFlags(5) = 4
            bitFlags(6) = 2
            bitFlags(7) = 1
                    
            'We want to focus on the bytes relevant to the current scanline only
            For y = 0 To yFinal
            
                'Move past the filter byte
                tmpStream.WriteByte m_Data.ReadByte()
                numPixelsProcessed = 0
                        
                'Read through (numOfBytesPerLine) entries, pushing values into the new stream as we go
                For x = 0 To xStride - 2
                
                    curByte = m_Data.ReadByte()
                    
                    'Parse each bit in turn
                    For i = 0 To 7
                        
                        'Ignore empty bytes at the end of each scanline
                        If (numPixelsProcessed <= xFinal) Then
                            If (bitFlags(i) = (curByte And bitFlags(i))) Then tmpStream.WriteByte 1 Else tmpStream.WriteByte 0
                            numPixelsProcessed = numPixelsProcessed + 1
                        End If
                        
                    Next i
                Next x
                
            Next y
                    
        Case 2
            
            Dim shiftTable() As Byte
            ReDim shiftTable(0 To 3) As Byte
            shiftTable(0) = 2 ^ 6
            shiftTable(1) = 2 ^ 4
            shiftTable(2) = 2 ^ 2
            shiftTable(3) = 1
                    
            'Unlike other methods, we want to extract the bytes relevant to the current scanline only
            For y = 0 To yFinal
            
                'Move past the filter byte (which has already been dealt with in a previous step)
                tmpStream.WriteByte m_Data.ReadByte()
                numPixelsProcessed = 0
                
                'Read through (numOfBytesPerLine) entries, pushing values into the new stream as we go
                For x = 0 To xStride - 2
                
                    curByte = m_Data.ReadByte()
                    
                    'Parse each two-bit pair in turn
                    For i = 0 To 3
                        
                        'Ignore empty bytes at the end of each scanline
                        If (numPixelsProcessed <= xFinal) Then
                            tmpStream.WriteByte (curByte \ shiftTable(i)) And &H3
                            numPixelsProcessed = numPixelsProcessed + 1
                        End If
                        
                    Next i
                Next x
            Next y
            
        Case 4
        
            'Unlike other methods, we want to extract the bytes relevant to the current scanline only
            For y = 0 To yFinal
            
                'Move past the filter byte (which has already been dealt with in a previous step)
                tmpStream.WriteByte m_Data.ReadByte()
                numPixelsProcessed = 0
                        
                'Read through (numOfBytesPerLine) entries, pushing values into the new stream as we go
                For x = 0 To xStride - 2
                
                    curByte = m_Data.ReadByte()
                    
                    'Parse each two-bit pair in turn, while ignoring empty bytes at the end of each scanline
                    tmpStream.WriteByte (curByte \ 16) And &HF
                    numPixelsProcessed = numPixelsProcessed + 1
                    
                    If (numPixelsProcessed <= xFinal) Then
                        tmpStream.WriteByte curByte And &HF
                        numPixelsProcessed = numPixelsProcessed + 1
                    End If
                    
                Next x
            Next y
    
    End Select
            
    'Overwrite our original data stream with the new one, and reset xStride to match
    Set m_Data = tmpStream
    xStride = srcHeader.Width + 1
    
End Sub

'To simplify rendering, PD treates grayscale data identically to palette data; we just generate a grayscale palette
' on-the-fly (one matching the current color depth), then render accordingly.
Private Sub ConstructGrayPalette(ByRef srcHeader As PD_PNGHeader, ByRef srcPalette() As RGBQuad, ByRef srcPaletteSize As Long, ByVal srcTrnsRed As Long)

    Dim palOK As Boolean: palOK = True
    
    Select Case srcHeader.BitDepth
        Case 1
            srcPaletteSize = 2
        Case 2
            srcPaletteSize = 4
        Case 4
            srcPaletteSize = 16
        Case 8
            srcPaletteSize = 256
        Case Else
            palOK = False
    End Select
            
    If palOK Then
    
        ReDim srcPalette(0 To srcPaletteSize - 1) As RGBQuad
        
        Dim tmpCalcInt As Long, tmpMax As Long
        tmpMax = srcPaletteSize - 1
        
        Dim i As Long
        For i = 0 To tmpMax
        
            tmpCalcInt = Int(CDbl(i / tmpMax) * 255# + 0.5)
            
            With srcPalette(i)
                .Red = tmpCalcInt
                .Green = tmpCalcInt
                .Blue = tmpCalcInt
                .Alpha = 255
            End With
            
        Next i
        
        'If tRNS is present, assign it now
        If (srcTrnsRed >= 0) And (srcTrnsRed <= 255) Then
            For i = 0 To tmpMax
                If (srcPalette(i).Red = srcTrnsRed) Then srcPalette(i).Alpha = 0
            Next i
        End If
        
    End If
            
End Sub

'If this chunk is a PLTE chunk, use this function to retrieve a palette.  (By default, all alpha values will be set to 255.)
Friend Function GetPalette(ByRef dstQuads() As RGBQuad, ByRef numOfColors As Long, ByRef warningStack As pdStringStack) As Boolean

    GetPalette = False
    
    If (m_Type = "PLTE") Then
    
        'Ensure chunk size is a multiple of 3.  (Non-multiples of 3 are a critical failure.)
        If ((m_Size Mod 3) = 0) Then
            
            numOfColors = m_Size \ 3
            ReDim dstQuads(0 To numOfColors - 1) As RGBQuad
            
            Dim srcBytes() As Byte, tmpSA As SafeArray1D
            m_Data.WrapArrayAroundMemoryStream srcBytes, tmpSA
            
            Dim i As Long
            For i = 0 To numOfColors - 1
                With dstQuads(i)
                    .Red = srcBytes(i * 3)
                    .Green = srcBytes(i * 3 + 1)
                    .Blue = srcBytes(i * 3 + 2)
                    .Alpha = 255
                End With
            Next i
            
            m_Data.UnwrapArrayFromMemoryStream srcBytes
            
            GetPalette = True
        
        Else
            warningStack.AddString "WARNING!  pdPNGChunk.GetPalette found a color count that is *not* a multiple of three.  Processing abandoned."
        End If
        
    Else
        warningStack.AddString "WARNING!  pdPNGChunk.GetPalette was requested of a non-PLTE chunk.  That doesn't work."
    End If

End Function

'If this chunk is a tRNS chunk, use this function to retrieve relevant transparency data.  Handling is automatic
' according to the header's color type.
Friend Function GetTRNSData(ByRef dstPalette() As RGBQuad, ByRef dstTransRed As Long, ByRef dstTransGreen As Long, ByRef dstTransBlue As Long, ByRef srcHeader As PD_PNGHeader, ByRef warningStack As pdStringStack) As Boolean

    GetTRNSData = False

    If (m_Type = "tRNS") Then
        
        m_Data.SetPosition 0, FILE_BEGIN
        
        'Depending on the color type of the source image, we interpret the contents of the tRNS chunk differently.
        ' Types 0 and 2 (grayscale, true-color) allow you to define a single "transparent color" which is always
        ' set to 0 alpha.  Type 3 (indexed color) allows you to set per-index transparency.  Images with full
        ' alpha channels should not have a tRNS chunk at all; we don't enforce this, but we do ignore tRNS if
        ' that's the case.
        Select Case srcHeader.ColorType
        
            Case png_Greyscale
            
                'Regardless of bit-depth, the chunk must be two-bytes long
                If (m_Data.GetStreamSize = 2) Then
                    dstTransRed = m_Data.ReadIntUnsigned_BE()
                    dstTransGreen = dstTransRed
                    dstTransBlue = dstTransRed
                    GetTRNSData = True
                Else
                    warningStack.AddString "WARNING!  pdPNGChunk.GetTRNSData found a grayscale tRNS that wasn't two bytes long (" & m_Data.GetStreamSize() & ")"
                End If
                
            Case png_Truecolor
            
                'Regardless of bit-depth, the chunk must be six-bytes long
                If (m_Data.GetStreamSize = 6) Then
                    dstTransRed = m_Data.ReadIntUnsigned_BE()
                    dstTransGreen = m_Data.ReadIntUnsigned_BE()
                    dstTransBlue = m_Data.ReadIntUnsigned_BE()
                    GetTRNSData = True
                Else
                    warningStack.AddString "WARNING!  pdPNGChunk.GetTRNSData found a truecolor tRNS that wasn't six bytes long (" & m_Data.GetStreamSize() & ")"
                End If
            
            Case png_Indexed
            
                'Indexed color tables support a variable number of alpha values, one byte per original table.
                ' This list of colors can be smaller than the palette itself - if that's the case, unlisted entries
                ' should be treated as fully opaque (which we do anyway, because the palette is initialized to 255).
                Dim i As Long, numBytes As Long
                numBytes = m_Data.GetStreamSize - 1
                If (numBytes <= 256) Then
                    
                    For i = 0 To numBytes
                        dstPalette(i).Alpha = m_Data.ReadByte()
                    Next i
                    
                    GetTRNSData = True
                    
                Else
                    warningStack.AddString "WARNING!  pdPNGChunk.GetTRNSData found a tRNS chunk that's too large to contain usable data (" & m_Data.GetStreamSize() & ")"
                End If
            
        End Select
        
    Else
        warningStack.AddString "WARNING!  pdPNGChunk.GetTRNSData was requested of a non-tRNS chunk.  That doesn't work."
    End If

End Function
