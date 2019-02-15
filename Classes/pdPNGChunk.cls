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
'Copyright 2018-2019 by Tanner Helland
'Created: 12/April/18
'Last updated: 23/April/18
'Last update: new helper function for cHRM chunks
'
'pdPNG leans on this class to handle the actual storage (and deciphering) of certain PNG chunks.
' This simplifies the decode process because we only deal with the chunks that matter to PD,
' and it gives us an convenient way to do things like concatenate IDAT chunks post-load.
' (PNG allows image data to be split over multiple IDAT chunks, which is perhaps useful for
' streaming over a network connection, but is pointless for loading files from a local drive.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
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

'Some external functions (e.g. upsampling low-bit-depth pixel data) may produce a new pdStream object for us to use.
' When they're finished, they can notify us of the new stream target via this function.  (Obviously, things like
' CRCs will not be relevant after the underlying data changes; don't reuse them after calling this function!)
Friend Sub SubmitNewDataStream(ByRef srcStream As pdStream)
    Set m_Data = srcStream
    m_Size = srcStream.GetStreamSize()
End Sub

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
        DecompressChunk = Compression.DecompressPtrToPtr(tmpBuffer.Peek_PointerOnly(0), inflateSize, m_Data.Peek_PointerOnly(0), m_Size, cf_Zlib)
        
        If DecompressChunk Then
            DecompressChunk = (initSize = inflateSize)
            If (Not DecompressChunk) Then warningStack.AddString "pdPNGChunk.DecompressChunk received mismatched sizes from libdeflate: " & initSize & ", " & inflateSize
        Else
            warningStack.AddString "pdPNGChunk.DecompressChunk received a failure state from libdeflate; FYI sizes were init: " & initSize & ", " & inflateSize
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
                
                'Decompression is going to fail - we know that - but it's okay.  We know 128 bytes isn't enough to extract
                ' the entire profile.  (We do, however, need to fail specifically because the buffer is too small, and not
                ' because of some other problem state.)
                Dim writeSize As Long: writeSize = 128
                Const ERR_BUFF_TOO_SMALL As Long = 3
                
                Dim ldResult As Long
                ldResult = Plugin_libdeflate.Decompress_ZLib(tmpBuffer.Peek_PointerOnly(0), writeSize, m_Data.Peek_PointerOnly(curOffset), m_Size - curOffset)
                
                If (ldResult = ERR_BUFF_TOO_SMALL) Then
                    
                    'The ICC header has now been dumped into our temporary stream.  Retrieve the profile header size.
                    tmpBuffer.SetSizeExternally 128
                    tmpBuffer.SetPosition 0, FILE_BEGIN
                    
                    Dim profSize As Long
                    profSize = tmpBuffer.ReadLong_BE()
                    
                    'We can now reset our temporary buffer to that size, and extract the *entire* ICC profile
                    tmpBuffer.StopStream True
                    tmpBuffer.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, , profSize
                    DecompressChunk = Compression.DecompressPtrToPtr(tmpBuffer.Peek_PointerOnly(0), profSize, m_Data.Peek_PointerOnly(curOffset), m_Size - curOffset, cf_Zlib)
                    tmpBuffer.SetSizeExternally profSize
                    
                    'Compression was successful!  Overwrite our stream with the contents of the ICC profile, then exit.
                    If DecompressChunk Then
                        m_Data.StopStream True
                        Set m_Data = tmpBuffer
                        m_Size = profSize
                    Else
                        warningStack.AddString "pdPNGChunk.DecompressChunk couldn't decompress the ICC chunk; ICC retrieval was *not* successful."
                    End If
                    
                Else
                    warningStack.AddString "pdPNGChunk.DecompressChunk couldn't extract a valid ICC header (" & CStr(ldResult) & "); ICC retrieval was *not* successful."
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
    
    'Reset the data stream pointer to the start of the stream
    m_Data.SetPosition 0, FILE_BEGIN
    
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
    
    'Multi-channel images (or images with multi-byte channel values) require offsets during the unfilter phase.
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
    
    'A note throughout this function: per the PNG spec:
    ' - "Unsigned arithmetic modulo 256 is used, so that both the inputs and outputs fit into bytes."
    ' - "Filters are applied to each byte regardless of bit depth."
    
    'As usual, interlaced images are their own beasts.  For performance reasons, we split handling for interlaced
    ' images into their own function (as they must be processed in a more time-consuming way).
    If imgIsInterlaced Then
    
        'Interlaced PNGs are actually stored as 7 independent "IDAT streams", one for each interlace pass.
        ' Each stream represents a "mini-PNG", a complete image comprised of scanlines unique to that pass.
        ' As a consequence, in very small images, some of the streams will be empty if the interlacing offsets
        ' produce no pixels that lie within image bounds.
        
        'Each interlace pass thus needs to be unfiltered independently, as if it were its own micro-image.
        ' This means that each scanline in an interlace pass is unfiltered against *other bytes in that
        ' interlace pass, only*.
        
        'To handle this cleanly, let's cheat.  Rather than writing complicated per-pass filtering code,
        ' let's break up the original interlaced IDAT stream into seven indepedent *non-interlaced* ones.
        ' We'll create the chunks on-the-fly, ask them to perform standard non-interlaced unfiltering,
        ' after which we'll reassemble their unfiltered bits into a coherent,
        ' "unfiltered-but-still-interlaced" stream.
        
        'Start by prepping a "fake" header; this will retain some data from the original PNG header
        ' (like color-type and bit-depth, which haven't changed), but we'll modify things like width/height
        ' on-the-fly to match each individual chunk.
        Dim tmpHeader As PD_PNGHeader
        tmpHeader = srcHeader
        tmpHeader.Interlaced = False
        
        'Notice how this function (UnfilterChunk) accepts arrays for x/y pixel count and x byte count (stride)?
        ' For interlaced images, these arrays contain unique data for each interlacing pass (0-6, 7 passes total).
        ' We need to pass arrays like this to each of our "fake" IDAT chunks, one per pass, but like normal
        ' non-interlaced images, the arrays we pass will need to be only 1 entry long.
        Dim tmpChunk As pdPNGChunk
        Dim fakeXPixelCount() As Long, fakeYPixelCount() As Long, fakeXByteCount() As Long
        ReDim fakeXPixelCount(0) As Long
        ReDim fakeYPixelCount(0) As Long
        ReDim fakeXByteCount(0) As Long
        
        Dim passOffset As Long, sizeOfPass As Long, sizeOfPassCheck As Long
        passOffset = 0
        
        'Process each interlacing pass in turn
        Dim i As Long
        For i = 0 To 6
            
            'We can entirely skip 0-length passes
            If (xByteCount(i) > 0) Then
            
                'Figure out the total size of this pass, in bytes.
                sizeOfPass = xByteCount(i) * yPixelCount(i)
                
                'Initialize the new chunk and prep its internal buffer to the required size
                Set tmpChunk = New pdPNGChunk
                tmpChunk.CreateChunk "IDAT", sizeOfPass
                
                'Use the size of this pass to "extract" the relevant portion of the original IDAT stream
                ' into the temporary chunk's internal buffer.  (Note that, like all .Read* functions, this call
                ' will increment the internal pointer of the m_Data stream.)  As a failsafe, we'll also
                ' double-check that the read size matches the size we requested - this ensures we never attempt
                ' to read past the end of the stream.
                sizeOfPassCheck = m_Data.ReadBytesToBarePointer(tmpChunk.BorrowData.Peek_PointerOnly(0), sizeOfPass)
                If (sizeOfPassCheck = sizeOfPass) Then
                    tmpChunk.BorrowData.SetSizeExternally sizeOfPass
                Else
                    tmpChunk.BorrowData.SetSizeExternally sizeOfPassCheck
                    warningStack.AddString "WARNING!  pdPNGChunk.UnfilterChunk failed to read an interlacing pass correctly (" & sizeOfPass & ", " & sizeOfPassCheck & ")"
                End If
                
                'With the temporary chunk assembled, we now need to update our "fake" PNG header to match this pass.
                ' (Specifically, width and height need to be updated.)
                tmpHeader.Width = xPixelCount(i)
                tmpHeader.Height = yPixelCount(i)
                
                'Similarly, we need to prep fake x/y pixel counts and stride arrays.
                fakeXPixelCount(0) = xPixelCount(i)
                fakeYPixelCount(0) = yPixelCount(i)
                fakeXByteCount(0) = xByteCount(i)
                
                'We now have everything we need to unfilter this pass.
                tmpChunk.UnfilterChunk warningStack, tmpHeader, fakeXPixelCount, fakeYPixelCount, fakeXByteCount
                
                'With unfiltering complete, we can now extract the unfiltered bytes *back* into *our* data stream.
                ' (Note that we copy them in-place over the original, unfiltered bytes - and to determine the
                ' position to copy to, we use the current position, minus the size of this pass (because that's how
                ' far the pointer was incremented on the previous .Read call).
                tmpChunk.BorrowData.ReadBytesToBarePointer m_Data.Peek_PointerOnly(m_Data.GetPosition() - sizeOfPass), sizeOfPass
                
                'We are now done with this chunk.  We don't need to clear it, as we'll automatically free it when
                ' we process the next chunk.
                
            End If
        
        Next i
        
        'We have successfully unfiltered all interlaced pixel data - but note that the data stream is still in
        ' *interlaced* order (with pixels arranged in subgroups, according to interlace pass).
    
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
        Dim p As Long, pa As Long, pb As Long, pc As Long
        
        Dim x As Long, y As Long, yOffset As Long
        For y = 0 To yFinal
        
            yOffset = y * xStride
            
            'Filtering changes based on the "filter byte" at this position
            Select Case srcBytes(yOffset)
                
                'Regardless of image properties, unfiltered bytes don't require additional processing
                Case png_None
                
                '"Sub" filtering filters according to the previous pixel
                Case png_Sub
                
                    yOffset = yOffset + 1
                    For x = pxOffset To xStride - 2
                        prevValue = srcBytes(yOffset + x - pxOffset)
                        curValue = srcBytes(yOffset + x)
                        srcBytes(yOffset + x) = (curValue + prevValue) And 255
                    Next x
                    
                '"Up" filtering filters according to the pixel above
                Case png_Up
                    
                    'First line has no "previous" line; values are absolute.
                    If (y <> 0) Then
                        yOffset = yOffset + 1
                        For x = 0 To xStride - 2
                            prevValue = srcBytes(yOffset + x - xStride)
                            curValue = srcBytes(yOffset + x)
                            srcBytes(yOffset + x) = (curValue + prevValue) And 255
                        Next x
                    End If
                    
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
                        
                        'The Paeth predictor was previously used as a separate function, but inlining it
                        ' provides a slight performance boost.  (Also, the function itself is simple - it just
                        ' tries to find the previous pixel with the smallest variance from the current one.)
                        p = (prevValue + prevValue2 - prevValue3)
                        pa = Abs(p - prevValue)
                        pb = Abs(p - prevValue2)
                        pc = Abs(p - prevValue3)
                        If (pa <= pb) Then
                            If (pa <= pc) Then p = prevValue Else p = prevValue3
                        Else
                            If (pb <= pc) Then p = prevValue2 Else p = prevValue3
                        End If
                        
                        srcBytes(yOffset + x) = (curValue + p) And 255
                        
                    Next x
            
            End Select
            
        Next y
        
        m_Data.UnwrapArrayFromMemoryStream srcBytes
        
    End If
    
End Function

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

'bKGD chunks return the pure, unadulterated bKGD contents.  Note that the length of the return varies based
' on the parent image's color type.  From the spec: "For colour type 3 (indexed-colour), the value is the
' palette index of the colour to be used as background.  For colour types 0 and 4 (greyscale, greyscale with alpha),
' the value is the grey level to be used as background in the range 0 to (2^bitdepth)-1. For colour types 2 and 6
' (truecolour, truecolour with alpha), the values are the colour to be used as background, given as RGB samples in
' the range 0 to (2^bitdepth)-1. In each case, for consistency, two bytes per sample are used regardless of the image
' bit depth. If the image bit depth is less than 16, the least significant bits are used and the others are 0.
'
'At present, RGB background colors in 48/64-bpp images are returned as a concatenated 24-bpp RGB Long.  (This may need
' to be revisited if internal 48/64-bpp tracking is ever expanded.)
Friend Function GetBackgroundColor(ByVal parentIsHighBitDepth As Boolean, ByRef warningStack As pdStringStack) As Long
    
    If (m_Type = "bKGD") Then
        
        m_Data.SetPosition 0, FILE_BEGIN
        
        '1-byte chunks are palette indices
        If (m_Size = 1) Then
            GetBackgroundColor = m_Data.ReadByte()
        
        '2-byte chunks are grayscale
        ElseIf (m_Size = 2) Then
        
            Dim l As Long
            l = m_Data.ReadIntUnsigned_BE()
            If parentIsHighBitDepth Then GetBackgroundColor = RGB(l \ 256, l \ 256, l \ 256) Else GetBackgroundColor = RGB(l, l, l)
            
        'The only remaining option is 6-byte RGB data.  Extract the bytes individually, then scale as necessary.
        Else
        
            Dim r As Long, g As Long, b As Long
            r = m_Data.ReadIntUnsigned_BE()
            g = m_Data.ReadIntUnsigned_BE()
            b = m_Data.ReadIntUnsigned_BE()
            If parentIsHighBitDepth Then GetBackgroundColor = RGB(r \ 256, g \ 256, b \ 256) Else GetBackgroundColor = RGB(r, g, b)
            
        End If
    Else
        warningStack.AddString "WARNING!  You called pdPNGChunk.GetBackgroundColor on a non-bKGD chunk."
    End If
    
End Function

'gAMA chunks will return the pure, unadulterated gAMA contents - which is, per the spec,
' "[a] four-byte unsigned integer, representing gamma times 100000."  We return the integer version to clarify
' the handling of the case gamma = 1.0, which is the most common case, and the one that allows us to avoid
' further processing.
Friend Function GetGammaData(ByRef warningStack As pdStringStack) As Long
    If (m_Type = "gAMA") Then
        m_Data.SetPosition 0, FILE_BEGIN
        GetGammaData = m_Data.ReadLong_BE()
    Else
        warningStack.AddString "WARNING!  You requested gamma data from a non-gAMA chunk; no value was returned."
    End If
End Function

'cHRM chunks return the pure, unadulterated cHRM contents - which are, per the spec,
' "the 1931 CIE x,y chromaticities of the red, green, and blue display primaries used in the image,
' and the referenced white point.  Each value is encoded as a four-byte PNG unsigned integer,
' representing the x or y value times 100000."
'
'The passed array will be redimmed [0, 7], and filled with chromaticity data in the following order:
' White point x/y, Red x/y, Green x/y, Blue x/y
Friend Function GetChromaticityData(ByRef dstData() As Long, ByRef warningStack As pdStringStack) As Boolean
    
    GetChromaticityData = (m_Type = "cHRM")
    If GetChromaticityData Then
        
        m_Data.SetPosition 0, FILE_BEGIN
        ReDim dstData(0 To 7) As Long
        
        Dim i As Long
        For i = 0 To 7
            dstData(i) = m_Data.ReadLong_BE()
        Next i
        
    Else
        warningStack.AddString "WARNING!  You requested chromaticity data from a non-cHRM chunk; no value was returned."
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
