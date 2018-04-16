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
