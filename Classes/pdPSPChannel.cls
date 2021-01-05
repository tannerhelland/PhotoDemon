VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPSPChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PSP (PaintShop Pro) Block Container
'Copyright 2020-2021 by Tanner Helland
'Created: 31/December/20
'Last updated: 04/January/21
'Last update: fix scanline alignment on 1- and 4-bit images (they *ARE* 32-bit aligned after all,
'             which is what the spec says - but 8/24/32/48/64-bpp images are not!  aaaarrrggghhh)
'
'This class describes a single "channel" inside a JASC/Corel Paint Shop Pro image file.
' It has been custom-built for PhotoDemon, with a special emphasis on parsing performance.
'
'Unless otherwise noted, all code in this class is my original work.  I've based my work off the
' "official" PSP spec at this URL (link good as of December 2020):
' ftp://ftp.corel.com/pub/documentation/PSP/
'
'Older PSP specs were also useful.  You may be able to find them here (link good as of December 2020);
' look for files with names like "psp8spec.pdf":
' http://www.telegraphics.com.au/svn/pspformat/trunk
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This helper struct is passed to the child channel class; it contains all information necessary to
' decompress an embedded channel into useable RGBA data.  Some fields are populated by the parent
' (e.g. compression type, which is the same for all channels in a frame and is stored in a
' parent header struct), while others are filled by the child pdPSPChannel class (e.g. compressed
' and uncompressed size, which are stored directly inside the channel header)
Private m_ChannelHeader As PSP_ChannelHeader

'Contents of the decompressed channel stream.  If LoadChannel() returns psp_Success, this array
' is guaranteed to be allocated, initialized, and filled with useable data from the original
' image stream.  (Note, however, that this data has *not* yet been converted to 8-bit RGBA data,
' so you'll need to consider e.g. the parent image's palette to interpret the bytes correctly.
' This is why there's a separate content size tracker - because the data may not have been
' upsampled from lower bit-depths yet, etc.)
Private m_ChannelContents() As Byte, m_ContentSize As Long

Friend Sub FreeChannelContents()
    Erase m_ChannelContents
End Sub

'Return value isn't meaningful until LoadChannel() has been called
Friend Function GetChannelDIBType() As PSPDIBType
    GetChannelDIBType = m_ChannelHeader.ch_dstBitmapType
End Function

'Return value isn't meaningful until LoadChannel() has been called
Friend Function GetChannelPtr() As Long
    GetChannelPtr = VarPtr(m_ChannelContents(0))
End Function

'Return value isn't meaningful until LoadChannel() has been called
Friend Function GetChannelSize() As Long
    GetChannelSize = m_ContentSize
End Function

'Return value isn't meaningful until LoadChannel() has been called
Friend Function GetChannelType() As PSPChannelType
    GetChannelType = m_ChannelHeader.ch_ChannelType
End Function

Friend Function IsChannelOK() As Boolean
    IsChannelOK = m_ChannelHeader.ch_ChannelOK
End Function

'Assuming the source stream is pointing at the start of a channel block, attempt to load the channel.
' Returns psp_Success if successful, psp_Warning if stream alignment is okay but channel data is not,
' psp_Failure if stream alignment is unsaveable.  (If psp_Failure is returned, check initial stream
' pointer alignment - it may not have been pointing at a channel block when you called this function!)
'
'IMPORTANTLY: on psp_Success or psp_Warning, the passed stream pointer will now point at the *end* of
' this block.  You can simply continue reading the file as-is.  On failure, however, stream position
' is *not* guaranteed (mostly because if initial block validation fails, we have no way to reorient the
' pointer in a meaningful way - we can only reset it).  On failure, you need to abandon further parsing.
Friend Function LoadChannel(ByRef srcStream As pdStream, ByRef srcWarnings As pdStringStack, ByRef srcHeader As PSP_ChannelHeader) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "LoadChannel"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'Start with basic block validation
    Const PSP_BLOCK_MARKER As Long = &H4B427E
    Dim blockCheck As Long, blockID As PSPBlockID, blockLength As Long
    
    blockCheck = srcStream.ReadLong()
    If (blockCheck <> PSP_BLOCK_MARKER) Then
        LoadChannel = psp_Failure
        InternalError funcName, "stream misaligned: " & Hex$(blockCheck) & " @ " & srcStream.GetPosition(), srcWarnings
        Exit Function
    End If
    
    blockID = srcStream.ReadIntUnsigned()
    blockLength = srcStream.ReadLong()
    
    'Early PSP versions use an asinine mechanism for struct length descriptors, so this is
    ' a stupid (but necessary) workaround
    If (srcHeader.ch_ParentVersionMajor <= 3) Then blockLength = srcStream.ReadLong()
    
    'If blockID is bad... I'm honestly not sure what to do.  Probably just report it,
    ' then skip to the end of the block?  (Clearly the reader didn't expect this block,
    ' and recovery code would be challenging... basically I'd need to manually investigate
    ' the file if this code is being called on a non-channel chunk.)
    '
    'For now, we do nothing, and leave it to the caller to decide.
    If (blockID <> PSP_CHANNEL_BLOCK) Then
        InternalError funcName, "not a channel!", srcWarnings
        LoadChannel = psp_Failure
        Exit Function
    End If
    
    If (blockLength <= 0) Then
        InternalError funcName, "bad block length!", srcWarnings
        LoadChannel = psp_Failure
        Exit Function
    End If
    
    'If we're still here, the channel block header passed basic validation.
    
    'Before continuing, make a note of the original stream position and the total block length.
    ' We'll need these to align the stream pointer before exiting (on success *or* failure).
    Dim origBlockPosition As Long, origBlockSize As Long
    origBlockPosition = srcStream.GetPosition()
    origBlockSize = blockLength
    
    'The block header is immediately followed by a "channel information chunk".  This describes
    ' basic details on the channel, like offsets to pixel data and un/compressed size.
    
    'Store chunk size separately; we'll use it for stream alignment later
    Dim channelChunkSize As Long
    
    'Early PSP versions use an asinine mechanism for struct length descriptors, so this is
    ' a stupid (but necessary) workaround
    If (srcHeader.ch_ParentVersionMajor <= 3) Then
        srcStream.SetPosition -8, FILE_CURRENT
        channelChunkSize = srcStream.ReadLong()
        srcStream.SetPosition 4, FILE_CURRENT
    Else
        channelChunkSize = srcStream.ReadLong()
    End If
    
    'Compressed size is critical since this is the only indicator of the length of the
    ' deflate or RLE-compressed stream
    srcHeader.ch_CompressedSize = srcStream.ReadLong()
    srcHeader.ch_UncompressedSize = srcStream.ReadLong()
    
    'Prepare the channel's decompression buffer
    m_ContentSize = srcHeader.ch_UncompressedSize
    ReDim m_ChannelContents(0 To m_ContentSize - 1) As Byte
    
    'Bitmap type is an internal PSP struct; we can usually infer this from other data,
    ' but it's a nice confirmation of where this channel should ultimately land
    srcHeader.ch_dstBitmapType = srcStream.ReadIntUnsigned()
    srcHeader.ch_ChannelType = srcStream.ReadIntUnsigned()
    
    'Future expansion members are explicitly allowed by the spec.  Use the previously
    ' retrieved chunk length to align the stream pointer.
    srcStream.SetPosition origBlockPosition + channelChunkSize, FILE_BEGIN
    
    'The struct now points at channel contents!  Time to decompress it.
    ' (Obviously, the manner in which we decompress it is entirely dependent on the compression
    ' scheme used.)
    Select Case srcHeader.ch_Compression
    
        'Uncompressed means we can just copy bytes out as-is
        Case PSP_COMP_NONE
            CopyMemoryStrict VarPtr(m_ChannelContents(0)), srcStream.Peek_PointerOnly(-1, srcHeader.ch_CompressedSize), srcHeader.ch_CompressedSize
            okToProceed = psp_Success
            
        'PSP format defines an extremely simple RLE scheme
        Case PSP_COMP_RLE
            
            'The PSP spec defines RLE behavior nicely; see it for details
            Dim srcByte As Byte, srcLong As Long, numBytesProcessed As Long
            numBytesProcessed = 0
            
            Dim numSafeBytesDst As Long, numSafeBytesSrc As Long
            numSafeBytesDst = srcHeader.ch_UncompressedSize
            numSafeBytesSrc = srcHeader.ch_CompressedSize
            
            'For PSP's RLE encoding, we can effectively ignore scanlines and just treat the data
            ' as an arbitrary stream of bytes.  Note that this function is *not* necessarily
            ' guaranteed to be robust against malformed data; e.g. it may crash if fed deliberately
            ' corrupted RLE data that produces runs extending beyond the end of the image.
            ' (I've tried to account for this as best I can, but I don't have an easy way to
            ' aggressively fuzz the decompressor!)
            Dim curOffset As Long
            curOffset = 0
            
            Do While (curOffset < numSafeBytesDst) And (numBytesProcessed < numSafeBytesSrc)
            
                'Get the RLE byte
                srcLong = srcStream.ReadByte()
                numBytesProcessed = numBytesProcessed + 1
                
                'This is a run.  Repeat the next byte [n-128] times
                If (srcLong > 128) Then
                    
                    srcLong = srcLong - 128
                    
                    'Safety check
                    If (curOffset + srcLong > numSafeBytesDst) Then
                        InternalError funcName, "bad RLE run: " & CStr(curOffset + srcLong) & " vs " & CStr(numSafeBytesDst), srcWarnings
                        srcLong = numSafeBytesDst - curOffset
                    End If
                    
                    'Use FillMemory for fast byte-filling
                    srcByte = srcStream.ReadByte()
                    numBytesProcessed = numBytesProcessed + 1
                    
                    VBHacks.FillMemory VarPtr(m_ChannelContents(curOffset)), srcLong, srcByte
                    curOffset = curOffset + srcLong
                
                'This is a segment of [n] uncompressed bytes.  Read the bytes directly into
                ' the target buffer.
                Else
                    
                    'Safety check
                    If (curOffset + srcLong > numSafeBytesDst) Then
                        InternalError funcName, "bad RLE chunk: " & CStr(curOffset + srcLong) & " vs " & CStr(numSafeBytesDst), srcWarnings
                        srcLong = numSafeBytesDst - curOffset
                    End If
                    
                    srcStream.ReadBytesToBarePointer VarPtr(m_ChannelContents(curOffset)), srcLong
                    curOffset = curOffset + srcLong
                    numBytesProcessed = numBytesProcessed + srcLong
                    
                End If
                
            Loop
            
            okToProceed = psp_Success
            
        'Per the spec, PSP's LZ77 scheme is the same as PNG's - so it's DEFLATE to the rescue!
        Case PSP_COMP_LZ77
            Compression.DecompressPtrToPtr VarPtr(m_ChannelContents(0)), srcHeader.ch_UncompressedSize, srcStream.Peek_PointerOnly(-1, srcHeader.ch_CompressedSize), srcHeader.ch_CompressedSize, cf_Zlib
            okToProceed = psp_Success
            
        'NOTE: you may notice that JPEG compression is missing.  That's on purpose, because JPEG
        ' encoding is invalid for individual channels.  JPEG compression is only allowed for
        ' composite thumbnails, and PhotoDemon's decoder has a separate function for handling those
        ' (inside the parent pdPSP class).  If we encounter JPEG compression inside a channel,
        ' it's a broken file.
        Case Else
            InternalError funcName, "bad compression", srcWarnings
            okToProceed = psp_Failure
    
    End Select
    
    'If decompression was successful, move the pointer to the end of the chunk
    If (okToProceed < psp_Failure) Then srcStream.SetPosition origBlockPosition + origBlockSize, FILE_BEGIN
    
    'Mark the channel as OK/not-OK
    srcHeader.ch_ChannelOK = (okToProceed = psp_Success)
    
    'Make a local copy of the passed header, as we may need it on subsequent calls
    m_ChannelHeader = srcHeader
    
    'If the channel is OK and < 8-bit, immediately upsample it to 8-bits
    If srcHeader.ch_ChannelOK Then
        If (srcHeader.ch_ParentBitDepth < 8) Then srcHeader.ch_ChannelOK = UpsampleLowBitDepth(m_ChannelHeader, srcWarnings)
    End If
    
    'HDR images are still TODO (need some for testing)
    
    'Our work here is done!
    LoadChannel = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description, srcWarnings
    srcWarnings.AddString "Internal error in pdPSPChannel." & funcName & ", #" & Err.Number & ": " & Err.Description
    LoadChannel = psp_Failure
    
End Function

'On 1-bpp and 4-bpp images, we want to upsample their channel data to 8-bpp before
' rendering a final DIB.
Friend Function UpsampleLowBitDepth(ByRef srcHeader As PSP_ChannelHeader, ByRef srcWarnings As pdStringStack) As Boolean
    
    Const funcName As String = "UpsampleLowBitDepth"
    
    If (srcHeader.ch_ParentBitDepth >= 8) Then
        InternalError funcName, "can't upsample an 8+bit channel", srcWarnings
        UpsampleLowBitDepth = True
        Exit Function
    End If
    
    'Make a local copy of the channel's contents
    Dim curBytes() As Byte, oldSize As Long
    oldSize = m_ContentSize
    
    ReDim curBytes(0 To oldSize - 1) As Byte
    CopyMemoryStrict VarPtr(curBytes(0)), VarPtr(m_ChannelContents(0)), oldSize
    
    'Weirdly, testing random images from online seems to show that PSP scanlines
    ' *ARE* aligned on 4-byte boundaries... but only for 1- and 4-bit images.
    ' (8-bit and up do *not* obey the spec's requested 4-byte alignment rule,
    ' and in fact enforcing 4-byte alignment on those images will break otherwise
    ' good ones.)
    Dim pxWidth As Long, pxHeight As Long
    pxWidth = srcHeader.ch_ParentWidth
    pxHeight = srcHeader.ch_ParentHeight
    
    Dim pxWidthAligned As Long
    pxWidthAligned = ((pxWidth + 3) \ 4) * 4
    
    'Figure out what size we want the final array to be
    m_ContentSize = srcHeader.ch_ParentWidth * srcHeader.ch_ParentHeight
    ReDim m_ChannelContents(0 To m_ContentSize - 1) As Byte
    
    Dim srcScanlineSize As Long, srcByte As Byte
    
    'NOTE: monochrome testing still required; I don't have monochrome images yet,
    ' so this conversion code is just a placeholder.  (It's possible source scanlines
    ' break on byte boundaries - if so, different code will be required.)
    Dim x As Long, y As Long, i As Long, numPixelsProcessed As Long
    If (srcHeader.ch_ParentBitDepth = 1) Then
        
        srcScanlineSize = (pxWidthAligned + 7) \ 8
        
        Dim bitFlags(0 To 7) As Byte
        For i = 0 To 7
            bitFlags(i) = 2 ^ (7 - i)
        Next i
        
        For y = 0 To pxHeight - 1
            numPixelsProcessed = 0
        For x = 0 To srcScanlineSize - 1
        
            srcByte = curBytes(y * srcScanlineSize + x)
            
            'Ignore empty bytes at the end of each scanline
            For i = 0 To 7
                If (numPixelsProcessed < pxWidth) Then
                    If (bitFlags(i) = (srcByte And bitFlags(i))) Then m_ChannelContents(y * pxWidth + numPixelsProcessed) = 1 Else m_ChannelContents(y * pxWidthAligned + numPixelsProcessed) = 0
                    numPixelsProcessed = numPixelsProcessed + 1
                Else
                    Exit For
                End If
            Next i
            
        Next x
        Next y
    
    'NOTE: 4-bite testing still required; I don't have 4-bit images yet,
    ' so this conversion code is just a placeholder.
    ElseIf (srcHeader.ch_ParentBitDepth = 4) Then
        
        srcScanlineSize = (pxWidthAligned + 1) \ 2
        
        For y = 0 To pxHeight - 1
            numPixelsProcessed = 0
        For x = 0 To srcScanlineSize - 1
            
            srcByte = curBytes(y * srcScanlineSize + x)
            
            'Ignore empty bytes at the end of each scanline
            m_ChannelContents(y * pxWidth + x * 2) = (srcByte \ 16) And &HF
            numPixelsProcessed = numPixelsProcessed + 1
            
            If (numPixelsProcessed < pxWidth) Then
                m_ChannelContents(y * pxWidth + x * 2 + 1) = srcByte And &HF
                numPixelsProcessed = numPixelsProcessed + 1
                If (numPixelsProcessed >= pxWidth) Then Exit For
            Else
                Exit For
            End If
            
        Next x
        Next y
        
    Else
        InternalError funcName, "unexpected bit-depth: " & srcHeader.ch_ParentBitDepth, srcWarnings
        UpsampleLowBitDepth = False
        Exit Function
    End If

End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, ByRef parentWarningStack As pdStringStack, Optional ByVal writeDebugLog As Boolean = True)
    
    Dim errText As String
    errText = "pdPSPChannel." & funcName & "() reported an error: " & errDescription
    If (Not parentWarningStack Is Nothing) Then parentWarningStack.AddString errText
    
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction errText
    Else
        Debug.Print errText
    End If
    
End Sub
