VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPSDLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PSD (PhotoShop Image) Layer Container and Parser
'Copyright 2019-2019 by Tanner Helland
'Created: 15/January/19
'Last updated: 15/January/19
'Last update: continued work on initial build
'
'This class contains layer-specific data pulled from a PSD file.  It is populated by a parent
' pdPSD instance.  It is not usable outside of a PSD parsing context.
'
'All code in this class is my original work.  It is based off the "official" Adobe spec at this URL
' (link good as of January 2019):
' https://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577409_72092
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Various rects; these have been translated to the proper order (Adobe uses top/left/bottom/right order)
Private m_Rect As RectL, m_MaskRect As RectL, m_RealMaskRect As RectL

'Channel count can vary on a per-layer basis
Private m_ChannelCount As Long

'Each channel in the layer has a unique ID (e.g. red/green/blue/alpha) and length.
Public Enum PSD_ChannelID
    cidRealUserLayerMask = -3
    cidUserLayerMask = -2
    cidAlpha = -1
    cidRed = 0
    cidGreen = 1
    cidBlue = 2
    cidCyan = 0
    cidMagenta = 1
    cidYellow = 2
    cidKey = 3
End Enum

#If False Then
    Private Const cidRealUserLayerMask = -3, cidUserLayerMask = -2, cidAlpha = -1, cidRed = 0, cidGreen = 1, cidBlue = 2, cidCyan = 0, cidMagenta = 1, cidYellow = 2, cidKey = 3
#End If

Private Enum PSD_ChannelCompression
    ccRaw = 0
    ccPackBits = 1
    ccZip = 2
    ccZipWithPrediction = 3
End Enum

#If False Then
    Private Const ccRaw = 0, ccPackBits = 1, ccZip = 2, ccZipWithPrediction = 3
#End If

Private Type PSD_ChannelInfo
    ciID As PSD_ChannelID
    ciSizeL As Long
    ciSizeH As Long    'Only used by PSB format; not currently supported
    
    'The following members are filled in a separate pass (their data is not stored in a different
    ' segment of the file)
    ciCompression As PSD_ChannelCompression
    ciDataLength As Long
    
    'The interpretation of ciData() depends on the value of ciCompression (above). It is relatively common
    ' for PSD data to be stored as raw-encoded, in which case ciData() provides relevant raw pixel values.
    ' If the data has been compressed via PackBits or Deflate, however, ciData() needs to be decoded before
    ' you can get usable information from it.  The DecodeChannels() function handles this.
    ciData() As Byte
    
    'After a channel has been decoded, the new raw data will be placed here.
    ciDataDecoded() As Byte
    
End Type

Private m_Channels() As PSD_ChannelInfo

'Layers can include additional optional data.  This comes in a variety of shapes and sizes,
' and while PD will successfully parse it all, we don't make use of every possible descriptor.
Private Type PSD_AdditionalInfo
    aiSignature As String * 4
    aiKey As String * 4
    aiDataLength As Long
    aiDataBytes() As Byte
End Type

Private m_AdditionalInfo() As PSD_AdditionalInfo
Private m_AdditionalInfoCount As Long

'Photoshop provides some blend modes that PD does not.  We attempt to convert all supported blend modes
' to our own internal constant, but if we fail, we'll fall back to "normal" mode.
Private m_BlendMode As PD_BlendMode

Private m_Opacity As Byte, m_Clipping As Byte, m_Flags As Byte, m_Visible As Boolean, m_LenExtraData As Long
Private m_LayerNameANSI As String, m_LayerNameUnicode As String

'The final result of all channel decoding is this pdDIB object, ready for attachment to a pdLayer.
Private m_LayerDIB As pdDIB

Friend Function DoesLayerHaveUsableData() As Boolean
    DoesLayerHaveUsableData = (m_ChannelCount > 0) And (Not m_LayerDIB Is Nothing)
End Function

Friend Function GetLayerName() As String
    If (LenB(m_LayerNameUnicode) <> 0) Then GetLayerName = m_LayerNameUnicode Else GetLayerName = m_LayerNameANSI
End Function

Friend Function GetLayerDIB() As pdDIB
    Set GetLayerDIB = m_LayerDIB
End Function

Friend Function GetLayerOffsetX() As Long
    GetLayerOffsetX = m_Rect.Left
End Function

Friend Function GetLayerOffsetY() As Long
    GetLayerOffsetY = m_Rect.Top
End Function

Friend Function GetLayerBlendMode() As PD_BlendMode
    GetLayerBlendMode = m_BlendMode
End Function

Friend Function GetLayerOpacity() As Single
    GetLayerOpacity = CSng(m_Opacity) / 2.55!
End Function

Friend Function GetLayerVisibility() As Boolean
    GetLayerVisibility = m_Visible
End Function

'After channel data has been retrieved from the PSD, we must decode it into something usable by PD.
' That's what this function handles.  If the underlying data is irrelevant to PD (e.g. adjustment or effect layers),
' this function simply frees the associated layer data to cut down on memory requirements.  Layers with relevant
' image data, however, will produce a usable pdDIB object from this step.
Friend Function DecodeChannels(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile) As PD_PSDResult
    Debug.Print "here"
    DecodeChannels = psd_Success
    
    'Support for additional formats are being added all the time.  Reject any unwanted combinations for now.
    If (bitsPC <> 8) Or (imgColorMode <> cm_RGB) Then
        warningStack.AddString "The current color mode (" & imgColorMode & " ) and bpc (" & bitsPC & " is still under construction."
        Debug.Print "abandoning load"
        Set m_LayerDIB = Nothing
        DecodeChannels = psd_Failure
        Exit Function
    Else
        Debug.Print "attempting load"
    End If
    
    On Error GoTo InternalVBError
    
    'All channels in a layer are typically stored using the same compression format.
    ' However, this is not guaranteed, so we must load and process each channel independently from the others.
    
    'Start by preparing a pdDIB object.  This will ultimately hold the results of our parsing, although the
    ' data may get convolved through a number of intermediary formats first.
    Set m_LayerDIB = New pdDIB
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    m_LayerDIB.CreateBlank layerWidth, layerHeight, 32, vbWhite, 255
    
    Dim bytesPerChannel As Long
    bytesPerChannel = (bitsPC \ 8)
    
    ' data is stored in planar format (e.g. not interleaved - RRRGGGBBB instead of RGBRGBRGB).
    ' As such, each channel must be loaded independent of other channels in the image.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        
        'If the target layer DIB was freed by a previous channel, exit immediately
        If (m_LayerDIB Is Nothing) Then Exit For
        
        With m_Channels(i)
        
            'Some channels contribute nothing to the final image (and are zero-length).  Ignore them.
            If (.ciDataLength = 0) Then
                Debug.Print "Zero-length channel found."
                GoTo SkipChannel
            End If
            
            'Our first goal is to translate the PSD channel stream into a standalone byte array
            ' at its native bit-per-channel size.  The amount of work required to do this varies based
            ' on the compression technique used.
            
            'We know the size of the finished channel data from the image's dimensions and current bit-depth.
            Dim numPixels As Long
            numPixels = layerWidth * layerHeight
            
            Dim finalArraySize As Long
            finalArraySize = numPixels * bytesPerChannel
            
            'Monochrome layers will need to be handled separately
            If (finalArraySize > 0) Then
                
                'Prep a temporary output buffer
                Dim tmpBytes() As Byte
                Dim x As Long, y As Long, tmpSA As SafeArray1D
                Dim scanWidth As Long, scanStart As Long
                scanStart = m_LayerDIB.GetDIBPointer()
                scanWidth = m_LayerDIB.GetDIBStride()
                
                Debug.Print "Compression type: " & .ciCompression
                
                Select Case .ciCompression
                
                    'Raw compression = no compression.  Use the bytes as-is.
                    Case ccRaw
                    
                        'Technically, we could do something clever like swap safearray headers between the
                        ' coded and decoded type members - but for now, simply perform a copy to keep
                        ' subsequent code simple.
                        ReDim .ciDataDecoded(0 To .ciDataLength - 1) As Byte
                        CopyMemoryStrict VarPtr(.ciDataDecoded(0)), VarPtr(.ciData(0)), .ciDataLength
                        
                    'PackBits = RLE compression.  Data must be decoded before transferring it to the image.
                    Case ccPackBits
                        
                        'We know the necessary size of the final array; prep the receiving buffer now.
                        ReDim .ciDataDecoded(0 To finalArraySize - 1) As Byte
                        
                        'For convenience, we're going to wrap a pdStream object around the source bytes.
                        ' This simplifies things like translating from big-endian values.
                        Dim rleStream As pdStream
                        Set rleStream = New pdStream
                        rleStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , .ciDataLength, VarPtr(.ciData(0))
                        
                        'RLE data starts with a list of scanline sizes (as each scanline gets compressed
                        ' individually).  The size of individual size records varies by image type.
                        Dim lineLengths() As Long
                        ReDim lineLengths(0 To layerHeight - 1) As Long
                        If imageIsPSB Then
                            For y = 0 To layerHeight - 1
                                lineLengths(y) = rleStream.ReadLong_BE()
                            Next y
                        Else
                            For y = 0 To layerHeight - 1
                                lineLengths(y) = rleStream.ReadIntUnsigned_BE()
                            Next y
                        End If
                        
                        'We now want to extract each scanline in turn.  (Technically, we don't need to
                        ' know RLE line lengths to do this, but they provide a nice failsafe mechanism
                        ' for ensuring we don't overrun line boundaries.)
                        Dim bytesReadTotal As Long, bytesReadLine As Long, rleCode As Long, tmpVal As Long
                        bytesReadTotal = 0
                        
                        For y = 0 To layerHeight - 1
                            
                            bytesReadLine = 0
                            
                            Do
                            
                                'Packbits is extremely simple as compression strategies go.  Treat the first
                                ' (signed) byte in each stream as a code.
                                rleCode = rleStream.ReadByte()
                                If (rleCode > 127) Then rleCode = rleCode - 256
                                
                                'If the code is >= 0, the next (code) bytes should be copied to the destination
                                ' buffer as-is.
                                If (rleCode >= 0) Then
                                    rleCode = rleCode + 1
                                    tmpVal = rleStream.ReadBytes_PointerOnly(rleCode)
                                    CopyMemoryStrict VarPtr(.ciDataDecoded(bytesReadTotal)), tmpVal, rleCode
                                    bytesReadLine = bytesReadLine + rleCode + 1
                                    bytesReadTotal = bytesReadTotal + rleCode
                                    
                                'If the code is on the range [-127,-1], inclusive, the next byte should be
                                ' repeated [1 - code] times in the destination buffer.
                                ElseIf (rleCode > -128) Then
                                    rleCode = 1 - rleCode
                                    tmpVal = rleStream.ReadByte()
                                    FillMemory VarPtr(.ciDataDecoded(bytesReadTotal)), rleCode, tmpVal
                                    bytesReadTotal = bytesReadTotal + rleCode
                                    bytesReadLine = bytesReadLine + 2
                                
                                '-128 means do nothing (nop).  I'm not sure why an encoder would set this, but if it does,
                                ' we'll simply move to the next byte.
                                'Else
                                    
                                End If
                            
                            'Technically, a malformed PSD could cause issues here, as a mis-reported RLE length
                            ' toward the end of the file could cause us to overrun the destination buffer.
                            ' Checks could be added both here, and on the above RLE checks (as both of those
                            ' could also overrun a destination line) for additional safety, at some cost to performance.
                            Loop While bytesReadLine < lineLengths(y)
                            
                        Next y
                        
                        'Manually free our temporary stream object, as it unsafely wraps a local array
                        Set rleStream = Nothing
                        DecodeChannels = psd_Success
                        
                    Case ccZip
                        DecodeChannels = psd_Failure
                        Set m_LayerDIB = Nothing
                    
                    Case ccZipWithPrediction
                        DecodeChannels = psd_Failure
                        Set m_LayerDIB = Nothing
                
                End Select
                
                'The source scanline has been successfully decoded.  "Paint" its contents into the destination DIB.
                If (Not m_LayerDIB Is Nothing) Then
                
                    'If this buffer is uncompressed, we can load it directly into our target DIB, depending on
                    ' the image's color-depth.  (Because 8-bpc is so prevalent, it's worth accelerating manually.)
                    If (bytesPerChannel = 1) Then
                        
                        Debug.Print "Transferring channel #" & .ciID & " to DIB..."
                        
                        m_LayerDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
                        
                        For y = 0 To layerHeight - 1
                            tmpSA.pvData = scanStart + (scanWidth * y)
                        For x = 0 To layerWidth - 1
                            If (.ciID = cidBlue) Then
                                tmpBytes(x * 4) = .ciDataDecoded(y * layerWidth + x)
                            ElseIf (.ciID = cidGreen) Then
                                tmpBytes(x * 4 + 1) = .ciDataDecoded(y * layerWidth + x)
                            ElseIf (.ciID = cidRed) Then
                                tmpBytes(x * 4 + 2) = .ciDataDecoded(y * layerWidth + x)
                            ElseIf (.ciID = cidAlpha) Then
                                tmpBytes(x * 4 + 3) = .ciDataDecoded(y * layerWidth + x)
                            End If
                        Next x
                        Next y
                        
                        m_LayerDIB.UnwrapArrayFromDIB tmpBytes
                        
                    Else
                        warningStack.AddString "BPC is currently unsupported; abandoning decode."
                        DecodeChannels = psd_Failure
                    End If
                    
                End If
                
            'Raw data is zero-length; this layer does not contain usable data
            Else
                Debug.Print "erasing layer DIB"
                Set m_LayerDIB = Nothing
            End If
            
        End With
    
SkipChannel:
    Next i
    
    'If a full layer image was constructed successfully, premultiply it before exiting
    If (Not m_LayerDIB Is Nothing) Then m_LayerDIB.SetAlphaPremultiplication True
    
    Exit Function
    
InternalVBError:
    InternalError "DecodeChannels", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in DecodeChannels, #" & Err.Number & ": " & Err.Description
    DecodeChannels = psd_Failure
End Function

'After *all* layer headers have been successfully parsed, the next chunk of the PSD contains image channel data.
' Pass the same source stream to this function to load all channel data for this layer.
Friend Function LoadChannels(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode) As PD_PSDResult

    On Error GoTo InternalVBError
    
    LoadChannels = psd_Success
    
    'Channel data is stored in planar format (e.g. not interleaved - RRRGGGBBB instead of RGBRGBRGB).
    ' As such, each channel must be loaded independent of other channels in the image.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        With m_Channels(i)
            .ciCompression = srcStream.ReadInt_BE()
            
            'The size of this chunk is "size of pixel data + size of compression flag".
            ' To avoid confusion, calculate a bare "size of pixel data" by subtracting two from
            ' the net chunk size.
            .ciDataLength = .ciSizeL - 2
            Debug.Print "loading channel of length " & .ciDataLength
            'It *is* possible to have zero-length layer data (for example, adjustment or effect layers
            ' may do this).  As such, it's critical to check channel size before attempting to read.
            If (.ciDataLength > 0) Then srcStream.ReadBytes .ciData, .ciDataLength, True
        End With
    Next i
    
    Exit Function
    
InternalVBError:
    InternalError "LoadChannels", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in LoadChannels, #" & Err.Number & ": " & Err.Description
    LoadChannels = psd_Failure
    
End Function

'Single-layer images in Photoshop have special rules.  This function is only called if a PSD is single-layer.
Friend Function NotifySingleLayerImage(ByRef srcStream As pdStream, ByVal imgNumChannels As Integer, ByVal imgWidthPx As Long, ByVal imgHeightPx As Long, ByVal imgBitsPerChannel As Integer, ByVal imgColorMode As PSD_ColorMode) As PD_PSDResult
    
    'Set default parameters
    With m_Rect
        .Left = 0
        .Top = 0
        .Right = imgWidthPx
        .Bottom = imgHeightPx
    End With
    
    m_Opacity = 255
    m_Clipping = 0
    m_Visible = True
    m_LayerNameANSI = "Background"      'TODO: figure out what this should *actually* be
    
    'Set channel data to match the parent image
    m_ChannelCount = imgNumChannels
    ReDim m_Channels(0 To m_ChannelCount - 1) As PSD_ChannelInfo
    
    'Calculate a scanline size for the current bit-depth; we need this to know how to separate the source data
    ' into usable per-channel chunks
    Dim slSize As Long
    slSize = CalculateBytesPerRow(imgBitsPerChannel, imgWidthPx)
    
    Dim i As Long
    
    'Retrieve compression method; this is identical to the compression IDs used in multilayer images
    Dim cmpMethod As PSD_ChannelCompression
    cmpMethod = srcStream.ReadInt_BE()
    
    Select Case cmpMethod
    
        Case ccRaw
            
            'In a raw-encoded image, channel size is fixed and easily calculated.
            Dim fixedChSize As Long
            fixedChSize = slSize * imgHeightPx
            
            For i = 0 To m_ChannelCount - 1
                With m_Channels(i)
                    .ciCompression = cmpMethod
                    .ciDataLength = fixedChSize
                    'Channels appear in sequential order, per the spec:
                    ' "Image data is stored in planar order: first all the red data, then all the green data, etc."
                    .ciID = i
                    srcStream.ReadBytes .ciData, fixedChSize, True
                End With
            Next i
        
        Case ccPackBits
        
            'In an RLE-encoded image, channel size cannot be pre-calculated.  It must instead be calculated
            ' dynamically using the RLE headers in the image data chunk.
            
            'The image data is prefaced by RLE scanline counts for *all* channels.  Load these into a
            ' standalone array.  (Note that this handling is a little messy on account of VB not having
            ' an unsigned-int type.  To simplify the math involved, we copy the unsigned ints into a
            ' long array, then reset the stream pointer - this allows us to just "copy" the RLE headers
            ' into each channel stream, and re-use our normal multi-layer RLE code to process it further.)
            Dim ptrBackup As Long
            ptrBackup = srcStream.GetPosition()
            
            Dim rleTable() As Long
            ReDim rleTable(0 To m_ChannelCount * imgHeightPx - 1) As Long
            For i = 0 To UBound(rleTable)
                rleTable(i) = srcStream.ReadIntUnsigned_BE()
            Next i
            
            'Reset the stream pointer
            srcStream.SetPosition ptrBackup, FILE_BEGIN
            
            'From the RLE table counts, above, we can calculate a buffer size for each channel
            Dim j As Long, chSizeSum As Long
            For i = 0 To m_ChannelCount - 1
                
                chSizeSum = 0
                For j = 0 To imgHeightPx - 1
                    chSizeSum = chSizeSum + rleTable(i * imgHeightPx + j)
                Next j
                
                'The required size for each channel is (rleTableForChannel + rleEncodedBytes)
                m_Channels(i).ciDataLength = (imgHeightPx * 2) + chSizeSum
                
            Next i
            
            'Now that we know how large each channel's raw data is, we can copy the data into place.
            For i = 0 To m_ChannelCount - 1
                With m_Channels(i)
                    .ciCompression = cmpMethod
                    'Channels appear in sequential order, per the spec:
                    ' "Image data is stored in planar order: first all the red data, then all the green data, etc."
                    .ciID = i
                    ReDim .ciData(0 To .ciDataLength - 1) As Byte
                    srcStream.ReadBytesToBarePointer VarPtr(.ciData(0)), (imgHeightPx * 2)
                End With
            Next i
            
            'Do one last loop to copy actual RLE bytes into place
            For i = 0 To m_ChannelCount - 1
                With m_Channels(i)
                    srcStream.ReadBytesToBarePointer VarPtr(.ciData(imgHeightPx * 2)), (.ciDataLength - (imgHeightPx * 2))
                End With
            Next i
            
        'TODO!
        Case ccZip
        Case ccZipWithPrediction
        
    End Select
    
    'Channel size is not known in advance, alas!  Instead, we have to decode *all* data in the image,
    ' inferring channel sizes from other factors.  This makes decompression noticeably more obnoxious compared to
    ' multi-layer images (backward, I know).
    
    
    
'Channel count can vary on a per-layer basis
'Private m_ChannelCount

End Function

'Given a stream object *THAT ALREADY POINTS TO A LAYER RECORD*, parse out whatever meaningful
' layer data we can.  You must also pass the parent class's warning stack - this class may add additional
' warnings to it.  PSD/PSB status is also required, as it affects the size of certain markers.
Friend Function ParseLayer(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean) As PD_PSDResult

    If (srcStream Is Nothing) Or (warningStack Is Nothing) Then
        ParseLayer = psd_Failure
        Exit Function
    End If
    
    On Error GoTo InternalVBError
    
    ParseLayer = psd_Success
    
    'The first descriptor in a layer record is a rectangle of the layer's coordinates.
    ' Per the spec, "Specified as top, left, bottom, right coordinates"
    With m_Rect
        .Top = srcStream.ReadLong_BE()
        .Left = srcStream.ReadLong_BE()
        .Bottom = srcStream.ReadLong_BE()
        .Right = srcStream.ReadLong_BE()
    End With
    
    'Number of channels in this layer
    m_ChannelCount = srcStream.ReadInt_BE()
    ReDim m_Channels(0 To m_ChannelCount - 1) As PSD_ChannelInfo
    
    'Each channel has its own channel ID and length indicator
    Dim i As Long
    For i = 0 To m_ChannelCount - 1
        m_Channels(i).ciID = srcStream.ReadInt_BE()
        m_Channels(i).ciSizeL = srcStream.ReadLong_BE()
        If imageIsPSB Then m_Channels(i).ciSizeH = srcStream.ReadLong_BE  'TODO for PSB format!
    Next i
    
    'After the channel descriptors comes a blend mode signature, which is great for confirming
    ' that our parser is still aligned correctly.
    Dim validateBM As String
    validateBM = srcStream.ReadString_ASCII(4)
    If (validateBM <> "8BIM") Then
        InternalError "ParseLayer", "Blend mode marker not present; parser is mis-aligned: " & validateBM
        ParseLayer = psd_Failure
        Exit Function
    End If
    
    'Next comes the actual blend mode; this is a 4-byte ascii identifier
    Dim bmCode As String
    bmCode = srcStream.ReadString_ASCII(4)
    
    Select Case bmCode
        Case "pass" ' = pass through
            m_BlendMode = BL_NORMAL 'TODO
        Case "norm" ' = normal
            m_BlendMode = BL_NORMAL
        Case "diss" ' = dissolve
            m_BlendMode = BL_NORMAL 'TODO
        Case "dark" ' = darken
            m_BlendMode = BL_DARKEN
        Case "mul " ' = multiply
            m_BlendMode = BL_MULTIPLY
        Case "idiv" ' = color burn
            m_BlendMode = BL_COLORBURN
        Case "lbrn" ' = linear burn
            m_BlendMode = BL_LINEARBURN
        Case "dkCl" ' = darker color
            m_BlendMode = BL_DARKEN 'TODO
        Case "lite" ' = lighten
            m_BlendMode = BL_LIGHTEN
        Case "scrn" ' = screen
            m_BlendMode = BL_SCREEN
        Case "div " ' = color dodge
            m_BlendMode = BL_COLORDODGE
        Case "lddg" ' = linear dodge
            m_BlendMode = BL_LINEARDODGE
        Case "lgCl" ' = lighter color
            m_BlendMode = BL_LIGHTEN    'TODO
        Case "over" ' = overlay
            m_BlendMode = BL_OVERLAY
        Case "sLit" ' = soft light
            m_BlendMode = BL_SOFTLIGHT
        Case "hLit" ' = hard light
            m_BlendMode = BL_HARDLIGHT
        Case "vLit" ' = vivid light
            m_BlendMode = BL_VIVIDLIGHT
        Case "lLit" ' = linear light
            m_BlendMode = BL_LINEARLIGHT
        Case "pLit" ' = pin light
            m_BlendMode = BL_PINLIGHT
        Case "hMix" ' = hard mix
            m_BlendMode = BL_HARDMIX
        Case "diff" ' = difference
            m_BlendMode = BL_DIFFERENCE
        Case "smud" ' = exclusion
            m_BlendMode = BL_EXCLUSION
        Case "fsub" ' = subtract
            m_BlendMode = BL_SUBTRACT
        Case "fdiv" ' = divide
            m_BlendMode = BL_DIVIDE
        Case "hue " ' = hue
            m_BlendMode = BL_HUE
        Case "sat " ' = saturation
            m_BlendMode = BL_SATURATION
        Case "colr" ' = color
            m_BlendMode = BL_COLOR
        Case "lum " ' = luminosity
            m_BlendMode = BL_LUMINOSITY
        Case Else
            warningStack.AddString "ParseLayer encountered an unknown blend mode code: " & bmCode
            m_BlendMode = BL_NORMAL
            ParseLayer = psd_Warning
    
    End Select
    
    'Next comes opacity (1-byte)
    m_Opacity = srcStream.ReadByte()
    
    'Clipping (0 = base, 1 = non-base)
    m_Clipping = srcStream.ReadByte()
    
    'Misc flags
    m_Flags = srcStream.ReadByte()
    
    'Visibility is stored as bit 1 in the flag byte we just retrieved.
    ' Note that the spec doesn't state this correctly; the bit actually stores INVISIBILITY state
    ' (e.g. if the bit is set, the layer is INVISIBLE).
    m_Visible = ((m_Flags And 2) = 0)
    
    'Filler byte (for alignment)
    srcStream.SetPosition 1, FILE_CURRENT
    
    'And finally, the length of "extra data field", or as the spec clarifies,
    ' "the total length of the next five fields".  I don't know what those five fields are,
    ' actually, as the spec only shows three:
    ' 1) Layer mask data
    ' 2) Layer blending range data
    ' 3) Layer name
    ' 4) Additional layer information
    ' 5) ???
    '
    'We grab this value anyway, as it's helpful for confirming pointer synchronization before exiting.
    m_LenExtraData = srcStream.ReadLong_BE()
    
    Dim finalPointerPos As Long
    finalPointerPos = srcStream.GetPosition() + m_LenExtraData
    
    'Next comes layer mask / adjustment layer data.  This is parsed separately.
    ParseLayer = ParseLayerMaskAdjustmentData(srcStream, warningStack, imageIsPSB)
    
    'Next, layer blending ranges data.  No idea what this does, but we still perform
    ' a correct parse.
    If (ParseLayer < psd_Failure) Then ParseLayer = ParseLayerMaskAdjustmentData(srcStream, warningStack, imageIsPSB)
    
    If (ParseLayer < psd_Failure) Then
    
        'Next, layer name.  This is a legacy Pascal ShortString, "padded to a multiple of 4 bytes".
        ' (Modern PSD files contain a separate tagged block with a Unicode copy of this string.)
        ' Pascal ShortStrings have a one-byte length value, followed by a string of chars
        ' (current system codepage) of length 0-255.
        Dim layerNameLength As Byte
        layerNameLength = srcStream.ReadByte()
        If (layerNameLength = 0) Then
            srcStream.SetPosition 3, FILE_CURRENT   'Forcibly advance the pointer by 3 more bytes
        Else
            
            'Retrieve the ANSI string, and advance by (n) additional bytes to mantain 4-byte alignment
            m_LayerNameANSI = srcStream.ReadString_ASCII(layerNameLength)
            
            Dim reqPadding As Long
            reqPadding = ((layerNameLength + 1) Mod 4)
            If (reqPadding <> 0) Then srcStream.SetPosition 4 - reqPadding, FILE_CURRENT
            Debug.Print "layer name: " & m_LayerNameANSI
            
        End If
        
    End If
    
    'Since Photoshop 4.0, a whole bunch of additional (optional) layer feature blocks have been defined.
    ' These follow the end of the layer records structure, and their existence can only be caught by
    ' looking for a mismatch between the current stream pointer and the expected "end of segment"
    ' calculated earlier.  These blocks use a PNG-like "chunk" system where each chunk guarantees a
    ' 4-char signature, 4-char ID code, length, then some variable amount of data.
    If (srcStream.GetPosition() < finalPointerPos) Then
        
        ReDim m_AdditionalInfo(0 To 3) As PSD_AdditionalInfo
        m_AdditionalInfoCount = 0
        
        Dim sigCheck As String
        
        Do While (srcStream.GetPosition() < finalPointerPos)
            
            'Verify the signature
            sigCheck = srcStream.ReadString_ASCII(4)
            
            If (sigCheck <> "8BIM") And (sigCheck <> "8B64") Then
                Debug.Print "Unknown signature: " & sigCheck
                warningStack.AddString "ParseLayer found an unknown additional segment signature: " & sigCheck
                ParseLayer = psd_Warning
            Else
                
                If (m_AdditionalInfoCount > UBound(m_AdditionalInfo)) Then ReDim Preserve m_AdditionalInfo(0 To m_AdditionalInfoCount * 2 - 1) As PSD_AdditionalInfo
                With m_AdditionalInfo(m_AdditionalInfoCount)
                    .aiKey = srcStream.ReadString_ASCII(4)
                    .aiDataLength = srcStream.ReadLong_BE()
                    'Debug.Print "Additional layer info found: " & .Key & " - " & .DataLength & " bytes"
                    If (.aiDataLength > 0) Then srcStream.ReadBytes .aiDataBytes, .aiDataLength, True
                End With
                
                m_AdditionalInfoCount = m_AdditionalInfoCount + 1
                
            End If
            
        Loop
        
        'Debug.Print "Found " & m_AdditionalInfoCount & " additional layer info records"
    
    Else
        'Debug.Print "No additional layer info records."
    End If
    
    'TODO: do something useful with additional info chunks that are relevant to PD features
    
    'Failsafe check to ensure our stream pointer is still aligned
    If (srcStream.GetPosition() <> finalPointerPos) Then
        warningStack.AddString "ParseLayer() may have misaligned stream pointer!"
        ParseLayer = psd_Failure
    End If
    
    Exit Function

InternalVBError:
    InternalError "ParseLayer", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ParseLayer, #" & Err.Number & ": " & Err.Description
    ParseLayer = psd_Failure
    
End Function

'Parse the layer blending range data section.  PD does not currently use this data.
Private Function ParseLayerBlendingRanges(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean) As PD_PSDResult
    
    On Error GoTo InternalVBError
    
    ParseLayerBlendingRanges = psd_Success
    
    'This segment provides its own length check.  This value can be zero (I think?), in which case
    ' the segment is skippable.
    Dim segmentLength As Long
    segmentLength = srcStream.ReadLong_BE()
    If (segmentLength = 0) Then
        ParseLayerBlendingRanges = psd_Success
        Exit Function
    End If
    
    'I don't actually know what layer blending ranges do, but the parser does correctly retrieve them
    ' (in case we find a use for them in the future).
     
    'Composite gray blend source. Contains 2 black values followed by 2 white values.
    ' Present but irrelevant for Lab & Grayscale.
    srcStream.ReadLong_BE
    
    'Composite gray blend destination range
    srcStream.ReadLong_BE
    
    'The remaining segment bytes consist of...
    ' (4 bytes) Nth channel source range
    ' (4 bytes) Nth channel destination range
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        srcStream.ReadLong_BE
        srcStream.ReadLong_BE
    Next i
    
    'Before exiting, do a failsafe check for pointer correctness
    Dim expectedFinalPointer As Long
    expectedFinalPointer = srcStream.GetPosition() + segmentLength
    
    'Perform a failsafe check for stream pointer correctness
    If (expectedFinalPointer <> srcStream.GetPosition()) Then
        warningStack.AddString "ParseLayerBlendingRanges() may have misaligned stream pointer!"
        ParseLayerBlendingRanges = psd_Warning
    End If
    
    Exit Function

InternalVBError:
    InternalError "ParseLayerBlendingRanges", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ParseLayerBlendingRanges, #" & Err.Number & ": " & Err.Description
    ParseLayerBlendingRanges = psd_Failure
    
End Function

'Parse the layer mask / adjustment layer data section.  PD will make better use of this data in the future,
' when these features are finally implemented locally.
Private Function ParseLayerMaskAdjustmentData(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean) As PD_PSDResult
    
    On Error GoTo InternalVBError
    
    ParseLayerMaskAdjustmentData = psd_Success
    
    'This segment provides its own length check.  This value can be zero, in which case the segment
    ' is skippable.
    Dim segmentLength As Long
    segmentLength = srcStream.ReadLong_BE()
    If (segmentLength = 0) Then
        ParseLayerMaskAdjustmentData = psd_Success
        Exit Function
    End If
    
    'Before exiting, we'll do a failsafe check for pointer correctness
    Dim expectedFinalPointer As Long
    expectedFinalPointer = srcStream.GetPosition() + segmentLength
    
    'Per the spec, this chunk "...can be 40 bytes, 24 bytes, or 4 bytes if no layer mask."
    ' I had previously attempted to validate this, but the spec is clearly wrong, as modern
    ' PSDs can have larger sizes than those reported here.  As such, we no longer attempt to
    ' validate the segment length; if it extends beyond what we know how to read, we'll forcibly
    ' reset the pointer at the end of this segment.
    
    'Layer mask rect, in "top, left, bottom, right" order
    With m_MaskRect
        .Top = srcStream.ReadLong_BE()
        .Left = srcStream.ReadLong_BE()
        .Bottom = srcStream.ReadLong_BE()
        .Right = srcStream.ReadLong_BE()
    End With
    
    'Remaining values may not be used in current build; where relevant, we just read and dump 'em
    ' to keep the parser properly aligned.
    
    'Default color (can only be 0 or 255; not current validated)
    srcStream.ReadByte
    
    'Flags
    ' bit 0 = position relative to layer
    ' bit 1 = layer mask disabled
    ' bit 2 = invert layer mask when blending (Obsolete)
    ' bit 3 = indicates that the user mask actually came from rendering other data
    ' bit 4 = indicates that the user and/or vector masks have parameters applied to them
    Dim maskFlags As Byte
    maskFlags = srcStream.ReadByte()
    
    'Per the spec, "[2 bytes] Padding. Only present if size = 20"
    If (segmentLength = 20) Then
        srcStream.SetPosition 2, FILE_CURRENT
        
    'When size=36, the padding is instead replaced by usable data
    Else
        
        'Real Flags. Same as Flags information above.
        srcStream.ReadByte
        
        'Real user mask background.  Must be 0 or 255 (not currently validated).
        srcStream.ReadByte
        
         'Rectangle enclosing the real user mask
        With m_RealMaskRect
            .Top = srcStream.ReadLong_BE()
            .Left = srcStream.ReadLong_BE()
            .Bottom = srcStream.ReadLong_BE()
            .Right = srcStream.ReadLong_BE()
        End With
        
    End If
    
    'Modern PSDs must have additional data here that is not present in the spec, because I've received
    ' unexpected misalignment errors.  As such, it's critical to re-align the pointer (based on the
    ' segment size value given above) before exiting.
    If (expectedFinalPointer <> srcStream.GetPosition()) Then srcStream.SetPosition expectedFinalPointer, FILE_BEGIN
    
    Exit Function

InternalVBError:
    InternalError "ParseLayerMaskAdjustmentData", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ParseLayerMaskAdjustmentData, #" & Err.Number & ": " & Err.Description
    ParseLayerMaskAdjustmentData = psd_Failure
    
End Function

'Given a bit-depth and image width, figure out how "long" a scanline is for a single color channel.
Private Function CalculateBytesPerRow(ByVal channelBitDepth As Long, ByVal imgWidth As Long) As Long
    
    'Monochrome images are the obnoxious exception
    If (channelBitDepth = 1) Then
        CalculateBytesPerRow = (imgWidth + 7) \ 8
    Else
        CalculateBytesPerRow = imgWidth * (channelBitDepth \ 8)
    End If
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdPSDLayer." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdPSDLayer." & funcName & "() reported an error: " & errDescription
    End If
End Sub
