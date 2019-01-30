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
'Last updated: 30/January/19
'Last update: add support for 32-bit channels (all color modes!)
'
'This class contains layer-specific data pulled from a PSD file.  It is populated by a parent
' pdPSD instance.  It has no purpose outside of a PSD parsing context; for layer handling inside
' PhotoDemon itself, refer to the pdLayer class.
'
'Unless otherwise noted, all code in this class is my original work.  It is based off the "official" Adobe spec
' at this URL (link good as of January 2019):
' https://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577409_72092
'
'Adobe's custom "zip-with-prediction" compression mode is currently undocumented in the PSD spec.  To my knowledge,
' the first 3rd-party to reverse-engineer it is the authors of Paint.NET's PSD plugin.  Their MIT-licensed work was
' used as a reference for this class's "zip-with-prediction" decoder.  Many thanks to those authors; their reference
' implementation is available here (link good as of January 2019): https://github.com/PsdPlugin/PsdPlugin
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PSD files contain a *lot* of information.  To aid debugging, you can activate "verbose" output; this will
' dump all kinds of diagnostic information to the debug log.  (Note that other PSD classes have their own
' version of this constant.)
Private Const PSD_DEBUG_VERBOSE As Boolean = True

'Various rects; these have been translated to the proper order (Adobe uses top/left/bottom/right order)
Private m_Rect As RectL, m_MaskRect As RectL, m_RealMaskRect As RectL

'Channel count can vary on a per-layer basis
Private m_ChannelCount As Long

'Color mode typically mirrors the parent image's color mode
Private m_ColorMode As PSD_ColorMode

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
    cidIndexed = 0
    cidLabL = 0
    cidLabA = 1
    cidLabB = 2
End Enum

#If False Then
    Private Const cidRealUserLayerMask = -3, cidUserLayerMask = -2, cidAlpha = -1, cidRed = 0, cidGreen = 1, cidBlue = 2, cidCyan = 0, cidMagenta = 1, cidYellow = 2, cidKey = 3, cidIndexed = 0, cidLabL = 0, cidLabA = 1, cidLabB = 2
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
    
    'If PD can successfully decode a channel's data (this is not guaranteed until all compression
    ' formats are implemented), this value will be set to TRUE
    ciDecodedOK As Boolean
    ciUnusedPadding As Boolean
    
    'The interpretation of ciData() depends on the value of ciCompression (above). It is relatively common
    ' for PSD data to be stored as raw-encoded, in which case ciData() provides relevant raw pixel values.
    ' If the data has been compressed via PackBits or Deflate, however, ciData() needs to be decoded before
    ' you can get usable information from it.  The DecodeChannels() function handles this.
    ciData() As Byte
    
    'After a channel has been decoded, the new raw data will be placed here.
    ciDataDecoded() As Byte
    
End Type

Private m_Channels() As PSD_ChannelInfo

'Additional layer info is parsed and stored by a dedicated class; this class is queried only for blocks
' that we know how to interpret.
Private m_LayerInfo As pdPSDLayerInfo

'Photoshop provides some blend modes that PD does not.  We attempt to convert all supported blend modes
' to our own internal constant, but if we fail, we'll fall back to "normal" mode.
Private m_BlendMode As PD_BlendMode

Private m_Opacity As Byte, m_Clipping As Byte, m_Flags As Byte, m_Visible As Boolean, m_LenExtraData As Long
Private m_LayerNameANSI As String, m_LayerNameUnicode As String

'The final result of all channel decoding is this pdDIB object, ready for attachment to a pdLayer.
Private m_LayerDIB As pdDIB

'Multichannel images represent a variation on the above rule.  Instead of creating one DIB, they create several -
' one for each channel.  The user can then decide how to assemble these into a final image.  (These values are
' unused in any other color mode.)
Private m_FinalChannelDIBs() As pdDIB, m_FinalChannelIDs() As PSD_ChannelID, m_FinalChannelCount As Long

Friend Function DoesLayerHaveUsableData() As Boolean
    DoesLayerHaveUsableData = (m_ChannelCount > 0) And ((m_Rect.Right - m_Rect.Left) <> 0) And ((m_Rect.Bottom - m_Rect.Top) <> 0)
    If DoesLayerHaveUsableData Then
        If (m_ColorMode = cm_Multichannel) Then
            DoesLayerHaveUsableData = (m_FinalChannelCount > 0)
        Else
            DoesLayerHaveUsableData = (Not m_LayerDIB Is Nothing)
        End If
    End If
End Function

'Certain color types can result in multiple layers within a single pdPSDLayer object (multichannel images, most notably).
' Check this value before assuming that only one pdLayer should be constructed from this PSD layer object.
Friend Function GetLayerCount() As Long
    If (m_ColorMode = cm_Multichannel) Then
        GetLayerCount = m_FinalChannelCount
    Else
        GetLayerCount = 1
    End If
End Function

Friend Function GetLayerDIB() As pdDIB
    Set GetLayerDIB = m_LayerDIB
End Function

Friend Function GetLayerName() As String
    If (LenB(m_LayerNameUnicode) <> 0) Then GetLayerName = m_LayerNameUnicode Else GetLayerName = m_LayerNameANSI
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

Friend Function GetLayerWidth() As Long
    GetLayerWidth = m_Rect.Right - m_Rect.Left
End Function

Friend Function GetLayerHeight() As Long
    GetLayerHeight = m_Rect.Bottom - m_Rect.Top
End Function

Friend Function GetLayerVisibility() As Boolean
    GetLayerVisibility = m_Visible
End Function

Friend Function GetMultiChannelLayerDIB(ByVal lIndex As Long) As pdDIB
    Set GetMultiChannelLayerDIB = m_FinalChannelDIBs(lIndex)
End Function

Friend Function GetMultiChannelLayerID(ByVal lIndex As Long) As PSD_ChannelID
    GetMultiChannelLayerID = m_FinalChannelIDs(lIndex)
End Function

'After channel data has been decoded (decompressed), it needs to be assembled into a final usable image.
' At present, this step is limited to certain color modes and bit-depths, although I am actively expanding coverage.
' If you have a PSD file that doesn't work, this step is likely the culprit - attach the PSD file to a new Issue
' at GitHub and I will investigate further.
'
'(Also note: if the DecodeChannels step fails for a given channel, don't waste time calling this step; it won't
' work if the underlying channel data hasn't been decoded properly!)
Friend Function ConstructImage(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD) As PD_PSDResult

    ConstructImage = psd_Success
    
    On Error GoTo InternalVBError
    
    m_ColorMode = imgColorMode
    
    'AFAIK, all bit-depths and color modes are now fully covered by PhotoDemon's importer.
    'If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Bit-depth and color mode are supported; attempting to decode now..."
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Image construction varies greatly depending on the format of the input data.  To better handle the unique needs
    ' of each color format, handling is sorted by color mode.  (This allows for acceleration that would otherwise
    ' be impractical in a "universal" image constructor.)
    If (imgColorMode = cm_RGB) Then
        ConstructImage = ConstructImage_RGB(srcStream, warningStack, imageIsPSB, bitsPC, imgColorMode, srcProfile, parentPSD)
        
    ElseIf (imgColorMode = cm_CMYK) Then
        ConstructImage = ConstructImage_CMYK(srcStream, warningStack, imageIsPSB, bitsPC, imgColorMode, srcProfile, parentPSD)
        
    ElseIf (imgColorMode = cm_Grayscale) Or (imgColorMode = cm_Indexed) Or (imgColorMode = cm_Duotone) Or (imgColorMode = cm_Bitmap) Then
        ConstructImage = ConstructImage_SingleChannel(srcStream, warningStack, imageIsPSB, bitsPC, imgColorMode, srcProfile, parentPSD)
    
    ElseIf (imgColorMode = cm_Lab) Then
        ConstructImage = ConstructImage_Lab(srcStream, warningStack, imageIsPSB, bitsPC, imgColorMode, srcProfile, parentPSD)
        
    ElseIf (imgColorMode = cm_Multichannel) Then
        ConstructImage = ConstructImage_Multichannel(srcStream, warningStack, imageIsPSB, bitsPC, imgColorMode, srcProfile, parentPSD)
        
    End If
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Total image assembly took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ConstructImage", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ConstructImage, #" & Err.Number & ": " & Err.Description
    ConstructImage = psd_Failure
    
End Function

Private Function ConstructImage_CMYK(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD) As PD_PSDResult

    ConstructImage_CMYK = psd_Success
    On Error GoTo InternalVBError
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Interpretation of the underlying data depends on bit-depth, obviously
    Dim bytesPerChannel As Long
    bytesPerChannel = (bitsPC \ 8)
    
    'Set a flag to determine ICC profile availability; PhotoDemon won't (currently) load CMYK images without
    ' an embedded profile as the results are likely to be garbage.  Photoshop will *never* write CMYK images
    ' an ICC profile, but amateur 3rd-party software may attempt to write CMYK like this - and they shouldn't.
    Dim iccAvailable As Boolean
    iccAvailable = (Not srcProfile Is Nothing)
    If iccAvailable Then iccAvailable = srcProfile.HasICCData()
    If (Not iccAvailable) Then
        warningStack.AddString "CMYK layer does not contain a usable ICC profile.  Pixel data is junk; this file was written incorrectly."
        ConstructImage_CMYK = psd_Failure
        Exit Function
    End If
    
    'If there's no embedded alpha channel, we can skip certain processing steps; determine alpha status now
    Dim alphaPresent As Boolean
    alphaPresent = False
    
    'TODO: figure out if we want to account for layer masks here; we technically don't need to, as the mask
    ' gets assembled separately and "composited" over the top - but we *would* need to premultiply in that case,
    ' so maybe a separate tracker for layer masks specifically is a better idea?
    
    'Note that alpha channels can be explicitly marked (with an ID of -1) or they can simply be "extra" channels
    ' without an obvious RGB mapping.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(i).ciID = cidAlpha) Or (m_Channels(i).ciID > cidKey) Then alphaPresent = True
    Next i
    
    'Layer dimensions determine the size of the constructed layer, obviously; layer sizes are completely independent
    ' of the parent image.
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    
    'Determine the size of a single channel; we'll need this to calculate offsets into our planar assembly array
    Dim channelSize As Long
    channelSize = bytesPerChannel * layerWidth * layerHeight
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Construction setup took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    Dim x As Long, y As Long
    
    'To keep memory requirements low, we're going to process the image one scanline at a time.  This prevents
    ' us from needing to allocate another (potentially massive) image-sized buffer.
    Dim channelScanlineSize As Long
    channelScanlineSize = bytesPerChannel * layerWidth
    
    'To simplify handling if channels are missing, start with a white, opaque destination.
    Set m_LayerDIB = New pdDIB
    If (Not m_LayerDIB.CreateBlank(layerWidth, layerHeight, 32, vbWhite, 255)) Then
        InternalError "ConstructImage_CMYK", "Couldn't allocate image buffer; image is likely too large."
        ConstructImage_CMYK = psd_Failure
        Exit Function
    End If
    
    Dim scanStart As Long, scanWidth As Long
    scanStart = m_LayerDIB.GetDIBPointer()
    scanWidth = m_LayerDIB.GetDIBStride()
    
    'Figure out how many source channels we have and where their indices lie.  (Channel order isn't
    ' necessarily guaranteed in PSDs.)
    Dim cIndex As Long, mIndex As Long, yIndex As Long, kIndex As Long, aIndex As Long
    cIndex = -1: mIndex = -1: yIndex = -1: kIndex = -1: aIndex = -1
    
    For x = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(x).ciID = cidCyan) Then
            cIndex = x
        ElseIf (m_Channels(x).ciID = cidMagenta) Then
            mIndex = x
        ElseIf (m_Channels(x).ciID = cidYellow) Then
            yIndex = x
        ElseIf (m_Channels(x).ciID = cidKey) Then
            kIndex = x
        ElseIf (m_Channels(x).ciID = cidAlpha) Or (m_Channels(x).ciID > cidKey) Then
            aIndex = x
        End If
    Next x
    
    'Prep a temporary buffer.  We'll assemble each scanline into a contiguous planar chunk in this buffer,
    ' then color-manage, interleave, swizzle, and (when necessary) swap endianness while merging the data
    ' into the final image buffer.
    Dim tmpPlanarBuffer() As Byte, srcLineSize As Long
    If alphaPresent Then srcLineSize = channelScanlineSize * 5 Else srcLineSize = channelScanlineSize * 4
    ReDim tmpPlanarBuffer(0 To srcLineSize - 1) As Byte
    
    'Prep source and destination ICC profiles
    Dim lcmsSrcProfile As pdLCMSProfile, lcmsDstProfile As pdLCMSProfile, lcmsTransform As pdLCMSTransform
    Set lcmsSrcProfile = New pdLCMSProfile
    lcmsSrcProfile.CreateFromPDICCObject srcProfile
    
    Set lcmsDstProfile = New pdLCMSProfile
    lcmsDstProfile.CreateSRGBProfile
    
    'Preparing the transform itself is a little more complicated, as it varies based on source color depth
    ' and the presence of alpha channels.  (Note that we also explicitly request blackpoint compensation to
    ' produce results similar to Adobe's CMM.)
    Dim srcColorMode As LCMS_PIXEL_FORMAT
    If (bytesPerChannel = 1) Then
        srcColorMode = TYPE_CMYK_8_PLANAR
    ElseIf (bytesPerChannel = 2) Then
        srcColorMode = TYPE_CMYK_16_PLANAR Or FLAG_SE
    ElseIf (bytesPerChannel = 4) Then
        srcColorMode = TYPE_CMYK_FLT Or FLAG_PLANAR
    End If
    If alphaPresent Then srcColorMode = srcColorMode Or FLAG_ALPHAPRESENT
    
    'Photoshop defines CMYK data as vanilla (e.g. 0 is white, 1 is black)
    srcColorMode = srcColorMode Or FLAG_MINISWHITE
    
    Set lcmsTransform = New pdLCMSTransform
    lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile, srcColorMode, TYPE_BGRA_8, lcmsSrcProfile.GetRenderingIntent(), cmsFLAGS_BLACKPOINTCOMPENSATION Or cmsFLAGS_COPY_ALPHA
    
    'Iterate through each scanline in the image
    For y = 0 To layerHeight - 1
    
        'Copy all present channels into the contiguous planar buffer
        If (cIndex >= 0) Then CopyMemoryStrict VarPtr(tmpPlanarBuffer(0)), VarPtr(m_Channels(cIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
        If (mIndex >= 0) Then CopyMemoryStrict VarPtr(tmpPlanarBuffer(channelScanlineSize)), VarPtr(m_Channels(mIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
        If (yIndex >= 0) Then CopyMemoryStrict VarPtr(tmpPlanarBuffer(channelScanlineSize * 2)), VarPtr(m_Channels(yIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
        If (kIndex >= 0) Then CopyMemoryStrict VarPtr(tmpPlanarBuffer(channelScanlineSize * 3)), VarPtr(m_Channels(kIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
        If (aIndex >= 0) Then CopyMemoryStrict VarPtr(tmpPlanarBuffer(channelScanlineSize * 4)), VarPtr(m_Channels(aIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
        
        'Use LCMS to transform the data into our final RGBA buffer
        lcmsTransform.ApplyTransformToArbitraryMemoryEx VarPtr(tmpPlanarBuffer(0)), scanStart + y * scanWidth, layerWidth, 1, srcLineSize, scanWidth, channelScanlineSize, 0&
        
    Next y
    
    'Free all source channel data to relieve memory pressure
    For x = 0 To Abs(m_ChannelCount) - 1
        Erase m_Channels(x).ciDataDecoded
    Next x
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Image assembly took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'If meaningful alpha data exists in the file, premultiply it now
    If alphaPresent Then
        m_LayerDIB.SetAlphaPremultiplication True
        If PSD_DEBUG_VERBOSE Then
            PDDebug.LogAction "Alpha premultiplication took: " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
    Else
        m_LayerDIB.SetInitialAlphaPremultiplicationState True
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ConstructImage_CMYK", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ConstructImage_CMYK, #" & Err.Number & ": " & Err.Description
    ConstructImage_CMYK = psd_Failure
    
End Function

Private Function ConstructImage_Lab(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD) As PD_PSDResult

    ConstructImage_Lab = psd_Success
    On Error GoTo InternalVBError
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Interpretation of the underlying data depends on bit-depth, obviously
    Dim bytesPerChannel As Long
    bytesPerChannel = (bitsPC \ 8)
    
    'Set a flag to determine ICC profile availability; this affects a *lot* of subsequent work in this function,
    ' so we need to reference it early (and often).
    Dim iccAvailable As Boolean
    iccAvailable = (Not srcProfile Is Nothing)
    If iccAvailable Then iccAvailable = srcProfile.HasICCData()
    
    'If there's no embedded alpha channel, we can skip certain processing steps; determine alpha status now
    Dim alphaPresent As Boolean
    alphaPresent = False
    
    'TODO: figure out if we want to account for layer masks here; we technically don't need to, as the mask
    ' gets assembled separately and "composited" over the top - but we *would* need to premultiply in that case,
    ' so maybe a separate tracker for layer masks specifically is a better idea?
    
    'Note that alpha channels can be explicitly marked (with an ID of -1) or they can simply be "extra" channels
    ' without an obvious RGB mapping.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(i).ciID = cidAlpha) Or (m_Channels(i).ciID > cidBlue) Then alphaPresent = True
    Next i
    
    'Unlike other color modes, Lab images are tricky because an accurate reconstruction depends on full
    ' color-management applied to the raw channel data.  LittleCMS needs channel data to be contiguous
    ' before it can process it into RGB, or we have to do something low-quality like ignore the ICC data
    ' and manually transform the channels ourselves into RGB mode (which results in a lower quality image).
    '
    'Assembling all channels in the image into a contiguous chunk would require a large block of memory,
    ' so we don't want to do that.  Instead, what we're going to do is assemble one scanline of Lab data
    ' at a time - into a single contiguous planar block - then color-manage that block into an RGBA
    ' buffer via LittleCMS.  This gives us the best combination of proper HDR handling without ballooning
    ' RAM requirements into problematic areas (for a 32-bit app, anyway).
    '
    'Because Photoshop doesn't guarantee a certain channel order in the file, we need to explicitly map
    ' L/a/b channel indices in advance.  (I don't know that Lab files can exist without all three channels
    ' being present, but just in case, we cover that possibility.)
    Dim lIndex As Long, aIndex As Long, bIndex As Long, channelsFound As Long
    lIndex = -1: aIndex = -1: bIndex = -1: channelsFound = 0
    For i = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(i).ciID = cidLabL) Then
            lIndex = i
            channelsFound = channelsFound + 1
        ElseIf (m_Channels(i).ciID = cidLabA) Then
            aIndex = i
            channelsFound = channelsFound + 1
        ElseIf (m_Channels(i).ciID = cidLabB) Then
            bIndex = i
            channelsFound = channelsFound + 1
        End If
    Next i
    
    If (channelsFound < 3) Then warningStack.AddString "Lab image is missing channels; PhotoDemon will try to assemble a usable image anyway."
    
    'Layer dimensions determine the size of the constructed layer, obviously; layer sizes are completely
    ' independent of the parent image.
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    
    'Determine the size of a single channel scanline; we'll need this to calculate offsets into our planar
    ' assembly array
    Dim channelScanlineSize As Long
    channelScanlineSize = bytesPerChannel * layerWidth
    
    'If alpha is present, we ignore it until all Lab data is processed first; this means we only need to
    ' allocate three contiguous channels worth of memory.
    Dim labBuffer() As Byte
    ReDim labBuffer(0 To channelScanlineSize * 3 - 1) As Byte
    
    'Prep the RGBA buffer in advance; it *does* need to be allocated into a single contiguous chunk.
    Set m_LayerDIB = New pdDIB
    If (Not m_LayerDIB.CreateBlank(layerWidth, layerHeight, 32, vbWhite, 255)) Then
        InternalError "ConstructImage_Lab", "Couldn't allocate image buffer; image is likely too large."
        ConstructImage_Lab = psd_Failure
        Exit Function
    End If
    
    Dim scanStart As Long, scanWidth As Long
    scanStart = m_LayerDIB.GetDIBPointer()
    scanWidth = m_LayerDIB.GetDIBStride()
    
    'Prep all required color management objects.  Note that special flags are set to mark the source data
    ' as planar instead of interleaved, and to specify transform parameters that more closely mimic the
    ' results of Adobe's CMM.
    Dim lcmsSrcProfile As pdLCMSProfile, lcmsDstProfile As pdLCMSProfile, lcmsTransform As pdLCMSTransform
    Set lcmsSrcProfile = New pdLCMSProfile
        
    If iccAvailable Then
        lcmsSrcProfile.CreateFromPDICCObject srcProfile
        
    'If the image came from Photoshop itself, it will always have an ICC profile attached.  Only garbage
    ' 3rd-party software would dare write Lab data without a profile.  Fortunately for us, LittleCMS can
    ' generate a proper Lab profile "on the fly" as required.
    Else
        lcmsSrcProfile.CreateLabProfile True
    End If
    
    Set lcmsDstProfile = New pdLCMSProfile
    lcmsDstProfile.CreateSRGBProfile
    
    Dim srcColorFormat As LCMS_PIXEL_FORMAT
    If (bytesPerChannel = 1) Then
        srcColorFormat = TYPE_Lab_8
    ElseIf (bytesPerChannel = 2) Then
        srcColorFormat = TYPE_Lab_16
    ElseIf (bytesPerChannel = 4) Then
        srcColorFormat = TYPE_Lab_FLT
    End If
    
    'Mark the data as planar, and if it's 2-bytes per channel, mark it as big-endian
    srcColorFormat = srcColorFormat Or FLAG_PLANAR
    If (bytesPerChannel = 2) Then srcColorFormat = srcColorFormat Or FLAG_SE
    
    'Explicitly request blackpoint compensation to produce results similar to Adobe's CMM.  (This shouldn't matter
    ' as the source profile should have perceptual rendering intent set, but for images "in the wild", who knows
    ' wtf we might find.)
    Set lcmsTransform = New pdLCMSTransform
    lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile, srcColorFormat, TYPE_BGRA_8, lcmsSrcProfile.GetRenderingIntent(), cmsFLAGS_BLACKPOINTCOMPENSATION
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Construction setup took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    Dim x As Long, y As Long
    
    'Instead of iterating through all channels in the image (as we do in other color modes), we're actually
    ' going to iterate through scanlines.
    For y = 0 To layerHeight - 1
    
        'Copy the source Lab data into our temporary contiguous buffer
        If (lIndex >= 0) Then CopyMemoryStrict VarPtr(labBuffer(0)), VarPtr(m_Channels(lIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
        If (aIndex >= 0) Then CopyMemoryStrict VarPtr(labBuffer(channelScanlineSize)), VarPtr(m_Channels(aIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
        If (bIndex >= 0) Then CopyMemoryStrict VarPtr(labBuffer(channelScanlineSize * 2)), VarPtr(m_Channels(bIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
        
        'Use lcms to color-manage the data into the RGBA buffer
        lcmsTransform.ApplyTransformToArbitraryMemoryEx VarPtr(labBuffer(0)), scanStart + scanWidth * y, layerWidth, 1, channelScanlineSize * 3, scanWidth, channelScanlineSize, 0&
        
    Next y
    
    'The lab data has been successfully assembled!  Immediately free all source channel arrays to reduce memory pressure.
    If (lIndex >= 0) Then Erase m_Channels(lIndex).ciDataDecoded
    If (aIndex >= 0) Then Erase m_Channels(aIndex).ciDataDecoded
    If (bIndex >= 0) Then Erase m_Channels(bIndex).ciDataDecoded
    
    'Our final step is to copy alpha data, if any, into the image buffer.
    If alphaPresent Then
        
        Dim tmpBytes() As Byte, tmpSA As SafeArray1D
    
        Dim chIndex As Long
        For chIndex = 0 To Abs(m_ChannelCount) - 1
            
            With m_Channels(chIndex)
                
                'There are a lot of reasons to ignore a given channel:
                ' 1) Channel data wasn't decoded successfully
                ' 2) Channel doesn't contribute pixel data to the final image (e.g. adjustment layers)
                ' 3) Layer masks, which have to be handled separately due to custom size and layout considerations
                If (Not .ciDecodedOK) Then GoTo SkipChannel
                If (.ciDataLength = 0) Then GoTo SkipChannel
                If (.ciID < cidAlpha) Then
                    warningStack.AddString "Found a layer mask, but PD can't interpret these yet - sorry."
                    GoTo SkipChannel
                End If
            
                'Make sure this is an alpha channel
                Dim cOffset As Long
                If (.ciID = cidAlpha) Or (.ciID > cidLabB) Then
                    
                    Dim ySrcOffset As Long, xLoopEnd As Long
                    m_LayerDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
                    xLoopEnd = layerWidth - 1
                    
                    '8-bit channel (most common case)
                    If (bytesPerChannel = 1) Then
                    
                        For y = 0 To layerHeight - 1
                            tmpSA.pvData = scanStart + (scanWidth * y)
                            ySrcOffset = y * layerWidth
                        For x = 0 To xLoopEnd
                            tmpBytes(x * 4 + 3) = .ciDataDecoded(ySrcOffset + x)
                        Next x
                        Next y
                        
                    '16-bit channel (take high-bit only, and note that the source data *is big-endian*
                    ElseIf (bytesPerChannel = 2) Then
                    
                        For y = 0 To layerHeight - 1
                            tmpSA.pvData = scanStart + (scanWidth * y)
                            ySrcOffset = y * layerWidth * 2
                        For x = 0 To xLoopEnd
                            tmpBytes(x * 4 + 3) = .ciDataDecoded(ySrcOffset + x * 2)
                        Next x
                        Next y
                    
                    '32-bit channel remains TODO
                    ElseIf (bytesPerChannel = 4) Then

                    End If
                    
                    'Manually free our unsafe array wrapper before proceeding, and immediately free our source data
                    ' to relieve memory pressure
                    m_LayerDIB.UnwrapArrayFromDIB tmpBytes
                    Erase .ciDataDecoded
            
                End If
                
            End With
            
SkipChannel:
        Next chIndex
    
    End If
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Lab transform took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'If meaningful alpha data exists in the file, premultiply it now
    If alphaPresent Then
        m_LayerDIB.SetAlphaPremultiplication True
        If PSD_DEBUG_VERBOSE Then
            PDDebug.LogAction "Alpha premultiplication took: " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
    Else
        m_LayerDIB.SetInitialAlphaPremultiplicationState True
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ConstructImage_Lab", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ConstructImage_Lab, #" & Err.Number & ": " & Err.Description
    ConstructImage_Lab = psd_Failure
    
End Function

'Multichannel images don't provide a clear mechanism for handling them (in an RGBA context, anyway).
' In PD, I follow convention of other parsers and load each channel to its own grayscale layer.
' The hope is that the user can rechannel each layer's contents according to their specialized knowledge,
' and ultimately obtain a usable image.  This approach means that channel data is handled similarly to
' grayscale channels, but each input layer will potentially produce *multiple* output layers.
Private Function ConstructImage_Multichannel(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD) As PD_PSDResult

    ConstructImage_Multichannel = psd_Success
    On Error GoTo InternalVBError
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Interpretation of the underlying data depends on bit-depth, obviously
    Dim bytesPerChannel As Long
    bytesPerChannel = (bitsPC \ 8)
    
    'Unlike other image constructors, this sub does *not* deal with ICC profiles.  I can revisit this if
    ' someone sends a multichannel file with a relevant embedded profile, but at present, I have no idea how
    ' to connect mode-less channel data to mode-specific ICC data.
    
    '(Same goes for alpha channels - their meaning is nebulous in a multichannel file, so we don't deal
    ' with them at present.  They'll be potentially easier to deal with once we have support for layer masks.)
    
    'Layer dimensions determine the size of the constructed layer, obviously; layer sizes are completely independent
    ' of the parent image.
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    
    'A unique target DIB will be created for each discovered channel.  Initialize our collection in advance.
    m_FinalChannelCount = 0
    ReDim m_FinalChannelDIBs(0) As pdDIB
    ReDim m_FinalChannelIDs(0) As PSD_ChannelID
        
    'Determine the size of a single channel; we'll need this to calculate offsets into our planar assembly array
    Dim channelSize As Long
    channelSize = bytesPerChannel * layerWidth * layerHeight
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Construction setup took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    Dim x As Long, y As Long
    Dim tmpBytes() As Byte, tmpLongs() As Long, tmpSA As SafeArray1D
    Dim scanStart As Long, scanWidth As Long, xLoopEnd As Long, ySrcOffset As Long
    scanStart = m_LayerDIB.GetDIBPointer()
    scanWidth = m_LayerDIB.GetDIBStride()
    
    'Layers will typically match the color-mode and bit-depth of the parent image, but it can be
    ' convenient to adjust these "on-the-fly" (e.g. bitmap/monochrome layers are inconvenient to
    ' render manually, so we translate them to a grayscale equivalent on the fly).  These parameters
    ' are reset as each channel is processed.
    Dim chnColorMode As PSD_ColorMode, chnBitDepth As Long
    
    'Prep a generic grayscale palette in advance, then copy it into a Long-type array that we can
    ' access more quickly inside loops.
    Dim imgPalette() As RGBQuad, palSize As Long, trnsIndex As Long
    ReDim imgPalette(0 To 255) As RGBQuad
    For x = 0 To 255
        imgPalette(x).Blue = x
        imgPalette(x).Green = x
        imgPalette(x).Red = x
        imgPalette(x).Alpha = 255
    Next x
    
    Dim palLookup() As Long
    ReDim palLookup(0 To 255) As Long
    For x = 0 To palSize - 1
        GetMem4 VarPtr(imgPalette(x)), palLookup(x)
    Next x
        
    'Iterate through all channels, copying their planar data into the interleaved final image buffer as we go
    Dim chIndex As Long
    For chIndex = 0 To Abs(m_ChannelCount) - 1
        
        'Reset any on-the-fly changes we've made to color mode or bit-depth
        chnColorMode = m_ColorMode
        chnBitDepth = bitsPC
        
        With m_Channels(chIndex)
            
            'There are a lot of reasons to ignore a given channel:
            ' 1) Channel data wasn't decoded successfully
            ' 2) Channel doesn't contribute pixel data to the final image (e.g. adjustment layers)
            ' 3) Layer masks, which have to be handled separately due to custom size and layout considerations
            If (Not .ciDecodedOK) Then GoTo SkipChannel
            If (.ciDataLength = 0) Then GoTo SkipChannel
            If (.ciID < cidAlpha) Then
                warningStack.AddString "Found a layer mask, but PD can't interpret these yet - sorry."
                GoTo SkipChannel
            End If
            
            'Prepare a target image for this channel
            ReDim Preserve m_FinalChannelDIBs(0 To m_FinalChannelCount) As pdDIB
            ReDim Preserve m_FinalChannelIDs(0 To m_FinalChannelCount) As PSD_ChannelID
            
            Set m_FinalChannelDIBs(m_FinalChannelCount) = New pdDIB
            If m_FinalChannelDIBs(m_FinalChannelCount).CreateBlank(layerWidth, layerHeight, 32, vbWhite, 255) Then
            
                m_FinalChannelIDs(m_FinalChannelCount) = .ciID
                scanStart = m_FinalChannelDIBs(m_FinalChannelCount).GetDIBPointer()
                scanWidth = m_FinalChannelDIBs(m_FinalChannelCount).GetDIBStride()
                
                'Use the grayscale palette to populate the target image
                m_FinalChannelDIBs(m_FinalChannelCount).WrapLongArrayAroundScanline tmpLongs, tmpSA, 0
                xLoopEnd = layerWidth - 1
                
                '8-bit channel
                If (bytesPerChannel = 1) Then
                    
                    For y = 0 To layerHeight - 1
                        tmpSA.pvData = scanStart + (scanWidth * y)
                        ySrcOffset = y * layerWidth
                    For x = 0 To xLoopEnd
                        tmpLongs(x) = palLookup(.ciDataDecoded(ySrcOffset + x))
                    Next x
                    Next y
                    
                '16-bit channel
                ElseIf (bytesPerChannel = 2) Then
                
                    For y = 0 To layerHeight - 1
                        tmpSA.pvData = scanStart + (scanWidth * y)
                        ySrcOffset = y * layerWidth * 2
                    For x = 0 To xLoopEnd
                        tmpLongs(x) = palLookup(.ciDataDecoded(ySrcOffset + x * 2))
                    Next x
                    Next y
                    
                '32-bit channel
                Else
                
                    Dim tmpSingle As Single
                    For y = 0 To layerHeight - 1
                        tmpSA.pvData = scanStart + (scanWidth * y)
                        ySrcOffset = y * layerWidth * 4
                    For x = 0 To xLoopEnd
                    
                        'Extract the source data into a single-type value
                        GetMem4_Ptr VarPtr(.ciDataDecoded(ySrcOffset + x * 4)), VarPtr(tmpSingle)
                        
                        'Clamp for safety before assigning
                        If (tmpSingle < 0!) Then tmpSingle = 0!
                        If (tmpSingle > 1!) Then tmpSingle = 1!
                        tmpLongs(x) = palLookup(Int(tmpSingle * 255!))
                        
                    Next x
                    Next y
                    
                End If
                
                m_FinalChannelDIBs(m_FinalChannelCount).UnwrapLongArrayFromDIB tmpLongs
                m_FinalChannelDIBs(m_FinalChannelCount).SetAlphaPremultiplication True
                
                m_FinalChannelCount = m_FinalChannelCount + 1
                
            Else
                warningStack.AddString "Couldn't create target DIB for this source channel; memory may be low."
                GoTo SkipChannel
            End If
            
            'Immediately free the source channel data to relieve memory pressure
            Erase .ciDataDecoded
            
        End With

SkipChannel:
    Next chIndex
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Image assembly took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ConstructImage_Multichannel", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ConstructImage_Multichannel, #" & Err.Number & ": " & Err.Description
    ConstructImage_Multichannel = psd_Failure
    
End Function

Private Function ConstructImage_RGB(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD) As PD_PSDResult

    ConstructImage_RGB = psd_Success
    On Error GoTo InternalVBError
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Interpretation of the underlying data depends on bit-depth, obviously
    Dim bytesPerChannel As Long
    bytesPerChannel = (bitsPC \ 8)
    
    'Set a flag to determine ICC profile availability; this affects a *lot* of subsequent work in this function,
    ' so we need to reference it early (and often).
    Dim iccAvailable As Boolean
    iccAvailable = (Not srcProfile Is Nothing)
    If iccAvailable Then iccAvailable = srcProfile.HasICCData()
    
    'If there's no embedded alpha channel, we can skip certain processing steps; determine alpha status now
    Dim alphaPresent As Boolean
    alphaPresent = False
    
    'TODO: figure out if we want to account for layer masks here; we technically don't need to, as the mask
    ' gets assembled separately and "composited" over the top - but we *would* need to premultiply in that case,
    ' so maybe a separate tracker for layer masks specifically is a better idea?
    
    'Note that alpha channels can be explicitly marked (with an ID of -1) or they can simply be "extra" channels
    ' without an obvious RGB mapping.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(i).ciID = cidAlpha) Or (m_Channels(i).ciID > cidBlue) Then alphaPresent = True
    Next i
    
    'Layer dimensions determine the size of the constructed layer, obviously; layer sizes are completely independent
    ' of the parent image.
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    
    'Determine the size of a single channel; we'll need this to calculate offsets into our planar assembly array
    Dim channelSize As Long
    channelSize = bytesPerChannel * layerWidth * layerHeight
    
    'To simplify handling if channels are missing, start with a white, opaque destination.
    Set m_LayerDIB = New pdDIB
    If (Not m_LayerDIB.CreateBlank(layerWidth, layerHeight, 32, vbWhite, 255)) Then
        InternalError "ConstructImage_RGB", "Couldn't allocate image buffer; image is likely too large."
        ConstructImage_RGB = psd_Failure
        Exit Function
    End If
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Construction setup took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    Dim x As Long, y As Long
    Dim tmpBytes() As Byte, tmpSA As SafeArray1D
    Dim scanStart As Long, scanWidth As Long
    scanStart = m_LayerDIB.GetDIBPointer()
    scanWidth = m_LayerDIB.GetDIBStride()
    
    'We now branch down two paths, depending on the presence of an ICC profile.  If an ICC profile exists,
    ' we can merge color management, interleaving, and swizzling into a single step, reducing import time
    ' by a meaningful amount.  If an ICC profile *doesn't* exist (meaning the file came from a 3rd-party
    ' PSD writer), we'll manually interleave and swizzle ourselves.
    If iccAvailable Then
        
        'To keep memory requirements low, we're going to process the image one scanline at a time.  This prevents
        ' us from needing to allocate another (potentially massive) image-sized buffer.
        Dim channelScanlineSize As Long
        channelScanlineSize = bytesPerChannel * layerWidth
        
        'Figure out how many source channels we have and where their indices lie.  (Channel order isn't
        ' necessarily guaranteed in PSDs.)
        Dim rIndex As Long, gIndex As Long, bIndex As Long, aIndex As Long
        rIndex = -1: gIndex = -1: bIndex = -1: aIndex = -1
        
        For x = 0 To Abs(m_ChannelCount) - 1
            If (m_Channels(x).ciID = cidBlue) Then
                bIndex = x
            ElseIf (m_Channels(x).ciID = cidGreen) Then
                gIndex = x
            ElseIf (m_Channels(x).ciID = cidRed) Then
                rIndex = x
            ElseIf (m_Channels(x).ciID = cidAlpha) Or (m_Channels(x).ciID > cidBlue) Then
                aIndex = x
            End If
        Next x
        
        'Prep a temporary buffer.  We'll assemble each scanline into a contiguous planar chunk in this buffer,
        ' then color-manage, interleave, swizzle, and (when necessary) swap endianness while merging the data
        ' into the final image buffer.
        Dim tmpPlanarBuffer() As Byte, srcLineSize As Long
        If alphaPresent Then srcLineSize = channelScanlineSize * 4 Else srcLineSize = channelScanlineSize * 3
        ReDim tmpPlanarBuffer(0 To srcLineSize - 1) As Byte
        
        'Prep source and destination ICC profiles
        Dim lcmsSrcProfile As pdLCMSProfile, lcmsDstProfile As pdLCMSProfile, lcmsTransform As pdLCMSTransform
        Set lcmsSrcProfile = New pdLCMSProfile
        lcmsSrcProfile.CreateFromPDICCObject srcProfile
        
        Set lcmsDstProfile = New pdLCMSProfile
        lcmsDstProfile.CreateSRGBProfile
        
        'Preparing the transform itself is a little more complicated, as it varies based on source color depth
        ' and the presence of alpha channels.  (Note that we also explicitly request blackpoint compensation to
        ' produce results similar to Adobe's CMM.)
        Dim srcColorMode As LCMS_PIXEL_FORMAT
        If (bytesPerChannel = 1) Then
            srcColorMode = TYPE_RGB_8_PLANAR
        ElseIf (bytesPerChannel = 2) Then
            srcColorMode = TYPE_RGB_16_PLANAR Or FLAG_SE
        ElseIf (bytesPerChannel = 4) Then
            srcColorMode = TYPE_RGB_FLT Or FLAG_PLANAR
        End If
        If alphaPresent Then srcColorMode = srcColorMode Or FLAG_ALPHAPRESENT
        
        Set lcmsTransform = New pdLCMSTransform
        lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile, srcColorMode, TYPE_BGRA_8, lcmsSrcProfile.GetRenderingIntent(), cmsFLAGS_BLACKPOINTCOMPENSATION Or cmsFLAGS_COPY_ALPHA
        
        'Iterate through each scanline in the image
        For y = 0 To layerHeight - 1
        
            'Copy all present channels into the contiguous planar buffer
            If (rIndex >= 0) Then CopyMemoryStrict VarPtr(tmpPlanarBuffer(0)), VarPtr(m_Channels(rIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
            If (gIndex >= 0) Then CopyMemoryStrict VarPtr(tmpPlanarBuffer(channelScanlineSize)), VarPtr(m_Channels(gIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
            If (bIndex >= 0) Then CopyMemoryStrict VarPtr(tmpPlanarBuffer(channelScanlineSize * 2)), VarPtr(m_Channels(bIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
            If (aIndex >= 0) Then CopyMemoryStrict VarPtr(tmpPlanarBuffer(channelScanlineSize * 3)), VarPtr(m_Channels(aIndex).ciDataDecoded(y * channelScanlineSize)), channelScanlineSize
            
            'Use LCMS to transform the data into our final RGBA buffer
            lcmsTransform.ApplyTransformToArbitraryMemoryEx VarPtr(tmpPlanarBuffer(0)), scanStart + y * scanWidth, layerWidth, 1, srcLineSize, scanWidth, channelScanlineSize, 0&
            
        Next y
        
    'No embedded color profile.  Swizzle and interleave manually on a channel-by-channel basis.
    Else
        
        'Iterate through all channels, copying their planar data into the interleaved final image buffer as we go
        Dim chIndex As Long
        For chIndex = 0 To Abs(m_ChannelCount) - 1
            
            With m_Channels(chIndex)
                
                'There are a lot of reasons to ignore a given channel:
                ' 1) Channel data wasn't decoded successfully
                ' 2) Channel doesn't contribute pixel data to the final image (e.g. adjustment layers)
                ' 3) Layer masks, which have to be handled separately due to custom size and layout considerations
                If (Not .ciDecodedOK) Then GoTo SkipChannel
                If (.ciDataLength = 0) Then GoTo SkipChannel
                If (.ciID < cidAlpha) Then
                    warningStack.AddString "Found a layer mask, but PD can't interpret these yet - sorry."
                    GoTo SkipChannel
                End If
                
                'Channel order in the file is not guaranteed, so we'll need to manually swizzle based on channel ID.
                Dim cOffset As Long, ySrcOffset As Long, xLoopEnd As Long
                If (.ciID = cidBlue) Then
                    cOffset = 0
                ElseIf (.ciID = cidGreen) Then
                    cOffset = 1
                ElseIf (.ciID = cidRed) Then
                    cOffset = 2
                ElseIf (.ciID = cidAlpha) Or (.ciID > cidBlue) Then
                    cOffset = 3
                End If
                    
                'Actually painting the color channel data is extremely simple; just copy the byte into its
                ' relevant position in the target DIB.
                m_LayerDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
                xLoopEnd = layerWidth - 1
                
                '8-bit channel (most common case)
                If (bytesPerChannel = 1) Then
                
                    For y = 0 To layerHeight - 1
                        tmpSA.pvData = scanStart + (scanWidth * y)
                        ySrcOffset = y * layerWidth
                    For x = 0 To xLoopEnd
                        tmpBytes(x * 4 + cOffset) = .ciDataDecoded(ySrcOffset + x)
                    Next x
                    Next y
                    
                '16-bit channel (take high-bit only, and note that the source data *is big-endian*
                ElseIf (bytesPerChannel = 2) Then
                
                    For y = 0 To layerHeight - 1
                        tmpSA.pvData = scanStart + (scanWidth * y)
                        ySrcOffset = y * layerWidth * 2
                    For x = 0 To xLoopEnd
                        tmpBytes(x * 4 + cOffset) = .ciDataDecoded(ySrcOffset + x * 2)
                    Next x
                    Next y
                
                '32-bit channels are not supported if an ICC profile is missing, as a manual transform will
                ' produce a garbage-quality image.  (PSDs from Photoshop will *always* have an ICC profile
                ' attached, so this is only relevant to garbage-quality 3rd-party implementations.)
                ElseIf (bytesPerChannel = 4) Then
                    warningStack.AddString "32-bit channels require an ICC profile for correct assembly."
                End If
                
                'Manually free our unsafe array wrapper before proceeding, and immediately free our source data
                ' to relieve memory pressure
                m_LayerDIB.UnwrapArrayFromDIB tmpBytes
                Erase .ciDataDecoded
                
            End With
    
SkipChannel:
        Next chIndex
    
    End If
    
    If PSD_DEBUG_VERBOSE Then
        If iccAvailable Then
            PDDebug.LogAction "Image assembly (ICC path) took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        Else
            PDDebug.LogAction "Image assembly (no ICC) took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        End If
        VBHacks.GetHighResTime startTime
    End If
    
    'If meaningful alpha data exists in the file, premultiply it now
    If alphaPresent Then
        m_LayerDIB.SetAlphaPremultiplication True
        If PSD_DEBUG_VERBOSE Then
            PDDebug.LogAction "Alpha premultiplication took: " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
    Else
        m_LayerDIB.SetInitialAlphaPremultiplicationState True
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ConstructImage_RGB", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ConstructImage_RGB, #" & Err.Number & ": " & Err.Description
    ConstructImage_RGB = psd_Failure
    
End Function

'Grayscale, duotone, and indexed images share a unified load function.  (Monochrome images are also handled here,
' with their encoding having been silently "upsampled" to grayscale in a previous stage.)
Private Function ConstructImage_SingleChannel(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD) As PD_PSDResult

    ConstructImage_SingleChannel = psd_Success
    On Error GoTo InternalVBError
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Interpretation of the underlying data depends on bit-depth, obviously
    Dim bytesPerChannel As Long
    bytesPerChannel = (bitsPC \ 8)
    
    'Set a flag to determine ICC profile availability; this affects a *lot* of subsequent work in this function,
    ' so we need to reference it early (and often).
    Dim lcmsSrcProfile As pdLCMSProfile, lcmsDstProfile As pdLCMSProfile, lcmsTransform As pdLCMSTransform
    Dim iccAvailable As Boolean
    iccAvailable = (Not srcProfile Is Nothing)
    If iccAvailable Then iccAvailable = srcProfile.HasICCData()
    If iccAvailable Then
        Set lcmsSrcProfile = New pdLCMSProfile: Set lcmsDstProfile = New pdLCMSProfile: Set lcmsTransform = New pdLCMSTransform
    End If
    
    'If there's no embedded alpha channel, we can skip certain processing steps; determine alpha status now
    Dim alphaPresent As Boolean
    alphaPresent = False
    
    'TODO: figure out if we want to account for layer masks here; we technically don't need to, as the mask
    ' gets assembled separately and "composited" over the top - but we *would* need to premultiply in that case,
    ' so maybe a separate tracker for layer masks specifically is a better idea?
    
    'Note that alpha channels can be explicitly marked (with an ID of -1) or they can simply be "extra" channels
    ' without an obvious RGB mapping.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(i).ciID = cidAlpha) Or (m_Channels(i).ciID > cidIndexed) Then alphaPresent = True
    Next i
    
    'In earlier builds, I attempted to construct an lcms transform very early in the assembly process,
    ' then use that transform to merge all assembly steps into a single color-managed transform handled
    ' by lcms.  This proved doubly problematic: it takes a *lot* of memory to assemble a complete planar
    ' representation of large 16-bit PSD files, and lcms is unpredictable about converting planar data
    ' to interleaved while simultaneously changing bit-depth.  (This isn't surprising, given what a weird
    ' edgecase it is.)  As such, PD now takes the slower but much safer route of manually downsampling
    ' the PSD data channel-by-channel into the final image buffer, then color-managing the final 8-bit
    ' buffer as a final step.  This solves all reliability issues and greatly reduces memory presssure,
    ' so I'm content to keep it as PD's "permanent" solution.
    
    'Layer dimensions determine the size of the constructed layer, obviously; layer sizes are completely independent
    ' of the parent image.
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    
    'Determine the size of a single channel; we'll need this to calculate offsets into our planar assembly array
    Dim channelSize As Long
    channelSize = bytesPerChannel * layerWidth * layerHeight
    
    'To simplify handling if channels are missing, start with a white, opaque destination.
    Set m_LayerDIB = New pdDIB
    If (Not m_LayerDIB.CreateBlank(layerWidth, layerHeight, 32, vbWhite, 255)) Then
        InternalError "ConstructImage_SingleChannel", "Couldn't allocate image buffer; image is likely too large."
        ConstructImage_SingleChannel = psd_Failure
        Exit Function
    End If
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Construction setup took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    Dim x As Long, y As Long
    Dim tmpBytes() As Byte, tmpLongs() As Long, tmpSA As SafeArray1D
    Dim scanStart As Long, scanWidth As Long, xLoopEnd As Long, ySrcOffset As Long
    scanStart = m_LayerDIB.GetDIBPointer()
    scanWidth = m_LayerDIB.GetDIBStride()
    
    'Layers will typically match the color-mode and bit-depth of the parent image, but it can be
    ' convenient to adjust these "on-the-fly" (e.g. bitmap/monochrome layers are inconvenient to
    ' render manually, so we translate them to a grayscale equivalent on the fly).  These parameters
    ' are reset as each channel is processed.
    Dim chnColorMode As PSD_ColorMode, chnBitDepth As Long
    
    'Unfortunately, alpha data may appear in the channel list before gray data.  Our accelerated approach to
    ' palette handling means that palette data may overwrite existing alpha data, if any.  As such, we make
    ' sure to process grayscale/palette/duotone data BEFORE processing alpha channels.
    
    'Iterate through all channels, copying their planar data into the interleaved final image buffer as we go
    Dim chIndex As Long
    For chIndex = 0 To Abs(m_ChannelCount) - 1
        
        'Reset any on-the-fly changes we've made to color mode or bit-depth
        chnColorMode = m_ColorMode
        chnBitDepth = bitsPC
        
        With m_Channels(chIndex)
            
            'There are a lot of reasons to ignore a given channel:
            ' 1) Channel data wasn't decoded successfully
            ' 2) Channel doesn't contribute pixel data to the final image (e.g. adjustment layers)
            ' 3) Layer masks, which have to be handled separately due to custom size and layout considerations
            If (Not .ciDecodedOK) Then GoTo SkipChannel
            If (.ciDataLength = 0) Then GoTo SkipChannel
            If (.ciID < cidAlpha) Then
                warningStack.AddString "Found a layer mask, but PD can't interpret these yet - sorry."
                GoTo SkipChannel
            End If
            
            'Monochrome layers are a special case.  If encountered, it's easiest to expand them to an 8-bit stream
            ' and treat the results as grayscale.  (This allows us to reuse our grayscale converter "as-is".)
            If (chnBitDepth = 1) Then
                ExpandMonochrome m_Channels(chIndex), layerHeight, layerWidth
                chnColorMode = cm_Grayscale
                chnBitDepth = 8
                bytesPerChannel = 1
            End If
            
            'Indexed color images can still have masks, so we need to check channel IDs.
            ' (Note that the "indexed" channel ID is simply 0; this refers to the pixel data
            ' channel of indexed, grayscale, and duotone color modes.)
            If (.ciID = cidIndexed) Then
                
                'Next step is to construct a palette for the image; this step varies
                ' based on channel type.
                Dim imgPalette() As RGBQuad, palSize As Long, trnsIndex As Long
                trnsIndex = -1
                
                'In indexed images, we can simply read the color table information from the parent image
                If (chnColorMode = cm_Indexed) Then
                    
                    parentPSD.GetColorTableData imgPalette, palSize, trnsIndex
                    If (palSize > 0) Then
                        If (trnsIndex >= 0) And (trnsIndex <= 255) Then imgPalette(trnsIndex).Alpha = 0
                    Else
                        ReDim imgPalette(0 To 255) As RGBQuad
                    End If
                
                'For grayscale, bitmap, or duotone images let's manually construct a linear grayscale palette
                Else
                
                    ReDim imgPalette(0 To 255) As RGBQuad
                    palSize = 256
                    For x = 0 To 255
                        imgPalette(x).Blue = x
                        imgPalette(x).Green = x
                        imgPalette(x).Red = x
                        If (x = trnsIndex) Then imgPalette(x).Alpha = 0 Else imgPalette(x).Alpha = 255
                    Next x
                    
                    'For grayscale channels, we want to do handle any ICC processing *before* converting
                    ' the data internally to RGBA.  (Grayscale profiles can only be applied to single-channel
                    ' grayscale data, so if we convert the data to RGBA format, the profile will be unusable.)
                    If iccAvailable Then
                    
                        lcmsSrcProfile.CreateFromPDICCObject srcProfile
                        lcmsDstProfile.CreateGenericGrayscaleProfile
                        
                        'Construct and apply the transform in-place over the existing grayscale data
                        If (bytesPerChannel = 1) Then
                            lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile, TYPE_GRAY_8, TYPE_GRAY_8, lcmsSrcProfile.GetRenderingIntent(), cmsFLAGS_BLACKPOINTCOMPENSATION
                            lcmsTransform.ApplyTransformToArbitraryMemory VarPtr(.ciDataDecoded(0)), VarPtr(.ciDataDecoded(0)), layerWidth, layerWidth, layerHeight, layerWidth
                        ElseIf (bytesPerChannel = 2) Then
                            lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile, TYPE_GRAY_16_SE, TYPE_GRAY_16_SE, lcmsSrcProfile.GetRenderingIntent(), cmsFLAGS_BLACKPOINTCOMPENSATION
                            lcmsTransform.ApplyTransformToArbitraryMemory VarPtr(.ciDataDecoded(0)), VarPtr(.ciDataDecoded(0)), layerWidth * 2, layerWidth * 2, layerHeight, layerWidth
                        ElseIf (bytesPerChannel = 4) Then
                            lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile, TYPE_GRAY_FLT, TYPE_GRAY_FLT, lcmsSrcProfile.GetRenderingIntent(), cmsFLAGS_BLACKPOINTCOMPENSATION
                            lcmsTransform.ApplyTransformToArbitraryMemory VarPtr(.ciDataDecoded(0)), VarPtr(.ciDataDecoded(0)), layerWidth * 4, layerWidth * 4, layerHeight, layerWidth
                        End If
                        
                    End If
                    
                '/End palette assembly
                End If
                
                'To accelerate the load process, we want to set all four bytes of individual pixels as Longs
                ' (rather than the slower step of assigning RGBA values individually).  To "trick" VB into
                ' doing this for us, we'll alias a Long array around the destination DIB, and we need a matching
                ' Long-type array for the palette data.
                Dim palLookup() As Long
                ReDim palLookup(0 To 255) As Long
                For x = 0 To palSize - 1
                    GetMem4 VarPtr(imgPalette(x)), palLookup(x)
                Next x
                
                m_LayerDIB.WrapLongArrayAroundScanline tmpLongs, tmpSA, 0
                xLoopEnd = layerWidth - 1
                
                '8-bit channel
                If (bytesPerChannel = 1) Then
                    
                    For y = 0 To layerHeight - 1
                        tmpSA.pvData = scanStart + (scanWidth * y)
                        ySrcOffset = y * layerWidth
                    For x = 0 To xLoopEnd
                        tmpLongs(x) = palLookup(.ciDataDecoded(ySrcOffset + x))
                    Next x
                    Next y
                    
                '16-bit channel
                ElseIf (bytesPerChannel = 2) Then
                    For y = 0 To layerHeight - 1
                        tmpSA.pvData = scanStart + (scanWidth * y)
                        ySrcOffset = y * layerWidth * 2
                    For x = 0 To xLoopEnd
                        tmpLongs(x) = palLookup(.ciDataDecoded(ySrcOffset + x * 2))
                    Next x
                    Next y
                    
                '32-bit (grayscale only, hopefully; other color modes may not produce desirable results)
                Else
                    
                    Dim tmpSingle As Single
                    For y = 0 To layerHeight - 1
                        tmpSA.pvData = scanStart + (scanWidth * y)
                        ySrcOffset = y * layerWidth * 4
                    For x = 0 To xLoopEnd
                    
                        'Extract the source data into a single-type value
                        GetMem4_Ptr VarPtr(.ciDataDecoded(ySrcOffset + x * 4)), VarPtr(tmpSingle)
                        
                        'Clamp for safety before assigning
                        If (tmpSingle < 0!) Then tmpSingle = 0!
                        If (tmpSingle > 1!) Then tmpSingle = 1!
                        tmpLongs(x) = palLookup(Int(tmpSingle * 255!))
                        
                    Next x
                    Next y
                    
                End If
                
                m_LayerDIB.UnwrapLongArrayFromDIB tmpLongs
                
                'Immediately free the source channel data to relieve memory pressure
                Erase .ciDataDecoded
            
            '/End .ciID = valid channel ID
            End If
            
        End With

SkipChannel:
    Next chIndex
    
    'If alpha is present, repeat the above steps, looking for alpha channels specifically.
    If alphaPresent Then
        
        For chIndex = 0 To Abs(m_ChannelCount) - 1
            
            With m_Channels(chIndex)
                
                If (Not .ciDecodedOK) Then GoTo SkipChannelAlpha
                If (.ciDataLength = 0) Then GoTo SkipChannelAlpha
                
                'Alpha channel found
                If (.ciID > cidIndexed) Or (.ciID = cidAlpha) Then
            
                    m_LayerDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
                    xLoopEnd = layerWidth - 1
                    
                    Const ONE_DIV_255 As Single = 1! / 255!
                    
                    '8-bit alpha channel
                    If (bytesPerChannel = 1) Then
                    
                        For y = 0 To layerHeight - 1
                            tmpSA.pvData = scanStart + (scanWidth * y)
                            ySrcOffset = y * layerWidth
                        For x = 0 To xLoopEnd
                            tmpBytes(x * 4 + 3) = CSng(tmpBytes(x * 4 + 3)) * (CSng(.ciDataDecoded(ySrcOffset + x)) * ONE_DIV_255)
                        Next x
                        Next y
                    
                    '16-bit alpha channel
                    ElseIf (bytesPerChannel = 2) Then
                    
                        For y = 0 To layerHeight - 1
                            tmpSA.pvData = scanStart + (scanWidth * y)
                            ySrcOffset = y * layerWidth * 2
                        For x = 0 To xLoopEnd
                            tmpBytes(x * 4 + 3) = CSng(tmpBytes(x * 4 + 3)) * (CSng(.ciDataDecoded(ySrcOffset + x * 2)) * ONE_DIV_255)
                        Next x
                        Next y
                    
                    '32-bit alpha channel
                    Else
                        
                        Dim tmpAlphaFloat As Single
                        For y = 0 To layerHeight - 1
                            tmpSA.pvData = scanStart + (scanWidth * y)
                            ySrcOffset = y * layerWidth * 4
                        For x = 0 To xLoopEnd
                        
                            'Extract the source data into a single-type value
                            GetMem4_Ptr VarPtr(.ciDataDecoded(ySrcOffset + x * 4)), VarPtr(tmpAlphaFloat)
                            
                            'Clamp for safety before assigning
                            If (tmpAlphaFloat < 0!) Then
                                tmpAlphaFloat = 0!
                            ElseIf (tmpAlphaFloat > 1!) Then
                                tmpAlphaFloat = 1!
                            End If
                            
                            tmpBytes(x * 4 + 3) = CSng(tmpBytes(x * 4 + 3)) * tmpAlphaFloat
                            
                        Next x
                        Next y
                        
                    End If
                
                    m_LayerDIB.UnwrapArrayFromDIB tmpBytes
                
                '/End alpha channel found
                End If
                
                'Immediately free the source channel data to relieve memory pressure
                Erase .ciDataDecoded
            
            End With

SkipChannelAlpha:
        Next chIndex
        
    End If
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Image assembly took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'With the destination image assembled successfully, we can now perform any post-processing
    ' (including color management).  Note that grayscale images have already handle color management as
    ' part of the decode process, so this step is only relevant for *color* images.
    If iccAvailable And (imgColorMode = cm_Indexed) Then
        
        lcmsSrcProfile.CreateFromPDICCObject srcProfile
        lcmsDstProfile.CreateSRGBProfile
        
        'Explicitly request blackpoint compensation to produce results similar to Adobe's CMM
        lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile, TYPE_BGRA_8, TYPE_BGRA_8, lcmsSrcProfile.GetRenderingIntent(), cmsFLAGS_BLACKPOINTCOMPENSATION Or cmsFLAGS_COPY_ALPHA
        lcmsTransform.ApplyTransformToPDDib m_LayerDIB
        
        'Transforms can consume a surprising amount of memory, depending on the way the ICC is built.  (For example,
        ' files can define giant 3D mapping tables - and this approach is actually very common in CMYK images.)
        ' To relieve memory pressure on subsequent post-processing functions, free all LCMS objects immediately.
        Set lcmsSrcProfile = Nothing: Set lcmsDstProfile = Nothing: Set lcmsTransform = Nothing
        
        If PSD_DEBUG_VERBOSE Then
            PDDebug.LogAction "Color management took: " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
    End If
    
    'If meaningful alpha data exists in the file, premultiply it now
    If alphaPresent Then
        m_LayerDIB.SetAlphaPremultiplication True
        If PSD_DEBUG_VERBOSE Then
            PDDebug.LogAction "Alpha premultiplication took: " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
    Else
        m_LayerDIB.SetInitialAlphaPremultiplicationState True
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ConstructImage_SingleChannel", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ConstructImage_SingleChannel, #" & Err.Number & ": " & Err.Description
    ConstructImage_SingleChannel = psd_Failure
    
End Function

'After channel data has been retrieved from the PSD, we must decode it into something usable by PD.
' That's what this function handles.  If the underlying data is irrelevant to PD (e.g. adjustment or effect layers),
' this function simply frees the associated layer data to cut down on memory requirements.  Layers with relevant
' image data, however, will produce a usable pdDIB object from this step.
Friend Function DecodeChannels(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD) As PD_PSDResult
    
    DecodeChannels = psd_Success
    
    On Error GoTo InternalVBError
    
    'Decoded data size is contingent on the layer's rectangle
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    
    'Layers will typically match the color-mode and bit-depth of the parent image, but it can be
    ' convenient to adjust these "on-the-fly" (e.g. bitmap/monochrome layers are inconvenient to
    ' render manually, so we translate them to a grayscale equivalent on the fly).  These parameters
    ' are reset as each channel is loaded, so that changes to one of them does not affect any
    ' other layers in the image.
    Dim chnColorMode As Long, chnBitDepth As Long
    
    'Channel data is stored in planar format (e.g. not interleaved - RRRGGGBBB instead of RGBRGBRGB).
    ' As such, each channel must be loaded independent of other channels in the image.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        
        'Reset any on-the-fly changes we've made to color mode or bit-depth
        chnColorMode = imgColorMode
        chnBitDepth = bitsPC
        
        With m_Channels(i)
            
            'Failsafe reset of decoding results; this tells subsequent steps to ignore this layer's data
            .ciDecodedOK = False
            
            'Some channels contribute nothing to the final image (and are zero-length).  Ignore them.
            If (.ciDataLength = 0) Then
                .ciDecodedOK = False
                GoTo SkipChannel
            End If
            
            'Layer masks represent a significant complication; we will deal with them in the future
            If (.ciID < cidAlpha) Then
                
                'TODO: calculate special decoding rect, as it is *not* the same as the regular rect.  We also have to
                ' deal with crap like "default color" which determines whether the encoded data represents hiding
                ' or showing the layer beneath.
                'Debug.Print .ciID
                'If .ciID = cidUserLayerMask Then
                '    Debug.Print m_MaskRect.Left, m_MaskRect.Top, m_MaskRect.Right - m_MaskRect.Left, m_MaskRect.Bottom - m_MaskRect.Top
                '    Debug.Print m_RealMaskRect.Left, m_RealMaskRect.Top, m_RealMaskRect.Right - m_RealMaskRect.Left, m_RealMaskRect.Bottom - m_RealMaskRect.Top
                'End If
                .ciDecodedOK = False
                GoTo SkipChannel
                
            End If
            
            'Our goal in this step is to translate the PSD channel stream into a standalone byte array
            ' at its native bit-per-channel size.  The amount of work required to do this varies based
            ' on the compression technique used.
            Dim bytesPerChannel As Long
            bytesPerChannel = (chnBitDepth \ 8)
            
            'For color channels, we can calculate the size of the finished channel data from the image's dimensions
            ' and current bit-depth.
            Dim numPixels As Long
            numPixels = layerWidth * layerHeight
            
            Dim finalArraySize As Long
            
            'Monochrome is a special case.
            If (bytesPerChannel = 0) Then
                layerWidth = (layerWidth + 7) \ 8
                If (layerWidth = 0) Then layerWidth = 1
                numPixels = layerWidth * layerHeight
                finalArraySize = numPixels
            Else
                finalArraySize = numPixels * bytesPerChannel
            End If
            
            'Monochrome layers will need to be handled separately
            If (finalArraySize > 0) Then
                
                'Prep a temporary output buffer
                Dim x As Long, y As Long
                If PSD_DEBUG_VERBOSE Then
                    PDDebug.LogAction "Channel #" & i & " (ID=" & .ciID & ") has compression type: " & .ciCompression
                    Dim decodeStartTime As Currency
                    VBHacks.GetHighResTime decodeStartTime
                End If
                
                'Compression mode "zip" is apparently always "zip + prediction" when the source data is 32-bit.
                ' Replace this on-the-fly to simplify subsequent processing.
                If (bytesPerChannel = 4) And (.ciCompression = ccZip) Then .ciCompression = ccZipWithPrediction
                
                Select Case .ciCompression
                
                    'Raw compression = no compression.  Use the bytes as-is.
                    Case ccRaw
                    
                        'Technically, we could do something clever like swap safearray headers between the
                        ' coded and decoded type members - but for now, simply perform a copy to keep
                        ' subsequent code simple.
                        ReDim .ciDataDecoded(0 To .ciDataLength - 1) As Byte
                        CopyMemoryStrict VarPtr(.ciDataDecoded(0)), VarPtr(.ciData(0)), .ciDataLength
                        .ciDecodedOK = True
                        
                    'PackBits = RLE compression.  Data must be decoded before transferring it to the image.
                    Case ccPackBits
                        
                        'We know the necessary size of the final array; prep the receiving buffer now.
                        ReDim .ciDataDecoded(0 To finalArraySize - 1) As Byte
                        
                        'For convenience, we're going to wrap a temporary pdStream object around the source bytes.
                        ' This simplifies translating a large series of big-endian values.
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
                        
                        'We got what we need from the stream.  Note its position (as this is where pixel data begins),
                        ' then free it.
                        Dim basePtrPosition As Long
                        basePtrPosition = rleStream.GetPosition()
                        Set rleStream = Nothing
                        
                        'We now want to extract each scanline in turn.  (Technically, we don't need to
                        ' know RLE line lengths to do this, but they provide a nice failsafe mechanism
                        ' for ensuring we don't overrun line boundaries.)
                        Dim bytesReadDst As Long, bytesReadSrc As Long, bytesReadLine As Long, rleCode As Long, tmpVal As Long
                        bytesReadDst = 0
                        bytesReadSrc = 0
                        
                        For y = 0 To layerHeight - 1
                            
                            bytesReadLine = 0
                            
                            Do
                            
                                'Packbits is extremely simple as compression strategies go.  Treat the first
                                ' (signed) byte in each stream as a code.
                                rleCode = .ciData(basePtrPosition + bytesReadSrc)
                                If (rleCode > 127) Then rleCode = rleCode - 256
                                bytesReadLine = bytesReadLine + 1
                                bytesReadSrc = bytesReadSrc + 1
                                
                                'If the code is >= 0, the next (code) bytes should be copied to the destination
                                ' buffer as-is.
                                If (rleCode > 0) Then
                                    
                                    rleCode = rleCode + 1
                                    bytesReadLine = bytesReadLine + rleCode
                                    'If (bytesReadLine > lineLengths(y)) Then
                                    '    Debug.Print "WARNING RLE OVERFLOW (1)! " & lineLengths(y), y, rleCode
                                    '    Exit Do
                                    'End If
                                    CopyMemoryStrict VarPtr(.ciDataDecoded(bytesReadDst)), VarPtr(.ciData(basePtrPosition + bytesReadSrc)), rleCode
                                    bytesReadDst = bytesReadDst + rleCode
                                    bytesReadSrc = bytesReadSrc + rleCode
                                    
                                'If the code is on the range [-127,-1], inclusive, the next byte should be
                                ' repeated [1 - code] times in the destination buffer.
                                ElseIf (rleCode > -128) Then
                                    
                                    rleCode = 1 - rleCode
                                    bytesReadLine = bytesReadLine + 1
                                    'If (bytesReadLine > lineLengths(y)) Then
                                    '    Debug.Print "WARNING RLE OVERFLOW (2)! " & lineLengths(y), y, rleCode
                                    '    Exit Do
                                    'End If
                                    FillMemory VarPtr(.ciDataDecoded(bytesReadDst)), rleCode, .ciData(basePtrPosition + bytesReadSrc) 'tmpVal
                                    bytesReadDst = bytesReadDst + rleCode
                                    bytesReadSrc = bytesReadSrc + 1
                                    
                                '-128 means do nothing (nop).  I'm not sure why an encoder would set this, but if it does,
                                ' the decoder will continue normally.
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
                        .ciDecodedOK = True
                        DecodeChannels = psd_Success
                    
                    'Zip compression has not actually been tested "in the wild", because I don't have a usable
                    ' test file with this encoding.  That said, I see no reason why the current code wouldn't
                    ' work as-is (famous last words)...
                    Case ccZip
                        
                        'Inflate the stream directly into its final buffer.
                        If Plugin_zLib.IsZLibAvailable() Then
                            ReDim .ciDataDecoded(0 To finalArraySize - 1) As Byte
                            If Plugin_zLib.ZlibDecompress_UnsafePtr(VarPtr(.ciDataDecoded(0)), finalArraySize, VarPtr(.ciData(0)), .ciDataLength) Then
                                .ciDecodedOK = True
                                DecodeChannels = psd_Success
                            Else
                                DecodeChannels = psd_Warning
                                InternalError "DecodeChannels", "zLib failed to inflate compressed channel stream."
                            End If
                        Else
                            DecodeChannels = psd_Failure
                            InternalError "DecodeChannels", "zLib missing"
                        End If
                        
                    Case ccZipWithPrediction
                        
                        Dim startTime As Currency
                        VBHacks.GetHighResTime startTime
                        
                        'Zip w/prediction is only available on 16-bit and 32-bit channels.
                        If (bitsPC = 16) Or (bitsPC = 32) Then
                        
                            'Start by inflating the stream.
                            If Plugin_zLib.IsZLibAvailable() Then
                                
                                ReDim .ciDataDecoded(0 To finalArraySize - 1) As Byte
                                If Plugin_zLib.ZlibDecompress_UnsafePtr(VarPtr(.ciDataDecoded(0)), finalArraySize, VarPtr(.ciData(0)), .ciDataLength) Then
                                    
                                    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Time spent in inflate: " & VBHacks.GetTimeDiffNowAsString(startTime)
                                    VBHacks.GetHighResTime startTime
                                    
                                    'Erase the original, compressed data as it is no longer required
                                    Erase .ciData
                                    
                                    'Rather than flip endianness "as we go", it's easier to just swap the entire
                                    ' stream in advance. (Note that we don't do this for 32-bit data, as delta encoding
                                    ' is done on a byte-by-byte process, with 1st-4th bytes manually assembled into
                                    ' their own sub-planes.)
                                    If (bitsPC = 16) Then
                                        VBHacks.SwapEndianness16 .ciDataDecoded
                                        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Time spent in endian swap: " & VBHacks.GetTimeDiffNowAsString(startTime)
                                        VBHacks.GetHighResTime startTime
                                    End If
                                    
                                    Dim curDelta As Long, deltaPrev As Long, newValue As Long, tmpIndex As Long
                                    Dim srcIndex As Long, srcInt As Integer, srcLong As Long, yOffset As Long, xOffset As Long
                                    
                                    '16-bit PSDs use an extremely primitive delta encoding scheme - each line starts with a
                                    ' 2-byte unsigned int value, and subsequent bytes are the delta between the previous
                                    ' pixel and the current one.  This makes decoding tedious, but straightforward.
                                    ' (Note that I deliberately mix data types here, as VB lacks unsigned ints but we
                                    ' can "cheat" and use longs by shifting memory strategically.)
                                    If (bitsPC = 16) Then
                                        
                                        srcIndex = 0
                                        
                                        For y = 0 To (layerHeight - 1)
    
                                            'Retrieve the first value in each row and treat it as an absolute value
                                            GetMem2_Ptr VarPtr(.ciDataDecoded(srcIndex)), VarPtr(deltaPrev)
                                            srcIndex = srcIndex + 2
                                            
                                            'While here, cache the current scanline pointer to avoid unnecessary multipliations
                                            ' inside the inner loop.
                                            tmpIndex = y * layerWidth * 2
                                            
                                            'All subsequent bytes are treated as deltas
                                            For x = 1 To layerWidth - 1
                                                
                                                'Retrieve the next two bytes into a signed int
                                                GetMem2_Ptr VarPtr(.ciDataDecoded(srcIndex)), VarPtr(srcInt)
                                                
                                                'Apply the delta from the previous pixel
                                                newValue = (CLng(srcInt) + deltaPrev)
                                                
                                                'Store the result in-place over the old value.  (A lack of easy dereferencing
                                                ' makes this tedious, but GetMem2 provides a slight speed boost over RtlMoveMemory.)
                                                GetMem2_Ptr VarPtr(newValue), VarPtr(.ciDataDecoded(srcIndex))
                                                
                                                'Store the delta for the next pixel to use and increment our index into the array
                                                deltaPrev = newValue And &HFFFF&
                                                srcIndex = srcIndex + 2
                                                
                                            Next x
    
                                        Next y
                                    
                                    '32-bit data is a totally different beast.
                                    ElseIf (bitsPC = 32) Then
                                        
                                        srcIndex = 0
                                        
                                        'We also need a temporary buffer the size of one scanline
                                        Dim tmpScanLine() As Byte
                                        ReDim tmpScanLine(0 To layerWidth * 4 - 1) As Byte
                                        
                                        '32-bit prediction works on a row-by-row basis.  Before dealing with things like
                                        ' byte layout, we can just un-predict each row "as-is".
                                        For y = 0 To (layerHeight - 1)
                                            
                                            yOffset = y * layerWidth * 4
                                            
                                            'First byte is a "pure" value; leave it untouched
                                            deltaPrev = .ciDataDecoded(srcIndex)
                                            srcIndex = srcIndex + 1
                                            
                                            'All subsequent bytes are treated as deltas
                                            For x = 1 To (layerWidth * 4) - 1
                                                srcLong = .ciDataDecoded(srcIndex)
                                                srcLong = (srcLong + deltaPrev) And 255
                                                .ciDataDecoded(srcIndex) = srcLong
                                                deltaPrev = srcLong
                                                srcIndex = srcIndex + 1
                                            Next x
                                            
                                            'With the scanline decoded, we lastly need to manually re-assemble the four
                                            ' planar byte chunks into actual floats.  (The encoder separated bytes 1/2/3/4
                                            ' of each float into their own little planar segment to improve compression.)
                                            For x = 0 To layerWidth - 1
                                                xOffset = x * 4
                                                tmpScanLine(xOffset) = .ciDataDecoded(yOffset + layerWidth * 3 + x)
                                                tmpScanLine(xOffset + 1) = .ciDataDecoded(yOffset + layerWidth * 2 + x)
                                                tmpScanLine(xOffset + 2) = .ciDataDecoded(yOffset + layerWidth + x)
                                                tmpScanLine(xOffset + 3) = .ciDataDecoded(yOffset + x)
                                            Next x
                                            
                                            'Copy the completed scanline over the original data
                                            CopyMemoryStrict VarPtr(.ciDataDecoded(yOffset)), VarPtr(tmpScanLine(0)), layerWidth * 4
                                            
                                        Next y
                                        
                                    End If
                                    
                                    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Time spent in delta decode: " & VBHacks.GetTimeDiffNowAsString(startTime)
                                    VBHacks.GetHighResTime startTime
                                    
                                    'Because the image constructor expects big-endian data, we now have to re-swap all
                                    ' decoded data back to big-endian format (16-bit only; 32-bit was dealt with already).
                                    If (bitsPC = 16) Then
                                        VBHacks.SwapEndianness16 .ciDataDecoded
                                        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Time spent in final endianness swap: " & VBHacks.GetTimeDiffNowAsString(startTime)
                                    End If
                                    
                                    .ciDecodedOK = True
                                    
                                Else
                                    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "zLib failed to inflate stream."
                                    .ciDecodedOK = False
                                    DecodeChannels = psd_Failure
                                End If
                            Else
                                If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "zLib is missing; can't deflate stream."
                                .ciDecodedOK = False
                                DecodeChannels = psd_Failure
                            End If
                            
                        Else
                            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "ZIP w/prediction compression isn't supported at this bit-depth.  This PSD is malformed."
                            .ciDecodedOK = False
                            DecodeChannels = psd_Failure
                        End If
                        
                End Select
                
                If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Total time to decode channel #" & i & ": " & VBHacks.GetTimeDiffNowAsString(decodeStartTime)
                
            'Raw data is zero-length; this layer does not contain usable data
            Else
                DecodeChannels = psd_Warning
            End If
            
            'If we haven't already, free the original encoded data; we no longer require it
            Erase .ciData
            
        End With
    
SkipChannel:
    Next i
    
    Exit Function
    
InternalVBError:
    InternalError "DecodeChannels", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in DecodeChannels, #" & Err.Number & ": " & Err.Description
    DecodeChannels = psd_Failure
End Function

'1-bit-per-channel images are internally expanded to 8-bit grayscale to simplify rendering
Private Function ExpandMonochrome(ByRef srcChannel As PSD_ChannelInfo, ByVal numScanlines As Long, ByVal numBitsInScanline As Long) As Boolean
    
    'By the time this function is called, the monochrome data stream has been expanded to raw pixel data.
    ' We want to expand it a second time, to full grayscale data, to simplify rendering a composite image.
    
    'We'll use pdStream to generate a new, 8-bpp version of the image; this simplifies the process of
    ' collecting individual bytes into a contiguous stream.
    Dim tmpStream As pdStream
    Set tmpStream = New pdStream
    tmpStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
    
    Dim srcStream As pdStream
    Set srcStream = New pdStream
    srcStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , UBound(srcChannel.ciDataDecoded) + 1, VarPtr(srcChannel.ciDataDecoded(0))
    
    Dim numPixelsProcessed As Long
    
    'There's not a performance-friendly way to mask flags in VB, so let's just use a byte array for clarity
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
    
    'Bytes are in scanline order
    Dim x As Long, y As Long, i As Long, curByte As Byte, xStride As Long
    xStride = (numBitsInScanline + 7) \ 8
    If (xStride = 0) Then xStride = 1
    
    For y = 0 To numScanlines - 1
    
        'Reset pixel count on each scanline
        numPixelsProcessed = 0
                
        'Read through (numOfBytesPerLine) entries, pushing values into the new stream as we go
        For x = 0 To xStride - 1
        
            curByte = srcStream.ReadByte()
            
            'Parse each bit in turn
            For i = 0 To 7
                
                'Ignore empty bytes at the end of each scanline
                If (numPixelsProcessed <= numBitsInScanline) Then
                    If (bitFlags(i) = (curByte And bitFlags(i))) Then tmpStream.WriteByte 0 Else tmpStream.WriteByte 255
                    numPixelsProcessed = numPixelsProcessed + 1
                End If
                
            Next i
        Next x
        
    Next y
    
    'With a new grayscale stream assembled, overwrite the decoded data stream with the new grayscale stream
    Set srcStream = Nothing
    srcChannel.ciDataLength = tmpStream.GetStreamSize()
    ReDim srcChannel.ciDataDecoded(0 To tmpStream.GetStreamSize() - 1) As Byte
    CopyMemoryStrict VarPtr(srcChannel.ciDataDecoded(0)), tmpStream.Peek_PointerOnly(0), tmpStream.GetStreamSize()
    Set tmpStream = Nothing
    
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
Friend Function NotifySingleLayerImage(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal imgNumChannels As Integer, ByVal imgWidthPx As Long, ByVal imgHeightPx As Long, ByVal imgBitsPerChannel As Integer, ByVal imgColorMode As PSD_ColorMode, ByRef srcAddInfo As pdPSDLayerInfo) As PD_PSDResult
    
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
    m_LayerNameANSI = "Background"      'TODO: figure out what this should *actually* be - do PSDs store a name elsewhere?
    
    'Set channel data to match the parent image
    m_ChannelCount = imgNumChannels
    ReDim m_Channels(0 To m_ChannelCount - 1) As PSD_ChannelInfo
    
    'Calculate a scanline size for the current bit-depth; we need this to know how to separate the source data
    ' into usable per-channel chunks
    Dim slSize As Long
    slSize = CalculateBytesPerRow(imgBitsPerChannel, imgWidthPx)
    
    Dim i As Long
    
    'High bit-depth images may store their layer data in a special "additional info" chunk; check for this case and
    ' silently redirect our stream accordingly
    If srcAddInfo.DoesKeyExist("Lr16") Then
        Set srcStream = srcAddInfo.GetStreamForKey("Lr16")
    ElseIf srcAddInfo.DoesKeyExist("Lr32") Then
        Set srcStream = srcAddInfo.GetStreamForKey("Lr32")
    End If
    
    'Retrieve compression method; this is identical to the compression IDs used in multilayer images
    Dim cmpMethod As PSD_ChannelCompression
    cmpMethod = srcStream.ReadInt_BE()
    
    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Single-layer image compression method: " & cmpMethod
    
    Select Case cmpMethod
    
        Case ccRaw
            
            'In a raw-encoded image, channel size is fixed and easily calculated.
            Dim fixedChSize As Long
            fixedChSize = slSize * imgHeightPx
            
            'Ensure the remaining size of the data chunk is *at least* as large as our internally calculated size.
            Dim failsafeChSize As Long
            failsafeChSize = (srcStream.GetStreamSize() - srcStream.GetPosition())
            If (fixedChSize * m_ChannelCount > failsafeChSize) Then
                InternalError "NotifySingleLayerImage", "stream size is smaller than required image size (" & CStr(fixedChSize * m_ChannelCount) & " vs " & CStr(failsafeChSize) & ")"
                fixedChSize = failsafeChSize
                NotifySingleLayerImage = psd_Failure
                Exit Function
            End If
            
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
            
        'TODO!  I have not encountered these segments "in the wild" on single-layer images, as the single-layer data
        ' is typically stored in a separate Lr16/Lr32 segment.
        Case ccZip
            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "ZIP compression is not yet supported."
            
        Case ccZipWithPrediction
            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "ZIP w/prediction compression is not yet supported."
        
    End Select
    
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
    ' actually, as the spec only shows three, with a fourth mentioned in an entirely different section:
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
    If (ParseLayer < psd_Failure) Then ParseLayer = ParseLayerBlendingRanges(srcStream, warningStack, imageIsPSB)
    
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
            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Layer name (ANSI): " & m_LayerNameANSI
            
        End If
        
    End If
    
    'Since Photoshop 4.0, a whole bunch of additional (optional) layer feature blocks have been defined.
    ' These follow the end of the layer records structure, and their existence can only be caught by
    ' looking for a mismatch between the current stream pointer and the expected "end of segment"
    ' calculated earlier.  These blocks use a PNG-like "chunk" system where each chunk guarantees a
    ' 4-char signature, 4-char ID code, length, then some variable amount of data.
    If (srcStream.GetPosition() < finalPointerPos) Then ParseLayer = m_LayerInfo.ParseAdditionalLayerInfo(srcStream, warningStack, imageIsPSB, finalPointerPos)
    
    'Search additional info chunks for settings useful to PD
    If (m_LayerInfo.GetInfoCount > 0) Then
        
        'Additional info structs may be large and complicated; pdStream simplifies traversal
        Dim tmpStream As pdStream
        
        'Unicode layer names are important; they supercede the (required) ANSI layer name, if any
        If m_LayerInfo.DoesKeyExist("luni") Then
        
            'Data chunk is a 4-byte string length (chars, not bytes), followed by [length] wchars
            Set tmpStream = m_LayerInfo.GetStreamForKey("luni")
            
            Dim layerNameLengthChars As Long
            layerNameLengthChars = tmpStream.ReadLong_BE()
                    
            If (layerNameLengthChars <= ((tmpStream.GetStreamSize() - 4) \ 2)) And (layerNameLengthChars > 0) Then
                m_LayerNameUnicode = tmpStream.ReadString_Unicode_BE(layerNameLengthChars)
                m_LayerNameUnicode = Strings.TrimNull(m_LayerNameUnicode)
            End If
                    
            'Important to free stream, as it unsafely wraps the underlying chunk data
            tmpStream.StopStream
                    
            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Layer name (Unicode): " & m_LayerNameUnicode
            
        End If
        
        Set tmpStream = Nothing
    
    End If
    
    'Failsafe check to ensure our stream pointer is still aligned
    If (srcStream.GetPosition() <> finalPointerPos) Then
        warningStack.AddString "ParseLayer() may have misaligned stream pointer!  Forcibly realigning..."
        srcStream.SetPosition finalPointerPos, FILE_BEGIN
        ParseLayer = psd_Warning
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
    Else
        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Found layer blending range data: " & segmentLength & " bytes."
    End If
    
    Dim expectedFinalPointer As Long
    expectedFinalPointer = srcStream.GetPosition() + segmentLength
    
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
    Do While srcStream.GetPosition() < expectedFinalPointer
        srcStream.ReadLong_BE
        srcStream.ReadLong_BE
    Loop
    
    'Before exiting, do a failsafe check for pointer correctness
    If (expectedFinalPointer <> srcStream.GetPosition()) Then
        warningStack.AddString "ParseLayerBlendingRanges() may have misaligned stream pointer!"
        ParseLayerBlendingRanges = psd_Warning
        srcStream.SetPosition expectedFinalPointer, FILE_BEGIN
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
    Else
        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Layer mask/adjustment data section found: " & segmentLength & " bytes."
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

'PD can successfully parse the layer global mask, but the data is currently unused.
Friend Function ParseGlobalLayerMaskInfo(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean) As PD_PSDResult

    On Error GoTo InternalVBError
    
    ParseGlobalLayerMaskInfo = psd_Success
    
    'This segment provides its own length check.  This value can be zero, in which case the segment
    ' is skippable.
    Dim segmentLength As Long
    segmentLength = srcStream.ReadLong_BE()
    If (segmentLength = 0) Then
        ParseGlobalLayerMaskInfo = psd_Success
        Exit Function
    Else
        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Global layer mask section found: " & segmentLength & " bytes."
    End If
    
    'Before exiting, we'll do a failsafe check for pointer correctness
    Dim expectedFinalPointer As Long
    expectedFinalPointer = srcStream.GetPosition() + segmentLength
    
    'Overlay color space (undocumented)
    srcStream.ReadInt_BE
    
    '4 * 2 byte color components
    Dim i As Long
    For i = 0 To 3
        srcStream.ReadInt_BE
    Next i
    
    'Opacity; 0 = transparent, 100 = opaque
    srcStream.ReadInt_BE
    
    '"Kind. 0 = Color selected--i.e. inverted; 1 = Color protected;128 = use value stored per layer.
    ' This value is preferred. The others are for backward compatibility with beta versions."
    srcStream.ReadByte
    
    'Variable amounts of padding follow; ignore this and simply align the pointer manually
    If (expectedFinalPointer <> srcStream.GetPosition()) Then srcStream.SetPosition expectedFinalPointer, FILE_BEGIN
    
    Exit Function

InternalVBError:
    InternalError "ParseGlobalLayerMaskInfo", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ParseGlobalLayerMaskInfo, #" & Err.Number & ": " & Err.Description
    ParseGlobalLayerMaskInfo = psd_Failure
    
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

'These names are for debug purposes only, and should *NOT* be localized!
Private Function GetNameOfColorMode(ByVal cMode As PSD_ColorMode) As String
    Select Case cMode
        Case cm_Bitmap
            GetNameOfColorMode = "Bitmap/monochrome"
        Case cm_Grayscale
            GetNameOfColorMode = "Grayscale"
        Case cm_Indexed
            GetNameOfColorMode = "Indexed"
        Case cm_RGB
            GetNameOfColorMode = "RGB"
        Case cm_CMYK
            GetNameOfColorMode = "CMYK"
        Case cm_Multichannel
            GetNameOfColorMode = "Multichannel"
        Case cm_Duotone
            GetNameOfColorMode = "Duotone"
        Case cm_Lab
            GetNameOfColorMode = "Lab"
        Case Else
            GetNameOfColorMode = "unknown (" & CStr(cMode) & ")"
    End Select
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdPSDLayer." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdPSDLayer." & funcName & "() reported an error: " & errDescription
    End If
End Sub

Private Sub Class_Initialize()
    Set m_LayerInfo = New pdPSDLayerInfo
End Sub
