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
'Copyright 2019-2026 by Tanner Helland
'Created: 15/January/19
'Last updated: 17/April/24
'Last update: improve compatibility by always writing RGBA layer data
'
'This class contains layer-specific data pulled from a PSD file.  It is populated by a parent
' pdPSD instance.  It has no purpose outside of a PSD parsing context; for layer handling inside
' PhotoDemon itself, refer to the pdLayer class.
'
'Unless otherwise noted, all code in this class is my original work.  It is based off the
' "official" Adobe spec at this URL (link good as of January 2019):
' https://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577409_72092
'
'Adobe's custom "zip-with-prediction" compression mode is currently undocumented in the PSD spec.
' To my knowledge, the first 3rd-party to reverse-engineer it was (were?) the authors of
' Paint.NET's PSD plugin.  Their MIT-licensed work was used as a reference for this class's
' "zip-with-prediction" decoder.  Many thanks to those authors; their reference implementation is
' available here (link good as of January 2019): https://github.com/PsdPlugin/PsdPlugin
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PSD files contain a *lot* of information.  To aid debugging, you can activate "verbose" output; this will
' dump all kinds of diagnostic information to the debug log.  (Note that other PSD classes have their own
' version of this constant.)
Private Const PSD_DEBUG_VERBOSE As Boolean = False

'Layer rect, automatically translated to the proper order (Adobe uses top/left/bottom/right order)
Private m_Rect As RectL

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

Public Enum PSD_ChannelCompression
    ccRaw = 0
    ccPackBits = 1
    ccZip = 2
    ccZipWithPrediction = 3
End Enum

#If False Then
    Private Const ccRaw = 0, ccPackBits = 1, ccZip = 2, ccZipWithPrediction = 3
#End If

Public Enum PSD_LayerGroupState
    lgs_None = 0
    lgs_GroupStartOpen = 1
    lgs_GroupStartClosed = 2
    lgs_GroupEnd = 3
End Enum

#If False Then
    Private Const lgs_None = 0, lgs_GroupStartOpen = 1, lgs_GroupStartClosed = 2, lgs_GroupEnd = 3
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
    
    'Most PSD files are stored in PackBits format.  When writing a PSD file, we store RLE scanline lengths
    ' separate from the scanlines themselves; this is merely a convenience, as the length table needs to be
    ' stored as individual ushorts, which VB6 doesn't natively support.  (We also need to embed the table in
    ' different orders, depending on whether we're writing layer or composite image data, so it's helpful
    ' to maintain it separately for that case too.)
    rleTable() As Long
    
End Type

Private m_Channels() As PSD_ChannelInfo

'Additional layer info is parsed and stored by a dedicated class; this class is queried only for blocks
' that we know how to interpret.
Private m_LayerInfo As pdPSDLayerInfo

'Photoshop provides some blend modes that PD does not.  We attempt to convert all supported blend modes
' to our own internal constant, but if we fail, we'll fall back to "normal" mode.
Private m_BlendMode As PD_BlendMode

Private m_Opacity As Byte, m_Clipping As Byte, m_Flags As Byte, m_Visible As Boolean, m_LenExtraData As Long
Private m_LayerNameANSI As String, m_LayerNameUnicode As String, m_GroupState As PSD_LayerGroupState

'The final result of all channel decoding is this pdDIB object, ready for attachment to a pdLayer.
Private m_LayerDIB As pdDIB

'If the original layer contained transparency data (whether in an alpha channel or layer mask),
' this will be set to TRUE.  PD hard-converts everything to RGBA format during loading, but when exporting,
' it can be helpful to know if a fully opaque alpha channel came from the source file or from us.
Private m_LayerHasAlpha As Boolean

'Multichannel images represent a variation on the above rule.  Instead of creating one DIB, they create several -
' one for each channel.  The user can then decide how to assemble these into a final image.  (These values are
' unused in any other color mode.)
Private m_FinalChannelDIBs() As pdDIB, m_FinalChannelIDs() As PSD_ChannelID, m_FinalChannelCount As Long

'If layer mask data is present, these trackers will be set to TRUE and relevant mask parameters will be auto-filled
Private m_LayerMaskExists As Boolean, m_RealLayerMaskExists As Boolean
Private Type LayerMaskData
    lmRect As RectL
    lmDefaultColor As Byte
    lmFlags As Byte
    lmParameters As Byte
    lmUserMaskDensity As Byte
    lmUserMaskFeather As Double
    lmVectorMaskFeather As Double   'This and the next parameter appear in the file in opposite order; I've reversed
    lmVectorMaskDensity As Byte     ' them here to avoid unwanted struct padding complications
    lmPadding1 As Byte
    lmParametersLoaded As Boolean
End Type
Private m_LayerMaskData As LayerMaskData, m_RealLayerMaskData As LayerMaskData

'On normal layers, layer masks get "merged" into that layer's alpha channel (because PD doesn't support
' layer masks yet).  However, I am working toward full mask support in stages, so right now, there is a
' special case for layer *groups* with masks - these masks will be correctly saved into a dedicated
' array at load-time.  At image assembly time, these can then be transferred into the placeholder layer's
' pdLayer.GetLayerMask object - and someday, the same will be done for "normal" layers too.
'
'Anyway, at present you need to query m_MaskWasMergedIntoDIB to know how to handle the layer mask.
' (If it was *not* merged into the layer DIB, retrieve it separately - this is *only* true for group markers.)
Private m_MaskWasMergedIntoDIB As Boolean, m_MaskRect As RectF
Private m_LayerMaskBytes() As Byte, m_sizeOfLayerMask As Long

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

Friend Function GetLayerClipping() As Boolean
    GetLayerClipping = (m_Clipping = 1)
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

Friend Function GetMaskDefaultColor() As Byte
    GetMaskDefaultColor = m_LayerMaskData.lmDefaultColor
End Function

Friend Function GetMaskExists() As Boolean
    GetMaskExists = m_LayerMaskExists Or m_RealLayerMaskExists
End Function

Friend Function GetMaskMergedIntoDIB() As Boolean
    GetMaskMergedIntoDIB = m_MaskWasMergedIntoDIB
End Function

Friend Function GetMaskRectF() As RectF
    GetMaskRectF = m_MaskRect
End Function

Friend Function GetMaskUnmergedPtr() As Long
    If (m_sizeOfLayerMask <> 0) Then GetMaskUnmergedPtr = VarPtr(m_LayerMaskBytes(0))
End Function

Friend Function GetMaskUnmergedSize() As Long
    GetMaskUnmergedSize = m_sizeOfLayerMask
End Function

Friend Function GetMultiChannelLayerDIB(ByVal lIndex As Long) As pdDIB
    Set GetMultiChannelLayerDIB = m_FinalChannelDIBs(lIndex)
End Function

Friend Function GetMultiChannelLayerID(ByVal lIndex As Long) As PSD_ChannelID
    GetMultiChannelLayerID = m_FinalChannelIDs(lIndex)
End Function

Friend Function HasAlpha() As Boolean
    HasAlpha = m_LayerHasAlpha
End Function

'If this layer marks the beginning of a layer group, this will return TRUE.
Friend Function IsLayerGroupStart() As Boolean
    IsLayerGroupStart = ((m_GroupState = lgs_GroupStartOpen) Or (m_GroupState = lgs_GroupStartClosed))
End Function

'If this layer marks the end of a layer group, this will return TRUE.
Friend Function IsLayerGroupEnd() As Boolean
    IsLayerGroupEnd = (m_GroupState = lgs_GroupEnd)
End Function

'After channel data has been decoded (decompressed), it needs to be assembled into a final usable image.
' At present, this step is limited to certain color modes and bit-depths, although I am actively expanding coverage.
' If you have a PSD file that doesn't work, this step is likely the culprit - attach the PSD file to a new Issue
' at GitHub and I will investigate further.
'
'(Also note: if the DecodeChannels step fails for a given channel, don't waste time calling this step;
' it won't work if the underlying channel data hasn't been decoded properly!)
Friend Function ConstructImage(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD, Optional ByVal loadGroupMasksAsPixels As Boolean = False) As PD_PSDResult
    
    On Error GoTo InternalVBError
    
    ConstructImage = psd_Success
    
    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "attempting to construct layer: " & Me.GetLayerName()
    
    m_ColorMode = imgColorMode
    
    'AFAIK, all bit-depths and color modes are now fully covered by PhotoDemon's importer.
    'If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Bit-depth and color mode are supported; attempting to decode now..."
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Many images may not contain alpha channels.  For these images, we don't want to waste time premultiplying
    ' alpha after converting the data to RGBA.  Image constructor will modify this value if alpha is encountered.
    m_LayerHasAlpha = False
    
    'Image construction varies greatly depending on the format of the input data.  To better handle the unique needs
    ' of each color format, handling is sorted by color mode.  (This allows for acceleration that would otherwise
    ' be impractical in a "universal" image constructor.)
    If (imgColorMode = cm_RGB) Then
        ConstructImage = ConstructImage_RGB(warningStack, imageIsPSB, bitsPC, imgColorMode, srcProfile, parentPSD, m_LayerHasAlpha)
        
    ElseIf (imgColorMode = cm_CMYK) Then
        ConstructImage = ConstructImage_CMYK(warningStack, imageIsPSB, bitsPC, imgColorMode, srcProfile, parentPSD, m_LayerHasAlpha)
        
    ElseIf (imgColorMode = cm_Grayscale) Or (imgColorMode = cm_Indexed) Or (imgColorMode = cm_Duotone) Or (imgColorMode = cm_Bitmap) Then
        ConstructImage = ConstructImage_SingleChannel(warningStack, imageIsPSB, bitsPC, imgColorMode, srcProfile, parentPSD, m_LayerHasAlpha)
    
    ElseIf (imgColorMode = cm_Lab) Then
        ConstructImage = ConstructImage_Lab(warningStack, imageIsPSB, bitsPC, imgColorMode, srcProfile, parentPSD, m_LayerHasAlpha)
        
    ElseIf (imgColorMode = cm_Multichannel) Then
        ConstructImage = ConstructImage_Multichannel(warningStack, imageIsPSB, bitsPC, imgColorMode, srcProfile, parentPSD)
        
    End If
    
    'After all meaningful color channels have been assembled, let's perform one final pass to deal with any layer
    ' masks present in the image.  (The data in layer masks is not contingent on the image's color mode, but they
    ' *are* contingent on color depth, so we can deal with them in a roughly uniform manner regardless of how
    ' channel data was encoded.)
    '
    'Note that this step is especially important for layer *groups*, as they will not have any layer pixel data,
    ' but they *may* have a mask associated with them.
    If (ConstructImage < psd_Failure) Then ConstructImage = ConstructLayerMasks(warningStack, imageIsPSB, bitsPC, imgColorMode, parentPSD, m_LayerHasAlpha, loadGroupMasksAsPixels)
    
    'If meaningful alpha data exists in the file (in the form of alpha channel(s) or masks), premultiply it now.
    ' (Note that we must check for the presence of a layer DIB; layer groups can potentially construct a mask
    ' in this step, but not an actual layer DIB.)
    If (Not m_LayerDIB Is Nothing) Then
        If m_LayerHasAlpha Then
            m_LayerDIB.SetAlphaPremultiplication True
        Else
            m_LayerDIB.SetInitialAlphaPremultiplicationState True
        End If
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

Private Function ConstructImage_CMYK(ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD, ByRef imageHasAlpha As Boolean) As PD_PSDResult

    ConstructImage_CMYK = psd_Success
    On Error GoTo InternalVBError
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Interpretation of the underlying data depends on bit-depth, obviously
    Dim bytesPerChannel As Long
    bytesPerChannel = (bitsPC \ 8)
    
    'Set a flag to determine ICC profile availability.  If I had my way, PhotoDemon wouldn't load
    ' CMYK images without an embedded color profile as the results will be borderline garbage.
    ' This is not generally an issue because Photoshop *never* writes CMYK images without embedded
    ' color profiles, but 3rd-party PSD encoders are a different story.
    '
    'That said, we try to make life easy for users, and given the number of bad 3rd-party PSDs
    ' floating around the web, we want to try and load bad CMYK data as best we can.  Currently,
    ' that means using a generic CMYK profile that ships with PhotoDemon.  It is a public-domain
    ' CMYK transfer built by Graeme Gill of ArgyllCMS fame, and originally shared in the LittleCMS
    ' mailing list (https://sourceforge.net/p/lcms/mailman/message/32755884/).
    '
    'PD extracts this profile to the Data/ColorProfiles at first-run, and if we can't find it
    ' there, we simply abandon ship.
    Dim iccAvailable As Boolean
    iccAvailable = (Not srcProfile Is Nothing)
    If iccAvailable Then iccAvailable = srcProfile.HasICCData()
    If (Not iccAvailable) Then
    
        warningStack.AddString "CMYK layer does not contain a usable ICC profile.  Looking for default PD CMYK profile..."
        
        Dim builtInProfilePath As String
        builtInProfilePath = UserPrefs.GetColorProfilePath() & "cmyk.icm"
        
        If Files.FileExists(builtInProfilePath) Then
            Set srcProfile = New pdICCProfile
            iccAvailable = srcProfile.LoadICCFromFile(builtInProfilePath)
            If iccAvailable Then iccAvailable = srcProfile.HasICCData()
        End If
        
        'If loading the built-in CMYK profile somehow failed, abandon the load process
        If iccAvailable Then
            warningStack.AddString "Default CMYK profile found.  Proceeding with PSD import..."
        Else
            warningStack.AddString "Default CMYK missing or broken.  PSD import abandoned."
            ConstructImage_CMYK = psd_Failure
            Exit Function
        End If
        
    End If
    
    'Note that alpha channels can be explicitly marked (with an ID of -1) or they can simply be "extra" channels
    ' without an obvious RGB mapping.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(i).ciID = cidAlpha) Or (m_Channels(i).ciID > cidKey) Then imageHasAlpha = True
    Next i
    
    'Layer dimensions determine the size of the constructed layer, obviously; layer sizes are completely independent
    ' of the parent image.
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    
    'Failsafe check for bad layer dimensions.  See detailed comments in ConstructImage_RGB(),
    ' which is where this layer group exception is most relevant.
    If (layerWidth <= 0) Or (layerHeight <= 0) Then
        If (Not Me.IsLayerGroupStart) And (Not Me.IsLayerGroupEnd) Then
            InternalError "ConstructImage_CMYK", "bad layer width/height"
            ConstructImage_CMYK = psd_Warning
        Else
            ConstructImage_CMYK = psd_Success
        End If
        Exit Function
    End If
    
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
    If imageHasAlpha Then srcLineSize = channelScanlineSize * 5 Else srcLineSize = channelScanlineSize * 4
    ReDim tmpPlanarBuffer(0 To srcLineSize - 1) As Byte
    
    'Prep source and destination ICC profiles
    Dim lcmsSrcProfile As pdLCMSProfile, lcmsDstProfile As pdLCMSProfile, lcmsTransform As pdLCMSTransform
    
    If iccAvailable Then
        Set lcmsSrcProfile = New pdLCMSProfile
        lcmsSrcProfile.CreateFromPDICCObject srcProfile
        Set lcmsDstProfile = New pdLCMSProfile
        lcmsDstProfile.CreateSRGBProfile
    End If
    
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
    If imageHasAlpha Then srcColorMode = srcColorMode Or FLAG_ALPHAPRESENT
    
    'Photoshop defines CMYK data as vanilla (e.g. 0 is white, 1 is black)
    srcColorMode = srcColorMode Or FLAG_MINISWHITE
    
    If iccAvailable Then
        Set lcmsTransform = New pdLCMSTransform
        lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile, srcColorMode, TYPE_BGRA_8, lcmsSrcProfile.GetRenderingIntent(), cmsFLAGS_BLACKPOINTCOMPENSATION Or cmsFLAGS_COPY_ALPHA
    End If
    
    'Iterate through each scanline in the image
    If iccAvailable Then
        
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
        
    End If
    
    'Free all source channel data to relieve memory pressure
    If (cIndex >= 0) Then Erase m_Channels(cIndex).ciDataDecoded
    If (mIndex >= 0) Then Erase m_Channels(mIndex).ciDataDecoded
    If (yIndex >= 0) Then Erase m_Channels(yIndex).ciDataDecoded
    If (kIndex >= 0) Then Erase m_Channels(kIndex).ciDataDecoded
    If (aIndex >= 0) Then Erase m_Channels(aIndex).ciDataDecoded
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Image assembly took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ConstructImage_CMYK", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ConstructImage_CMYK, #" & Err.Number & ": " & Err.Description
    ConstructImage_CMYK = psd_Failure
    
End Function

Private Function ConstructImage_Lab(ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD, ByRef imageHasAlpha As Boolean) As PD_PSDResult

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
    
    'Note that alpha channels can be explicitly marked (with an ID of -1) or they can simply be "extra" channels
    ' without an obvious color channel mapping.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(i).ciID = cidAlpha) Or (m_Channels(i).ciID > cidBlue) Then imageHasAlpha = True
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
    
    'Failsafe check for bad layer dimensions.  See detailed comments in ConstructImage_RGB(),
    ' which is where this layer group exception is most relevant.
    If (layerWidth <= 0) Or (layerHeight <= 0) Then
        If (Not Me.IsLayerGroupStart) And (Not Me.IsLayerGroupEnd) Then
            InternalError "ConstructImage_Lab", "bad layer width/height"
            ConstructImage_Lab = psd_Warning
        Else
            ConstructImage_Lab = psd_Success
        End If
        Exit Function
    End If
    
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
    If imageHasAlpha Then
        
        Dim tmpBytes() As Byte, tmpSA As SafeArray1D
    
        Dim chIndex As Long
        For chIndex = 0 To Abs(m_ChannelCount) - 1
            
            With m_Channels(chIndex)
                
                If (Not .ciDecodedOK) Then GoTo SkipChannel
                If (.ciDataLength = 0) Then GoTo SkipChannel
                If (.ciID < cidAlpha) Then GoTo SkipChannel
                
                'Make sure this is an alpha channel
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
                    
                    '32-bit channel
                    ElseIf (bytesPerChannel = 4) Then
                            
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
                            tmpBytes(x * 4 + 3) = Int(tmpSingle * 255!)
                            
                        Next x
                        Next y
                        
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
Private Function ConstructImage_Multichannel(ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD) As PD_PSDResult

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
    
    'Failsafe check for bad layer dimensions.  See detailed comments in ConstructImage_RGB(),
    ' which is where this layer group exception is most relevant.
    If (layerWidth <= 0) Or (layerHeight <= 0) Then
        If (Not Me.IsLayerGroupStart) And (Not Me.IsLayerGroupEnd) Then
            InternalError "ConstructImage_Multichannel", "bad layer width/height"
            ConstructImage_Multichannel = psd_Warning
        Else
            ConstructImage_Multichannel = psd_Success
        End If
        Exit Function
    End If
    
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
    Dim tmpLongs() As Long, tmpSA As SafeArray1D
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
    Dim imgPalette() As RGBQuad, palSize As Long
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
            If (.ciID < cidAlpha) Then GoTo SkipChannel
            
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
                m_FinalChannelDIBs(m_FinalChannelCount).SetInitialAlphaPremultiplicationState True
                
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

Private Function ConstructImage_RGB(ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD, ByRef imageHasAlpha As Boolean) As PD_PSDResult

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
    If iccAvailable Then iccAvailable = srcProfile.HasICCData() And ColorManagement.UseEmbeddedICCProfiles()
    
    'Note that alpha channels can be explicitly marked (with an ID of -1) or they can simply be "extra" channels
    ' with a positive ID, but no corresponding RGB mapping.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(i).ciID = cidAlpha) Or (m_Channels(i).ciID > cidBlue) Then imageHasAlpha = True
    Next i
    
    'Layer dimensions determine the size of the constructed layer, obviously; layer sizes are completely independent
    ' of the parent image.
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    
    'Failsafe check for bad layer dimensions.  Note that this case *will* be triggered by layer groups
    ' with a layer mask, and that's okay - we provide special handling for this in the mask retrieval code.
    ' (We still want to exit here, however, because there will not be any useable layer pixel data beyond
    ' this point.)
    If (layerWidth <= 0) Or (layerHeight <= 0) Then
        If (Not Me.IsLayerGroupStart) And (Not Me.IsLayerGroupEnd) Then
            InternalError "ConstructImage_RGB", "bad layer width/height"
            ConstructImage_RGB = psd_Warning
        Else
            ConstructImage_RGB = psd_Success
        End If
        Exit Function
    End If
    
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
            With m_Channels(x)
                If .ciDecodedOK Then
                    If (.ciID = cidBlue) Then
                        bIndex = x
                    ElseIf (.ciID = cidGreen) Then
                        gIndex = x
                    ElseIf (.ciID = cidRed) Then
                        rIndex = x
                    ElseIf (.ciID = cidAlpha) Or (.ciID > cidBlue) Then
                        aIndex = x
                    End If
                End If
            End With
        Next x
        
        'Prep a temporary buffer.  We'll assemble each scanline into a contiguous planar chunk in this buffer,
        ' then color-manage, interleave, swizzle, and (when necessary) swap endianness while merging the data
        ' into the final image buffer.
        Dim tmpPlanarBuffer() As Byte, srcLineSize As Long
        If imageHasAlpha Then srcLineSize = channelScanlineSize * 4 Else srcLineSize = channelScanlineSize * 3
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
        If imageHasAlpha Then srcColorMode = srcColorMode Or FLAG_ALPHAPRESENT
        
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
        
        'Channel data is potentially huge; free all of it before exiting
        If (rIndex >= 0) Then Erase m_Channels(rIndex).ciDataDecoded
        If (gIndex >= 0) Then Erase m_Channels(gIndex).ciDataDecoded
        If (bIndex >= 0) Then Erase m_Channels(bIndex).ciDataDecoded
        If (aIndex >= 0) Then Erase m_Channels(aIndex).ciDataDecoded
        
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
                If (.ciID < cidAlpha) Then GoTo SkipChannel
                
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
                ' attached, so this is only relevant to bad 3rd-party implementations.)
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
    
    Exit Function
    
InternalVBError:
    InternalError "ConstructImage_RGB", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ConstructImage_RGB, #" & Err.Number & ": " & Err.Description
    ConstructImage_RGB = psd_Failure
    
End Function

'Grayscale, duotone, and indexed images share a unified load function.  (Monochrome images are also handled here,
' with their encoding having been silently "upsampled" to grayscale in a previous stage.)
Private Function ConstructImage_SingleChannel(ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef srcProfile As pdICCProfile, ByRef parentPSD As pdPSD, ByRef imageHasAlpha As Boolean) As PD_PSDResult

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
    
    'Note that alpha channels can be explicitly marked (with an ID of -1) or they can simply be "extra" channels
    ' without an obvious RGB mapping.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(i).ciID = cidAlpha) Or (m_Channels(i).ciID > cidIndexed) Then imageHasAlpha = True
    Next i
    
    'Layer dimensions determine the size of the constructed layer, obviously; layer sizes are completely independent
    ' of the parent image.
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    
    'Failsafe check for bad layer dimensions.  See detailed comments in ConstructImage_RGB(),
    ' which is where this layer group exception is most relevant.
    If (layerWidth <= 0) Or (layerHeight <= 0) Then
        If (Not Me.IsLayerGroupStart) And (Not Me.IsLayerGroupEnd) Then
            InternalError "ConstructImage_SingleChannel", "bad layer width/height"
            ConstructImage_SingleChannel = psd_Warning
        Else
            ConstructImage_SingleChannel = psd_Success
        End If
        Exit Function
    End If
    
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
            If (.ciID < cidAlpha) Then GoTo SkipChannel
            
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
    If imageHasAlpha Then
        
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
                    
                    'Immediately free the source channel data to relieve memory pressure
                    Erase .ciDataDecoded
                    
                '/End alpha channel found
                End If
                
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
    
    Exit Function
    
InternalVBError:
    InternalError "ConstructImage_SingleChannel", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ConstructImage_SingleChannel, #" & Err.Number & ": " & Err.Description
    ConstructImage_SingleChannel = psd_Failure
    
End Function

'Retrieve layer masks, if any.  The original design of this function involved merging layer mask data
' directly into the alpha channel of this layer (because PhotoDemon didn't layer masks yet).  Now, however,
' there are some contingencies for placing layer mask data directly into its own structure.  As this function
' is rapidly evolving, be sure to check comments before modifying.
Private Function ConstructLayerMasks(ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef parentPSD As pdPSD, ByRef imageHasAlpha As Boolean, Optional ByVal loadGroupMasksAsPixels As Boolean = False) As PD_PSDResult
    
    Const FUNC_NAME As String = "ConstructLayerMasks"
    
    ConstructLayerMasks = psd_Success
    On Error GoTo InternalVBError
    
    'Perform a few failsafe checks to detect any issues that arose during channel assembly
    Dim exitEarly As Boolean: exitEarly = False
    If (m_LayerDIB Is Nothing) Then
        exitEarly = True
    Else
        If (m_LayerDIB.GetDIBWidth <= 0) Or (m_LayerDIB.GetDIBHeight <= 0) Then exitEarly = True
    End If
    
    'Multichannel images have unique constraints; I have not seen a multichannel + mask image "in the wild",
    ' and I'd like to have one for testing before enabling mask + multichannel behavior.
    If (imgColorMode = cm_Multichannel) Then exitEarly = True
    
    'After all that "exit early" business, we now need to handle a special case - layer groups with masks.
    ' Layer groups will not have a layer image, but they *may* have a mask.  This is a valid case, and we
    ' want to cover it (as best we can, at least until we support proper groups and masks).
    Dim specialLayerGroupMaskCase As Boolean
    specialLayerGroupMaskCase = False
    
    If exitEarly Then
        If (Me.IsLayerGroupStart Or Me.IsLayerGroupEnd) Then
            exitEarly = False
            specialLayerGroupMaskCase = True
        End If
    End If
    
    If exitEarly Then
        ConstructLayerMasks = psd_Warning
        Exit Function
    End If
    
    'Before proceeding further, see if this layer even contains masks
    Dim i As Long, maskFound As Boolean
    For i = 0 To Abs(m_ChannelCount) - 1
        If (m_Channels(i).ciID = cidUserLayerMask) Or (m_Channels(i).ciID = cidRealUserLayerMask) Then maskFound = True
    Next i
    
    If (Not maskFound) Then
        ConstructLayerMasks = psd_Success
        Exit Function
    Else
        imageHasAlpha = maskFound
    End If
    
    'If we're still here, one (or more) masks were found.
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Interpretation of the underlying data depends on bit-depth, obviously
    Dim bytesPerChannel As Long
    bytesPerChannel = (bitsPC \ 8)
    
    'Unlike color channels, which inherit the dimensions of their parent layer, layer masks provide their
    ' own dimensions.  These are not constrained by layer dimensions, and areas exclusive to the mask rect
    ' have a pre-defined value of 0 or 255 (defined by a separate value in the underlying PSD).
    Dim maskLeft As Long, maskTop As Long
    Dim maskWidth As Long, maskHeight As Long
    
    'Prep pointers into the destination image; we'll obviously need these to properly apply the layer mask.
    ' (Note that we don't yet have a DIB object for masks on layer groups - this will be handled as a
    ' special case, later in the function.)
    Dim x As Long, y As Long
    Dim tmpBytes() As Byte, tmpSA As SafeArray1D
    Dim scanStart As Long, scanWidth As Long, xLoopEnd As Long, ySrcOffset As Long
    If (Not m_LayerDIB Is Nothing) Then
        scanStart = m_LayerDIB.GetDIBPointer()
        scanWidth = m_LayerDIB.GetDIBStride()
    End If
    
    'Iterate through all channels, looking for mask data
    Dim chIndex As Long
    For chIndex = 0 To Abs(m_ChannelCount) - 1
        
        With m_Channels(chIndex)
            
            'As with normal layers, skip any channels that had decoding failures
            If (Not .ciDecodedOK) Then GoTo SkipChannel
            If (.ciDataLength = 0) Then GoTo SkipChannel
            If (.ciID >= cidAlpha) Then GoTo SkipChannel
            If (.ciID <> cidUserLayerMask) And (.ciID <> cidRealUserLayerMask) Then GoTo SkipChannel
            
            'Boundaries are unique to the different types of layer mask (determined by ID)
            If (.ciID = cidUserLayerMask) Then
                maskLeft = m_LayerMaskData.lmRect.Left
                maskTop = m_LayerMaskData.lmRect.Top
                maskWidth = m_LayerMaskData.lmRect.Right - m_LayerMaskData.lmRect.Left
                maskHeight = m_LayerMaskData.lmRect.Bottom - m_LayerMaskData.lmRect.Top
            ElseIf (.ciID = cidRealUserLayerMask) Then
                maskLeft = m_RealLayerMaskData.lmRect.Left
                maskTop = m_RealLayerMaskData.lmRect.Top
                maskWidth = m_RealLayerMaskData.lmRect.Right - m_RealLayerMaskData.lmRect.Left
                maskHeight = m_RealLayerMaskData.lmRect.Bottom - m_RealLayerMaskData.lmRect.Top
            End If
            
            'Special flags inside the mask section can indicate that a mask is "disabled";
            ' if it is, we just want to ignore it
            Dim maskDisabled As Boolean
            maskDisabled = False
            If (.ciID = cidUserLayerMask) Then
                If ((m_LayerMaskData.lmFlags And &H2&) <> 0) Then maskDisabled = True
            ElseIf (.ciID = cidRealUserLayerMask) Then
                If ((m_RealLayerMaskData.lmFlags And &H2&) <> 0) Then maskDisabled = True
            End If
            
        End With
            
        If maskDisabled Then GoTo SkipChannel
        
        'We now enter a branch that is a temporary fix due to PhotoDemon's lack of full mask support.
        '
        'In the future, we will want to extract layer masks to their own special mask structure.
        ' As a preliminary test, PhotoDemon *will* do this for layer groups right now (but *not* for
        ' normal layers).  The mask will then be stored with the layer group and exported back out
        ' to PSD as desired.
        '
        'But because PD's viewport renderer (and many other parts of the program) do not have full
        ' support for masks yet, "normal" layers instead have their mask(s) merged into the layer's
        ' alpha channel.  This guarantees "correct" appearances in PD, but note that this layer will
        ' be permanently modified by this change.
        
        'If this is a *layer group*, we have not created a layer DIB yet (because no layer image data exists).
        ' As such, this is our chance to finally create one at the size specified by the mask.  Note that we
        ' only do this if the caller specified "load layer group masks as pixels" (which is very helpful for
        ' analyzing mask contents, but not at all helpful for correct image display).
        If (specialLayerGroupMaskCase And loadGroupMasksAsPixels) Then
            If (Not m_LayerDIB Is Nothing) Then InternalError FUNC_NAME, "layer image exists for group?"
            Set m_LayerDIB = New pdDIB
            m_LayerDIB.CreateBlank maskWidth, maskHeight, 32, vbWhite, 255
            scanStart = m_LayerDIB.GetDIBPointer()
            scanWidth = m_LayerDIB.GetDIBStride()
        End If
        
        'Populate mask rect, and note that it's always clipped to this layer's size.
        ' (Group masks are a little weird this way - TODO!)
        With m_MaskRect
            .Left = maskLeft
            .Top = maskTop
            .Width = maskWidth
            .Height = maskHeight
        End With
        
        'There are a lot of different ways to go about applying the mask to the image.  AFAIK, no guarantees
        ' are made about things like rect intersections between the mask and layer, although common sense
        ' suggests the layer is typically inclusive of the mask... but this is a PSD, so expect the worst.
        
        'To simplify handling (especially for the case where a mask obscures a large portion of the layer
        ' via the "default color" mask byte), let's prep a full buffer at the size of the layer, fill it with
        ' the default value, then simply overwrite the relevant rect with the mask's actual data.
        Dim tmpBuffer() As Byte
        
        With m_Channels(chIndex)
        
            'Normal layers can use the underlying DIB dimensions for their reference; layer groups may not have
            ' such a DIB, so we need to prep them differently
            If (specialLayerGroupMaskCase And (Not loadGroupMasksAsPixels)) Then
                ReDim tmpBuffer(0 To maskWidth - 1, 0 To maskHeight - 1) As Byte
                If (.ciID = cidUserLayerMask) Then
                    FillMemory VarPtr(tmpBuffer(0, 0)), maskWidth * maskHeight, m_LayerMaskData.lmDefaultColor
                ElseIf (.ciID = cidRealUserLayerMask) Then
                    FillMemory VarPtr(tmpBuffer(0, 0)), maskWidth * maskHeight, m_RealLayerMaskData.lmDefaultColor
                End If
            Else
                ReDim tmpBuffer(0 To m_LayerDIB.GetDIBWidth - 1, 0 To m_LayerDIB.GetDIBHeight - 1) As Byte
                If (.ciID = cidUserLayerMask) Then
                    FillMemory VarPtr(tmpBuffer(0, 0)), m_LayerDIB.GetDIBWidth * m_LayerDIB.GetDIBHeight, m_LayerMaskData.lmDefaultColor
                ElseIf (.ciID = cidRealUserLayerMask) Then
                    FillMemory VarPtr(tmpBuffer(0, 0)), m_LayerDIB.GetDIBWidth * m_LayerDIB.GetDIBHeight, m_RealLayerMaskData.lmDefaultColor
                End If
            End If
                
            'There's a bit in the layer flags that supposedly indicates whether the mask position is in absolute
            ' vs relative coordinates (to its parent layer).  Limited testing shows that this bit doesn't seem
            ' to matter, but if someone can turn up an image that shows otherwise, I'm happy to modify this code.
            Dim rIntersect As RectL
            Dim srcMaskRect As RectL
            If (.ciID = cidUserLayerMask) Then srcMaskRect = m_LayerMaskData.lmRect Else srcMaskRect = m_RealLayerMaskData.lmRect
            
            'By default, mask rects are in IMAGE coordinates - not LAYER coordinates.  Modify this now by adding layer
            ' offsets to the mask rectangle.
            srcMaskRect.Left = srcMaskRect.Left - m_Rect.Left
            srcMaskRect.Right = srcMaskRect.Right - m_Rect.Left
            srcMaskRect.Top = srcMaskRect.Top - m_Rect.Top
            srcMaskRect.Bottom = srcMaskRect.Bottom - m_Rect.Top
        
        End With
        
        'On a normal layer, ensure the layer and mask overlap (and also, *where* they overlap).
        ' (On a layer group mask, this step is technically irrelevant, because the layer DIB will have
        ' been created according to the mask's dimensions.)
        Dim intersectOK As Boolean
        If specialLayerGroupMaskCase Then
            
            intersectOK = True
            With rIntersect
                .Left = 0
                .Top = 0
                .Right = maskWidth - 1
                .Bottom = maskHeight - 1
            End With
            
        Else
            
            intersectOK = PDMath.IntersectRectL(rIntersect, PopulateRectL(0, 0, m_LayerDIB.GetDIBWidth - 1, m_LayerDIB.GetDIBHeight - 1), srcMaskRect)
            
            'Validate the final rect, just in case.  (In the next step we need to use it to index into an array,
            ' and overwrites there are catastrophic.)
            '
            '(This change was added in June 2023 to fix crashes on the test-case image "vector-mask2.psd" from the psd-tools repo.)
            If ((rIntersect.Right - rIntersect.Left) >= maskWidth) Then rIntersect.Right = rIntersect.Left + (maskWidth - 1)
            If ((rIntersect.Bottom - rIntersect.Top) >= maskHeight) Then rIntersect.Bottom = rIntersect.Top + (maskHeight - 1)
            
        End If
        
        If intersectOK Then
        
            'We now need to merge the mask data with the "default color" array.  Handling varies by
            ' parent layer color depth, but at the end of this we'll have an 8-bit copy the correctly
            ' represents the mask (regardless of invert status).
            Dim mOffsetLeft As Long, mOffsetTop As Long
            If specialLayerGroupMaskCase Then
                mOffsetLeft = 0
                mOffsetTop = 0
            Else
                mOffsetLeft = srcMaskRect.Left
                mOffsetTop = srcMaskRect.Top
            End If
            
            With m_Channels(chIndex)
            
                '8-bit channel
                If (bytesPerChannel = 1) Then
                    
                    For y = rIntersect.Top To rIntersect.Bottom
                    For x = rIntersect.Left To rIntersect.Right
                        tmpBuffer(x, y) = .ciDataDecoded((y - mOffsetTop) * maskWidth + (x - mOffsetLeft))
                    Next x
                    Next y
                    
                '16-bit channel
                ElseIf (bytesPerChannel = 2) Then
                
                    For y = rIntersect.Top To rIntersect.Bottom
                    For x = rIntersect.Left To rIntersect.Right
                        tmpBuffer(x, y) = .ciDataDecoded((y - mOffsetTop) * maskWidth * 2 + (x - mOffsetLeft) * 2)
                    Next x
                    Next y
                    
                '32-bit (grayscale only, hopefully; other color modes may not produce desirable results)
                Else
                    
                    Dim tmpSingle As Single
                    For y = rIntersect.Top To rIntersect.Bottom
                    For x = rIntersect.Left To rIntersect.Right
                        
                        'Extract the source data into a single-type value
                        GetMem4_Ptr VarPtr((y - mOffsetTop) * maskWidth * 2 + (x - mOffsetLeft) * 2), VarPtr(tmpSingle)
                        
                        'Clamp for safety before assigning
                        If (tmpSingle < 0!) Then tmpSingle = 0!
                        If (tmpSingle > 1!) Then tmpSingle = 1!
                        tmpBuffer(x, y) = Int(tmpSingle * 255!)
                        
                    Next x
                    Next y
                    
                End If
            
            End With
        
        'Rects do not overlap; only the "default color" of the mask will be used
        Else
            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Mask found, but does not overlap the layer!"
        End If
        
        'Immediately free the source channel data to relieve memory pressure
        Erase m_Channels(chIndex).ciDataDecoded
        
        'On layer masks, we now want to populate the dedicated layer mask array; on image layers,
        ' we'll instead apply the mask directly to the target layer.
        If (specialLayerGroupMaskCase And (Not loadGroupMasksAsPixels)) Then
            
            'Just copy the temporary array to a dedicated, module-level mask array
            m_sizeOfLayerMask = maskWidth * maskHeight
            ReDim m_LayerMaskBytes(0 To m_sizeOfLayerMask - 1) As Byte
            VBHacks.CopyMemoryStrict VarPtr(m_LayerMaskBytes(0)), VarPtr(tmpBuffer(0, 0)), m_sizeOfLayerMask
            
            m_MaskWasMergedIntoDIB = False
            
        'On image layers, actually applying the mask is very simple.  We simply want to multiply any
        ' existing alpha values by the values found in the mask.
        Else
            
            Dim layerWidth As Long, layerHeight As Long
            layerWidth = m_LayerDIB.GetDIBWidth
            layerHeight = m_LayerDIB.GetDIBHeight
            
            m_LayerDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
            xLoopEnd = layerWidth - 1
            
            Const ONE_DIV_255 As Single = 1! / 255!
            
            For y = 0 To layerHeight - 1
                tmpSA.pvData = scanStart + (scanWidth * y)
                ySrcOffset = y * layerWidth
            For x = 0 To xLoopEnd
                tmpSingle = CSng(tmpBuffer(x, y)) * ONE_DIV_255
                tmpBytes(x * 4 + 3) = Int(CSng(tmpBytes(x * 4 + 3)) * tmpSingle)
            Next x
            Next y
            
            m_LayerDIB.UnwrapArrayFromDIB tmpBytes
            
            m_MaskWasMergedIntoDIB = True
            
        End If
        
SkipChannel:
    Next chIndex
    
    If PSD_DEBUG_VERBOSE Then
        PDDebug.LogAction "Mask construction + merge took: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    Exit Function
    
InternalVBError:
    InternalError FUNC_NAME, "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in " & FUNC_NAME & ", #" & Err.Number & ": " & Err.Description
    ConstructLayerMasks = psd_Failure
    
End Function

'After channel data has been retrieved from the PSD, we must decode it into something usable by PD.
' That's what this function handles.  If the underlying data is irrelevant to PD (e.g. adjustment or effect layers),
' this function simply frees the associated layer data to cut down on memory requirements.  Layers with relevant
' image data, however, will produce a usable pdDIB object from this step.
Friend Function DecodeChannels(ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal bitsPC As Long, ByVal imgColorMode As PSD_ColorMode, ByRef parentPSD As pdPSD) As PD_PSDResult
    
    DecodeChannels = psd_Success
    
    On Error GoTo InternalVBError
    
    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "attempting to decode channels for layer: " & Me.GetLayerName()
    
    'Decoded data size is contingent on the layer's rectangle
    Dim origLayerWidth As Long, origLayerHeight As Long
    Dim layerWidth As Long, layerHeight As Long
    origLayerWidth = m_Rect.Right - m_Rect.Left
    origLayerHeight = m_Rect.Bottom - m_Rect.Top
    
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
        
        'Same goes for layer dimensions
        layerWidth = origLayerWidth
        layerHeight = origLayerHeight
        
        With m_Channels(i)
            
            'Failsafe reset of decoding results; this tells subsequent steps to ignore this layer's data
            .ciDecodedOK = False
            
            'Some channels contribute nothing to the final image (and are zero-length).  Ignore them.
            If (.ciDataLength = 0) Then
                .ciDecodedOK = False
                GoTo SkipChannel
            End If
            
            'Layer masks are *not* required to have the same dimensions as their parent layer.  (In fact, they are
            ' often much smaller.)  If one is found, silently substitute the mask's dimensions in place of the
            ' layer's dimensions; this gives us proper boundaries for decompression overrun checks.
            If (.ciID < cidAlpha) Then
                
                If (.ciID = cidUserLayerMask) Then
                    With m_LayerMaskData.lmRect
                        layerWidth = (.Right - .Left)
                        layerHeight = (.Bottom - .Top)
                    End With
                ElseIf (.ciID = cidRealUserLayerMask) Then
                    With m_RealLayerMaskData.lmRect
                        layerWidth = (.Right - .Left)
                        layerHeight = (.Bottom - .Top)
                    End With
                Else
                    .ciDecodedOK = False
                    GoTo SkipChannel
                End If
                
                'Perform a final check to ensure calculated boundaries are acceptable
                If (layerWidth <= 0) Or (layerHeight <= 0) Then
                    InternalError "DecodeChannels", "layer width or height is <= 0; layer will not be decoded"
                    .ciDecodedOK = False
                    GoTo SkipChannel
                End If
                
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
                If PSD_DEBUG_VERBOSE Or Me.IsLayerGroupEnd Or Me.IsLayerGroupStart Then
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
                        Dim bytesReadDst As Long, bytesReadSrc As Long, bytesReadLine As Long, rleCode As Long
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
                                    
                                    'Failsafe RLE checks can be performed here, but they have an obvious negative perf impact;
                                    ' we currently assume the PSD is well-formed.
                                    'If (bytesReadLine > lineLengths(y)) Then Exit Do
                                    'If (bytesReadDst + rleCode > UBound(.ciDataDecoded)) Then Exit Do
                                    
                                    CopyMemoryStrict VarPtr(.ciDataDecoded(bytesReadDst)), VarPtr(.ciData(basePtrPosition + bytesReadSrc)), rleCode
                                    bytesReadDst = bytesReadDst + rleCode
                                    bytesReadSrc = bytesReadSrc + rleCode
                                    
                                'If the code is on the range [-127,-1], inclusive, the next byte should be
                                ' repeated [1 - code] times in the destination buffer.
                                ElseIf (rleCode > -128) Then
                                    
                                    rleCode = 1 - rleCode
                                    bytesReadLine = bytesReadLine + 1
                                    
                                    'See above statement on the performance impact of "safe" boundary checking
                                    'If (bytesReadLine > lineLengths(y)) Then Exit Do
                                    'If (bytesReadDst + rleCode > UBound(.ciDataDecoded)) Then Exit Do
                                    
                                    FillMemory VarPtr(.ciDataDecoded(bytesReadDst)), rleCode, .ciData(basePtrPosition + bytesReadSrc) 'tmpVal
                                    bytesReadDst = bytesReadDst + rleCode
                                    bytesReadSrc = bytesReadSrc + 1
                                    
                                '-128 means do nothing (nop).  I'm not sure why an encoder would set this, but if it does,
                                ' we can just continue normally.
                                'Else
                                    
                                End If
                            
                            'Technically, a malformed PSD could cause issues here, as a mis-reported RLE length
                            ' toward the end of the file could cause us to overrun the destination buffer.
                            ' Checks could be added both here, and on the above RLE checks (as both of those
                            ' could also overrun a destination line) for additional safety, at some cost to performance.
                            Loop While bytesReadLine < lineLengths(y)
                            
                        Next y
                        
                        'Decompression was successful!
                        .ciDecodedOK = True
                        DecodeChannels = psd_Success
                    
                    'Zip compression has not actually been tested "in the wild", because I don't have a usable
                    ' test file with this encoding.  That said, I see no reason why the current code wouldn't
                    ' work as-is (famous last words)...
                    Case ccZip
                        
                        'Inflate the stream directly into its final buffer.
                        If Compression.IsFormatSupported(cf_Zlib) Then
                            ReDim .ciDataDecoded(0 To finalArraySize - 1) As Byte
                            If Compression.DecompressPtrToPtr(VarPtr(.ciDataDecoded(0)), finalArraySize, VarPtr(.ciData(0)), .ciDataLength, cf_Zlib) Then
                                .ciDecodedOK = True
                                DecodeChannels = psd_Success
                            Else
                                DecodeChannels = psd_Warning
                                InternalError "DecodeChannels", "libdeflate failed to inflate compressed channel stream."
                            End If
                        Else
                            DecodeChannels = psd_Failure
                            InternalError "DecodeChannels", "libdeflate missing"
                        End If
                        
                    Case ccZipWithPrediction
                        
                        Dim startTime As Currency
                        VBHacks.GetHighResTime startTime
                        
                        'Zip w/prediction is only available on 16-bit and 32-bit channels.
                        If (bitsPC = 16) Or (bitsPC = 32) Then
                        
                            'Start by inflating the stream.
                            If Compression.IsFormatSupported(cf_Zlib) Then
                                
                                ReDim .ciDataDecoded(0 To finalArraySize - 1) As Byte
                                If Compression.DecompressPtrToPtr(VarPtr(.ciDataDecoded(0)), finalArraySize, VarPtr(.ciData(0)), .ciDataLength, cf_Zlib) Then
                                    
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
                                    
                                    Dim deltaPrev As Long, newValue As Long, tmpIndex As Long, basePtr As Long
                                    Dim srcIndex As Long, srcInt As Integer, srcLong As Long, yOffset As Long, xOffset As Long
                                    Dim lyrBoundY As Long, lyrBoundX As Long
                                    lyrBoundY = layerHeight - 1
                                    lyrBoundX = layerWidth - 1
                                    basePtr = VarPtr(.ciDataDecoded(0))
                                    
                                    '16-bit PSDs use an extremely primitive delta encoding scheme - each line starts with a
                                    ' 2-byte unsigned int value, and subsequent bytes are the delta between the previous
                                    ' pixel and the current one.  This makes decoding tedious, but straightforward.
                                    ' (Note that I deliberately mix data types here, as VB lacks unsigned ints but we
                                    ' can "cheat" and use longs by shifting memory strategically.)
                                    '
                                    'TODO: wrap an integer array around the decoded data stream; we could use this to
                                    ' "trick" VB into reading/placing signed ints
                                    If (bitsPC = 16) Then
                                        
                                        srcIndex = 0
                                        
                                        For y = 0 To lyrBoundY
    
                                            'Retrieve the first value in each row and treat it as an absolute value
                                            GetMem2_Ptr basePtr + srcIndex, VarPtr(deltaPrev)
                                            srcIndex = srcIndex + 2
                                            
                                            'While here, cache the current scanline pointer to avoid unnecessary multipliations
                                            ' inside the inner loop.
                                            tmpIndex = y * layerWidth * 2
                                            
                                            'All subsequent bytes are treated as deltas
                                            For x = 1 To lyrBoundX
                                                
                                                'Retrieve the next two bytes into a signed int
                                                GetMem2_Ptr basePtr + srcIndex, VarPtr(srcInt)
                                                
                                                'Apply the delta from the previous pixel
                                                newValue = CLng(srcInt) + deltaPrev
                                                
                                                'Store the result in-place over the old value.  (A lack of easy dereferencing
                                                ' makes this tedious, but GetMem2 provides a slight speed boost over RtlMoveMemory.)
                                                GetMem2_Ptr VarPtr(newValue), basePtr + srcIndex
                                                
                                                'Store the delta for the next pixel to use and increment our index into the array
                                                deltaPrev = newValue And &HFFFF&
                                                srcIndex = srcIndex + 2
                                                
                                            Next x
    
                                        Next y
                                        
                                    '32-bit data is a totally different beast.
                                    ElseIf (bitsPC = 32) Then
                                        
                                        srcIndex = 0
                                        
                                        'We need a temporary buffer the size of one scanline; we'll reassemble individual
                                        ' floats into this line as we go, then copy the finished result back into the
                                        ' original array.
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
                                    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "libdeflate failed to inflate stream."
                                    .ciDecodedOK = False
                                    DecodeChannels = psd_Failure
                                End If
                            Else
                                If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "libdeflate is missing; can't deflate stream."
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
    Dim bitFlags(0 To 7) As Byte
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
            
            'The size of this chunk is "size of pixel data + size of compression flag", with the
            ' latter value always being two bytes (an unsigned int).  To avoid confusion, we
            ' pre-calculate the raw "size of pixel data" by subtracting two from the net chunk size.
            .ciDataLength = .ciSizeL - 2
            
            'Want to see basic channel data (helpful for debugging)?
            'Debug.Print "data length: ", .ciDataLength, .ciID
            
            'It *is* possible to have zero-length layer data (for example, layer groups and
            ' adjustment/effect layers can do this).  As such, it's critical to check channel size
            ' before attempting to retrieve raw channel data.
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
    m_LayerNameANSI = "background"      'TODO: figure out what this should *actually* be - do PSDs store a name elsewhere?
    
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
                If imageIsPSB Then
                    rleTable(i) = srcStream.ReadLong_BE()
                Else
                    rleTable(i) = srcStream.ReadIntUnsigned_BE()
                End If
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
                If imageIsPSB Then
                    m_Channels(i).ciDataLength = (imgHeightPx * 4) + chSizeSum
                Else
                    m_Channels(i).ciDataLength = (imgHeightPx * 2) + chSizeSum
                End If
                
            Next i
            
            'Now that we know how large each channel's raw data is, we can copy the data into place.
            For i = 0 To m_ChannelCount - 1
                With m_Channels(i)
                    .ciCompression = cmpMethod
                    'Channels appear in sequential order, per the spec:
                    ' "Image data is stored in planar order: first all the red data, then all the green data, etc."
                    .ciID = i
                    ReDim .ciData(0 To .ciDataLength - 1) As Byte
                    
                    If imageIsPSB Then
                        srcStream.ReadBytesToBarePointer VarPtr(.ciData(0)), (imgHeightPx * 4)
                    Else
                        srcStream.ReadBytesToBarePointer VarPtr(.ciData(0)), (imgHeightPx * 2)
                    End If
                    
                End With
            Next i
            
            'Do one last loop to copy actual RLE bytes into place
            For i = 0 To m_ChannelCount - 1
                With m_Channels(i)
                    If imageIsPSB Then
                        srcStream.ReadBytesToBarePointer VarPtr(.ciData(imgHeightPx * 4)), (.ciDataLength - (imgHeightPx * 4))
                    Else
                        srcStream.ReadBytesToBarePointer VarPtr(.ciData(imgHeightPx * 2)), (.ciDataLength - (imgHeightPx * 2))
                    End If
                End With
            Next i
            
        'TODO!  I have not encountered these segments "in the wild" on single-layer images,
        ' as the single-layer data is typically stored in a separate Lr16/Lr32 segment.
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
    If (m_ChannelCount > 0) Then
    
        ReDim m_Channels(0 To m_ChannelCount - 1) As PSD_ChannelInfo
        
        'Each channel has its own channel ID and length indicator
        Dim i As Long
        For i = 0 To m_ChannelCount - 1
        
            m_Channels(i).ciID = srcStream.ReadInt_BE()
            m_Channels(i).ciSizeL = srcStream.ReadLong_BE()
            
            'PSB entries are actually 8-bytes; grab the next byte while we're here, but note that 32-bit limitations
            ' mean we can still only reasonably use the 4 low bytes of the size.
            If imageIsPSB Then
                m_Channels(i).ciSizeH = m_Channels(i).ciSizeL
                m_Channels(i).ciSizeL = srcStream.ReadLong_BE
            End If
                
            'Random fact: layer group start/end markers will always list 4 RGBA channels,
            ' with a constant size of 2 bytes per channel (whether PSD or PSB format)
            
        Next i
        
    End If
    
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
    m_BlendMode = GetBlendModeFromKey(bmCode)
    
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
    
    'Some layer states (e.g. belonging to a group) are optional; set any of these optional parameters
    ' to default values in advance.  We'll overwrite their values if we encounter the relevant
    ' optional chunk(s).
    m_GroupState = lgs_None
    
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
        
        'Layer group data is tricky; PD doesn't support groups (at present), but we still want to
        ' properly handle things like group visibility when importing PSD documents.  Plugins in
        ' other software that isn't group-aware (e.g. PDN) list group start/stop layers as
        ' blank layers... maybe there's wisdom in this?
        '
        'Anyway, we parse group state if it exists, but we do not construct actual groups from
        ' the data.
        '
        'Note that the PSD spec only lists "lsct" as a valid marker for this data, but files
        ' "in the wild" (perhaps from older PS versions?) have also used "lsdk" with the same
        ' data arrangement.  We cover both cases "just in case".
        If m_LayerInfo.DoesKeyExist("lsct") Or m_LayerInfo.DoesKeyExist("lsdk") Then
            
            'Group data exists; set up a stream around it
            If m_LayerInfo.DoesKeyExist("lsct") Then
                Set tmpStream = m_LayerInfo.GetStreamForKey("lsct")
            Else
                Set tmpStream = m_LayerInfo.GetStreamForKey("lsdk")
            End If
            
            'The first four bytes describe "section divider type"; from the spec:
            ' 4 possible values, 0 = any other type of layer, 1 = open "folder",
            ' 2 = closed "folder", 3 = bounding section divider, hidden in the UI
            m_GroupState = tmpStream.ReadLong_BE()
            If ((m_GroupState < 0) Or (m_GroupState > 3)) Then m_GroupState = lgs_None
            
            'If the chunk is 12-bytes or longer, then it encodes the following:
            '4-byte signature: '8BIM'
            '4-byte "key": see blend mode keys in <Layer records>.
            If (tmpStream.GetStreamSize >= 12) Then
                If (tmpStream.ReadString_ASCII(4) = "8BIM") Then
                    'Next four bytes contain group blend mode.  We don't currently retrieve this data,
                    ' but my limited testing has shown that this is always listed as "pass" for
                    ' "passthrough mode" in the wild.
                End If
            End If
            
            'Important to free stream, as it unsafely wraps the underlying chunk data
            tmpStream.StopStream
            
        End If
        
        Set tmpStream = Nothing
    
    End If
    
'    Debug.Print "Channel count", m_LayerNameANSI, m_ChannelCount
'    If (m_LayerNameANSI = "Group 1") Or (m_LayerNameANSI = "</Layer group>") Then
'        For i = 0 To m_ChannelCount - 1
'            With m_Channels(i)
'                Debug.Print .ciID, .ciSizeL
'            End With
'        Next i
'        Debug.Print m_Rect.Left, m_Rect.Top, m_Rect.Right, m_Rect.Bottom
'    End If
    
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

Private Function GetBlendModeFromKey(ByVal srcKey As String) As PD_BlendMode
    
    Select Case srcKey
        Case "pass" ' = pass through
            GetBlendModeFromKey = BM_Normal 'TODO
        Case "norm" ' = normal
            GetBlendModeFromKey = BM_Normal
        Case "diss" ' = dissolve
            GetBlendModeFromKey = BM_Normal 'TODO
        Case "dark" ' = darken
            GetBlendModeFromKey = BM_Darken
        Case "mul " ' = multiply
            GetBlendModeFromKey = BM_Multiply
        Case "idiv" ' = color burn
            GetBlendModeFromKey = BM_ColorBurn
        Case "lbrn" ' = linear burn
            GetBlendModeFromKey = BM_LinearBurn
        Case "dkCl" ' = darker color
            GetBlendModeFromKey = BM_Darken 'TODO
        Case "lite" ' = lighten
            GetBlendModeFromKey = BM_Lighten
        Case "scrn" ' = screen
            GetBlendModeFromKey = BM_Screen
        Case "div " ' = color dodge
            GetBlendModeFromKey = BM_ColorDodge
        Case "lddg" ' = linear dodge
            GetBlendModeFromKey = BM_LinearDodge
        Case "lgCl" ' = lighter color
            GetBlendModeFromKey = BM_Lighten    'TODO
        Case "over" ' = overlay
            GetBlendModeFromKey = BM_Overlay
        Case "sLit" ' = soft light
            GetBlendModeFromKey = BM_SoftLight
        Case "hLit" ' = hard light
            GetBlendModeFromKey = BM_HardLight
        Case "vLit" ' = vivid light
            GetBlendModeFromKey = BM_VividLight
        Case "lLit" ' = linear light
            GetBlendModeFromKey = BM_LinearLight
        Case "pLit" ' = pin light
            GetBlendModeFromKey = BM_PinLight
        Case "hMix" ' = hard mix
            GetBlendModeFromKey = BM_HardMix
        Case "diff" ' = difference
            GetBlendModeFromKey = BM_Difference
        Case "smud" ' = exclusion
            GetBlendModeFromKey = BM_Exclusion
        Case "fsub" ' = subtract
            GetBlendModeFromKey = BM_Subtract
        Case "fdiv" ' = divide
            GetBlendModeFromKey = BM_Divide
        Case "hue " ' = hue
            GetBlendModeFromKey = BM_Hue
        Case "sat " ' = saturation
            GetBlendModeFromKey = BM_Saturation
        Case "colr" ' = color
            GetBlendModeFromKey = BM_Color
        Case "lum " ' = luminosity
            GetBlendModeFromKey = BM_Luminosity
        Case Else
            InternalError "GetBlendModeFromKey", "ParseLayer encountered an unknown blend mode code: " & srcKey
            GetBlendModeFromKey = BM_Normal
            
    End Select
    
End Function

Private Function GetKeyFromBlendMode(ByVal srcBlendMode As PD_BlendMode) As String
    
    Select Case srcBlendMode
        
        Case BM_Normal ' = normal
            GetKeyFromBlendMode = "norm"
        Case BM_Darken ' = darken
            GetKeyFromBlendMode = "dark"
        Case BM_Multiply ' = multiply
            GetKeyFromBlendMode = "mul "
        Case BM_ColorBurn ' = color burn
            GetKeyFromBlendMode = "idiv"
        Case BM_LinearBurn ' = linear burn
            GetKeyFromBlendMode = "lbrn"
        Case BM_Lighten ' = lighten
            GetKeyFromBlendMode = "lite"
        Case BM_Screen ' = screen
            GetKeyFromBlendMode = "scrn"
        Case BM_ColorDodge ' = color dodge
            GetKeyFromBlendMode = "div "
        Case BM_LinearDodge ' = linear dodge
            GetKeyFromBlendMode = "lddg"
        Case BM_Overlay ' = overlay
            GetKeyFromBlendMode = "over"
        Case BM_SoftLight ' = soft light
            GetKeyFromBlendMode = "sLit"
        Case BM_HardLight ' = hard light
            GetKeyFromBlendMode = "hLit"
        Case BM_VividLight ' = vivid light
            GetKeyFromBlendMode = "vLit"
        Case BM_LinearLight ' = linear light
            GetKeyFromBlendMode = "lLit"
        Case BM_PinLight ' = pin light
            GetKeyFromBlendMode = "pLit"
        Case BM_HardMix ' = hard mix
            GetKeyFromBlendMode = "hMix"
        Case BM_Difference ' = difference
            GetKeyFromBlendMode = "diff"
        Case BM_Exclusion ' = exclusion
            GetKeyFromBlendMode = "smud"
        Case BM_Subtract ' = subtract
            GetKeyFromBlendMode = "fsub"
        Case BM_Divide ' = divide
            GetKeyFromBlendMode = "fdiv"
        Case BM_Hue ' = hue
            GetKeyFromBlendMode = "hue "
        Case BM_Saturation ' = saturation
            GetKeyFromBlendMode = "sat "
        Case BM_Color ' = color
            GetKeyFromBlendMode = "colr"
        Case BM_Luminosity ' = luminosity
            GetKeyFromBlendMode = "lum "
        
        'Some Adobe blend modes are not currently implemented in PD; these are still TODO
'        Case "pass" ' = pass through
'        Case "diss" ' = dissolve
'        Case "lgCl" ' = lighter color
'        Case "dkCl" ' = darker color
            
        Case Else
            InternalError "GetKeyFromBlendMode", "Someone passed a bad blend mode ID: " & srcBlendMode
            GetKeyFromBlendMode = "norm"
            
    End Select
    
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
    With m_LayerMaskData.lmRect
        .Top = srcStream.ReadLong_BE()
        .Left = srcStream.ReadLong_BE()
        .Bottom = srcStream.ReadLong_BE()
        .Right = srcStream.ReadLong_BE()
    End With
    
    'Remaining values may not be used in current build; where relevant, we just read and dump 'em
    ' to keep the parser properly aligned.
    
    'Default color (should only be 0 or 255; not current validated)
    m_LayerMaskData.lmDefaultColor = srcStream.ReadByte()
    
    'Flags
    ' bit 0 = position relative to layer
    ' bit 1 = layer mask disabled
    ' bit 2 = invert layer mask when blending (Obsolete)
    ' bit 3 = indicates that the user mask actually came from rendering other data
    ' bit 4 = indicates that the user and/or vector masks have parameters applied to them
    m_LayerMaskData.lmFlags = srcStream.ReadByte()
    
    m_LayerMaskExists = True
    
    'Per the spec, "[2 bytes] Padding. Only present if size = 20"
    If (segmentLength = 20) Then
        srcStream.SetPosition 2, FILE_CURRENT
        
    'When size > 20, padding may instead be replaced by usable data
    Else
        
        If (srcStream.GetPosition() <= expectedFinalPointer - 18) Then
            
            'Real Flags. Same as Flags information above.
            m_RealLayerMaskData.lmFlags = srcStream.ReadByte()
            
            'Real user mask background.  Must be 0 or 255 (not currently validated).
            m_RealLayerMaskData.lmDefaultColor = srcStream.ReadByte()
            
            'Rectangle enclosing the real user mask
            With m_RealLayerMaskData.lmRect
                .Top = srcStream.ReadLong_BE()
                .Left = srcStream.ReadLong_BE()
                .Bottom = srcStream.ReadLong_BE()
                .Right = srcStream.ReadLong_BE()
            End With
            
            'Mark the "real" layer mask as existing
            m_RealLayerMaskExists = True
            
        End If
        
    End If
    
    'Now the spec has this wrong, but after the real layer mask data, there can potentially be more information.
    ' This information is contingent on the m_LayerMaskData.lmFlags byte, above:
    ' "1 byte | Mask Parameters. Only present if bit 4 of Flags set above."
    If ((m_LayerMaskData.lmFlags And &H10&) <> 0) And (srcStream.GetPosition < expectedFinalPointer) Then

        With m_LayerMaskData
            .lmParameters = srcStream.ReadByte()
            .lmParametersLoaded = True
        
            'The parameters byte determines the presence of even MORE data.  (Whoever thought up this
            ' segment of the file should be smacked; saving a few bytes is not worth this much trouble.)
            If ((.lmParameters And &H1&) <> 0) Then .lmUserMaskDensity = srcStream.ReadByte()
            If ((.lmParameters And &H2&) <> 0) Then .lmUserMaskFeather = srcStream.ReadDouble_BE()
            If ((.lmParameters And &H4&) <> 0) Then .lmVectorMaskDensity = srcStream.ReadByte()
            If ((.lmParameters And &H8&) <> 0) Then .lmVectorMaskFeather = srcStream.ReadDouble_BE()
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

Friend Sub WriteLayerData(ByRef cStream As pdStream, ByRef srcLayer As pdLayer, ByRef channelCompression As PSD_ChannelCompression, Optional ByVal writePSB As Boolean = False, Optional ByVal layerGroupState As PSD_LayerGroupState = lgs_None)

    'Layer records are complicated beasts.  Some of the information we place is fixed; other pieces are not.
    ' I don't have a lot of helpful information other than "review the spec" if you have questions.
    
    'Note that group start/end markers are written as semi-normal layer data (presumably for
    ' backward compatibility with old PS versions).
    Dim isGroupMarker As Boolean
    isGroupMarker = (layerGroupState <> lgs_None)
    
    'The first chunk of a layer record is a (mostly) fixed-length segment describing basic layer parameters.
    With srcLayer
        
        'Because PhotoDemon supports non-destructive affine layer transformations (rotate, resize, etc),
        ' and we don't have a way to embed this data inside PSDs (that I know of), we need to first generate
        ' a temporary copy of the source layer's data, with all affine transforms accounted for.  (This is
        ' also useful as we will generate planar channel data in this step, and we want to un-premultiply a
        ' temporary copy of the image, to avoid modifying source image data as part of the save.)
        Dim tmpLayerDIB As pdDIB, layerOffsetX As Long, layerOffsetY As Long
        Set tmpLayerDIB = New pdDIB
        If srcLayer.AffineTransformsActive(True) Then
            If (Not srcLayer.GetAffineTransformedDIB(tmpLayerDIB, layerOffsetX, layerOffsetY)) Then
                InternalError "WriteLayerData", "couldn't generate affine-transformed layer DIB!"
                Exit Sub
            End If
        Else
            tmpLayerDIB.CreateFromExistingDIB srcLayer.GetLayerDIB
            layerOffsetX = srcLayer.GetLayerOffsetX()
            layerOffsetY = srcLayer.GetLayerOffsetY()
        End If
        
        'Before continuing, we want to determine if alpha is relevant in this layer.  If a layer is fully opaque,
        ' we can ignore alpha data, as Photoshop will automatically plug-in full opacity for us.
        m_LayerHasAlpha = DIBs.IsDIBTransparent(tmpLayerDIB)
        
        'Unpremultiply the temporary DIB before continuing, if alpha values are relevant
        If m_LayerHasAlpha Then tmpLayerDIB.SetAlphaPremultiplication False
        
        'UPDATE April 2024
        '
        'PhotoDemon has traditionally written out 24-bit RGB data if a layer doesn't use meaningful alpha,
        ' but this produces unexpected results on some versions of Photoshop.  Unfortunately, I haven't been
        ' able to track down a consistent pattern in why this happens - all other software can open PD's PSDs
        ' without trouble when this occurs, but PS shows some kind of file offset problem, with weird layer
        ' settings and unreliable handling of layer metadata.
        '
        'GIMP avoids this problem by *always* writing an alpha channel, regardless of relevancy,
        ' and indeed this solves the same problem when exporting PSD files from PD.
        '
        'Until I can figure out why only some layered images struggle with 24-bit data, I've switched to
        ' *always* writing RGBA channels for each layer.  (Note that we deliberately set this forcible alpha
        ' value here, *after* unpremultiplying alpha, because only opaque layers will mark m_LayerHasAlpha
        ' as FALSE and their alpha channels consist of only 255, so unpremultiplying is irrelevant.)
        m_LayerHasAlpha = True
        
        'First is layer rect.  Note that Adobe uses the non-standard top/left/bottom/right order.
        ' (Note also that layer groups are always written as 0-size rects.)
        If isGroupMarker Then
            cStream.WritePadding 16
        Else
            cStream.WriteLong_BE layerOffsetY
            cStream.WriteLong_BE layerOffsetX
            cStream.WriteLong_BE (layerOffsetY + tmpLayerDIB.GetDIBHeight())
            cStream.WriteLong_BE (layerOffsetX + tmpLayerDIB.GetDIBWidth())
        End If
        
        'Next is number of channels in the layer; PD always writes RGB/A data.
        ' (Note that this is still true for layer group markers; Adobe always writes them as RGBA data.)
        
        'Group channel count depends on the presence of a mask
        If isGroupMarker Then
            
            'Extra channel required for mask
            If srcLayer.GetLayerMaskExists Then
                cStream.WriteInt_BE 5
                
            'Non-masked groups write empty RGB/A channels
            Else
                cStream.WriteInt_BE 4
            End If
            
        Else
            
            'PD will write RGB layers as appropriate
            If m_LayerHasAlpha Then
                cStream.WriteInt_BE 4
            Else
                cStream.WriteInt_BE 3
            End If
            
        End If
            
        'Next is a list of channel information, 6 bytes per channel (so 6*4 for typical PhotoDemon RGBA layer)
        ' - 2 bytes: Channel ID
        ' - 4 bytes: Channel length (PSB is 8 bytes)
        Dim rleSize As Long
        
        'Channel lengths are POST-compression, so we can't know these values without performing actual
        ' channel compression.  We may as well handle that now, so we don't have to revisit these numbers later.
        ' (This could be better optimized to ignore the alpha channel if we aren't going to write it to file - TODO!)
        ReDim m_Channels(0 To 3) As PSD_ChannelInfo
        If isGroupMarker Then
        
            'Write four dummy RGBA channel ID values, no more or less.  (This matches Photoshop's behavior.)
            cStream.WriteInt_BE -1
            cStream.WriteLong_BE 2
            
            cStream.WriteInt_BE 0
            cStream.WriteLong_BE 2
            
            cStream.WriteInt_BE 1
            cStream.WriteLong_BE 2
            
            cStream.WriteInt_BE 2
            cStream.WriteLong_BE 2
            
            'If we have a mask, write it last
            If srcLayer.GetLayerMaskExists Then
                
                PrepBareMaskForWrite srcLayer, channelCompression
                cStream.WriteInt_BE cidUserLayerMask    'Fixed layer mask ID
                
                'See notes on PackBits behavior, below
                If (channelCompression = ccPackBits) Then
                    rleSize = m_Channels(0).ciDataLength + 2
                    If writePSB Then rleSize = rleSize + (UBound(m_Channels(0).rleTable) + 1) * 4 Else rleSize = rleSize + (UBound(m_Channels(0).rleTable) + 1) * 2
                    cStream.WriteLong_BE rleSize
                Else
                    cStream.WriteLong_BE m_Channels(0).ciDataLength + 2
                End If
                
            End If
        
        Else
            
            PrepChannelsForWrite tmpLayerDIB, channelCompression, (Not m_LayerHasAlpha), False, Nothing
            
            'We can now write channel information for each channel
            Dim i As Long
            For i = 0 To 3
                
                'If the layer's alpha channel is irrelevant, don't write it.
                If (i = 3) And (Not m_LayerHasAlpha) Then Exit For
                
                'Preface this segment with each channel's ID
                cStream.WriteInt_BE m_Channels(i).ciID
                
                'Note that regardless of compression type, we must add "2" to the length of the compressed channel data,
                ' because PSDs require that channel data be prefaced with a 2-byte "compression type" indicator.
                If (channelCompression = ccPackBits) Then
                    
                    'PackBits requires us to write a separate scanline length table at the start of the channel segment
                    rleSize = m_Channels(i).ciDataLength + 2
                    If writePSB Then rleSize = rleSize + (UBound(m_Channels(i).rleTable) + 1) * 4 Else rleSize = rleSize + (UBound(m_Channels(i).rleTable) + 1) * 2
                    cStream.WriteLong_BE rleSize
                    
                Else
                    cStream.WriteLong_BE m_Channels(i).ciDataLength + 2
                End If
                
            Next i
            
        End If
        
        'Blend mode signature is hard-coded
        cStream.WriteString_ASCII "8BIM"
        
        'Blend mode itself is also a hard-coded ANSI string
        cStream.WriteString_ASCII GetKeyFromBlendMode(.GetLayerBlendMode)
        
        'Opacity [0,255]
        cStream.WriteByte Int(.GetLayerOpacity() * 2.55! + 0.5!)
        
        'Clipping
        cStream.WriteByte 0
        
        'Flags; at present, the only flag we care about is the "invisibility" flag (bit 2)
        Dim psdFlags As Byte
        If (Not .GetLayerVisibility) Then psdFlags = psdFlags Or 2
        cStream.WriteByte psdFlags
        
        'Padding byte (required by the spec)
        cStream.WriteByte 0
        
        'Length of the next five fields.  We don't know this value in advance, so simply flag the current position for now
        Dim startExtraFieldsPosition As Long
        startExtraFieldsPosition = cStream.GetPosition()
        cStream.WriteLong_BE 0
        
        'Layer mask data comes next.
        ' PhotoDemon doesn't support masks on normal layers but it *does* support group masks,
        ' so branch accordingly
        If (isGroupMarker And .GetLayerMaskExists) Then
            
            'Note the start of the group section
            Dim startExtraGroupField As Long
            startExtraGroupField = cStream.GetPosition
            
            'Per the spec, this chunk "...can be 40 bytes, 24 bytes, or 4 bytes if no layer mask."
            ' We always write the shorter 24-byte version, so we can simply mark that size in advance.
            cStream.WriteLong_BE 20&
            
            'The next 16-bytes are the mask rect, in "top, left, bottom, right" order
            cStream.WriteLong_BE srcLayer.GetLayerMask.GetMaskTop
            cStream.WriteLong_BE srcLayer.GetLayerMask.GetMaskLeft
            cStream.WriteLong_BE srcLayer.GetLayerMask.GetMaskTop + srcLayer.GetLayerMask.GetMaskHeight
            cStream.WriteLong_BE srcLayer.GetLayerMask.GetMaskLeft + srcLayer.GetLayerMask.GetMaskWidth
            
            'Default color (translate from PD's [0.0, 100.0] scale)
            cStream.WriteByte Int(srcLayer.GetLayerMask.GetOpacityOutsideMask * 2.55! + 0.5!)
            
            'Mask flags (always 0 when coming from PD)
            cStream.WriteByte 0
            
            'Two bytes padding
            cStream.WritePadding 2
            
        'No mask
        Else
            cStream.WriteLong_BE 0
        End If
        
        'Layer blending ranges are also not currently supported (or understood, if I'm being honest)
        cStream.WriteLong_BE 0
        
        'Layer name is a Pascal-style string, "padded to a multiple of 4 bytes".  By my testing, this name is not used
        ' by any modern PSD-compatible program (including Adobe itself).  Instead, an optional appended block allows
        ' us to provide a variable-length Unicode layer name.
        '
        'We could use WideCharToMultiByte() to produce a string in the local codepage, but given the general
        ' uselessness of this string in the first place, let's just write a UTF-8 string (which plays nicely
        ' with some 3rd-party readers written in e.g. Python) and assume that very few programs, if any,
        ' will use this data in a way that makes UTF-8 problematic.
        Dim lName As String, lNameLengthBytes As Long, tmpBytes() As Byte
        lName = .GetLayerName()
        
        'Layer group start/end markers need additional parsing
        If isGroupMarker Then
            
            Dim groupStartText As String, gstPosition As Long
            groupStartText = g_Language.TranslateMessage("Group start:")
            
            If (layerGroupState = lgs_GroupEnd) Then
                lName = "</Layer Group>"
            Else
                gstPosition = InStr(1, lName, groupStartText, vbBinaryCompare)
                If (gstPosition > 0) Then lName = Trim$(Right$(lName, Len(lName) - (gstPosition + Len(groupStartText))))
            End If
        End If
        
        If (LenB(lName) > 0) Then
        
            Strings.UTF8FromString lName, tmpBytes, lNameLengthBytes
            If (lNameLengthBytes > 255) Then lNameLengthBytes = 255
            cStream.WriteByte lNameLengthBytes
            cStream.WriteByteArray tmpBytes, lNameLengthBytes
            
            'Calculate and write padding, as necessary
            Dim reqPadding As Long
            reqPadding = ((lNameLengthBytes + 1) Mod 4)
            If (reqPadding <> 0) Then cStream.WritePadding 4 - reqPadding
        
        Else
            cStream.WriteLong_BE 0 '1 byte = 0 as length, 3 null bytes for padding
        End If
        
        'All mandatory layer information has now been written.  We can now write any optional
        ' additional descriptors.  These blocks all use the same general layout.
        
        'One optional descriptor we always write is a Unicode layer name.
        If (LenB(lName) > 0) Then
        
            'Write a 4-byte hard-coded "start of block" flag
            cStream.WriteString_ASCII "8BIM"
            
            'Write a 4-byte hard-coded ASCII identifier of the optional block
            cStream.WriteString_ASCII "luni"
            
            'Write the length of the *entire* data section, in bytes
            ' (4-byte string length + string itself + 2-byte null terminator)
            cStream.WriteLong_BE 4 + LenB(lName) + 2
            
            'Write the string itself (which is prefixed by the length IN CODE POINTS, ignoring the terminating null)
            cStream.WriteLong_BE Len(lName)
            cStream.WriteString_UnicodeBE lName, True
            
            'Optional blocks are supposed to be even-padded, but we don't have to worry about this with
            ' Unicode strings (2 bytes per char, remember!)
            
        End If
        
        'Layer group data is stored as an extra chunk as well
        If isGroupMarker Then
            
            'Write a 4-byte hard-coded "start of block" flag
            cStream.WriteString_ASCII "8BIM"
            
            'Write a 4-byte hard-coded ASCII identifier of the optional block
            cStream.WriteString_ASCII "lsct"
            
            'Write chunk length (this is a fixed magic number at present), 4-bytes for end group markers
            ' 4 or 12 or 16 for start group markers (we use 16 in PD).
            If (layerGroupState = lgs_GroupEnd) Then
                cStream.WriteLong_BE 4
            Else
                cStream.WriteLong_BE 16
            End If
            
            'Start and end markers have different magic number values
            cStream.WriteLong_BE layerGroupState
            If (layerGroupState = lgs_GroupStartOpen) Or (layerGroupState = lgs_GroupStartClosed) Then
                
                'Next, write the standard PSD blend-mode marker, followed by "passthrough" mode
                ' (this is always used "in the wild" in my limited testing of PSDs with group data).
                cStream.WriteString_ASCII ("8BIM")
                cStream.WriteString_ASCII ("pass")
                
                'Finally, a "section divider sub-type" where "0 = normal"
                cStream.WriteLong_BE 0
                
            End If
            
        End If
        
        'END OF LAYER BLOCK
        
        'Before exiting, we need to go back and overwrite the "length of extra data field" value from
        ' earlier in this function.
        Dim finalPosition As Long
        finalPosition = cStream.GetPosition()
        
        cStream.SetPosition startExtraFieldsPosition, FILE_BEGIN
        cStream.WriteLong_BE (finalPosition - startExtraFieldsPosition) - 4  'Do not include the length value's 4 bytes in total
        cStream.SetPosition finalPosition, FILE_BEGIN
        
    End With

End Sub

'This function is *ONLY SUPPORTED FOR LAYER GROUP MARKERS* at present; regular layer masks remain TODO.
' (As such, this function will likely need modification since it assumes no other meaningful channels
' for the source layer, by design.)
Friend Sub PrepBareMaskForWrite(ByRef srcLayer As pdLayer, ByVal channelCompression As PSD_ChannelCompression)
    
    'Failsafes only; the caller should check these
    If (srcLayer Is Nothing) Then Exit Sub
    If (Not srcLayer.GetLayerMaskExists) Then Exit Sub
    If (srcLayer.GetLayerMask Is Nothing) Then Exit Sub
    If (srcLayer.GetLayerMask.GetMaskWidth = 0) Or (srcLayer.GetLayerMask.GetMaskHeight = 0) Then Exit Sub
    If (srcLayer.GetLayerMask.GetPtrToMaskBytes = 0) Then Exit Sub
    
    'TODO: add check for mask active (that flag won't be relevant until masks are *fully* implemented in PD)
    
    'Because we don't have a reference DIB, this function can be much shorter than a full RGB/A write -
    ' but note that we still need things like scanline markers, even for a mask, for PackBits compression
    ' (since its RLE markers are scanline-based).
    
    'Ensure a channel object exists (since we aren't writing color channels, *this* function must
    ' initialize the channel collection)
    ReDim m_Channels(0) As PSD_ChannelInfo
    m_Channels(0).ciID = cidUserLayerMask
    
    Dim maskWidth As Long, maskHeight As Long
    maskWidth = srcLayer.GetLayerMask.GetMaskWidth
    maskHeight = srcLayer.GetLayerMask.GetMaskHeight
    
    'Grab pointers to the mask data
    Dim tmpBytes() As Byte, tmpSA As SafeArray1D
    Dim scanStart As Long, scanWidth As Long
    scanStart = srcLayer.GetLayerMask.GetPtrToMaskBytes()
    scanWidth = maskWidth
    
    VBHacks.WrapArrayAroundPtr_Byte tmpBytes, tmpSA, scanStart, maskWidth * maskHeight
    
    Dim y As Long
    Dim dstOffset As Long
    
    Dim xLoopSrc As Long, yLoopSrc As Long
    xLoopSrc = maskWidth - 1
    yLoopSrc = maskHeight - 1
    
    Select Case channelCompression
    
        'RAW compression just means "no compression"; write mask as-is
        Case ccRaw
            ReDim m_Channels(0).ciData(0 To maskWidth * maskHeight - 1) As Byte
            m_Channels(0).ciDataLength = maskWidth * maskHeight
            VBHacks.CopyMemoryStrict VarPtr(m_Channels(0).ciData(0)), scanStart, maskWidth * maskHeight
            
        'PackBits is simple RLE compression
        Case ccPackBits
        
            'See the PrepChannelsForWrite function for full details on PackBits; this code is adopted from there
            With m_Channels(0)
                .ciDataLength = maskWidth * maskHeight + Int(CDbl(maskWidth * maskHeight) * 0.1 + 0.5)
                ReDim .ciData(0 To .ciDataLength) As Byte
                ReDim .rleTable(0 To maskHeight - 1) As Long
            End With
            
            'To simplify the PackBits algorithm, we'll manually pull RGBA bytes into their own planar arrays
            ' before packing them.  We'll also use a temporary destination array that is plenty long enough
            ' to hold any pathologically bad lines.
            Dim tmpSrcBytes() As Byte, tmpDstBytes() As Byte, dstSize As Long
            ReDim tmpSrcBytes(0 To maskWidth - 1) As Byte
            ReDim tmpDstBytes(0 To maskWidth * 2) As Byte
            
            'dstOffset points to our current location in the RLE-encoded array
            dstOffset = 0
            
            'Process each scanline as a separate entity
            For y = 0 To yLoopSrc
                
                'Move this mask scanline into the temporary source bytes array
                VBHacks.CopyMemoryStrict VarPtr(tmpSrcBytes(0)), VarPtr(tmpBytes(y * scanWidth)), scanWidth
                
                'Hand the temporary source array off to the packbits encoder; it handles the rest!
                GeneratePackbitsLine maskWidth, dstSize, tmpSrcBytes, tmpDstBytes
                
                'Take the retrieved data and copy into the destination data array, and store the RLE line size
                CopyMemoryStrict VarPtr(m_Channels(0).ciData(dstOffset)), VarPtr(tmpDstBytes(0)), dstSize
                m_Channels(0).rleTable(y) = dstSize
                
                'Update the running offset into the destination array
                dstOffset = dstOffset + dstSize
                
            Next y
            
            'Store the finished length of this PackBits stream and trim the completed array
            ' (to save a bit of memory)
            m_Channels(0).ciDataLength = dstOffset
            ReDim Preserve m_Channels(0).ciData(0 To dstOffset - 1) As Byte
            
        Case ccZip
            
            'Simple - just zip the data into the channel array!
            With m_Channels(0)
                Compression.CompressPtrToDstArray .ciData, .ciDataLength, VarPtr(tmpBytes(0)), maskWidth * maskHeight, cf_Zlib, -1, False, True
            End With
            
        Case ccZipWithPrediction
        'Other compression methods TODO
        
    End Select
    
    VBHacks.UnwrapArrayFromPtr_Byte tmpBytes
    
End Sub

'Note that this function may be called externally; this is by design, as it is shared by our parent pdPSD class
' when writing a composite image copy out to the end of a PSD file.  (Also note that masks are *not* fully
' supported yet - at present, PhotoDemon only supports them for layer group markers, and only if the group markers
' came from a PSD with embedded group masks.)
Friend Sub PrepChannelsForWrite(ByRef srcLayerDIB As pdDIB, ByVal channelCompression As PSD_ChannelCompression, Optional ByVal skipAlpha As Boolean = False, Optional ByVal prepMaskToo As Boolean = False, Optional ByRef srcLayer As pdLayer = Nothing)
    
    'Grab pointers to the layer's image data
    Dim tmpBytes() As Byte, tmpSA As SafeArray1D
    Dim scanStart As Long, scanWidth As Long
    If (Not srcLayerDIB Is Nothing) Then
        srcLayerDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
        scanStart = tmpSA.pvData
        scanWidth = srcLayerDIB.GetDIBStride
    End If
    
    Dim chIndex As Long, x As Long, y As Long
    Dim lWidth As Long, lHeight As Long, xLoopSrc As Long, yLoopSrc As Long
    Dim dstOffset As Long
    
    If (Not srcLayerDIB Is Nothing) Then
        lWidth = srcLayerDIB.GetDIBWidth
        lHeight = srcLayerDIB.GetDIBHeight
    End If
    
    xLoopSrc = lWidth - 1
    yLoopSrc = lHeight - 1
    
    'This function may be called by an outside caller (as part of writing a compressed composite image
    ' at the end of a PSD file), so perform a failsafe check for channel count.
    If (m_ChannelCount <= 0) Then
        If (srcLayerDIB.GetDIBColorDepth = 32) Then m_ChannelCount = 4 Else m_ChannelCount = 3
        If skipAlpha Then m_ChannelCount = 3
        ReDim m_Channels(0 To m_ChannelCount - 1) As PSD_ChannelInfo
    End If
    
    'Mark channel IDs in advance; we'll use a fixed RGBA order for these
    m_Channels(0).ciID = cidRed
    m_Channels(1).ciID = cidGreen
    m_Channels(2).ciID = cidBlue
    If (m_ChannelCount > 3) Then m_Channels(3).ciID = cidAlpha
    
    Select Case channelCompression
    
        'RAW compression just means "no compression"; extract all source channels as-is
        Case ccRaw
            
            For chIndex = 0 To m_ChannelCount - 1
                ReDim m_Channels(chIndex).ciData(0 To lWidth * lHeight - 1) As Byte
                m_Channels(chIndex).ciDataLength = lWidth * lHeight
            Next chIndex
            
            For y = 0 To yLoopSrc
                tmpSA.pvData = scanStart + scanWidth * y
            For x = 0 To xLoopSrc
                
                dstOffset = y * lWidth + x
                
                'Destination channels are in RGBA order, not BGRA order
                m_Channels(2).ciData(dstOffset) = tmpBytes(x * 4)
                m_Channels(1).ciData(dstOffset) = tmpBytes(x * 4 + 1)
                m_Channels(0).ciData(dstOffset) = tmpBytes(x * 4 + 2)
                If (m_ChannelCount = 4) Then m_Channels(3).ciData(dstOffset) = tmpBytes(x * 4 + 3)
                
            Next x
            Next y
        
        'PackBits is simple RLE compression
        Case ccPackBits
        
            'Writing a PackBits stream is fairly straightforward.  There's no "official" guideline to writing
            ' a PackBits stream (and in fact, even "official" implementations like Apple vs Adobe vs libtiff
            ' use slightly different patterns), so we're mostly left to our own devices.
            '
            'For purposes of this encoder, just note that PackBits streams are generated on a per-scanline basis.
            ' (Said another way, each scanline is its own unique stream; the next line "starts fresh" with its
            ' own stream.)  This approach would make this process easy to multithread if desired, but at present,
            ' it's a low-usage scenario so I'm not too obsessed with improving its performance further.
            '
            'Start by initializing separate RLE tables for each channel.  Note also that PackBits does have a
            ' pathological worst-case memory consumption of one extra byte for each 127 input bytes (if the
            ' input data has 0% contiguous runs), so the potential output space *is* larger than the number of
            ' pixels in the channel.
            For chIndex = 0 To m_ChannelCount - 1
                With m_Channels(chIndex)
                    .ciDataLength = lWidth * lHeight + Int(CDbl(lWidth * lHeight) * 0.1)
                    ReDim .ciData(0 To .ciDataLength) As Byte
                    ReDim .rleTable(0 To lHeight - 1) As Long
                End With
            Next chIndex
            
            'To simplify the PackBits algorithm, we'll manually pull RGBA bytes into their own planar arrays
            ' before packing them.  We'll also use a temporary destination array that is plenty long enough
            ' to hold any pathologically bad lines.
            Dim tmpSrcBytes() As Byte, tmpDstBytes() As Byte, srcOffset As Long, dstSize As Long
            ReDim tmpSrcBytes(0 To lWidth - 1) As Byte
            ReDim tmpDstBytes(0 To lWidth * 2) As Byte
            
            'First test: before doing anything else, let's just make sure we can write a literal-only
            ' (no run) PackBits that produces a valid output stream
            For chIndex = 0 To m_ChannelCount - 1
            
                'Source data is in BGRA order; we want it in RGBA order
                If (chIndex = 0) Then
                    srcOffset = 2
                ElseIf (chIndex = 2) Then
                    srcOffset = 0
                Else
                    srcOffset = chIndex
                End If
                
                dstOffset = 0
                
                With m_Channels(chIndex)
                    
                    For y = 0 To yLoopSrc
                        
                        'Point an array at this scanline in the original image
                        tmpSA.pvData = scanStart + scanWidth * y
                    
                        'Start by populating a temporary source bytes array
                        For x = 0 To xLoopSrc
                            tmpSrcBytes(x) = tmpBytes(x * 4 + srcOffset)
                        Next x
                        
                        'Hand the temporary source array off to the packbits encoder; it handles the rest!
                        GeneratePackbitsLine lWidth, dstSize, tmpSrcBytes, tmpDstBytes
                        
                        'Take the retrieved data and copy into the destination data array, and store the RLE line size
                        CopyMemoryStrict VarPtr(.ciData(dstOffset)), VarPtr(tmpDstBytes(0)), dstSize
                        .rleTable(y) = dstSize
                        
                        'Update the running offset into the destination array
                        dstOffset = dstOffset + dstSize
                        
                    Next y
                    
                    'Store the finished length of this PackBits stream and trim the completed array
                    ' (to save a bit of memory)
                    .ciDataLength = dstOffset
                    ReDim Preserve .ciData(0 To dstOffset - 1) As Byte
                
                End With
            
            Next chIndex
        
        Case ccZip
            
            'Start by generating planar copies of each channel
            For chIndex = 0 To m_ChannelCount - 1
                ReDim m_Channels(chIndex).ciDataDecoded(0 To lWidth * lHeight - 1) As Byte
                m_Channels(chIndex).ciDataLength = lWidth * lHeight
            Next chIndex
            
            For y = 0 To yLoopSrc
                tmpSA.pvData = scanStart + scanWidth * y
            For x = 0 To xLoopSrc
                
                dstOffset = y * lWidth + x
                
                'Destination channels are in RGBA order, not BGRA order
                m_Channels(2).ciDataDecoded(dstOffset) = tmpBytes(x * 4)
                m_Channels(1).ciDataDecoded(dstOffset) = tmpBytes(x * 4 + 1)
                m_Channels(0).ciDataDecoded(dstOffset) = tmpBytes(x * 4 + 2)
                If (m_ChannelCount = 4) Then m_Channels(3).ciDataDecoded(dstOffset) = tmpBytes(x * 4 + 3)
                
            Next x
            Next y
            
            'With all channels separated into planar arrays, we can now produce zipped copies of them
            For chIndex = 0 To m_ChannelCount - 1
                With m_Channels(chIndex)
                    Compression.CompressPtrToDstArray .ciData, .ciDataLength, VarPtr(.ciDataDecoded(0)), lWidth * lHeight, cf_Zlib, -1, False, True
                End With
            Next chIndex
            
        Case ccZipWithPrediction
        'Other compression methods TODO
        
    End Select
    
    srcLayerDIB.UnwrapArrayFromDIB tmpBytes
    
End Sub

'Generate a PackBits-compressed line of pixels.  Source and destination arrays must be non-null and properly sized
' (especially the destination line - it must be large enough to hold any possible PackBits output).
Private Sub GeneratePackbitsLine(ByVal scanlineWidth As Long, ByRef dstFinalSize As Long, ByRef srcBytes() As Byte, ByRef dstBytes() As Byte)
    
    'Cover the special case of 1-pixel lines
    If (scanlineWidth = 1) Then
        dstBytes(0) = 0
        dstBytes(1) = srcBytes(0)
        dstFinalSize = 2
        Exit Sub
    End If
    
    'Some upper limits are defined by PackBits itself.
    Const MAXIMUM_RUN_LENGTH As Long = 128
    Const MAXIMUM_LITERAL_LENGTH As Long = 128
    
    Dim dstOffset As Long: dstOffset = 0
    Dim srcOffset As Long: srcOffset = 0
    Dim curRunLength As Long: curRunLength = 0
    Dim targetByte As Byte
    
    Do While (srcOffset < scanlineWidth)
        
        'Previous iterations, if any, will have advanced the pointer to a new byte, one that does not belong
        ' to the previous segment (repeats or literals).  Our job is to analyze this byte and its immediate
        ' neighbor to determine if we should start a new repeat or literal run.
        
        'Start with the special case of "only one byte left in the stream".  This will only occur if we just
        ' ended a repeat run, and the last byte in the line is either 1) another repeat, but the previous run
        ' hit the maximum repeat length, or 2) a unique value that requires its own segment.
        '
        'To avoid slow EOF checks in the inner loop, we want to cover this as a special case (that a branch
        ' predictor can predictably ignore).
        If (srcOffset = scanlineWidth - 1) Then
        
            'Write out the final byte as a 1-byte literal
            dstBytes(dstOffset) = 0
            dstBytes(dstOffset + 1) = srcBytes(srcOffset)
            dstOffset = dstOffset + 2
            srcOffset = srcOffset + 1
            
        'We have room to check neighboring pixels for repeats, without worrying about EOF
        Else
            
            'Compare this byte to its immediate neighbor.  If they match, start a new repeat run.
            targetByte = srcBytes(srcOffset)
            If (targetByte = srcBytes(srcOffset + 1)) Then
            
                'This byte and its neighbor match.  See how long the run continues.
                curRunLength = 2
                
                'Make sure we don't run past the end of the scanline
                Dim maxRunLength As Long
                If (scanlineWidth - srcOffset < MAXIMUM_RUN_LENGTH) Then maxRunLength = scanlineWidth - srcOffset Else maxRunLength = MAXIMUM_RUN_LENGTH
                
                'Count the number of repeats
                Do While (curRunLength < maxRunLength)
                    If (srcBytes(srcOffset + curRunLength) = targetByte) Then curRunLength = curRunLength + 1 Else Exit Do
                Loop
                
                'curRunLength is now one larger than the actual size of the run.  Encode it in the destination
                ' and advance all pointers to match.
                dstBytes(dstOffset) = 256 - (curRunLength - 1)
                dstBytes(dstOffset + 1) = targetByte
                dstOffset = dstOffset + 2
                srcOffset = srcOffset + curRunLength
            
            'This byte and its neighbor do not match.  Start a new literal run.
            Else
                
                curRunLength = 2
                
                'Make sure we don't run past the end of the scanline
                'Dim maxRunLength As Long
                If (scanlineWidth - srcOffset < MAXIMUM_LITERAL_LENGTH) Then maxRunLength = scanlineWidth - srcOffset Else maxRunLength = MAXIMUM_LITERAL_LENGTH
                
                'Count the number of non-repeats
                Do While (curRunLength < maxRunLength)
                    If (srcBytes(srcOffset + curRunLength - 1) <> srcBytes(srcOffset + curRunLength)) Then
                        curRunLength = curRunLength + 1
                    Else
                    
                        'See if this new repeat run is at least three bytes long.  If it isn't, it can actually improve
                        ' compression by keeping those bytes inside this literal run, because 1) we've already eaten the
                        ' "cost" of the flag byte for this run, and 2) if the 3rd byte starts a new literal run, it incurs
                        ' a new flag byte cost - which we can avoid by keeping the current run intact!
                        If (curRunLength + 1 < maxRunLength) Then
                            If (srcBytes(srcOffset + curRunLength) = srcBytes(srcOffset + curRunLength + 1)) Then
                                
                                'Subtract 1, so that the first byte of the forthcoming repeat run is *not* included in
                                ' the current literal run (e.g. so the repeat run starts at the *first* char in the repeat
                                ' sequence, not the *second* one!)
                                curRunLength = curRunLength - 1
                                Exit Do
                                
                            Else
                                curRunLength = curRunLength + 1
                            End If
                        
                        'We're already at maximum length for this run, so we couldn't extend it even if we wanted to.
                        ' Subtract 1 per the note above (to ensure the repeated byte is not included in the
                        ' current literal run), then immediately exit.
                        Else
                            curRunLength = curRunLength - 1
                            Exit Do
                        End If
                        
                    End If
                Loop
                
                'curRunLength is now one larger than the actual size of the run.  Encode it in the destination
                ' and advance all pointers to match.
                dstBytes(dstOffset) = (curRunLength - 1)    '- 1 per the spec, so e.g. an encoded "0" actually represents a length of "1"
                CopyMemoryStrict VarPtr(dstBytes(dstOffset + 1)), VarPtr(srcBytes(srcOffset)), curRunLength
                dstOffset = dstOffset + curRunLength + 1    '+ 1 for the flag byte, of course!
                srcOffset = srcOffset + curRunLength
                
            End If
        
        End If
        
    Loop
    
    'Return the final size before exiting
    dstFinalSize = dstOffset
    
End Sub

Friend Sub WriteChannelData(ByRef cStream As pdStream, ByRef channelCompression As PSD_ChannelCompression, Optional ByVal writePSB As Boolean = False)
    
    'Handle masks for layer groups as a special case
    
    
    'Because channel data has already been separated into planes and (optionally) compressed,
    ' writing the actual data out to file is easy.  The only caveat involves PackBits compression,
    ' because we must first write out RLE tables for *all* channels before writing out actual channel data.
    Dim i As Long
    
    Select Case channelCompression
    
        Case ccRaw
        
            'Dump channel bytes out as-is
            For i = 0 To 3
                'If alpha is irrelevant, do not write it
                If (i = 3) And (Not m_LayerHasAlpha) Then Exit For
                cStream.WriteInt_BE channelCompression
                cStream.WriteByteArray m_Channels(i).ciData, m_Channels(i).ciDataLength
            Next i
        
        Case ccPackBits
        
            'Write RLE tables before writing PackBits stream
            For i = 0 To 3
                
                'If alpha is irrelevant, do not write it
                If (i = 3) And (Not m_LayerHasAlpha) Then Exit For
                
                cStream.WriteInt_BE channelCompression
                
                With m_Channels(i)
                    
                    'The RLE table is stored as Longs, but we need to write it as big-endian ushorts
                    Dim y As Long
                    For y = 0 To UBound(.rleTable)
                        cStream.WriteIntU_BE .rleTable(y)
                    Next y
                    
                    'Dump the channel after it
                    cStream.WriteBytesFromPointer VarPtr(.ciData(0)), .ciDataLength
                    
                End With
                
            Next i
        
        Case ccZip
        
            'Dump channel bytes out as-is
            For i = 0 To 3
                'If alpha is irrelevant, do not write it
                If (i = 3) And (Not m_LayerHasAlpha) Then Exit For
                cStream.WriteInt_BE channelCompression
                cStream.WriteByteArray m_Channels(i).ciData, m_Channels(i).ciDataLength
            Next i
        
        Case ccZipWithPrediction
    
    End Select

End Sub

'Layer groups don't typically have channel data, *unless* they have a mask.  (Either way,
' we do still need to write *some* bytes here.)
Friend Sub WriteChannelData_Group(ByRef srcLayer As pdLayer, ByRef cStream As pdStream, ByRef channelCompression As PSD_ChannelCompression, Optional ByVal writePSB As Boolean = False)
    
    'Failsafe only
    If (srcLayer Is Nothing) Then Exit Sub
    
    If srcLayer.GetLayerMaskExists Then
    
        'Write four dummy entries like we do for un-masked groups...
        cStream.WritePadding 8
        
        '...but then, write out actual mask data as a real entry!
        Select Case channelCompression
        
            'Dump channel bytes out as-is
            Case ccRaw
                cStream.WriteInt_BE channelCompression
                cStream.WriteByteArray m_Channels(0).ciData, m_Channels(0).ciDataLength
            
            'Write RLE tables before writing PackBits stream
            Case ccPackBits
                cStream.WriteInt_BE channelCompression
                
                With m_Channels(0)
                    
                    'The RLE table is stored as Longs, but we need to write it as big-endian ushorts
                    Dim y As Long
                    For y = 0 To UBound(.rleTable)
                        cStream.WriteIntU_BE .rleTable(y)
                    Next y
                    
                    'Dump the channel after it
                    cStream.WriteBytesFromPointer VarPtr(.ciData(0)), .ciDataLength
                    
                End With
            
            'Dump channel bytes out as-is
            Case ccZip
                cStream.WriteInt_BE channelCompression
                cStream.WriteByteArray m_Channels(0).ciData, m_Channels(0).ciDataLength
                
            Case ccZipWithPrediction
                'TODO
        
        End Select
        
    'If the mask doesn't exist, write dummy entries for four placeholder RGB/A channels
    ' (this mirrors what Adobe does)
    Else
        cStream.WritePadding 8
    End If
    
End Sub

'Only use this when writing a composite image at the *end* of a PSD file; it is NOT correct for normal layer data!
Friend Sub WriteSpecialMergedChannelData(ByRef cStream As pdStream, ByRef channelCompression As PSD_ChannelCompression, Optional ByVal writePSB As Boolean = False, Optional ByVal writeAlpha As Boolean = True)

    'Because channel data has already been separated into planes and (optionally) compressed, writing the actual
    ' data out to file is super easy.  The only caveat involves PackBits compression, because we must first write out
    ' RLE tables for *all* channels before writing out actual channel data.
    Dim i As Long
    
    Select Case channelCompression
    
        Case ccRaw
        
            'Dump channel bytes out as-is
            For i = 0 To 3
                If (i = 3) And (Not writeAlpha) Then Exit For
                cStream.WriteByteArray m_Channels(i).ciData, m_Channels(i).ciDataLength
            Next i
        
        Case Else
        
            'Unlike normal layers, we write *all* RLE tables out first, followed by *all* PackBits streams
            For i = 0 To 3
                If (i = 3) And (Not writeAlpha) Then Exit For
                With m_Channels(i)
                    Dim y As Long
                    For y = 0 To UBound(.rleTable)
                        cStream.WriteIntU_BE .rleTable(y)
                    Next y
                End With
            Next i
            
            For i = 0 To 3
                If (i = 3) And (Not writeAlpha) Then Exit For
                With m_Channels(i)
                    cStream.WriteBytesFromPointer VarPtr(.ciData(0)), .ciDataLength
                End With
            Next i
            
    End Select

End Sub

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
