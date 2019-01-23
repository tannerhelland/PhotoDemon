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
'Last updated: 23/January/19
'Last update: implement CMYK support
'
'This class contains layer-specific data pulled from a PSD file.  It is populated by a parent
' pdPSD instance.  It has no purpose outside of a PSD parsing context; for layer handling inside PhotoDemon,
' refer to the pdLayer class.
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

'PSD files contain a *lot* of information.  To aid debugging, you can set output to "verbose"; this will
' dump all kinds of diagnostic information to the debug log.  (Note that the parent pdPSD class has its
' own version of this constant.)
Private Const PSD_DEBUG_VERBOSE As Boolean = True

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
    cidIndexed = 0
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
    DoesLayerHaveUsableData = (m_ChannelCount > 0) And (Not m_LayerDIB Is Nothing) And ((m_Rect.Right - m_Rect.Left) <> 0) And ((m_Rect.Bottom - m_Rect.Top) <> 0)
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
    
    'Support for additional formats are being added all the time.  Reject any unwanted combinations for now.
    Dim mustCancel As Boolean
    mustCancel = False
    If (Not mustCancel) Then mustCancel = (bitsPC > 8)
    'If (Not mustCancel) Then mustCancel = (imgColorMode = cm_CMYK)
    If (Not mustCancel) Then mustCancel = (imgColorMode = cm_Lab)
    If (Not mustCancel) Then mustCancel = (imgColorMode = cm_Multichannel)
    If mustCancel Then
        warningStack.AddString "The current color mode (" & GetNameOfColorMode(imgColorMode) & ") and bpc (" & bitsPC & ") combination is still under construction."
        Set m_LayerDIB = Nothing
        ConstructImage = psd_Failure
        Exit Function
    Else
        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Bit-depth and color mode are supported; attempting to decode now..."
    End If
    
    'Set a flag to determine ICC profile availability; this affects a *lot* of subsequent work in this function,
    ' so we need to reference it early (and often).
    Dim iccAvailable As Boolean
    iccAvailable = (Not srcProfile Is Nothing)
    If iccAvailable Then iccAvailable = srcProfile.HasICCData()
    
    Dim lcmsSrcProfile As pdLCMSProfile, lcmsDstProfile As pdLCMSProfile, lcmsTransform As pdLCMSTransform
    If iccAvailable Then
        
        Set lcmsSrcProfile = New pdLCMSProfile
        lcmsSrcProfile.CreateFromPDICCObject srcProfile
        
        'Destination profile and transform parameters depend on what we find inside the image; they will be
        ' constructed on-the-fly when appropriate.
        Set lcmsDstProfile = New pdLCMSProfile
        Set lcmsTransform = New pdLCMSTransform
        
    End If
                                    
    'Start by preparing a pdDIB object.  This will ultimately hold the results of our parsing, although the
    ' data may get convolved through various intermediary formats first.
    Dim layerWidth As Long, layerHeight As Long
    layerWidth = m_Rect.Right - m_Rect.Left
    layerHeight = m_Rect.Bottom - m_Rect.Top
    
    Set m_LayerDIB = New pdDIB
    m_LayerDIB.CreateBlank layerWidth, layerHeight, 32, vbWhite, 255
    
    'When loading CMYK data, we need to use a different intermediary format: a CMYKA buffer that can
    ' hold five channels worth of data.  (This data will then be color-managed into an RGBA buffer.)
    Dim cmykaBuffer8() As Byte, cmykaBuffer16() As Integer
    If (imgColorMode = cm_CMYK) Then
        
        If (bitsPC = 8) Then
            ReDim cmykaBuffer8(0 To (layerWidth * 5) - 1, 0 To layerHeight - 1) As Byte
            FillMemory VarPtr(cmykaBuffer8(0, 0)), (layerWidth * 5) * layerHeight, 255
            
        'TODO
        ElseIf (bitsPC = 16) Then
        
        End If
        
    End If
    
    'Layers will typically match the color-mode and bit-depth of the parent image, but it can be
    ' convenient to adjust these "on-the-fly" (e.g. bitmap/monochrome layers are inconvenient to
    ' render manually, so we translate them to a grayscale equivalent on the fly).  These parameters
    ' are reset as each channel is loaded, so that changes to one of them does not affect any
    ' other layers in the image.
    Dim chnColorMode As PSD_ColorMode, chnBitDepth As Long
    
    'Channel data is stored in planar format (e.g. not interleaved - RRRGGGBBB instead of RGBRGBRGB).
    ' As such, each channel must be loaded independent of other channels in the image.
    Dim i As Long
    For i = 0 To Abs(m_ChannelCount) - 1
        
        'If the target layer DIB was freed by a previous step, exit immediately
        If (m_LayerDIB Is Nothing) Then Exit For
        
        'Reset any on-the-fly changes we've made to color mode or bit-depth
        chnColorMode = imgColorMode
        chnBitDepth = bitsPC
        
        'Interpretation of the underlying data depends on bit-depth, obviously
        Dim bytesPerChannel As Long
        bytesPerChannel = (chnBitDepth \ 8)
        
        With m_Channels(i)
            
            'If channel data wasn't decoded properly, skip this channel entirely
            If (Not .ciDecodedOK) Then GoTo SkipChannel
            
            'Some channels contribute nothing to the final image (and are zero-length).  Ignore them.
            If (.ciDataLength = 0) Then GoTo SkipChannel
            
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
                
                GoTo SkipChannel
                
            End If
            
            'Monochrome layers are a special case.  If encountered, it's easiest to expand them to an 8-bit stream
            ' and treat the results as grayscale.  (This allows us to reuse the grayscale code "as-is".)
            If (chnBitDepth = 1) Then
                ExpandMonochrome m_Channels(i), layerHeight, layerWidth
                chnColorMode = cm_Grayscale
                chnBitDepth = 8
                bytesPerChannel = 1
            End If
            
            'For color channels, we can calculate the size of the finished channel data from the image's dimensions
            ' and current bit-depth.
            Dim numPixels As Long
            numPixels = layerWidth * layerHeight
            
            Dim finalArraySize As Long
            finalArraySize = numPixels * bytesPerChannel
            
            'Monochrome layers are the only sub-8-bit channel depth that PSDs support.
            ' They also need to be handled separately as their alignment requires specialized handling.
            If (finalArraySize > 0) Then
                
                Dim x As Long, y As Long
                
                Dim tmpBytes() As Byte, tmpSA As SafeArray1D
                Dim scanWidth As Long, scanStart As Long, srcIndex As Long
                scanStart = m_LayerDIB.GetDIBPointer()
                scanWidth = m_LayerDIB.GetDIBStride()
                
                Dim cOffset As Long
                
                'Separate "painting" steps by bit-depth
                
                '8-bit channel (most common case)
                If (bytesPerChannel = 1) Then
                    
                    'RGB color space
                    If (chnColorMode = cm_RGB) Then
                        
                        'Channel order in the file is likely to be RGB, not BGR, so we need to manually
                        ' swizzle order as we go.
                        If (.ciID = cidBlue) Then
                            cOffset = 0
                        ElseIf (.ciID = cidGreen) Then
                            cOffset = 1
                        ElseIf (.ciID = cidRed) Then
                            cOffset = 2
                        ElseIf (.ciID = cidAlpha) Then
                            cOffset = 3
                        End If
                        
                        'TODO: handle masks here, if present; we can simply merge their results with the existing
                        ' alpha channel data, as premultiplication will sort out the rest.
                        
                        'NOTE!  masks represent a more complex problem; Photoshop only stores a sub-portion of the
                        ' mask in the file.  The rect and default mask color setting need to be read to know how
                        ' to paint areas outside said rect, and the number of scanlines is likely to be very different
                        ' (in both width and height) than the layer itself.  Handling this will be potentially complex.
                        
                        'Actually painting the color channel data is extremely simple; just copy the byte into its
                        ' relevant position in the target DIB.
                        m_LayerDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
                        
                        For y = 0 To layerHeight - 1
                            tmpSA.pvData = scanStart + (scanWidth * y)
                        For x = 0 To layerWidth - 1
                            tmpBytes(x * 4 + cOffset) = .ciDataDecoded(y * layerWidth + x)
                        Next x
                        Next y
                        
                        m_LayerDIB.UnwrapArrayFromDIB tmpBytes
                    
                    'CMYK is passed into an intermediary buffer, then color-managed into an RGBA container.
                    ElseIf (chnColorMode = cm_CMYK) Then
                    
                        If (.ciID = cidCyan) Then
                            cOffset = 0
                        ElseIf (.ciID = cidMagenta) Then
                            cOffset = 1
                        ElseIf (.ciID = cidYellow) Then
                            cOffset = 2
                        ElseIf (.ciID = cidKey) Then
                            cOffset = 3
                        ElseIf (.ciID = cidAlpha) Then
                            cOffset = 4
                        End If
                        
                        For y = 0 To layerHeight - 1
                        For x = 0 To layerWidth - 1
                            cmykaBuffer8(x * 5 + cOffset, y) = .ciDataDecoded(y * layerWidth + x)
                        Next x
                        Next y
                        
                    '256-color, duotone, and grayscale images are handled very similarly
                    ElseIf (chnColorMode = cm_Indexed) Or (chnColorMode = cm_Grayscale) Or (chnColorMode = cm_Duotone) Then
                        
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
                                
                            'For grayscale or duotone we manually construct a standard grayscale palette
                            Else
                            
                                ReDim imgPalette(0 To 255) As RGBQuad
                                palSize = 256
                                For x = 0 To 255
                                    imgPalette(x).Blue = x
                                    imgPalette(x).Green = x
                                    imgPalette(x).Red = x
                                    If (x = trnsIndex) Then imgPalette(x).Alpha = 0 Else imgPalette(x).Alpha = 255
                                Next x
                                
                                'While here, if a color profile exists, we want to apply color management prior to
                                ' applying the palette to the image.  We can do this in-place over the decoded
                                ' grayscale pixel data.
                                If iccAvailable Then
                                    
                                    'Build a destination lcms profile and relevant transform
                                    lcmsDstProfile.CreateGenericGrayscaleProfile
                                    lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile, TYPE_GRAY_8, TYPE_GRAY_8, INTENT_PERCEPTUAL
                                    
                                    'Apply the transform in-place to the existing grayscale data
                                    lcmsTransform.ApplyTransformToArbitraryMemory VarPtr(.ciDataDecoded(0)), VarPtr(.ciDataDecoded(0)), layerWidth, layerWidth, layerHeight, layerWidth, False
                                    
                                End If
                                
                            End If
                            
                            'Now we can simply paint the palette into the destination 32-bpp DIB.
                            ' (TODO: accelerate this using a Long array; painting individual bytes is stupid.)
                            m_LayerDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
                            
                            For y = 0 To layerHeight - 1
                                tmpSA.pvData = scanStart + (scanWidth * y)
                            For x = 0 To layerWidth - 1
                                srcIndex = .ciDataDecoded(y * layerWidth + x)
                                tmpBytes(x * 4) = imgPalette(srcIndex).Blue
                                tmpBytes(x * 4 + 1) = imgPalette(srcIndex).Green
                                tmpBytes(x * 4 + 2) = imgPalette(srcIndex).Red
                                tmpBytes(x * 4 + 3) = imgPalette(srcIndex).Alpha
                            Next x
                            Next y
                            
                            m_LayerDIB.UnwrapArrayFromDIB tmpBytes
                        
                        'Any other channels in an indexed image, if they exist, are treated as alpha data.
                        ' TODO: sort out mask handling into a separate function.
                        Else
                        
                            m_LayerDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
                            
                            For y = 0 To layerHeight - 1
                                tmpSA.pvData = scanStart + (scanWidth * y)
                            For x = 0 To layerWidth - 1
                                tmpBytes(x * 4 + 3) = CSng(tmpBytes(x * 4 + 3)) * (CSng(.ciDataDecoded(y * layerWidth + x)) / 255!)
                            Next x
                            Next y
                            
                            m_LayerDIB.UnwrapArrayFromDIB tmpBytes
                            
                        End If
                    
                    'Other color spaces can be handled here in the future
                    'ElseIf (chnColorMode = cm_CMYK) Then    'etc
                        
                    End If
                
                'BPC other than 1 are still under construction.
                Else
                    warningStack.AddString "BPC is currently unsupported; abandoning decode."
                    ConstructImage = psd_Failure
                End If
                
            
            'Failsafe final array size (non zero-length); this layer does not contain usable data
            Else
                ConstructImage = psd_Warning
            End If
            
        End With
    
SkipChannel:
    Next i
    
    'If a full layer image was constructed successfully, apply any relevant post-processing before exiting
    If (Not m_LayerDIB Is Nothing) Then
        
        'With all color channels assembled, 8-bit RGBA images can now be color-managed in-place
        If (imgColorMode = cm_RGB) Then
            If (bitsPC = 8) And iccAvailable Then
                lcmsDstProfile.CreateSRGBProfile
                lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile
                lcmsTransform.ApplyTransformToPDDib m_LayerDIB
            End If
        
        'CMYK data needs to be converted to RGBA for use in PD.
        ElseIf (imgColorMode = cm_CMYK) Then
            
            'TODO: failsafe CMYK path if an ICC profile is not embedded.  This will *never* happen if the PSD
            ' came from Photoshop (and in fact, it should never happen period unless a writer is broken), but
            ' I imagine there are plenty of broken PSDs in the wild...
            If (bitsPC = 8) And iccAvailable Then
                lcmsDstProfile.CreateSRGBProfile
                lcmsTransform.CreateTwoProfileTransform lcmsSrcProfile, lcmsDstProfile, TYPE_CMYKA_8, TYPE_BGRA_8
                lcmsTransform.ApplyTransformToArbitraryMemory VarPtr(cmykaBuffer8(0, 0)), m_LayerDIB.GetDIBPointer, layerWidth * 5, layerWidth * 4, layerHeight, layerWidth
                Erase cmykaBuffer8
            End If
        End If
        
        'TODO: only premultiply if an alpha channel was found in the image; otherwise, we can guarantee full opacity
        'Premultiply alpha before exiting
        m_LayerDIB.SetAlphaPremultiplication True
    
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ConstructImage", "internal VB error #" & Err.Number & ": " & Err.Description
    warningStack.AddString "Internal error in ConstructImage, #" & Err.Number & ": " & Err.Description
    ConstructImage = psd_Failure
    
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
                
                If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Channel #" & i & " (ID=" & .ciID & ") has compression type: " & .ciCompression
                
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
                                bytesReadLine = bytesReadLine + 1
                                
                                'If the code is >= 0, the next (code) bytes should be copied to the destination
                                ' buffer as-is.
                                If (rleCode > 0) Then
                                    rleCode = rleCode + 1
                                    tmpVal = rleStream.ReadBytes_PointerOnly(rleCode)
                                    bytesReadLine = bytesReadLine + rleCode
                                    'If (bytesReadLine > lineLengths(y)) Then
                                    '    Debug.Print "WARNING RLE OVERFLOW (1)! " & lineLengths(y), y, rleCode
                                    '    Exit Do
                                    'End If
                                    CopyMemoryStrict VarPtr(.ciDataDecoded(bytesReadTotal)), tmpVal, rleCode
                                    bytesReadTotal = bytesReadTotal + rleCode
                                    
                                'If the code is on the range [-127,-1], inclusive, the next byte should be
                                ' repeated [1 - code] times in the destination buffer.
                                ElseIf (rleCode > -128) Then
                                    rleCode = 1 - rleCode
                                    tmpVal = rleStream.ReadByte()
                                    bytesReadLine = bytesReadLine + 1
                                    'If (bytesReadLine > lineLengths(y)) Then
                                    '    Debug.Print "WARNING RLE OVERFLOW (2)! " & lineLengths(y), y, rleCode
                                    '    Exit Do
                                    'End If
                                    FillMemory VarPtr(.ciDataDecoded(bytesReadTotal)), rleCode, tmpVal
                                    bytesReadTotal = bytesReadTotal + rleCode
                                    
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
                        
                    Case ccZip
                        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "ZIP compression isn't supported yet."
                        .ciDecodedOK = False
                        DecodeChannels = psd_Failure
                        
                    Case ccZipWithPrediction
                        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "ZIP w/prediction compression isn't supported yet."
                        .ciDecodedOK = False
                        DecodeChannels = psd_Failure
                        
                End Select
                
            'Raw data is zero-length; this layer does not contain usable data
            Else
                DecodeChannels = psd_Warning
            End If
            
            'Before moving to the next channel, free the encoded data; we no longer require it
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
            
        'TODO!
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
    If (srcStream.GetPosition() < finalPointerPos) Then
        
        ReDim m_AdditionalInfo(0 To 3) As PSD_AdditionalInfo
        m_AdditionalInfoCount = 0
        
        Dim sigCheck As String
        
        Do While (srcStream.GetPosition() < finalPointerPos)
            
            'Verify the signature
            sigCheck = srcStream.ReadString_ASCII(4)
            
            If (sigCheck <> "8BIM") And (sigCheck <> "8B64") Then
                warningStack.AddString "ParseLayer found an unknown additional segment signature: " & sigCheck
                ParseLayer = psd_Warning
            Else
                
                If (m_AdditionalInfoCount > UBound(m_AdditionalInfo)) Then ReDim Preserve m_AdditionalInfo(0 To m_AdditionalInfoCount * 2 - 1) As PSD_AdditionalInfo
                With m_AdditionalInfo(m_AdditionalInfoCount)
                    .aiKey = srcStream.ReadString_ASCII(4)
                    .aiDataLength = srcStream.ReadLong_BE()
                    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Additional layer info found: " & .aiKey & " - " & .aiDataLength & " bytes"
                    If (.aiDataLength > 0) Then srcStream.ReadBytes .aiDataBytes, .aiDataLength, True
                End With
                
                m_AdditionalInfoCount = m_AdditionalInfoCount + 1
                
            End If
            
        Loop
        
        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Layer contained " & m_AdditionalInfoCount & " additional info records"
    
    'Layers are not required to have additional info records; this is okay
    Else
        
    End If
    
    'Search additional info chunks for settings useful to PD
    If (m_AdditionalInfoCount > 0) Then
        
        'Additional info structs may be large and complicated; pdStream simplifies traversal
        Dim tmpStream As pdStream
        Set tmpStream = New pdStream
                    
        For i = 0 To m_AdditionalInfoCount - 1
        
            Select Case m_AdditionalInfo(i).aiKey
            
                'Unicode layer names are important; they supercede the (required) ANSI layer name, if any
                Case "luni"
                
                    'Data chunk is a 4-byte string length (chars, not bytes), followed by [length] wchars
                    tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , m_AdditionalInfo(i).aiDataLength, VarPtr(m_AdditionalInfo(i).aiDataBytes(0))
                    
                    Dim layerNameLengthChars As Long
                    layerNameLengthChars = tmpStream.ReadLong_BE()
                    
                    If (layerNameLengthChars <= ((m_AdditionalInfo(i).aiDataLength - 4) \ 2)) And (layerNameLengthChars > 0) Then
                        m_LayerNameUnicode = tmpStream.ReadString_Unicode_BE(layerNameLengthChars)
                        m_LayerNameUnicode = Strings.TrimNull(m_LayerNameUnicode)
                    End If
                    
                    'Important to free stream, as it unsafely wraps the underlying chunk data
                    tmpStream.StopStream
                    
                    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Layer name (Unicode): " & m_LayerNameUnicode
            
            End Select
        
        Next i
        
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
    
    'Variable amounts of padding; ignore this and simply align the pointer manually
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
