Attribute VB_Name = "Plugin_jxl"
'***************************************************************************
'JPEG-XL Reference Library (libjxl) Interface
'Copyright 2022-2022 by Tanner Helland
'Created: 28/September/22
'Last updated: 13/October/22
'Last update: importing static images is working well!  I want to support animations too, but don't have any
'             good ones to test with - so I think I'm going to move over to export support so I can generate
'             animations for testing purposes.
'
'libjxl (available at https://github.com/libjxl/libjxl) is the official reference library implementation
' for the modern JPEG-XL format.  Support for this format was added during the PhotoDemon 10.0 release cycle.
'
'Unfortunately for Windows XP users, libjxl currently requires Windows Vista or later.  PhotoDemon will
' detect this automatically and gracefully hide JPEG XL support for XP users.  (If anyone knows how to build
' libjxl in an XP-compatible way, I would happily welcome a pull request...)
'
'PhotoDemon tries to support most JPEG XL features, but esoteric ones (like animation) remain a WIP.
' If you encounter any issues with JPEG XL images, please file an issue on GitHub and attach the image(s)
' in question so I can investigate further.
'
'This wrapper class uses a shorthand wrapper to DispCallFunc originally written by Olaf Schmidt.
' Many thanks to Olaf, whose original version can be found here (link good as of Feb 2019):
' http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)&p=4795471&viewfull=1#post4795471
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'DO NOT enable verbose logging in production builds
Private Const JXL_DEBUG_VERBOSE As Boolean = True

'Return values when attempting to assess "is this data in JXL format?"
Private Enum JxlSignature
    JXL_SIG_NOT_ENOUGH_BYTES = 0 '/* Not enough bytes were passed to determine if a valid signature was found.
    JXL_SIG_INVALID = 1          '/* No valid JPEG XL header was found.
    JXL_SIG_CODESTREAM = 2       '/* A valid JPEG XL codestream signature was found, that is a JPEG XL image without container.
    JXL_SIG_CONTAINER = 3        '/* A valid container signature was found, that is a JPEG XL image embedded in a box format container.
End Enum

#If False Then
    Private Const JXL_SIG_NOT_ENOUGH_BYTES = 0, JXL_SIG_INVALID = 1, JXL_SIG_CODESTREAM = 2, JXL_SIG_CONTAINER = 3
#End If

'Return value for JxlDecoderProcessInput.
' The values from JXL_DEC_BASIC_INFO onwards are optional informative events that can be subscribed to,
' they are never returned if they have not been registered with @ref JxlDecoderSubscribeEvents.
Private Enum JxlDecoderStatus
    
    'Function call finished successfully, or decoding is finished and there is nothing more to be done.
    ' (Note that JxlDecoderProcessInput will return JXL_DEC_SUCCESS if all events that were registered
    ' with JxlDecoderSubscribeEvents were processed, even before the end of the JPEG XL codestream.
    ' In this case, the return value JxlDecoderReleaseInput will be the same as it was at the last
    ' signaled event. E.g. if JXL_DEC_FULL_IMAGE was subscribed to, then all bytes from the end of the
    ' JPEG XL codestream (including possible boxes needed for jpeg reconstruction) will be returned
    ' as unprocessed.)
    JXL_DEC_SUCCESS = 0
    
    'An error occurred, for example invalid input file or out of memory.
    JXL_DEC_ERROR = 1
    
    'The decoder needs more input bytes to continue.
    ' Before the next JxlDecoderProcessInput call, more input data must be set, by calling JxlDecoderReleaseInput
    ' (if input was set previously) and then calling JxlDecoderSetInput. JxlDecoderReleaseInput returns how many
    ' bytes are not yet processed, before a next call to JxlDecoderProcessInput all unprocessed bytes must be
    ' provided again (the address need not match, but the contents must), and more bytes must be concatenated
    ' after the unprocessed bytes.
    '
    'In most cases, JxlDecoderReleaseInput will return no unprocessed bytes at this event, the only exceptions
    ' are if the previously set input ended within (a) the raw codestream signature, (b) the signature box,
    ' (c) a box header, or (d) the first 4 bytes of a brob, ftyp, or jxlp box. In any of these cases the number
    ' of unprocessed bytes is less than 20.
    JXL_DEC_NEED_MORE_INPUT = 2
    
    'The decoder is able to decode a preview image and requests setting a preview output buffer using
    ' JxlDecoderSetPreviewOutBuffer.  This occurs if JXL_DEC_PREVIEW_IMAGE is requested and it is possible
    ' to decode a preview image from the codestream and the preview out buffer was not yet set. There is
    ' maximum one preview image in a codestream.
    '
    'In this case, JxlDecoderReleaseInput will return all bytes from the end of the frame header (including ToC)
    ' of the preview frame as unprocessed.
    JXL_DEC_NEED_PREVIEW_OUT_BUFFER = 3
    
    'The decoder is able to decode a DC image and requests setting a DC output buffer using
    ' JxlDecoderSetDCOutBuffer. This occurs if JXL_DEC_DC_IMAGE is requested and it is possible to decode a
    ' DC image from the codestream and the DC out buffer was not yet set. This event re-occurs for new frames
    ' if there are multiple animation frames.
    '
    'The DC feature in this form will be removed. For progressive rendering, JxlDecoderFlushImage should be used.
    JXL_DEC_NEED_DC_OUT_BUFFER = 4
    
    'The decoder requests an output buffer to store the full resolution image, which can be set with
    ' JxlDecoderSetImageOutBuffer or with JxlDecoderSetImageOutCallback. This event re-occurs for new frames
    ' if there are multiple animation frames and requires setting an output again. In this case,
    ' JxlDecoderReleaseInput will return all bytes from the end of the frame header (including ToC) as
    ' unprocessed.
    JXL_DEC_NEED_IMAGE_OUT_BUFFER = 5
    
    'The JPEG reconstruction buffer is too small for reconstructed JPEG codestream to fit.
    ' JxlDecoderSetJPEGBuffer must be called again to make room for remaining bytes. This event may occur
    ' multiple times after JXL_DEC_JPEG_RECONSTRUCTION.
    JXL_DEC_JPEG_NEED_MORE_OUTPUT = 6
    
    'The box contents output buffer is too small. JxlDecoderSetBoxBuffer must be called again to make room
    ' for remaining bytes. This event may occur multiple times after JXL_DEC_BOX.
    JXL_DEC_BOX_NEED_MORE_OUTPUT = 7
    
    'Informative event by JxlDecoderProcessInput.
    ' Basic information such as image dimensions and extra channels. This event occurs max once per image.
    ' In this case, JxlDecoderReleaseInput will return all bytes from the end of the basic info as unprocessed
    ' (including the last byte of basic info if it did not end on a byte boundary).
    JXL_DEC_BASIC_INFO = &H40&
    
    'Informative event by JxlDecoderProcessInput.
    ' User extensions of the codestream header. This event occurs max once per image and always later than
    ' JXL_DEC_BASIC_INFO and earlier than any pixel data.
    '
    'DEPRECATED: The decoder no longer returns this.  The header extensions, if any, are available at the
    ' JXL_DEC_BASIC_INFO event.
    JXL_DEC_EXTENSIONS = &H80&
    
    'Informative event by JxlDecoderProcessInput
    ' Color encoding or ICC profile from the codestream header. This event occurs max once per image and
    ' always later than JXL_DEC_BASIC_INFO and earlier than any pixel data. In this case, JxlDecoderReleaseInput
    ' will return all bytes from the end of the image header (which is the start of the first frame) as unprocessed.
    JXL_DEC_COLOR_ENCODING = &H100&
    
    'Informative event by JxlDecoderProcessInput
    ' Preview image, a small frame, decoded. This event can only happen if the image has a preview frame encoded.
    ' This event occurs max once for the codestream and always later than JXL_DEC_COLOR_ENCODING and before
    ' JXL_DEC_FRAME.  In this case, JxlDecoderReleaseInput will return all bytes from the end of the preview frame
    ' as unprocessed.
    JXL_DEC_PREVIEW_IMAGE = &H200&
    
    'Informative event by JxlDecoderProcessInput.
    ' Beginning of a frame. JxlDecoderGetFrameHeader can be used at this point.
    '
    'A note on frames: a JPEG XL image can have internal frames that are not intended to be displayed
    ' (e.g. used for compositing a final frame), but this only returns displayed frames, unless
    ' JxlDecoderSetCoalescing was set to JXL_FALSE: in that case, the individual layers are returned without blending.
    ' Note that even when coalescing is disabled, only frames of type kRegularFrame are returned; frames of type
    ' kReferenceOnly and kLfFrame are always for internal purposes only and cannot be accessed. A displayed frame
    ' either has an animation duration or is the only or last frame in the image. This event occurs max once per
    ' displayed frame, always later than JXL_DEC_COLOR_ENCODING, and always earlier than any pixel data.
    '
    'While JPEG XL supports encoding a single frame as the composition of multiple internal sub-frames also
    ' called frames, this event is not indicated for the internal frames.
    '
    'In this case, JxlDecoderReleaseInput will return all bytes from the end of the frame header (including ToC)
    ' as unprocessed.
    JXL_DEC_FRAME = &H400&
    
    'Informative event by JxlDecoderProcessInput.
    ' DC image, 8x8 sub-sampled frame, decoded. It is not guaranteed that the decoder will always return DC
    ' separately, but when it does it will do so before outputting the full frame. JxlDecoderSetDCOutBuffer
    ' must be used after getting the basic image information to be able to get the DC pixels, if not this
    ' return status only indicates we're past this point in the codestream. This event occurs max once per frame
    ' and always later than JXL_DEC_FRAME and other header events and earlier than full resolution pixel data.
    '
    'DEPRECATED: The DC feature in this form will be removed. For progressive rendering, JxlDecoderFlushImage
    ' should be used.
    JXL_DEC_DC_IMAGE = &H800&
    
    'Informative event by JxlDecoderProcessInput.
    ' full frame (or layer, in case coalescing is disabled) is decoded. JxlDecoderSetImageOutBuffer must be used
    ' after getting the basic image information to be able to get the image pixels, if not this return status only
    ' indicates we're past this point in the codestream. This event occurs max once per frame and always later than
    ' JXL_DEC_DC_IMAGE.  In this case, JxlDecoderReleaseInput will return all bytes from the end of the frame
    ' (or if JXL_DEC_JPEG_RECONSTRUCTION is subscribed to, from the end of the last box that is needed for jpeg
    ' reconstruction) as unprocessed.
    JXL_DEC_FULL_IMAGE = &H1000&
    
    'Informative event by JxlDecoderProcessInput.
    ' JPEG reconstruction data decoded.  JxlDecoderSetJPEGBuffer may be used to set a JPEG reconstruction buffer
    ' after getting the JPEG reconstruction data. If a JPEG reconstruction buffer is set a byte stream identical
    ' to the JPEG codestream used to encode the image will be written to the JPEG reconstruction buffer instead
    ' of pixels to the image out buffer. This event occurs max once per image and always before JXL_DEC_FULL_IMAGE.
    ' In this case, JxlDecoderReleaseInput will return all bytes from the end of the 'jbrd' box as unprocessed.
    JXL_DEC_JPEG_RECONSTRUCTION = &H2000&
    
    'Informative event by JxlDecoderProcessInput.
    ' The header of a box of the container format (BMFF) is decoded. The following API functions related to boxes
    ' can be used after this event:
    ' - JxlDecoderSetBoxBuffer and JxlDecoderReleaseBoxBuffer
    ' - "JxlDecoderReleaseBoxBuffer": set and release a buffer to get the box data.
    ' - JxlDecoderGetBoxType: get the 4-character box typename.
    ' - JxlDecoderGetBoxSizeRaw get the size of the box as it appears in the container file, not decompressed.
    ' - JxlDecoderSetDecompressBoxes to configure whether to get the box data decompressed, or possibly compressed.
    '
    'Boxes can be compressed. This is so when their box type is "brob". In that case, they have an underlying
    ' decompressed box type and decompressed data. JxlDecoderSetDecompressBoxes allows configuring which data to get.
    ' Decompressing requires Brotli. JxlDecoderGetBoxType has a flag to get the compressed box type, which can be
    ' "brob", or the decompressed box type. If a box is not compressed (its compressed type is not "brob"), then the
    ' output decompressed box type and data is independent of what setting is configured.
    '
    'The buffer set with JxlDecoderSetBoxBuffer must be set again for each next box to be obtained, or can be left
    ' unset to skip outputting this box.  The output buffer contains the full box data when the next JXL_DEC_BOX
    ' event or JXL_DEC_SUCCESS occurs. JXL_DEC_BOX occurs for all boxes, including non-metadata boxes such as the
    ' signature box or codestream boxes. To check whether the box is a metadata type for respectively EXIF, XMP or
    ' JUMBF, use JxlDecoderGetBoxType and check for types "Exif", "xml " and "jumb" respectively.
    '
    'In this case, JxlDecoderReleaseInput will return all bytes from the start of the box header as unprocessed.
    JXL_DEC_BOX = &H4000&

    'Informative event by JxlDecoderProcessInput.
    ' A progressive step in decoding the frame is reached. When calling JxlDecoderFlushImage at this point,
    ' the flushed image will correspond exactly to this point in decoding, and not yet contain partial results
    ' (such as partially more fine detail) of a next step. By default, this event will trigger maximum once per frame,
    ' when a 8x8th resolution (DC) image is ready (the image data is still returned at full resolution, giving upscaled
    ' DC). Use JxlDecoderSetProgressiveDetail to configure more fine-grainedness. The event is not guaranteed to trigger,
    ' not all images have progressive steps or DC encoded.
    '
    'In this case, JxlDecoderReleaseInput will return all bytes from the end of the section that was needed to produce
    ' this progressive event as unprocessed.
    JXL_DEC_FRAME_PROGRESSION = &H8000&

End Enum

#If False Then
    Private Const JXL_DEC_SUCCESS = 0, JXL_DEC_ERROR = 1, JXL_DEC_NEED_MORE_INPUT = 2, JXL_DEC_NEED_PREVIEW_OUT_BUFFER = 3, JXL_DEC_NEED_DC_OUT_BUFFER = 4, JXL_DEC_NEED_IMAGE_OUT_BUFFER = 5, JXL_DEC_JPEG_NEED_MORE_OUTPUT = 6, JXL_DEC_BOX_NEED_MORE_OUTPUT = 7, JXL_DEC_BASIC_INFO = &H40&, JXL_DEC_EXTENSIONS = &H80&, JXL_DEC_COLOR_ENCODING = &H100&, JXL_DEC_PREVIEW_IMAGE = &H200&, JXL_DEC_FRAME = &H400&, JXL_DEC_DC_IMAGE = &H800&, JXL_DEC_FULL_IMAGE = &H1000&, JXL_DEC_JPEG_RECONSTRUCTION = &H2000&, JXL_DEC_BOX = &H4000&, JXL_DEC_FRAME_PROGRESSION = &H8000&
#End If

'Image orientation metadata. Values 1..8 match the EXIF definitions.
' The name indicates the operation to perform to transform from the encoder image to the display image.
Private Enum JxlOrientation
    JXL_ORIENT_IDENTITY = 1
    JXL_ORIENT_FLIP_HORIZONTAL = 2
    JXL_ORIENT_ROTATE_180 = 3
    JXL_ORIENT_FLIP_VERTICAL = 4
    JXL_ORIENT_TRANSPOSE = 5
    JXL_ORIENT_ROTATE_90_CW = 6
    JXL_ORIENT_ANTI_TRANSPOSE = 7
    JXL_ORIENT_ROTATE_90_CCW = 8
End Enum

#If False Then
    Private Const JXL_ORIENT_IDENTITY = 1, JXL_ORIENT_FLIP_HORIZONTAL = 2, JXL_ORIENT_ROTATE_180 = 3, JXL_ORIENT_FLIP_VERTICAL = 4, JXL_ORIENT_TRANSPOSE = 5, JXL_ORIENT_ROTATE_90_CW = 6, JXL_ORIENT_ANTI_TRANSPOSE = 7, JXL_ORIENT_ROTATE_90_CCW = 8
#End If

'Given type of an extra channel.
Private Enum JxlExtraChannelType
    JXL_CHANNEL_ALPHA
    JXL_CHANNEL_DEPTH
    JXL_CHANNEL_SPOT_COLOR
    JXL_CHANNEL_SELECTION_MASK
    JXL_CHANNEL_BLACK
    JXL_CHANNEL_CFA
    JXL_CHANNEL_THERMAL
    JXL_CHANNEL_RESERVED0
    JXL_CHANNEL_RESERVED1
    JXL_CHANNEL_RESERVED2
    JXL_CHANNEL_RESERVED3
    JXL_CHANNEL_RESERVED4
    JXL_CHANNEL_RESERVED5
    JXL_CHANNEL_RESERVED6
    JXL_CHANNEL_RESERVED7
    JXL_CHANNEL_UNKNOWN
    JXL_CHANNEL_OPTIONAL
End Enum

#If False Then
    Private Const JXL_CHANNEL_ALPHA = 0, JXL_CHANNEL_DEPTH = 0, JXL_CHANNEL_SPOT_COLOR = 0, JXL_CHANNEL_SELECTION_MASK = 0, JXL_CHANNEL_BLACK = 0, JXL_CHANNEL_CFA = 0, JXL_CHANNEL_THERMAL = 0, JXL_CHANNEL_RESERVED0 = 0, JXL_CHANNEL_RESERVED1 = 0, JXL_CHANNEL_RESERVED2 = 0, JXL_CHANNEL_RESERVED3 = 0, JXL_CHANNEL_RESERVED4 = 0, JXL_CHANNEL_RESERVED5 = 0, JXL_CHANNEL_RESERVED6 = 0, JXL_CHANNEL_RESERVED7 = 0, JXL_CHANNEL_UNKNOWN = 0, JXL_CHANNEL_OPTIONAL = 0
#End If

'Codestream preview header
Private Type JxlPreviewHeader
    
    'Preview width in pixels
    xsize As Long
    
    'Preview height in pixels
    ysize As Long
    
End Type

'Intrinsic size header
Private Type JxlIntrinsicSizeHeader
    
    'Intrinsic width in pixels
    xsize As Long
    
    'Intrinsic height in pixels
    ysize As Long
    
End Type

'Codestream animation header, optionally present in the beginning of the codestream, and if it is it applies
' to all animation frames (unlike JxlFrameHeader which applies to an individual frame).
Private Type JxlAnimationHeader
    
    'Numerator of ticks per second of a single animation frame time unit
    tps_numerator As Long
    
    'Denominator of ticks per second of a single animation frame time unit
    tps_denominator As Long
    
    'Amount of animation loops, or 0 to repeat infinitely
    num_loops As Long
    
    'Whether animation time codes are present at animation frames in the codestream
    have_timecodes As Long
    
End Type

'Basic image information. This information is available from the file
' signature and first part of the codestream header.
Private Type JxlBasicInfo
    
    'Whether the codestream is embedded in the container format. If true, metadata information and
    ' extensions may be available in addition to the codestream.
    have_container As Long
    
    'Width of the image in pixels, before applying orientation.
    xsize As Long
    
    'Height of the image in pixels, before applying orientation.
    ysize As Long
    
    'Original image color channel bit depth.
    bits_per_sample As Long
    
    'Original image color channel floating point exponent bits, or 0 if they are unsigned integer.
    ' For example, if the original data is half-precision (binary16) floating point, bits_per_sample is 16
    ' and exponent_bits_per_sample is 5, and so on for other floating point precisions.
    exponent_bits_per_sample As Long
    
    'Upper bound on the intensity level present in the image in nits. For unsigned integer pixel encodings,
    ' this is the brightness of the largest representable value. The image does not necessarily contain a pixel
    ' actually this bright. An encoder is allowed to set 255 for SDR images without computing a histogram.
    '
    'Leaving this set to its default of 0 lets libjxl choose a sensible default value based on the color encoding.
    intensity_target As Single
    
    'Lower bound on the intensity level present in the image. This may be loose, i.e. lower than the actual
    ' darkest pixel. When tone mapping, a decoder will map [min_nits, intensity_target] to the display range.
    min_nits As Single
    
    'See the description of @see linear_below.
    relative_to_max_display As Long

    'The tone mapping will leave unchanged (linear mapping) any pixels whose brightness is strictly below this.
    ' The interpretation depends on relative_to_max_display. If true, this is a ratio [0, 1] of the maximum
    ' display brightness [nits], otherwise an absolute brightness [nits].
    linear_below As Single
    
    'Whether the data in the codestream is encoded in the original color profile that is attached to the
    ' codestream metadata header, or is encoded in an internally supported absolute color space (which the
    ' decoder can always convert to linear or non-linear sRGB or to XYB).
    '
    'If the original profile is used, the decoder outputs pixel data in the color space matching that profile,
    ' but doesn't convert it to any other color space. If the original profile is not used, the decoder only
    ' outputs the data as sRGB (linear if outputting to floating point, nonlinear with standard sRGB transfer
    ' function if outputting to unsigned integers) but will not convert it to to the original color profile.
    ' The decoder also does not convert to the target display color profile.
    '
    'To convert the pixel data produced by the decoder to the original color profile, one of the
    ' JxlDecoderGetColor* functions needs to be called with @ref JXL_COLOR_PROFILE_TARGET_DATA to get the
    ' color profile of the decoder output, and then an external CMS can be used for conversion.
    '
    'Note that for lossy compression, this should be set to false for most use cases, and if needed, the image
    ' should be converted to the original color profile after decoding, as described above.
    uses_original_profile As Long
    
    'Indicates a preview image exists near the beginning of the codestream. The preview itself or its
    ' dimensions are not included in the basic info.
    have_preview As Long
    
    'Indicates animation frames exist in the codestream. The animation information is not included in the
    ' basic info.
    have_animation As Long
    
    'Image orientation, value 1-8 matching the values used by JEITA CP-3451C (Exif version 2.3).
    orientation_image As JxlOrientation
    
    'Number of color channels encoded in the image, this is either 1 for grayscale data or 3 for color data.
    ' This count does not include the alpha channel or other extra channels. To check presence of an alpha
    ' channel, such as in the case of RGBA color, check alpha_bits != 0.
    '
    'If and only if this is 1, the JxlColorSpace in the JxlColorEncoding is JXL_COLOR_SPACE_GRAY.
    num_color_channels As Long
    
    'Number of additional image channels. This includes the main alpha channel, but can also include
    ' additional channels such as depth, additional alpha channels, spot colors, and so on.
    '
    'Information about extra channels can be queried with JxlDecoderGetExtraChannelInfo. The main alpha channel,
    ' if it exists, also has its information available in the alpha_bits, alpha_exponent_bits and
    ' alpha_premultiplied fields in this JxlBasicInfo.
    num_extra_channels As Long
    
    'Bit depth of the encoded alpha channel, or 0 if there is no alpha channel.  If present, matches the
    ' alpha_bits value of the JxlExtraChannelInfo associated with this alpha channel.
    alpha_bits As Long
    
    'Alpha channel floating point exponent bits, or 0 if they are unsigned. If present, matches the alpha_bits
    ' value of the JxlExtraChannelInfo associated with this alpha channel.
    alpha_exponent_bits As Long
    
    'Whether the alpha channel is premultiplied. Only used if there is a main alpha channel.
    ' Matches the alpha_premultiplied value of the JxlExtraChannelInfo associated with this alpha channel.
    alpha_premultiplied As Long

    'Dimensions of encoded preview image, only used if have_preview is JXL_TRUE.
    preview_header As JxlPreviewHeader
    
    'Animation header with global animation properties for all frames, only used if have_animation is JXL_TRUE.
    animation_header As JxlAnimationHeader
    
    'Intrinsic width of the image.
    ' The intrinsic size can be different from the actual size in pixels (as given by xsize and ysize)
    ' and it denotes the recommended dimensions for displaying the image, i.e. applications are advised
    ' to resample the decoded image to the intrinsic dimensions.
    intrinsic_xsize As Long
    
    'Intrinsic height of the image.
    ' The intrinsic size can be different from the actual size in pixels (as given by xsize and ysize)
    ' and it denotes the recommended dimensions for displaying the image, i.e. applications are advised
    ' to resample the decoded image to the intrinsic dimensions.
    intrinsic_ysize As Long
    
    'Padding for forwards-compatibility, in case more fields are exposed in a future version of the library.
    jxl_padding(100) As Byte
    
End Type

'Information for a single extra channel.
Private Type JxlExtraChannelInfo
    
    'Given type of an extra channel.
    type_of_channel As JxlExtraChannelType
    
    'Total bits per sample for this channel.
    bits_per_sample As Long
    
    'Floating point exponent bits per channel, or 0 if they are unsigned integer.
    exponent_bits_per_sample As Long
    
    'The exponent the channel is downsampled by on each axis.
    dim_shift As Long
    
    'Length of the extra channel name in bytes, or 0 if no name.  Excludes null termination character.
    name_length As Long
    
    'Whether alpha channel uses premultiplied alpha. Only applicable if type is JXL_CHANNEL_ALPHA.
    alpha_premultiplied As Long
    
    'Spot color of the current spot channel in linear RGBA. Only applicable if type is JXL_CHANNEL_SPOT_COLOR.
    spot_color(4) As Single
    
    'Only applicable if type is JXL_CHANNEL_CFA.
    cfa_channel As Long
    
End Type

'Extensions in the codestream header.
Private Type JxlHeaderExtensions
    extensions As Currency
End Type

'Frame blend modes.
' (When decoding, if coalescing is enabled (default), this can be ignored.)
Private Enum JxlBlendMode
    JXL_BLEND_REPLACE = 0
    JXL_BLEND_ADD = 1
    JXL_BLEND_BLEND = 2
    JXL_BLEND_MULADD = 3
    JXL_BLEND_MUL = 4
End Enum

#If False Then
    Private Const JXL_BLEND_REPLACE = 0, JXL_BLEND_ADD = 1, JXL_BLEND_BLEND = 2, JXL_BLEND_MULADD = 3, JXL_BLEND_MUL = 4
#End If

'The information about blending the color channels or a single extra channel.
' When decoding, if coalescing is enabled (default), this can be ignored and the blend mode is considered to
' be JXL_BLEND_REPLACE. When encoding, these settings apply to the pixel data given to the encoder.
Private Type JxlBlendInfo
  
    'Blend mode.
    jBlendMode As JxlBlendMode
    
    'Reference frame ID to use as the 'bottom' layer (0-3).
    sourceID As Long
    
    'Which extra channel to use as the 'alpha' channel for blend modes JXL_BLEND_BLEND and JXL_BLEND_MULADD.
    alphaChannelID As Long
    
    'Clamp values to [0,1] for the purpose of blending.
    bool_clamp As Long
    
End Type

'Information about layers.
' When decoding, if coalescing is enabled (default), this can be ignored.
' When encoding, these settings apply to the pixel data given to the encoder.  The encoder may choose an
' internal representation that differs.
Private Type JxlLayerInfo

    'Whether cropping is applied for this frame.
    ' When decoding, if false, crop_x0 and crop_y0 are set to zero, and xsize and ysize to the main image
    ' dimensions.  If coalescing is enabled (default), this is always false, regardless of the internal
    ' encoding in the JPEG XL codestream.)
    ' When encoding and this is false, those fields are ignored.
    bool_have_crop As Long
    
    'Horizontal offset of the frame (can be negative).
    crop_x0 As Long
    
    'Vertical offset of the frame (can be negative).
    crop_y0 As Long
    
    'Width of the frame (number of columns).
    xsize As Long
    
    'Height of the frame (number of rows).
    ysize As Long
    
    'The blending info for the color channels. Blending info for extra channels has to be retrieved
    ' separately using JxlDecoderGetExtraChannelBlendInfo.
    blend_info As JxlBlendInfo
    
    'After blending, save the frame as reference frame with this ID (0-3).
    ' Special case: if the frame duration is nonzero, ID 0 means "will not be referenced in the future".
    ' This value is not used for the last frame.
    save_as_reference As Long
    
End Type

'The header of one displayed frame or non-coalesced layer.
Private Type JxlFrameHeader

    'How long to wait after rendering in ticks. The duration in seconds of a tick is given by
    ' tps_numerator and tps_denominator in JxlAnimationHeader.
    duration As Long
    
    'SMPTE timecode of the current frame in form 0xHHMMSSFF, or 0. The bits are interpreted from
    ' most-significant to least-significant as hour, minute, second, and frame. If timecode is nonzero,
    ' it is strictly larger than that of a previous frame with nonzero duration. These values are only
    ' available if have_timecodes in JxlAnimationHeader is JXL_TRUE.
    timecode As Long
    
    'Length of the frame name in bytes, or 0 if no name.  Excludes null termination character.
    ' This value is set by the decoder. For the encoder, this value is ignored and JxlEncoderSetFrameName
    ' is used instead to set the name and the length.
    name_length As Long
    
    'Indicates this is the last animation frame. This value is set by the decoder to indicate no further
    ' frames follow. For the encoder, it is not required to set this value and it is ignored;
    ' JxlEncoderCloseFrames is used to indicate the last frame to the encoder instead.
    bool_is_last As Long
    
    'Information about the layer in case of no coalescing.
    layer_info As JxlLayerInfo

End Type

'Data type for the sample values per channel per pixel.
Private Enum JxlDataType
    
    'Use 32-bit single-precision floating point values, with range 0.0-1.0 (within gamut, may go outside
    ' this range for wide color gamut). Floating point output, either JXL_TYPE_FLOAT or JXL_TYPE_FLOAT16,
    ' is recommended for HDR and wide gamut images when color profile conversion is required.
    JXL_TYPE_FLOAT = 0
    
    'Use type uint8_t. May clip wide color gamut data.
    JXL_TYPE_UINT8 = 2

    'Use type uint16_t. May clip wide color gamut data.
    JXL_TYPE_UINT16 = 3
    
    'Use 16-bit IEEE 754 half-precision floating point values
    JXL_TYPE_FLOAT16 = 5
    
End Enum

#If False Then
    Private Const JXL_TYPE_FLOAT = 0, JXL_TYPE_UINT8 = 2, JXL_TYPE_UINT16 = 3, JXL_TYPE_FLOAT16 = 5
#End If

'Ordering of multi-byte data.
Private Enum JxlEndianness
  
    'Use the endianness of the system, either little endian or big endian, without forcing either
    ' specific endianness. Do not use if pixel data should be exported to a well defined format.
    JXL_NATIVE_ENDIAN = 0
    
    'Force little endian
    JXL_LITTLE_ENDIAN = 1
    
    'Force big endian
    JXL_BIG_ENDIAN = 2
    
End Enum

#If False Then
    Private Const JXL_NATIVE_ENDIAN = 0, JXL_LITTLE_ENDIAN = 1, JXL_BIG_ENDIAN = 2
#End If

'Data type for the sample values per channel per pixel for the output buffer for pixels.
' This is not necessarily the same as the data type encoded in the codestream.
' The channels are interleaved per pixel.
' The pixels are organized row by row, left to right, top to bottom.
Private Type JxlPixelFormat

    'Amount of channels available in a pixel buffer.
    ' 1: single-channel data, e.g. grayscale or a single extra channel
    ' 2: single-channel + alpha
    ' 3: trichromatic, e.g. RGB
    ' 4: trichromatic + alpha
    ' TODO: this needs finetuning. It is not yet defined how the user chooses output color space. CMYK+alpha needs 5 channels.
    num_channels As Long
    
    'Data type of each channel.
    data_type As JxlDataType
    
    'Whether multi-byte data types are represented in big endian or little endian format.
    ' This applies to JXL_TYPE_UINT16, JXL_TYPE_UINT32 and JXL_TYPE_FLOAT.
    endianness As JxlEndianness
    
    'Align scanlines to a multiple of align bytes, or 0 to require no alignment at all
    ' (which has the same effect as value 1).
    align_scanline As Long
    
End Type

'Current full-image header, if any
Private m_Header As JxlBasicInfo

'Library handle will be non-zero if libjxl is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE
Private m_LibHandle As Long, m_LibAvailable As Boolean

'libjxl has very specific compiler needs in order to produce maximum perf code, so rather than
' compile myself, I stick with the prebuilt Windows binaries and wrap 'em using DispCallFunc
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum LibJXL_ProcAddress
    JxlDecoderCloseInput
    JxlDecoderCreate
    JxlDecoderDCOutBufferSize
    JxlDecoderDestroy
    JxlDecoderExtraChannelBufferSize
    JxlDecoderFlushImage
    JxlDecoderGetBasicInfo
    JxlDecoderGetBoxSizeRaw
    JxlDecoderGetBoxType
    JxlDecoderGetColorAsICCProfile
    JxlDecoderGetExtraChannelInfo
    JxlDecoderGetExtraChannelName
    JxlDecoderGetFrameHeader
    JxlDecoderGetFrameName
    JxlDecoderGetICCProfileSize
    JxlDecoderImageOutBufferSize
    JxlDecoderPreviewOutBufferSize
    JxlDecoderProcessInput
    JxlDecoderReleaseBoxBuffer
    JxlDecoderReleaseInput
    JxlDecoderReleaseJPEGBuffer
    JxlDecoderReset
    JxlDecoderSetBoxBuffer
    JxlDecoderSetDCOutBuffer
    JxlDecoderSetDecompressBoxes
    JxlDecoderSetExtraChannelBuffer
    JxlDecoderSetImageOutBuffer
    JxlDecoderSetInput
    JxlDecoderSetJPEGBuffer
    JxlDecoderSetPreviewOutBuffer
    JxlDecoderSizeHintBasicInfo
    JxlDecoderSubscribeEvents
    JxlDecoderVersion
    JxlSignatureCheck
    
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Current JXL decoder, if any.  Created once-per-image, and released when the load process terminates
' (either successfully or unsuccessfully).
Private m_JxlDecoder As Long

'Current file stream manager, if any.  Created once-per-image and released when the load process terminates.
Private m_Stream As pdStream

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to the maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

'Initialize the library.  Do not call this until you have verified its existence (typically via the PluginManager module)
Public Function InitializeLibJXL(ByRef pathToDLLFolder As String) As Boolean
    
    InitializeLibJXL = False
    
    'I don't currently know how to build libxjl in an XP-compatible way.
    ' As a result, its support is limited to Win Vista and above.
    If (Not OS.IsVistaOrLater) Then
        debugMsg "libjxl does not currently work on Windows XP."
        Exit Function
    End If
    
    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim libPath As String
    libPath = pathToDLLFolder & "libjxl.dll"
    m_LibHandle = VBHacks.LoadLib(libPath)
    InitializeLibJXL = (m_LibHandle <> 0)
    m_LibAvailable = InitializeLibJXL
    
    'Initialize all module-level arrays
    ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
    ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
    ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
    
    'If we initialized the library successfully, cache some library-specific data
    If InitializeLibJXL Then
        
        'Pre-load all relevant proc addresses
        m_ProcAddresses(JxlDecoderVersion) = GetProcAddress(m_LibHandle, "JxlDecoderVersion")
        m_ProcAddresses(JxlSignatureCheck) = GetProcAddress(m_LibHandle, "JxlSignatureCheck")
        m_ProcAddresses(JxlDecoderCreate) = GetProcAddress(m_LibHandle, "JxlDecoderCreate")
        m_ProcAddresses(JxlDecoderDestroy) = GetProcAddress(m_LibHandle, "JxlDecoderDestroy")
        m_ProcAddresses(JxlDecoderReset) = GetProcAddress(m_LibHandle, "JxlDecoderReset")
        m_ProcAddresses(JxlDecoderCloseInput) = GetProcAddress(m_LibHandle, "JxlDecoderCloseInput")
        m_ProcAddresses(JxlDecoderDCOutBufferSize) = GetProcAddress(m_LibHandle, "JxlDecoderDCOutBufferSize")
        m_ProcAddresses(JxlDecoderExtraChannelBufferSize) = GetProcAddress(m_LibHandle, "JxlDecoderExtraChannelBufferSize")
        m_ProcAddresses(JxlDecoderFlushImage) = GetProcAddress(m_LibHandle, "JxlDecoderFlushImage")
        m_ProcAddresses(JxlDecoderGetBasicInfo) = GetProcAddress(m_LibHandle, "JxlDecoderGetBasicInfo")
        m_ProcAddresses(JxlDecoderGetBoxSizeRaw) = GetProcAddress(m_LibHandle, "JxlDecoderGetBoxSizeRaw")
        m_ProcAddresses(JxlDecoderGetBoxType) = GetProcAddress(m_LibHandle, "JxlDecoderGetBoxType")
        m_ProcAddresses(JxlDecoderGetColorAsICCProfile) = GetProcAddress(m_LibHandle, "JxlDecoderGetColorAsICCProfile")
        m_ProcAddresses(JxlDecoderGetExtraChannelInfo) = GetProcAddress(m_LibHandle, "JxlDecoderGetExtraChannelInfo")
        m_ProcAddresses(JxlDecoderGetExtraChannelName) = GetProcAddress(m_LibHandle, "JxlDecoderGetExtraChannelName")
        m_ProcAddresses(JxlDecoderGetFrameHeader) = GetProcAddress(m_LibHandle, "JxlDecoderGetFrameHeader")
        m_ProcAddresses(JxlDecoderGetFrameName) = GetProcAddress(m_LibHandle, "JxlDecoderGetFrameName")
        m_ProcAddresses(JxlDecoderGetICCProfileSize) = GetProcAddress(m_LibHandle, "JxlDecoderGetICCProfileSize")
        m_ProcAddresses(JxlDecoderImageOutBufferSize) = GetProcAddress(m_LibHandle, "JxlDecoderImageOutBufferSize")
        m_ProcAddresses(JxlDecoderPreviewOutBufferSize) = GetProcAddress(m_LibHandle, "JxlDecoderPreviewOutBufferSize")
        m_ProcAddresses(JxlDecoderProcessInput) = GetProcAddress(m_LibHandle, "JxlDecoderProcessInput")
        m_ProcAddresses(JxlDecoderReleaseBoxBuffer) = GetProcAddress(m_LibHandle, "JxlDecoderReleaseBoxBuffer")
        m_ProcAddresses(JxlDecoderReleaseInput) = GetProcAddress(m_LibHandle, "JxlDecoderReleaseInput")
        m_ProcAddresses(JxlDecoderReleaseJPEGBuffer) = GetProcAddress(m_LibHandle, "JxlDecoderReleaseJPEGBuffer")
        m_ProcAddresses(JxlDecoderSetBoxBuffer) = GetProcAddress(m_LibHandle, "JxlDecoderSetBoxBuffer")
        m_ProcAddresses(JxlDecoderSetDCOutBuffer) = GetProcAddress(m_LibHandle, "JxlDecoderSetDCOutBuffer")
        m_ProcAddresses(JxlDecoderSetDecompressBoxes) = GetProcAddress(m_LibHandle, "JxlDecoderSetDecompressBoxes")
        m_ProcAddresses(JxlDecoderSetExtraChannelBuffer) = GetProcAddress(m_LibHandle, "JxlDecoderSetExtraChannelBuffer")
        m_ProcAddresses(JxlDecoderSetImageOutBuffer) = GetProcAddress(m_LibHandle, "JxlDecoderSetImageOutBuffer")
        m_ProcAddresses(JxlDecoderSetInput) = GetProcAddress(m_LibHandle, "JxlDecoderSetInput")
        m_ProcAddresses(JxlDecoderSetJPEGBuffer) = GetProcAddress(m_LibHandle, "JxlDecoderSetJPEGBuffer")
        m_ProcAddresses(JxlDecoderSetPreviewOutBuffer) = GetProcAddress(m_LibHandle, "JxlDecoderSetPreviewOutBuffer")
        m_ProcAddresses(JxlDecoderSizeHintBasicInfo) = GetProcAddress(m_LibHandle, "JxlDecoderSizeHintBasicInfo")
        m_ProcAddresses(JxlDecoderSubscribeEvents) = GetProcAddress(m_LibHandle, "JxlDecoderSubscribeEvents")
        
    Else
        debugMsg "WARNING!  LoadLibrary failed to load libjxl.  Last DLL error: " & Err.LastDllError
        debugMsg "(FYI, the attempted path was: " & libPath & ")"
    End If
    
End Function

'Forcibly disable library interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetLibJXLVersion() As String
    
    'Do not attempt to retrieve version info unless the library was loaded successfully.
    If (m_LibHandle <> 0) And m_LibAvailable Then
        
        Dim ptrVersion As Long
        ptrVersion = CallCDeclW(JxlDecoderVersion, vbLong)
        
        'From the docs (https://libjxl.readthedocs.io/en/latest/api_decoder.html):
        ' Returns the decoder library version as an integer:
        ' MAJOR_VERSION * 1000000 + MINOR_VERSION * 1000 + PATCH_VERSION.
        ' (For example, version 1.2.3 would return 1002003.)
        GetLibJXLVersion = Trim$(Str$(ptrVersion \ 1000000)) & "." & Trim$(Str$((ptrVersion \ 1000) Mod 1000)) & "." & Trim$(Str$(ptrVersion Mod 1000)) & ".0"
        
    End If
        
End Function

Public Function IsLibJXLAvailable() As Boolean
    IsLibJXLAvailable = (m_LibHandle <> 0)
End Function

Public Function IsLibJXLEnabled() As Boolean
    IsLibJXLEnabled = m_LibAvailable
End Function

'When PD closes, make sure to release our open library handle
Public Sub ReleaseLibJXL()
    If (m_LibHandle <> 0) Then
        VBHacks.FreeLib m_LibHandle
        m_LibHandle = 0
    End If
End Sub

'Import/Export functions follow
Public Function IsFileJXL(ByRef srcFile As String) As Boolean
    
    IsFileJXL = False
    
    'Failsafe check
    If (Not Plugin_jxl.IsLibJXLEnabled()) Then Exit Function
    
    'libjxl provides a built-in validation function, *but* we need to manually pull some bytes into memory first
    Dim tmpStream As pdStream
    Set tmpStream = New pdStream
    If tmpStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
        
        'The spec does not suggest how many bytes need to be read before validation can occur.
        ' This arbitrary number is hopefully big enough (but if it isn't, we'll iterate accordingly).
        Const NUM_BYTES_TO_TEST As Long = 1024
        
        Dim numBytesAvailable As Long
        numBytesAvailable = NUM_BYTES_TO_TEST
        
        'Use file size as an upper limit (simple failsafe against reading past EOF)
        Dim sizeOfWholeFile As Long
        sizeOfWholeFile = Files.FileLenW(srcFile)
        If (numBytesAvailable > sizeOfWholeFile) Then numBytesAvailable = sizeOfWholeFile
        
        'Pull [numBytesAvailable] into memory
        Dim ptrRawBytes As Long
        ptrRawBytes = tmpStream.ReadBytes_PointerOnly(numBytesAvailable)
        
        'Attempt JXL validation
        Dim retSignature As JxlSignature
        retSignature = CallCDeclW(JxlSignatureCheck, vbLong, ptrRawBytes, numBytesAvailable)
        
        'Repeat with ever-larger chunks of the file if the decoder requires it
        If (retSignature = JXL_SIG_NOT_ENOUGH_BYTES) Then
            
            Do
                
                'Calculate how many new bytes to pull in
                numBytesAvailable = numBytesAvailable * 2
                If (numBytesAvailable > sizeOfWholeFile) Then numBytesAvailable = sizeOfWholeFile
                
                'Reset stream pointer and pull in the new, larger [numBytesAvailable] for validation
                tmpStream.SetPosition 0, FILE_BEGIN
                ptrRawBytes = tmpStream.ReadBytes_PointerOnly(numBytesAvailable)
                
                'Continue validating until EOF is reached, if necessary
                retSignature = CallCDeclW(JxlSignatureCheck, vbLong, ptrRawBytes, numBytesAvailable)
                
            Loop While (retSignature = JXL_SIG_NOT_ENOUGH_BYTES) And (numBytesAvailable < sizeOfWholeFile)
            
        End If
        
        IsFileJXL = (retSignature = JXL_SIG_CONTAINER) Or (retSignature = JXL_SIG_CODESTREAM)
        
        'Release the stream regardless of success/failure; we'll re-create it as necessary in a later step
        tmpStream.StopStream True
        
    End If
    
End Function

'Load a JPEG XL file from disk.  srcFile must be a fully qualified path.  In the case of animated images,
' dstImage will be populated with all embedded frames, one frame per layer.
Public Function LoadJXL(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "LoadJXL"
    LoadJXL = False
    
    'Failsafe check
    If (Not Plugin_jxl.IsLibJXLEnabled()) Then Exit Function
    
    'Next, we need to validate the file format as JPEG-XL.
    If Plugin_jxl.IsFileJXL(srcFile) Then
        
        If JXL_DEBUG_VERBOSE Then debugMsg "JXL format found.  Proceeding with load..."
        If (dstImage Is Nothing) Then Set dstImage = New pdImage
        
        'Create a generic JXL decoder.  This (opaque struct) must be kept alive for the duration
        ' of the load process.
        '
        'Note that to improve performance, we will simply reset an existing decoder if one exists
        ' (rather than create a new one).  This approach obviously makes this implementation non-thread-safe
        ' - but we could obviously delegate this to individual classes for a thread-safe solution in the future.
        JXL_CreateDecoder
        If (m_JxlDecoder = 0) Then
            InternalError FUNC_NAME, "can't continue without a JxlDecoder instance; abandoning import"
            Exit Function
        End If
        
        'We can now start feeding data into the decoder.  libjxl uses an interesting design where the caller
        ' can "subscribe" to "events".  These events are just special return codes for the "feed more data into
        ' the decoder" function, but they are convenient for parsing because we can simply wait for the "events"
        ' to occur before doing hefty tasks like allocating buffers for pixels, etc.
        Dim eventsWanted As JxlDecoderStatus
        eventsWanted = JXL_DEC_BASIC_INFO Or JXL_DEC_FRAME Or JXL_DEC_FULL_IMAGE
        
        Dim jxlReturn As JxlDecoderStatus
        jxlReturn = CallCDeclW(JxlDecoderSubscribeEvents, vbLong, m_JxlDecoder, eventsWanted)
        If (jxlReturn = JXL_DEC_ERROR) Then
            InternalError FUNC_NAME, "couldn't subscribe events"
            Exit Function
        ElseIf (jxlReturn <> JXL_DEC_SUCCESS) Then
            InternalError FUNC_NAME, "unexpected return: " & jxlReturn
            Exit Function
        Else
            If JXL_DEBUG_VERBOSE Then debugMsg "Successfully subscribed to events: " & eventsWanted
        End If
        
        '(more events are a potential future TODO)
        
        'Open a stream on the underlying JXL file
        Const JXL_CHUNK_SIZE As Long = 2 ^ 19   '0.5 MB at a time seems like a reasonable modern default?
        Set m_Stream = New pdStream
        If (Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, JXL_CHUNK_SIZE, optimizeAccess:=OptimizeSequentialAccess)) Then
            InternalError FUNC_NAME, "no stream", True
            Exit Function
        End If
        
        If JXL_DEBUG_VERBOSE Then debugMsg "Started generic file stream on " & srcFile
        
        'Start feeding data into libjxl.  For now, we're just gonna dump as much of the file into libjxl as we can
        ' until we receive the event(s) we've subscribed to.
        Dim nextEvent As JxlDecoderStatus
        nextEvent = JXL_ProcessUntilEvent(JXL_CHUNK_SIZE)
        If (nextEvent = JXL_DEC_ERROR) Then GoTo LoadFailed
        
        'The first event we expect is "basic image header" retrieval.
        If (nextEvent = JXL_DEC_BASIC_INFO) Then
            
            jxlReturn = CallCDeclW(JxlDecoderGetBasicInfo, vbLong, m_JxlDecoder, VarPtr(m_Header))
            If (jxlReturn <> JXL_DEC_SUCCESS) Then
                InternalError FUNC_NAME, "couldn't get basic info"
                GoTo LoadFailed
            End If
            
        Else
            InternalError FUNC_NAME, "unexpected event instead of basic info"
        End If
        
        'If we're still here, basic image info retrieval was successful.
        If JXL_DEBUG_VERBOSE Then
            debugMsg "Image dimensions: " & m_Header.xsize & "x" & m_Header.ysize
            debugMsg "Image channel bit-depth: " & m_Header.bits_per_sample
            debugMsg "Image num color channels: " & m_Header.num_color_channels
            debugMsg "Image num extra channels: " & m_Header.num_extra_channels
            debugMsg "Image is animated: " & (m_Header.have_animation <> 0)
        End If
        
        'Validate the image header before continuing.
        If (m_Header.xsize <= 0) Or (m_Header.ysize <= 0) Then
            InternalError FUNC_NAME, "bad image size"
            GoTo LoadFailed
        End If
        
        'We now have enough to initialize a basic pdImage object.
        If (dstImage Is Nothing) Then Set dstImage = New pdImage
        dstImage.Width = m_Header.xsize
        dstImage.Height = m_Header.ysize
        
        'DPI is not encoded in JXL files, but ExifTool will try to pick it up later during processing if
        ' it encounters it...
        'dstImage.SetDPI 72, 72
        
        'Formal animation support remains TODO pending images to test with!
        Dim imgIsAnimated As Boolean
        imgIsAnimated = (m_Header.have_animation <> 0)
        
        'dstImage.ImgStorage.AddEntry "animation-loop-count", Trim$(Str$(m_AnimationInfo.loop_count))
        Dim idxFrame As Long, numFramesOK As Long
        idxFrame = 0
        numFramesOK = 0
        
        'We also need to flag the underlying format in advance, since it changes the way layer
        ' names are assigned (animation layers are called "frames" instead of "pages")
        dstImage.SetOriginalFileFormat PDIF_JXL
        
        'Animated images will be auto-loaded as separate layers
        Dim tmpLayer As pdLayer, tmpDIB As pdDIB
        Dim curFrameHeader As JxlFrameHeader
        
        'We will now continue iterating through the file, one frame at a time, until the full image
        ' is loaded.  (Note that JPEG XL files don't give the number of frames up front, which is a pain -
        ' so we have no choice but to iterate until we hit the frame marked as "last frame".)
        '
        'Note that we will also quit if we successfully load the frame marked as "last frame".
        Dim letsQuitEarly As Boolean
        letsQuitEarly = False
        
        nextEvent = JXL_ProcessUntilEvent(JXL_CHUNK_SIZE)
        Do While (nextEvent <> JXL_DEC_SUCCESS)
        
            'For now, halt on all errors.  (They are assumed to be unrecoverable at this point,
            ' since libjxl's API provides no mechanism for retrieving *what* the error was -
            ' that's a TODO item per their docs.)
            If (nextEvent = JXL_DEC_ERROR) Then GoTo LoadFailed
            
            'Handle events according to type.
            
            'A new frame header has been encountered.  Prep a pdLayer object to receive it.
            If (nextEvent = JXL_DEC_FRAME) Then
                
                'Yes, this text uses "page" instead of "frame" - this is purely to reduce localization burdens
                If imgIsAnimated Then
                    Dim unknownText As String
                    If (LenB(unknownText) = 0) Then unknownText = g_Language.TranslateMessage("unknown")
                    Message "Loading page %1 of %2...", CStr(idxFrame + 1), unknownText, "DONOTLOG"
                End If
                
                'Retrieve the current frame header
                jxlReturn = CallCDeclW(JxlDecoderGetFrameHeader, vbLong, m_JxlDecoder, VarPtr(curFrameHeader))
                If (jxlReturn = JXL_DEC_SUCCESS) Then
                
                    'Success!  Create a new layer in the destination image, then copy the pixel data and
                    ' timestamp (if relevant) into it.
                    Dim newLayerID As Long, newLayerName As String
                    newLayerID = dstImage.CreateBlankLayer()
                    Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                    
                    If imgIsAnimated Then
                        newLayerName = Layers.GenerateInitialLayerName(vbNullString, vbNullString, True, dstImage, dstDIB, idxFrame)
                    Else
                        newLayerName = Layers.GenerateInitialLayerName(srcFile, vbNullString, False, dstImage, dstDIB)
                    End If
                    
                    tmpLayer.InitializeNewLayer PDL_Image, newLayerName, Nothing, True
                    tmpLayer.SetLayerVisibility (idxFrame = 0)
                    
                    'Optional layer name from embedded frame name?
                    If (curFrameHeader.name_length <> 0) Then
                        'TODO
                    End If
                    
                    'TODO: pull animation timing info if relevant
                    
'                    'As part of storing frametime, update the layer's name with ([time] ms) at the end
'                    frameTimeInMS = frameTimestamp - lastFrameTimestamp
'                    tmpLayer.SetLayerFrameTimeInMS frameTimeInMS
'                    tmpLayer.SetLayerName tmpLayer.GetLayerName & " (" & CStr(frameTimeInMS) & "ms)"
'                    lastFrameTimestamp = frameTimestamp
'                    tmpLayer.NotifyOfDestructiveChanges
                    
                    'Prep a (reusable) buffer to receive this frame's pixel data
                    If (tmpDIB Is Nothing) Then Set tmpDIB = New pdDIB
                    If (tmpDIB.GetDIBWidth <> m_Header.xsize) Or (tmpDIB.GetDIBHeight <> m_Header.ysize) Then
                        tmpDIB.CreateBlank m_Header.xsize, m_Header.ysize, 32, 0, 0
                    Else
                        tmpDIB.ResetDIB 0
                    End If
                    
                Else
                    InternalError FUNC_NAME, "bad frame header"
                    GoTo LoadFailed
                End If
            
            'Pixel data is ready, but we need to specify a buffer first
            ElseIf (nextEvent = JXL_DEC_NEED_IMAGE_OUT_BUFFER) Then
                
                Dim pxFormat As JxlPixelFormat
                With pxFormat
                    .align_scanline = 4     'Windows requires 4-byte alignment, but this is redundant when decoding to RGBA8...
                    .data_type = JXL_TYPE_UINT8
                    .num_channels = 4

                    'Only matters if we support higher bit-depths in the future, obviously
                    .endianness = JXL_LITTLE_ENDIAN
                End With

                'Ensure our placeholder DIB is valid
                Dim dibReady As Boolean
                dibReady = Not (tmpDIB Is Nothing)
                If dibReady Then dibReady = (tmpDIB.GetDIBWidth = m_Header.xsize) And (tmpDIB.GetDIBHeight = m_Header.ysize)
                
                'Test only: see what size is required by the decoder?
                Dim reqSize As Long
                jxlReturn = CallCDeclW(JxlDecoderImageOutBufferSize, vbLong, m_JxlDecoder, VarPtr(pxFormat), VarPtr(reqSize))
                debugMsg "Size allocated for frame: " & Files.GetFormattedFileSize(tmpDIB.GetDIBStride * tmpDIB.GetDIBHeight)
                
                'If a valid buffer exists, pass its information to the decoder
                If dibReady Then
                    jxlReturn = CallCDeclW(JxlDecoderSetImageOutBuffer, vbLong, m_JxlDecoder, VarPtr(pxFormat), tmpDIB.GetDIBPointer, tmpDIB.GetDIBStride * tmpDIB.GetDIBHeight)
                    If (jxlReturn <> JXL_DEC_SUCCESS) Then
                        InternalError FUNC_NAME, "bad SetImageOutBuffer"
                        GoTo LoadFailed
                    End If
                Else
                    InternalError FUNC_NAME, "bad pixel buffer"
                    GoTo LoadFailed
                End If

            'The current frame has been decoded successfully.
            ElseIf (nextEvent = JXL_DEC_FULL_IMAGE) Then
                
                'Premultiply the DIB (TODO: see if the decoder can do this for us? might be faster)
                tmpDIB.SetAlphaPremultiplication True
                
                'Swizzle R/B channels
                DIBs.SwizzleBR tmpDIB
                
                'Store the finished DIB inside the temporary layer, then assign that layer to the parent image
                tmpLayer.SetLayerDIB tmpDIB
                Set tmpDIB = New pdDIB
                
                'Increment frame count and reset current frame state.
                idxFrame = idxFrame + 1
                numFramesOK = numFramesOK + 1
                If JXL_DEBUG_VERBOSE Then debugMsg "Successfully finished frame #" & idxFrame
                
                'If this frame was marked as the last frame in the image, do not attempt to retrieve
                ' another frame - instead, just exit after this.
                If (curFrameHeader.bool_is_last <> 0) Then letsQuitEarly = True
            
            Else
                InternalError FUNC_NAME, "unexpected event: " & nextEvent
                Exit Do
            End If
            
            'Retrieve the next event.  Note that libjxl will return JXL_DEC_SUCCESS if all requested events
            ' have been returned, *even if EOF has not been reached*.  That's okay for our purposes - but if we
            ' expand coverage in the future, we need to manually request new events accordingly.
            If letsQuitEarly Then
                Exit Do
            Else
                nextEvent = JXL_ProcessUntilEvent(JXL_CHUNK_SIZE)
            End If
        
        Loop
        
        'Report success if at least one frame was retrieved correctly
        LoadJXL = (numFramesOK > 0)
        If LoadJXL Then
            dstImage.NotifyImageChanged UNDO_Everything
            If JXL_DEBUG_VERBOSE Then debugMsg "JXL loaded successfully; " & numFramesOK & " frames processed."
        End If
        
    Else
        Exit Function
    End If
    
    If (Not LoadJXL) And JXL_DEBUG_VERBOSE Then debugMsg "Plugin_jxl.LoadJXL failed."
    Exit Function
    
LoadFailed:
    
    LoadJXL = False
    InternalError FUNC_NAME, "terminating due to error"
    
    'Free the decoder, if any
    JXL_DestroyDecoder
    
    'Free our file importer too
    Set m_Stream = Nothing
    
End Function

'Create a new JPEG XL decoder (fills m_JxlDecoder with a pointer to an opaque JxlDecoder struct)
Private Function JXL_CreateDecoder() As Boolean
    
    Const FUNC_NAME = "JXL_CreateDecoder"
    
    If (m_JxlDecoder = 0) Then
        m_JxlDecoder = CallCDeclW(JxlDecoderCreate, vbLong, 0&)
        If (m_JxlDecoder = 0) Then
            InternalError FUNC_NAME, "couldn't create decoder"
            Exit Function
        Else
            If JXL_DEBUG_VERBOSE Then debugMsg "Created decoder: " & m_JxlDecoder
        End If
    Else
        JXL_CreateDecoder = JXL_ResetDecoder()
    End If
    
    JXL_CreateDecoder = (m_JxlDecoder <> 0)
    
End Function

'Destroy the current JPEG XL decoder (m_JxlDecoder)
Private Function JXL_DestroyDecoder() As Boolean
    If (m_JxlDecoder <> 0) Then
        CallCDeclW JxlDecoderDestroy, vbEmpty, m_JxlDecoder
        If JXL_DEBUG_VERBOSE Then debugMsg "Destroyed decoder: " & m_JxlDecoder
        m_JxlDecoder = 0
    End If
    JXL_DestroyDecoder = (m_JxlDecoder = 0)
End Function

'Continue loading data into the active decoder until an event state is returned.  This function automatically
' tracks underlying file position to ensure correct read behavior.  File is read in [chunkSize] chunks using
' memory mapping.
Private Function JXL_ProcessUntilEvent(Optional ByVal chunkSize As Long = 1024) As JxlDecoderStatus
        
    If JXL_DEBUG_VERBOSE Then debugMsg "Starting ProcessUntilEvent..."
        
    Const FUNC_NAME As String = "JXL_ProcessUntilEvent"
    
    'Note the current file pointer
    Dim origPosition As Long
    origPosition = m_Stream.GetPosition()
    
    'Ensure [chunkSize] does not extend past EOF
    Dim overflowCheck As Long
    overflowCheck = origPosition + chunkSize
    If (overflowCheck > m_Stream.GetStreamSize) Then chunkSize = m_Stream.GetStreamSize - origPosition
    If (chunkSize < 0) Then
        InternalError FUNC_NAME, "read past EOF"
        JXL_ProcessUntilEvent = JXL_DEC_ERROR
        Exit Function
    End If
    
    'Map [chunkSize] bytes into memory
    Dim ptrToSource As Long
    ptrToSource = m_Stream.ReadBytes_PointerOnly(chunkSize)
    
    'Validate the number of bytes read, just in case our attempted read extended past the end of the file.
    Dim numBytesRead As Long
    numBytesRead = m_Stream.GetPosition() - origPosition
    
    'For extreme details on the read process (so you can see right down to the byte where a file passes/fails),
    ' you can uncomment these lines...
    'If JXL_DEBUG_VERBOSE Then debugMsg "Read " & numBytesRead & " bytes into memory, handing off to libjxl..."
    'If JXL_DEBUG_VERBOSE Then debugMsg "Asking for " & numBytesRead & " bytes from libjxl, file offsets " & origPosition & " to " & (m_Stream.GetPosition - 1) & ", ptr: " & ptrToSource & "..."
    
    'Notify the decoder of new input.  (This step is pass/fail.)
    Dim jxlReturn As JxlDecoderStatus
    jxlReturn = CallCDeclW(JxlDecoderSetInput, vbLong, m_JxlDecoder, ptrToSource, numBytesRead)
    If (jxlReturn <> JXL_DEC_SUCCESS) Then
        InternalError FUNC_NAME, "bad JxlDecoderSetInput"
        Exit Function
    End If
    
    If JXL_DEBUG_VERBOSE Then debugMsg "libjxl SetInput successful."
    
    'Ask the decoder to process the input we've sent.
    jxlReturn = CallCDeclW(JxlDecoderProcessInput, vbLong, m_JxlDecoder)
    If JXL_DEBUG_VERBOSE Then debugMsg "libjxl ProcessInput returned: " & jxlReturn
    
    'If the decoder requires more input before raising a requested event, pass it more input
    Dim numBytesStillRequired As Long
    Do While (jxlReturn = JXL_DEC_NEED_MORE_INPUT)
        
        If JXL_DEBUG_VERBOSE Then debugMsg "libjxl ProcessInput needs more output.  Loading another chunk..."
        
        'Before adding new input, we must release the current input.
        numBytesStillRequired = CallCDeclW(JxlDecoderReleaseInput, vbLong, m_JxlDecoder)
        'If JXL_DEBUG_VERBOSE And (numBytesStillRequired <> 0) Then debugMsg "numBytesStillRequired: " & numBytesStillRequired
        
        'The decoder may align bytes in its own way.  It has returned the number of bytes from the
        ' *last* set we passed it that it *still* needs access to.  We can pass as many bytes as we
        ' want on our next call, but we need to make sure the bytes *start* at the place requested
        ' by the release input call.
        If (numBytesStillRequired <> 0) Then m_Stream.SetPosition numBytesStillRequired * -1, FILE_CURRENT
        
        'Pull more data from file, and once again note the actual number of bytes read.
        origPosition = m_Stream.GetPosition
        
        'Ensure [chunkSize] does not extend past EOF
        overflowCheck = origPosition + chunkSize
        If (overflowCheck > m_Stream.GetStreamSize) Then chunkSize = m_Stream.GetStreamSize - origPosition
        If (chunkSize < 0) Then
            InternalError FUNC_NAME, "read past EOF"
            JXL_ProcessUntilEvent = JXL_DEC_ERROR
            Exit Function
        End If
        
        ptrToSource = m_Stream.ReadBytes_PointerOnly(chunkSize)
        numBytesRead = m_Stream.GetPosition() - origPosition
        
        'For extreme details on the read process (so you can see right down to the byte where a file passes/fails),
        ' you can uncomment this line as well...
        'If JXL_DEBUG_VERBOSE Then debugMsg "(inner) Asking for " & numBytesRead & " bytes from libjxl, file offsets " & origPosition & " to " & (m_Stream.GetPosition - 1) & ", ptr: " & ptrToSource & "..."
        
        'Set the new input (only pass/fail is returned; fail may occur if we didn't release previous input)
        jxlReturn = CallCDeclW(JxlDecoderSetInput, vbLong, m_JxlDecoder, ptrToSource, numBytesRead)
        If (jxlReturn <> JXL_DEC_SUCCESS) Then
            InternalError FUNC_NAME, "bad JxlDecoderSetInput"
            Exit Function
        End If
        
        'Request a new round of processing
        jxlReturn = CallCDeclW(JxlDecoderProcessInput, vbLong, m_JxlDecoder)
        
    Loop
    
    'We are now guaranteed to have raised *some* kind of event (or error).
    JXL_ProcessUntilEvent = jxlReturn
    If JXL_DEBUG_VERBOSE Then debugMsg "libjxl Event ready: " & jxlReturn
    
    'Release any input we have supplied, but be sure to *align the underlying file stream pointer* accordingly
    numBytesStillRequired = CallCDeclW(JxlDecoderReleaseInput, vbLong, m_JxlDecoder)
    If (numBytesStillRequired > 0) Then
        If JXL_DEBUG_VERBOSE Then debugMsg "Release input required modifying file offset by -" & numBytesStillRequired
        m_Stream.SetPosition numBytesStillRequired * -1, FILE_CURRENT
    End If

End Function

'Reset the current JPEG XL decoder (m_JxlDecoder).  Frees any image-specific information already inside the decoder,
' so *do not call* unless you have everything you need from the current m_JxlDecoder instance.
Private Function JXL_ResetDecoder() As Boolean
    
    If (m_JxlDecoder <> 0) Then
        CallCDeclW JxlDecoderReset, vbEmpty, m_JxlDecoder
        If JXL_DEBUG_VERBOSE Then debugMsg "Reset decoder: " & m_JxlDecoder
    
    'Failsafe only; just call JXL_CreateDecoder if you need a new instance
    Else
        JXL_ResetDecoder = JXL_CreateDecoder()
    End If
    
    JXL_ResetDecoder = (m_JxlDecoder <> 0)
    
End Function

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As LibJXL_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

    Dim i As Long, vTemp() As Variant, hResult As Long
    
    Dim numParams As Long
    If (UBound(pa) < LBound(pa)) Then numParams = 0 Else numParams = UBound(pa) + 1
    
    If IsMissing(pa) Then
        ReDim vTemp(0) As Variant
    Else
        vTemp = pa 'make a copy of the params to prevent problems with VT_ByRef members in the ParamArray
    End If
    
    For i = 0 To numParams - 1
        m_vType(i) = VarType(vTemp(i))
        m_vPtr(i) = VarPtr(vTemp(i))
    Next i
    
    Const CC_CDECL As Long = 1
    hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    
End Function

'The following two functions are for logging errors (always active) and/or informational processing messages
' (only when JXL_DEBUG_VERBOSE = True).
'
' To use these functions outside PhotoDemon, substitute PDDebug.LogAction with your own logger.
Private Sub debugMsg(ByRef msgText As String)
    PDDebug.LogAction msgText, PDM_External_Lib, True
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then debugMsg "Plugin_jxl." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "Plugin_jxl." & funcName & "() reported an error: " & errDescription
    End If
End Sub
