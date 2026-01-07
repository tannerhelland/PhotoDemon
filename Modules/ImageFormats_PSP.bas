Attribute VB_Name = "ImageFormats_PSP"
'***************************************************************************
'PhotoDemon PSP (PaintShop Pro) Container and Parser
'Copyright 2020-2026 by Tanner Helland
'Created: 30/December/20
'Last updated: 03/February/21
'Last update: wrap up finishing touches on export support
'
'This module (and associated pdPSP- classes) handle JASC/Corel Paint Shop Pro image parsing.
' All code has been custom-built for PhotoDemon, with a special emphasis on parsing performance.
'
'Both import and export of PSP files are supported.  I have attempted to support all possible
' versions of PSP files (PSP 5 was the version that "invented" the PSP format, and all versions
' since have made slight modifications to the format).  Unfortunately, Corel stopped publishing
' public specs for the PSP format after PSP 8, so support for all versions beyond that point relies
' on guesswork and heuristics instead of an authoritative "Spec".
'
'As with all 3rd-party PSP engines, Paint Shop Pro has many features that don't have direct analogs
' in PhotoDemon.  Many of these settings are still parsed by PD's PSP engine, but they will not
' "appear" in the final loaded image.  My ongoing goal is to expand support in this class as various
' PSP features are implemented in PD itself.
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

'PSP files contain a *lot* of information.  To aid debugging, you can activate "verbose" output;
' this will dump all kinds of diagnostic information to the debug log.
Public Const PSP_DEBUG_VERBOSE As Boolean = False

'PSP loading is complicated, and a lot of things can go wrong.  Instead of returning binary "success/fail"
' values, we return specific flags; "warnings" may be recoverable and you can still attempt to load the file.
' "Failure" returns are unrecoverable and processing *must* be abandoned.  (As a convenience, you can treat
' the "warning" and "failure" values as flags; specific warning/failure states in each category will share
' the same high flag bit.)
'
'As I get deeper into this class, I may expand this enum to include more detailed states.
Public Enum PD_PSPResult
    psp_Success = &H0
    psp_Warning = &H10
    psp_Failure = &H100
    psp_FileNotPSP = &H1000
End Enum

#If False Then
    Private Const psp_Success = &H0, psp_Warning = &H10, psp_Failure = &H100, psp_FileNotPSP = &H1000
#End If

Public Enum PSPBlockID
    PSP_IMAGE_BLOCK = 0             '// General Image Attributes Block (main)
    PSP_CREATOR_BLOCK               '// Creator Data Block (main)
    PSP_COLOR_BLOCK                 '// Color Palette Block (main and sub)
    PSP_LAYER_START_BLOCK           '// Layer Bank Block (main)
    PSP_LAYER_BLOCK                 '// Layer Block (sub)
    PSP_CHANNEL_BLOCK               '// Channel Block (sub)
    PSP_SELECTION_BLOCK             '// Selection Block (main)
    PSP_ALPHA_BANK_BLOCK            '// Alpha Bank Block (main)
    PSP_ALPHA_CHANNEL_BLOCK         '// Alpha Channel Block (sub)
    PSP_COMPOSITE_IMAGE_BLOCK       '// Composite Image Block (sub)
    PSP_EXTENDED_DATA_BLOCK         '// Extended Data Block (main)
    PSP_TUBE_BLOCK                  '// Picture Tube Data Block (main)
    PSP_ADJUSTMENT_EXTENSION_BLOCK  '// Adjustment Layer Block (sub)
    PSP_VECTOR_EXTENSION_BLOCK      '// Vector Layer Block (sub)
    PSP_SHAPE_BLOCK                 '// Vector Shape Block (sub)
    PSP_PAINTSTYLE_BLOCK            '// Paint Style Block (sub)
    PSP_COMPOSITE_IMAGE_BANK_BLOCK  '// Composite Image Bank (main)
    PSP_COMPOSITE_ATTRIBUTES_BLOCK  '// Composite Image Attr. (sub)
    PSP_JPEG_BLOCK                  '// JPEG Image Block (sub)
    PSP_LINESTYLE_BLOCK             '// Line Style Block (sub)
    PSP_TABLE_BANK_BLOCK            '// Table Bank Block (main)
    PSP_TABLE_BLOCK                 '// Table Block (sub)
    PSP_PAPER_BLOCK                 '// Vector Table Paper Block (sub)
    PSP_PATTERN_BLOCK               '// Vector Table Pattern Block (sub)
    PSP_GRADIENT_BLOCK              '// Vector Table Gradient Block (not used)
    PSP_GROUP_EXTENSION_BLOCK       '// Group Layer Block (sub)
    PSP_MASK_EXTENSION_BLOCK        '// Mask Layer Block (sub)
    PSP_BRUSH_BLOCK                 '// Brush Data Block (main)
    PSP_ART_MEDIA_BLOCK             '// Art Media Layer Block (main)
    PSP_ART_MEDIA_MAP_BLOCK         '// Art Media Layer map data Block (main)
    PSP_ART_MEDIA_TILE_BLOCK        '// Art Media Layer map tile Block (main)
    PSP_ART_MEDIA_TEXTURE_BLOCK     '// AM Layer map texture Block (main)
    PSP_COLORPROFILE_BLOCK          '// ICC Color profile block
End Enum

#If False Then
    Private Const PSP_IMAGE_BLOCK = 0, PSP_CREATOR_BLOCK = 1, PSP_COLOR_BLOCK = 2, PSP_LAYER_START_BLOCK = 3, PSP_LAYER_BLOCK = 4, PSP_CHANNEL_BLOCK = 5, PSP_SELECTION_BLOCK = 6, PSP_ALPHA_BANK_BLOCK = 7, PSP_ALPHA_CHANNEL_BLOCK = 8, PSP_COMPOSITE_IMAGE_BLOCK = 9
    Private Const PSP_EXTENDED_DATA_BLOCK = 10, PSP_TUBE_BLOCK = 11, PSP_ADJUSTMENT_EXTENSION_BLOCK = 12, PSP_VECTOR_EXTENSION_BLOCK = 13, PSP_SHAPE_BLOCK = 14, PSP_PAINTSTYLE_BLOCK = 15, PSP_COMPOSITE_IMAGE_BANK_BLOCK = 16, PSP_COMPOSITE_ATTRIBUTES_BLOCK = 17, PSP_JPEG_BLOCK = 18, PSP_LINESTYLE_BLOCK = 19
    Private Const PSP_TABLE_BANK_BLOCK = 20, PSP_TABLE_BLOCK = 21, PSP_PAPER_BLOCK = 22, PSP_PATTERN_BLOCK = 23, PSP_GRADIENT_BLOCK = 24, PSP_GROUP_EXTENSION_BLOCK = 25, PSP_MASK_EXTENSION_BLOCK = 26, PSP_BRUSH_BLOCK = 27, PSP_ART_MEDIA_BLOCK = 28, PSP_ART_MEDIA_MAP_BLOCK = 29
    Private Const PSP_ART_MEDIA_TILE_BLOCK = 30, PSP_ART_MEDIA_TEXTURE_BLOCK = 31, PSP_COLORPROFILE_BLOCK = 32
#End If

'/* Graphic contents flags.
Public Enum PSPGraphicContents
    '// Layer types
    keGCRasterLayers = &H1                  '// At least one raster layer
    keGCVectorLayers = &H2                  '// At least one vector layer
    keGCAdjustmentLayers = &H4              '// At least one adjustment layer
    keGCGroupLayers = &H8                   '// at least one group layer
    keGCMaskLayers = &H10                   '// at least one mask layer
    keGCArtMediaLayers = &H20               '// at least one art media layer
    '// Additional attributes
    keGCMergedCache = &H800000              '// merged cache (composite image)
    keGCThumbnail = &H1000000               '// Has a thumbnail
    keGCThumbnailTransparency = &H2000000   '// Thumbnail transparency
    keGCComposite = &H4000000               '// Has a composite image
    keGCCompositeTransparency = &H8000000   '// Composite transparency
    keGCFlatImage = &H10000000              '// Just a background
    keGCSelection = &H20000000              '// Has a selection
    keGCFloatingSelectionLayer = &H40000000 '// Has float. selection
    keGCAlphaChannels = &H80000000          '// Has alpha channel(s)
End Enum

#If False Then
    Private Const keGCRasterLayers = &H1, keGCVectorLayers = &H2, keGCAdjustmentLayers = &H4, keGCGroupLayers = &H8, keGCMaskLayers = &H10, keGCArtMediaLayers = &H20
    Private Const keGCMergedCache = &H800000, keGCThumbnail = &H1000000, keGCThumbnailTransparency = &H2000000, keGCComposite = &H4000000, keGCCompositeTransparency = &H8000000, keGCFlatImage = &H10000000, keGCSelection = &H20000000, keGCFloatingSelectionLayer = &H40000000, keGCAlphaChannels = &H80000000
#End If

'/* Possible metrics used to measure resolution.  */
Public Enum PSP_METRIC
    PSP_METRIC_UNDEFINED = 0    '// Metric unknown
    PSP_METRIC_INCH             '// Resolution is in inches
    PSP_METRIC_CM               '// Resolution is in centimeters
End Enum

#If False Then
    Private Const PSP_METRIC_UNDEFINED = 0, PSP_METRIC_INCH = 1, PSP_METRIC_CM = 2
#End If

'/* Channel types
Public Enum PSPChannelType
    PSP_CHANNEL_COMPOSITE = 0   '// Channel of single channel bitmap
    PSP_CHANNEL_RED             '// Red channel of 8, 16 bpc bitmap
    PSP_CHANNEL_GREEN           '// Green channel of 8, 16 bpc bitmap
    PSP_CHANNEL_BLUE            '// Blue channel of 8, 16 bpc bitmap
End Enum

#If False Then
    Private Const PSP_CHANNEL_COMPOSITE = 0, PSP_CHANNEL_RED = 1, PSP_CHANNEL_GREEN = 2, PSP_CHANNEL_BLUE = 3
#End If

'/* Compression types
Public Enum PSPCompression
    PSP_COMP_NONE = 0   '// No compression
    PSP_COMP_RLE        '// RLE compression
    PSP_COMP_LZ77       '// LZ77 compression
    PSP_COMP_JPEG       '// JPEG compression (only used by thumbnail and composite image, invalid in image header)
End Enum

#If False Then
    Private Const PSP_COMP_NONE = 0, PSP_COMP_RLE = 1, PSP_COMP_LZ77 = 2, PSP_COMP_JPEG = 3
#End If

'/* DIB (raster) types
Public Enum PSPDIBType
    PSP_DIB_IMAGE = 0               '// Layer color bitmap
    PSP_DIB_TRANS_MASK              '// Layer transparency mask bitmap
    PSP_DIB_USER_MASK               '// Layer user mask bitmap
    PSP_DIB_SELECTION               '// Selection mask bitmap
    PSP_DIB_ALPHA_MASK              '// Alpha channel mask bitmap
    PSP_DIB_THUMBNAIL               '// Thumbnail bitmap
    PSP_DIB_THUMBNAIL_TRANS_MASK    '// Thumbnail transparency mask
    PSP_DIB_ADJUSTMENT_LAYER        '// Adjustment layer bitmap
    PSP_DIB_COMPOSITE               '// Composite image bitmap
    PSP_DIB_COMPOSITE_TRANS_MASK    '// Composite image transparency
    PSP_DIB_PAPER                   '// Paper bitmap
    PSP_DIB_PATTERN                 '// Pattern bitmap
    PSP_DIB_PATTERN_TRANS_MASK      '// Pattern transparency mask
End Enum

#If False Then
    Private Const PSP_DIB_IMAGE = 0, PSP_DIB_TRANS_MASK = 1, PSP_DIB_USER_MASK = 2, PSP_DIB_SELECTION = 3, PSP_DIB_ALPHA_MASK = 4, PSP_DIB_THUMBNAIL = 5, PSP_DIB_THUMBNAIL_TRANS_MASK = 6, PSP_DIB_ADJUSTMENT_LAYER = 7, PSP_DIB_COMPOSITE = 8, PSP_DIB_COMPOSITE_TRANS_MASK = 9
    Private Const PSP_DIB_PAPER = 10, PSP_DIB_PATTERN = 11, PSP_DIB_PATTERN_TRANS_MASK = 12
#End If

'/* PSP Layer types.  */
Public Enum PSPLayerType
    keGLTUndefined = 0              '// Undefined layer type
    keGLTRaster                     '// Standard raster layer
    keGLTFloatingRasterSelection    '// Floating selection (raster)
    keGLTVector                     '// Vector layer
    keGLTAdjustment                 '// Adjustment layer
    keGLTGroup                      '// Group layer
    keGLTMask                       '// Mask layer
    keGLTArtMedia                   '// Art media layer
End Enum

#If False Then
    Private Const keGLTUndefined = 0, keGLTRaster = 1, keGLTFloatingRasterSelection = 2, keGLTVector = 3, keGLTAdjustment = 4, keGLTGroup = 5, keGLTMask = 6, keGLTArtMedia = 7
#End If

'/* PSP Layer flags.  */
Public Enum PSPLayerProperties
    keVisibleFlag = &H1             '// Layer is visible
    keMaskPresenceFlag = &H2        '// Layer has a mask
End Enum

#If False Then
    Private Const keVisibleFlag = &H1, keMaskPresenceFlag = &H2
#End If

'/* PSP Blend modes.  */
Public Enum PSPBlendModes
    bmLAYER_BLEND_NORMAL = 0
    bmLAYER_BLEND_DARKEN
    bmLAYER_BLEND_LIGHTEN
    bmLAYER_BLEND_LEGACY_HUE
    bmLAYER_BLEND_LEGACY_SATURATION
    bmLAYER_BLEND_LEGACY_COLOR
    bmLAYER_BLEND_LEGACY_LUMINOSITY
    bmLAYER_BLEND_MULTIPLY
    bmLAYER_BLEND_SCREEN
    bmLAYER_BLEND_DISSOLVE
    bmLAYER_BLEND_OVERLAY
    bmLAYER_BLEND_HARD_LIGHT
    bmLAYER_BLEND_SOFT_LIGHT
    bmLAYER_BLEND_DIFFERENCE
    bmLAYER_BLEND_DODGE
    bmLAYER_BLEND_BURN
    bmLAYER_BLEND_EXCLUSION
    bmLAYER_BLEND_TRUE_HUE
    bmLAYER_BLEND_TRUE_SATURATION
    bmLAYER_BLEND_TRUE_COLOR
    bmLAYER_BLEND_TRUE_LIGHTNESS
    bmLAYER_BLEND_ADJUST = 255
End Enum

#If False Then
    Private Const bmLAYER_BLEND_NORMAL = 0, bmLAYER_BLEND_DARKEN = 1, bmLAYER_BLEND_LIGHTEN = 2, bmLAYER_BLEND_LEGACY_HUE = 3, bmLAYER_BLEND_LEGACY_SATURATION = 4, bmLAYER_BLEND_LEGACY_COLOR = 5, bmLAYER_BLEND_LEGACY_LUMINOSITY = 6, bmLAYER_BLEND_MULTIPLY = 7, bmLAYER_BLEND_SCREEN = 8, bmLAYER_BLEND_DISSOLVE = 9
    Private Const bmLAYER_BLEND_OVERLAY = 10, bmLAYER_BLEND_HARD_LIGHT = 11, bmLAYER_BLEND_SOFT_LIGHT = 12, bmLAYER_BLEND_DIFFERENCE = 13, bmLAYER_BLEND_DODGE = 14, bmLAYER_BLEND_BURN = 15, bmLAYER_BLEND_EXCLUSION = 16, bmLAYER_BLEND_TRUE_HUE = 17, bmLAYER_BLEND_TRUE_SATURATION = 18, bmLAYER_BLEND_TRUE_COLOR = 19
    Private Const bmLAYER_BLEND_TRUE_LIGHTNESS = 20, bmLAYER_BLEND_ADJUST = 255
#End If

Public Type PSP_ChannelHeader
    ch_ParentVersionMajor As Long
    ch_ParentWidth As Long
    ch_ParentHeight As Long
    ch_ParentBitDepth As Long
    ch_MaskWidth As Long        'Masks must use *these* measurements instead of the parent measurement
    ch_MaskHeight As Long
    ch_Compression As PSPCompression
    ch_CompressedSize As Long
    ch_UncompressedSize As Long
    ch_dstBitmapType As PSPDIBType
    ch_ChannelType As PSPChannelType
    ch_ChannelOK As Boolean   'Internal value; set to TRUE if decoding appeared to have worked
End Type

'Image header, constructed from the "General Image Attributes" block.
' (Note the similarities to a Windows DIB header, including unused values like "plane count")
Public Type PSPImageHeader

    psph_VersionMajor As Long           'Some parsing behavior needs to be modified based on version, alas
    psph_VersionMinor As Long
    
    psph_HeaderSize As Long             'Used only for double-checking embedded header in file; do NOT use after validation
    psph_Width As Long
    psph_Height As Long
    psph_Resolution As Double           'Interpretation relies on ResolutionUnit
    psph_ResolutionUnit As PSP_METRIC   'Inch or cm
    psph_Compression As PSPCompression  'CANNOT be JPEG-compressed (JPEG is only for the thumbnail)
    psph_BitDepth As Long               'must be 1, 4, 8, 24, or 48
    psph_PlaneCount As Long             'must be 1
    psph_ColorCount As Long             '2 ^ bit-depth
    psph_IsGrayscale As Boolean         '0 = not greyscale, 1 = greyscale, embedded in file as BYTE
    psph_TotalSize As Long              'Sum of the sizes of all layer color bitmaps
    
    'Layers were added in PSP5, which was also the first appearance of a dedicated "PSP" format
    psph_ActiveLayer As Long            'Identifies the layer that was active when the image document was saved
    psph_LayerCount As Long             'Total layer count, embedded in file as WORD
    
    'Added in PSP6
    psph_ContentFlags As PSPGraphicContents 'See enum for flag details
    
    'The spec allows for future header expansion, so the top SIZE member is critical for correct reading;
    ' don't assume a fixed size!
    
End Type

'Given a PSP channel collection and a sample PSP DIB header (from either the parent layer or
' the composite/thumbnail image being rendered), construct a useable pdDIB.
'
'Returns: TRUE if creation appears successful; FALSE otherwise
Public Function PSP_BuildDIBFromChannels(ByVal numPSPChannels As Long, ByRef srcChannels() As pdPSPChannel, ByRef srcChannelHeader As PSP_ChannelHeader, ByRef dstDIB As pdDIB, ByRef dstPalette() As RGBQuad, Optional ByVal dstPaletteSize As Long = 256) As Boolean
    
    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Reconstructing DIB from " & numPSPChannels & " source channels now..."
    
    'Start by initializing the DIB (and make it opaque, as many PSP images do not encode an alpha channel)
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateBlank srcChannelHeader.ch_ParentWidth, srcChannelHeader.ch_ParentHeight, 32, 0, 255
    
    Dim localCopy() As Byte
    ReDim localCopy(0) As Byte
    
    'Grab pointers to the destination image
    Dim dstSA As SafeArray2D, dstPixels() As RGBQuad
    dstDIB.WrapRGBQuadArrayAroundDIB dstPixels, dstSA
    
    Dim pxWidth As Long, pxHeight As Long
    pxWidth = dstDIB.GetDIBWidth - 1
    pxHeight = dstDIB.GetDIBHeight - 1
    
    Dim x As Long, y As Long
    
    'PSP stride is still under investigation.  The spec claims 4-byte scanline alignment,
    ' but 3rd-party readers claim this is untrue.  A final decision here is PENDING TESTING\
    Dim srcStride As Long
    srcStride = dstDIB.GetDIBWidth
    
    'Iterate through each channel in turn, placing its bits where they belong
    Dim i As Long
    For i = 0 To numPSPChannels - 1
        
        'Skip broken channels
        If (Not srcChannels(i).IsChannelOK) Then GoTo NextChannel
        
        'Skip selection channels
        If (srcChannels(i).GetChannelDIBType = PSP_DIB_SELECTION) Then GoTo NextChannel
        
        'Always start by copying the contents of this channel into a local struct (for faster access)
        If (srcChannels(i).GetChannelSize > UBound(localCopy) + 1) Then ReDim localCopy(0 To srcChannels(i).GetChannelSize - 1) As Byte
        CopyMemoryStrict VarPtr(localCopy(0)), srcChannels(i).GetChannelPtr(), srcChannels(i).GetChannelSize
        srcChannels(i).FreeChannelContents
        
        
        'If this channel is for a mask, it may require special offsets; flag this in advance.
        Dim isMaskChannel As Boolean
        'isMaskChannel = (srcChannels(i).GetChannelDIBType = PSP_DIB_ALPHA_MASK)
        'isMaskChannel = isMaskChannel Or (srcChannels(i).GetChannelDIBType = PSP_DIB_COMPOSITE_TRANS_MASK)
        isMaskChannel = isMaskChannel Or (srcChannels(i).GetChannelDIBType = PSP_DIB_PATTERN_TRANS_MASK)
        'isMaskChannel = isMaskChannel Or (srcChannels(i).GetChannelDIBType = PSP_DIB_THUMBNAIL_TRANS_MASK)
        'isMaskChannel = isMaskChannel Or (srcChannels(i).GetChannelDIBType = PSP_DIB_TRANS_MASK)
        isMaskChannel = isMaskChannel Or (srcChannels(i).GetChannelDIBType = PSP_DIB_USER_MASK)
        
        If PSP_DEBUG_VERBOSE Then
            PDDebug.LogAction "Channel DIB type is: " & srcChannels(i).GetChannelDIBType
            If isMaskChannel Then PDDebug.LogAction "(Note: this is a mask channel)"
        End If
        
        'Alpha bytes are not returned as-such; instead, they are a composite channel that is not
        ' the *first* channel in the image
        Select Case srcChannels(i).GetChannelType
            
            Case PSP_CHANNEL_COMPOSITE
                
                'Grayscale and indexed image
                If (i = 0) Then
                    
                    'Failsafe check for palette existence; if it's missing, grab a grayscale one as that's
                    ' likely what's required.
                    If (dstPaletteSize = 0) Then
                        If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "creating grayscale palette for composite channel..."
                        dstPaletteSize = 256
                        Palettes.GetPalette_Grayscale dstPalette
                    End If
                    
                    If isMaskChannel Then
                        If PSP_DEBUG_VERBOSE Then PDDebug.LogAction vbTab & "unexpected mask branch"
                    Else
                        If PSP_DEBUG_VERBOSE Then PDDebug.LogAction vbTab & "Generating RGB data from gray or indexed channel..."
                        For y = 0 To pxHeight
                        For x = 0 To pxWidth
                            dstPixels(x, y) = dstPalette(localCopy(y * srcStride + x))
                        Next x
                        Next y
                    End If
                
                'Actual alpha channel (i = 1 on gray/indexed images, i = 3 on RGBA)
                Else
                    
                    If isMaskChannel Then
                        If PSP_DEBUG_VERBOSE Then PDDebug.LogAction vbTab & "unexpected mask branch"
                    Else
                    
                        'TODO: masks will also appear here.  If we've already generated an alpha channel for
                        ' this image (as evidenced by a previous hit on this branch), we need to adjust the
                        ' existing alpha channel by the mask rather than overwriting it completely!
                        If PSP_DEBUG_VERBOSE Then PDDebug.LogAction vbTab & "Generating alpha channel..."
                        For y = 0 To pxHeight
                        For x = 0 To pxWidth
                            dstPixels(x, y).Alpha = localCopy(y * srcStride + x)
                        Next x
                        Next y
                        
                    End If
                    
                End If
            
            Case PSP_CHANNEL_RED
                If PSP_DEBUG_VERBOSE Then PDDebug.LogAction vbTab & "Generating red channel..."
                For y = 0 To pxHeight
                For x = 0 To pxWidth
                    dstPixels(x, y).Red = localCopy(y * srcStride + x)
                Next x
                Next y
                
            Case PSP_CHANNEL_GREEN
                If PSP_DEBUG_VERBOSE Then PDDebug.LogAction vbTab & "Generating green channel..."
                For y = 0 To pxHeight
                For x = 0 To pxWidth
                    dstPixels(x, y).Green = localCopy(y * srcStride + x)
                Next x
                Next y
            
            Case PSP_CHANNEL_BLUE
                If PSP_DEBUG_VERBOSE Then PDDebug.LogAction vbTab & "Generating blue channel..."
                For y = 0 To pxHeight
                For x = 0 To pxWidth
                    dstPixels(x, y).Blue = localCopy(y * srcStride + x)
                Next x
                Next y
            
        End Select
    
NextChannel:
    Next i
    
    dstDIB.UnwrapRGBQuadArrayFromDIB dstPixels
    dstDIB.SetAlphaPremultiplication True
    
    PSP_BuildDIBFromChannels = True

End Function

Public Function PSP_GetDIBTypeName(ByVal srcType As PSPDIBType) As String
    
    Select Case srcType
        Case PSP_DIB_IMAGE
            PSP_GetDIBTypeName = "Layer color bitmap"
        Case PSP_DIB_TRANS_MASK
            PSP_GetDIBTypeName = "Layer transparency mask bitmap"
        Case PSP_DIB_USER_MASK
            PSP_GetDIBTypeName = "Layer user mask bitmap"
        Case PSP_DIB_SELECTION
            PSP_GetDIBTypeName = "Selection mask bitmap"
        Case PSP_DIB_ALPHA_MASK
            PSP_GetDIBTypeName = "Alpha channel mask bitmap"
        Case PSP_DIB_THUMBNAIL
            PSP_GetDIBTypeName = "Thumbnail bitmap"
        Case PSP_DIB_THUMBNAIL_TRANS_MASK
            PSP_GetDIBTypeName = "Thumbnail transparency mask"
        Case PSP_DIB_ADJUSTMENT_LAYER
            PSP_GetDIBTypeName = "Adjustment layer bitmap"
        Case PSP_DIB_COMPOSITE
            PSP_GetDIBTypeName = "Composite image bitmap"
        Case PSP_DIB_COMPOSITE_TRANS_MASK
            PSP_GetDIBTypeName = "Composite image transparency"
        Case PSP_DIB_PAPER
            PSP_GetDIBTypeName = "Paper bitmap"
        Case PSP_DIB_PATTERN
            PSP_GetDIBTypeName = "Pattern bitmap"
        Case PSP_DIB_PATTERN_TRANS_MASK
            PSP_GetDIBTypeName = "Pattern transparency mask"
        Case Else
            PSP_GetDIBTypeName = "(unknown)"
    End Select
    
End Function

Public Function PSP_GetLayerTypeName(ByVal srcType As PSPLayerType) As String
    
    Select Case srcType
        Case keGLTUndefined
            PSP_GetLayerTypeName = "Undefined layer type"
        Case keGLTRaster
            PSP_GetLayerTypeName = "Standard raster layer"
        Case keGLTFloatingRasterSelection
            PSP_GetLayerTypeName = "Floating selection (raster)"
        Case keGLTVector
            PSP_GetLayerTypeName = "Vector layer"
        Case keGLTAdjustment
            PSP_GetLayerTypeName = "Adjustment layer"
        Case keGLTGroup
            PSP_GetLayerTypeName = "Group layer"
        Case keGLTMask
            PSP_GetLayerTypeName = "Mask layer"
        Case keGLTArtMedia
            PSP_GetLayerTypeName = "Art media layer"
        Case Else
            PSP_GetLayerTypeName = "(unknown)"
    End Select
    
End Function
