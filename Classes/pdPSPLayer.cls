VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPSPLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PSP (PaintShop Pro) Block Container
'Copyright 2020-2026 by Tanner Helland
'Created: 31/December/20
'Last updated: 04/January/21
'Last update: continue work on initial build
'
'This class describes a single "layer" inside a JASC/Corel Paint Shop Pro image file.
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

'Initial offset (*BEFORE* all layer data but *AFTER* the sub-block header) and block length
Private m_InitialOffset As Long, m_TotalLength As Long

'Layer header!  Not all values are used by PD, but all values are imported regardless
' (in case we find a use for them in the future).  Note also that the header is expliticly
' marked as "potentially expandable in the future" per the spec; when reading, PD will use
' the prefacing length marker to ensure the header is always loaded correctly.  (Same goes
' for old PSP versions that may not fill this entire header struct.)
Private Type PSP_LayerHeader
    pspl_Name As String     'PD will automatically generate this from ASCII or UTF-8 source as necessary (it varies by PSP version)
    pspl_LayerType As PSPLayerType
    pspl_ImageRect As RectL
    pspl_SavedImageRect As RectL
    pspl_Opacity As Single
    pspl_BlendMode As PSPBlendModes
    pspl_Flags As PSPLayerProperties
    pspl_AlphaLocked As Boolean
    pspl_GroupID As Byte
    pspl_MaskRect As RectL
    pspl_SavedMaskRect As RectL
    pspl_MaskLinked As Boolean
    pspl_MaskDisabled As Boolean
    pspl_MaskInverted As Boolean
    pspl_BlendRangeCount As Long
    pspl_BlendRanges(0 To 39) As Byte
    
    'Only in later PSP versions...
    pspl_UseHighlightColor As Boolean
    pspl_HighlightColor As Long
    
    'Remember - expansion fields may be added in the future!
    
    'Any following fields are internal PhotoDemon ones; they are not part of actual PSP Files.
    ' (These fields may also be filled using data from embedded sub-blocks in the file)
    pspl_LayerOK As Boolean
    pspl_LayerHasPixels As Boolean
    pspl_OriginalActiveLayer As Boolean
    pspl_GroupStart As Boolean
    pspl_VectorOK As Boolean
    
End Type

Private m_LayerHeader As PSP_LayerHeader

'Some layer types (vector, adjustment, group, etc) store extra layer information in a sub-block
' following the layer header.  This class is guaranteed to be non-null if a sub-block was found,
' but it's obviously up to this class to make sense of the information found in the sub-block.
Private m_LayerSubBlock As pdPSPBlock

'Group-type layers behave a differently from their analog in other formats (e.g. PSD).  Group start
' markers contain a count of how many "child" layers are in that group, *not* counting "grandchildren"
' (e.g. nested groups).  PhotoDemon will attempt to use this to restore groups in the final image
' as best it can.
Private m_NumLayersInGroup As Long, m_GroupLinked As Boolean

'Vector layers are comprised of one or more vector shapes (which can, in turn, be comprised of
' one or more vector shapes, ad infinitum).  This collection stores any such shapes.  Note that they
' can be standard vector shapes, or text items - PSP groups these under the same "vector" heading.
Private m_NumVectorShapes As Long, m_Shapes() As pdPSPShape

'Each layer is comprised of 0-5 channels.  Note that even in a 3-color RGB image, certain layer types
' (masks, adjustment layers, etc) may only have one channel - so you can't use the color-type of the
' parent image to interpret individual layer channel contents.  You need to check each channel's
' type in its block header.
Private m_BitmapCount As Long, m_ChannelCount As Long
Private m_Channels() As pdPSPChannel

'If this layer is successfully loaded, its raster contents will be placed in this DIB
Private m_LayerDIB As pdDIB

'Only valid after LoadLayer has been called (and the function has returned),
' and the header's LayerOK flags has been set to TRUE.
Friend Function ConstructMatchingDIB(ByRef srcPalette() As RGBQuad, ByVal srcPaletteSize As Long, ByRef srcHeader As PSPImageHeader) As Boolean
    
    If m_LayerHeader.pspl_LayerOK Then
        
        'Populate a temporary header to pass to the DIB renderer.  This contains essential
        ' details about the parent image, as well as external details relevant to interpreting
        ' channel contents - like compression type!
        Dim tmpHeader As PSP_ChannelHeader
        With tmpHeader
            .ch_ParentVersionMajor = srcHeader.psph_VersionMajor
            .ch_ParentWidth = m_LayerHeader.pspl_SavedImageRect.Right - m_LayerHeader.pspl_SavedImageRect.Left
            .ch_ParentHeight = m_LayerHeader.pspl_SavedImageRect.Bottom - m_LayerHeader.pspl_SavedImageRect.Top
            .ch_MaskWidth = m_LayerHeader.pspl_SavedMaskRect.Right - m_LayerHeader.pspl_SavedMaskRect.Left
            .ch_MaskHeight = m_LayerHeader.pspl_SavedMaskRect.Bottom - m_LayerHeader.pspl_SavedMaskRect.Top
            .ch_ParentBitDepth = srcHeader.psph_BitDepth
            .ch_Compression = srcHeader.psph_Compression
        End With
        
        'Some DIB types may be relevant in the future (e.g. adjustment layers), but right now, PD doesn't have
        ' a useful way to render them.  Skip such layers now.
        Dim skipLayer As Boolean
        skipLayer = (m_LayerHeader.pspl_LayerType = keGLTAdjustment)
        
        If skipLayer Then
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Skipping generation of raster data for layer : " & m_LayerHeader.pspl_Name & " because layer type is: " & m_LayerHeader.pspl_LayerType
        Else
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Attempting to generate raster data for layer: " & m_LayerHeader.pspl_Name & ", " & m_LayerHeader.pspl_LayerType
            ConstructMatchingDIB = ImageFormats_PSP.PSP_BuildDIBFromChannels(m_ChannelCount, m_Channels, tmpHeader, m_LayerDIB, srcPalette, srcPaletteSize)
        End If
        
        If ConstructMatchingDIB Then
            
            'To save space (presumably), PaintShop Pro is clever about "cropping out" only the
            ' relevant non-transparent rect from each layer, if one exists.  Only those pixels
            ' get saved to file, while the other transparent region is simply omitted.  (This
            ' is kind of stupid, because even their garbage RLE compression will effectively
            ' render those large transparent regions meaningless in terms of filespace - but I
            ' digress.)
            '
            'To correctly restore those transparent regions, PD needs to compare two embedded
            ' rects inside the layer: both the actual layer rect, and then the "saved" image
            ' rect.  The union of those two is what needs to be present in the final layer.
            With m_LayerHeader
                
                Dim unionNeeded As Boolean
                unionNeeded = (.pspl_ImageRect.Left <> .pspl_SavedImageRect.Left)
                unionNeeded = unionNeeded Or (.pspl_ImageRect.Top <> .pspl_SavedImageRect.Top)
                unionNeeded = unionNeeded Or (.pspl_ImageRect.Right <> .pspl_SavedImageRect.Right)
                unionNeeded = unionNeeded Or (.pspl_ImageRect.Bottom <> .pspl_SavedImageRect.Bottom)
                
                If unionNeeded Then
                    
                    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction vbTab & "Restoring original layer boundaries..."
                    
                    'Construct a new RGBA DIB at the size of the *image* layer
                    Dim newDIB As pdDIB
                    Set newDIB = New pdDIB
                    newDIB.CreateBlank .pspl_ImageRect.Right - .pspl_ImageRect.Left, .pspl_ImageRect.Bottom - .pspl_ImageRect.Top, 32, 0, 0
                    newDIB.SetInitialAlphaPremultiplicationState True
                    
                    'Copy the already-created DIB into its proper "final" location
                    Dim newLeft As Long, newTop As Long
                    newLeft = .pspl_SavedImageRect.Left
                    newTop = .pspl_SavedImageRect.Top
                    GDI.BitBltWrapper newDIB.GetDIBDC, newLeft, newTop, m_LayerDIB.GetDIBWidth, m_LayerDIB.GetDIBHeight, m_LayerDIB.GetDIBDC, 0, 0, vbSrcCopy
                    
                    'Swap references
                    Set m_LayerDIB = newDIB
                    
                End If
                
            End With
            
        Else
            If (Not ConstructMatchingDIB) Then m_LayerHeader.pspl_LayerOK = False
        End If
        
    End If
    
End Function

'Only valid after LoadLayer has been called (and the function has returned),
' and the header's LayerOK flags has been set to TRUE.
Friend Function ConstructMatchingDIB_Vector(ByRef srcPalette() As RGBQuad, ByVal srcPaletteSize As Long, ByRef srcHeader As PSPImageHeader) As Boolean
    
    Const funcName As String = "ConstructMatchingDIB_Vector"
    
    If (Me.IsLayerOK() And (m_NumVectorShapes > 0)) Then
        
        'Vector layers do not store layer boundaries, which makes them complicated to convert
        ' to raster objects.  Start with an image-sized layer; we'll trim it layer if possible.
        Set m_LayerDIB = New pdDIB
        m_LayerDIB.CreateBlank srcHeader.psph_Width, srcHeader.psph_Height, 32, 0, 0
        m_LayerDIB.SetInitialAlphaPremultiplicationState True
        
        'Ask each shape to render itself into the raster layer
        Dim i As Long, numSuccessfulRenders As Long
        For i = 0 To m_NumVectorShapes - 1
            If (Not m_Shapes(i) Is Nothing) Then
                If m_Shapes(i).IsShapeOK Then
                    If m_Shapes(i).RenderShapeIntoDIB(m_LayerDIB, srcPalette, srcPaletteSize, srcHeader) Then
                        numSuccessfulRenders = numSuccessfulRenders + 1
                    Else
                        InternalError funcName, "shape render failed: " & i, Nothing
                    End If
                End If
            End If
        Next i
        
        m_LayerHeader.pspl_VectorOK = (numSuccessfulRenders > 0)
        ConstructMatchingDIB_Vector = m_LayerHeader.pspl_VectorOK
        
    End If
    
End Function

'PSP files are confusing because text can be intermixed with other shapes in a given layer, so it's not
' as simple as "layer = text" vs "layer = raster".  We have to iterate the shape collection and hunt
' for text objects.
Friend Function DoesVectorLayerContainText() As Boolean
    If (m_LayerHeader.pspl_LayerType = keGLTVector) Then
        If (m_NumVectorShapes > 0) Then
            Dim i As Long
            For i = 0 To m_NumVectorShapes - 1
                If (Not m_Shapes(i) Is Nothing) Then
                    DoesVectorLayerContainText = m_Shapes(i).IsShapeText()
                    If DoesVectorLayerContainText Then Exit Function
                End If
            Next i
        End If
    End If
End Function

'All Get-prefixed functions are only valid after LoadLayer has been called (and the function has returned)
Friend Function GetLayerBlendMode_PSP() As PSPBlendModes
    GetLayerBlendMode_PSP = m_LayerHeader.pspl_BlendMode
End Function

Friend Function GetLayerBlendMode() As PD_BlendMode
    
    Select Case m_LayerHeader.pspl_BlendMode
        Case bmLAYER_BLEND_NORMAL
            GetLayerBlendMode = BM_Normal
        Case bmLAYER_BLEND_DARKEN
            GetLayerBlendMode = BM_Darken
        Case bmLAYER_BLEND_LIGHTEN
            GetLayerBlendMode = BM_Lighten
        Case bmLAYER_BLEND_LEGACY_HUE
            GetLayerBlendMode = BM_Hue
        Case bmLAYER_BLEND_LEGACY_SATURATION
            GetLayerBlendMode = BM_Saturation
        Case bmLAYER_BLEND_LEGACY_COLOR
            GetLayerBlendMode = BM_Color
        Case bmLAYER_BLEND_LEGACY_LUMINOSITY
            GetLayerBlendMode = BM_Luminosity
        Case bmLAYER_BLEND_MULTIPLY
            GetLayerBlendMode = BM_Multiply
        Case bmLAYER_BLEND_SCREEN
            GetLayerBlendMode = BM_Screen
        Case bmLAYER_BLEND_DISSOLVE
            GetLayerBlendMode = BM_Normal
        Case bmLAYER_BLEND_OVERLAY
            GetLayerBlendMode = BM_Overlay
        Case bmLAYER_BLEND_HARD_LIGHT
            GetLayerBlendMode = BM_HardLight
        Case bmLAYER_BLEND_SOFT_LIGHT
            GetLayerBlendMode = BM_SoftLight
        Case bmLAYER_BLEND_DIFFERENCE
            GetLayerBlendMode = BM_Difference
        Case bmLAYER_BLEND_DODGE
            GetLayerBlendMode = BM_ColorDodge
        Case bmLAYER_BLEND_BURN
            GetLayerBlendMode = BM_ColorBurn
        Case bmLAYER_BLEND_EXCLUSION
            GetLayerBlendMode = BM_Exclusion
        Case bmLAYER_BLEND_TRUE_HUE
            GetLayerBlendMode = BM_Hue
        Case bmLAYER_BLEND_TRUE_SATURATION
            GetLayerBlendMode = BM_Saturation
        Case bmLAYER_BLEND_TRUE_COLOR
            GetLayerBlendMode = BM_Color
        Case bmLAYER_BLEND_TRUE_LIGHTNESS
            GetLayerBlendMode = BM_Luminosity
        Case bmLAYER_BLEND_ADJUST
            GetLayerBlendMode = BM_Normal 'No one knows what this actually does (it's possibly a flag reserved for adjustment layers?)
        Case Else
            InternalError "GetLayerBlendMode", "unknown blendmode: " & m_LayerHeader.pspl_BlendMode, Nothing
            GetLayerBlendMode = BM_Normal
    End Select
    
End Function

Friend Function GetLayerDIB() As pdDIB
    Set GetLayerDIB = m_LayerDIB
End Function

Friend Function GetLayerName() As String
    GetLayerName = m_LayerHeader.pspl_Name
End Function

Friend Function GetLayerOffsetX() As Long
    GetLayerOffsetX = m_LayerHeader.pspl_ImageRect.Left
End Function

Friend Function GetLayerOffsetY() As Long
    GetLayerOffsetY = m_LayerHeader.pspl_ImageRect.Top
End Function

Friend Function GetLayerOpacity() As Single
    GetLayerOpacity = m_LayerHeader.pspl_Opacity
End Function

Friend Function GetLayerType() As PSPLayerType
    GetLayerType = m_LayerHeader.pspl_LayerType
End Function

Friend Function GetLayerVisibility() As Boolean
    GetLayerVisibility = (m_LayerHeader.pspl_Flags And keVisibleFlag) = keVisibleFlag
End Function

Friend Function GetLayerWasActiveLayer() As Boolean
    GetLayerWasActiveLayer = m_LayerHeader.pspl_OriginalActiveLayer
End Function

'Places all text shape indices in a pdStack object; return is the number of text shapes found.
' If return is 0, dstList is *not* guaranteed to be initialized.
Friend Function GetListOfTextIndices(ByRef dstList As pdStack) As Long
    
    If Me.DoesVectorLayerContainText() Then
    
        Set dstList = New pdStack
        
        Dim i As Long
        For i = 0 To m_NumVectorShapes - 1
            If m_Shapes(i).IsShapeText() Then dstList.AddInt i Else Debug.Print "not text: " & i
        Next i
        
        GetListOfTextIndices = dstList.GetNumOfInts()
        
    End If
    
End Function

Friend Function GetNumLayersInGroup() As Long
    GetNumLayersInGroup = m_NumLayersInGroup
End Function

'Bad indices will result in a null return.  Plan accordingly!
Friend Function GetVectorShape(ByVal shpIndex As Long) As pdPSPShape
    If (shpIndex >= 0) And (shpIndex < m_NumVectorShapes) Then Set GetVectorShape = m_Shapes(shpIndex)
End Function

Friend Function IsLayerGroupStart() As Boolean
    IsLayerGroupStart = m_LayerHeader.pspl_GroupStart
End Function

'Only valid after LoadLayer has been called (and the function has returned)
Friend Function IsLayerOK() As Boolean
    
    'Vector layer "OK" status varies by vector layer type
    If (m_LayerHeader.pspl_LayerType = keGLTVector) Then
        
        'On vector layers, we need at least one valid shape object to consider a layer "OK"
        If (m_NumVectorShapes > 0) Then
            Dim i As Long
            For i = 0 To m_NumVectorShapes - 1
                If m_Shapes(i).IsShapeOK Then
                    IsLayerOK = True
                    Exit For
                End If
            Next i
        End If
        
    'Raster layers are simpler: if they have good channel data, consider the whole layer OK
    Else
        IsLayerOK = m_LayerHeader.pspl_LayerOK
    End If
    
End Function

Friend Function IsLayerPixelBased() As Boolean
    IsLayerPixelBased = m_LayerHeader.pspl_LayerHasPixels And (m_LayerHeader.pspl_LayerType <> keGLTVector)
End Function

Friend Function IsVectorLayerAvailableAsRaster() As Boolean
    IsVectorLayerAvailableAsRaster = m_LayerHeader.pspl_VectorOK
End Function

'Assuming the source stream is pointing at the start of a layer block, attempt to load the layer.
' Returns psp_Success if successful, psp_Warning if stream alignment is okay but channel data is not,
' psp_Failure if stream alignment is unsaveable.  (If psp_Failure is returned, check initial stream
' pointer alignment - it may not have been pointing at a layer block when you called this function!)
'
'IMPORTANTLY: on psp_Success or psp_Warning, the passed stream pointer will now point at the *end* of
' this block.  You can simply continue reading the file as-is.  On failure, however, stream position
' is *not* guaranteed (mostly because if initial block validation fails, we have no way to reorient the
' pointer in a meaningful way - we can only reset it).  On failure, you need to abandon further parsing.
Friend Function LoadLayer(ByRef srcStream As pdStream, ByRef srcWarnings As pdStringStack, ByRef srcHeader As PSPImageHeader) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "LoadLayer"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'This layer will only be marked "OK" if we can retrieve at least one valid channel for it
    m_LayerHeader.pspl_LayerOK = False
    
    'The caller should have already performed block validation, but we want to be extra careful.
    Const PSP_BLOCK_MARKER As Long = &H4B427E
    If (srcStream.ReadLong() <> PSP_BLOCK_MARKER) Then
        LoadLayer = psp_Failure
        InternalError funcName, "stream misaligned", srcWarnings
        Exit Function
    End If
    
    If (srcStream.ReadIntUnsigned() <> PSP_LAYER_BLOCK) Then
        LoadLayer = psp_Failure
        InternalError funcName, "not a layer block!", srcWarnings
        Exit Function
    End If
    
    'PSPv5 files use a different header layout, so we must cache values differently
    ' depending on parent image version.
    Dim layerChunkSize As Long
    
    Dim isPSPv5 As Boolean
    isPSPv5 = (srcHeader.psph_VersionMajor <= 3)
    
    If isPSPv5 Then
        layerChunkSize = srcStream.ReadLong()
        m_TotalLength = srcStream.ReadLong()
        m_InitialOffset = srcStream.GetPosition()
    Else
        m_TotalLength = srcStream.ReadLong()
        m_InitialOffset = srcStream.GetPosition()
        layerChunkSize = srcStream.ReadLong()
    End If
    
    'Validate all lengths retrieved thus far
    If (m_TotalLength <= 0) Or (layerChunkSize <= 0) Then
        LoadLayer = psp_Failure
        InternalError funcName, "bad initial chunk or header size", srcWarnings
        Exit Function
    End If
    
    'Catastrophic failure states have now (mostly) been dealt with.
    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Layer block looks okay.  Parsing header now..."
    
    'Now we can finally read this layer's header!  Note that some header behavior
    ' may vary between PSP versions; PD tries to account for this wherever it can.
    
    'Layer name
    If isPSPv5 Then
        m_LayerHeader.pspl_Name = srcStream.ReadString_ASCII(256)
        m_LayerHeader.pspl_Name = Strings.TrimNull(m_LayerHeader.pspl_Name)
    Else
        Dim tmpNameLength As Long
        tmpNameLength = srcStream.ReadIntUnsigned()
        If (tmpNameLength > 0) Then m_LayerHeader.pspl_Name = srcStream.ReadString_UTF8(tmpNameLength)
    End If
    
    'I won't comment all the following fields.  Refer to the spec for details.
    m_LayerHeader.pspl_LayerType = srcStream.ReadByte()
    
    'Very old PSP files (v5?) may use the "undefined" layer type for raster images.  Silently convert
    ' these to raster layer types to simplify further parsing.  (If the layer has unusable data, the
    ' raster parser will automatically detect this and abandon processing as necessary.)
    If (m_LayerHeader.pspl_LayerType = keGLTUndefined) Then m_LayerHeader.pspl_LayerType = keGLTRaster
    
    'Image rects are strange beasts in PSP files.  They are only used for certain layer types, and you
    ' have to perform union ops between the rects to calculate actual layer contents.  I've left debug
    ' code related to them as I may need it if/when I implement mask support.
    srcStream.ReadBytesToBarePointer VarPtr(m_LayerHeader.pspl_ImageRect), 16
    srcStream.ReadBytesToBarePointer VarPtr(m_LayerHeader.pspl_SavedImageRect), 16
    'If PSP_DEBUG_VERBOSE Then Debug.Print m_LayerHeader.pspl_Name, m_LayerHeader.pspl_ImageRect.Left, m_LayerHeader.pspl_ImageRect.Top, m_LayerHeader.pspl_ImageRect.Right, m_LayerHeader.pspl_ImageRect.Bottom
    'If PSP_DEBUG_VERBOSE Then Debug.Print m_LayerHeader.pspl_Name, m_LayerHeader.pspl_SavedImageRect.Left, m_LayerHeader.pspl_SavedImageRect.Top, m_LayerHeader.pspl_SavedImageRect.Right, m_LayerHeader.pspl_SavedImageRect.Bottom
    
    m_LayerHeader.pspl_Opacity = CDbl(srcStream.ReadByte()) / 2.55
    m_LayerHeader.pspl_BlendMode = srcStream.ReadByte()
    
    If isPSPv5 Then
        If (srcStream.ReadByte() <> 0) Then m_LayerHeader.pspl_Flags = keVisibleFlag
    Else
        m_LayerHeader.pspl_Flags = srcStream.ReadByte()
    End If
    
    m_LayerHeader.pspl_AlphaLocked = (srcStream.ReadByte() <> 0)
    m_LayerHeader.pspl_GroupID = srcStream.ReadByte()
    srcStream.ReadBytesToBarePointer VarPtr(m_LayerHeader.pspl_MaskRect), 16
    srcStream.ReadBytesToBarePointer VarPtr(m_LayerHeader.pspl_SavedMaskRect), 16
    m_LayerHeader.pspl_MaskLinked = (srcStream.ReadByte() <> 0)
    m_LayerHeader.pspl_MaskDisabled = (srcStream.ReadByte() <> 0)
    m_LayerHeader.pspl_MaskInverted = (srcStream.ReadByte() <> 0)
    m_LayerHeader.pspl_BlendRangeCount = srcStream.ReadIntUnsigned()
    srcStream.ReadBytesToBarePointer VarPtr(m_LayerHeader.pspl_BlendRanges(0)), 40  '4 bytes * 5 ranges * 2 sets
    
    'Now we need to branch for v5 files; they use a totally different encoding scheme.
    ' (Bitmap and channel count are in a totally different chunk in later files.)
    If isPSPv5 Then
        m_BitmapCount = srcStream.ReadIntUnsigned()
        m_ChannelCount = srcStream.ReadIntUnsigned()
    Else
        
        'All subsequent header reads need to first validate header bytes remaining
        If (layerChunkSize - (srcStream.GetPosition() - m_InitialOffset) >= 1) Then m_LayerHeader.pspl_UseHighlightColor = (srcStream.ReadByte <> 0)
        If (layerChunkSize - (srcStream.GetPosition() - m_InitialOffset) >= 4) Then m_LayerHeader.pspl_HighlightColor = srcStream.ReadLong()
        
        'Bitmap and channel count are in a separate chunk; do not attempt to retrieve them yet!
        
    End If
    
    'The header has now been fully parsed.
    
    'Still more bytes may remain in future versions of the file, so use the initial
    ' chunk size to forcibly align the pointer to the end of the header.
    srcStream.SetPosition m_InitialOffset + layerChunkSize, FILE_BEGIN
    
    'Make a note of this position.  We may need to reference it if unknown blocks
    ' are encountered later on.
    Dim endOfHeader As Long
    endOfHeader = srcStream.GetPosition()
    
    'Next, we need to validate some header components before continuing with layer parsing.
    ' Note that how we validate depends on the layer type - some layer types (e.g. groups)
    ' don't need to embed things like saved pixel data, so 0x0 dimensions are OK.
    Dim dimensionsExpected As Boolean
    dimensionsExpected = (m_LayerHeader.pspl_LayerType = keGLTRaster)
    dimensionsExpected = dimensionsExpected Or (m_LayerHeader.pspl_LayerType = keGLTFloatingRasterSelection)
    dimensionsExpected = dimensionsExpected Or (m_LayerHeader.pspl_LayerType = keGLTMask)
    
    'Note that you'd think you'd want to check vector layers here, but vector layers do *not*
    ' store a layer boundary rect - an egregious oversight, I think, and a nightmare for validation.
    
    'First, discard any layers with obviously bad attributes.
    If ((m_LayerHeader.pspl_SavedImageRect.Right - m_LayerHeader.pspl_SavedImageRect.Left) <= 0) And dimensionsExpected Then
        okToProceed = psp_Warning
        If dimensionsExpected Then InternalError funcName, "invalid width", srcWarnings
    End If
    
    If ((m_LayerHeader.pspl_SavedImageRect.Bottom - m_LayerHeader.pspl_SavedImageRect.Top) <= 0) And dimensionsExpected Then
        okToProceed = psp_Warning
        If dimensionsExpected Then InternalError funcName, "invalid height", srcWarnings
    End If
    
    'If any validation tests failed, exit now - but critically, remember to restore stream position
    ' before bailing!
    If (okToProceed <> psp_Success) Then
        
        If dimensionsExpected Then
            InternalError funcName, "WARNING: one or more critical validations failed.  This layer will not be imported.", srcWarnings
        Else
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Skipping layer """ & m_LayerHeader.pspl_Name & """, type: " & ImageFormats_PSP.PSP_GetLayerTypeName(m_LayerHeader.pspl_LayerType) & "..."
        End If
        
        m_LayerHeader.pspl_LayerHasPixels = False
        
        LoadLayer = okToProceed
        GoTo EarlyTermination
        
    End If
    
    'potential TODO: more validations depending on layer type?  It's hard to guess at "valid"
    ' parameters for layers I can't easily render and see... I'll return to this later.
    
    'Unfortunately, the PSP spec has another poor implementation decision here.
    '
    ' What follows the layer header is variable.  It can either be a...
    ' - "layer extension sub-block", which contains additional detail for *some* layer types, or...
    ' - "layer bitmap chunk" which is exactly what it sounds like.
    '
    'The spec only allows the extension sub-block for certain layer types.  For other types,
    ' it is explicitly disallowed.  They also reserve the right to add more layer types in the
    ' future that may or may not use an extension block.
    '
    'But there is no guaranteed-unique special key, or header, or ID for these extension blocks
    ' (especially ones that "may exist in the future"), so there's no foolproof way to tell if
    ' you're looking at a bitmap layer chunk or some kind of extension sub-block.
    '
    'Even worse, from what I've seen in 3rd-party forums...
    ' (https://forums.getpaint.net/topic/22817-paint-shop-pro-filetype-2015-01-05/?do=findComment&comment=410768)
    ' ...Corel at some point reversed the spec's explicit disallowal of sub-blocks for certain
    ' layer types (default raster layers!), which means that none of the current spec's rules
    ' are guaranteed to be observed in the future.  (Which sort of defeats the whole purpose
    ' of a specification, no?)
    '
    'So what is a 3rd-party decoder like PD supposed to do?
    '
    'Well, we obviously work around these issues as best we can.  PSP sub-blocks do start with a
    ' fixed (for now) 4-byte ID, which we can obviously check.  This ID appears in the same place
    ' that a chunk-length value would appear if a sub-block is *not* present.  So there's only
    ' potential for issues if a bitmap chunk length in a non-sub-block-containing layer were to
    ' coincedentally equal the fixed 4-byte ID value of a sub-block identifier. The odds of this
    ' aren't great, but because the ID value is just a normal 4-byte integer (4,932,222), the odds
    ' *are* non-zero!
    '
    'Anyway, I mention all this because I don't have a reference for how to "best" implement this.
    ' When writing files, I'd probably just never include a sub-block on raster layers (especially
    ' because all publicly available specs explicitly disallow this), and if a raster layer's size
    ' somehow happened to match the sub-block ID, I'd just pad it with dummy bytes to ensure its
    ' value was different.
    If (Not isPSPv5) Then
    
        Dim subBlockExists As Boolean
        subBlockExists = (srcStream.ReadLong() = PSP_BLOCK_MARKER)
        
        'For now, cache the sub-block's offset, type, and length; this makes it trivial for us to
        ' add support for additional parsing as I attempt to support more custom layer types
        If subBlockExists Then
            
            'Cache the contents of the sub-block; we'll decide later what to do with it.
            Set m_LayerSubBlock = New pdPSPBlock
            
            Dim subBlockPos As Long, subBlockID As PSPBlockID, subBlockLength As Long
            subBlockID = srcStream.ReadIntUnsigned()
            subBlockLength = srcStream.ReadLong()
            subBlockPos = srcStream.GetPosition()
            
            m_LayerSubBlock.InitializeBlock subBlockPos, subBlockID, subBlockLength
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Found layer sub-block: " & m_LayerSubBlock.GetBlockIDAsString() & " (" & subBlockID & "), length=" & subBlockLength
            
            'Attempt to deal with certain sub-block types now
            okToProceed = ParseLayerSubBlock(srcStream, srcWarnings, srcHeader)
            
            'In the future, we may want to abandon processing of certain block types if we can't
            ' make sense of their associated sub-blocks.  But for now, sub-blocks are mostly a
            ' curiosity, so as long as we didn't outright fail while parsing it, carry on normally.
            If (okToProceed < psp_Failure) Then okToProceed = psp_Success
            
            'Regardless of sub-block contents, always skip to the end of the sub-block after processing.
            ' (This ensures stream alignment for blocks we don't understand, or new blocks that may be
            ' added in the future.)
            srcStream.SetPosition subBlockPos + subBlockLength, FILE_BEGIN
        
        'If a sub-block doesn't exist, retreat the stream pointer
        Else
            srcStream.SetPosition -4, FILE_CURRENT
        End If
    
    End If
    
    'Any sub-blocks have now been read or skipped, meaning we're pointing at this layer's
    ' first channel block (if a PSPv5 image), or a bitmap header "chunk" (PSPv6+).
    If (Not isPSPv5) Then
    
        'Modern versions of PSP may embed additional blocks here!  This can happen on all types
        ' of layers, including non-raster ones.  The meaning and/or interpretation of these blocks
        ' is unknown (owing to Corel's unwillingness to release modern format specs, argh).
        
        'If we encounter an unexpected block here, simply skip over it.
        Dim failsafeCheck As Boolean
        failsafeCheck = (srcStream.ReadLong() = PSP_BLOCK_MARKER)
        If failsafeCheck Then
            subBlockID = srcStream.ReadIntUnsigned()
            subBlockLength = srcStream.ReadLong()
            srcStream.SetPosition subBlockLength, FILE_CURRENT
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Mystery block encountered in unexpected position; skipping it now..."
        
        'No "mystery" block found.  Retreat the cursor and proceed normally.
        Else
            srcStream.SetPosition -4, FILE_CURRENT
        End If
        
        Dim bmpHeaderOffset As Long
        bmpHeaderOffset = srcStream.GetPosition()
        
        Dim bmpHeaderChunkSize As Long
        bmpHeaderChunkSize = srcStream.ReadLong()
        
        If (bmpHeaderChunkSize <= 0) Then
            If PSP_DEBUG_VERBOSE Then InternalError funcName, "bitmap doesn't exist, which is okay for some layer types; skipping to end of layer chunk...", srcWarnings
            LoadLayer = psp_Warning
            GoTo EarlyTermination
        End If
        
        m_BitmapCount = srcStream.ReadIntUnsigned()
        m_ChannelCount = srcStream.ReadIntUnsigned()
        
        'The spec explicitly allows future expansion fields.  Use the chunk length to skip ahead.
        srcStream.SetPosition bmpHeaderOffset + bmpHeaderChunkSize, FILE_BEGIN
        
        'If bitmaps and/or channels do not exist, mark the layer as good (if it's not expected to
        ' have pixels, obviously) but do not attempt to process it further!
        If (m_BitmapCount <= 0) Or (m_ChannelCount <= 0) Then
            
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "This layer has no raster data (" & m_BitmapCount & ":" & m_ChannelCount & "); skipping to end of layer block..."
            
            m_LayerHeader.pspl_LayerHasPixels = False
            If (Not dimensionsExpected) Then
                m_LayerHeader.pspl_LayerOK = True
                LoadLayer = psp_Success
            Else
                m_LayerHeader.pspl_LayerOK = False
                LoadLayer = psp_Warning
            End If
            
            GoTo EarlyTermination
            
        Else
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Layer contains non-zero bitmaps/channels: " & m_BitmapCount & ", " & m_ChannelCount
        End If
        
    End If
    
    'With all headers parsed, we are now pointing at this layer's channel data.
    ' (Layers without channel data have already exited the function.)
    
    'Per the spec:
    ' "Layer Channel Sub-Blocks: Contains all the layer's channels.
    '  Possible channels include
    '  1) one or three Channel Sub-Blocks defining the layer's color bitmap
    '  2) one channel defining the layer’s transparency mask
    '  3) one channel defining the layer’s user mask or the adjustment layer bitmap."
    '
    'Regardless of layer type, we're going to load all encountered channels.  Our parent
    ' class will later decide how to handle them (including merging masks with bitmap contents, etc)
    ReDim m_Channels(0 To m_ChannelCount - 1) As pdPSPChannel
    
    Dim numGoodChannels As Long
    numGoodChannels = 0
    
    Dim i As Long
    For i = 0 To m_ChannelCount - 1
    
        Set m_Channels(i) = New pdPSPChannel
        
        'Populate a temporary header to pass to the target channel.  This contains essential
        ' details about the parent image, as well as external details relevant to interpreting
        ' channel contents - like compression type!
        Dim tmpHeader As PSP_ChannelHeader
        With tmpHeader
            .ch_ParentVersionMajor = srcHeader.psph_VersionMajor
            .ch_ParentWidth = m_LayerHeader.pspl_SavedImageRect.Right - m_LayerHeader.pspl_SavedImageRect.Left
            .ch_ParentHeight = m_LayerHeader.pspl_SavedImageRect.Bottom - m_LayerHeader.pspl_SavedImageRect.Top
            .ch_MaskWidth = m_LayerHeader.pspl_SavedMaskRect.Right - m_LayerHeader.pspl_SavedMaskRect.Left
            .ch_MaskHeight = m_LayerHeader.pspl_SavedMaskRect.Bottom - m_LayerHeader.pspl_SavedMaskRect.Top
            .ch_ParentBitDepth = srcHeader.psph_BitDepth
            .ch_Compression = srcHeader.psph_Compression
        End With
        
        If (m_Channels(i).LoadChannel(srcStream, srcWarnings, tmpHeader) >= psp_Failure) Then
            InternalError funcName, "catastrophic failure; channel #" & (i + 1) & " will be ignored...", srcWarnings
        Else
            'If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Channel " & (i + 1) & " (" & ImageFormats_PSP.PSP_GetDIBTypeName(m_Channels(i).GetChannelDIBType()) & ") looks okay"
            numGoodChannels = numGoodChannels + 1
        End If
        
    Next i
    
    'All channels have now been loaded!  If one or more were valid, construct a matching DIB for
    ' this layer.
    If (numGoodChannels > 0) Then
        
        'We can finally mark this layer as OK; yay!
        m_LayerHeader.pspl_LayerOK = True
        m_LayerHeader.pspl_LayerHasPixels = True
        
    Else
        InternalError funcName, "no valid channels found; layer will be ignored", srcWarnings
        okToProceed = psp_Warning
    End If
    
    'Our work here is done!
    LoadLayer = okToProceed
    
    'Before exiting, move the stream pointer to the end of this layer block.
EarlyTermination:
    srcStream.SetPosition m_InitialOffset + m_TotalLength, FILE_BEGIN
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description, srcWarnings
    srcWarnings.AddString "Internal error in pdPSPLayer." & funcName & ", #" & Err.Number & ": " & Err.Description
    LoadLayer = psp_Failure
    
End Function

'This dummy function is used internally by PD to mark the start of layer groups
' (since PSP files don't embed this automatically; instead, they use ref-counting and only
' embed *end* markers uuuuuugh).
Friend Sub LoadLayer_GroupMarker(ByRef srcLayerName As String, ByVal srcLayerVisibility As Long, ByVal srcLayerBlendMode As PSPBlendModes, ByVal srcLayerOpacity As Single)
    With m_LayerHeader
        .pspl_BlendMode = srcLayerBlendMode
        .pspl_LayerHasPixels = False
        .pspl_LayerOK = True
        .pspl_LayerType = keGLTGroup
        .pspl_Name = srcLayerName
        .pspl_Opacity = srcLayerOpacity
        If srcLayerVisibility Then .pspl_Flags = keVisibleFlag
        .pspl_GroupStart = True
    End With
End Sub

Friend Sub MarkAsActiveLayer()
    m_LayerHeader.pspl_OriginalActiveLayer = True
End Sub

'Note the requirements for stream positioning before calling this function!  It requires
' specific stream alignment to work.
Private Function ParseLayerSubBlock(ByRef srcStream As pdStream, ByRef srcWarnings As pdStringStack, ByRef srcHeader As PSPImageHeader) As PD_PSPResult
    
    Const funcName As String = "ParseLayerSubBlock"
    
    'Failsafe check only
    If (m_LayerSubBlock Is Nothing) Then
        InternalError funcName, "sub-block wasn't initialized?", srcWarnings
        ParseLayerSubBlock = psp_Failure
        Exit Function
    End If
    
    Dim streamPosAtEntry As Long
    streamPosAtEntry = srcStream.GetPosition()
    
    'Handling is obviously contingent on block type
    Select Case m_LayerSubBlock.GetBlockID()
    
        'Adjustment layer
        Case PSP_ADJUSTMENT_EXTENSION_BLOCK
        
        'Vector layer
        Case PSP_VECTOR_EXTENSION_BLOCK
        
            'Vector layers comprise both vector shapes *and* text layers.  (Much like PD,
            ' these layer types are closely related in PSP files.)  We parse them similarly.
            Dim chunkSize As Long
            chunkSize = srcStream.ReadLong()
            
            If (chunkSize > 4) Then
                
                m_NumVectorShapes = srcStream.ReadLong()
                If (m_NumVectorShapes > 0) Then
                
                    ReDim m_Shapes(0 To m_NumVectorShapes - 1) As pdPSPShape
                    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Layer reports " & m_NumVectorShapes & " shapes"
                    
                    'Future expansions fields are explicitly allowed.  Skip to the end of the chunk.
                    srcStream.SetPosition streamPosAtEntry + chunkSize, FILE_BEGIN
                    
                    'Attempt to parse [n] vector shapes
                    Dim i As Long
                    For i = 0 To m_NumVectorShapes - 1
                        Set m_Shapes(i) = New pdPSPShape
                        If (m_Shapes(i).LoadShape(srcStream, srcWarnings, srcHeader) >= psp_Failure) Then
                            InternalError funcName, "failed to parse shape #" & (i + 1), srcWarnings
                            ParseLayerSubBlock = psp_Failure
                            Exit Function
                        Else
                            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Parsed vector shape #" & (i + 1) & " successfully"
                        End If
                    Next i
                    
                Else
                    InternalError funcName, "no vector shapes in block?", srcWarnings
                End If
                    
            Else
                InternalError funcName, "empty vector block?", srcWarnings
                ParseLayerSubBlock = psp_Success
            End If
        
        'Group layer (start markers only)
        Case PSP_GROUP_EXTENSION_BLOCK
            
            'Group layers are incredibly simple: they just store a chunk size (which we can ignore;
            ' the parent layer will take care of stream alignment for us), number of layers in the
            ' group, and whether the group is linked.
            If (srcStream.ReadLong() > 4) Then
                m_NumLayersInGroup = srcStream.ReadLong()
                m_GroupLinked = (srcStream.ReadByte() <> 0)
                ParseLayerSubBlock = psp_Success
            End If
        
        'Mask layer
        Case PSP_MASK_EXTENSION_BLOCK
        
        'Art media
        Case PSP_ART_MEDIA_BLOCK
        
        'Future-proofing
        Case Else
            
            'Block "33" shows up constantly in modern PSP files.  No one knows what it does.
            ' Don't pollute the debug log with reports for that particular mystery ID.
            If (m_LayerSubBlock.GetBlockID() > 33) Then InternalError funcName, "unknown layer sub-block: " & m_LayerSubBlock.GetBlockID(), srcWarnings
            ParseLayerSubBlock = psp_Warning
            Exit Function
    
    End Select
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, ByRef parentWarningStack As pdStringStack, Optional ByVal writeDebugLog As Boolean = True)
    
    Dim errText As String
    errText = "pdPSPLayer." & funcName & "() reported an error: " & errDescription
    If (Not parentWarningStack Is Nothing) Then parentWarningStack.AddString errText
    
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction errText
    Else
        Debug.Print errText
    End If
    
End Sub
