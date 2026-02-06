VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdXCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon XCF (GIMP native file format) Container and Parser
'Copyright 2022-2026 by Tanner Helland
'Created: 30/March/22
'Last updated: 25/May/22
'Last update: add support for GZip variants (.xcfgz, .xcf.gz)
'
'This class handles GIMP XCF format parsing duties.  It is custom-built for PhotoDemon, with an emphasis
' on performance and proper color-management of all imported data.
'
'To my knowledge, all color models (indexed, gray, and RGB, all with/without alpha) and precisions
' (a huge list of both integer- and floating-point) are covered by this importer.  This level of coverage
' was a huge project and there may be outliers for esoteric combinations, particularly those not exposed
' in GIMP itself (like double-precision floating point images).  Detailed testing of such combinations
' will have to wait until GIMP allows me to generate them!
'
'As with all 3rd-party XCF engines, GIMP has many features that don't have direct analogs in PhotoDemon.
' Many of these features are still parsed by this class (making it simple to support them in the future),
' but they may not "appear" in PhotoDemon's current version, at least not in a way users can easily access.
' My ongoing goal is to expose these features directly as they are implemented in PD itself.
'
'Unless otherwise noted, all code in this class is my original work.  I've based my work off the
' "official" XCF spec at this URL (link good as of March 2022):
' https://gitlab.gnome.org/GNOME/gimp/-/blob/master/devel-docs/specifications/xcf.txt
'
'Many thanks to the authors of that XCF spec - it was *immensely* helpful in understanding various
' quirks of the XCF format (which are many).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output;
' this will dump a variety of diagnostic information to the debug log.
Private Const XCF_DEBUG_VERBOSE As Boolean = False

'XCF-specific enums are split between this class and the child pdXCFLayer class
Private Enum xcf_ColorMode
    xcf_RGB = 0
    xcf_Grayscale = 1
    xcf_Indexed = 2
End Enum

#If False Then
    Private Const xcf_RGB = 0, xcf_Grayscale = 1, xcf_Indexed = 2
#End If

'For inexplicable reasons, layers have their own color mode enum (which is different than the image-wide mode!)
Private Enum xcf_ColorModeLayer
    xcf_RGBX = 0
    xcf_RGBA = 1
    xcf_GrayX = 2
    xcf_GrayA = 3
    xcf_IndexedX = 4
    xcf_IndexedA = 5
End Enum

#If False Then
    Private Const xcf_RGBX = 0, xcf_RGBA = 1, xcf_GrayX = 2, xcf_GrayA = 3, xcf_IndexedX = 4, xcf_IndexedA = 5
#End If

Private Enum xcf_Compression
    xcf_Compress_None = 0       'Supposedly never appears "in the wild"
    xcf_Compress_RLE = 1        'Default
    xcf_Compress_ZLib = 2       'Optional on modern XCF
    xcf_Compress_Fractal = 3    'Reserved for future use; not actually implemented yet
End Enum

#If False Then
    Private Const xcf_Compress_None = 0, xcf_Compress_RLE = 1, xcf_Compress_ZLib = 2, xcf_Compress_Fractal = 3
#End If

Private Enum xcf_Precision
    xcf_08bitIntLinear = 0
    xcf_08bitIntGamma = 1
    xcf_16bitIntLinear = 2
    xcf_16bitIntGamma = 3
    xcf_32bitIntLinear = 4
    xcf_32bitIntGamma = 5
    xcf_16bitFltLinear = 6
    xcf_16bitFltGamma = 7
    xcf_32bitFltLinear = 8
    xcf_32bitFltGamma = 9
    xcf_64bitFltLinear = 10
    xcf_64bitFltGamma = 11
End Enum

#If False Then
    Private Const xcf_08bitIntLinear = 0, xcf_08bitIntGamma = 0, xcf_16bitIntLinear = 0, xcf_16bitIntGamma = 0, xcf_32bitIntLinear = 0, xcf_32bitIntGamma = 0, xcf_16bitFltLinear = 0, xcf_16bitFltGamma = 0, xcf_32bitFltLinear = 0, xcf_32bitFltGamma = 0, xcf_64bitFltLinear = 0, xcf_64bitFltGamma = 0
#End If

Private Enum xcf_PropertyID
    xcf_PROP_END = 0
    xcf_PROP_COLORMAP = 1
    xcf_PROP_ACTIVE_LAYER = 2
    xcf_PROP_ACTIVE_CHANNEL = 3
    xcf_PROP_SELECTION = 4
    xcf_PROP_FLOATING_SELECTION = 5
    xcf_PROP_OPACITY = 6
    xcf_PROP_BLEND_MODE = 7
    xcf_PROP_VISIBLE = 8
    xcf_PROP_LINKED = 9
    xcf_PROP_LOCK_ALPHA = 10
    xcf_PROP_APPLY_MASK = 11
    xcf_PROP_EDIT_MASK = 12
    xcf_PROP_SHOW_MASK = 13
    xcf_PROP_SHOW_MASKED = 14
    xcf_PROP_OFFSETS = 15
    xcf_PROP_COLOR = 16
    xcf_PROP_COMPRESSION = 17
    xcf_PROP_GUIDES = 18
    xcf_PROP_RESOLUTION = 19
    xcf_PROP_TATTOO = 20
    xcf_PROP_PARASITES = 21
    xcf_PROP_UNIT = 22
    xcf_PROP_PATHS = 23
    xcf_PROP_USER_UNIT = 24
    xcf_PROP_VECTORS = 25
    xcf_PROP_TEXT_LAYER_FLAGS = 26
    xcf_PROP_LOCK_CONTENT = 28
    xcf_PROP_GROUP_ITEM = 29
    xcf_PROP_ITEM_PATH = 30
    xcf_PROP_GROUP_ITEM_FLAGS = 31
    xcf_PROP_LOCK_POSITION = 32
    xcf_PROP_FLOAT_OPACITY = 33
    xcf_PROP_COLOR_TAG = 34
    xcf_PROP_COMPOSITE_MODE = 35
    xcf_PROP_COMPOSITE_SPACE = 36
    xcf_PROP_BLEND_SPACE = 37
    xcf_PROP_FLOAT_COLOR = 38
    xcf_PROP_SAMPLE_POINTS = 39
    xcf_PROP_ITEM_SET = 40
    xcf_PROP_ITEM_SET_ITEM = 41
    xcf_PROP_LOCK_VISIBILITY = 42
End Enum

#If False Then
    Private Const xcf_PROP_END = 0, xcf_PROP_COLORMAP = 1, xcf_PROP_ACTIVE_LAYER = 2, xcf_PROP_ACTIVE_CHANNEL = 3, xcf_PROP_SELECTION = 4, xcf_PROP_FLOATING_SELECTION = 5, xcf_PROP_OPACITY = 6, xcf_PROP_BLEND_MODE = 7, xcf_PROP_VISIBLE = 8, xcf_PROP_LINKED = 9
    Private Const xcf_PROP_LOCK_ALPHA = 10, xcf_PROP_APPLY_MASK = 11, xcf_PROP_EDIT_MASK = 12, xcf_PROP_SHOW_MASK = 13, xcf_PROP_SHOW_MASKED = 14, xcf_PROP_OFFSETS = 15, xcf_PROP_COLOR = 16, xcf_PROP_COMPRESSION = 17, xcf_PROP_GUIDES = 18, xcf_PROP_RESOLUTION = 19
    Private Const xcf_PROP_TATTOO = 20, xcf_PROP_PARASITES = 21, xcf_PROP_UNIT = 22, xcf_PROP_PATHS = 23, xcf_PROP_USER_UNIT = 24, xcf_PROP_VECTORS = 25, xcf_PROP_TEXT_LAYER_FLAGS = 26, xcf_PROP_LOCK_CONTENT = 28, xcf_PROP_GROUP_ITEM = 29
    Private Const xcf_PROP_ITEM_PATH = 30, xcf_PROP_GROUP_ITEM_FLAGS = 31, xcf_PROP_LOCK_POSITION = 32, xcf_PROP_FLOAT_OPACITY = 33, xcf_PROP_COLOR_TAG = 34, xcf_PROP_COMPOSITE_MODE = 35, xcf_PROP_COMPOSITE_SPACE = 36, xcf_PROP_BLEND_SPACE = 37, xcf_PROP_FLOAT_COLOR = 38, xcf_PROP_SAMPLE_POINTS = 39
    Private Const xcf_PROP_ITEM_SET = 40, xcf_PROP_ITEM_SET_ITEM = 41, xcf_PROP_LOCK_VISIBILITY = 42
#End If

'Theses structs are implemented in a PhotoDemon-specific way.  They are not meant to store *all* GIMP features;
' just the ones PD is capable of using.  (Note also that they deliberately store things like 32-bit pointers
' instead of 64-bit ones; these are the type of implementation details that are PD-specific.)
Private Type xcf_Property
    propID As xcf_PropertyID
    propSize As Long
    propData() As Byte
End Type

Private Type xcf_Channel
    ptrInFile As Long
    cWidth As Long      'Must match width of parent object (layer or image)
    cHeight As Long     'Must match height of parent object (layer or image)
    cName As String
    cNumProperties As Long
    cProperties() As xcf_Property
    cBytesPerPixel As Long
    numTilesX As Long
    numTilesY As Long
    numTilesTotal As Long
    ptrToTiles() As Long
    
    'If this channel is fully decoded by PD, the end result (in 8-bit format) will be stored here.
    ' The array is guaranteed to have dimensions [0 to cWidth - 1, 0 to cHeight - 1] - but ONLY if the channel
    ' was actually decoded.  (It's assumed the caller has a way of knowing that, since channels are
    ' only retrieved if expliticly requested by their parent object.)
    cBytes() As Byte
    
End Type

Private Type xcf_Layer
    
    'These members are retrieved from a dedicated table immediately following the image header:
    ptrInFile As Long
    
    'These members are retrieved from the layer block pointed to by ptrInFile:
    lWidth As Long
    lHeight As Long
    lColorMode As xcf_ColorModeLayer
    lName As String
    lNumProperties As Long
    lProperties() As xcf_Property
    ptrToPixels As Long
    ptrToMask As Long
    
    'These members are retrieved from (and/or calculated by) the pixel hierarchy struct located
    ' at ptrToPixels:
    lBytesPerPixel As Long
    numTilesX As Long
    numTilesY As Long
    numTilesTotal As Long
    ptrToTiles() As Long    'Guaranteed dimensioned to [0 to numTilesTotal - 1], IIF numTilesX/Y are non-zero
    
    'The layer mask, if any, will be stored here:
    lMask As xcf_Channel
    
    'These members are retrieved from individual layer property blocks (many of which are optional).
    ' They must also be translated from GIMP values to PD ones; see the Import_Stage4c_DecodeLayerProps function
    ' for detailsl
    lOpacity As Single
    lOffsetX As Long
    lOffsetY As Long
    lBlendMode As PD_BlendMode
    lAlphaLocked As Boolean
    lIsGroupMarker As Boolean
    lHasMask As Boolean
    lMaskActive As Boolean
    lIsActiveLayer As Boolean
    lIsVisible As Boolean
    
    'These members are generated by PhotoDemon using all data given above:
    lDIB As pdDIB
    
End Type

Private m_Layers() As xcf_Layer, m_numOfLayers As Long

'The image itself can embed channels (selections are one example).  Currently the pointers of all
' channels are retrieved during parsing, but image-wide channels are *not* processed further.
' Layer-specific channels (like masks) *are* parsed, but they are stored inside their parent
' xcf_Layer object, not in this "image-level" array.
Private m_Channels() As xcf_Channel, m_numOfChannels As Long

'Embedded ICC profiles are read and used by PD.
Private m_Profile As pdICCProfile

'GIMP encodes images as individual 64x64 tiles (only right/bottom tiles are allowed to be
' smaller than this).  To avoid repeat memory allocations, we reuse a single persistent
' tile DIB during image assembly.
Private m_Tile As pdDIB

'Image-wide properties include things like a global palette, tile compression IDs, etc
Private m_numImageProperties As Long, m_ImageProperties() As xcf_Property

'Canvas width/height, color model, and precision all come from the file header
Private m_ImageWidth As Long, m_ImageHeight As Long
Private m_imageColorMode As xcf_ColorMode, m_imagePrecision As xcf_Precision
Private m_imageResolutionPPI As Single

'Palette, if any
Private m_numPaletteColors As Long, m_Palette() As RGBQuad

'XCF version is critical to correct parsing; different versions have different fields in different places,
' and use different sizes for low-level values like internal file pointers
Private m_xcfVersion As Long

'Byte-by-byte access is provided, as always, by a pdStream instance
Private m_Stream As pdStream

'If the XCF file is embedded inside some other container (like .gz, as in "image.xcf.gz"), we will try to
' extract the bytes to a temporary array, then parse that.  This may or may not work depending on memory constraints.
Private m_tmpXcfCopy() As Byte

'Validate a source filename as XCF format.  Validation *does* touch the file -
' we must validate a "magic number" in the header.
Friend Function IsFileXCF(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True, Optional ByVal onSuccessLeaveStreamOpen As Boolean = False) As Boolean
    
    Dim potentiallyXCF As Boolean
    potentiallyXCF = Files.FileExists(srcFilename)
    If potentiallyXCF Then potentiallyXCF = (Files.FileLenW(srcFilename) > 26)
    
    'Check extension up front, if requested.
    Dim useGZMode As Boolean
    If potentiallyXCF Then
        
        'Regardless of extension validation, we need to determine if a GZ-encoded file is being used
        ' (because that changes how we parse the file).
        useGZMode = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "xcfgz", True) Or Strings.StringsEqualRight(srcFilename, "xcf.gz", True)
        
        'Validate against a normal "XCF" extension or the various gz-encoded sub-types
        If requireValidFileExtension Then
            potentiallyXCF = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "xcf", True)
            If (Not potentiallyXCF) Then potentiallyXCF = useGZMode
        End If
        
    End If
    
    'Proceed with deeper validation as necessary
    If potentiallyXCF Then
        
        'Attempt to load the file.  How we do this varies by XCF file type (GIMP allows you to embed an XCF file
        ' inside various containers, e.g. xcf.gz).
        Set m_Stream = New pdStream
        Dim streamOK As Boolean
        
        If useGZMode Then
            
            If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "XCF is inside GZip wrapper.  Attempting decompression..."
            
            'Load the file (which is currently just a compressed stream) into memory
            Dim origBytes() As Byte
            If Files.FileLoadAsByteArray(srcFilename, origBytes) Then
                
                'If the original file was < 2 GB, its size will be *hypothetically* stored in the last 4 bytes of the stream.
                Dim decompressedSize As Long
                GetMem4 VarPtr(origBytes(UBound(origBytes) - 3)), decompressedSize
                
                'To reduce the possibility of atrociously bad guesses in a malformed file (which would likely have us
                ' reading a CRC value instead of the original file size), mask the size against a safe upper limit.
                decompressedSize = decompressedSize And &HFFFFFFF
                
                'If the size field is bad, attempt an arbitrary buffer size (which we'll increase as necessary).
                If (decompressedSize < 0) Then
                    Const ARBITRARY_INIT_SIZE As Long = 16000000
                    decompressedSize = ARBITRARY_INIT_SIZE
                End If
                
                'Attempt to decode the .gz stream into a local array, and keep attempting until we have a big enough
                ' buffer to hold the result (or we run out of memory entirely).
                Dim actualDecompressedSize As Long
                Do
                    
                    'Prep a buffer to receive the file copy, then attempt decompression
                    ReDim m_tmpXcfCopy(0 To decompressedSize - 1) As Byte
                    
                    Const ERR_BUFF_TOO_SMALL As Long = 3
                    Dim ldResult As Long
                    ldResult = Plugin_libdeflate.Decompress_GZip(VarPtr(m_tmpXcfCopy(0)), decompressedSize, VarPtr(origBytes(0)), UBound(origBytes) + 1, actualDecompressedSize)
                    
                    decompressedSize = decompressedSize * 2
                    
                Loop While (ldResult = ERR_BUFF_TOO_SMALL) And (decompressedSize < 2 ^ 30)
                
                Const DECOMPRESSION_OK As Long = 0
                If (ldResult = DECOMPRESSION_OK) And (actualDecompressedSize > 0) Then
                    
                    If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "GZip decompression worked.  Proceeding normally..."
                        
                    'We decompressed the file!  Because we may have allocated a *lot* more space than we needed
                    ' (and we have a bunch of parsing to do that's gonna accumulate even more memory), reduce the
                    ' size of the placeholder array as small as possible.
                    If (UBound(m_tmpXcfCopy) <> actualDecompressedSize - 1) Then ReDim Preserve m_tmpXcfCopy(0 To actualDecompressedSize - 1) As Byte
                    
                    'Attempt to start a stream on the temporary array, then let parsing proceed normally from there.
                    streamOK = m_Stream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, actualDecompressedSize, VarPtr(m_tmpXcfCopy(0)))
                    
                Else
                    If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "GZip decompression failed; file parsing abandoned."
                End If
                
            End If
            
        Else
            streamOK = m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename)
        End If
        
        If streamOK Then
            
            'The first 9 bytes of an XCF file must be the ASCII values "gimp xcf "
            potentiallyXCF = (m_Stream.ReadString_ASCII(9) = "gimp xcf ")
            If (potentiallyXCF And XCF_DEBUG_VERBOSE) Then PDDebug.LogAction "Valid XCF file found"
            
        End If
        
    End If
    
    IsFileXCF = potentiallyXCF
    If (Not IsFileXCF) Or (Not onSuccessLeaveStreamOpen) Then
        Set m_Stream = Nothing
        Erase m_tmpXcfCopy
    End If
    
End Function

'This function is ONLY VALID AFTER A SUCCESSFUL CALL TO LoadXCF_FromFile!
Friend Function GetICCProfile() As pdICCProfile
    Set GetICCProfile = m_Profile
End Function

'This function is ONLY VALID AFTER A SUCCESSFUL CALL TO LoadXCF_FromFile!
Friend Function GetOriginalAlphaState() As Boolean
    If (m_numOfLayers > 1) Then
        GetOriginalAlphaState = True
    Else
        GetOriginalAlphaState = (m_Layers(0).lColorMode = xcf_GrayA) Or (m_Layers(0).lColorMode = xcf_IndexedA) Or (m_Layers(0).lColorMode = xcf_RGBA)
    End If
End Function

'This function is ONLY VALID AFTER A SUCCESSFUL CALL TO LoadXCF_FromFile!
Friend Function GetOriginalColorDepth() As Long
    If (m_imageColorMode = xcf_Indexed) Then
        GetOriginalColorDepth = 8
    Else
    
        Select Case m_imagePrecision
            Case xcf_08bitIntLinear
                GetOriginalColorDepth = 8
            Case xcf_08bitIntGamma
                GetOriginalColorDepth = 8
            Case xcf_16bitIntLinear
                GetOriginalColorDepth = 16
            Case xcf_16bitIntGamma
                GetOriginalColorDepth = 16
            Case xcf_32bitIntLinear
                GetOriginalColorDepth = 32
            Case xcf_32bitIntGamma
                GetOriginalColorDepth = 32
            Case xcf_16bitFltLinear
                GetOriginalColorDepth = 16
            Case xcf_16bitFltGamma
                GetOriginalColorDepth = 16
            Case xcf_32bitFltLinear
                GetOriginalColorDepth = 32
            Case xcf_32bitFltGamma
                GetOriginalColorDepth = 32
            Case xcf_64bitFltLinear
                GetOriginalColorDepth = 64
            Case xcf_64bitFltGamma
                GetOriginalColorDepth = 64
        End Select
        
        If (m_imageColorMode = xcf_RGB) Then
            If Me.GetOriginalAlphaState Then GetOriginalColorDepth = GetOriginalColorDepth * 4 Else GetOriginalColorDepth = GetOriginalColorDepth * 3
        End If
        
    End If
    
End Function

'This function is ONLY VALID AFTER A SUCCESSFUL CALL TO LoadXCF_FromFile!
Friend Function GetOriginalDPI() As Single
    GetOriginalDPI = m_imageResolutionPPI
End Function

'This function is ONLY VALID AFTER A SUCCESSFUL CALL TO LoadXCF_FromFile!
Friend Function HasGrayscale() As Boolean
    HasGrayscale = (m_imageColorMode = xcf_Grayscale)
End Function

'Validate and load a candidate XCF file
Friend Function LoadXCF_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    LoadXCF_FromFile = False
    
    'Validate the file
    If Me.IsFileXCF(srcFile, False, True) Then
        
        'Validation only checks the first 9-bytes of the file for a magic ASCII string.
        ' If we're still here, that validation string passed. (We can still reject the file
        ' if the header contains invalid members.)
        LoadXCF_FromFile = Import_Stage1_ParseHeader(srcFile, dstImage, dstDIB)
        If (Not LoadXCF_FromFile) Then Exit Function
        
        'Still here?  Header looks okay.  Time to proceed with image property retrieval!
        LoadXCF_FromFile = Import_Stage2_LoadProps(srcFile, dstImage, dstDIB)
        If (Not LoadXCF_FromFile) Then Exit Function
        
        'Image properties can't really fail (short of a grossly malformed file), but we'll quickly
        ' know if the file is okay in the next step: grabbing all offsets for layer and channel blocks.
        ' (Note that this next stage will also fail if the source file is 2+ GB in size.)
        LoadXCF_FromFile = Import_Stage3_LoadLayersAndChannels(srcFile, dstImage, dstDIB)
        If (Not LoadXCF_FromFile) Then Exit Function
        
        'With all offsets retrieved, it's now time to use those offsets to traverse actual layer data.
        LoadXCF_FromFile = Import_Stage4_ParseLayerHeaders(srcFile, dstImage, dstDIB)
        If (Not LoadXCF_FromFile) Then Exit Function
        
        'Final step is assembling a complete pdImage object from the retrieved layers!
        LoadXCF_FromFile = Import_Stage5_BuildImage(srcFile, dstImage, dstDIB)
        
    'No penalty on failed validation; exit immediately
    Else
        Exit Function
    End If

End Function

'Import stages follow.
' IMPORTANT: all Import-prefixed stages must be called in succession, and they must only be called
' from a parent LoadXCF_ function(s).  (These functions rely on correct stream alignment from previous steps,
' and they *will break* if called any other way.

'Import step 5: with all XCF file data parsed, it's time to assemble a finished pdImage object!
Private Function Import_Stage5_BuildImage(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "Import_Stage5_BuildImage"
    Import_Stage5_BuildImage = True
    
    On Error GoTo BrokenImage
    
    'We now have everything we need out of the XCF file.  Time to build a pdImage object!
    
    'Start with basics, like width/height
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    dstImage.Width = m_ImageWidth
    dstImage.Height = m_ImageHeight
    
    'Resolution TODO
    Dim tmpLayer As pdLayer
    
    'Assemble layers!  Note that we must iterate in reverse order to match XCF embedded order
    Dim i As Long
    For i = m_numOfLayers - 1 To 0 Step -1
        
        'Ask the parent image to create a blank layer for us
        Dim newLayerID As Long
        newLayerID = dstImage.CreateBlankLayer()
        Set tmpLayer = dstImage.GetLayerByID(newLayerID)
        
        'Initialize the layer, and hand off our local DIB to the pdLayer object for future management
        tmpLayer.InitializeNewLayer PDL_Image, m_Layers(i).lName, m_Layers(i).lDIB
        
        'The DIB, as it comes from GIMP, will *not* be premultiplied.  Premultiply alpha now.
        tmpLayer.GetLayerDIB.SetAlphaPremultiplication True
        
        'Fill in any remaining layer properties
        With m_Layers(i)
            If .lAlphaLocked Then tmpLayer.SetLayerAlphaMode AM_Locked Else tmpLayer.SetLayerAlphaMode AM_Normal
            tmpLayer.SetLayerVisibility .lIsVisible
            tmpLayer.SetLayerBlendMode .lBlendMode
            tmpLayer.SetLayerOpacity .lOpacity
            tmpLayer.SetLayerOffsetX .lOffsetX
            tmpLayer.SetLayerOffsetY .lOffsetY
            'TODO: layer masks
            'TODO: layer groups
        End With
        
    Next i
    
    'With all layers assembled, set the active layer (if any) before exiting.
    ' (Note also that XCF files can have multiple simultaneously active layers.  We just grab
    '  the first tagged one.)
    For i = 0 To m_numOfLayers - 1
        If m_Layers((m_numOfLayers - 1) - i).lIsActiveLayer Then
            dstImage.SetActiveLayerByIndex i
            Exit For
        End If
    Next i
    
    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage5_BuildImage = False
    
End Function

'Import step 4: parse layer headers (and properties).  Offsets to each layer block were retrieved in the
' previous step.  Now we must actually traverse those offsets and pull layer header data and property lists
' from each block.  These offsets *should* be in roughly sequential order matching the original layer order
' but the spec does not require this, so we may need to bounce around the file as we proceed.  (This is
' especially true for importing tile data, which uses a complex layout in XCF files.)
Private Function Import_Stage4_ParseLayerHeaders(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "Import_Stage4_ParseLayerHeaders"
    Import_Stage4_ParseLayerHeaders = True
    
    On Error GoTo BrokenImage
    
    'Modern GIMP files use 64-bit pointers; legacy use 32-bit
    Dim ptrIs64bit As Boolean, ptrFileH As Long
    ptrIs64bit = (m_xcfVersion >= 11)
    
    'Iterate each layer in turn
    Dim i As Long
    For i = 0 To m_numOfLayers - 1
        
        'Start by forcibly aligning the stream pointer to this layer's offset
        m_Stream.SetPosition m_Layers(i).ptrInFile, FILE_BEGIN
        
        'Layer headers are reasonably simple compared to other parts of XCF files.
        
        'Start by retrieving fixed per-layer data
        With m_Layers(i)
        
            'Dimensions and color-mode
            .lWidth = m_Stream.ReadLong_BE()
            .lHeight = m_Stream.ReadLong_BE()
            .lColorMode = m_Stream.ReadLong_BE()
            
            'Layer name is encoded as a "GIMP string", which are UTF-8 strings preceded by a 4-byte length
            Dim lenName As Long
            lenName = m_Stream.ReadLong_BE()    'this value *includes* a null-terminator, so string itself is [n-1] chars
            If (lenName > 0) Then .lName = m_Stream.ReadString_UTF8(lenName - 1)
            
            '+1 for null terminator
            m_Stream.SetPosition 1, FILE_CURRENT
            
            'Next is a standard list of GIMP properties.  This includes things like layer offset,
            ' opacity, visibility, blend mode, etc.  This is a variable-length list that we must iterate
            ' until we reach the special PROP_END marker (0 ID followed by 0 length).
            Const INIT_PROP_COUNT As Long = 16
            .lNumProperties = 0
            ReDim .lProperties(0 To INIT_PROP_COUNT - 1) As xcf_Property
            
            Dim propID As xcf_PropertyID
            propID = m_Stream.ReadLong_BE()
            
            Do While (propID <> xcf_PROP_END)
                    
                'Store this property ID, then retrieve length and payload
                If (.lNumProperties > UBound(.lProperties)) Then ReDim Preserve .lProperties(0 To .lNumProperties * 2 - 1) As xcf_Property
                .lProperties(.lNumProperties).propID = propID
                .lProperties(.lNumProperties).propSize = m_Stream.ReadLong_BE()
                If (.lProperties(.lNumProperties).propSize > 0) Then m_Stream.ReadBytes .lProperties(.lNumProperties).propData, .lProperties(.lNumProperties).propSize, True
                .lNumProperties = .lNumProperties + 1
                
                'Retrieve the next property ID
                propID = m_Stream.ReadLong_BE()
                
            Loop
            
            'The last ID we found was PROP_END.  It will still be followed by a 0-length marker.
            If (m_Stream.ReadLong_BE() <> 0) Then InternalError FUNC_NAME, "PROP_END had a non-zero payload"
            
            'Fix the property collection to its final size
            If (.lNumProperties < UBound(.lProperties) + 1) Then ReDim Preserve .lProperties(0 To .lNumProperties - 1) As xcf_Property
            
            'After the property list comes two important pointers (file offsets):
            ' 1) pointer to the pixel hierarchy, and...
            If ptrIs64bit Then ptrFileH = m_Stream.ReadLong_BE()
            .ptrToPixels = m_Stream.ReadLong_BE()
            
            ' 2) pointer to the layer mask (a channel structure, 0 if none exists)
            If ptrIs64bit Then ptrFileH = m_Stream.ReadLong_BE()
            .ptrToMask = m_Stream.ReadLong_BE()
            
            If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "Found layer: " & .lName & " - (" & .lWidth & "x" & .lHeight & "), " & .lNumProperties & " props and ptrs are " & .ptrToPixels & ", " & .ptrToMask
            
            'GIMP files are typically structured with interleaved layer and pixel data, so a (complex) hierarchy
            ' of pixel mapping structs will appear here, before the next layer.  To improve performance, we want
            ' to grab the full hierarchy of pixel mappings before moving the file pointer to the next layer block.
            If Import_Stage4_ParseLayerHeaders Then Import_Stage4_ParseLayerHeaders = Import_Stage4a_PrepTileRetrieval(.ptrToPixels, .lWidth, .lHeight, .lBytesPerPixel, .numTilesX, .numTilesY, .numTilesTotal, .ptrToTiles)
            
        End With
        
        'With all pixel mapping data retrieved, we now have enough information to retrieve and parse
        ' individual pixel tiles for this layer.  (GIMP encodes all images as 64x64 px tiles.)
        ' This next step will generate a ready-to-go layer DIB from the source GIMP tiles, with full
        ' support for all known color-depths and precision values.
        If Import_Stage4_ParseLayerHeaders Then Import_Stage4_ParseLayerHeaders = Import_Stage4b_GenerateLayerPixels(i, dstImage, dstDIB)
        
        'In the future, we could look at (gracefully?) recovering if one XCF layer is bad,
        ' but for now, let's abandon parsing if any individual layer is bad.
        ' (This simplifies testing parsing robustness.)
        If (Not Import_Stage4_ParseLayerHeaders) Then
            InternalError FUNC_NAME, "generate layer pixels failed"
            Exit Function
        End If
        
        'If this layer was constructed successfully, it's time to pull meaningful properties from the
        ' property list.  This includes things like layer offsets, blend modes, opacity, visibility, etc.
        Import_Stage4_ParseLayerHeaders = Import_Stage4c_DecodeLayerProps(i, dstImage, dstDIB)
        
        'If this layer has a mask, we also want to retrieve it before exiting.
        If (m_Layers(i).ptrToMask <> 0) Then Import_Stage4_ParseLayerHeaders = Import_Stage4d_DecodeLayerMask(i, dstImage, dstDIB)
        
        'Errors during property parsing are not critical; we will simply supply default properties
        ' at pdLayer generation time.
        If (Not Import_Stage4_ParseLayerHeaders) Then
            InternalError FUNC_NAME, "one or more layer props was broken"
            Import_Stage4_ParseLayerHeaders = True
        End If
        
        'This layer is now finished.  Proceed with the next layer (if any).
        
    'Continue with the next layer (and note that stream offset *does not matter* at this point;
    ' it will be forcibly reset by the next layer pass).
    Next i
    
    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage4_ParseLayerHeaders = False
    
End Function

'ONLY CALL THIS FUNCTION FROM Import_Stage4, above.
'
'Given a valid channel struct index, retrieve the "pixels" associated with that channel.  I put "pixels"
' in quotes because the channel may not map to pixels at all (e.g. a mask is really just a byte-stream
' that gets merged into pixel data).
'
'This function shares many attributes with the standard "GenerateLayerPixels" function, but there are
' necessary differences because this function produces a byte array in linear space - unlike color data,
' which produces an interleaved image that is (almost) always in some gamma-corrected space.  Because of
' these differences, we can't rely on things like an embedded color profile or lcms-2 transforms.
' Rather than clutter up the core decoder with these variances, I have moved this channel-specific
' decoder here.  (At present, it is only used for layer masks, but other channels could be retrieved
' as-is - we would just need to find a use for them in PD!)=
'
'Incoming *and* outgoing stream pointer alignment is *not* guaranteed.  This function will move
' the pointer around as necessary to decode individual tiles.  Pointer alignment, if required,
' must be handled by the caller.
Private Function Import_Stage4e_GenerateSingleChannel(ByRef srcChannel As xcf_Channel) As Boolean

    Const FUNC_NAME As String = "Import_Stage4e_GenerateSingleChannel"
    Import_Stage4e_GenerateSingleChannel = False
    
    On Error GoTo BrokenImage
    
    'If we've made it this far, previous functions successfully retrieved a list of file offsets
    ' for the individual 64x64 tiles that comprise this channel.  It is now our job to decompress
    ' those tiles, and from the results, construct a useable channel bytestream.
    If (srcChannel.cWidth <= 0) Or (srcChannel.cHeight <= 0) Then
        InternalError FUNC_NAME, "bad channel dimensions"
        Import_Stage4e_GenerateSingleChannel = False
        Exit Function
    End If
    
    'For a normal "layer", we paint our finished results into a pdDIB object.  For a standalone channel, however,
    ' we instead want to "paint" our results into a final channel byte array.  We still need a placeholder
    ' 64x64 "tile" to hold the current tile as it's decoded - but that tile will ultimately just get copied
    ' line-by-line into ito the parent channel array.
    Dim bTile() As Byte
    ReDim bTile(0 To 63, 0 To 63) As Byte
    
    ReDim srcChannel.cBytes(0 To srcChannel.cWidth - 1, 0 To srcChannel.cHeight - 1) As Byte
    
    'We need a ton of information from previous steps (like color-depth and compression mode)
    ' to inform the processing of tile data.  We also need to be careful with right-most and
    ' bottom-most tiles, because they are unlikely to have full 64x64 px dimensions.
    Dim xTileMax As Long, yTileMax As Long
    xTileMax = srcChannel.numTilesX - 1
    yTileMax = srcChannel.numTilesY - 1
    
    Dim idxTile As Long, numTilesTotal As Long
    idxTile = 0
    numTilesTotal = srcChannel.numTilesTotal
    
    'Tiles in channels use the same compression as layer data
    Dim idxProp As Long
    idxProp = GetIndexOfProperty(xcf_PROP_COMPRESSION, m_ImageProperties)
    
    Dim srcCompression As xcf_Compression
    If (idxProp >= 0) Then
        srcCompression = m_ImageProperties(idxProp).propData(0)
    Else
        srcCompression = xcf_Compress_RLE
    End If
    If (srcCompression < 0) Or (srcCompression > xcf_Compress_ZLib) Then srcCompression = xcf_Compress_RLE
    
    'Unlike layer pixel data, we do not need a complex support matrix here.  All channels encode "grayscale" data,
    ' with no additional modifiers (like alpha).
    
    'Note, however, that while the XCF spec constantly refers to channels in terms of "bytes", layer masks
    ' are actually encoded with the same precision as their parent layer.  Thus we *do* need to account for
    ' integer vs float precision, and 8/16/32/64-bit sizes.
    Dim precisionInt As Boolean, precisionFloat As Boolean
    precisionInt = (m_imagePrecision = xcf_08bitIntGamma) Or (m_imagePrecision = xcf_08bitIntLinear) Or (m_imagePrecision = xcf_16bitIntGamma) Or (m_imagePrecision = xcf_16bitIntLinear) Or (m_imagePrecision = xcf_32bitIntGamma) Or (m_imagePrecision = xcf_32bitIntLinear)
    precisionFloat = (Not precisionInt)
    
    Dim numBitsPerChannel As Long
    numBitsPerChannel = GetNumBitsFromImagePrecision()
    
    Dim numBytesPerChannel As Long
    numBytesPerChannel = numBitsPerChannel \ 8
    
    'Pre-calculate a "worst-case" source size.  This formula is arbitrary; maybe there is a better way,
    ' but I derived this from the following clue in the XCF spec:
    '
    '"The RLE encoding can cause degenerated encodings in which the original data stream may double in size
    ' (or grow to arbitrarily large sizes if (128,0,0) operations are inserted). Such encodings must be avoided,
    ' as GIMP's XCF reader expects that the size of an encoded tile is never more than 24 KB, which is only 1.5
    ' times the unencoded size of a 64x64 RGBA tile.
    '
    'Based on this advice, I've gone with a similar formula of "50% overhead for worst-case compression".
    Dim srcSize As Long, maxSrcSize As Long, srcBytes() As Byte
    maxSrcSize = Int((64 * 64 * numBytesPerChannel) * 1.5 + 0.5)
    
    'Because we'll potentially be swapping endianness, ensure it's a solid multiple of 8
    maxSrcSize = ((maxSrcSize + 7) \ 8) * 8
    ReDim srcBytes(0 To maxSrcSize - 1) As Byte
    
    'Decompressed destination size is fixed, thankfully, but note that not all tiles will use the full
    ' tile size.  Right- and bottom-most tiles may be smaller if the image dimensions are not a clean
    ' multiple of 64.
    '
    'Note also that we declare this array as "byte" but it may actually contain shorts, floats, etc -
    ' we'll alias it as necessary when the time comes.
    Dim dstSize As Long, maxDstSize As Long, dstBytes() As Byte
    maxDstSize = 64 * 64 * numBytesPerChannel
    ReDim dstBytes(0 To maxDstSize - 1) As Byte
    
    'Various HDR formats require extra handling because they use silly big-endian encoding for floats.
    ' Because we are not using LCMS in this function, we want an additional swizzling buffer available
    ' for HDR formats, at the same size as the decompressed pixel data.
    Dim tmpSwizzle() As Byte
    If (numBitsPerChannel > 8) Then ReDim tmpSwizzle(0 To maxDstSize - 1) As Byte
    
    'Precalculate dimensions of the last row/column of tiles; this saves us some computation effort
    ' inside the tile loop.
    Dim tileWidth As Long, tileHeight As Long, numTilePixels As Long
    
    Dim tileWidthLast As Long, tileHeightLast As Long
    tileWidthLast = srcChannel.cWidth - ((srcChannel.cWidth \ 64) * 64)
    If (tileWidthLast <= 0) Then tileWidthLast = 64
    tileHeightLast = srcChannel.cHeight - ((srcChannel.cHeight \ 64) * 64)
    If (tileHeightLast <= 0) Then tileHeightLast = 64
    
    'As usual, HDR data requires more complex handling.  We will manually convert data to little-endian format,
    ' but after that we still need to deal with int vs float precision, including coverage for halfs (which VB6
    ' doesn't provide natively, argh).
    
    'LittleCMS is very helpful here, but note that we need to approach it differently because this is *not*
    ' a true color-management transform - so we need to take care to stick to linear transforms only.
    Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile
    Dim srcTransform As pdLCMSTransform
    
    Set srcProfile = New pdLCMSProfile
    Set dstProfile = New pdLCMSProfile
    
    If (numBitsPerChannel > 8) Then
        srcProfile.CreateGenericGrayscaleProfile 1#
        dstProfile.CreateGenericGrayscaleProfile 1#
    End If
    
    Dim srcFormat As LCMS_PIXEL_FORMAT
    
    'Start iterating tiles
    Dim xTile As Long, yTile As Long
    For yTile = 0 To yTileMax
    For xTile = 0 To xTileMax
        
        'PDDebug.LogAction "Parsing tile (" & xTile & ", " & yTile & ")"
        
        'Unfortunately, GIMP does not encode compressed tile sizes into the XCF file.  We can infer the
        ' size for most tiles (by subtracting this tile's file offset from the next tile's offset),
        ' but for the final tile in the image, we just have to grab a "worst-case" chunk of bytes from the
        ' file and hope for the best.
        If (idxTile < numTilesTotal - 1) Then
            srcSize = srcChannel.ptrToTiles(idxTile + 1) - srcChannel.ptrToTiles(idxTile)
        Else
            srcSize = maxSrcSize
        End If
        
        'Destination size is fixed, and contingent only on the current tile's dimensions and the color-depth
        ' of the image.
        If (xTile < xTileMax) Then tileWidth = 64 Else tileWidth = tileWidthLast
        If (yTile < yTileMax) Then tileHeight = 64 Else tileHeight = tileHeightLast
        numTilePixels = tileWidth * tileHeight
        dstSize = numTilePixels * numBytesPerChannel
        
        'Forcibly align the file stream pointer for this tile, then pull the required amount of source bytes
        ' into a local array.  (Manual RLE decoding is faster this way, vs pulling byte-by-byte from the file.)
        m_Stream.SetPosition srcChannel.ptrToTiles(idxTile), FILE_BEGIN
        
        '(PD's stream object automatically protects against OOB reads, but I still prefer to check for EOF and
        ' shrink the source buffer accordingly, especially since XCF requires such a huge "safety" margin for
        ' RLE-compressed tiles.)
        If (srcSize > m_Stream.GetStreamSize - m_Stream.GetPosition) Then srcSize = m_Stream.GetStreamSize - m_Stream.GetPosition
        srcSize = m_Stream.ReadBytesToBarePointer(VarPtr(srcBytes(0)), srcSize)
        
        'We are now done with the file stream and can happily ignore it until the next tile.
        
        'PDDebug.LogAction "Read " & srcSize & " bytes from position " & m_Layers(idxLayer).ptrToTiles(idxTile)
        
        'What happens next varies according to compression mode.  For DEFLATE, for example, we can just
        ' decompress as-is.  For RLE, however, we need to manually handle decompression.
        Select Case srcCompression
            
            'Shouldn't exist in the wild; will revisit if I can find a way to produce such a file
            Case xcf_Compress_None
                
            'Default setting for XCF files; after RLE decompression, HDR images require unshuffling
            ' (since RLE operates on individual bytes, not channels)
            Case xcf_Compress_RLE
                DecompressRLE numBytesPerChannel, numTilePixels, srcBytes, dstBytes
                If (numBitsPerChannel > 8) Then UnshuffleHDRChannels numBitsPerChannel, numTilePixels, False, True, tmpSwizzle, dstBytes, maxDstSize
                
            'Available as an optional setting in modern XCF files, but rare "in the wild"
            Case xcf_Compress_ZLib
                
                'zLib compression is straightforward: decode the source bytes directly into the destination array,
                ' using a pre-calculated destination size as our fixed indicator for "completion".  (Note that we
                ' may still need to swap endianness after decompression, but we at least won't have to deal with
                ' planar-to-interleaved conversion.)
                Dim dcmpResult As Long
                dcmpResult = Plugin_libdeflate.Decompress_ZLib(VarPtr(dstBytes(0)), dstSize, VarPtr(srcBytes(0)), srcSize)
                If (dcmpResult <> 0) Then InternalError FUNC_NAME, "bad decompress: " & dcmpResult
                
        End Select
        
        'With the tile successfully decompressed, we now need to generate an 8-bpp copy of it.  For 8-bpp data,
        ' we're pretty much done!  HDR images require extra work here, however.
        
        'Manually handle BE to LE conversion for HDR float formats (and note that 32-bit integer will be forcibly
        ' downsampled to 16-bit, because 32-bit integers are not supported natively supported by LCMS - something that
        ' doesn't matter for masks, but which happens "for free" because we reuse code from the layer parser).
        SwapEndiannessHDR srcCompression, numBitsPerChannel, precisionFloat, True, numTilePixels, False, tmpSwizzle, dstBytes, maxDstSize
        
        'dstBytes() now contains the final tile bytes, in proper scanline order.  HDR formats need to be
        ' forcibly downsampled to 8-bit values, but if the data is already in 8-bit format, we can just
        ' copy it as-is out to the final byte stream!
        Dim x As Long, y As Long
        If (numBitsPerChannel = 8) Then
            For y = 0 To tileHeight - 1
                CopyMemoryStrict VarPtr(srcChannel.cBytes(xTile * 64, yTile * 64 + y)), VarPtr(dstBytes(y * tileWidth)), tileWidth
            Next y
        
        'HDR formats...
        Else
            
            'Integer data must be 16-bit (because 8-bit was already handled, above, and 32-bit is forcibly downsampled
            ' to 16-bit in a previous step).
            If precisionInt Then
                
                'We could use LCMS here, but honestly it's just as fast to assign the bytes ourselves
                ' (especially given the issues with data locality).
                For y = 0 To tileHeight - 1
                    For x = 0 To tileWidth - 1
                        srcChannel.cBytes(xTile * 64 + x, yTile * 64 + y) = dstBytes(y * tileWidth * 2 + x * 2)
                    Next x
                Next y
            
            'Floating-point may as well be done via LCMS, since it simplifies dealing with unsupported formats
            ' like halfs
            Else
                
                If (numBitsPerChannel = 16) Then
                    srcFormat = TYPE_GRAY_HALF_FLT
                ElseIf (numBitsPerChannel = 32) Then
                    srcFormat = TYPE_GRAY_FLT
                Else
                    srcFormat = TYPE_GRAY_DBL
                End If
                
                'Render to the *temporary placeholder tile*
                Set srcTransform = New pdLCMSTransform
                srcTransform.CreateTwoProfileTransform srcProfile, dstProfile, srcFormat, TYPE_GRAY_8, INTENT_ABSOLUTE_COLORIMETRIC, 0&
                srcTransform.ApplyTransformToArbitraryMemoryEx VarPtr(dstBytes(0)), VarPtr(bTile(0, 0)), tileWidth, tileHeight, (numBitsPerChannel \ 8) * tileWidth, 64, (numBitsPerChannel \ 8) * tileWidth * tileHeight, 0&
                
                'We now need to render the temporary placeholder tile to its actual location in the fully assembled mask
                For y = 0 To tileHeight - 1
                    CopyMemoryStrict VarPtr(srcChannel.cBytes(xTile * 64, yTile * 64 + y)), VarPtr(bTile(0, y)), tileWidth
                Next y

            End If
            
        End If
        
        'Increment our pointer into the source tile offset table
        idxTile = idxTile + 1
        
    Next xTile
    Next yTile
    
    Import_Stage4e_GenerateSingleChannel = True
    
    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage4e_GenerateSingleChannel = False
    
End Function

'ONLY CALL THIS FUNCTION FROM Import_Stage4, above.
'
'Given a valid layer index with a non-zero mask pointers, retrieve the associated mask and store it inside a
' layer-level pdLayerMask object.
'
'Note that this function will move the file stream pointer.  The caller *must* correctly reposition the
' stream pointer, if desired, when this function returns.
Private Function Import_Stage4d_DecodeLayerMask(ByVal idxLayer As Long, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    Const FUNC_NAME As String = "Import_Stage4d_DecodeLayerMask"
    Import_Stage4d_DecodeLayerMask = True
    
    On Error GoTo BrokenImage
    
    'Failsafe check for a valid layer pointer for the requested index
    If (m_Layers(idxLayer).ptrToMask = 0) Then
        Import_Stage4d_DecodeLayerMask = False
        Exit Function
    End If
    
    'Modern GIMP files use 64-bit pointers; legacy use 32-bit
    Dim ptrIs64bit As Boolean, ptrFileH As Long
    ptrIs64bit = (m_xcfVersion >= 11)
    
    'Retrieving a layer mask is much like retrieving the layer itself, but with slightly different structs
    ' (because we are retrieving a channel, not a layer).
    
    'Start by moving the file pointer to the mask offset
    m_Stream.SetPosition m_Layers(idxLayer).ptrToMask, FILE_BEGIN
    
    'The stream now (theoretically) points at a channel structure describing the mask.  Retrieve the various members,
    ' and validate as-we-go.  (This is helpful because if attributes like mask width/height don't match the parent
    ' layer's width/height, we know the file is corrupt.)
    m_Layers(idxLayer).lMask.cWidth = m_Stream.ReadLong_BE()
    m_Layers(idxLayer).lMask.cHeight = m_Stream.ReadLong_BE()
    
    'Validate dimensions before proceeding
    If (m_Layers(idxLayer).lMask.cWidth <> m_Layers(idxLayer).lWidth) Or (m_Layers(idxLayer).lMask.cHeight <> m_Layers(idxLayer).lHeight) Then
        InternalError FUNC_NAME, "mask and layer dimensions don't match"
        Import_Stage4d_DecodeLayerMask = False
        Exit Function
    End If
    
    'Still here?  Channel name comes next (this does not have relevance in PD, but we retrieve it anyway).
    
    'Layer name is encoded as a "GIMP string", which are UTF-8 strings preceded by a 4-byte length
    Dim lenName As Long
    lenName = m_Stream.ReadLong_BE()    'this value *includes* a null-terminator, so string itself is [n-1] chars
    If (lenName > 0) Then m_Layers(idxLayer).lMask.cName = m_Stream.ReadString_UTF8(lenName - 1)
    
    '+1 for null terminator
    m_Stream.SetPosition 1, FILE_CURRENT
    
    'Next is a standard list of GIMP properties.  These mostly include UI-type attributes specific to GIMP,
    ' but we retrieve them "just in case" there is useful data in the collection.  Note that this is a
    ' variable-length list that we must iterate until we reach the special PROP_END marker (0 ID followed
    ' by 0 length).
    Const INIT_PROP_COUNT As Long = 16
    m_Layers(idxLayer).lMask.cNumProperties = 0
    ReDim m_Layers(idxLayer).lMask.cProperties(0 To INIT_PROP_COUNT - 1) As xcf_Property
    
    Dim propID As xcf_PropertyID
    propID = m_Stream.ReadLong_BE()
    
    With m_Layers(idxLayer).lMask
        
        Do While (propID <> xcf_PROP_END)
                
            'Store this property ID, then retrieve length and payload
            If (.cNumProperties > UBound(.cProperties)) Then ReDim Preserve .cProperties(0 To .cNumProperties * 2 - 1) As xcf_Property
            .cProperties(.cNumProperties).propID = propID
            .cProperties(.cNumProperties).propSize = m_Stream.ReadLong_BE()
            If (.cProperties(.cNumProperties).propSize > 0) Then m_Stream.ReadBytes .cProperties(.cNumProperties).propData, .cProperties(.cNumProperties).propSize, True
            .cNumProperties = .cNumProperties + 1
            
            'Retrieve the next property ID
            propID = m_Stream.ReadLong_BE()
            
        Loop
        
    End With
    
    'The last ID we found was PROP_END.  It will still be followed by a 0-length marker.
    If (m_Stream.ReadLong_BE() <> 0) Then InternalError FUNC_NAME, "PROP_END had a non-zero payload"
    
    'Fix the property collection to its final size
    If (m_Layers(idxLayer).lMask.cNumProperties < UBound(m_Layers(idxLayer).lMask.cProperties) + 1) Then ReDim Preserve m_Layers(idxLayer).lMask.cProperties(0 To m_Layers(idxLayer).lMask.cNumProperties - 1) As xcf_Property
    
    'After the property list comes a pointer (file offset) to the pixel hierarchy for this channel
    If ptrIs64bit Then ptrFileH = m_Stream.ReadLong_BE()
    Dim ptrToTileHierarchy As Long
    ptrToTileHierarchy = m_Stream.ReadLong_BE()
    
    'Parsing the full hierarchy is a complex job, so let's pass it off to a dedicated function.
    With m_Layers(idxLayer)
        Import_Stage4d_DecodeLayerMask = Import_Stage4a_PrepTileRetrieval(ptrToTileHierarchy, .lWidth, .lHeight, .lMask.cBytesPerPixel, .lMask.numTilesX, .lMask.numTilesY, .lMask.numTilesTotal, .lMask.ptrToTiles)
    End With
    
    'We now have pointers to every tile comprising this mask.  We now need to convert those tiles into a normal
    ' pixel stream (at PD's current bit-depth).  Note that this step must *not* be color-managed, because mask
    ' data is always linear (unlike color data, which is almost always in a gamma-corrected space).
    If Import_Stage4d_DecodeLayerMask Then Import_Stage4d_DecodeLayerMask = Import_Stage4e_GenerateSingleChannel(m_Layers(idxLayer).lMask)
    
    'If the mask was decoded successfully, merge it down now.  (In the future, we will transfer the mask to a
    ' dedicated mask object.)
    If Import_Stage4d_DecodeLayerMask Then
        
        Dim layerWidth As Long, layerHeight As Long
        layerWidth = m_Layers(idxLayer).lDIB.GetDIBWidth
        layerHeight = m_Layers(idxLayer).lDIB.GetDIBHeight
        
        Dim tmpBytes() As Byte, tmpSA As SafeArray1D
        m_Layers(idxLayer).lDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
        
        Dim scanStart As Long, scanWidth As Long
        scanStart = tmpSA.pvData
        scanWidth = m_Layers(idxLayer).lDIB.GetDIBStride
        
        Dim xLoopEnd As Long
        xLoopEnd = layerWidth - 1
        
        Const ONE_DIV_255 As Single = 1! / 255!
        
        Dim tmpSingle As Single
        Dim x As Long, y As Long, ySrcOffset As Long
        For y = 0 To layerHeight - 1
            tmpSA.pvData = scanStart + (scanWidth * y)
            ySrcOffset = y * layerWidth
        For x = 0 To xLoopEnd
            tmpSingle = CSng(m_Layers(idxLayer).lMask.cBytes(x, y)) * ONE_DIV_255
            tmpBytes(x * 4 + 3) = Int(CSng(tmpBytes(x * 4 + 3)) * tmpSingle)
        Next x
        Next y
        
        m_Layers(idxLayer).lDIB.UnwrapArrayFromDIB tmpBytes
        
    End If
    
    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage4d_DecodeLayerMask = False
    
End Function

'ONLY CALL THIS FUNCTION FROM Import_Stage4, above.
'
'Given a valid layer index, iterate all properties retrieved from that layer and try to recover as many
' useful layer attributes as we can.
'
'Unlike other functions, this one does *not* touch the master stream pointer.  It only operates on the
' property bytes cached in the target layer object.
Private Function Import_Stage4c_DecodeLayerProps(ByVal idxLayer As Long, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    Const FUNC_NAME As String = "Import_Stage4c_DecodeLayerProps"
    Import_Stage4c_DecodeLayerProps = True
    
    On Error GoTo BrokenImage
    
    'This step is pretty simple: iterate the property list and parse any meaningful properties,
    ' then translate them into their nearest PhotoDemon equivalent.
    
    'Before decoding anything, set default layer properties in advance.  (Some properties are not mandatory,
    ' and we want to assume sane defaults for these.)
    With m_Layers(idxLayer)
        .lOpacity = 100!
        .lOffsetX = 0
        .lOffsetY = 0
        .lBlendMode = BM_Normal
        .lAlphaLocked = False
        .lIsGroupMarker = False
        .lHasMask = False
        .lIsActiveLayer = False
        .lIsVisible = True
    End With
    
    Dim tmpInt As Long, tmpFloat As Single
    
    'Traversing individual property byte streams is much easier c/o pdStream
    Dim tmpStream As pdStream
    Set tmpStream = New pdStream
    
    Dim i As Long
    For i = 0 To m_Layers(idxLayer).lNumProperties - 1
        
        'As a convenience, always point the temporary stream at this property's bytes;
        ' there's no harm if we don't end up using it.
        If (m_Layers(idxLayer).lProperties(i).propSize > 0) Then tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, startingBufferSize:=m_Layers(idxLayer).lProperties(i).propSize, baseFilePointerOffset:=VarPtr(m_Layers(idxLayer).lProperties(i).propData(0))
        
        Select Case m_Layers(idxLayer).lProperties(i).propID
            
            'Legacy opacity can appear in two ways: a legacy indicator (using [0, 255] int scale),
            ' or a modern indicator ([0.0, 1.0] float).  The legacy indicator must always appear
            ' first in the file; the modern indicator must appear later (and is optional).
            Case xcf_PROP_OPACITY
                If (m_Layers(idxLayer).lProperties(i).propSize >= 4) Then
                    tmpInt = tmpStream.ReadLong_BE()
                    If (tmpInt < 0) Then tmpInt = 0
                    If (tmpInt > 255) Then tmpInt = 255
                    m_Layers(idxLayer).lOpacity = tmpInt / 2.55!
                End If
                
            Case xcf_PROP_FLOAT_OPACITY
                If (m_Layers(idxLayer).lProperties(i).propSize >= 4) Then
                    tmpFloat = tmpStream.ReadFloat_BE()
                    If (tmpFloat < 0!) Then tmpFloat = 0!
                    If (tmpFloat > 1!) Then tmpFloat = 1!
                    m_Layers(idxLayer).lOpacity = tmpFloat * 100!
                End If
                
            Case xcf_PROP_VISIBLE
                If (m_Layers(idxLayer).lProperties(i).propSize >= 4) Then
                    tmpInt = tmpStream.ReadLong_BE()
                    m_Layers(idxLayer).lIsVisible = (tmpInt <> 0)
                End If
            
            'Can appear in multiple layers!
            Case xcf_PROP_ACTIVE_LAYER
                m_Layers(idxLayer).lIsActiveLayer = True
            
            'Does not guarantee presence of mask
            Case xcf_PROP_APPLY_MASK
                If (m_Layers(idxLayer).lProperties(i).propSize >= 4) Then
                    tmpInt = tmpStream.ReadLong_BE()
                    m_Layers(idxLayer).lMaskActive = (tmpInt <> 0)
                End If
            
            Case xcf_PROP_GROUP_ITEM
                m_Layers(idxLayer).lIsGroupMarker = True
            
            'Critical (?) for group handling, but forensics will be necessary to figure out interpretation
            Case xcf_PROP_ITEM_PATH
                'uint32  30       Type identification
                'uint32  plength  Total length of the following payload in bytes
                'item-path        List of pointers, represented as uint32 values
                '
                'PROP_ITEM_PATH indicates the path of the layer if inside a group,
                'i.e. its position within the group (last element of the list), but
                'also the position of the group itself within its own level, up to the
                'top-level position (first element).

            Case xcf_PROP_LOCK_ALPHA
                If (m_Layers(idxLayer).lProperties(i).propSize >= 4) Then
                    tmpInt = tmpStream.ReadLong_BE()
                    m_Layers(idxLayer).lAlphaLocked = (tmpInt <> 0)
                End If
            
            'Blend mode is ugly, alas; GIMP supplies a huge amount of legacy vs modern blend-modes,
            ' and we just try to map them as best we can.
            Case xcf_PROP_BLEND_MODE
                If (m_Layers(idxLayer).lProperties(i).propSize >= 4) Then
                    tmpInt = tmpStream.ReadLong_BE()
                    m_Layers(idxLayer).lBlendMode = GetPDBlendFromGIMPBlend(tmpInt)
                End If
                
            Case xcf_PROP_OFFSETS
                If (m_Layers(idxLayer).lProperties(i).propSize >= 8) Then
                    m_Layers(idxLayer).lOffsetX = tmpStream.ReadLong_BE()
                    m_Layers(idxLayer).lOffsetY = tmpStream.ReadLong_BE()
                End If
                
        End Select
        
    Next i

    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage4c_DecodeLayerProps = False
    
End Function

'Private helper(s) for the layer property parser, above.
Private Function GetPDBlendFromGIMPBlend(ByVal srcGIMPBlend As Long) As PD_BlendMode
    
    Select Case srcGIMPBlend
        'Normal (Legacy)
        Case 0
            GetPDBlendFromGIMPBlend = BM_Normal
        'Dissolve (legacy) [random dithering to discrete alpha)
        Case 1
            GetPDBlendFromGIMPBlend = BM_Normal
        'Behind (legacy) [not selectable in the GIMP UI]
        Case 2
            GetPDBlendFromGIMPBlend = BM_Normal
        'Multiply (legacy)
        Case 3
            GetPDBlendFromGIMPBlend = BM_Multiply
        'Screen (legacy)
        Case 4
            GetPDBlendFromGIMPBlend = BM_Screen
        'Old broken Overlay
        Case 5
            GetPDBlendFromGIMPBlend = BM_Overlay
        'Difference (legacy)
        Case 6
            GetPDBlendFromGIMPBlend = BM_Difference
        'Addition (legacy)
        Case 7
            GetPDBlendFromGIMPBlend = BM_Normal
        'Subtract (legacy)
        Case 8
            GetPDBlendFromGIMPBlend = BM_Subtract
        'Darken only (legacy)
        Case 9
            GetPDBlendFromGIMPBlend = BM_Darken
        'Lighten only (legacy)
        Case 10
            GetPDBlendFromGIMPBlend = BM_Lighten
        'Hue (HSV) (legacy)
        Case 11
            GetPDBlendFromGIMPBlend = BM_Hue
        'Saturation (HSV) (legacy)
        Case 12
            GetPDBlendFromGIMPBlend = BM_Saturation
        'Color (HSL) (legacy)
        Case 13
            GetPDBlendFromGIMPBlend = BM_Color
        'Value (HSV) (legacy)
        Case 14
            GetPDBlendFromGIMPBlend = BM_Luminosity
        'Divide (Legacy)
        Case 15
            GetPDBlendFromGIMPBlend = BM_Divide
        'Dodge (Legacy)
        Case 16
            GetPDBlendFromGIMPBlend = BM_ColorDodge
        'Burn (Legacy)
        Case 17
            GetPDBlendFromGIMPBlend = BM_ColorBurn
        'Hard light(Legacy)
        Case 18
            GetPDBlendFromGIMPBlend = BM_HardLight
        'Soft light(Legacy)
        Case 19
            GetPDBlendFromGIMPBlend = BM_SoftLight
        'Grain Extract(Legacy)
        Case 20
            GetPDBlendFromGIMPBlend = BM_GrainExtract
        'Grain Merge(Legacy)
        Case 21
            GetPDBlendFromGIMPBlend = BM_GrainMerge
        'Color erase (legacy)
        Case 22
            GetPDBlendFromGIMPBlend = BM_Erase
        'Overlay
        Case 23
            GetPDBlendFromGIMPBlend = BM_Overlay
        'Hue (LCH)
        Case 24
            GetPDBlendFromGIMPBlend = BM_Hue
        'chroma (LCH)
        Case 25
            GetPDBlendFromGIMPBlend = BM_Saturation
        'Color (LCH)
        Case 26
            GetPDBlendFromGIMPBlend = BM_Color
        'Lightness (LCH)
        Case 27
            GetPDBlendFromGIMPBlend = BM_Luminosity
        'Normal
        Case 28
            GetPDBlendFromGIMPBlend = BM_Normal
        'Behind
        Case 29
            GetPDBlendFromGIMPBlend = BM_Behind
        'Multiply
        Case 30
            GetPDBlendFromGIMPBlend = BM_Multiply
        'Screen
        Case 31
            GetPDBlendFromGIMPBlend = BM_Screen
        'Difference
        Case 32
            GetPDBlendFromGIMPBlend = BM_Difference
        'Addition
        Case 33
            GetPDBlendFromGIMPBlend = BM_Normal
        'Subtract
        Case 34
            GetPDBlendFromGIMPBlend = BM_Subtract
        'Darken only
        Case 35
            GetPDBlendFromGIMPBlend = BM_Darken
        'Lighten only
        Case 36
            GetPDBlendFromGIMPBlend = BM_Lighten
        'Hue (HSV)
        Case 37
            GetPDBlendFromGIMPBlend = BM_Hue
        'Saturation (HSV)
        Case 38
            GetPDBlendFromGIMPBlend = BM_Saturation
        'Color (HSL)
        Case 39
            GetPDBlendFromGIMPBlend = BM_Color
        'Value (HSV)
        Case 40
            GetPDBlendFromGIMPBlend = BM_Luminosity
        'Divide
        Case 41
            GetPDBlendFromGIMPBlend = BM_Divide
        'Dodge
        Case 42
            GetPDBlendFromGIMPBlend = BM_ColorDodge
        'Burn
        Case 43
            GetPDBlendFromGIMPBlend = BM_ColorBurn
        'Hard light
        Case 44
            GetPDBlendFromGIMPBlend = BM_HardLight
        'Soft light
        Case 45
            GetPDBlendFromGIMPBlend = BM_SoftLight
        'Grain Extract
        Case 46
            GetPDBlendFromGIMPBlend = BM_GrainExtract
        'Grain Merge
        Case 47
            GetPDBlendFromGIMPBlend = BM_GrainMerge
        'Vivid light
        Case 48
            GetPDBlendFromGIMPBlend = BM_VividLight
        'Pin light
        Case 49
            GetPDBlendFromGIMPBlend = BM_PinLight
        'Linear light
        Case 50
            GetPDBlendFromGIMPBlend = BM_LinearLight
        'Hard mix
        Case 51
            GetPDBlendFromGIMPBlend = BM_HardMix
        'Exclusion
        Case 52
            GetPDBlendFromGIMPBlend = BM_Exclusion
        'Linear burn
        Case 53
            GetPDBlendFromGIMPBlend = BM_LinearBurn
        'Luma/Luminance darken only
        Case 54
            GetPDBlendFromGIMPBlend = BM_Darken
        'Luma/Luminance lighten only
        Case 55
            GetPDBlendFromGIMPBlend = BM_Lighten
        'Luminance
        Case 56
            GetPDBlendFromGIMPBlend = BM_Luminosity
        'Color erase
        Case 57
            GetPDBlendFromGIMPBlend = BM_Erase
        'Erase
        Case 58
            GetPDBlendFromGIMPBlend = BM_Erase
        'Merge
        Case 59
            GetPDBlendFromGIMPBlend = BM_Normal
        'Split
        Case 60
            GetPDBlendFromGIMPBlend = BM_Normal
        'Pass-through
        Case 61
            GetPDBlendFromGIMPBlend = BM_Normal
        
    End Select
    
End Function

'ONLY CALL THIS FUNCTION FROM Import_Stage4, above.
'
'Given a valid layer index, retrieve the pixels associated with that layer.  This involves reading
' all tiles for the layer, decompressing them (multiple compression schemes are possible), converting
' from planar to interleaved formats, swapping endianness on HDR schemes, manually translating
' unsupported formats like half-floats to VB-compatible ones, color-managing the resulting data from
' a huge matrix of potential integer and/or floating-point representations to standard 32-bit sRGB,
' then assembling a useable layer pixel stream from all that data.  FML!
'
'Incoming *and* outgoing stream pointer alignment is *not* guaranteed.  This function will move
' the pointer around as necessary to decode individual tiles.  Pointer alignment, if required,
' must be handled by the caller.
Private Function Import_Stage4b_GenerateLayerPixels(ByVal idxLayer As Long, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    Const FUNC_NAME As String = "Import_Stage4b_GenerateLayerPixels"
    Import_Stage4b_GenerateLayerPixels = True
    
    On Error GoTo BrokenImage
    
    'If we've made it this far, previous functions successfully retrieved a list of file offsets
    ' for the individual 64x64 tiles that comprise this XCF file.  It is now our job to decompress
    ' those tiles, and from the results, construct a useable layer pixel stream.
    
    'Start by creating a blank tile object.  All tiles will first get "painted" into this placeholder.
    Set m_Tile = New pdDIB
    m_Tile.CreateBlank 64, 64, 32, vbBlack, 255
    
    'Next, create the target layer DIB.  The full-size layer image will be assembled tile-by-tile from
    ' the placeholder m_Tile object, above.
    Set m_Layers(idxLayer).lDIB = New pdDIB
    m_Layers(idxLayer).lDIB.CreateBlank m_Layers(idxLayer).lWidth, m_Layers(idxLayer).lHeight, 32, vbBlack, 255
    
    'We need a ton of information from previous steps (like color-depth and compression mode)
    ' to inform the processing of tile data.  We also need to be careful with right-most and
    ' bottom-most tiles, because they are unlikely to have full 64x64 px dimensions.
    Dim xTileMax As Long, yTileMax As Long
    xTileMax = m_Layers(idxLayer).numTilesX - 1
    yTileMax = m_Layers(idxLayer).numTilesY - 1
    
    Dim idxTile As Long, numTilesTotal As Long
    idxTile = 0
    numTilesTotal = m_Layers(idxLayer).numTilesTotal
    
    Dim idxProp As Long
    idxProp = GetIndexOfProperty(xcf_PROP_COMPRESSION, m_ImageProperties)
    
    'Determine compression mode.  Compression is a mandatory property, but we provide failsafe checks
    ' for broken and/or absent compression mode indicators.
    Dim srcCompression As xcf_Compression
    If (idxProp >= 0) Then
        srcCompression = m_ImageProperties(idxProp).propData(0)
    Else
        srcCompression = xcf_Compress_RLE
    End If
    If (srcCompression < 0) Or (srcCompression > xcf_Compress_ZLib) Then srcCompression = xcf_Compress_RLE
    
    'XCF supports a complex matrix of potential color models and precisions.  Let's simplify the image-
    ' and layer-level flags into branchable toggles that simplify the process of converting to PD's internal
    ' 32-bpp BGRA requirement.  (Note also that some color model details are layer-specific, like alpha
    ' channel presence, while others are image-specific, like int vs float precisions.)
    Dim isAlpha As Boolean, isGray As Boolean, isIndexed As Boolean
    Dim precisionInt As Boolean, precisionFloat As Boolean, isGamma As Boolean, isLinear As Boolean
    Dim numBitsPerChannel As Long
    
    FillPixelFormats idxLayer, isAlpha, isGray, isIndexed, precisionInt, isGamma, numBitsPerChannel
    isLinear = (Not isGamma)
    precisionFloat = (Not precisionInt)
    
    'Source bits and destination bits will be stored in dedicated arrays.  We will "extract" source bytes
    ' from the base XCF file into this local "source" array (this is important because we don't always know
    ' the true size of a tile's data, so we just have to grab a "big enough" chunk from the file which will
    ' almost always contain some amount of unwanted trailing bytes).
    '
    'We will then blindly decompress the bytes from our local "source" array into a "destination" array
    ' (whose size *is* known and fixed, based on calculations we can make on tile size, color format, and
    ' precision).  We will then swizzle and down/upsample destination bytes as necessary, then color-manage
    ' everything into the dedicated tile DIB (which will ultimately get blted into the assembled "full layer
    ' image" as you see it in the app).
    Dim xcfBytesPerPixel As Long
    xcfBytesPerPixel = m_Layers(idxLayer).lBytesPerPixel
    
    'Some swizzling and up/downsampling can be performed with help from LCMS.  (Note that for many formats,
    ' we still need to massage the data to get it to a place where LCMS can work with it - this is especially
    ' true for formats unsupported by LCMS, like 32-bit integer channel precision.)
    '
    'To make this work, however, we need to define all possible LCMS color formats that we may be relying on.
    Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile, tmpTransform As pdLCMSTransform
    Set srcProfile = New pdLCMSProfile
    
    'If the XCF embedded a color profile, use that; otherwise, we'll supply a generic source profile
    ' (based on the linear or gamma color space encoding supplied by the header)
    Dim usedBasicProfile As Boolean: usedBasicProfile = False
    
    If (Not m_Profile Is Nothing) Then
        
        'If the source profile is broken, fall back to a default one.
        If (Not srcProfile.CreateFromPDICCObject(m_Profile)) Then
            usedBasicProfile = True
            If isLinear Then srcProfile.CreateLinearRGBProfile Else srcProfile.CreateSRGBProfile True
        End If
        
    Else
        
        'Source data can be in linear or gamma-encoded format.  Gamma-encoded format is close to sRGB and the
        ' spec recommends just assuming that on legacy files anyway.
        usedBasicProfile = True
        If isGray Then
            If isLinear Then srcProfile.CreateGenericGrayscaleProfile 1# Else srcProfile.CreateGenericGrayscaleProfile 2.2
        Else
            If isLinear Then srcProfile.CreateLinearRGBProfile Else srcProfile.CreateSRGBProfile True
        End If
        
    End If
    
    'If we had to create our own color profile on-the-fly, override the class-level color profile
    ' with the one we created; we will cache the profile in PD's central cache so other color-managed
    ' portions of the program can access it.
    If usedBasicProfile Then
        Set m_Profile = New pdICCProfile
        m_Profile.LoadICCFromLCMSProfile srcProfile
    End If
    
    'For now, do a hard-convert into sRGB format
    Set dstProfile = New pdLCMSProfile
    dstProfile.CreateSRGBProfile
    
    'Some color modes (e.g. indexed images) don't require color management, so we must check for a
    ' null-format before attempting to create a transform.
    Dim srcFormat As LCMS_PIXEL_FORMAT
    srcFormat = GetSrcLCMSFormat(isIndexed, isGray, precisionInt, isAlpha, numBitsPerChannel, srcCompression)
    
    If (srcFormat <> 0) Then
        Dim flgTransform As LCMS_TRANSFORM_FLAGS
        If isAlpha Then flgTransform = cmsFLAGS_COPY_ALPHA Else flgTransform = 0
        Set tmpTransform = New pdLCMSTransform
        tmpTransform.CreateTwoProfileTransform srcProfile, dstProfile, srcFormat, TYPE_BGRA_8, INTENT_PERCEPTUAL, flgTransform
    End If
    
    'Pre-calculate a "worst-case" source size.  This formula is arbitrary; maybe there is a better way,
    ' but I derived this from the following clue in the XCF spec:
    '
    '"The RLE encoding can cause degenerated encodings in which the original data stream may double in size
    ' (or grow to arbitrarily large sizes if (128,0,0) operations are inserted). Such encodings must be avoided,
    ' as GIMP's XCF reader expects that the size of an encoded tile is never more than 24 KB, which is only 1.5
    ' times the unencoded size of a 64x64 RGBA tile.
    '
    'Based on this advice, I've gone with a similar formula of "50% overhead for worst-case compression".
    Dim srcSize As Long, maxSrcSize As Long, srcBytes() As Byte
    maxSrcSize = Int((64 * 64 * xcfBytesPerPixel) * 1.5 + 0.5)
    ReDim srcBytes(0 To maxSrcSize - 1) As Byte
    
    'Decompressed destination size is fixed, thankfully, but note that not all tiles will use the full
    ' tile size.  Right- and bottom-most tiles may be smaller if the image dimensions are not a clean
    ' multiple of 64.
    '
    'Note also that we declare this array as "byte" but it may actually contain shorts, floats, etc -
    ' we'll alias it as necessary when the time comes.
    Dim dstSize As Long, maxDstSize As Long, dstBytes() As Byte
    maxDstSize = 64 * 64 * xcfBytesPerPixel
    ReDim dstBytes(0 To maxDstSize - 1) As Byte
    
    'Various HDR formats require extra handling because they either do not use LCMS-supported color formats
    ' (e.g. 32-bit integer precision), or because they use silly big-endian encoding for floats (LCMS only
    ' offers endianness swapping for int formats).  Due to the complexity of these transforms, we want an
    ' additional swizzling buffer available for HDR formats, at the same size as the decompressed pixel data.
    Dim tmpSwizzle() As Byte
    If (numBitsPerChannel > 8) Then ReDim tmpSwizzle(0 To maxDstSize - 1) As Byte
    
    'Precalculate dimensions of the last row/column of tiles; this saves us some computation effort
    ' inside the tile loop.
    Dim tileWidth As Long, tileHeight As Long, numTilePixels As Long
    
    Dim tileWidthLast As Long, tileHeightLast As Long
    tileWidthLast = m_Layers(idxLayer).lWidth - ((m_Layers(idxLayer).lWidth \ 64) * 64)
    If (tileWidthLast <= 0) Then tileWidthLast = 64
    tileHeightLast = m_Layers(idxLayer).lHeight - ((m_Layers(idxLayer).lHeight \ 64) * 64)
    If (tileHeightLast <= 0) Then tileHeightLast = 64
    
    'Start iterating tiles
    Dim xTile As Long, yTile As Long
    For yTile = 0 To yTileMax
    For xTile = 0 To xTileMax
        
        'PDDebug.LogAction "Parsing tile (" & xTile & ", " & yTile & ")"
        
        'Unfortunately, GIMP does not encode compressed tile sizes into the XCF file.  We can infer the
        ' size for most tiles (by subtracting this tile's file offset from the next tile's offset),
        ' but for the final tile in the image, we just have to grab a "worst-case" chunk of bytes from the
        ' file and hope for the best.
        If (idxTile < numTilesTotal - 1) Then
            srcSize = m_Layers(idxLayer).ptrToTiles(idxTile + 1) - m_Layers(idxLayer).ptrToTiles(idxTile)
        Else
            srcSize = maxSrcSize
        End If
        
        'Destination size is fixed, and contingent only on the current tile's dimensions and the color-depth
        ' of the image.
        If (xTile < xTileMax) Then tileWidth = 64 Else tileWidth = tileWidthLast
        If (yTile < yTileMax) Then tileHeight = 64 Else tileHeight = tileHeightLast
        numTilePixels = tileWidth * tileHeight
        dstSize = numTilePixels * xcfBytesPerPixel
        
        'Forcibly align the file stream pointer for this tile, then pull the required amount of source bytes
        ' into a local array.  (Manual RLE decoding is faster this way, vs pulling byte-by-byte from the file.)
        m_Stream.SetPosition m_Layers(idxLayer).ptrToTiles(idxTile), FILE_BEGIN
        
        '(PD's stream object automatically protects against OOB reads, but I still prefer to check for EOF and
        ' shrink the source buffer accordingly, especially since XCF requires such a huge "safety" margin for
        ' RLE-compressed tiles.)
        If (srcSize > m_Stream.GetStreamSize - m_Stream.GetPosition) Then srcSize = m_Stream.GetStreamSize - m_Stream.GetPosition
        srcSize = m_Stream.ReadBytesToBarePointer(VarPtr(srcBytes(0)), srcSize)
        
        'We are now done with the file stream and can happily ignore it until the next tile.
        
        'PDDebug.LogAction "Read " & srcSize & " bytes from position " & m_Layers(idxLayer).ptrToTiles(idxTile)
        
        'What happens next varies according to compression mode.  For DEFLATE, for example, we can just
        ' decompress as-is.  For RLE, however, we need to manually handle decompression.
        Select Case srcCompression
            
            'Shouldn't exist in the wild; will revisit if I can find a way to produce such a file
            Case xcf_Compress_None
                
            'Default setting for XCF files; after RLE decompression, HDR images require unshuffling
            ' (since RLE operates on individual bytes, not channels)
            Case xcf_Compress_RLE
                DecompressRLE xcfBytesPerPixel, numTilePixels, srcBytes, dstBytes
                If (numBitsPerChannel > 8) Then UnshuffleHDRChannels numBitsPerChannel, numTilePixels, isAlpha, isGray, tmpSwizzle, dstBytes, maxDstSize
                
            'Available as an optional setting in modern XCF files, but rare "in the wild"
            Case xcf_Compress_ZLib
                
                'zLib compression is straightforward: decode the source bytes directly into the destination array,
                ' using a pre-calculated destination size as our fixed indicator for "completion".  (Note that we
                ' may still need to swap endianness after decompression, but we at least won't have to deal with
                ' planar-to-interleaved conversion.)
                Dim dcmpResult As Long
                dcmpResult = Plugin_libdeflate.Decompress_ZLib(VarPtr(dstBytes(0)), dstSize, VarPtr(srcBytes(0)), srcSize)
                If (dcmpResult <> 0) Then InternalError FUNC_NAME, "bad decompress: " & dcmpResult
                
        End Select
        
        'If the source image is in indexed mode, we must handle it manually; otherwise, we can use
        ' LittleCMS to swizzle for us.
        
        'Indexed (can only by 8-bpp)
        If isIndexed Then
            RenderIndexedColor dstBytes, tileWidth, tileHeight, isAlpha
        
        'Gray and true-color are handled by color profile
        Else
            
            'Manually handle BE to LE conversion for HDR float formats (and 32-bit integer, which is not
            ' natively supported by LCMS)
            SwapEndiannessHDR srcCompression, numBitsPerChannel, precisionFloat, isGray, numTilePixels, isAlpha, tmpSwizzle, dstBytes, maxDstSize
            
            'Finalize plane and scanline calculations for LCMS conversion
            Dim lcmsBytesPerLine As Long, lcmsBytesPerPlane As Long
            FinalizeLCMSTransform lcmsBytesPerLine, lcmsBytesPerPlane, srcCompression, precisionInt, isAlpha, isGray, tileWidth, tileHeight, xcfBytesPerPixel, numBitsPerChannel
            
            'Note that the creation of this transform is extremely specific.  In particular, we have to explicitly
            ' address planar-to-interleaved (on RLE-compressed images), and we are also rendering into a fixed-size
            ' 64x64 image buffer - but the incoming tile *may be smaller than this*, especially on row/column ends.
            ' (So the source size is allowed to float per tile size, but the destination size is *not*.)
            '
            'This is all dealt with in previous steps, but I mention it here because the input settings to this call
            ' vary *wildly* by source pixel format and color mode, and this line is a potential source of headaches
            ' for anyone who wants to adapt this code for their own purposes.
            tmpTransform.ApplyTransformToArbitraryMemoryEx VarPtr(dstBytes(0)), m_Tile.GetDIBPointer, tileWidth, tileHeight, lcmsBytesPerLine, 64 * 4, lcmsBytesPerPlane, 0&
            
        End If
        
        'This tile is now fully rendered in color-managed 32-bit BGRA format!
        ' Blt it into its final location in the merged layer surface.
        GDI.BitBltWrapper m_Layers(idxLayer).lDIB.GetDIBDC, xTile * 64, yTile * 64, tileWidth, tileHeight, m_Tile.GetDIBDC, 0, 0, vbSrcCopy
        
        'Increment our pointer into the source tile offset table
        idxTile = idxTile + 1
        
    Next xTile
    Next yTile
    
    Import_Stage4b_GenerateLayerPixels = True
    
    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage4b_GenerateLayerPixels = False
    
End Function

'A number of support functions for tile decompression follow
Private Sub RenderIndexedColor(ByRef dstBytes() As Byte, ByVal tileWidth As Long, ByVal tileHeight As Long, ByVal isAlpha As Boolean)

    'Indexed color images must be manually assembled into a 32-bpp BGRA tile.
    Dim pxTile() As RGBQuad, pxSA As SafeArray2D
    m_Tile.WrapRGBQuadArrayAroundDIB pxTile, pxSA
    
    Dim idxSrc As Long
    idxSrc = 0
    
    Dim x As Long, y As Long
    For y = 0 To tileHeight - 1
    For x = 0 To tileWidth - 1
        If isAlpha Then
            pxTile(x, y) = m_Palette(dstBytes(idxSrc))
            pxTile(x, y).Alpha = dstBytes(tileWidth * tileHeight + idxSrc)
            idxSrc = idxSrc + 1
        Else
            pxTile(x, y) = m_Palette(dstBytes(idxSrc))
            idxSrc = idxSrc + 1
        End If
    Next x
    Next y
    
    m_Tile.UnwrapRGBQuadArrayFromDIB pxTile
    
End Sub

Private Sub FinalizeLCMSTransform(ByRef lcmsBytesPerLine As Long, ByRef lcmsBytesPerPlane As Long, ByVal srcCompression As xcf_Compression, ByVal precisionInt As Boolean, ByVal isAlpha As Boolean, ByVal isGray As Boolean, ByVal tileWidth As Long, ByVal tileHeight As Long, ByVal xcfBytesPerPixel As Long, ByVal numBitsPerChannel As Long)
    
    'We need to specify some custom values here because XCF files support some precisions that are not
    ' natively supported by LCMS (like 32-bit integer).
    If (numBitsPerChannel = 8) Then
        
        '8-bit XCF data has not been preprocessed by us (unlike HDR data).  This means that it will
        ' still be in planar mode if compressed via RLE.
        If (srcCompression = xcf_Compress_ZLib) Then
            lcmsBytesPerLine = tileWidth * xcfBytesPerPixel
            lcmsBytesPerPlane = tileWidth * tileHeight
        Else
            lcmsBytesPerLine = tileWidth
            lcmsBytesPerPlane = (numBitsPerChannel \ 8) * tileWidth * tileHeight
        End If
    
    'HDR data has been preprocessed by us (whether coming from zLib or RLE compression),
    ' but we still need to declare some "special" modes that LCMS doesn't support (and which we have
    ' manually solved in a previous step)
    Else
        
        '32-byte int data has been pre-downsampled to shorts
        If precisionInt And (numBitsPerChannel > 16) Then
            numBitsPerChannel = 16
            xcfBytesPerPixel = xcfBytesPerPixel \ 2
        End If
            
        'Everything else has been manually converted from planar to interleaved and endianness-swapped as necessary
        lcmsBytesPerLine = tileWidth * xcfBytesPerPixel
        lcmsBytesPerPlane = (numBitsPerChannel \ 8) * tileWidth * tileHeight
        
    End If
    
End Sub

Private Sub SwapEndiannessHDR(ByVal srcCompression As Long, ByVal numBitsPerChannel As Long, ByVal precisionFloat As Boolean, ByVal isGray As Boolean, ByVal numTilePixels As Long, ByVal isAlpha As Boolean, ByRef tmpSwizzle() As Byte, ByRef dstBytes() As Byte, ByVal maxDstSize As Long)

    'HDR float formats are in big-endian notation.  We must manually swizzle before converting,
    ' because LCMS does not support this format natively.
    If (numBitsPerChannel > 8) Then
    
        'With swizzling finished, we now need to manually convert from big- to little-endian
        ' for float formats.
        If precisionFloat Then
            If (numBitsPerChannel = 16) Then
                VBHacks.SwapEndianness16 dstBytes
            ElseIf (numBitsPerChannel = 32) Then
                VBHacks.SwapEndianness32 dstBytes
            Else
                VBHacks.SwapEndianness64 dstBytes
            End If
        
        'Integer precision
        Else
            
            '32-bit integer data has an additional consideration.  LCMS does not provide native
            ' support for this color format (because it's a stupid and wasteful format, frankly,
            ' when 32-bit floats are RIGHT THERE), so we're going to downsample 32-bit integer data
            ' to 16-bit data, since it's just going to get smushed down to 8-bit for PD anyway.
            If (numBitsPerChannel > 16) Then
                
                Dim x As Long
                
                'Swap buffers again, for simplicity's sake.  (Note that the array size check is a failsafe only.)
                If (UBound(tmpSwizzle) <> UBound(dstBytes)) Then ReDim tmpSwizzle(0 To UBound(dstBytes)) As Byte
                CopyMemoryStrict VarPtr(tmpSwizzle(0)), VarPtr(dstBytes(0)), maxDstSize
                
                Dim pxSize As Long
                If isGray Then pxSize = 2 Else pxSize = 6
                If isAlpha Then pxSize = pxSize + 2
                
                For x = 0 To numTilePixels - 1
                    dstBytes(x * pxSize + 0) = tmpSwizzle(x * pxSize * 2 + 1)
                    dstBytes(x * pxSize + 1) = tmpSwizzle(x * pxSize * 2 + 0)
                    If isGray Then
                        If isAlpha Then
                            dstBytes(x * pxSize + 2) = tmpSwizzle(x * pxSize * 2 + 5)
                            dstBytes(x * pxSize + 3) = tmpSwizzle(x * pxSize * 2 + 4)
                        End If
                    Else
                        dstBytes(x * pxSize + 2) = tmpSwizzle(x * pxSize * 2 + 5)
                        dstBytes(x * pxSize + 3) = tmpSwizzle(x * pxSize * 2 + 4)
                        dstBytes(x * pxSize + 4) = tmpSwizzle(x * pxSize * 2 + 9)
                        dstBytes(x * pxSize + 5) = tmpSwizzle(x * pxSize * 2 + 8)
                        If isAlpha Then
                            dstBytes(x * pxSize + 6) = tmpSwizzle(x * pxSize * 2 + 13)
                            dstBytes(x * pxSize + 7) = tmpSwizzle(x * pxSize * 2 + 12)
                        End If
                    End If
                Next x
            
            'For plain 16-bit integer data, just swap endianness in-place
            Else
                VBHacks.SwapEndianness16 dstBytes
            End If
                    
        End If
        
    End If
    
End Sub

Private Sub UnshuffleHDRChannels(ByVal numBitsPerChannel As Long, ByVal numTilePixels As Long, ByVal isAlpha As Boolean, ByVal isGray As Boolean, ByRef tmpSwizzle() As Byte, ByRef dstBytes() As Byte, ByVal maxDstSize As Long)

    If (numBitsPerChannel > 8) Then
        
        'Half-float RGB/A data appears to behave differently.  I'm not sure why?
        CopyMemoryStrict VarPtr(tmpSwizzle(0)), VarPtr(dstBytes(0)), maxDstSize
        
        'Channels vary by color model (gray vs RGB)
        Dim nChannels As Long
        If isGray Then
            nChannels = 1
        Else
            nChannels = 3
        End If
        
        'Alpha is always encoded as an additional channel
        If isAlpha Then nChannels = nChannels + 1
        
        'From final channel count, figure out the size of a single pixel *in bytes*
        Dim nBytesPerPixel As Long
        nBytesPerPixel = nChannels * (numBitsPerChannel \ 8)
        
        Dim nBytesPerChannel As Long
        nBytesPerChannel = numBitsPerChannel \ 8
        
        'Unshuffle accordingly
        Dim idxSrcOffset As Long
        For idxSrcOffset = 0 To nBytesPerPixel - 1
            Dim x As Long
            For x = 0 To numTilePixels - 1
                dstBytes(x * nBytesPerPixel + idxSrcOffset) = tmpSwizzle(idxSrcOffset * numTilePixels + x)
            Next x
        Next idxSrcOffset
    
    End If
    
End Sub

Private Sub DecompressRLE(ByVal xcfBytesPerPixel As Long, ByVal numTilePixels As Long, ByRef srcBytes() As Byte, ByRef dstBytes() As Byte)

    'GIMP's RLE strategy uses a homebrew design (sigh).  As with most RLE schemes, channels are encoded
    ' individually (planar not interleaved), and at higher bit-depths they are RLE-encoded byte-by-byte,
    ' big-endian order, so we also need to sort that shit out.
    '
    'Even worse, the format provides no markers to align scanlines (or any source data whatsoever),
    ' so we have to manually watch decompressed size and stop once all pixels have been assembled.
    ' (Yes, this makes the format extremely susceptible to malicious encoding - but that's not
    ' our fault!)
    Dim idxRLE As Long, rleMarker As Byte, rleMult As Long
    
    'I call this "idxChannel" but really, it operates on a byte-level.  HDR images (which are encoded
    ' as RLE by default, ugh) place each byte of a given channel into its own plain - so in a 32-bit
    ' RGBA channel, all the 1st R bytes will be placed together, then the 2nd R bytes, etc.
    Dim idxChannel As Long
    For idxChannel = 0 To xcfBytesPerPixel - 1
        
        'As a failsafe against faulty RLE runs, forcibly align the destination pointer at the start
        ' of each "channel" (but leave the source pointer where it is).
        Dim endOfThisChannel As Long
        endOfThisChannel = (idxChannel + 1) * numTilePixels
        
        Dim idxDst As Long
        idxDst = idxChannel * numTilePixels
        
        'Success is noted by reaching the end of the destination channel.  Note that we do *not* always
        ' validate destination size before writing (because that kills performance).  Instead, we are
        ' content to simply crash-out on faulty RLE runs.
        Do While (idxDst < endOfThisChannel)
            
            'Always start with an RLE marker
            rleMarker = srcBytes(idxRLE)
            idxRLE = idxRLE + 1
            
            '"Short run of identical bytes"
            If (rleMarker <= 126) Then
                rleMarker = rleMarker + 1
                FillMemory VarPtr(dstBytes(idxDst)), rleMarker, srcBytes(idxRLE)
                idxRLE = idxRLE + 1
                idxDst = idxDst + rleMarker
                
            '"Long run of identical bytes"
            ElseIf (rleMarker = 127) Then
                rleMult = srcBytes(idxRLE)
                rleMult = rleMult * 256 + srcBytes(idxRLE + 1)
                FillMemory VarPtr(dstBytes(idxDst)), rleMult, srcBytes(idxRLE + 2)
                idxRLE = idxRLE + 3
                idxDst = idxDst + rleMult
            
            '"Long run of different bytes"
            ElseIf (rleMarker = 128) Then
                rleMult = srcBytes(idxRLE)
                rleMult = rleMult * 256 + srcBytes(idxRLE + 1)
                CopyMemoryStrict VarPtr(dstBytes(idxDst)), VarPtr(srcBytes(idxRLE + 2)), rleMult
                idxRLE = idxRLE + rleMult + 2
                idxDst = idxDst + rleMult
                
            '"Short run of different bytes"
            Else
                rleMarker = 256 - rleMarker
                CopyMemoryStrict VarPtr(dstBytes(idxDst)), VarPtr(srcBytes(idxRLE)), rleMarker
                idxRLE = idxRLE + rleMarker
                idxDst = idxDst + rleMarker
            End If
            
        Loop
        
    Next idxChannel

End Sub

Private Function GetSrcLCMSFormat(ByVal isIndexed As Boolean, ByVal isGray As Boolean, ByVal precisionInt As Boolean, ByVal isAlpha As Boolean, ByVal numBitsPerChannel As Long, ByVal srcCompression As xcf_Compression) As LCMS_PIXEL_FORMAT

    'Indexed mode is not handled via profile; it is handled manually so we don't need to handle it here
    If isIndexed Then
        GetSrcLCMSFormat = 0
        Exit Function
    End If
    
    'Gray (can be any bit-depth)
    If isGray Then
        
        If precisionInt Then
            
            'Normal 8-bit data will be planar, by default.  We manually convert HDR images to interleaved while
            ' swapping between big-endian and little-endian source values, so in this function you'll notice that
            ' 8-bit data is left planar, while HDR data uses normal interleaved mode (because we handled interleaving).
            If (numBitsPerChannel = 8) Then
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_GRAYA_8_PLANAR
                Else
                    GetSrcLCMSFormat = TYPE_GRAY_8
                End If
                
            ElseIf (numBitsPerChannel = 16) Then
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_GRAYA_16
                Else
                    GetSrcLCMSFormat = TYPE_GRAY_16
                End If
            Else
                
                'LCMS doesn't support 32-bit integer components.  We will manually downsample to
                ' 16-bit integer prior to color-management, so expressly request 16-bit here.
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_GRAYA_16
                Else
                    GetSrcLCMSFormat = TYPE_GRAY_16
                End If
            
            End If
        
        'Floating-point encoding
        Else
            If (numBitsPerChannel = 16) Then
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_GRAYA_HALF_FLT
                Else
                    GetSrcLCMSFormat = TYPE_GRAY_HALF_FLT
                End If
            ElseIf (numBitsPerChannel = 32) Then
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_GRAYA_FLT
                Else
                    GetSrcLCMSFormat = TYPE_GRAY_FLT
                End If
            Else
                
                'Double-precision is only "theoretically" supported today, because GIMP doesn't actually
                ' allow you to specify double-precision inside the app (though the spec allows for it as a
                ' valid color-depth).
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_GRAYA_DBL
                Else
                    GetSrcLCMSFormat = TYPE_GRAY_DBL
                End If
            
            End If
            
        End If
        
    'RGB (can be any bit-depth)
    Else
        If precisionInt Then
            If (numBitsPerChannel = 8) Then
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_RGBA_8_PLANAR
                Else
                    GetSrcLCMSFormat = TYPE_RGB_8_PLANAR
                End If
            ElseIf (numBitsPerChannel = 16) Then
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_RGBA_16
                Else
                    GetSrcLCMSFormat = TYPE_RGB_16
                End If
            Else
                
                'LCMS doesn't support 32-bit integer components.  We will manually downsample to
                ' 16-bit integer prior to color-management, so expressly request 16-bit here.
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_RGBA_16
                Else
                    GetSrcLCMSFormat = TYPE_RGB_16
                End If
            
            End If
        
        'Floating-point encoding
        Else
            If (numBitsPerChannel = 16) Then
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_RGBA_HALF_FLT
                Else
                    GetSrcLCMSFormat = TYPE_RGB_HALF_FLT
                End If
            ElseIf (numBitsPerChannel = 32) Then
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_RGBA_FLT
                Else
                    GetSrcLCMSFormat = TYPE_RGB_FLT
                End If
            Else
                
                'Double-precision is only "theoretically" supported today, because GIMP doesn't actually
                ' allow you to specify double-precision inside the app (though the spec allows for it as a
                ' valid color-depth).
                If isAlpha Then
                    GetSrcLCMSFormat = TYPE_RGBA_DBL
                Else
                    GetSrcLCMSFormat = TYPE_RGB_DBL
                End If
            
            End If
            
        End If
    
    End If
    
    'If the image is in zLib format, it was encoded using normal interleaved format (RGBARGBARGBA)
    ' instead of planar format (RRRGGGBBBAAA), even in normal 8-bit color mode.  Deactivate that bit accordingly.
    If (srcCompression = xcf_Compress_ZLib) Then GetSrcLCMSFormat = GetSrcLCMSFormat And (Not &H1000&)
    
End Function

Private Sub FillPixelFormats(ByVal idxLayer As Long, ByRef isAlpha As Boolean, ByRef isGray As Boolean, ByRef isIndexed As Boolean, ByRef precisionInt As Boolean, ByRef isGamma As Boolean, ByRef numBitsPerChannel As Long)
    
    With m_Layers(idxLayer)
        isAlpha = (.lColorMode = xcf_GrayA) Or (.lColorMode = xcf_IndexedA) Or (.lColorMode = xcf_RGBA)
        isGray = (.lColorMode = xcf_GrayA) Or (.lColorMode = xcf_GrayX)
        isIndexed = (.lColorMode = xcf_IndexedA) Or (.lColorMode = xcf_IndexedX)
    End With
    
    precisionInt = (m_imagePrecision = xcf_08bitIntGamma) Or (m_imagePrecision = xcf_08bitIntLinear) Or (m_imagePrecision = xcf_16bitIntGamma) Or (m_imagePrecision = xcf_16bitIntLinear) Or (m_imagePrecision = xcf_32bitIntGamma) Or (m_imagePrecision = xcf_32bitIntLinear)
    isGamma = (m_imagePrecision = xcf_08bitIntGamma) Or (m_imagePrecision = xcf_16bitIntGamma) Or (m_imagePrecision = xcf_32bitIntGamma) Or (m_imagePrecision = xcf_16bitFltGamma) Or (m_imagePrecision = xcf_32bitFltGamma) Or (m_imagePrecision = xcf_64bitFltGamma)
    numBitsPerChannel = GetNumBitsFromImagePrecision()
    
End Sub

Private Function GetNumBitsFromImagePrecision() As Long

    If ((m_imagePrecision = xcf_08bitIntGamma) Or (m_imagePrecision = xcf_08bitIntLinear)) Then
        GetNumBitsFromImagePrecision = 8
    ElseIf ((m_imagePrecision = xcf_16bitFltGamma) Or (m_imagePrecision = xcf_16bitFltLinear) Or (m_imagePrecision = xcf_16bitIntGamma) Or (m_imagePrecision = xcf_16bitIntLinear)) Then
        GetNumBitsFromImagePrecision = 16
    ElseIf ((m_imagePrecision = xcf_32bitFltGamma) Or (m_imagePrecision = xcf_32bitFltLinear) Or (m_imagePrecision = xcf_32bitIntGamma) Or (m_imagePrecision = xcf_32bitIntLinear)) Then
        GetNumBitsFromImagePrecision = 32
    Else
        GetNumBitsFromImagePrecision = 64
    End If
    
End Function

'ONLY CALL THIS FUNCTION FROM Import_Stage4, above.
'
'Given a valid layer index, traverse the complex pixel hierarchy associated with said layer.
' From this hierarchy, we will ultimately end up with a list of file offsets to each sub-tile
' that comprises this layer's image.
'
'Incoming *and* outgoing stream pointer alignment is *not* guaranteed.  This function will move
' the pointer around as necessary to decode layer information.  Pointer alignment, if required,
' *must* be handled by the caller.
Private Function Import_Stage4a_PrepTileRetrieval(ByVal ptrStart As Long, ByVal validationWidth As Long, ByVal validationHeight As Long, ByRef dstBytesPerPixel As Long, ByRef dstNumTilesX As Long, ByRef dstNumTilesY As Long, ByRef dstNumTilesTotal As Long, ByRef dstTilePtrs() As Long) As Boolean

    Const FUNC_NAME As String = "Import_Stage4a_PrepTileRetrieval"
    Import_Stage4a_PrepTileRetrieval = True
    
    On Error GoTo BrokenImage
    
    'Modern GIMP files use 64-bit pointers; legacy use 32-bit
    Dim ptrIs64bit As Boolean, ptrFileH As Long
    ptrIs64bit = (m_xcfVersion >= 11)
    
    'GIMP uses an (unnecessarily) complicated mechanism for encoding pixel data.  To wit:
    ' Each layer provides an offset to a...
    ' "Hierarchy" (their term), which is a small header followed by an array of pointers to...
    ' "Levels" (their term), which are small headers followed by an array of pointers to...
    ' "Tiles", which are exactly what they sound like (and contain actual pixel data).
    '
    'As part of this setup, GIMP files appear to have been designed with mipmapping in mind,
    ' but that's never actually been implemented so there is just one "actual" level of pixel data
    ' followed by an arbitrary amount of "dummy" layers (with null tile pointers to non-existent
    ' mipmaps).
    '
    'So actual pixel tiles are only meaningful for one "actual" level of pixels, and tiles are
    ' additionally cumbersome to work with because we don't actually know the length of source data
    ' available for each tile except through inference (difference between neighboring offsets,
    ' which works for all tiles except the final one so special handling is required anyway).
    ' This ill-conceived design means we must manually track output bytes and stop processing
    ' once an expected "tile size" worth of destination pixels is decoded.
    '
    'On top of all this mess is a huge matrix of potential incoming color-spaces and precisions
    ' (from 8-bit grayscale to 256-bit double-precision RGBA) with no clear spec limitations on
    ' which compression schemes work with which depths, so we're mostly stuck flailing around
    ' on our own to try and solve all combinations.  Yay?
    '
    'But I'm getting ahead of myself.  In this step, we want to parse this layer's "hierarchy"
    ' block, then the relevant "levels" block, then retrieve all tile pointers and cache them
    ' in a persistent table (that subsequent steps can use to actually retrieve tile data).
    
    'Start by forcibly aligning the hierarchy pointer.
    m_Stream.SetPosition ptrStart, FILE_BEGIN
    
    'Validate the hierarchy header
    Dim checkWidth As Long, checkHeight As Long
    checkWidth = m_Stream.ReadLong_BE()
    checkHeight = m_Stream.ReadLong_BE()
    dstBytesPerPixel = m_Stream.ReadLong_BE()
    
    If (checkWidth <> validationWidth) Or (checkHeight <> validationHeight) Then
        InternalError FUNC_NAME, "bad hierarchy dimensions"
        Import_Stage4a_PrepTileRetrieval = False
        Exit Function
    End If
    
    If (dstBytesPerPixel <= 0) Then
        InternalError FUNC_NAME, "bad bytes-per-pixel"
        Import_Stage4a_PrepTileRetrieval = False
        Exit Function
    End If
    
    'Next is a list of "level" pointers.  We retrieve all of them, but really only want the first one.
    ' (We retrieve all in case GIMP eventually implements this feature.  I wouldn't hold my breath lol)
    Const INIT_LEVEL_COUNT As Long = 16
    Dim lstLevels() As Long, numLevels As Long
    ReDim lstLevels(0 To INIT_LEVEL_COUNT - 1) As Long
    numLevels = 0
    
    'Pointers in this function are confusing.  In most cases, they are variable-size like other pointers
    ' in an XCF file (64-bit on modern XCF files, 32-bit on earlier version).  However, on some of the
    ' "dummy" structs GIMP insists on writing, pointers remain 32-bit regardless of file version.
    ' Pay close attention to where we use which pointer length.
    Dim ptrNext As Long, ptrH As Long
    If ptrIs64bit Then ptrH = m_Stream.ReadLong_BE()
    ptrNext = m_Stream.ReadLong_BE()
    
    Do While (ptrNext <> 0) And (ptrH = 0)
        If (numLevels > UBound(lstLevels)) Then ReDim Preserve lstLevels(0 To numLevels * 2 - 1) As Long
        lstLevels(numLevels) = ptrNext
        numLevels = numLevels + 1
        If ptrIs64bit Then ptrH = m_Stream.ReadLong_BE()
        ptrNext = m_Stream.ReadLong_BE()
    Loop
    
    'Files over 2GB are not useable in PD
    If (ptrH <> 0) Then
        InternalError FUNC_NAME, "2 GB limit exceeded"
        Import_Stage4a_PrepTileRetrieval = False
        Exit Function
    End If
    
    'Ensure the first retrieved level pointer is valid
    If (lstLevels(0) = 0) Then
        InternalError FUNC_NAME, "null base level"
        Import_Stage4a_PrepTileRetrieval = False
        Exit Function
    End If
    
    'If we're still here, we have a valid level pointer (and a bunch of dummy level pointers, too).
    ' This next code block is set up to iterate levels, but I've added a hard branch to only really
    ' process the first iteration.  (Subsequent iterations were mostly used to test XCF layout to
    ' study details of the "dummy" layers in case I decide to add XCF export support someday.)
    Dim idxLevel As Long
    For idxLevel = 0 To numLevels - 1
        
        'Hard branch for now; you can inspect non-base levels if curious, but they do not contain
        ' useful data (as of April 2022).
        If (idxLevel > 0) Then Exit For
        
        'Move to the specified level offset
        m_Stream.SetPosition lstLevels(idxLevel), FILE_BEGIN
        
        'Validate width/height yet again
        checkWidth = m_Stream.ReadLong_BE()
        checkHeight = m_Stream.ReadLong_BE()
        
        If (checkWidth <> validationWidth) Or (checkHeight <> validationHeight) Then
            InternalError FUNC_NAME, "bad hierarchy dimensions"
            Import_Stage4a_PrepTileRetrieval = False
            Exit Function
        End If
        
        'We now need to retrieve pointers to all tiles in this level of this layer.
        
        'Before doing this, calculate the number of tiles in the x and y direction.
        ' GIMP always uses 64x64 px tiles regardless of image size, with only the right/bottom
        ' tiles allowed to encode fewer pixels than this.
        dstNumTilesX = (checkWidth + 63) \ 64
        dstNumTilesY = (checkHeight + 63) \ 64
        
        'Use these to retrieve all tile pointers
        dstNumTilesTotal = dstNumTilesX * dstNumTilesY
        
        'Failsafe checks before allocation
        If ((dstNumTilesX <= 0) Or (dstNumTilesY <= 0)) Then
            InternalError FUNC_NAME, "null tile count"
            Import_Stage4a_PrepTileRetrieval = False
            Exit Function
        ElseIf (dstNumTilesTotal <= 0) Then
            InternalError FUNC_NAME, "bad tile count"
            Import_Stage4a_PrepTileRetrieval = False
            Exit Function
        End If
        
        ReDim dstTilePtrs(0 To dstNumTilesTotal - 1) As Long
        
        Dim idxTile As Long
        For idxTile = 0 To dstNumTilesTotal - 1
            
            'Only the base level uses 64-bit pointers; subsequent "dummy" levels use uint32 no matter what
            If (idxLevel = 0) And ptrIs64bit Then ptrFileH = m_Stream.ReadLong_BE()
            dstTilePtrs(idxTile) = m_Stream.ReadLong_BE()
            
        Next idxTile
        
        'With all tiles received, the tile list must be terminated by a null pointer
        If (m_Stream.ReadLong_BE() <> 0) Then
            InternalError FUNC_NAME, "tile segment terminated unexpectedly"
            Import_Stage4a_PrepTileRetrieval = False
            Exit Function
        End If
        
    Next idxLevel
    
    'If we're still here, tile extraction was (presumably) successful!
    If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "Successfully retrieved " & dstNumTilesX & "x" & dstNumTilesY & " (" & dstNumTilesTotal & ") tile pointers"
        
    Import_Stage4a_PrepTileRetrieval = True
    
    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage4a_PrepTileRetrieval = False
    
End Function

'Import step 3: load all layers.  After this step, the stream pointer is no longer relevant.  (It will likely
' point at the first layer header, but the XCF format declares layers using file offsets, so arbitrary padding
' between segments *is* allowable.  We will forcibly align the stream pointer for all reads after this step.)
Private Function Import_Stage3_LoadLayersAndChannels(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "Import_Stage3_LoadLayersAndChannels"
    
    'This stage will only fail if we encounter 2+ GB files; otherwise, we assume success
    ' (and won't know otherwise until we start parsing layers+channels from their file offsets)
    Import_Stage3_LoadLayersAndChannels = True
    
    On Error GoTo BrokenImage
    
    'XCF files encode all layers as pointers to dedicated layer segments.  We need to track proper
    ' stream alignment here since we'll be jumping all over the file.
    m_numOfLayers = 0
    m_numOfChannels = 0
    
    '(Initial size is arbitrary; we'll resize these collections as we go)
    Const INIT_COLLECTION_COUNT As Long = 16
    ReDim m_Layers(0 To INIT_COLLECTION_COUNT - 1) As xcf_Layer
    ReDim m_Channels(0 To INIT_COLLECTION_COUNT - 1) As xcf_Channel
    
    'Note also that this segment uses a bunch of pointers, and per the XCF spec...
    ' "A POINTER is stored as a 32-bit integer (4 bytes) in big-endian order
    '  for XCF up to 10, and 64-bit (8 bytes), still big-endian, for XCF 11
    '  and over, allowing higher than 4GB XCF files since GIMP 2.10.0."
    '
    'Because of this, if we see 64-bit pointers with non-zero values in their high-word, we'll fail
    ' the file and exit immediately.
    Dim ptrFileL As Long, ptrFileH As Long
    
    Dim ptrIs64bit As Boolean
    ptrIs64bit = (m_xcfVersion >= 11)
    If ptrIs64bit Then ptrFileH = m_Stream.ReadLong_BE()
    ptrFileL = m_Stream.ReadLong_BE()
    
    Do While (ptrFileL <> 0) And (ptrFileH = 0)
        
        'Store this layer's pointer (absolute offset within the file)
        If (m_numOfLayers > UBound(m_Layers)) Then ReDim Preserve m_Layers(0 To m_numOfLayers * 2 - 1) As xcf_Layer
        m_Layers(m_numOfLayers).ptrInFile = ptrFileL
        m_numOfLayers = m_numOfLayers + 1
        
        'Grab the next layer pointer
        If ptrIs64bit Then ptrFileH = m_Stream.ReadLong_BE()
        ptrFileL = m_Stream.ReadLong_BE()
        
    Loop
    
    'If the file contains 2GB+ offsets, fail immediately
    If (ptrFileH <> 0) Then
        Import_Stage3_LoadLayersAndChannels = False
        Exit Function
    End If
    
    'With the layer collection complete, we can resize it precisely
    If (m_numOfLayers <> 0) Then ReDim Preserve m_Layers(0 To m_numOfLayers - 1) As xcf_Layer
    
    'If we're still here, let's also populate channel pointers.  (These have limited relevance to PD,
    ' but it costs nothing to retrieve them.)
    If ptrIs64bit Then ptrFileH = m_Stream.ReadLong_BE()
    ptrFileL = m_Stream.ReadLong_BE()
    
    Do While (ptrFileL <> 0) And (ptrFileH = 0)
        
        'Store this layer's pointer (absolute offset within the file)
        If (m_numOfChannels > UBound(m_Channels)) Then ReDim Preserve m_Channels(0 To m_numOfChannels * 2 - 1) As xcf_Channel
        m_Channels(m_numOfChannels).ptrInFile = ptrFileL
        m_numOfChannels = m_numOfChannels + 1
        
        'Grab the next channel pointer
        If ptrIs64bit Then ptrFileH = m_Stream.ReadLong_BE()
        ptrFileL = m_Stream.ReadLong_BE()
        
    Loop
    
    'We could technically fail 2+ GB files here as well, but since PD doesn't technically need channel data,
    ' we will attempt to load layer data (and pray there's enough memory to do so).
    
    'With the channel collection complete, we can resize it precisely
    If (m_numOfChannels <> 0) Then ReDim Preserve m_Channels(0 To m_numOfLayers - 1) As xcf_Channel
    
    'Report success based on non-zero layer counts
    Import_Stage3_LoadLayersAndChannels = (m_numOfLayers > 0)
    
    If XCF_DEBUG_VERBOSE Then
        If Import_Stage3_LoadLayersAndChannels Then
            PDDebug.LogAction "Found " & m_numOfLayers & " layer(s) and " & m_numOfChannels & " channel(s)"
            Dim i As Long
            For i = 0 To m_numOfLayers - 1
                PDDebug.LogAction "Layer offset: " & m_Layers(i).ptrInFile
            Next i
            If (m_numOfChannels > 0) Then
                For i = 0 To m_numOfChannels - 1
                    PDDebug.LogAction "Channel offset: " & m_Layers(i).ptrInFile
                Next i
            End If
        Else
            PDDebug.LogAction "WARNING: bad layer count: " & m_numOfLayers
        End If
    End If
    
    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage3_LoadLayersAndChannels = False
    
End Function

'Only call from Import_Stage2_LoadProps!
Private Function Import_Stage2a_CacheKeyProps(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    Const FUNC_NAME As String = "Import_Stage2a_CacheKeyProps"
    Import_Stage2a_CacheKeyProps = True
    
    On Error GoTo BrokenImage
    
    'GIMP embeds property lists for the image, individual layers, and individual channels.  Some properties
    ' are restricted to only a specific subset; others are available anywhere.
    
    'This function only traverses and caches key properties from the IMAGE property list.  These are particularly
    ' crucial to cache, because they can affect the subsequent decoding of layers and/or channels.
    Dim tmpStream As pdStream
    Set tmpStream = New pdStream
    
    'Let's start with the global palette, if any.
    Dim idxProp As Long
    idxProp = GetIndexOfProperty(xcf_PROP_COLORMAP, m_ImageProperties)
    If (idxProp >= 0) Then
        
        'Wrap a pdStream object around the property array for convenient traversal
        tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, startingBufferSize:=m_ImageProperties(idxProp).propSize, baseFilePointerOffset:=VarPtr(m_ImageProperties(idxProp).propData(0))
        
        'First 4 bytes are color count
        m_numPaletteColors = tmpStream.ReadLong_BE()
        ReDim m_Palette(0 To m_numPaletteColors - 1) As RGBQuad
        
        'Colors are stored as RGB trios - NOT quads - so we need to manually traverse the struct
        Dim i As Long
        For i = 0 To m_numPaletteColors - 1
            With m_Palette(i)
                .Red = tmpStream.ReadByte()
                .Green = tmpStream.ReadByte()
                .Blue = tmpStream.ReadByte()
                .Alpha = 255    'Indexed images can use alpha, but it's encoded as a separate channel
            End With
        Next i
        
        'Free the temporary stream (before potential reuse)
        tmpStream.StopStream
        
    End If
    
    'Next, image resolution
    m_imageResolutionPPI = 96!
    
    idxProp = GetIndexOfProperty(xcf_PROP_RESOLUTION, m_ImageProperties)
    If (idxProp >= 0) Then
    
        tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, startingBufferSize:=m_ImageProperties(idxProp).propSize, baseFilePointerOffset:=VarPtr(m_ImageProperties(idxProp).propData(0))
        
        'Floats need to be converted from BE to LE before we can interpret them
        Dim srcBytes() As Byte
        tmpStream.ReadBytes srcBytes, 8, True
        VBHacks.SwapEndianness32 srcBytes
        
        Dim hRes As Single, vRes As Single
        VBHacks.GetMem4_Ptr VarPtr(srcBytes(0)), VarPtr(hRes)
        VBHacks.GetMem4_Ptr VarPtr(srcBytes(4)), VarPtr(vRes)
        
        m_imageResolutionPPI = (hRes + vRes) / 2!
        
        tmpStream.StopStream
        
    End If
    
    'Next, ICC profile.  This one must be manually retrieved, because it's stored as a parasite - NOT a property.
    ' (Retrieving it will require us to traverse a separate parasite tree.)
    idxProp = GetIndexOfProperty(xcf_PROP_PARASITES, m_ImageProperties)
    If (idxProp >= 0) Then
    
        tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, startingBufferSize:=m_ImageProperties(idxProp).propSize, baseFilePointerOffset:=VarPtr(m_ImageProperties(idxProp).propData(0))
        
        'Parasites are their own sub-tree within the larger "parasite" property.  They use string IDs
        ' and are individual, variable-length custom storage items.
        Do While tmpStream.GetPosition() < tmpStream.GetStreamSize() - 4
            
            'Parasite names are encoded as a "GIMP string", which are UTF-8 strings preceded by a 4-byte length
            Dim strParasite As String, lenName As Long
            lenName = tmpStream.ReadLong_BE()    'this value *includes* a null-terminator, so string itself is [n-1] chars
            If (lenName > 0) Then strParasite = tmpStream.ReadString_UTF8(lenName - 1)
            
            '+1 for null terminator
            tmpStream.SetPosition 1, FILE_CURRENT
            
            '4-byte flags follow (unnecessary) followed by 4-byte length (extremely necessary)
            Dim sizeOfData As Long
            tmpStream.ReadLong_BE
            sizeOfData = tmpStream.ReadLong_BE()
                
            Dim origOffset As Long
            origOffset = tmpStream.GetPosition
            
            If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "Found image parasite: " & strParasite & " (" & sizeOfData & " bytes)"
            If (LCase$(strParasite) = "icc-profile") Then
                
                If (sizeOfData > 0) Then
                    If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "Found ICC profile: " & sizeOfData & " bytes"
                    Set m_Profile = New pdICCProfile
                    m_Profile.LoadICCFromPtr sizeOfData, tmpStream.ReadBytes_PointerOnly(sizeOfData)
                End If
                
                'We could exit the loop here, but I'm leaving it for now in case we want to utilize
                ' other parasites in the future
                'Exit Do
                
            End If
            
            tmpStream.SetPosition origOffset, FILE_BEGIN
            tmpStream.SetPosition sizeOfData, FILE_CURRENT
            
        Loop
        
        tmpStream.StopStream
        
    End If
    
    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage2a_CacheKeyProps = False
    
End Function

'Import step 2: load all image "properties".  After this step, the stream pointer will be aligned with the
' start of the layer segment.
Private Function Import_Stage2_LoadProps(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    'Unlike some import functions, we assume success here (because the only way to fail
    ' is a malformed file)
    Const FUNC_NAME As String = "Import_Stage2_LoadProps"
    Import_Stage2_LoadProps = True
    
    On Error GoTo BrokenImage
    
    'Properties are basically a list of chunks with numeric IDs.  Each chunk self-describes
    ' its own length (although some historic XCF files may not report size correctly - I'm not sure
    ' how to address this and the spec is unclear, so this is TODO pending further testing)/
    
    'In this step, we basically just want to iterate all properties and copy them into a local collection.
    Dim curPropID As xcf_PropertyID, curPropSize As Long
    curPropID = m_Stream.ReadLong_BE()
    
    'Start with an arbitrary collection size
    Const INIT_PROP_LIST_COUNT As Long = 16
    ReDim m_ImageProperties(0 To INIT_PROP_LIST_COUNT - 1) As xcf_Property
    
    Do While (curPropID <> xcf_PROP_END)
        
        'Ensure we have room for this property, then store the new ID
        If (m_numImageProperties > UBound(m_ImageProperties)) Then ReDim Preserve m_ImageProperties(0 To m_numImageProperties * 2 - 1) As xcf_Property
        m_ImageProperties(m_numImageProperties).propID = curPropID
        
        'After ID comes property length
        curPropSize = m_Stream.ReadLong_BE()
        m_ImageProperties(m_numImageProperties).propSize = curPropSize
        
        'Special handling for legacy properties with potentially invalid payload sizes follows.
        Dim ptrOrig As Long
        ptrOrig = m_Stream.GetPosition()
        
            'The XCF spec specifically calls out image palettes as a place to watch for bad length markers.
            ' Specifically, it says:
            ' "Beware that the payload length of the PROP_COLORMAP in particular cannot be trusted: some historic
            '  releases of GIMP erroneously wrote n+4 instead of 3*n+4 into the length word (but still actually
            '  followed it by 3*n+4 bytes of payload)."
            '
            'For this property, we want to manually retrieve the number of colors, and use the larger of 3*n+4
            ' and the actual encoded property.
            If (curPropID = xcf_PROP_COLORMAP) Then
                curPropSize = PDMath.Max2Int(m_Stream.ReadLong_BE() * 3 + 4, curPropSize)
                m_ImageProperties(m_numImageProperties).propSize = curPropSize
                m_Stream.SetPosition ptrOrig, FILE_BEGIN
            End If
            
        'After ID comes payload
        If (curPropSize > 0) Then m_Stream.ReadBytes m_ImageProperties(m_numImageProperties).propData, curPropSize, True
        
        'Debug info (please disable in production)
        If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "Found image property: " & GetPropertyName(curPropID) & " (" & curPropSize & " bytes)"
        
        'Increment total image property count
        m_numImageProperties = m_numImageProperties + 1
        
        'Read next property ID
        curPropID = m_Stream.ReadLong_BE()
        
        'Failsafe check for EOF (which usually means something catastrophically bad happened)
        If (m_Stream.GetPosition >= m_Stream.GetStreamSize) Then
            Import_Stage2_LoadProps = False
            Exit Do
        End If
        
    Loop
    
    'PROP_END has ID 0, but it is still followed by a length indicator (always 0).  We have not yet
    ' read the length indicator, so we must do so now - and we also need to validate that it is, indeed, 0.
    If Import_Stage2_LoadProps Then Import_Stage2_LoadProps = (m_Stream.ReadLong_BE() = 0)
    
    'After successful traversal of the full property list, we want to cache some key properties in useable structs.
    If Import_Stage2_LoadProps Then Import_Stage2_LoadProps = Import_Stage2a_CacheKeyProps(srcFile, dstImage, dstDIB)
    
    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage2_LoadProps = False
    
End Function

'Import step 1: after "magic number" validation, parse the short XCF header and validate core members
' like image width/height
Private Function Import_Stage1_ParseHeader(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "Import_Stage1_ParseHeader"
    Import_Stage1_ParseHeader = False
    
    On Error GoTo BrokenImage
    
    'Magic number is followed by version (as a string).  We remap this to a long for convenience.
    Dim strVersion As String
    strVersion = LCase$(m_Stream.ReadString_ASCII(4))
    
    If (strVersion = "file") Then
        m_xcfVersion = 0
    Else
        m_xcfVersion = CLng(Right$(strVersion, 3))
    End If
    
    If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "XCF version is " & m_xcfVersion
    
    '1-byte null-padding follows
    m_Stream.ReadByte
    
    'Canvas width/height
    m_ImageWidth = m_Stream.ReadLong_BE
    m_ImageHeight = m_Stream.ReadLong_BE
    
    'Color model
    m_imageColorMode = m_Stream.ReadLong_BE
    
    'If file is v4 or later, a 4-byte precision value follows next.  (For earlier versions,
    ' this field is omitted and a precision of "8-bit gamma integer" is assumed.)
    If (m_xcfVersion >= 4) Then
        
        Dim tmpInt As Long
        tmpInt = m_Stream.ReadLong_BE
        
        'This field is mapped weirdly; see the spec for details
        Select Case tmpInt
            
            Case 100
                m_imagePrecision = xcf_08bitIntLinear
            Case 150
                m_imagePrecision = xcf_08bitIntGamma
            Case 200
                m_imagePrecision = xcf_16bitIntLinear
            Case 250
                m_imagePrecision = xcf_16bitIntGamma
            Case 300
                m_imagePrecision = xcf_32bitIntLinear
            Case 350
                m_imagePrecision = xcf_32bitIntGamma
            Case 500
                m_imagePrecision = xcf_16bitFltLinear
            Case 550
                m_imagePrecision = xcf_16bitFltGamma
            Case 600
                m_imagePrecision = xcf_32bitFltLinear
            Case 650
                m_imagePrecision = xcf_32bitFltGamma
            Case 700
                m_imagePrecision = xcf_64bitFltLinear
            Case 750
                m_imagePrecision = xcf_64bitFltGamma
            Case Else
                InternalError FUNC_NAME, "unknown image precision: " & tmpInt
                Exit Function
        End Select
        
    Else
        m_imagePrecision = xcf_08bitIntGamma
    End If
    
    'And that's it for the header!  Do a few basic validations before continuing.
    If (m_ImageWidth <= 0) Or (m_ImageHeight <= 0) Then
        InternalError FUNC_NAME, "bad image width/height"
        Exit Function
        
    ElseIf (m_ImageWidth >= 65535) Or (m_ImageHeight >= 65535) Then
        InternalError FUNC_NAME, "image width/height too big"
        Exit Function
    
    ElseIf (m_imageColorMode < 0) Or (m_imageColorMode > 2) Then
        InternalError FUNC_NAME, "bad color mode: " & m_imageColorMode
        Exit Function
    End If
    
    'Still here?  Header looks okay!
    If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "XCF is " & m_ImageWidth & "x" & m_ImageHeight & " with color mode " & m_imageColorMode & " and precision " & m_imagePrecision
    Import_Stage1_ParseHeader = True
    
    Exit Function
    
BrokenImage:
    InternalError FUNC_NAME, "Internal VB error #" & Err.Number & ": " & Err.Description
    Import_Stage1_ParseHeader = False
    
End Function

'Returns index >= 0 if the requested property exists in the collection; -1 otherwise.
' Note that the GIMP spec doesn't like you to write multiple copies of the same property,
' but it specifically describes the *possibility* of encountering multiple copies of a property.
' In these cases, you are supposed to take the later copy of the property, if any.  (With a
' specific exception for list-type properties, which should be concatenated - we leave those
' handling nuances to the caller, however.)  I mention the "multiple copies" state specifically
' to explain why we traverse the property list *backward*.
Private Function GetIndexOfProperty(ByVal propID As xcf_PropertyID, ByRef srcProperties() As xcf_Property) As Long
    
    GetIndexOfProperty = -1
    
    Dim i As Long
    For i = UBound(srcProperties) To 0 Step -1
        If (srcProperties(i).propID = propID) Then
            GetIndexOfProperty = i
            Exit Function
        End If
    Next i
    
End Function

Private Function GetPropertyName(ByVal propID As xcf_PropertyID) As String
    Select Case propID
        Case xcf_PROP_END
            GetPropertyName = "end"
        Case xcf_PROP_COLORMAP
            GetPropertyName = "color map"
        Case xcf_PROP_ACTIVE_LAYER
            GetPropertyName = "active layer"
        Case xcf_PROP_ACTIVE_CHANNEL
            GetPropertyName = "active channel"
        Case xcf_PROP_SELECTION
            GetPropertyName = "selection"
        Case xcf_PROP_FLOATING_SELECTION
            GetPropertyName = "floating selection"
        Case xcf_PROP_OPACITY
            GetPropertyName = "opacity"
        Case xcf_PROP_BLEND_MODE
            GetPropertyName = "blend mode"
        Case xcf_PROP_VISIBLE
            GetPropertyName = "visibility"
        Case xcf_PROP_LINKED
            GetPropertyName = "linked"
        Case xcf_PROP_LOCK_ALPHA
            GetPropertyName = "lock alpha"
        Case xcf_PROP_APPLY_MASK
            GetPropertyName = "apply mask"
        Case xcf_PROP_EDIT_MASK
            GetPropertyName = "edit mask"
        Case xcf_PROP_SHOW_MASK
            GetPropertyName = "show mask"
        Case xcf_PROP_SHOW_MASKED
            GetPropertyName = "show masked"
        Case xcf_PROP_OFFSETS
            GetPropertyName = "offsets"
        Case xcf_PROP_COLOR
            GetPropertyName = "color"
        Case xcf_PROP_COMPRESSION
            GetPropertyName = "compression"
        Case xcf_PROP_GUIDES
            GetPropertyName = "guides"
        Case xcf_PROP_RESOLUTION
            GetPropertyName = "resolution"
        Case xcf_PROP_TATTOO
            GetPropertyName = "tattoo"
        Case xcf_PROP_PARASITES
            GetPropertyName = "parasite"
        Case xcf_PROP_UNIT
            GetPropertyName = "unit"
        Case xcf_PROP_PATHS
            GetPropertyName = "paths"
        Case xcf_PROP_USER_UNIT
            GetPropertyName = "user unit"
        Case xcf_PROP_VECTORS
            GetPropertyName = "vectors"
        Case xcf_PROP_TEXT_LAYER_FLAGS
            GetPropertyName = "text layer flags"
        Case xcf_PROP_LOCK_CONTENT
            GetPropertyName = "lock content"
        Case xcf_PROP_GROUP_ITEM
            GetPropertyName = "group item"
        Case xcf_PROP_ITEM_PATH
            GetPropertyName = "item path"
        Case xcf_PROP_GROUP_ITEM_FLAGS
            GetPropertyName = "item flags"
        Case xcf_PROP_LOCK_POSITION
            GetPropertyName = "lock position"
        Case xcf_PROP_FLOAT_OPACITY
            GetPropertyName = "float opacity"
        Case xcf_PROP_COLOR_TAG
            GetPropertyName = "color tag"
        Case xcf_PROP_COMPOSITE_MODE
            GetPropertyName = "composite mode"
        Case xcf_PROP_COMPOSITE_SPACE
            GetPropertyName = "composite space"
        Case xcf_PROP_BLEND_SPACE
            GetPropertyName = "blend space"
        Case xcf_PROP_FLOAT_COLOR
            GetPropertyName = "float color"
        Case xcf_PROP_SAMPLE_POINTS
            GetPropertyName = "sample points"
        Case xcf_PROP_ITEM_SET
            GetPropertyName = "item set"
        Case xcf_PROP_ITEM_SET_ITEM
            GetPropertyName = "item set item"
        Case xcf_PROP_LOCK_VISIBILITY
            GetPropertyName = "lock visibility"
    End Select
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdXCF." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdXCF." & funcName & "() reported an error: " & errDescription
    End If
End Sub
