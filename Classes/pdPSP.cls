VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PSP (PaintShop Pro) Container and Parser
'Copyright 2020-2026 by Tanner Helland
'Created: 30/December/20
'Last updated: 04/February/21
'Last update: get ICC profile embedding working when exporting
'
'This class (and its associated pdPSP- child classes) handle import and export of PSP
' (JASC/Corel Paint Shop Pro) image files.  All code here has been custom-built for PhotoDemon,
' with a particular emphasis on parsing performance.  This class will often outperform PSP's own
' file parser, especially on images with vector layers.
'
'As with all 3rd-party PSP engines, Paint Shop Pro has many features that don't have direct analogs
' in PhotoDemon.  Such features are still extracted by this class, but they may not "appear" in the
' final loaded image.  My ongoing goal is to expand support in this class as various PSP features are
' implemented in PD itself.
'
'I've based all work in this class off the "official" PSP spec at this URL (link good as of December 2020):
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

'When debugging, you can use this constant to have PD load the composite (merged) image to its
' own layer.  I find this helpful for confirming that the multi-layer projection matches the
' saved composite image.
Private Const PSP_LOAD_COMPOSITE_AS_BASE_LAYER As Boolean = False

'Until PD supports layer groups, we have to resort to workarounds for PSP layer group information.
' When loading layers, PD will automatically handle things like group opacity (by hiding all members
' of a layer group if said group is invisible), but as for the groups themselves, there are two options.
' The easiest option is to just ignore the groups and lose them when importing into PD.  The other
' option is to mimic groups by importing them as "dummy" layers.  This gives the user a way to
' rearrange them (and see them, if nothing else), and we can use a specific naming scheme to ensure
' we preserve group order at export time.  Toggling these behaviors is controlled by this constant,
' with a matching bool to track whether an exported image contains layers.  (If it does, we need to
' modify some writing behavior.)
Private Const PSP_GROUPS_AS_DUMMY_LAYERS As Boolean = False
Private m_ExportGroups As Boolean

'Like many other formats, PSP files consist of "blocks".  Blocks have hard-coded 2-byte IDs.
Private Const PSP_BLOCK_MARKER As Long = &H4B427E

'/* Type of image in the composite image bank block.  */
Private Enum PSPCompositeImageType
    PSP_IMAGE_COMPOSITE = 0 '// Composite Image
    PSP_IMAGE_THUMBNAIL     '// Thumbnail Image
End Enum

#If False Then
    Private Const PSP_IMAGE_COMPOSITE = 0, PSP_IMAGE_THUMBNAIL = 1
#End If

'PSP images are ultimately comprised of a collection of blocks.  These are analagous to
' something like "chunks" in a PNG file (although PSP "blocks" can contain sub-blocks and
' chunks within themselves, which is confusing to say the least).  As an early parse step,
' we quickly scan all blocks and cache their IDs and offsets.  This is helpful for both
' validation purposes, and general parsing purposes, because it lets us operate at a
' more helpful "block" level.
Private m_numOfBlocks As Long, m_Blocks() As pdPSPBlock
Private m_ImageHeader As PSPImageHeader

'When exporting PSP files, we set a module-level version flag so we can adjust write behavior
' as necessary.
Private m_ExportVersion As Long

'Files from PSPv5 use a different block header format.  This causes a lot of parsing headaches
' so we set a flag up-front and refer to it when reading block headers.
Private m_PSPisV5 As Boolean

'PSP files may embed one or more "composite" images, which are either thumbnails or
' copies of the merged image.  PD can retrieve these, and under certain circumstances it may
' preferentially choose to use the composite image instead of individual layers
' (explanation TBD pending true layer support)
Private Type PSPCompositeImage
    cih_HeaderSize As Long
    cih_Width As Long
    cih_Height As Long
    cih_BitDepth As Long                        'must be 1, 4, 8, 24, or 48
    cih_Compression As PSPCompression           'CAN be JPEG-compressed (unlike image layers)
    cih_PlaneCount As Long                      'must be 1
    cih_ColorCount As Long                      '2 ^ bit-depth
    cih_CompositeType As PSPCompositeImageType  'Used to interpret composite contents
    
    'The spec allows for future header expansion, so the top SIZE member is critical for correct reading;
    ' don't assume a fixed size!
    
    'This member is PD-specific, and is simply used to flag each composite entry as
    ' valid or invalid (invalid means something in the header looks awry)
    cih_LooksOK As Boolean
    
    'If commanded to, the composite image parser can retrieve the associated composite image
    ' and decode it to a standard pdDIB object.  Do *not* attempt to use the DIB unless the
    ' .cih_DIBOK value is TRUE.
    cih_DIB As pdDIB
    cih_DIBOK As Boolean
    
End Type

Private m_numCompositeImages As Long, m_CompositeImages() As PSPCompositeImage

'Composite images are allowed to use their own custom palette, separate from the parent image's palette
Private m_CompositePalette() As RGBQuad, m_CompositePaletteCount As Long

'Indexed images will provide a "parent" palette for the full image.
' If one exists, m_ParentPaletteCount will be some value > 0.
Private m_ParentPalette() As RGBQuad, m_ParentPaletteCount As Long

'Central layer collection.  The total number of layers is mirrored by the image header, but note
' that PhotoDemon may choose to "deactivate" layers it doesn't understand - so post-parsing,
' it's critical to defer to this layer count instead of the one embedded in the image header.
Private m_LayerCount As Long, m_Layers() As pdPSPLayer

'If an ICC profile exists, it will be loaded into this object
Private m_ICCProfile As pdICCProfile

'Byte-by-byte access is provided, as always, by a pdStream instance
Private m_Stream As pdStream

'At present, we require the caller to pass an identical source file path to every load function.
' (This is a cheap and easy way to ensure no funny business.)  If the PSP is loaded directly from memory,
' we flag this with a special name.
Private m_SourceFilename As String
Private Const PSP_LOADED_FROM_MEMORY As String = "LoadFromPtr*"
Private m_SourcePtr As Long, m_SourcePtrLen As Long

'If warnings are encountered during processing, I push their messages onto a string stack.  (I may
' decide to report these to the user... but haven't decided yet.  Either way, it's very helpful
' while debugging; see associated Warnings_XYZ functions for details.)
Private m_Warnings As pdStringStack

Friend Function GetICCProfile() As pdICCProfile
    Set GetICCProfile = m_ICCProfile
End Function

Friend Function GetOriginalColorDepth() As Long
    GetOriginalColorDepth = m_ImageHeader.psph_BitDepth
    If (GetOriginalColorDepth > 8) And Me.HasAlpha() Then
        If (GetOriginalColorDepth = 24) Then
            GetOriginalColorDepth = 32
        ElseIf (GetOriginalColorDepth = 48) Then
            GetOriginalColorDepth = 64
        End If
    End If
End Function

Friend Function HasAlpha() As Boolean
    HasAlpha = ((m_ImageHeader.psph_ContentFlags And keGCAlphaChannels) <> 0)
End Function

Friend Function HasICCProfile() As Boolean
    HasICCProfile = (Not m_ICCProfile Is Nothing)
End Function

'Perform basic validation on a potential PSP file.  For strict file extension matching (e.g. only
' .psp or .pspimage extension allowed), set checkExtension to TRUE.
Friend Function IsFilePSP(ByRef srcFile As String) As Boolean
    IsFilePSP = (Step1_ValidateHeader(srcFile) < psp_Failure)
End Function

Friend Function IsPSPGrayscale() As Boolean
    IsPSPGrayscale = m_ImageHeader.psph_IsGrayscale
End Function

'Simplified wrapper to load a PSD automatically.
Friend Function LoadPSP(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal loadFromPtr As Long = 0, Optional ByVal loadFromPtrLen As Long = 0) As PD_PSPResult
    
    'Reset some internal parameters to ensure subsequent reads are accurate.
    ' (This is critical if multiple PSPs are read back-to-back.)
    Me.Reset
    
    'We support PSD loading from both file and memory; if the passed loadFromPtr value is non-zero,
    ' treat it as a pointer and we'll wrap our stream around it instead.
    If (loadFromPtr <> 0) And (loadFromPtrLen <> 0) Then
        srcFile = PSP_LOADED_FROM_MEMORY
        m_SourcePtr = loadFromPtr
        m_SourcePtrLen = loadFromPtrLen
    Else
        m_SourcePtr = 0
        m_SourcePtrLen = 0
    End If

    'Try to validate the source file
    Dim keepLoading As PD_PSPResult
    keepLoading = Step1_ValidateHeader(srcFile)
    
    'If this identifies as a PSP file, build and validate the block directory.
    ' (Otherwise, exit immediately.)
    If (keepLoading < psp_Failure) Then
        keepLoading = Step2_BuildBlockDirectory(srcFile)
    Else
        LoadPSP = False
    End If
    
    'If the block directory looks good, parse the general image header.
    If (keepLoading < psp_Failure) Then keepLoading = Step3_ReadImageAttributes(srcFile)
    
    'If the block directory looks good, load the attached ICC profile (if any - they're rare in PSP files)
    If (keepLoading < psp_Failure) Then keepLoading = Step3a_ReadICCProfile()
    
    'TODO: extended data block (need to figure out what we want to retrieve from it)
    'TODO: tube data block (need relevant image(s) for testing)
    'TODO: creator data block (may be able to rely on ExifTool for now?)
    
    'Next, we previously examined the Composite Image block.  Per v8 of the spec,
    ' "All PSP files with more than one layer or with one layer with transparency
    '  must have a full size composite image."
    '
    'Note that single-opaque-layer images are *not* required to have a composite block.
    '
    'Note also that earlier PSP versions don't describe a composite block at all (v5),
    ' and from v6 onward the composite bank underwent various modifications.  PD accounts
    ' for these subtle differences and will simply return a warning (*not* a failure)
    ' if something goes wrong when reading the composite block.  You can keep parsing
    ' the rest of the image just fine, regardless of composite image state.
    '
    'With all that said, now that PhotoDemon is perfectly capable of importing individual
    ' PSP layers, it no longer parses the composite block except as a last-resort failsafe
    ' (if it can't find any useable layers, for example), or if you've activated the debug-only
    ' "import merged image as layer" constant at the top of this class.  If those conditions
    ' are met, the composite image retrieval will trigger later in the load process, after the
    ' importer detects that a composite block *is* actually wanted.  (Obviously, skipping the
    ' import of a giant image-sized layer can greatly improve import time.)
    'If (keepLoading < psp_Failure) Then keepLoading = Step4_ReadCompositeBlock(srcFile, False)
    
    'Next, if the image is < 24-bpp and non-grayscale, we need to retrieve the parent palette.
    ' (At this point, note that block order is no longer guaranteed - so the palette block could
    ' lie anywhere inside the file.)
    If (keepLoading < psp_Failure) Then keepLoading = Step5_ReadParentPalette(srcFile)
    
    'With the palette read, it's time to parse layers!
    If (keepLoading < psp_Failure) Then keepLoading = Step6_LoadLayers(srcFile)
    
    'With all layers loaded, we can finally assemble the final image.  Normally, this will consist
    ' of all layers from the file that PD knows how to interpret.  If errors were encountered
    ' during import, however, PD may decide to fall back to the embedded composite image (if one
    ' exists) to provide a more useable image to the user.
    If (keepLoading < psp_Failure) Then keepLoading = Step7_AssembleFinalImage(srcFile, dstImage, dstDIB)
    
    LoadPSP = keepLoading
    
    'Even if the PSP loaded, one (or more) surprises may have affected our ability to render
    ' the file correctly.  Regardless of this class's debug verbosity setting, warnings always
    ' get dumped out to the debug log.  This is especially useful when new versions of PaintShop Pro
    ' release, as they often bring new parsing surprises/nightmares.
    If (Me.Warnings_GetCount() > 0) Then
        PDDebug.LogAction "PhotoDemon's PSP parser generated one or more warnings.  Here is a full list:"
        Me.Warnings_DumpToDebugger
    End If
    
End Function

Friend Sub Reset()
    Set m_Warnings = New pdStringStack
End Sub

'Save a new PSP file to disk.  (With minor modifications, this could also be used to save to a memory stream,
' but this is *not* currently implemented.)
'
'If the destination file exists, it will be forcibly overwritten.
' Safe save measures should be implemented externally.
Friend Function SavePSP(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal targetPSPVersion As Long = 8, Optional ByVal writeCompositeImage As Boolean = True, Optional ByVal dstCompressionLevel As Long = -1, Optional ByVal writeICCProfile As Boolean = True) As Boolean
    
    On Error GoTo CouldNotSaveFile
    
    If (srcPDImage Is Nothing) Or (LenB(dstFile) = 0) Then
        InternalError "SavePSP", "invalid function parameters"
        SavePSP = False
        Exit Function
    End If
    
    'Failsafe check only; PD enforces safe overwriting in the parent function.
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'So, a few quick notes before we begin.
    
    '1) Saving a PSP file is much easier than loading it, primarily because we don't need to cover every
    ' color-mode and bit-depth option.  Instead, we focus on exporting 32-bit RGBA data in the cleanest
    ' manner possible.  Other bit-depths and/or color modes may be covered in future updates.
    
    '2) PSPs use many structs with non-DWORD alignments.  This causes issues in VB6 as struct members
    ' are typically aligned to DWORD boundaries (with some caveats; it's complicated and basically don't
    ' worry about it if you aren't a VB6 developer).  To avoid complications or convoluted struct-shifting,
    ' this function manually writes out most struct members individually.  This also aids in writing
    ' struct members that are defined differently in VB, like bools (which are bytes with 0/1 values in a
    ' PSP file).
    
    '3) This code was designed against the latest public PSP spec at the time of its writing
    ' (v8 from October 2006, available here: ftp://ftp.corel.com/pub/documentation/PSP/).  I make no
    ' guarantee that this code is compatible with earlier or later spec versions, although attempts have
    ' been made to support PSP files back to v6 and through the most recent modern version at the time of
    ' this writing (v "2021").
    
    'All good?  Then let's do this!
    
    'Start by ensuring that incoming flags match target version capabilities.
    If (targetPSPVersion > 8) Then
        targetPSPVersion = 8
    ElseIf (targetPSPVersion < 6) Then
        targetPSPVersion = 6
    End If
    
    m_ExportVersion = targetPSPVersion
    
    'Color profiles weren't supported until v8
    If (targetPSPVersion < 8) Then writeICCProfile = False
    
    Dim i As Long
    
    'Before exporting, we need to run some quick heuristics on exported layer data.  This approach will
    ' likely change as PD is able to reproduce more PSP behavior, but for now, we sometimes need to
    ' perform weird conversions between PD data and PSD data.
    
    'Scan layer names in the source image for group identifiers.  If we find any, set a module-level
    ' flag to note that layer groups need to be exported.  (They require additional header information
    ' and optional chunks.)  Note that group export behavior can be controlled by a module-level constant.
    m_ExportGroups = False
    If PSP_GROUPS_AS_DUMMY_LAYERS Then
        
        'When importing PSP files, layer groups are imported as "dummy" 1x1 layers with special names.
        Dim targetStartName As String, targetEndName As String
        targetStartName = g_Language.TranslateMessage("Group start:")
        targetEndName = g_Language.TranslateMessage("Group end:")
        
        Dim startTargetsFound As Long: startTargetsFound = 0
        Dim endTargetsFound As Long: endTargetsFound = 0
        
        For i = 0 To srcPDImage.GetNumOfLayers - 1
            If Strings.StringsEqual(Left$(srcPDImage.GetLayerByIndex(i).GetLayerName(), Len(targetStartName)), targetStartName, True) Then
                startTargetsFound = startTargetsFound + 1
            Else
                If Strings.StringsEqual(Left$(srcPDImage.GetLayerByIndex(i).GetLayerName(), Len(targetEndName)), targetEndName, True) Then endTargetsFound = endTargetsFound + 1
            End If
        Next i
        
        'We will export groups if:
        ' 1) At least one group is found in the document, and...
        ' 2) There are the same number of group start markers as group end markers.  (A mismatch in these
        '    values means the resulting PSP would be malformed, and we *never* want to risk that.)
        If (startTargetsFound > 0) And (endTargetsFound > 0) Then
            m_ExportGroups = (startTargetsFound = endTargetsFound)
        End If
        
    End If
    
    'As usual, all writing is handled by a pdStream instance.  Note that we do not request typical
    ' sequential perf-optimized mode - this is because PSPs use a ton of "length" entries that describe
    ' following data segments.  Lengths are not easily known in advance, so the writer must periodically
    ' "retreat" and update previous length entries.  This incurs potentially large penalties when
    ' sequential mode is requested, and random-access shows no meaningful benefit, so we stick to
    ' default caching behavior.
    Dim cStream As pdStream
    Set cStream = New pdStream
    
    Dim initBufferSize As Double
    initBufferSize = srcPDImage.EstimateRAMUsage()
    If (initBufferSize > CDbl(LONG_MAX)) Then initBufferSize = 0#
    cStream.StartStream PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFile, Int(initBufferSize)
    
    'Start by writing out a basic PSP header.
    Dim keepGoing As PD_PSPResult
    keepGoing = ExportStep1_FileHeaderAndGeneralAttributes(cStream, srcPDImage, writeCompositeImage)
    
    'Next comes the ICC profile.
    If (keepGoing < psp_Failure) And writeICCProfile Then keepGoing = ExportStep2_ColorProfile(cStream, srcPDImage)
    
    'Next, we need to write a composite image (merged image).  The spec requires this for all
    ' multi-layer or alpha-supporting images, e.g. "All PSP files with more than one layer or
    ' with one layer with transparency must have a full size composite image."
    If (keepGoing < psp_Failure) And writeCompositeImage Then keepGoing = ExportStep3_CompositeImage(cStream, srcPDImage, dstCompressionLevel)
    
    'Next comes the layer bank block
    If (keepGoing < psp_Failure) Then keepGoing = ExportStep4_LayerBank(cStream, srcPDImage, dstCompressionLevel)
    
    'The write is finished.  Close the file handle and exit.
    cStream.StopStream
    Set cStream = Nothing
    
    SavePSP = (keepGoing < psp_Failure)
    
    Exit Function
    
CouldNotSaveFile:
    InternalError "SavePSP", "Internal VB error #" & Err.Number & ": " & Err.Description
    SavePSP = False
    
End Function

Private Function ExportStep1_FileHeaderAndGeneralAttributes(ByRef cStream As pdStream, ByRef srcImage As pdImage, ByVal writeCompositeImage As Boolean) As PD_PSPResult

    On Error GoTo InternalVBError
    Const funcName As String = "ExportStep1"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'As noted previously, VB6 struct alignment issues prevent us from easily writing out structs "as-is";
    ' instead, we write out members one-at-a-time.
    
    'Start with the "magic number" 32-byte PSP identifier.
    Dim pspID() As Byte
    If Strings.BytesFromHex(pspID, "5061696E742053686F702050726F20496D6167652046696C650A1A0000000000") Then
        cStream.WriteBytesFromPointer VarPtr(pspID(0)), 32
    Else
        InternalError funcName, "BytesFromHex failed?"
        okToProceed = psp_Failure
    End If
    
    If (okToProceed < psp_Failure) Then
        
        'Major/minor versions follow.  The version number is always 2 less than the target PSP version,
        ' starting with PSP 5 as the base (with a corresponding version number of 3).  PD only supports
        ' versions for which a spec is publicly available, so v8 is the latest format we write.
        cStream.WriteInt m_ExportVersion - 2
        cStream.WriteInt 0
        
    End If
    
    'File header is complete.  Exit now if any failures were encountered.
    If (okToProceed >= psp_Failure) Then Exit Function
    
    'Next comes the "general image attributes" block.  This is written as a normal PSP block,
    ' with a prefixed header ID, block ID, and block length.  We obviously cannot write the
    ' block length until the header itself has been written, so we need to flag position in advance.
    cStream.WriteLong PSP_BLOCK_MARKER
    cStream.WriteInt PSP_IMAGE_BLOCK
    
    'Note stream position *immediately after the offset*; we'll return here to write the block length
    ' after it's written.
    cStream.WriteLong 0
    Dim posJustAfterBlockLength As Long
    posJustAfterBlockLength = cStream.GetPosition()
    
    'The block header is followed by a PSP "chunk", which has its own length marker (sigh).
    ' Chunks have different rules than blocks.  Chunk sizes include the 4-byte chunk length;
    ' block sizes do not.
    ' So this means, we can just mirror the block length, above - so I don't store a separate
    ' offset calculation here.  (When we return to write block length, we'll just write it 2x.)
    cStream.WriteLong 0
    
    'Image dimensions
    cStream.WriteLong srcImage.Width
    cStream.WriteLong srcImage.Height
    
    'Resolution value is "pixels per metric", with the metric being a unique ID that follows this value.
    If (srcImage.GetDPI <= 0#) Then
        cStream.WriteDouble 96#
    Else
        cStream.WriteDouble srcImage.GetDPI
    End If
    cStream.WriteByte PSP_METRIC_INCH
    
    'Compression type.  As with version descriptors, above, this could be user-controllable, but since
    ' zlib-compression provides far and away the best compression ratio, we always default to it currently.
    cStream.WriteInt PSP_COMP_LZ77
    
    'Bit-depth is currently always 32-bit, but PSP doesn't include alpha channels in this calculation
    cStream.WriteInt 24
    
    'Planes and color count are silly holdovers from old Windows bmp headers (must always be
    ' 1 and 2^bit-depth, respectively)
    cStream.WriteInt 1
    cStream.WriteLong 2 ^ 24
    
    'PD does not currently write grayscale-format images (bool byte)
    cStream.WriteByte 0
    
    'Total image size is sum of the sizes of all layer color bitmaps.  The spec doesn't describe what the
    ' purposes of this is (or how to calculate it), so I assume it's another nonsense Windows BMP holdover.
    ' A cursory examination of PSPv6 files shows that it appears to use the formula of [3 * w * h] for
    ' each layer, summed together over the full image.  I don't know if non-raster layers should appear here.
    '
    'Regardless, the value we write doesn't appear to affect compatibility with PSP itself, so this doesn't
    ' seem especially critical.
    Dim i As Long, sumOfLayerContents As Long
    For i = 0 To srcImage.GetNumOfLayers - 1
        sumOfLayerContents = sumOfLayerContents + 3 * srcImage.GetLayerByIndex(i).GetLayerWidth(True) * srcImage.GetLayerByIndex(i).GetLayerHeight(True)
    Next i
    cStream.WriteLong sumOfLayerContents
    
    'Active layer index
    cStream.WriteLong srcImage.GetActiveLayerIndex
    
    'Total layer count
    cStream.WriteIntU srcImage.GetNumOfLayers
    
    'Content flags
    Dim imgFlags As PSPGraphicContents
    imgFlags = keGCAlphaChannels Or keGCRasterLayers
    If writeCompositeImage Then imgFlags = imgFlags Or keGCComposite Or keGCCompositeTransparency
    cStream.WriteLong imgFlags
    
    'Future expansion fields are possible.  For now, calculate size of the header up to this point.
    Dim posEndOfHeader As Long, curSizeOfHeader As Long
    posEndOfHeader = cStream.GetPosition()
    curSizeOfHeader = posEndOfHeader - posJustAfterBlockLength
    
    'Return to the block header and write the current size *twice* (once as the block size, then again
    ' as the prefixed chunk size)
    cStream.SetPosition posJustAfterBlockLength - 4
    cStream.WriteLong curSizeOfHeader
    cStream.WriteLong curSizeOfHeader
    
    'Restore original stream alignment
    cStream.SetPosition posEndOfHeader
    
    'Header complete, and the struct is correctly aligned for the next block
    ExportStep1_FileHeaderAndGeneralAttributes = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    m_Warnings.AddString "Internal error in export step 1, #" & Err.Number & ": " & Err.Description
    ExportStep1_FileHeaderAndGeneralAttributes = psp_Failure
    
End Function

Private Function ExportStep2_ColorProfile(ByRef cStream As pdStream, ByRef srcImage As pdImage) As PD_PSPResult

    On Error GoTo InternalVBError
    Const funcName As String = "ExportStep2"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'We are 100% reliant on previous steps aligning the stream pointer correctly.  PSP files do not provide
    ' an easy mechanism for guaranteeing stream alignment, so we are at the mercy of preceding steps.
    
    'Start by writing the bank block header.  This is written as a normal PSP block, with a prefixed
    ' header ID, block ID, and block length.  We obviously cannot write the block length until the block
    ' itself has been written, so we need to flag position in advance.
    cStream.WriteLong PSP_BLOCK_MARKER
    cStream.WriteInt PSP_COLORPROFILE_BLOCK
    
    'Note stream position *immediately after the offset*; we'll return here to write the block length
    ' after it's written.
    cStream.WriteLong 0
    Dim posJustAfterBlockLength As Long
    posJustAfterBlockLength = cStream.GetPosition()
    
    'Now we need to write a "color profile data chunk".
    
    'Chunk size is unknown until the chunk is actually written
    cStream.WriteLong 0
    
    '"Description" comes next (a silly field), but we try to write it correctly.
    '
    'Start by pulling a reference to the source image's active color profile
    Dim srcProfile As pdICCProfile
    Set srcProfile = ColorManagement.GetProfile_ByHash(srcImage.GetColorProfile_WorkingSpace())
    If (srcProfile Is Nothing) Then Set srcProfile = ColorManagement.GetProfile_Default()
    
    'We need to use LCMS to pull detailed information out of the profile
    Dim cProfile As pdLCMSProfile
    Set cProfile = New pdLCMSProfile
    If (Not cProfile.CreateFromPDICCObject(srcProfile)) Then cProfile.CreateSRGBProfile False
    Set srcProfile = Nothing
    
    Dim strProfileDescription As String
    strProfileDescription = Trim$(Strings.TrimNull(cProfile.GetInfoString(cmsInfoDescription)))
    
    'We can finally write the description out to file
    cStream.WriteIntU Len(strProfileDescription)
    If (LenB(strProfileDescription) > 0) Then cStream.WriteString_ASCII strProfileDescription
    
    'Next comes size of the profile
    Dim profBytes() As Byte, profSize As Long
    profSize = cProfile.GetRawProfileBytes(profBytes)
    cStream.WriteLong profSize
    
    'Now the spec is actually wrong here - it shows the ICC color profile being immediately
    ' written as part of the color profile data chunk, but this is apparently incorrect.
    ' Instead, we need to end the chunk here (and return to the start of the chunk to fill in
    ' the current length calculation), then follow the chunk with a separate, not-inside-a-chunk
    ' copy of the profile itself.
    Dim posEndOfChunk As Long
    posEndOfChunk = cStream.GetPosition()
    cStream.SetPosition posJustAfterBlockLength
    cStream.WriteLong posEndOfChunk - posJustAfterBlockLength
    cStream.SetPosition posEndOfChunk
    
    'Finally, write the profile itself
    If (profSize > 0) Then cStream.WriteBytesFromPointer VarPtr(profBytes(0)), profSize
    Set cProfile = Nothing
    
    'The color profile block is complete.  Return to the start of the block and fill in the final size.
    Dim posEndOfBlock As Long, finalSizeOfBlock As Long
    posEndOfBlock = cStream.GetPosition()
    finalSizeOfBlock = posEndOfBlock - posJustAfterBlockLength
    cStream.SetPosition posJustAfterBlockLength - 4
    cStream.WriteLong finalSizeOfBlock
    
    'Restore original stream alignment
    cStream.SetPosition posEndOfBlock
    
    'Color profile block complete, and the struct is correctly aligned for the next block
    ExportStep2_ColorProfile = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    m_Warnings.AddString "Internal error in export step 2, #" & Err.Number & ": " & Err.Description
    ExportStep2_ColorProfile = psp_Failure
    
End Function

Private Function ExportStep3_CompositeImage(ByRef cStream As pdStream, ByRef srcImage As pdImage, Optional ByVal dstCompressionLevel As Long = -1) As PD_PSPResult

    On Error GoTo InternalVBError
    Const funcName As String = "ExportStep3"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'We are 100% reliant on previous steps aligning the stream pointer correctly.  PSP files do not provide
    ' an easy mechanism for guaranteeing stream alignment, so we are at the mercy of preceding steps.
    
    'Start by writing the bank block header.  This is written as a normal PSP block, with a prefixed
    ' header ID, block ID, and block length.  We obviously cannot write the block length until the block
    ' itself has been written, so we need to flag position in advance.
    cStream.WriteLong PSP_BLOCK_MARKER
    cStream.WriteInt PSP_COMPOSITE_IMAGE_BANK_BLOCK
    
    'Note stream position *immediately after the offset*; we'll return here to write the block length
    ' after it's written.
    cStream.WriteLong 0
    Dim posJustAfterBlockLength As Long
    posJustAfterBlockLength = cStream.GetPosition()
    
    'Now we need to write each composite image as its own chunk.  The PSP format typically likes to
    ' see a thumbnail (300px max dimensions in JPEG format) and a full-size merged image.
    ' (The full-size merged image is useful for non-PSP apps that claim to support PSP format but
    ' don't understand layers.)
    
    'PD will always attempt to write both a thumbnail and a merged image, so we fill in "2" as the
    ' number of composite images.
    cStream.WriteLong 8     '4 + 4
    cStream.WriteLong 2     '2 composite images (one thumbnail, one merged)
    
    'We now write (n) composite image attribute blocks, followed by (n) composite image blocks.
    ' (Note that blocks and chunks are constantly intermixed in the PSP format, in endlessly
    ' obnoxious and nonsensical ways.)
    
    '1) Thumbnail attributes
    
    'Write the thumbnail attributes block first.
    cStream.WriteLong PSP_BLOCK_MARKER
    cStream.WriteIntU PSP_COMPOSITE_ATTRIBUTES_BLOCK
    cStream.WriteLong 24     'Length is fixed and known up-front
    
    'Next is image attributes information chunk, which lives inside the image attributes block
    
    'Chunk size is fixed
    cStream.WriteLong 24
    
    'Thumbnail size is not specified; PSP itself appears to use 300px (aspect-ratio preserved).
    ' PD uses 256 because powers-of-two are more satisfying.
    Dim imgThumbnailDIB As pdDIB
    srcImage.RequestThumbnail imgThumbnailDIB, 256, False
    imgThumbnailDIB.ConvertTo24bpp vbWhite
    
    cStream.WriteLong imgThumbnailDIB.GetDIBWidth
    cStream.WriteLong imgThumbnailDIB.GetDIBHeight
    
    'Bit-depth never includes alpha channels (and the thumbnail is usually 24-bpp anyway)
    cStream.WriteIntU 24
    
    'Thumbnail is JPEG-compressed
    cStream.WriteIntU PSP_COMP_JPEG
    
    'Plane and color count are dumb leftovers from old Windows BMP headers
    cStream.WriteIntU 1
    cStream.WriteLong 2 ^ 24
    
    'Composite image type is important for distinguishing between thumbnail and merged copies
    cStream.WriteIntU PSP_IMAGE_THUMBNAIL
    
    '2) Merged image attributes
    cStream.WriteLong PSP_BLOCK_MARKER
    cStream.WriteIntU PSP_COMPOSITE_ATTRIBUTES_BLOCK
    cStream.WriteLong 24     'Length is fixed and known up-front
    
    'Next is image attributes information chunk, which lives inside the image attributes block
    
    'Chunk size is fixed
    cStream.WriteLong 24
    
    'Image size
    cStream.WriteLong srcImage.Width
    cStream.WriteLong srcImage.Height
    
    'Bit-depth never includes alpha channels
    cStream.WriteIntU 24
    
    'Merged image is zlib-compressed
    cStream.WriteIntU PSP_COMP_LZ77
    
    'Plane and color count are dumb leftovers from old Windows BMP headers
    cStream.WriteIntU 1
    cStream.WriteLong 2 ^ 24
    
    'Composite image type is important for distinguishing between thumbnail and merged copies
    cStream.WriteIntU PSP_IMAGE_COMPOSITE
    
    'Both image attribute blocks are now complete.  Time to write the images themselves.
    ' Note that the format of these blocks is quite different, since the JPEG-format thumbnail has
    ' different requirements and encoding behavior than a normal zlib-compressed image.
    
    '1) Thumbnail image (JPEG format)
    
    'Start with a block header
    cStream.WriteLong PSP_BLOCK_MARKER
    cStream.WriteIntU PSP_JPEG_BLOCK
    cStream.WriteLong 0     'Block length is unknown
    
    'Note the end of the header
    Dim posSubImageBlockHeader As Long
    posSubImageBlockHeader = cStream.GetPosition()
    
    'JPEG information chunk is a fixed size
    cStream.WriteLong 14
    
    'We now need to generate the actual JPEG
    Dim jpegBytes() As Byte
    
    Dim tmpSurface As pd2DSurface
    Set tmpSurface = New pd2DSurface
    tmpSurface.WrapSurfaceAroundDC imgThumbnailDIB.GetDIBDC, 0, imgThumbnailDIB.GetDIBWidth, imgThumbnailDIB.GetDIBHeight
    tmpSurface.SaveSurfaceToArray jpegBytes, P2_FFE_JPEG, 85
    
    'With JPEG bytes now available, we can write the compressed size (followed by the uncompressed size)
    cStream.WriteLong UBound(jpegBytes) + 1
    cStream.WriteLong imgThumbnailDIB.GetDIBStride * imgThumbnailDIB.GetDIBHeight
    
    'Image "type" is a unique PSP enum
    cStream.WriteIntU PSP_DIB_THUMBNAIL
    
    'The JPEG information chunk is complete.  Write the JPEG stream next.
    cStream.WriteByteArray jpegBytes, UBound(jpegBytes) + 1
    Erase jpegBytes
    
    'The thumbnail image is complete.  Return to the start of the block and fill in the final size.
    Dim posEndOfBlock As Long, finalSizeOfBlock As Long
    posEndOfBlock = cStream.GetPosition()
    finalSizeOfBlock = posEndOfBlock - posSubImageBlockHeader
    
    cStream.SetPosition posSubImageBlockHeader - 4
    cStream.WriteLong finalSizeOfBlock
    cStream.SetPosition posEndOfBlock
    
    'We now need to repeat the above steps, but for a much-larger "merged" image copy
    
    '1) Merged image (zlib format)
    
    'Start with a block header
    cStream.WriteLong PSP_BLOCK_MARKER
    cStream.WriteIntU PSP_COMPOSITE_IMAGE_BLOCK
    cStream.WriteLong 0     'Block length is unknown
    
    'Note the end of the header
    posSubImageBlockHeader = cStream.GetPosition()
    
    'Composite information chunk is a fixed size
    cStream.WriteLong 8
    cStream.WriteIntU 2     'Number of bitmaps (2 because alpha is treated as a "mask")
    cStream.WriteIntU 4     'Number of channels (RGBA)
    
    'We now need to generate the actual merged image.  Let's reuse the thumbnail DIB object,
    ' since we don't need the thumbnail itself anymore
    Set imgThumbnailDIB = New pdDIB
    srcImage.GetCompositedImage imgThumbnailDIB, False
    
    'We now need to iterate all channels and write them individually to the stream, just like we
    ' do with individual layers.
    Dim i As Long, srcChannelBytes() As Byte
    ReDim srcChannelBytes(0) As Byte
    
    For i = 0 To 3
        okToProceed = ExportStep4b_WriteChannel(cStream, imgThumbnailDIB, i, srcChannelBytes, dstCompressionLevel, True)
        If (okToProceed >= psp_Failure) Then Exit For
    Next i
    
    'Free the merged copy
    Set imgThumbnailDIB = Nothing
    
    'The merged image is complete.  Return to the start of the block and fill in the final size.
    posEndOfBlock = cStream.GetPosition()
    finalSizeOfBlock = posEndOfBlock - posSubImageBlockHeader
    
    cStream.SetPosition posSubImageBlockHeader - 4
    cStream.WriteLong finalSizeOfBlock
    cStream.SetPosition posEndOfBlock
    
    'With all composite and thumbnail images written, we can finally return to the start of the
    ' composite image block and mark the final block length.
    posEndOfBlock = cStream.GetPosition()
    finalSizeOfBlock = posEndOfBlock - posJustAfterBlockLength
    
    'Return to the block header and fill in the final block size
    cStream.SetPosition posJustAfterBlockLength - 4
    cStream.WriteLong finalSizeOfBlock
    
    'Restore original stream alignment
    cStream.SetPosition posEndOfBlock
    
    'Composite bank complete, and the struct is correctly aligned for the next block
    ExportStep3_CompositeImage = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    m_Warnings.AddString "Internal error in export step 3, #" & Err.Number & ": " & Err.Description
    ExportStep3_CompositeImage = psp_Failure
    
End Function

Private Function ExportStep4_LayerBank(ByRef cStream As pdStream, ByRef srcImage As pdImage, Optional ByVal dstCompressionLevel As Long = -1) As PD_PSPResult

    On Error GoTo InternalVBError
    Const funcName As String = "ExportStep4"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'We are 100% reliant on previous steps aligning the stream pointer correctly.  PSP files do not provide
    ' an easy mechanism for guaranteeing stream alignment, so we are at the mercy of preceding steps.
    
    'Start by writing the bank block header.  This is written as a normal PSP block, with a prefixed
    ' header ID, block ID, and block length.  We obviously cannot write the block length until the block
    ' itself has been written, so we need to flag position in advance.
    cStream.WriteLong PSP_BLOCK_MARKER
    cStream.WriteInt PSP_LAYER_START_BLOCK
    
    'Note stream position *immediately after the offset*; we'll return here to write the block length
    ' after it's written.
    cStream.WriteLong 0
    Dim posJustAfterBlockLength As Long
    posJustAfterBlockLength = cStream.GetPosition()
    
    'Now we need to write each layer as its own layer block.  These are complex multi-part beasts,
    ' so we'll defer to a separate function.
    Dim i As Long
    For i = 0 To srcImage.GetNumOfLayers - 1
        okToProceed = ExportStep4a_WriteLayer(cStream, srcImage, i, dstCompressionLevel)
        If (okToProceed >= psp_Failure) Then Exit For
    Next i
    
    'With all layers written, we can finally return to the start of the layer block and mark the
    ' final block length.
    Dim posEndOfBlock As Long, finalSizeOfBlock As Long
    posEndOfBlock = cStream.GetPosition()
    finalSizeOfBlock = posEndOfBlock - posJustAfterBlockLength
    
    'Return to the block header and write the current size *twice* (once as the block size, then again
    ' as the prefixed chunk size)
    cStream.SetPosition posJustAfterBlockLength - 4
    cStream.WriteLong finalSizeOfBlock
    
    'Restore original stream alignment
    cStream.SetPosition posEndOfBlock
    
    'Layer bank complete, and the struct is correctly aligned for the next block
    ExportStep4_LayerBank = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    m_Warnings.AddString "Internal error in export step 4, #" & Err.Number & ": " & Err.Description
    ExportStep4_LayerBank = psp_Failure
    
End Function

'Write a single layer to file.  Should only be called from its parent ExportStep4 function,
' because this function relies on stream alignment from its parent to work correctly.
Private Function ExportStep4a_WriteLayer(ByRef cStream As pdStream, ByRef srcImage As pdImage, ByVal srcLayerIndex As Long, Optional ByVal dstCompressionLevel As Long = -1) As PD_PSPResult

    On Error GoTo InternalVBError
    Const funcName As String = "ExportStep4a_WriteLayer"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'We are 100% reliant on previous steps aligning the stream pointer correctly.  PSP files do not provide
    ' an easy mechanism for guaranteeing stream alignment, so we are at the mercy of preceding steps.
    
    'Start by writing the layer's sub-block header.  This is written as a normal PSP block, with a prefixed
    ' header ID, block ID, and block length.  We obviously cannot write the block length until the block
    ' itself has been written, so we need to flag position in advance.
    cStream.WriteLong PSP_BLOCK_MARKER
    cStream.WriteInt PSP_LAYER_BLOCK
    
    'Note stream position *immediately after the offset*; we'll return here to write the block length
    ' after it's written.
    cStream.WriteLong 0
    Dim posJustAfterBlockLength As Long
    posJustAfterBlockLength = cStream.GetPosition()
    
    'Next comes the layer information chunk.  PSP chunks start with a chunk size that is not known
    ' in advance (because the header contains variable-length fields, like layer name).  We'll return
    ' and write a valid length after the chunk is complete.
    cStream.WriteLong 0
    
    'Create a generic target layer.  If the current layer uses non-destructive transforms, we'll create
    ' a destructively-transformed copy.
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.CopyExistingLayer srcImage.GetLayerByIndex(srcLayerIndex)
    
    If tmpLayer.AffineTransformsActive(True) Then
        tmpLayer.ConvertToNullPaddedLayer srcImage.Width, srcImage.Height
        tmpLayer.CropNullPaddedLayer
    End If
    
    'Layer name.  Note that the PSP spec does not specify string format.  We write UTF-8 bytes and
    ' just hope for the best *shrug*
    cStream.WriteInt Len(tmpLayer.GetLayerName())
    cStream.WriteString_UTF8 tmpLayer.GetLayerName()
    
    'Layer type (groups and vector are TODO)
    Dim layerType As PSPLayerType
    layerType = keGLTRaster
    cStream.WriteByte layerType
    
    'Confusingly-named image rectangle.  To save space (presumably), PSP files embed both a layer
    ' boundary rect, and a separate - different! - "saved image rect" that defines the rect where
    ' pixel data is actually stored.  Hypothetically this allow you to not save fully transparent
    ' boundary pixels (which is silly because compression handles such runs brilliantly)... but my
    ' point is, this rect basically describes layer offset and dimensions.  We don't waste energy
    ' attempting to trim layer boundaries; instead, we always write these as identical rects.
    cStream.WriteLong tmpLayer.GetLayerOffsetX()
    cStream.WriteLong tmpLayer.GetLayerOffsetY()
    cStream.WriteLong tmpLayer.GetLayerOffsetX() + tmpLayer.GetLayerWidth()
    cStream.WriteLong tmpLayer.GetLayerOffsetY() + tmpLayer.GetLayerHeight()
    
    'Repeat for "significant" image rect, but note that this rect uses *layer* coordinate space
    cStream.WriteLong 0
    cStream.WriteLong 0
    cStream.WriteLong tmpLayer.GetLayerWidth()
    cStream.WriteLong tmpLayer.GetLayerHeight()
    
    'Opacity and blend mode (must use the PSP blend mode enum, obviously)
    cStream.WriteByte Int(tmpLayer.GetLayerOpacity() * 2.55 + 0.5)
    cStream.WriteByte GetPSPBlendModeFromPD(tmpLayer.GetLayerBlendMode())
    
    'Layer flags (includes visibility)
    Dim lyrFlags As PSPLayerProperties
    If tmpLayer.GetLayerVisibility Then lyrFlags = keVisibleFlag
    cStream.WriteByte lyrFlags
    
    'Protected transparency (roughly correlates to locked alpha in PD)
    If (tmpLayer.GetLayerAlphaMode = AM_Locked) Then cStream.WriteByte 1 Else cStream.WriteByte 0
    
    'Group ID (TODO?)
    cStream.WriteByte 0
    
    'Mask rect (unused) and "significant" mask rect
    cStream.WritePadding 32     '4 * 4 * 2
    
    'Mask linked, disabled, invert on blend
    cStream.WritePadding 3      '1 + 1 + 1
    
    'Blend range count and values (unused, but written in a PSP-compatible way for default blending)
    cStream.WriteIntU 0
    Dim i As Long
    For i = 0 To 9
        cStream.WriteLong -65536    '0xFFFF0000
    Next i
    
    'PSPv8 files also support a "use highlight color" and "highlight color" flag;
    ' PD doesn't support these, so write null bytes
    If (m_ExportVersion >= 8) Then
        cStream.WriteByte 0
        cStream.WriteLong 0
    End If
    
    'Layer information chunk is complete!  Note the chunk size, then return to the start of the
    ' stream and fill in the preceding "chunk size" parameter.
    Dim posEndOfChunk As Long
    posEndOfChunk = cStream.GetPosition()
    cStream.SetPosition posJustAfterBlockLength
    cStream.WriteLong posEndOfChunk - posJustAfterBlockLength
    cStream.SetPosition posEndOfChunk
    
    'With the header chunk complete, we could now write a layer extension sub-block for
    ' non-raster layers.  (PD does not currently do this.)
    
    'Instead, we proceed with writing the "layer bitmap chunk"
    Dim posStartOfChunk As Long
    posStartOfChunk = cStream.GetPosition()
    
    'Chunk length (fixed at 8 - chunk size field is included in the calculation!)
    cStream.WriteLong 8
    
    'Bitmap count (always 2 in PD - one for RGB image, another for transparency mask)
    cStream.WriteIntU 2
    
    'Number of channels (always 4 in PD)
    cStream.WriteIntU 4
    
    'Chunk complete.  Next, we need to write one channel block for each channel in this layer
    ' (typically 4 channels: RGBA).  To improve performance, we reuse a shared cache for each channel.
    Dim srcChannelBytes() As Byte
    
    'Un-premultiply the source image before writing it to file
    tmpLayer.GetLayerDIB.SetAlphaPremultiplication False
    
    'Write each channel in turn
    For i = 0 To 3
        okToProceed = ExportStep4b_WriteChannel(cStream, tmpLayer.GetLayerDIB, i, srcChannelBytes, dstCompressionLevel)
        If (okToProceed >= psp_Failure) Then Exit For
    Next i
    
    'With all channels written, we can finally return to the start of this layer block and mark
    ' its final size.
    Dim posEndOfBlock As Long
    posEndOfBlock = cStream.GetPosition()
    cStream.SetPosition posJustAfterBlockLength - 4
    cStream.WriteLong posEndOfBlock - posJustAfterBlockLength
    cStream.SetPosition posEndOfBlock
    
    'This layer is complete, and the struct is correctly aligned for the next block
    ExportStep4a_WriteLayer = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    m_Warnings.AddString "Internal error in export step 4a, #" & Err.Number & ": " & Err.Description
    ExportStep4a_WriteLayer = psp_Failure
    
End Function

'Write a single channel from a source DIB to file.  Should only be called from its parent ExportStep4 function,
' (or in a special case, ExportStep3 for writing a merged image copy) because this function relies on stream
' alignment from its parent to work correctly.
Private Function ExportStep4b_WriteChannel(ByRef cStream As pdStream, ByRef srcDIB As pdDIB, ByVal idxChannel As Long, ByRef srcChannelBytes() As Byte, Optional ByVal dstCompressionLevel As Long = -1, Optional ByVal isCompositeLayer As Boolean = False) As PD_PSPResult

    On Error GoTo InternalVBError
    Const funcName As String = "ExportStep4B_WriteChannel"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'We are 100% reliant on previous steps aligning the stream pointer correctly.  PSP files do not provide
    ' an easy mechanism for guaranteeing stream alignment, so we are at the mercy of preceding steps.
    
    'Start by writing the channel block header.  This is written as a normal PSP block, with a prefixed
    ' header ID, block ID, and block length.  We obviously cannot write the block length until the block
    ' itself has been written, so we need to flag position in advance.
    cStream.WriteLong PSP_BLOCK_MARKER
    cStream.WriteInt PSP_CHANNEL_BLOCK
    
    'Note stream position *immediately after the offset*; we'll return here to write the block length
    ' after it's written.
    cStream.WriteLong 0
    Dim posJustAfterBlockLength As Long
    posJustAfterBlockLength = cStream.GetPosition()
    
    'Next comes the channel information chunk.  PSP chunks typically start with a chunk size that is
    ' not known in advance (because headers often contain variable-length fields, including future
    ' expansion fields), but this chunk is an exception.  It has a fixed size for the PSP version PD
    ' is attempting to support (PSPv8).
    cStream.WriteLong 16
    
    'Next is the compressed and uncompressed size of this channel.  We need to produce an uncompressed
    ' planar stream first (PD always uses interleaved RGBA data internally), then compress it (obviously)
    ' to obtain these values.
    If (Not VBHacks.IsArrayInitialized(srcChannelBytes)) Then ReDim srcChannelBytes(0) As Byte
    
    'Only 32-bpp RGBA data is handled at present
    Dim uncompressedSize As Long
    uncompressedSize = srcDIB.GetDIBWidth * srcDIB.GetDIBHeight
    If (UBound(srcChannelBytes) < uncompressedSize - 1) Then ReDim srcChannelBytes(0 To uncompressedSize - 1) As Byte
    
    'Compressed size is obviously not known in advance.  Skip this field - we'll return and fill it
    ' in later, after compression is complete.
    cStream.WriteLong 0
    cStream.WriteLong uncompressedSize
    
    'The destination buffer must be sized to a "worst-case" scenario for the current uncompressed
    ' size value.  (Worst-case size may vary by compression level.)
    If (dstCompressionLevel < 0) Then dstCompressionLevel = Compression.GetMaxCompressionLevel(cf_Zlib)
    
    Dim compressedSize As Long
    compressedSize = Compression.GetWorstCaseSize(uncompressedSize, cf_Zlib, dstCompressionLevel)
    
    'We now have empty source and destination buffers for this channel.  Fill the source buffer with
    ' the appropriate color plane.
    Dim pxOffset As Long, cID As PSPChannelType, bmpType As PSPDIBType
    If (idxChannel = 0) Then
        pxOffset = 2
        cID = PSP_CHANNEL_RED
        If isCompositeLayer Then bmpType = PSP_DIB_COMPOSITE Else bmpType = PSP_DIB_IMAGE
    ElseIf (idxChannel = 1) Then
        pxOffset = 1
        cID = PSP_CHANNEL_GREEN
        If isCompositeLayer Then bmpType = PSP_DIB_COMPOSITE Else bmpType = PSP_DIB_IMAGE
    ElseIf (idxChannel = 2) Then
        pxOffset = 0
        cID = PSP_CHANNEL_BLUE
        If isCompositeLayer Then bmpType = PSP_DIB_COMPOSITE Else bmpType = PSP_DIB_IMAGE
    ElseIf (idxChannel = 3) Then
        pxOffset = 3
        cID = PSP_CHANNEL_COMPOSITE
        If isCompositeLayer Then bmpType = PSP_DIB_COMPOSITE_TRANS_MASK Else bmpType = PSP_DIB_TRANS_MASK
    End If
    
    'Before compressing the data, write out the rest of the channel information chunk.  This will
    ' allow us to compress the chnnel stream directly into the saved file, saving the need for a
    ' (potentially large) extra destination buffer.
    
    'Write bitmap type and channel type (precalculated above)
    cStream.WriteIntU bmpType
    cStream.WriteIntU cID
    
    'The destination stream now points at the start of the channel content chunk.
    
    'Create a planar representation of the requested color channel.
    Dim pxBytes() As Byte, pxSA As SafeArray1D
    srcDIB.WrapArrayAroundDIB_1D pxBytes, pxSA
    
    Dim i As Long
    For i = 0 To uncompressedSize - 1
        srcChannelBytes(i) = pxBytes(i * 4 + pxOffset)
    Next i
    
    srcDIB.UnwrapArrayFromDIB pxBytes
    
    'We now have a planar representation of the requested color channel.  Compress it directly into
    ' the target stream.  (But first, ensure the stream has enough room to store any worst-case
    ' compression results.)
    cStream.EnsureBufferSpaceAvailable compressedSize
    If Compression.CompressPtrToPtr(cStream.Peek_PointerOnly(peekLength:=compressedSize), compressedSize, VarPtr(srcChannelBytes(0)), uncompressedSize, cf_Zlib, dstCompressionLevel) Then
        
        'Note that the actual compressed size is now stored in "compressedSize"
        okToProceed = psp_Success
        cStream.SetSizeExternally cStream.GetPosition() + compressedSize
        cStream.SetPosition cStream.GetPosition() + compressedSize
        
    'If compression fails, we are screwed because PSPs require a uniform compression method for
    ' *all* channels in a given image (e.g. we can't randomly write an uncompressed channel without
    ' making *all* channels uncompressed to match).
    Else
        InternalError funcName, "Compression failed!  Epic failure, abandoning save."
        okToProceed = psp_Failure
    End If
    
    'Only proceed if compression succeeded.
    If (okToProceed < psp_Failure) Then
        
        'This block is complete.  Note the end-of-block position.
        Dim posEndOfBlock As Long
        posEndOfBlock = cStream.GetPosition()
        
        'Return to the initial block header and fill in the final block size.
        cStream.SetPosition posJustAfterBlockLength - 4
        cStream.WriteLong posEndOfBlock - posJustAfterBlockLength
        
        'The stream now points at the start of the Channel Information Chunk.  Skip 4 bytes
        ' (these are the chunk header), then fill in the final compressed channel size.
        cStream.SetPosition 4, FILE_CURRENT
        cStream.WriteLong compressedSize
        
        'Channel chunk is complete!  Realign the stream pointer, then exit
        cStream.SetPosition posEndOfBlock
        
    End If
    
    'This layer is complete, and the struct is correctly aligned for the next block
    ExportStep4b_WriteChannel = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    m_Warnings.AddString "Internal error in export step 4b, #" & Err.Number & ": " & Err.Description
    ExportStep4b_WriteChannel = psp_Failure
    
End Function

'Only valid after the image header has been parsed (step 2)
Private Function GetDPI() As Double
    Dim useEmbeddedResolution As Boolean
    useEmbeddedResolution = (m_ImageHeader.psph_Resolution > 0#)
    useEmbeddedResolution = useEmbeddedResolution And ((m_ImageHeader.psph_ResolutionUnit = PSP_METRIC_INCH) Or (m_ImageHeader.psph_ResolutionUnit = PSP_METRIC_CM))
    If useEmbeddedResolution Then
        If (m_ImageHeader.psph_ResolutionUnit = PSP_METRIC_INCH) Then
            GetDPI = m_ImageHeader.psph_Resolution
        Else
            GetDPI = Units.GetCMFromInches(m_ImageHeader.psph_Resolution)
        End If
    Else
        GetDPI = 96#
    End If
End Function

Private Function GetPSPBlendModeFromPD(ByVal srcBlendMode As PD_BlendMode) As PSPBlendModes

    Select Case srcBlendMode
        Case BM_Normal
            GetPSPBlendModeFromPD = bmLAYER_BLEND_NORMAL
        Case BM_Darken
            GetPSPBlendModeFromPD = bmLAYER_BLEND_DARKEN
        Case BM_Lighten
            GetPSPBlendModeFromPD = bmLAYER_BLEND_LIGHTEN
        Case BM_Hue
            GetPSPBlendModeFromPD = bmLAYER_BLEND_LEGACY_HUE    'Using the legacy variants ensures backward-compatibility for more PSP versions
        Case BM_Saturation
            GetPSPBlendModeFromPD = bmLAYER_BLEND_LEGACY_SATURATION
        Case BM_Color
            GetPSPBlendModeFromPD = bmLAYER_BLEND_LEGACY_COLOR
        Case BM_Luminosity
            GetPSPBlendModeFromPD = bmLAYER_BLEND_LEGACY_LUMINOSITY
        Case BM_Multiply
            GetPSPBlendModeFromPD = bmLAYER_BLEND_MULTIPLY
        Case BM_Screen
            GetPSPBlendModeFromPD = bmLAYER_BLEND_SCREEN
        Case BM_Overlay
            GetPSPBlendModeFromPD = bmLAYER_BLEND_OVERLAY
        Case BM_HardLight
            GetPSPBlendModeFromPD = bmLAYER_BLEND_HARD_LIGHT
        Case BM_SoftLight
            GetPSPBlendModeFromPD = bmLAYER_BLEND_SOFT_LIGHT
        Case BM_Difference
            GetPSPBlendModeFromPD = bmLAYER_BLEND_DIFFERENCE
        Case BM_ColorDodge
            GetPSPBlendModeFromPD = bmLAYER_BLEND_DODGE
        Case BM_ColorBurn
            GetPSPBlendModeFromPD = bmLAYER_BLEND_BURN
        Case BM_Exclusion
            GetPSPBlendModeFromPD = bmLAYER_BLEND_EXCLUSION
        Case Else
            InternalError "GetPSPBlendModeFromPD", "PSP doesn't support blendmode: " & srcBlendMode
            GetPSPBlendModeFromPD = bmLAYER_BLEND_NORMAL
    End Select
    
End Function

Private Function Step1_ValidateHeader(ByRef srcFile As String) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "Step1_ValidateHeader"
    
    'If the passed path is zero, assume the caller is loading the PSP from memory.
    If (LenB(srcFile) = 0) Then m_SourceFilename = PSP_LOADED_FROM_MEMORY Else m_SourceFilename = srcFile
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'Open a stream over the source data.
    Set m_Stream = New pdStream
    
    If (m_SourceFilename = PSP_LOADED_FROM_MEMORY) Then
        If Not m_Stream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_SourcePtrLen, m_SourcePtr) Then
            InternalError funcName, "can't start in-memory stream on pointer: " & m_SourcePtr
            okToProceed = psp_Failure
        End If
    Else
        If Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, m_SourceFilename, , , OptimizeSequentialAccess) Then
            InternalError funcName, "can't access file; it may be locked or in an inaccessible location."
            okToProceed = psp_Failure
        End If
    End If
    
    'The stream is open.  Validate both the PSP's ASCII identifier and its version.
    ' (We will attempt to parse on a version failure, but I have not tested old PSP file versions
    ' as aggressively so no promises are made on correctness!)
    Dim pspValidated As Boolean
    If (okToProceed < psd_Failure) Then
        
        'PSP files use a lovely 32-byte identifier that completely eliminates the possibility
        ' of false-positives.  Yay!
        Dim pspID() As Byte, checkBytes() As Byte
        
        '(Note: this hex string translates to “Paint Shop Pro Image File\n\x1a”, padded with zeroes
        ' to 32 bytes).
        If Strings.BytesFromHex(pspID, "5061696E742053686F702050726F20496D6167652046696C650A1A0000000000") Then
            If m_Stream.AreBytesAvailable(32) Then
                If (m_Stream.ReadBytes(checkBytes, 32, True) = 32) Then
                    pspValidated = VBHacks.MemCmp(VarPtr(pspID(0)), VarPtr(checkBytes(0)), 32)
                    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "File """ & srcFile & """ initial PSP validation: " & pspValidated
                Else
                    InternalError funcName, "32 bytes weren't returned?"
                End If
            Else
                If PSP_DEBUG_VERBOSE Then InternalError funcName, "file is <32 bytes - not a PSP file!"
            End If
        Else
            InternalError funcName, "BytesFromHex failed?"
        End If
        
        If pspValidated Then
            okToProceed = psp_Success
        Else
            okToProceed = psp_FileNotPSP
        End If
        
        'Validate version number only if the original checksum passed
        If (okToProceed < psp_Failure) Then
            
            PDDebug.LogAction "File is Paintshop Pro (PSP) format."
            
            m_ImageHeader.psph_VersionMajor = m_Stream.ReadInt()
            m_ImageHeader.psph_VersionMinor = m_Stream.ReadInt()
            m_PSPisV5 = (m_ImageHeader.psph_VersionMajor <= 3)
            
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "PSP file found: version is " & m_ImageHeader.psph_VersionMajor & "." & m_ImageHeader.psph_VersionMinor
            
            'Per v8 of the spec, PSP versions are as follows:
            'Current version of the spec (v8) is 6.0
            ' 3.0 = Paint Shop Pro 5
            ' 4.0 = Paint Shop Pro 6
            ' 5.0 = Paint Shop Pro 7
            ' 6.0 = Paint Shop Pro 8
            ' 7.0 = Paint Shop Pro 9
            
            'Higher major version numbers are expected (e.g. I've encountered "13" in images
            ' created in 2020 found "in the wild"), but I don't have a spec for those versions
            ' so correct loading is not necessarily guaranteed.  PD won't raise errors on such
            ' files, but I note that here so that we can add custom handling if something breaks
            ' in the future
            If (m_ImageHeader.psph_VersionMajor < 3) Then m_Warnings.AddString "Unexpected PSP major version encountered: " & m_ImageHeader.psph_VersionMajor
            If (m_ImageHeader.psph_VersionMinor <> 0) Then m_Warnings.AddString "Unexpected PSP minor version encountered: " & m_ImageHeader.psph_VersionMinor
            
        Else
            If PSP_DEBUG_VERBOSE Then m_Warnings.AddString "PSP header failed basic validation.  (This is not a PSP file.)"
            okToProceed = psp_FileNotPSP
        End If
        
    End If
    
    'Validation complete.  If the file validated successfully, this function guarantees that m_Stream
    ' points at the first byte PAST the valid PSP file header.
    Step1_ValidateHeader = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 1, #" & Err.Number & ": " & Err.Description
    Step1_ValidateHeader = psp_Failure

End Function

'Do *NOT* call import steps out-of-order!
Private Function Step2_BuildBlockDirectory(ByRef srcFile As String) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "Step2_BuildBlockDirectory"
    
    Step2_BuildBlockDirectory = psp_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError funcName, "filename has changed since original validation!"
        Step2_BuildBlockDirectory = psp_Failure
        Exit Function
    End If
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'PSP files consist of a series of "blocks".  Total block count is unknown; the only way
    ' to know how many blocks are in the file is to "walk" the file manually.
    
    'In this step, we build an initial "block" collection.  This lets us perform some
    ' additional validation (like ensuring that critical blocks are in valid order).
    
    'Start by resetting our current block collection, if any
    Const DEFAULT_BLOCK_COUNT As Long = 8
    m_numOfBlocks = 0
    ReDim m_Blocks(0 To DEFAULT_BLOCK_COUNT - 1) As pdPSPBlock
    
    Dim curBlockID As PSPBlockID, curBlockLength As Long
    
    'After validation, we will be pointing at the first block in our collection.
    
    'All blocks use a uniform 32-bit "flag" ID, followed by a block ID and length
    Do While (m_Stream.ReadLong() = PSP_BLOCK_MARKER)
        
        'Ensure there's room in our block collection
        If (m_numOfBlocks > UBound(m_Blocks)) Then ReDim Preserve m_Blocks(0 To m_numOfBlocks * 2 - 1) As pdPSPBlock
        
        'Retrieve 2-byte ID
        curBlockID = m_Stream.ReadIntUnsigned()
        curBlockLength = m_Stream.ReadLong()
        
        'IMPORTANTLY: in very old PSP versions, chunks use two length descriptors.  The first describes
        ' the first sub-chunk, while the second describes the full chunk (dumb as hell, I know).  If we
        ' don't catch this at import time, it'll upset everything that follows.
        If m_PSPisV5 Then curBlockLength = m_Stream.ReadLong()
        
        'Add this block to the collection
        Set m_Blocks(m_numOfBlocks) = New pdPSPBlock
        m_Blocks(m_numOfBlocks).InitializeBlock m_Stream.GetPosition(), curBlockID, curBlockLength
        m_numOfBlocks = m_numOfBlocks + 1
        
        If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Block #" & m_numOfBlocks & ": " & m_Blocks(m_numOfBlocks - 1).GetBlockIDAsString()
        
        'Move the pointer ahead by the specified block length; we should be pointing at either a
        ' new block, or the end of the file!
        m_Stream.SetPosition curBlockLength, FILE_CURRENT
        If (m_Stream.GetPosition() >= m_Stream.GetStreamSize()) Then Exit Do
        
    Loop
    
    'With all blocks flagged, we can now do a bit of minimal validation on blocks and block order.
    ' (The PSP spec requires certain blocks appear in certain places and/or orders.)
    
    'From the spec:
    ' With regard to the order of the ten main blocks, there are two hard and fast rules.
    '  - First, the General Image Attributes block must come immediately after the File Header.
    '  - Second, the Composite Image Bank Block, if present, must come before the Layer Bank Block.
    '  - The order of other blocks is not mandated.
    
    'Start by ensuring the first block is the General Image Attributes block
    If (m_numOfBlocks > 0) Then
        If (Not m_Blocks(0) Is Nothing) Then
            If (m_Blocks(0).GetBlockID <> PSP_IMAGE_BLOCK) Then
                okToProceed = psp_Failure
                InternalError funcName, "first block is not General Image Attributes"
            End If
        Else
            okToProceed = psp_Failure
            InternalError funcName, "first block is corrupt"
        End If
    Else
        okToProceed = psp_Failure
        InternalError funcName, "no valid blocks found"
    End If
    
    'If the first block passed, look for the composite image bank block and the layer bank block.
    ' We need to ensure the former appears before the latter.
    Dim posCompositeBank As Long, posLayerBank As Long
    If (okToProceed < psp_Failure) Then
        
        'Set the composite and layer bank indices to invalid values
        posLayerBank = -2
        posCompositeBank = -1
        
        Dim i As Long
        For i = 0 To m_numOfBlocks - 1
            If (Not m_Blocks(i) Is Nothing) Then
                If (m_Blocks(i).GetBlockID = PSP_COMPOSITE_IMAGE_BANK_BLOCK) Then posCompositeBank = i
                If (m_Blocks(i).GetBlockID = PSP_LAYER_START_BLOCK) Then posLayerBank = i
            End If
        Next i
        
        If (posLayerBank <= posCompositeBank) And (posCompositeBank >= 0) Then
            okToProceed = psp_Failure
            InternalError funcName, "critical blocks out of order"
        Else
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "PSP block directory valid."
        End If
        
    End If
    
    'Finally, PSP images could theoretically be "empty" (e.g. containing just a header, but no
    ' layer or pixel data), but we don't want to mess with this possibility in PD.  Throw out
    ' such files.
    If (okToProceed < psp_Failure) Then
        If (m_numOfBlocks < 2) Then
            okToProceed = psp_Failure
            InternalError funcName, "too few blocks"
        End If
    End If
    
    'Block directory validation and assembly is complete!
    Step2_BuildBlockDirectory = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 2, #" & Err.Number & ": " & Err.Description
    Step2_BuildBlockDirectory = psp_Failure

End Function

'Do *NOT* call import steps out-of-order!
Private Function Step3_ReadImageAttributes(ByRef srcFile As String) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "Step3_ReadImageAttributes"
    
    Step3_ReadImageAttributes = psp_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError funcName, "filename has changed since original validation!"
        Step3_ReadImageAttributes = psp_Failure
        Exit Function
    End If
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'With a block directory now constructed, we can proceed with parsing individual blocks.
    
    'First, we need to parse the "General Image Attributes Block", which is the first block
    ' in all valid PSP files.
    Dim startOfHeader As Long
    startOfHeader = m_Blocks(0).GetBlockOffset()
    m_Stream.SetPosition startOfHeader, FILE_BEGIN
    
    'Start by simply pulling the header into a VB-friendly struct.  (We'll validate it after.)
    With m_ImageHeader
        
        'Early PSP versions use an asinine mechanism for struct length descriptors, so this is
        ' a stupid (but necessary) workaround
        If m_PSPisV5 Then
            m_Stream.SetPosition -8, FILE_CURRENT
            .psph_HeaderSize = m_Stream.ReadLong()
            m_Stream.SetPosition 4, FILE_CURRENT
        Else
            .psph_HeaderSize = m_Stream.ReadLong()
        End If
        
        'The header must be at least 36 bytes long (technically 42, but we have workarounds
        ' for corrupt layer data)
        If (.psph_HeaderSize < 36) Then
            InternalError funcName, "bad header"
            Step3_ReadImageAttributes = psp_Failure
            Exit Function
        End If
        
        .psph_Width = m_Stream.ReadLong()
        .psph_Height = m_Stream.ReadLong()
        .psph_Resolution = m_Stream.ReadDouble()
        .psph_ResolutionUnit = m_Stream.ReadByte()
        .psph_Compression = m_Stream.ReadIntUnsigned()
        .psph_BitDepth = m_Stream.ReadIntUnsigned()
        .psph_PlaneCount = m_Stream.ReadIntUnsigned()
        .psph_ColorCount = m_Stream.ReadLong()
        .psph_IsGrayscale = (m_Stream.ReadByte() = 1)
        .psph_TotalSize = m_Stream.ReadLong()
        PDDebug.LogAction "TOTAL SIZE: " & .psph_TotalSize
        
        'Technically these are just paranoid failsafe checks; layer data should *always* exist,
        ' but it doesn't hurt to account for broken 3rd-party software
        If (.psph_HeaderSize >= 42) Then
            .psph_ActiveLayer = m_Stream.ReadLong()
            .psph_LayerCount = m_Stream.ReadIntUnsigned()
        Else
            .psph_ActiveLayer = 0
            .psph_LayerCount = 1
        End If
        
        'Content flags may not exist in very old PSP versions
        If (.psph_HeaderSize >= 46) Then
            .psph_ContentFlags = m_Stream.ReadLong()
        Else
            .psph_ContentFlags = 0
        End If
        
        'Future expansion fields are explicitly supported by the spec, but I haven't found
        ' a spec past v8 so I don't currently know how to interpret any extra bytes here.
        
        'With the header successfully imported, validate as many relevant members as we can.
        If (.psph_Width <= 0) Or (.psph_Height <= 0) Then
            InternalError funcName, "bad width/height"
            okToProceed = psp_Failure
        End If
        
        If (.psph_Width <= 0) Or (.psph_Height <= 0) Then
            InternalError funcName, "bad width/height"
            okToProceed = psp_Failure
        End If
        
        If (.psph_Compression >= PSP_COMP_JPEG) Then
            InternalError funcName, "bad compression"
            okToProceed = psp_Failure
        End If
        
        If .psph_IsGrayscale Then
            If (.psph_BitDepth <> 1) And (.psph_BitDepth <> 4) And (.psph_BitDepth <> 8) And (.psph_BitDepth <> 16) Then
                InternalError funcName, "bad gray bit-depth"
                okToProceed = psp_Failure
            End If
        Else
            If (.psph_BitDepth <> 1) And (.psph_BitDepth <> 4) And (.psph_BitDepth <> 8) And (.psph_BitDepth <> 24) And (.psph_BitDepth <> 48) Then
                InternalError funcName, "bad color bit-depth"
                okToProceed = psp_Failure
            End If
        End If
        
        If PSP_DEBUG_VERBOSE Then
            PDDebug.LogAction "PSP dimensions: " & .psph_Width & "x" & .psph_Height & ", bit-depth: " & .psph_BitDepth & ", " & .psph_LayerCount & " layer(s), compression: " & GetCompressionName(.psph_Compression) & ", isGrayscale: " & .psph_IsGrayscale
        End If
        
    End With
    
    'Block directory validation and assembly is complete!
    
    'Forcibly align the stream pointer before continuing
    m_Stream.SetPosition startOfHeader + m_Blocks(0).GetBlockLength(), FILE_BEGIN
    Step3_ReadImageAttributes = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 3, #" & Err.Number & ": " & Err.Description
    Step3_ReadImageAttributes = psp_Failure

End Function

'Do *NOT* call import steps out-of-order!
Private Function Step3a_ReadICCProfile() As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "Step3a_ReadICCProfile"
    Step3a_ReadICCProfile = psp_Success
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'Obviously we can only load an ICC profile if one actually exists in the file
    Dim i As Long, iccIndex As Long
    For i = 0 To m_numOfBlocks - 1
        If (m_Blocks(i).GetBlockID = PSP_COLORPROFILE_BLOCK) Then
            iccIndex = i
            Exit For
        End If
    Next i
    
    'If a profile block doesn't exist, exit now.  (Note also that v5 files do *not* support
    ' embedded color profiles, which is actually awesome as it spares us dealing with their
    ' bullshit file layout differences!)
    If (iccIndex <= 0) Then
        If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "No ICC profile found (sRGB assumed)"
        Set m_ICCProfile = Nothing
        Step3a_ReadICCProfile = psp_Success
        Exit Function
    End If
    
    'If a profile block was found, its layout is trivial to parse!
    Dim origBlockPos As Long
    m_Stream.SetPosition m_Blocks(iccIndex).GetBlockOffset, FILE_BEGIN
    origBlockPos = m_Stream.GetPosition()
    
    Dim iccChunkSize As Long
    iccChunkSize = m_Stream.ReadLong()
    
    'Description isn't especially relevant to PD
    Dim iccDescriptionSize As Long, iccDescription As String
    iccDescriptionSize = m_Stream.ReadIntUnsigned()
    iccDescription = m_Stream.ReadString_UTF8(iccDescriptionSize)
    
    Dim iccLength As Long, iccBytes() As Byte
    iccLength = m_Stream.ReadLong()
    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "ICC profile found: " & iccDescription & " (" & iccDescriptionSize & ", " & iccLength & ", " & iccChunkSize & ", " & m_Blocks(iccIndex).GetBlockLength() & ")"
    
    'The spec is incorrect here.  The ICC profile looks as though it is part of the "color profile
    ' data chunk" but it's not.  Instead, it immediately follows the data chunk - so we need to skip
    ' past the data chunk's length before reading it.
    m_Stream.SetPosition origBlockPos + iccChunkSize, FILE_BEGIN
    
    If (iccLength = m_Stream.ReadBytes(iccBytes, iccLength, True)) Then
        
        'Cache the retrieved ICC bytes inside a dedicated profile class
        Set m_ICCProfile = New pdICCProfile
        m_ICCProfile.LoadICCFromPtr iccLength, VarPtr(iccBytes(0))
        If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "ICC profile loaded successfully!"
        okToProceed = psp_Success
    
    'Read failed
    Else
        Set m_ICCProfile = Nothing
        okToProceed = psp_Warning
        InternalError funcName, "bad ICC length descriptor"
    End If
    
    'Before exiting, set the pointer to align with the next block.
    m_Stream.SetPosition origBlockPos + m_Blocks(iccIndex).GetBlockLength()
    
    Step3a_ReadICCProfile = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 3a, #" & Err.Number & ": " & Err.Description
    Step3a_ReadICCProfile = psp_Failure
    
End Function

'Do *NOT* call import steps out-of-order!
Private Function Step4_ReadCompositeBlock(ByRef srcFile As String, Optional ByVal retrieveImages As Boolean = False) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "Step4_ReadCompositeBlock"
    
    Step4_ReadCompositeBlock = psp_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError funcName, "filename has changed since original validation!"
        Step4_ReadCompositeBlock = psp_Failure
        Exit Function
    End If
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'Certain images are not required to have a composite block at all.  Per the spec,
    ' "The Composite Image Bank Block is a block that contains one or more bitmaps of
    '  the merged document. Typically, the Composite Image Bank Block will contain one
    '  miniature bitmap of the merged document (a thumbnail) and one full size bitmap
    '  of the merged document. The full size bitmap must have the same bit depth and
    '  resolution as the entire image. All PSP files with more than one layer or with
    '  one layer with transparency must have a full size composite image."
    
    'As such, let's start by seeing if a composite block even exists for this image.
    Dim compositeBlockIndex As Long
    
    Dim i As Long
    For i = 0 To m_numOfBlocks - 1
        If (m_Blocks(i).GetBlockID = PSP_COMPOSITE_IMAGE_BANK_BLOCK) Then
            compositeBlockIndex = i
            Exit For
        End If
    Next i
    
    'Old PSP files didn't use a full "composite bank" with combination merged and thumbnail files;
    ' instead, they just provided a thumbnail segment.  Hypothetically we could retrieve these
    ' thumbnails some day, but for now, I don't know that it's worth the trouble.  (Maybe if PD
    ' someday provides a "Browse" tool I can revisit...)
    '
    'Anyway, if a composite block doesn't exist, exit now.
    If (compositeBlockIndex <= 0) Then
        
        If m_PSPisV5 Then
            m_Warnings.AddString "This is an ancient PSP file; composite blocks don't exist in this version!"
        Else
            m_Warnings.AddString "No composite block found; this may be an opaque single-layer image?"
        End If
        
        Step4_ReadCompositeBlock = psp_Warning
        Exit Function
        
    End If
    
    'If a composite block was found, query the number of composite images in the file
    Dim origBankInfoPos As Long
    m_Stream.SetPosition m_Blocks(compositeBlockIndex).GetBlockOffset, FILE_BEGIN
    origBankInfoPos = m_Stream.GetPosition()
    
    Dim cmpBankInfoSize As Long
    
    'Early PSP versions use an asinine mechanism for struct length descriptors, so this is
    ' a stupid (but necessary) workaround
    If m_PSPisV5 Then
        m_Stream.SetPosition -8, FILE_CURRENT
        cmpBankInfoSize = m_Stream.ReadLong()
        m_Stream.SetPosition 4, FILE_CURRENT
    Else
        cmpBankInfoSize = m_Stream.ReadLong()
    End If
    
    m_numCompositeImages = m_Stream.ReadLong()
    
    'Perform basic validation on both values
    If (cmpBankInfoSize < 4) Or (m_numCompositeImages <= 0) Then
        InternalError funcName, "bad composite block header: " & cmpBankInfoSize & ", " & m_numCompositeImages
        Step4_ReadCompositeBlock = psp_Warning
        Exit Function
    End If
    
    'The composite image bank provides for future expansion fields, so it's critical
    ' to use the chunk size to move the pointer into position.
    m_Stream.SetPosition origBankInfoPos, FILE_BEGIN
    m_Stream.SetPosition cmpBankInfoSize, FILE_CURRENT
    
    'The stream now points at the Composite Image Attributes sub-block.  This is a list
    ' of attribute structs, one for each composite image.  The attribute struct bears
    ' strong similarity to a normal image header, with basic fields like width, height,
    ' color-depth and compression.  We want to retrieve all structs, then search the list
    ' for the most appropriate image to pull (typically the largest and/or highest-quality).
    ReDim m_CompositeImages(0 To m_numCompositeImages - 1) As PSPCompositeImage
    
    'Note that each image attribute entry is stored inside a dedicated block, with a full
    ' block header.  We will validate each block as we read it to ensure our stream pointer
    ' remains properly aligned.
    Dim startOfBlock As Long, blockLength As Long, blockOK As Boolean
    
    For i = 0 To m_numCompositeImages - 1
        
        blockOK = True
        
        'Validate the universal block ID
        If (m_Stream.ReadLong() <> PSP_BLOCK_MARKER) Then
            InternalError funcName, "bad attribute block header"
            blockOK = False
        End If
        
        'Validate the composite-attributes ID
        If (m_Stream.ReadIntUnsigned() <> PSP_COMPOSITE_ATTRIBUTES_BLOCK) Then
            InternalError funcName, "bad attribute block ID"
            blockOK = False
        End If
        
        'Grab the total block length, then note the current stream pointer
        blockLength = m_Stream.ReadLong()
        If (blockLength <= 0) Then
            InternalError funcName, "bad attribute block length"
            blockOK = False
        End If
        
        startOfBlock = m_Stream.GetPosition()
        
        'If basic block validation fails, stream alignment isn't guaranteed - so we must
        ' abandon composite parsing entirely.
        If (Not blockOK) Then
            InternalError funcName, "Stream pointer misaligned; abandoning composite parsing"
            Step4_ReadCompositeBlock = psp_Warning
            Exit Function
        End If
        
        'If we're still here, the block header appears OK.
        ' Retrieve the composite image attribute information chunk.
        With m_CompositeImages(i)
            .cih_HeaderSize = m_Stream.ReadLong()
            .cih_Width = m_Stream.ReadLong()
            .cih_Height = m_Stream.ReadLong()
            .cih_BitDepth = m_Stream.ReadIntUnsigned()
            .cih_Compression = m_Stream.ReadIntUnsigned()
            .cih_PlaneCount = m_Stream.ReadIntUnsigned()
            .cih_ColorCount = m_Stream.ReadLong()
            .cih_CompositeType = m_Stream.ReadIntUnsigned()
            
            'The spec allows for future expansion, so the block size must be used for alignment
            
            'Start out assuming this frame is OK
            .cih_LooksOK = True
            
            'Next, perform basic validation on this header
            If (.cih_HeaderSize > blockLength) Or (.cih_HeaderSize < 24) Then
                InternalError funcName, "bad composite header size"
                .cih_LooksOK = False
            End If
            
            If (.cih_Width <= 0) Or (.cih_Height <= 0) Then
                InternalError funcName, "bad composite image dimensions"
                .cih_LooksOK = False
            End If
            
            If (.cih_BitDepth <> 1) And (.cih_BitDepth <> 4) And (.cih_BitDepth <> 8) And (.cih_BitDepth <> 16) And (.cih_BitDepth <> 24) And (.cih_BitDepth <> 48) Then
                InternalError funcName, "bad composite image bit-depth"
                .cih_LooksOK = False
            End If
            
            If (.cih_Compression > PSP_COMP_JPEG) Then
                InternalError funcName, "bad composite compression"
                .cih_LooksOK = False
            End If
            
            If (.cih_PlaneCount <> 1) Then InternalError funcName, "bad composite plane count; will attempt to read anyway"
            
            If (.cih_CompositeType <> PSP_IMAGE_COMPOSITE) And (.cih_CompositeType <> PSP_IMAGE_THUMBNAIL) Then
                InternalError funcName, "bad composite type"
                .cih_LooksOK = False
            End If
            
            'Debug output only
            If PSP_DEBUG_VERBOSE Then
                If .cih_LooksOK Then
                    PDDebug.LogAction "Composite #" & (i + 1) & ": " & .cih_Width & "x" & .cih_Height & ", bit-depth: " & .cih_BitDepth & ", type: " & IIf(.cih_CompositeType = PSP_IMAGE_COMPOSITE, "composite", "thumbnail") & ", compression: " & GetCompressionName(.cih_Compression)
                Else
                    PDDebug.LogAction "Composite #" & (i + 1) & ": marked invalid"
                End If
            End If
            
        End With
        
        'Regardless of whether this image looks okay or not, reorient the stream pointer
        ' using the start-of-block position and the block length.
        m_Stream.SetPosition startOfBlock, FILE_BEGIN
        m_Stream.SetPosition blockLength, FILE_CURRENT
        
    Next i
    
    'With all composite frame attributes parsed, our stream should now be pointing at the
    ' first composite image sub-block.  These come in two varieties:
    ' - JPEG blocks (typically used for thumbnails)
    ' - "Normal" blocks (encoded identically to PSP layers)
    
    'Further handling is contingent on the caller requesting us to retrieve composite images.
    ' (By default, we don't.)
    If retrieveImages Then
    
        'OK, let's retrieve each composite image in turn!
        For i = 0 To m_numCompositeImages - 1
            
            blockOK = True
            
            'Validate the universal block ID
            If (m_Stream.ReadLong() <> PSP_BLOCK_MARKER) Then
                InternalError funcName, "bad composite image block header"
                blockOK = False
            End If
            
            'Validate the composite-attributes ID
            Dim compositeImgType As PSPBlockID
            compositeImgType = m_Stream.ReadIntUnsigned()
            If (compositeImgType <> PSP_JPEG_BLOCK) And (compositeImgType <> PSP_COMPOSITE_IMAGE_BLOCK) Then
                InternalError funcName, "bad composite image block ID"
                blockOK = False
            End If
            
            If Not ((compositeImgType = PSP_JPEG_BLOCK) Xor (m_CompositeImages(i).cih_Compression <> PSP_COMP_JPEG)) Then
                InternalError funcName, "bad composite image compression"
                blockOK = False
            End If
            
            'Grab the total block length, then note the current stream pointer
            blockLength = m_Stream.ReadLong()
            If (blockLength <= 0) Then
                InternalError funcName, "bad composite image block length"
                blockOK = False
            End If
            
            startOfBlock = m_Stream.GetPosition()
            
            'Further handling depends on JPEG vs "normal" block encoding
            
            'JPEG frame
            If (compositeImgType = PSP_JPEG_BLOCK) Then
                okToProceed = RetrieveComposite_JPEG(i, blockLength)
                
            '"Normal" frame
            Else
                okToProceed = RetrieveComposite_Normal(i, blockLength)
                
            End If
            
            'After decoding the block, forcibly move the pointer to the end of the
            ' current block, using the original block length as our guide.
            m_Stream.SetPosition startOfBlock + blockLength, FILE_BEGIN
        
        Next i
        
    End If
    
    'Composite block is finished!
    Step4_ReadCompositeBlock = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 4, #" & Err.Number & ": " & Err.Description
    Step4_ReadCompositeBlock = psp_Failure

End Function

'Do *NOT* call import steps out-of-order!
Private Function Step5_ReadParentPalette(ByRef srcFile As String) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "Step5_ReadParentPalette"
    
    Step5_ReadParentPalette = psp_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError funcName, "filename has changed since original validation!"
        Step5_ReadParentPalette = psp_Failure
        Exit Function
    End If
    
    'Note stream position before proceeding
    Dim origStreamPosition As Long
    origStreamPosition = m_Stream.GetPosition()
    
    'Not all PSP files will contain a palette block. Per the spec:
    ' "The Color Palette Block contains a global color palette to be used with the color bitmaps
    '  in all layers of the document. This block is required for image documents that contain
    '  paletted bitmaps and should not be written for image documents that contain 24-bit or
    '  48-bit bitmaps (the block can be ignored in the latter cases)."
    
    'PD will always attempt to retrieve a parent palette if one exists.  24/48-bpp images
    ' will simply ignore it when generating layers.
    Dim paletteBlockIndex As Long
    paletteBlockIndex = -1
    
    Dim i As Long
    For i = 0 To m_numOfBlocks - 1
        If (m_Blocks(i).GetBlockID = PSP_COLOR_BLOCK) Then
            paletteBlockIndex = i
            Exit For
        End If
    Next i
    
    'If no palette block is found, that's fine - just mark it as such and carry on!
    If (paletteBlockIndex < 0) Then
        
        Step5_ReadParentPalette = psp_Success
        
        'If the current image is grayscale (regardless of color-depth), generate a default
        ' gray palette matching the current bit-depth.
        If m_ImageHeader.psph_IsGrayscale Then
        
            If (m_ImageHeader.psph_BitDepth >= 8) Then
                m_ParentPaletteCount = 256
                Palettes.GetPalette_Grayscale m_ParentPalette
            ElseIf (m_ImageHeader.psph_BitDepth = 4) Then
                m_ParentPaletteCount = 16
                Palettes.GetPalette_GrayscaleEx m_ParentPalette, m_ParentPaletteCount
            ElseIf (m_ImageHeader.psph_BitDepth = 1) Then
                m_ParentPaletteCount = 2
                Palettes.GetPalette_GrayscaleEx m_ParentPalette, m_ParentPaletteCount
            Else
                InternalError funcName, "unknown bit-depth and grayscale combination"
                Step5_ReadParentPalette = psp_Failure
            End If
            
            If PSP_DEBUG_VERBOSE Then
                If (Step5_ReadParentPalette < psp_Failure) Then PDDebug.LogAction "Parent grayscale palette constructed successfully."
            End If
        
        'This is a color image
        Else
            
            m_ParentPaletteCount = 0
            
            'The error state of this case depends on parent image bit-depth, obviously.
            If (m_ImageHeader.psph_BitDepth <= 8) Then
                InternalError funcName, "no palette found; image is unsaveable"
                Step5_ReadParentPalette = psp_Failure
                
            'This image is full-color, so just carry on loading!
            Else
                Step5_ReadParentPalette = psp_Success
            End If
            
        End If
        
        'If no palette is found, we can exit immediately
        Exit Function
    
    'A palette block was found
    Else
        
        'Set the pointer to this block's offset, then proceed with reading the palette
        m_Stream.SetPosition m_Blocks(paletteBlockIndex).GetBlockOffset() - 10, FILE_BEGIN
        
        'As always, workarounds are required for v5 stream alignment
        If m_PSPisV5 Then m_Stream.SetPosition -4, FILE_CURRENT
        
        If RetrievePalette(m_ParentPalette, m_ParentPaletteCount) Then
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Parent palette retrieved successfully (" & m_ParentPaletteCount & " colors)"
            Step5_ReadParentPalette = psp_Success
        Else
            InternalError funcName, "bad palette block"
            Step5_ReadParentPalette = psp_Failure
        End If
        
    End If
    
    'Palette retrieval is finished!  Regardless of what happened, reset stream position
    m_Stream.SetPosition origStreamPosition, FILE_BEGIN
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 5, #" & Err.Number & ": " & Err.Description
    Step5_ReadParentPalette = psp_Failure

End Function

'Do *NOT* call import steps out-of-order!
Private Function Step6_LoadLayers(ByRef srcFile As String) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "Step6_LoadLayers"
    
    Step6_LoadLayers = psp_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError funcName, "filename has changed since original validation!"
        Step6_LoadLayers = psp_Failure
        Exit Function
    End If
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'The last step in loading a PSP file is loading the "layer bank".
    
    'Before proceeding, backup the stream position.  (We can restore this if things go
    ' catastrophically wrong.)
    Dim entryStreamPosition As Long
    entryStreamPosition = m_Stream.GetPosition()
    
    'Next, look for the layer bank block
    Dim idxLayerBank As Long
    idxLayerBank = -1
    
    Dim i As Long
    For i = 0 To m_numOfBlocks - 1
        If (m_Blocks(i).GetBlockID = PSP_LAYER_START_BLOCK) Then
            idxLayerBank = i
            Exit For
        End If
    Next i
    
    'If layers aren't found, PD will automatically attempt recovery by using an embedded composite image
    ' (if one exists)
    If (idxLayerBank < 0) Then
        InternalError funcName, "no layer bank!"
        Step6_LoadLayers = psp_Failure
        Exit Function
    End If
    
    'Move the stream pointer into position at the start of the layer bank block.
    m_Stream.SetPosition m_Blocks(idxLayerBank).GetBlockOffset(), FILE_BEGIN
    
    'Each layer in this image has its own sub-block within this segment.  PSPs support many
    ' different layer types, and unlike PSDs, they may not supply raster "backup" copies of
    ' vector, text, or adjustment layers.  As such, we need to be careful about what we
    ' attempt to load.
    
    'To start, we first want to just load some header data for each layer into a central
    ' layer directory.
    m_LayerCount = m_ImageHeader.psph_LayerCount
    ReDim m_Layers(0 To m_LayerCount - 1) As pdPSPLayer
    
    Dim numLayersFound As Long, numLayersAttempted As Long
    numLayersFound = 0
    numLayersAttempted = 0
    
    'Before doing anything else, ensure we're actually pointing at a layer block now.
    Dim testBlockID As PSPBlockID, testBlockLength As Long
    If PeekNextBlock(PSP_BLOCK_MARKER, testBlockID, testBlockLength) Then
        If (testBlockID <> PSP_LAYER_BLOCK) Then
            InternalError funcName, "not pointing at the layer bank block!"
            okToProceed = psp_Failure
        End If
    Else
        InternalError funcName, "not pointing at a block!"
        okToProceed = psp_Failure
    End If
    
    If (okToProceed >= psp_Failure) Then Exit Function
    
    'Groups can be imported, but they're complex to handle correctly.  Note that group behavior
    ' is *not* guaranteed to work in malformed files, as PSPs use a weird counting mechanism for
    ' establishing group boundaries, and in broken files, we can't count layers accurately.
    Dim groupCount() As Long, groupReferences() As pdPSPLayer, idxGroup As Long, addGroupStartMarker As Boolean
    ReDim groupCount(0 To m_LayerCount - 1) As Long
    ReDim groupReferences(0 To m_LayerCount - 1) As pdPSPLayer
    idxGroup = 0
    
    'We're pointing at a layer block (hopefully, lol)
    Do
        
        addGroupStartMarker = False
        
        'It shouldn't be possible to find more layers than the header specifies, but if it happens,
        ' we attempt to deal with it gracefully.
        If (numLayersFound > UBound(m_Layers)) Then ReDim Preserve m_Layers(0 To numLayersFound * 2 - 1) As pdPSPLayer
        Set m_Layers(numLayersFound) = New pdPSPLayer
        
        'If this layer is parsed successfully, increment the layer counter.  (If something is wrong
        ' with this layer, or PD doesn't know how to interpret it, we'll just skip it.)
        If (m_Layers(numLayersFound).LoadLayer(m_Stream, m_Warnings, m_ImageHeader) = psp_Success) Then
            
            'Regardless of layer correctness, we need to deal with group boundaries.
            ' If we're inside a running group, decrement the current group counter no matter what,
            ' and flag this layer with the current group ID (if any).  Note that we do this even
            ' if the current layer is a group marker, so that nested groups work correctly.
            If (idxGroup > 0) Then
            
                groupCount(idxGroup) = groupCount(idxGroup) - 1
                
                'Last layer in group; move the index down one
                If (groupCount(idxGroup) = 0) Then
                    addGroupStartMarker = True
                    idxGroup = idxGroup - 1
                End If
                
            End If
            
            'If this layer appears to be valid, keep processing it
            If m_Layers(numLayersFound).IsLayerOK Then
                
                'If the layer actually has pixels, generate a DIB for it
                If m_Layers(numLayersFound).IsLayerPixelBased() Then
                
                    If m_Layers(numLayersFound).ConstructMatchingDIB(m_ParentPalette, m_ParentPaletteCount, m_ImageHeader) Then
                        
                        'Apply color-management, if any
                        If (Not m_ICCProfile Is Nothing) And ColorManagement.UseEmbeddedICCProfiles() Then
                            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction vbTab & "Color-managing layer contents..."
                            ColorManagement.ConvertDIBToSRGB m_Layers(numLayersFound).GetLayerDIB, m_ICCProfile
                        End If
                        
                    End If
                    
                'Non-pixel-based layers are fine (e.g. groups), but we don't need to produce
                ' a DIB for them, obviously!
                Else
                    
                    'On group markers, update our running group count table.  (See a few lines up
                    ' for how we use this to "tag" incoming layers)
                    If (m_Layers(numLayersFound).GetLayerType = keGLTGroup) Then
                        
                        'Increase the depth of the nested group tree
                        idxGroup = idxGroup + 1
                        If (idxGroup > UBound(groupCount)) Then
                            ReDim Preserve groupCount(0 To idxGroup * 2 - 1) As Long
                            ReDim Preserve groupReferences(0 To idxGroup * 2 - 1) As pdPSPLayer
                        End If
                        
                        'Store a reference to this marker; when we create a matching "start group"
                        ' layer, we need to copy relevant properties (e.g. visibility) of this layer.
                        groupCount(idxGroup) = m_Layers(numLayersFound).GetNumLayersInGroup()
                        Set groupReferences(idxGroup) = m_Layers(numLayersFound)
                        
                    'On vector layers, attempt to build a raster version of the underlying vector data
                    ElseIf (m_Layers(numLayersFound).GetLayerType = keGLTVector) Then
                        
                        If m_Layers(numLayersFound).ConstructMatchingDIB_Vector(m_ParentPalette, m_ParentPaletteCount, m_ImageHeader) Then
                            
                            'Apply color-management, if any
                            If (Not m_ICCProfile Is Nothing) And ColorManagement.UseEmbeddedICCProfiles() Then
                                If PSP_DEBUG_VERBOSE Then PDDebug.LogAction vbTab & "Color-managing layer contents..."
                                ColorManagement.ConvertDIBToSRGB m_Layers(numLayersFound).GetLayerDIB, m_ICCProfile
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
                'This layer was successfully constructed!  If its index matches the original "active layer index",
                ' mark it as such, and be sure to increment the "good layer count" before exiting
                If (numLayersAttempted = m_ImageHeader.psph_ActiveLayer) Then m_Layers(numLayersFound).MarkAsActiveLayer
                numLayersFound = numLayersFound + 1
                
            End If
            
        '/layer didn't load correctly
        Else
            
            'If we're inside a running group, decrement the current group counter even
            ' though we didn't load this image.  This ensures that our group boundaries
            ' persist in the correct places.
            If (idxGroup > 0) Then
                groupCount(idxGroup) = groupCount(idxGroup) - 1
                If (groupCount(idxGroup) = 0) Then
                    addGroupStartMarker = True
                    idxGroup = idxGroup - 1
                End If
            End If
            
        End If
        
        'If we need to add a new "start group" marker, do so now
        If addGroupStartMarker Then
            If (numLayersFound > UBound(m_Layers)) Then ReDim Preserve m_Layers(0 To numLayersFound * 2 - 1) As pdPSPLayer
            Set m_Layers(numLayersFound) = New pdPSPLayer
            With groupReferences(idxGroup + 1)
                m_Layers(numLayersFound).LoadLayer_GroupMarker .GetLayerName, .GetLayerVisibility, .GetLayerBlendMode_PSP, .GetLayerOpacity
            End With
            numLayersFound = numLayersFound + 1
        End If
        
        'Keep track of the number of layers we've attempted to load; this is relevant for tracking
        ' original layer indices, which we need to restore the saved "active" layer index
        numLayersAttempted = numLayersAttempted + 1
        
    'If we're still pointing at a layer block, continue loading!
    Loop While PeekNextBlock(PSP_BLOCK_MARKER, testBlockID, testBlockLength) And (testBlockID = PSP_LAYER_BLOCK)
    
    'All layers loaded!
    
    'Reset the central "layer count" to match only the number of *valid* layers found
    m_LayerCount = numLayersFound
    
    'Forcibly align the stream pointer before continuing
    m_Stream.SetPosition m_Blocks(idxLayerBank).GetBlockOffset() + m_Blocks(idxLayerBank).GetBlockLength(), FILE_BEGIN
    Step6_LoadLayers = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 6, #" & Err.Number & ": " & Err.Description
    Step6_LoadLayers = psp_Failure

End Function

'Do *NOT* call import steps out-of-order!
Private Function Step7_AssembleFinalImage(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "Step7_AssembleFinalImage"
    
    Step7_AssembleFinalImage = psp_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError funcName, "filename has changed since original validation!"
        Step7_AssembleFinalImage = psp_Failure
        Exit Function
    End If
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'We now have enough data to assemble the final image.
    Dim newLayerID As Long, tmpLayer As pdLayer
    Dim compositeLayerLoaded As Boolean
    
    Dim i As Long
    
    If (m_LayerCount <= 0) Or PSP_LOAD_COMPOSITE_AS_BASE_LAYER Then
    
        okToProceed = Step4_ReadCompositeBlock(srcFile, True)
        
        If (okToProceed < psp_Failure) And (m_numCompositeImages > 0) Then
            
            'Find the largest composite image that was loaded successfully
            Dim bestStats As Currency, bestIndex As Long
            bestIndex = -1
            
            For i = 0 To m_numCompositeImages - 1
                If m_CompositeImages(i).cih_DIBOK Then
                    
                    'Favor larger images (with a bonus for lossless compression)
                    Dim tmpStats As Currency
                    tmpStats = m_CompositeImages(i).cih_DIB.GetDIBWidth * m_CompositeImages(i).cih_DIB.GetDIBHeight * m_CompositeImages(i).cih_BitDepth
                    If (m_CompositeImages(i).cih_Compression <> PSP_COMP_JPEG) Then tmpStats = tmpStats + 1
                    
                    If (tmpStats > bestStats) Then
                        bestStats = tmpStats
                        bestIndex = i
                    End If
                    
                End If
            Next i
            
            'If we found a useable composite image, dump it into a pdImage struct so we have
            ' *something* to work with for the moment
            If (bestIndex >= 0) Then
                
                If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Using composite image #" & bestIndex & " (" & GetCompressionName(m_CompositeImages(bestIndex).cih_Compression) & ") to load PSP file..."
                
                With m_CompositeImages(bestIndex)
                    
                    'Set basic image attributes
                    dstImage.Width = .cih_Width
                    dstImage.Height = .cih_Height
                    
                    'Prep a new layer object and initialize it with the composite image
                    newLayerID = dstImage.CreateBlankLayer()
                    Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                    
                    Dim compositeLayerName As String
                    If (m_LayerCount <= 0) Then
                        compositeLayerName = g_Language.TranslateMessage("Composite image")
                    Else
                        compositeLayerName = "Composite image (debug purposes only)"
                    End If
                    tmpLayer.InitializeNewLayer PDL_Image, "Composite image (debug purposes only)", .cih_DIB
                    
                    'If layers exist in the file, set the composite layer to invisible (since it's only being
                    ' loaded for debugging purposes)
                    tmpLayer.SetLayerVisibility (m_LayerCount <= 0)
                    compositeLayerLoaded = True
                    
                End With
                
                okToProceed = psp_Success
                
            '/composite images exist but don't appear to be useable?
            Else
                InternalError funcName, "composite image exists, but PD can't decode it"
                okToProceed = psp_Failure
            End If
        
        '/no composite image found.  That's just fine - we just won't have the composite image to
        ' "fall back" on if we can't retrieve any usable layer data.
        'Else
            '(do nothing)
        End If
    
    '/end composite image handling
    End If
    
    'We don't actually need anything from the source PSP file now, so go ahead and close the stream
    If m_Stream.IsOpen() Then m_Stream.StopStream True
    
    'Before proceeding, ensure that at least one good layer was found
    If (m_LayerCount > 0) And (okToProceed < psp_Failure) Then
    
        'We'll now attempt to import as many layers as we can!
        
        'Layer groups are not currently supported by PD, but we do use group data to enable group
        ' visibility/invisibility when loading groups.  Note that PSPs support nested layer groups,
        ' so a layer can be within e.g. a whole bunch of parent groups with varying depths.
        ' I currently use a formula where if *any* of the current parent layers are invisible,
        ' the current layer gets hidden too.  This produces the most "useful-looking" document
        ' (in lieu of actually supporting groups ourselves, ha).
        Dim curGroupDepth As Long
        curGroupDepth = 0
        
        'I don't know if PSPs enforce a "maximum" group depth.  Photoshop limits group depth to 10,
        ' so we start there.  (Note that PD will auto-increment this if as it encounters complex
        ' trees; we just need an arbitrary starting size for tracking!)
        Const MAX_GROUP_DEPTH As Long = 10
        Dim groupVisibility() As Boolean
        ReDim groupVisibility(0 To MAX_GROUP_DEPTH - 1) As Boolean
        
        'If any group *above* the current one is invisible, the current layer's visibility flag
        ' gets overriden with the parent group.  This is my preferred behavior as it ensures that
        ' layers belonging to invisible groups are not visible by default in the image PD constructs.
        ' (We update this flag whenever group depth changes.)
        Dim curGroupVisibility As Boolean
        curGroupVisibility = True
        
        'Start by setting basic image attributes
        dstImage.Width = m_ImageHeader.psph_Width
        dstImage.Height = m_ImageHeader.psph_Height
        
        'Dummy 1x1 DIBs are used for group markers
        Dim dummyDIB As pdDIB
        Set dummyDIB = New pdDIB
        dummyDIB.CreateBlank 1, 1, 32, 0, 0
        dummyDIB.SetInitialAlphaPremultiplicationState True
                        
        'If possible, we'll attempt to restore the original active layer index
        Dim activeLayerIndex As Long
        activeLayerIndex = -1
        
        For i = 0 To m_LayerCount - 1
        
            'Double-check layer is OK
            If (m_Layers(i) Is Nothing) Then GoTo SkipToNextLayer
            If (Not m_Layers(i).IsLayerOK) Then GoTo SkipToNextLayer
            
            'For layers with pixel contents, load them next
            If m_Layers(i).IsLayerPixelBased() Then
                
                'Prep a new layer object
                newLayerID = dstImage.CreateBlankLayer()
                Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                
                'Initialize the layer DIB (using the source layer's name, too!)
                tmpLayer.InitializeNewLayer PDL_Image, m_Layers(i).GetLayerName(), m_Layers(i).GetLayerDIB()
                
                'Set matching layer properties
                tmpLayer.SetLayerOpacity m_Layers(i).GetLayerOpacity()
                tmpLayer.SetLayerOffsetX m_Layers(i).GetLayerOffsetX()
                tmpLayer.SetLayerOffsetY m_Layers(i).GetLayerOffsetY()
                tmpLayer.SetLayerVisibility m_Layers(i).GetLayerVisibility()
                tmpLayer.SetLayerBlendMode m_Layers(i).GetLayerBlendMode()
                
                'If we're inside a layer group, override the current layer visibility with its
                ' parent group visibility.
                If (curGroupDepth > 0) Then
                    If (Not curGroupVisibility) Then tmpLayer.SetLayerVisibility False
                End If
                
                'Check active layer status (and note that we have to manually increment the index if debug mode
                ' is active and we loaded the composite layer as a "temporary" base layer)
                If m_Layers(i).GetLayerWasActiveLayer() Then
                    If compositeLayerLoaded Then activeLayerIndex = i + 1 Else activeLayerIndex = i
                End If
                
            'This layer doesn't have pixels
            Else
                
                'Is this a group marker?
                If (m_Layers(i).GetLayerType = keGLTGroup) Then
                    
                    'Add this group's visibility and name to our running collection.  (We track this info
                    ' so we can properly handle complex nested group arrangements.)
                    If (Not m_Layers(i).IsLayerGroupStart()) Then
                        
                        If (curGroupDepth > UBound(groupVisibility)) Then ReDim Preserve groupVisibility(0 To curGroupDepth * 2 - 1) As Boolean
                        groupVisibility(curGroupDepth) = m_Layers(i).GetLayerVisibility()
                        curGroupVisibility = curGroupVisibility And groupVisibility(curGroupDepth)
                        
                        'Increment nested group counter
                        curGroupDepth = curGroupDepth + 1
                    
                    'Because layers are traversed in reverse-order, we actually use group *end* markers
                    ' for tracking visibility state.  When we reach a group *start* marker, we need to
                    ' reset current visibility flags to match the new nested depth.
                    Else
                            
                        'Decrement depth counter, with a failsafe check for malformed files
                        curGroupDepth = curGroupDepth - 1
                        If (curGroupDepth < 0) Then
                            curGroupDepth = 0
                            InternalError funcName, "malformed layer groups found; curGroupDepth below zero!"
                        End If
                        
                        'Reset the group visibility flag
                        curGroupVisibility = True
                        
                        'If we're still inside a group, look for *any* invisibility flags above us.
                        ' If any are found, set the current visibility flag to FALSE.
                        If (curGroupDepth > 0) Then
                            Dim gvIndex As Long
                            For gvIndex = 0 To curGroupDepth - 1
                                If (Not groupVisibility(gvIndex)) Then
                                    curGroupVisibility = False
                                    Exit For
                                End If
                            Next gvIndex
                        End If
                
                    End If
                    
                    'Because layers are traversed from bottom-to-top, this indicator
                    ' marks the *end* of a group.  Insert a dummy group-end layer.
                    If PSP_GROUPS_AS_DUMMY_LAYERS Then
                            
                        'Prep a new "dummy" layer object to represent this group
                        newLayerID = dstImage.CreateBlankLayer()
                        Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                        
                        If m_Layers(i).IsLayerGroupStart Then
                            tmpLayer.InitializeNewLayer PDL_Image, g_Language.TranslateMessage("Group start:") & " " & m_Layers(i).GetLayerName(), dummyDIB, True
                        Else
                            tmpLayer.InitializeNewLayer PDL_Image, g_Language.TranslateMessage("Group end:") & " " & m_Layers(i).GetLayerName(), dummyDIB, True
                        End If
                        
                        'Fill in any remaining layer properties
                        With m_Layers(i)
                            tmpLayer.SetLayerBlendMode .GetLayerBlendMode()
                            tmpLayer.SetLayerOpacity .GetLayerOpacity()
                            tmpLayer.SetLayerVisibility m_Layers(i).GetLayerVisibility()
                        End With
                
                    End If
                    
                'Is this a vector layer?
                ElseIf (m_Layers(i).GetLayerType = keGLTVector) Then
                    
                    'Load any non-text vector contents first
                    If m_Layers(i).IsVectorLayerAvailableAsRaster() Then
                        
                        'Prep a new layer object
                        newLayerID = dstImage.CreateBlankLayer()
                        Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                        
                        'Initialize the layer DIB (using the source layer's name, too!)
                        tmpLayer.InitializeNewLayer PDL_Image, m_Layers(i).GetLayerName(), m_Layers(i).GetLayerDIB()
                        
                        'Trim the layer to its smallest relevant size; this will set x/y offsets automatically.
                        ' Note that some vector layers may not be rendered correctly by PD (typically due to
                        ' esoteric brushes being used, which we can't reproduce accurately).  If this occurs,
                        ' the layer will be fully transparent, with no useable contents.  There's not really a
                        ' point to keeping layers like this, so delete them on-the-fly if we encounter them.
                        If (Not tmpLayer.CropNullPaddedLayer) Then
                            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "WARNING: skipping vector layer " & m_Layers(i).GetLayerName() & "; no useable render"
                            dstImage.DeleteLayerByID newLayerID
                        
                        'The layer exists and is useable!
                        Else
                        
                            'Set matching layer properties
                            tmpLayer.SetLayerOpacity m_Layers(i).GetLayerOpacity()
                            tmpLayer.SetLayerVisibility m_Layers(i).GetLayerVisibility()
                            tmpLayer.SetLayerBlendMode m_Layers(i).GetLayerBlendMode()
                            
                            'Check active layer status (and note that we have to manually increment the index if debug mode
                            ' is active and we loaded the composite layer as a "temporary" base layer)
                            If m_Layers(i).GetLayerWasActiveLayer() Then
                                If compositeLayerLoaded Then activeLayerIndex = i + 1 Else activeLayerIndex = i
                            End If
                            
                        End If
                        
                    End If
                    
                    'Run a second pass through the object, looking for (and extracting) text shapes
                    ' into corresponding PD text layers.
                    If m_Layers(i).DoesVectorLayerContainText() Then
                        
                        Dim listOfTextObjects As pdStack, numOfTextObjects As Long
                        numOfTextObjects = m_Layers(i).GetListOfTextIndices(listOfTextObjects)
                        
                        If (numOfTextObjects > 0) Then
                            
                            'Iterate all discovered text layers and add them to the image as PD-format text layers
                            Dim idxText As Long, idxShape As Long
                            For idxText = 0 To numOfTextObjects - 1
                                idxShape = listOfTextObjects.GetInt(idxText)
                                newLayerID = m_Layers(i).GetVectorShape(idxShape).CreateTextLayerNow(dstImage, m_Layers(i))
                            Next idxText
                            
                        End If
                        
                    End If
                
                End If
                
            End If
        
SkipToNextLayer:
        Next i
        
        'If an active layer index exists, set it now
        If (activeLayerIndex >= 0) Then dstImage.SetActiveLayerByIndex activeLayerIndex
        
        'All layers have been loaded!
        okToProceed = psp_Success
        
    '/end useable layers exist in file
    End If
    
    'Set any remaining image parameters
    If (okToProceed < psp_Failure) Then
        
        'DPI is automatically converted from PSP's format to PD's by the GetDPI() function in this class
        dstImage.SetDPI GetDPI(), GetDPI()
        
    End If
    
    'Our work here is done!
    Step7_AssembleFinalImage = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 7, #" & Err.Number & ": " & Err.Description
    Step7_AssembleFinalImage = psp_Failure

End Function

'Given an index into the m_CompositeImages() array, retrieve and decode the associated JPEG data
' into a standard pdDIB instance.  The stream pointer *MUST ALREADY BE IN THE CORRECT POSITION*,
' and *YOU MUST MANUALLY RE-POSITION THE POINTER AFTER CALLING THIS FUNCTION*.  The PSP spec is
' ambiguous about things like padding between blocks, so block length must *always* be used to
' forcibly align the stream pointer between blocks.
Private Function RetrieveComposite_JPEG(ByVal imgIndex As Long, ByVal parentBlockLength As Long) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "RetrieveComposite_JPEG"
    
    RetrieveComposite_JPEG = psp_Success
    
    'Failsafe check for JPEG format
    If (m_CompositeImages(imgIndex).cih_Compression <> PSP_COMP_JPEG) Then
        InternalError funcName, "JPEG compression mismatch"
        RetrieveComposite_JPEG = psp_Failure
        Exit Function
    End If
    
    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Attempting to retrieve composite image (" & m_CompositeImages(imgIndex).cih_Width & "x" & m_CompositeImages(imgIndex).cih_Height & ") in JPEG format..."
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'The JPEG information chunk is pretty straightforward: just a few size descriptors,
    ' and then the entire embedded JPEG file as a byte-stream.
    Dim startOffset As Long
    startOffset = m_Stream.GetPosition()
    
    'Start by validating the block length against the reported chunk size.
    ' (The JPEG chunk must always fit within its parent block.)
    Dim chunkSize As Long
    chunkSize = m_Stream.ReadLong()
    If (parentBlockLength < chunkSize) Or (chunkSize <= 0) Then
        InternalError funcName, "JPEG chunk size mismatch"
        RetrieveComposite_JPEG = psp_Failure
        Exit Function
    End If
    
    Dim compressedSize As Long, uncompressedSize As Long, imgType As PSPDIBType
    compressedSize = m_Stream.ReadLong()
    uncompressedSize = m_Stream.ReadLong()
    imgType = m_Stream.ReadIntUnsigned()
    Debug.Print "JPEG", compressedSize, uncompressedSize, imgType
    'Expansion bytes are explicitly allowed by the spec, so reorient the stream pointer manually
    ' using the previous chunk length.
    m_Stream.SetPosition startOffset + chunkSize, FILE_BEGIN
    
    'The stream now contains [compressedSize] bytes of JPEG data.  Use pd2D to load the image into
    ' a standard pdDIB.
    Dim tmpSurface As pd2DSurface
    Set tmpSurface = New pd2DSurface
    If tmpSurface.CreateSurfaceFromPtr(m_Stream.Peek_PointerOnly(peekLength:=compressedSize), compressedSize) Then
        
        'Point this frame's DIB at the decoded image, then free the underlying surface
        Set m_CompositeImages(imgIndex).cih_DIB = tmpSurface.GetSurfaceDIB
        Set tmpSurface = Nothing
        
        'Ensure the loaded DIB matches required PD settings
        m_CompositeImages(imgIndex).cih_DIBOK = (Not m_CompositeImages(imgIndex).cih_DIB Is Nothing)
        If m_CompositeImages(imgIndex).cih_DIBOK Then
            If (m_CompositeImages(imgIndex).cih_DIB.GetDIBColorDepth < 32) Then m_CompositeImages(imgIndex).cih_DIB.ConvertTo32bpp
            If (Not m_CompositeImages(imgIndex).cih_DIB.GetAlphaPremultiplication) Then m_CompositeImages(imgIndex).cih_DIB.SetAlphaPremultiplication True
            okToProceed = psp_Success
        Else
            okToProceed = psp_Warning
        End If
    
    'Failure; possibly a bad JPEG stream, or an esoteric implementation incompatible with GDI+
    Else
        m_CompositeImages(imgIndex).cih_DIBOK = False
        okToProceed = psp_Warning
    End If
    
    'Our work here is done!
    RetrieveComposite_JPEG = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in JPEG frame retrieval, #" & Err.Number & ": " & Err.Description
    RetrieveComposite_JPEG = psp_Failure
    
End Function

'Given an index into the m_CompositeImages() array, retrieve and decode the associated image data
' into a standard pdDIB instance.  The stream pointer *MUST ALREADY BE IN THE CORRECT POSITION*,
' and *YOU MUST MANUALLY RE-POSITION THE POINTER AFTER CALLING THIS FUNCTION*.  The PSP spec is
' ambiguous about things like padding between blocks, so block length must *always* be used to
' forcibly align the stream pointer between blocks.
Private Function RetrieveComposite_Normal(ByVal imgIndex As Long, ByVal parentBlockLength As Long) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "RetrieveComposite_Normal"
    
    RetrieveComposite_Normal = psp_Success
    
    'Failsafe check for valid compression schemes
    If (m_CompositeImages(imgIndex).cih_Compression <> PSP_COMP_NONE) And (m_CompositeImages(imgIndex).cih_Compression <> PSP_COMP_RLE) And (m_CompositeImages(imgIndex).cih_Compression <> PSP_COMP_LZ77) Then
        InternalError funcName, "frame compression mismatch"
        RetrieveComposite_Normal = psp_Failure
        Exit Function
    End If
    
    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Attempting to retrieve composite image in ""normal"" format..."
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'The composite image information chunk is pretty straightforward: a small header chunk,
    ' followed by an optional palette chunk, then individual channel blocks (with possible
    ' alpha channel block following).
    
    'Start by noting the initial offset of the stream object (we'll orient all positioning
    ' decisions against this)
    Dim startOffset As Long
    startOffset = m_Stream.GetPosition()
    
    'Next, validate the parent block length against the reported chunk size.
    ' (All subesquent sub-blocks and chunks must always fit within the parent block's boundaries.)
    Dim chunkSize As Long
    chunkSize = m_Stream.ReadLong()
    If (parentBlockLength < chunkSize) Or (chunkSize <= 0) Then
        InternalError funcName, "composite image chunk size mismatch"
        RetrieveComposite_Normal = psp_Failure
        Exit Function
    End If
    
    'The header is small: it simply tells us...
    ' 1) the number of bitmaps that follow (1 for opaque images, 2 for transparent images)
    ' 2) the number of channels that follow (1, 2, 3, or 4 for gray/indexed and/or RGB with/without alpha)
    Dim bitmapCount As Long, channelCount As Long
    bitmapCount = m_Stream.ReadIntUnsigned()
    channelCount = m_Stream.ReadIntUnsigned()
    
    'Validate both values
    If (bitmapCount < 1) Or (bitmapCount > 2) Then
        InternalError funcName, "bad bitmapCount (" & bitmapCount & ")"
        RetrieveComposite_Normal = psp_Warning
        Exit Function
    End If
    
    If (channelCount < 1) Or (channelCount > 4) Then
        InternalError funcName, "bad channelCount (" & channelCount & ")"
        RetrieveComposite_Normal = psp_Warning
        Exit Function
    End If
    
    'The spec mentions future expansion fields as an explicit possibility, so use the chunk length
    ' to position ourselves at the start of the next segment.
    m_Stream.SetPosition startOffset + chunkSize, FILE_BEGIN
    
    'We now point at one of two things: either a palette sub-block (only valid if the parent image
    ' is non-grayscale with a color-depth less than 24-bit), or the channel sub-blocks themselves.
    Dim blockID As PSPBlockID, blockLength As Long
    
    Dim paletteAvailable As Boolean
    paletteAvailable = PeekNextBlock(PSP_BLOCK_MARKER, blockID, blockLength)
    paletteAvailable = paletteAvailable And (blockID = PSP_COLOR_BLOCK)
    
    If PSP_DEBUG_VERBOSE Then
        PDDebug.LogAction "Composite image has " & bitmapCount & " bitmaps, " & channelCount & " channels, and palette status is " & paletteAvailable
    End If
    
    'Ensure that palette status matches color-count
    If (paletteAvailable And ((m_ImageHeader.psph_BitDepth > 8) Or (channelCount > 2))) Then
        InternalError funcName, "palette exists but shouldn't!"
        RetrieveComposite_Normal = psp_Warning
        Exit Function
    End If
    
    'If a palette is not available, but the image is monochrome or grayscale,
    ' generate a default gray palette.
    If (m_ImageHeader.psph_IsGrayscale Or m_ImageHeader.psph_BitDepth <= 8) And (Not paletteAvailable) Then
        
        If (m_ImageHeader.psph_BitDepth = 8) Or (m_ImageHeader.psph_BitDepth = 16) Then
            m_CompositePaletteCount = 256
            Palettes.GetPalette_Grayscale m_CompositePalette
        ElseIf (m_ImageHeader.psph_BitDepth = 4) Then
            m_CompositePaletteCount = 16
            Palettes.GetPalette_GrayscaleEx m_CompositePalette, m_CompositePaletteCount
        Else
            m_CompositePaletteCount = 2
            Palettes.GetPalette_GrayscaleEx m_CompositePalette, m_CompositePaletteCount
        End If
        
    End If
    
    'If the palette exists, immediately parse it - we'll need it momentarily to interpret
    ' channel data.
    If paletteAvailable Then
        If (Not RetrievePalette(m_CompositePalette, m_CompositePaletteCount)) Then
            InternalError funcName, "bad palette - can't retrieve composite image"
            RetrieveComposite_Normal = psp_Warning
            Exit Function
        End If
    End If
    
    'A gray or color palette should now exist for any image that requires it.
    
    'Note also that the palette reader forcibly aligns the stream reader correctly, so we don't
    ' need to mess with that here.
    
    'Now we are pointing at the composite image channel sub-blocks.  Per the spec, there will be...
    ' 1) 1-3 channel sub-blocks with pixel data.  1 channel for paletted bitmap, 3 for RGB data
    ' 2) optional 1 bonus channel for the transparency mask (stored as an 8-bit greyscale bitmap)
    '
    'We first want to retrieve and decompress all channels into their own channel objects
    Dim cmpChannels() As pdPSPChannel
    ReDim cmpChannels(0 To channelCount - 1) As pdPSPChannel
    
    Dim i As Long
    For i = 0 To channelCount - 1
        
        Set cmpChannels(i) = New pdPSPChannel
        
        'Perform a failsafe peek before attempting to load the channel
        
        'Block marker found
        If PeekNextBlock(PSP_BLOCK_MARKER, blockID, blockLength) Then
            
            'Block is a channel block
            If (blockID = PSP_CHANNEL_BLOCK) Then
            
                'Populate a temporary header to pass to the target channel.  This contains essential
                ' details about the parent image, as well as external details relevant to interpreting
                ' channel contents - like compression type!
                Dim tmpHeader As PSP_ChannelHeader
                With tmpHeader
                    .ch_ParentVersionMajor = m_ImageHeader.psph_VersionMajor
                    .ch_ParentWidth = m_CompositeImages(imgIndex).cih_Width
                    .ch_ParentHeight = m_CompositeImages(imgIndex).cih_Height
                    .ch_ParentBitDepth = m_CompositeImages(imgIndex).cih_BitDepth
                    .ch_Compression = m_CompositeImages(imgIndex).cih_Compression
                End With
                
                If (cmpChannels(i).LoadChannel(m_Stream, m_Warnings, tmpHeader) >= psp_Failure) Then
                    InternalError funcName, "catastrophic failure; abandoning file entirely..."
                    RetrieveComposite_Normal = psp_Failure
                    Exit Function
                
                'This branch was useful during initial channel parse creation, but it's overkill now
                'Else
                '    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Channel " & (i + 1) & " looks okay"
                End If
            
            '/Channel block ID incorrect
            Else
                InternalError funcName, "channel block expected but found " & blockID & " instead??  Attempting recovery..."
                m_Stream.SetPosition blockLength + 10, FILE_CURRENT
            End If
        
        '/Start-of-block marker incorrect
        Else
            InternalError funcName, "channel block expected but bad block marker encountered"
            RetrieveComposite_Normal = psp_Failure
            Exit Function
        End If
        
    Next i
    
    'With all channels decoded, construct a DIB for this entry!
    If (okToProceed < psp_Failure) Then
        m_CompositeImages(imgIndex).cih_DIBOK = ImageFormats_PSP.PSP_BuildDIBFromChannels(channelCount, cmpChannels, tmpHeader, m_CompositeImages(imgIndex).cih_DIB, m_CompositePalette, m_CompositePaletteCount)
    End If
    
    'Our work here is done!
    RetrieveComposite_Normal = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in JPEG frame retrieval, #" & Err.Number & ": " & Err.Description
    RetrieveComposite_Normal = psp_Failure
    
End Function

'If the stream is pointing at a palette block (PSP_COLOR_BLOCK), call this function to read
' the palette into an RGBQuad array.
Private Function RetrievePalette(ByRef dstPalette() As RGBQuad, ByRef dstPaletteSize As Long) As Boolean
    
    On Error GoTo InternalVBError
    Const funcName As String = "RetrievePalette"
    
    'This function is pessimistic - assume failure unless otherwise noted
    RetrievePalette = False
    
    Dim origStreamPosition As Long
    origStreamPosition = m_Stream.GetPosition()
    
    'Ensure correct chunk ID and size
    If (m_Stream.ReadLong() <> PSP_BLOCK_MARKER) Then
        InternalError funcName, "stream misaligned"
        m_Stream.SetPosition origStreamPosition, FILE_BEGIN
        Exit Function
    End If
    
    If (m_Stream.ReadIntUnsigned() <> PSP_COLOR_BLOCK) Then
        InternalError funcName, "not a palette block"
        m_Stream.SetPosition origStreamPosition, FILE_BEGIN
        Exit Function
    End If
    
    Dim blockLength As Long, palHeaderSize As Long
    
    'Early PSP versions use an asinine mechanism for struct length descriptors, so this is
    ' a stupid (but necessary) workaround
    If (m_ImageHeader.psph_VersionMajor <= 3) Then
        palHeaderSize = m_Stream.ReadLong()
        blockLength = m_Stream.ReadLong()
    Else
        blockLength = m_Stream.ReadLong()
    End If
    
    If (blockLength <= 0) Then
        InternalError funcName, "bad palette length"
        m_Stream.SetPosition origStreamPosition, FILE_BEGIN
        Exit Function
    End If
    
    'If we're still here, this appears to be a valid palette block.  Note the current stream position,
    ' then parse palette information.
    Dim posPaletteStart As Long
    posPaletteStart = m_Stream.GetPosition()
    
    'Palettes have a brief header.  The spec explicitly allows the header to have variable-size,
    ' so we must use its embedded length to align the stream at the start of the color list.
    If (m_ImageHeader.psph_VersionMajor > 3) Then palHeaderSize = m_Stream.ReadLong()
    If (palHeaderSize <= 0) Then
        InternalError funcName, "bad palette info header size"
        m_Stream.SetPosition origStreamPosition, FILE_BEGIN
        Exit Function
    End If
    
    dstPaletteSize = m_Stream.ReadLong()
    If (dstPaletteSize <= 0) Then
        InternalError funcName, "bad palette color count"
        m_Stream.SetPosition origStreamPosition, FILE_BEGIN
        Exit Function
    End If
    
    'Manually align the stream pointer
    m_Stream.SetPosition posPaletteStart + palHeaderSize, FILE_BEGIN
    
    'Retrieve (dstPaletteSize) colors, in RGBA format
    ReDim dstPalette(0 To dstPaletteSize - 1) As RGBQuad
    
    Dim i As Long, tmpColor As Long
    For i = 0 To dstPaletteSize - 1
    
        tmpColor = m_Stream.ReadLong()
        PutMem4 VarPtr(dstPalette(i)), tmpColor
        
        'PSP spec defines RGBQuads as RGBx; manually fill alpha to compensate
        dstPalette(i).Alpha = 255
        
    Next i
    
    'Palette read successfully!
    RetrievePalette = True
    
    'Before exiting, forcibly align the stream using the original block length
    m_Stream.SetPosition posPaletteStart + blockLength
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in palette retrieval, #" & Err.Number & ": " & Err.Description
    RetrievePalette = False
    
End Function

Private Sub Class_Initialize()
    Set m_Stream = New pdStream
    Me.Reset
End Sub

Private Sub Class_Terminate()
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen() Then m_Stream.StopStream True
    End If
End Sub

Private Function GetCompressionName(ByVal cmpType As PSPCompression) As String
    Select Case cmpType
        Case PSP_COMP_NONE
            GetCompressionName = "uncompressed"
        Case PSP_COMP_RLE
            GetCompressionName = "RLE"
        Case PSP_COMP_LZ77
            GetCompressionName = "zLib"
        Case PSP_COMP_JPEG
            GetCompressionName = "JPEG"
        Case Else
            GetCompressionName = "(unknown)"
    End Select
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    
    Dim errText As String
    errText = "pdPSP." & funcName & "() reported an error on file """ & m_SourceFilename & """: " & errDescription
    
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction errText
    Else
        Debug.Print errText
    End If
    
End Sub

'Peek the next 10 bytes in the stream to see if...
' 1) They represent a block (pass expectedHeader with e.g. 0x004b427e)
' 2) What the block ID is (only relevant if (1) passes)
' 3) What the block's length is (only relevant if (2) passes)
'
'IMPORTANTLY: regardless of what this function returns, the stream pointer will be restored
' to its original location at load-time.  This removes ambiguity based on the return value
' of the function or various potential failure states.
Private Function PeekNextBlock(ByVal expectedHeader As Long, ByRef dstBlockID As PSPBlockID, ByRef dstBlockLength As Long) As Boolean
    
    'Make a note of the entry stream position
    Dim origPosition As Long
    origPosition = m_Stream.GetPosition()
    
    'Check header first
    PeekNextBlock = (m_Stream.ReadLong() = expectedHeader)
    If PeekNextBlock Then
    
        'Valid block ID.  Retrieve 2-byte ID and 4-byte length.
        dstBlockID = m_Stream.ReadIntUnsigned()
        dstBlockLength = m_Stream.ReadLong()
        
        'Early PSP versions use an asinine mechanism for struct length descriptors, so this is
        ' a stupid (but necessary) workaround
        If (m_ImageHeader.psph_VersionMajor <= 3) Then dstBlockLength = m_Stream.ReadLong()
        
    End If
    
    'Reset the stream pointer before exiting
    m_Stream.SetPosition origPosition
    
End Function

'Want data on warnings?  Use these helper functions.
Friend Function Warnings_GetCount() As Long
    Warnings_GetCount = m_Warnings.GetNumOfStrings()
End Function

Friend Sub Warnings_CopyList(ByRef dstStack As pdStringStack)
    Set dstStack = m_Warnings
End Sub

Friend Sub Warnings_DumpToDebugger()
    If (m_Warnings.GetNumOfStrings() > 0) Then
        Dim i As Long
        For i = 0 To m_Warnings.GetNumOfStrings() - 1
            PDDebug.LogAction "(" & CStr(i + 1) & ") WARNING: pdPSP reported: " & m_Warnings.GetString(i)
        Next i
    End If
End Sub
