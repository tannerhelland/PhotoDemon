VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PSD (PhotoShop Image) Container and Parser
'Copyright 2019-2026 by Tanner Helland
'Created: 13/January/19
'Last updated: 31/August/22
'Last update: gracefully skip 0-length image resource blocks to fix crashes on Photopea-exported PSD files
'
'This class (and its associated pdPSD- child classes) handle Adobe PSD parsing duties.  It is custom-built
' for PhotoDemon, with an emphasis on performance and proper color-management of all imported data.
'
'To my knowledge, this is the only 3rd-party PSD parser that passes the full Apple color management
' test suite (https://developer.apple.com/library/archive/technotes/tn2115/_index.html).  This is a large
' point of pride for the author. ;)
'
'Currently, all color modes (nine in the spec), and all color depths (1 through 32-bits-per-channel)
' are believed to be fully covered by the import engine.  The export engine, by design, only writes
' 32-bpp RGBA images (as this is all PhotoDemon itself supports internally), but supporting other
' color modes or depths would require only minor modifications.
'
'As with all 3rd-party PSD engines, Photoshop has many features that don't have direct analogs in PhotoDemon.
' Such features are still extracted by this class, but they will not "appear" in the final loaded image.
' My ongoing goal is to expand support in this class as various PSD features are implemented in PD itself.
'
'Unless otherwise noted, all code in this class is my original work.  I've based my work off the
' "official" Adobe spec at this URL (link good as of January 2019):
' https://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577409_72092
'
'In places where Adobe's spec is either out-of-date or purely inaccurate, the open-source, MIT-licensed
' Paint.NET PSD plugin proved to be a valuable resource.  Its source code is available here
' (link good as of January 2019):
' https://github.com/PsdPlugin/PsdPlugin
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PSD files contain a *lot* of information.  To aid debugging, you can activate "verbose" output; this will
' dump all kinds of diagnostic information to the debug log.  (Note that other PSD-adjacent classes have
' their own version of this constant.)
Private Const PSD_DEBUG_VERBOSE As Boolean = False

'Until PD supports layer groups, we have to resort to workarounds for PSD layer group information.
' When loading layers, PD will automatically handle things like group opacity (by hiding all members
' of a layer group if said group is invisible), but as for the groups themselves, there are two options.
' The easiest option is to just ignore the groups and lose them when importing into PD.  The other
' option is to mimic Paint.NET's PSD plugin and import groups as "dummy" layers.  This gives the user
' a way to rearrange them (and see them, if nothing else), and we can use a specific naming scheme to
' ensure we preserve group order at export time.  Toggling these behaviors is controlled by this
' constant, with a matching bool to track whether an exported image contains layers.  (If it does,
' we need to modify some writing behavior.)
Private Const PSD_GROUPS_AS_DUMMY_LAYERS As Boolean = True
Private m_ExportGroups As Boolean

'Similarly, layer groups can contain additional features like masks.  If we allow round-tripping layer groups,
' we should also allow round-tripping layer group masks, but how to handle this in PD is tricky.  (If we just
' load the mask as a layer group's pixel data, the image doesn't display correctly.)  This debug will do
' just that - load group masks, if present, as layer data - so you can inspect the results, but I do not
' recommend using this in production code just yet.
Private Const PSD_LOAD_GROUP_MASKS_AS_PIXELS As Boolean = False

'PSD loading is complicated, and a lot of things can go wrong.  Instead of returning binary "success/fail"
' values, we return specific flags; "warnings" may be recoverable and you can still attempt to load the file.
' "Failure" returns are unrecoverable and processing *must* be abandoned.  (As a convenience, you can treat
' the "warning" and "failure" values as flags; specific warning/failure states in each category will share
' the same high flag bit.)
'
'As I get deeper into this class, I may expand this enum to include more detailed states.
Public Enum PD_PSDResult
    psd_Success = &H0
    psd_Warning = &H10
    psd_Failure = &H100
    psd_FileNotPSD = &H1000
End Enum

#If False Then
    Private Const psd_Success = &H0, psd_Warning = &H10, psd_Failure = &H100, psd_FileNotPSD = &H1000
#End If

'PSDs support many different color modes.  Some color modes (e.g. Duotone) are proprietary and not
' publicly documented; we follow Adobe's recommendation and simply load the data in an alternative format
' (e.g. grayscale in the case of Duotone).
Public Enum PSD_ColorMode
    cm_Bitmap = 0
    cm_Grayscale = 1
    cm_Indexed = 2
    cm_RGB = 3
    cm_CMYK = 4
    cm_Multichannel = 7
    cm_Duotone = 8
    cm_Lab = 9
End Enum

#If False Then
    Private Const cm_Bitmap = 0, cm_Grayscale = 1, cm_Indexed = 2, cm_RGB = 3, cm_CMYK = 4, cm_Multichannel = 7, cm_Duotone = 8, cm_Lab = 9
#End If

'The list of potential resource IDs is enormous:
' https://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577409_38034
' Most of these are irrelevant to PD, so for now, I've only listed ones that look potentially useful.
Private Enum PSD_ResourceID
    rid_ResolutionInfo = 1005
    rid_Caption = 1008
    rid_BackgroundColor = 1010  'Old PS versions required a background layer, if any, to be opaque
    rid_QuickMaskInfo = 1022    'Not sure if this needs to be used during compositing
    rid_LayerStateInfo = 1024   'Could be useful for determining auto-selected layer
    rid_LayersGroupInfo = 1026  'Could be useful for IDing and removing group-specific blank layers
    rid_IptcMetadata = 1028     'Compare against ExifTool's results?
    rid_JpegQuality = 1030      'We could auto-suggest this if saving as a JPEG from PD?
    rid_GridAndGuidesInfo = 1032   'Guides are on my to-do list
    rid_IccProfile = 1039       'Speaks for itself; this is critical to PD's color-managed import process
    rid_IccUntagged = 1041      'Spec is unclear on what this means; I'd need to see it "in the wild" to know how to handle in PD
    rid_IndexedColorTableCount = 1046   'Number of used colors - important for indexed images!
    rid_TransparentIndex = 1047 'Need to test against an indexed image PSD
    rid_AlphaIdentifiers = 1053 'No idea what this does, but we should check it against alpha channels
    rid_VersionInfo = 1057      'Lets us know if PS was the actual image source; will need to write when saving
    rid_EXIFData1 = 1058        'Check integration w/ ExifTool
    rid_EXIFData3 = 1059        'Check integration w/ ExifTool
    rid_XMPMetadata = 1060      'Check integration w/ ExifTool
    rid_PixelAspectRatio = 1064 'Could theoretically correct for this
    rid_LayerSelectionIDs = 1069    'Possibly useful if understood
    rid_LayerGroupsEnabledID = 1072 'Probably need this to decode layer group behavior
    rid_MeasurementScale = 1074     'Could potentially auto-adjust rulers against this? idk
    rid_DisplayInfo = 1077  '"DisplayInfo structure to support floating point colors. See Appendix A in Photoshop API Guide.pdf"
    rid_PathInfo = 2000 'Actually comprises 2000-2997; could be useful in the future
    rid_ClippingPathName = 2999 'We could potentially support this, if found
End Enum

#If False Then
    Private Const rid_ResolutionInfo = 1005, rid_Caption = 1008, rid_BackgroundColor = 1010, rid_QuickMaskInfo = 1022, rid_LayerStateInfo = 1024, rid_LayersGroupInfo = 1026, rid_IptcMetadata = 1028, rid_JpegQuality = 1030, rid_GridAndGuidesInfo = 1032, rid_IccProfile = 1039, rid_IccUntagged = 1041, rid_IndexedColorTableCount = 1046, rid_TransparentIndex = 1047, rid_AlphaIdentifiers = 1053, rid_VersionInfo = 1057, rid_EXIFData1 = 1058, rid_EXIFData3 = 1059, rid_XMPMetadata = 1060, rid_PixelAspectRatio = 1064, rid_LayerSelectionIDs = 1069, rid_LayerGroupsEnabledID = 1072, rid_MeasurementScale = 1074, rid_DisplayInfo = 1077, rid_PathInfo = 2000, rid_ClippingPathName = 2999
#End If

'Standard PSD header.  At load-time, only populated if the target PSD validates.
' (Note that this struct has major alignment issues under VB if attempting to write the whole thing
' as a block; PD writes members individually to avoid these issues.
Private Type PSD_Header
    Signature As String * 4
    Version As Integer
    Reserved(0 To 5) As Byte
    NumChannels As Integer
    ImageHeightPx As Long
    ImageWidthPx As Long
    BitsPerChannel As Integer
    ColorMode As PSD_ColorMode
End Type

Private m_Header As PSD_Header

'If the header defines a PSB file instead of a PSD file, this value will be set to TRUE early in the parse process.
' Subsequent functions should use this value to modify behavior according to PSD/PSB differences.
Private m_PSDisPSB As Boolean

'PhotoShop uses a detailed struct for resolution info; the spec tells you to dive into the separate
' SDK guide for this info, which is obtainable from https://www.adobe.com/devnet/photoshop/sdk.html
Private Enum PSD_ResolutionUnit
    ruInches = 1
    ruCentimeters = 2
    ruPoints = 3
    ruPicas = 4
    ruColumns = 5
End Enum

#If False Then
    Private Const ruInches = 1, ruCentimeters = 2, ruPoints = 3, ruPicas = 4, ruColumns = 5
#End If

'Per the spec:
' 1=display [sic] horitzontal resolution in pixels per inch; 2=display [sic] horitzontal resolution in pixels per cm.
Private Enum PSD_ResolutionUnitDisplay
    rudPPI = 0
    rudPPCm = 1
End Enum

#If False Then
    Private Const rudPPI = 0, rudPPCm = 1
#End If

Private Type PSD_ResolutionInfo
    riHRes As Single                        'Fixed32 in the file
    riHResUnit As PSD_ResolutionUnitDisplay 'Integer in the file
    riWidthUnit As PSD_ResolutionUnit       'Integer in the file
    riVRes As Single                        'Fixed32 in the file
    riVResUnit As PSD_ResolutionUnitDisplay 'Integer in the file
    riHeightUnit As PSD_ResolutionUnit      'Integer in the file
End Type

Private m_ResolutionInfo As PSD_ResolutionInfo

'If the PSD file uses indexed color mode, its palette will be initialized into this array
Private m_ColorTable() As RGBQuad, m_ColorTableCount As Long, m_TransparentIndex As Long

'Any stored image resources are retrieved as-is.  Some of these may be parsed further
' and stored in specialized (more usable) types.
Private Type PSD_ImageResource
    'irSignature As String * 4  'Always '8BIM', so we don't waste bytes storing it
    irID As PSD_ResourceID      '2-byte hex ID
    irName As String            'Pascal string, always "null" IRL as far as I can tell
    irDataLength As Long        'Size of data array
    irDataBytes() As Byte       '"Padded to make the size even"; PD will *not* retrieve padding bytes if they occur
End Type

Private m_ImageResources() As PSD_ImageResource
Private m_NumImageResources As Long

'If the PSD file contains layers (null-layer images *are* technically possible), each layer's contents
' will be accessible from this array.  If m_NumOfLayers is non-zero, this array is guaranteed to be sized
' as (0 to Abs(m_NumOfLayers) - 1).
Private m_Layers() As pdPSDLayer

'Per the spec, the number of layers *CAN BE NEGATIVE*.  A negative count has special meaning; see Step 4 for details.
Private m_numOfLayers As Long

'Embedded ICC profiles are read and used by PD.  (These can be found in the "image resource" segment.)
Private m_Profile As pdICCProfile

'During loading, we mark the offset to the composite image data.  PD doesn't always need this data,
' but if layer data is "unsalvageable" (e.g. layers are vector or fill layers that PD doesn't understand),
' we can use the composite image to at least give the user *something*.
Private m_CompositeImageOffset As Long

'High bit-depth images store data in a totally different way that requires us to jump around the source file.
' We flag this state so that we can reset the underlying stream pointer accordingly.
Private m_HighBitDepthToggle As Boolean

'Byte-by-byte access is provided, as always, by a pdStream instance
Private m_Stream As pdStream

'At present, we require the caller to pass an identical source file path to every load function.
' (This is a cheap and easy way to ensure no funny business.)  If the PSD is loaded directly from memory,
' we flag this with a special name.
Private m_SourceFilename As String
Private Const PSD_LOADED_FROM_MEMORY As String = "LoadFromPtr*"
Private m_SourcePtr As Long, m_SourcePtrLen As Long

'If warnings are encountered during processing, I push their messages onto a string stack.  (I may
' decide to report these to the user... but haven't decided yet.  Either way, it's very helpful
' while debugging; see associated Warnings_XYZ functions for details.)
Private m_Warnings As pdStringStack

'Perform basic validation on a potential PSD file.  For strict file extension matching (e.g. only
' PSD or PSB extension allowed), set checkExtension to TRUE.
Friend Function IsFilePSD(ByRef srcFile As String, Optional ByVal checkExtension As Boolean = False) As Boolean
    IsFilePSD = (Step1_ValidateHeader(srcFile, checkExtension) < psd_Failure)
End Function

'Simplified wrapper to load a PSD automatically.
Friend Function LoadPSD(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal checkExtension As Boolean = False, Optional ByVal loadFromPtr As Long = 0, Optional ByVal loadFromPtrLen As Long = 0) As PD_PSDResult
    
    'Reset some internal parameters to ensure subsequent reads are accurate.  (This is critical if multiple PSDs
    ' are read back-to-back.)
    Me.Reset
    
    'We support PSD loading from both file and memory; if the passed loadFromPtr value is non-zero,
    ' treat it as a pointer and we'll wrap our stream around it instead.
    If (loadFromPtr <> 0) And (loadFromPtrLen <> 0) Then
        srcFile = PSD_LOADED_FROM_MEMORY
        m_SourcePtr = loadFromPtr
        m_SourcePtrLen = loadFromPtrLen
    Else
        m_SourcePtr = 0
        m_SourcePtrLen = 0
    End If
    
    'Try to validate the source file
    Dim keepLoading As PD_PSDResult
    keepLoading = Step1_ValidateHeader(srcFile, checkExtension)
    
    If (keepLoading < psd_Failure) Then
        
        'The file validated well enough to continue (e.g. any warnings were considered non-fatal).
        ' Further processing is handled on a block-by-block basis, and a critical error at any step
        ' causes the entire file to be abandoned.
        PDDebug.LogAction "PSD file detected. Size is " & m_Header.ImageWidthPx & "x" & m_Header.ImageHeightPx & ", color mode is " & GetColorModeName(m_Header.ColorMode) & ", " & m_Header.BitsPerChannel & " bits per channel."
        keepLoading = Step2_RetrieveColorModeData(srcFile)
        If PSD_DEBUG_VERBOSE Then If (keepLoading < psd_Failure) Then PDDebug.LogAction "PSD parsing step 2 successful." Else PDDebug.LogAction "PSD parsing step 2 unsuccessful."
        
        'The color table, if any, has been retrieved.  The next segment of the file is a massive chunk
        ' called "Image Resources".  Much of the data in this section is specific to Photoshop, but we will
        ' attempt to retrieve what we can from it.
        If (keepLoading < psd_Failure) Then
            keepLoading = Step3_GatherImageResources(srcFile, dstImage)
            If PSD_DEBUG_VERBOSE Then If (keepLoading < psd_Failure) Then PDDebug.LogAction "PSD parsing step 3 successful." Else PDDebug.LogAction "PSD parsing step 3 unsuccessful."
        End If
        
        'Image resources have been retrieved.  The next segment of the file is layer data!
        If (keepLoading < psd_Failure) Then
            keepLoading = Step4_GatherLayersAndMasks(srcFile)
            If PSD_DEBUG_VERBOSE Then If (keepLoading < psd_Failure) Then PDDebug.LogAction "PSD parsing step 4 successful." Else PDDebug.LogAction "PSD parsing step 4 unsuccessful."
        End If
        
        'The final segment of a PSD file is a full copy of the composited layered image.  When Photoshop writes
        ' a PSD file, the contents of this segment are tied to a "maximize compatibility" setting. If "maximize
        ' compatibility" is checked (default = true), a full copy of the composited image is written out to the
        ' final block of the file.  If "maximize compatibility" is *not* set, PhotoShop writes out a plain white
        ' image to that final segment - a strange decision, but pure white compresses well and this guarantees
        ' that "lazy" readers still have something to read.
        
        'Anyway, we don't need this data in PD so we ignore it at present.  You could, however, parse it as the
        ' final stage in the load process, if desired - but note that there are caveats to the data in this segment.
        ' (For example, Photoshop may choose to write some single-layer images as zero-layer images, with the image
        ' data actually stored in the composited image segment.  PhotoDemon supports this condition, so the current
        ' stream point may not point where you think it does - consider yourself warned!)
        
        'If we made it this far successfully, we have enough data to construct a pdImage object!
        If (keepLoading < psd_Failure) Then
            keepLoading = Step5_AssemblePDImage(srcFile, dstImage)
            If PSD_DEBUG_VERBOSE Then If (keepLoading < psd_Failure) Then PDDebug.LogAction "PSD parsing step 5 successful." Else PDDebug.LogAction "PSD parsing step 5 unsuccessful."
        End If
        
        LoadPSD = keepLoading
        
    End If
    
    'Regardless of success/failure, ensure the underlying stream gets closed
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen() Then m_Stream.StopStream True
    End If
    
    'Even if the PSD loaded, one (or more) surprises may have affected our ability to render everything correctly.
    ' Regardless of this class's debug verbosity setting, warnings always get dumped out to the debug log.  This is
    ' especially useful when new versions of Photoshop release, as they often bring new parsing surprises/nightmares.
    If (Me.Warnings_GetCount() > 0) Then
        PDDebug.LogAction "PhotoDemon's PSD parser generated one or more warnings.  Here is a full list:"
        Me.Warnings_DumpToDebugger
    End If

End Function

'Save a new PSD file to disk.  (With minor modifications, this could also be used to save to a memory stream,
' but this is *not* currently implemented.)
'
'If the destination file exists, it will be forcibly overwritten.  Safe save measures should be implemented externally.
Friend Function SavePSD(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal useMaxCompatibility As Boolean = True, Optional ByVal compressionType As Long = 0, Optional ByVal writePSB As Boolean = False) As Boolean

    On Error GoTo CouldNotSaveFile
    
    If (srcPDImage Is Nothing) Or (LenB(dstFile) = 0) Then
        InternalError "SavePSD", "invalid function parameters"
        SavePSD = False
        Exit Function
    End If
    
    'Failsafe check only; PD enforces safe overwriting in the parent function.
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'So, a few quick notes before we begin.
    
    '1) Saving a PSD file is much easier than loading it, primarily because we don't need to cover every
    ' color-mode and bit-depth option.  Instead, we focus on exporting 32-bit RGBA data in the cleanest
    ' manner possible.  Other bit-depths and/or color modes may be covered in future updates.
    
    '2) PSDs use a ton of structs with non-DWORD alignments.  This causes issues in VB6 as struct members
    ' are typically aligned to DWORD boundaries (with some caveats; it's complicated and basically don't
    ' worry about it if you aren't a VB6 developer).  To avoid complications or convoluted struct-shifting,
    ' this function manually writes out most struct members individually.
    
    '3) This code was designed against the latest PSD spec at the time of its writing (August 2016,
    ' available here: https://www.adobe.com/devnet-apps/photoshop/fileformatashtml/).  No guarantee is
    ' made that this approach will be compatible with earlier or later spec versions.
    
    '4) If padding issues with strings (or any other data type) seem messy, that's because they are.
    ' Adobe uses many different string alignment/padding practices throughout PSDs, and only by a careful
    ' reading of the spec can you see what is required where.  Any messiness is their doing - not mine.
    
    '5) All PSD data is big-endian, so you will see a ton of _BE suffixes in stream function calls.
    
    'All good?  Then let's do this!
    Dim i As Long
    
    'Before exporting, we need to run some quick heuristics on exported layer data.  This approach will
    ' likely change as PD is able to reproduce more Photoshop behavior, but for now, we sometimes need
    ' to perform weird conversions between PD data and PSD data.
    
    'Scan layer names in the source image for group identifiers.  If we find any, set a module-level
    ' flag to note that layer groups need to be exported.  (They require additional header information
    ' and optional chunks.)  Note that group export behavior can be controlled by a module-level constant.
    m_ExportGroups = False
    If PSD_GROUPS_AS_DUMMY_LAYERS Then
        
        'When importing PSD files, layer groups are stored as "dummy" 1x1 layers with special names.
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
        '    values means the resulting PSD would be malformed, and we *never* want to risk that.)
        If (startTargetsFound > 0) And (endTargetsFound > 0) Then
            m_ExportGroups = (startTargetsFound = endTargetsFound)
        End If
        
    End If
    
    'As usual, all writing is handled by a pdStream instance.  Note that we do not request typical
    ' sequential perf-optimized mode - this is because PSDs use a ton of "length" entries that describe
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
    
    'Start by writing out a basic PSD header.
    Dim keepGoing As PD_PSDResult
    keepGoing = ExportStep1_WriteHeaderAndColorTable(cStream, srcPDImage, useMaxCompatibility, writePSB)
    
    'Next comes the image resource segment.  This is a variable-length segment comprised of individual "blocks".
    ' Blocks are similar to chunks in other formats; they are self-contained descriptors of variable length
    ' and type.  We write much fewer of these than Photoshop typically does, obviously.
    If (keepGoing < psd_Failure) Then keepGoing = ExportStep2_WriteImageResources(cStream, srcPDImage, useMaxCompatibility, writePSB)
    
    'Next comes the (potentially very large) layer and mask information section.  This is the most time-consuming
    ' segment to write, on account of writing all layer pixel data.
    If (keepGoing < psd_Failure) Then keepGoing = ExportStep3_WriteLayerAndMaskInformation(cStream, srcPDImage, compressionType, writePSB)
    
    'Finally, if "maximum compatibility" is selected, write out a copy of the merged image data.
    If (keepGoing < psd_Failure) Then keepGoing = ExportStep4_WriteMergedImage(cStream, srcPDImage, compressionType, useMaxCompatibility, writePSB)
    
    'The write is finished.  Close the file handle and exit.
    cStream.StopStream
    Set cStream = Nothing
    
    SavePSD = (keepGoing < psd_Failure)
    
    Exit Function
    
CouldNotSaveFile:
    InternalError "SavePSD", "Internal VB error #" & Err.Number & ": " & Err.Description
    SavePSD = False
    
End Function

Private Function ExportStep4_WriteMergedImage(ByRef cStream As pdStream, ByRef srcImage As pdImage, Optional ByVal compressionType As Long = 0, Optional ByVal useMaxCompatibility As Boolean = True, Optional ByVal writePSB As Boolean = False) As PD_PSDResult

    'Start by retrieving a merged image copy
    Dim tmpDIB As pdDIB
    If useMaxCompatibility Then
        srcImage.GetCompositedImage tmpDIB, False
    
    'When "max compatibility" is *not* set in Photoshop, Adobe just embeds a blank white image at the size of the
    ' original image.  This compresses much better and allows 3rd-party decoders to still retrieve *something*
    ' without crashing, and if a program can read layer data anyway, the composite is just a waste of space.
    ' (See also: https://feedback.photoshop.com/photoshop_family/topics/maximize_file_compatibility_bloats_files_need_smaller_option)
    ' We mimic Adobe's approach.
    Else
        Set tmpDIB = New pdDIB
        tmpDIB.CreateBlank srcImage.Width, srcImage.Height, 32, vbWhite
    End If
    
    'Determine if the composite image has meaningful alpha values (e.g. if it is *not* fully opaque)
    Dim alphaMatters As Boolean
    alphaMatters = DIBs.IsDIBTransparent(tmpDIB)
    
    'Write the compression method first.  Note that only raw and PackBits are currently supported in the
    ' merged image data segment; PD *could* technically write zipped data, but the spec is unclear on how
    ' the zip stream is handled in this segment - and it wouldn't be compatible with most applications
    ' anyway which defeats the whole point of the "maximize compatibility" setting!
    Const PSD_COMPRESSION_PACKBITS As Long = 1
    If (compressionType > PSD_COMPRESSION_PACKBITS) Then compressionType = PSD_COMPRESSION_PACKBITS
    cStream.WriteInt_BE compressionType
    
    'RLE requires special handling
    If (compressionType = PSD_COMPRESSION_PACKBITS) Then
    
        'Generating RLE data is convoluted and obnoxious, so let's just reuse a standard layer object to do it!
        Dim tmpPSDLayer As pdPSDLayer
        Set tmpPSDLayer = New pdPSDLayer
        tmpPSDLayer.PrepChannelsForWrite tmpDIB, compressionType
        
        'The PSD layer object can also write out the channel data for us - how nice!
        tmpPSDLayer.WriteSpecialMergedChannelData cStream, compressionType, writePSB, alphaMatters
    
    'The only other supported compression type at present is raw bytes (no compression).
    Else
    
        'We simply need to write all data to file in planar format (e.g. RRRGGGBBBAAA instead of RGBARGBARGBA...)
        Dim x As Long, y As Long
        Dim lWidth As Long, lHeight As Long, xLoopSrc As Long, yLoopSrc As Long
        Dim dstOffset As Long
        
        lWidth = tmpDIB.GetDIBWidth
        lHeight = tmpDIB.GetDIBHeight
        xLoopSrc = lWidth - 1
        yLoopSrc = lHeight - 1
        
        'Pre-cache each plane individually before writing, to improve performance.
        Dim chSize As Long
        chSize = lWidth * lHeight
        
        Dim planarBytes() As Byte
        If alphaMatters Then ReDim planarBytes(0 To chSize * 4 - 1) As Byte Else ReDim planarBytes(0 To chSize * 3 - 1) As Byte
        
        Dim tmpBytes() As Byte, tmpSA As SafeArray1D
        Dim scanStart As Long, scanWidth As Long
        tmpDIB.WrapArrayAroundScanline tmpBytes, tmpSA, 0
        scanStart = tmpSA.pvData
        scanWidth = tmpDIB.GetDIBStride
            
        For y = 0 To yLoopSrc
            tmpSA.pvData = scanStart + scanWidth * y
        For x = 0 To xLoopSrc
            
            dstOffset = y * lWidth + x
            
            'Destination channels are in RGBA order, not BGRA order
            planarBytes(dstOffset) = tmpBytes(x * 4 + 2)
            planarBytes(chSize + dstOffset) = tmpBytes(x * 4 + 1)
            planarBytes(chSize * 2 + dstOffset) = tmpBytes(x * 4)
            If alphaMatters Then planarBytes(chSize * 3 + dstOffset) = tmpBytes(x * 4 + 3)
            
        Next x
        Next y
        
        tmpDIB.UnwrapArrayFromDIB tmpBytes
        
        'Write the planar data to file
        cStream.WriteByteArray planarBytes
        
    End If
    
    'If alpha does *not* matter, we want to do something kind of hacky - we want to return to very near
    ' the start of the file, and embed a "3" instead of "4" for the number of channels in the image.
    If (Not alphaMatters) Then
        
        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "dropping alpha channel to conserve space"
        
        Dim curStreamPosition As Long
        curStreamPosition = cStream.GetPosition()
        
        'This is effectively a magic number: 4 bytes ASCII signature + 2 bytes version no. + 6 reserved bytes = 12
        cStream.SetPosition 12, FILE_BEGIN
        cStream.WriteInt_BE 3
        cStream.SetPosition curStreamPosition, FILE_BEGIN
        
    End If
        

End Function

Private Function ExportStep3_WriteLayerAndMaskInformation(ByRef cStream As pdStream, ByRef srcImage As pdImage, Optional ByVal compressionType As Long = 0, Optional ByVal writePSB As Boolean = False) As PD_PSDResult

    'As with other segments in a PSD file, the first marker in this segment is a "net length" marker.
    ' We can't possibly predict this value in advance, so just write a dummy value and we'll overwrite
    ' it at the end of this function with a correct value.
    Dim startPosition As Long
    startPosition = cStream.GetPosition()
    cStream.WriteLong_BE 0
    
    'The layer and mask info section is broken into several sub-sections:
    ' 1) Layer info
    ' 2) Global mask info
    ' 3) Misc tagged blocks
    
    'Each of these sections contains multiple sub-sections of their own.
    ExportStep3_WriteLayerAndMaskInformation = ExportStep3a_WriteLayerInfo(cStream, srcImage, compressionType, writePSB)
    
    'With all layer and channel data written, the last segment we need to write is global layer mask info.
    ' This is always 0.
    cStream.WriteLong_BE 0
    
    'No misc blocks are written in this global section (at present)
    
    'All writes are complete, which means we can (finally) calculate the size of this segment.
    Dim finalPosition As Long, segmentSize As Long
    finalPosition = cStream.GetPosition()
    segmentSize = (finalPosition - startPosition) - 4
    
    'Return to the initial length parameter and write it out correctly.
    cStream.SetPosition startPosition, FILE_BEGIN
    cStream.WriteLong_BE segmentSize    'Do not include the length value itself in the size calculation
    cStream.SetPosition finalPosition, FILE_BEGIN
    
End Function

Private Function ExportStep3a_WriteLayerInfo(ByRef cStream As pdStream, ByRef srcImage As pdImage, Optional ByVal compressionType As Long = 0, Optional ByVal writePSB As Boolean = False) As PD_PSDResult

    'As always, we start with a (currently unknown) length marker.  Unlike other segments, this one
    ' should actually be rounded up to a multiple of 2.
    Dim startPosition As Long
    startPosition = cStream.GetPosition()
    cStream.WriteLong_BE 0
    
    'Next comes layer count.  When reading PSD files from Photoshop, it may use a negative layer count to describe
    ' special alpha packing (per the spec, "If [layer count] is a negative number, its absolute value is the number
    ' of layers and the first alpha channel contains the transparency data for the merged result.")  We don't need
    ' to write this case, so we only ever write the *actual* layer count.
    cStream.WriteInt_BE srcImage.GetNumOfLayers()
    
    'If we're gonna try and export layer groups, we need to create some group flags in advance
    Dim targetStartName As String, targetEndName As String
    If m_ExportGroups Then
        targetStartName = g_Language.TranslateMessage("Group start:")
        targetEndName = g_Language.TranslateMessage("Group end:")
    End If
    
    'Next comes a variable-length segment called "layer records".  One of these records exists for each layer.
    ' Because all PSD layer-specific data is managed by a dedicated class, it's easier for us to offload
    ' layer embedding tasks to that same class.  Note that we want the classes to be persistent for the duration
    ' of this function because each layer header must define the length of the layer's pixel data; it's easiest
    ' to calculate both of those simultaneously, then keep the channel data in-memory for subsequent embedding
    ' after all layer segments are written.
    Dim tmpPSDLayers() As pdPSDLayer
    ReDim tmpPSDLayers(0 To srcImage.GetNumOfLayers - 1) As pdPSDLayer
    
    Dim layerGroupTypes() As PSD_LayerGroupState
    ReDim layerGroupTypes(0 To srcImage.GetNumOfLayers - 1) As PSD_LayerGroupState
    
    Dim i As Long
    For i = 0 To srcImage.GetNumOfLayers() - 1
        
        Set tmpPSDLayers(i) = New pdPSDLayer
        layerGroupTypes(i) = lgs_None
        
        'Are we allowed to write layer groups?  If we are, see if this is a group start/end marker
        If m_ExportGroups Then
            If Strings.StringsEqual(Left$(srcImage.GetLayerByIndex(i).GetLayerName(), Len(targetStartName)), targetStartName, True) Then layerGroupTypes(i) = lgs_GroupStartOpen
            If Strings.StringsEqual(Left$(srcImage.GetLayerByIndex(i).GetLayerName(), Len(targetEndName)), targetEndName, True) Then layerGroupTypes(i) = lgs_GroupEnd
        End If
        
        tmpPSDLayers(i).WriteLayerData cStream, srcImage.GetLayerByIndex(i), compressionType, writePSB, layerGroupTypes(i)
        
    Next i
    
    'During the previous step, each layer generated a compressed byte stream of each channel's pixel data
    ' (remember: PSDs store channels as planar, not interleaved).  That makes the next step - storing channel
    ' data - trivial, as we can just ask each layer object to dump channel data out to file.
    For i = 0 To srcImage.GetNumOfLayers() - 1
        
        'Non-group-layers are always written
        If (layerGroupTypes(i) = lgs_None) Then
            tmpPSDLayers(i).WriteChannelData cStream, compressionType, writePSB
        
        'Group layers just write dummy values (2x4=8) - UNLESS a mask is present (then they write more)
        Else
            tmpPSDLayers(i).WriteChannelData_Group srcImage.GetLayerByIndex(i), cStream, compressionType, writePSB
        End If
        
    Next i
    
    'All writes are complete, which means we can (finally) calculate the size of this segment.
    Dim finalPosition As Long, segmentSize As Long
    finalPosition = cStream.GetPosition()
    segmentSize = (finalPosition - startPosition) - 4
    
    'The spec says that the segment size should be a multiple of 2.  Enforce this now.
    If (segmentSize And 1) = 1 Then
        finalPosition = finalPosition + 1
        segmentSize = segmentSize + 1
        cStream.WriteByte 0
    End If
    
    'Return to the initial length parameter and write it out correctly.
    cStream.SetPosition startPosition, FILE_BEGIN
    cStream.WriteLong_BE segmentSize    'Do not include the length value itself in the size calculation
    cStream.SetPosition finalPosition, FILE_BEGIN
    
    ExportStep3a_WriteLayerInfo = psd_Success
    
End Function

Private Function ExportStep2_WriteImageResources(ByRef cStream As pdStream, ByRef srcImage As pdImage, Optional ByVal useMaxCompatibility As Boolean = True, Optional ByVal writePSB As Boolean = False) As PD_PSDResult
    
    'Before doing anything else, we need to make a note of the current stream position.  The first entry
    ' in this segment is the TOTAL LENGTH of the image resource segment.  We can't predict this value in
    ' advance (not easily, anyway), so we just write a dummy value for now, then return and fill in the
    ' "correct" value when the section is finished.
    Dim startPosition As Long
    startPosition = cStream.GetPosition()
    cStream.WriteLong_BE 0
    
    'Image resource blocks all use a fixed structure:
    ' - 4 byte signature '8BIM'
    ' - 2 byte unique ID (hard-coded and defined by Adobe)
    ' - Variable-length Pascal string w/ resource name (in practice, this always seems to be a two-byte null-string)
    ' - 4 byte size of resource data
    ' - actual resource data (even-padded according to the spec, but this does *not* appear to be practiced IRL)
    
    'A cleaner implementation would use dedicated sub-classes for each possible resource block type.  We don't do
    ' this because VB6 makes it impossible to cleanly organize large class collections.  (Also, we don't actually
    ' write many image resource blocks, so this function is manageable for now.)  That's why you'll see a bunch
    ' of ugly naked stream writes here.
    
    'PSD images from Photoshop typically contain dozens of image resource blocks.  PhotoDemon doesn't attempt to
    ' mimic this behavior; instead, we only write blocks that contain information relevant to the current PD session.
    ' (Photoshop will fill in default values for any blocks we don't supply.)
    
    'Here are the blocks we currently write.
    
    'First, a resolution info block (0x03ED).  The PSD spec does not define this; instead, you have to track down
    ' Appendix A of the Photoshop API guide.  (Fortunately, it's an easy segment to write, aside from the hassle
    ' of using fixed-point fractional values.)
    Dim blockSize As Long
    blockSize = 16
    WriteImageResourceHeader cStream, &H3ED, blockSize
    
    'Horizontal resolution (as fixed 16-16), resolution unit (1=ppi), and width unit (1=inches)
    cStream.WriteFixed1616_BE srcImage.GetDPI()
    cStream.WriteInt_BE 1
    cStream.WriteInt_BE 1
    
    'Vertical resolution is exactly the same
    cStream.WriteFixed1616_BE srcImage.GetDPI()
    cStream.WriteInt_BE 1
    cStream.WriteInt_BE 1
    
    'Second, an ICC profile (0x040F).  At present, all images exported from PhotoDemon are hard-converted to sRGB.
    ' PSD is the rare format where an embedded profile would actually (most likely?) be used by target software,
    ' instead of being completely ignored as usual... but some internal changes would have to be made to make
    ' that feasible.  Instead, we just generate a stock sRGB profile on-the-fly and embed it.
    '
    '(I've gone back and forth on the compatibility complications of writing this profile in modern v4 format.
    ' v4 profiles date back to 2001, but some people use *really* old software, so for now, PD forcibly generates
    ' a larger-but-more-compatible v2.1 sRGB profile for embedding whenever "maximum compatibility" is checked
    ' (as it is by default).  Note that the contents of v2 vs v4 sRGB profiles are identical in how the image
    ' is ultimately rendered, but the v4 profile is much smaller on account of v4 descriptors that represent
    ' sRGB's weird gamma ramp in a compact way.)
    Dim tmpProfile As pdLCMSProfile
    Set tmpProfile = New pdLCMSProfile
    tmpProfile.CreateSRGBProfile (Not useMaxCompatibility)
    
    Dim profBytes() As Byte
    blockSize = tmpProfile.GetRawProfileBytes(profBytes)
    WriteImageResourceHeader cStream, &H40F, blockSize
    cStream.WriteByteArray profBytes, blockSize
    
    'Image resource blocks must report their sizes accurately (no padding).  If an image resource block size
    ' is ODD, however, an extra padding byte needs to be inserted after the block, so that the *next* block
    ' lies on an even-position boundary.  (Again, this padding is *not* reported in the resource block size.)
    If ((blockSize And 1) = 1) Then cStream.WriteByte 0
    
    'Third (and last), version information (0x0421).  Version information describes the program that wrote
    ' the PSD file.  The spec doesn't mark this block as required (in fact, *no* resource blocks are required
    ' in a PSD file), but as with other formats, we try to follow Best Practices and take responsibility for
    ' the files we write.
    
    'Size of the version number is calculated as follows:
    ' 4 bytes - version number (always 1)
    ' 1 byte - real merged data included in file (e.g. "max compatibility" used at export time)
    ' (variable bytes) - name of PSD writer as a "Unicode string" (4 byte len + 2 bytes * num chars + 2-byte trailing null)
    ' (variable bytes) - name of PSD reader as a "Unicode string" (4 byte len + 2 bytes * num chars + 2-byte trailing null)
    ' 4 bytes - file version (always 1)
    blockSize = 4 + 1 + 4 'Fixed-size entries
    blockSize = blockSize + 4 + 4 'Size descriptors for two Unicode strings
    
    Dim writerName As String, readerName As String
    writerName = "PhotoDemon"
    readerName = writerName & " " & Updates.GetPhotoDemonVersion()
    blockSize = blockSize + LenB(writerName) + 2 + LenB(readerName) + 2  'Add string sizes to calculation, including trailing nulls
    
    WriteImageResourceHeader cStream, &H421, blockSize
    cStream.WriteLong_BE 1
    If useMaxCompatibility Then cStream.WriteByte 1 Else cStream.WriteByte 0
    cStream.WriteLong_BE Len(writerName)
    cStream.WriteString_UnicodeBE writerName, True
    cStream.WriteLong_BE Len(readerName)
    cStream.WriteString_UnicodeBE readerName, True
    cStream.WriteLong_BE 1
    
    'See above note about image resource block padding
    If ((blockSize And 1) = 1) Then cStream.WriteByte 0
    
    'If layer groups will be present in the destination file, we need to write some group indicator blocks now
    If m_ExportGroups Then
    
        'Two blocks are always present if groups are active:
        ' 1) "Layer groups enabled", and
        ' 2) "Layer groups info"
        
        'These probably have some meaning to Photoshop, but every document I've loaded just uses a string
        ' of constant numbers for both.  We mimic that behavior here.
        
        'Start with group info
        blockSize = srcImage.GetNumOfLayers * 2     'Unsigned ints
        WriteImageResourceHeader cStream, &H402, blockSize
        
        Dim i As Long
        For i = 0 To srcImage.GetNumOfLayers - 1
            cStream.WriteIntU_BE 0
        Next i
        
        'Next, groups enabled
        blockSize = srcImage.GetNumOfLayers
        WriteImageResourceHeader cStream, &H430, blockSize
        For i = 0 To srcImage.GetNumOfLayers - 1
            cStream.WriteByte 1
        Next i
        
        'See above note about image resource block padding
        If ((blockSize And 1) = 1) Then cStream.WriteByte 0
        
    End If
    
    'All image resource blocks have been written successfully!
    
    'Remember the beginning of this function?  We now need to return to the start of this segment and write out
    ' a *correct* "length of segment" marker.
    Dim finalPosition As Long, fullSizeOfSegment As Long
    finalPosition = cStream.GetPosition()
    fullSizeOfSegment = (finalPosition - startPosition) - 4  'Subtract 4 because the length marker itself is not included in the size calculation
    
    cStream.SetPosition startPosition, FILE_BEGIN
    cStream.WriteLong_BE fullSizeOfSegment
    cStream.SetPosition finalPosition, FILE_BEGIN
    
    ExportStep2_WriteImageResources = psd_Success
    
End Function

Private Sub WriteImageResourceHeader(ByRef cStream As pdStream, ByVal resID As Integer, ByVal resDataSize As Long)
    
    'Always start with the hard-coded signature
    Const IMAGE_RESOURCE_SIGNATURE As String = "8BIM"
    cStream.WriteString_ASCII IMAGE_RESOURCE_SIGNATURE
    
    'Next, the resource ID
    cStream.WriteIntU_BE resID
    
    'Next, a null Pascal string as the "resource name"
    cStream.WriteByte 0 '0 length
    cStream.WriteByte 0 'Extra byte for padding
    
    'Next, size of the resource data
    cStream.WriteLong_BE resDataSize
    
    'Writing of the actual data is left to the caller
    
End Sub

Private Function ExportStep1_WriteHeaderAndColorTable(ByRef cStream As pdStream, ByRef srcImage As pdImage, Optional ByVal useMaxCompatibility As Boolean = True, Optional ByVal writePSB As Boolean = False) As PD_PSDResult
    
    'As noted previously, VB6 struct alignment issues prevent us from easily writing out structs "as-is";
    ' instead, we write out members one-at-a-time.
    
    'Magic number
    cStream.WriteString_ASCII "8BPS"
    
    'Version: 1 for PSD, 2 for PSB
    cStream.WriteInt_BE 1
    
    '6 reserved bytes; must be 0
    cStream.WritePadding 6
    
    'Number of channels in the image, including alpha.  This is a little convoluted to write,
    ' because we only want to embed alpha data in the base layer if it's actually relevant
    ' (e.g. if the merged image is *not* fully opaque).  Because generating an extra merged copy
    ' of the full image is expensive, we only want to do it once, at the end of this function -
    ' so for now, just assume 4-channel RGBA and we will revisit this value later if we determine
    ' that the image's alpha data is pointless.
    cStream.WriteInt_BE 4
    
    'Height and width of the image, in pixels
    cStream.WriteLong_BE srcImage.Height
    cStream.WriteLong_BE srcImage.Width
    
    'Number of bits per channel; PD always writes 8bx4 channels at present
    cStream.WriteInt_BE 8
    
    'Color mode; PD always writes RGB images at present.
    cStream.WriteInt_BE cm_RGB
    
    'While here, let's also write the color mode data segment.  In indexed images, this segment stores palette data.
    ' It is unused by PhotoDemon exports so we only need to populate the first field - segment length, which is 0.
    cStream.WriteLong_BE 0
    
    ExportStep1_WriteHeaderAndColorTable = psd_Success

End Function

'Only valid after LoadPSD has been called
Friend Sub GetColorTableData(ByRef dstColors() As RGBQuad, ByRef dstColorCount As Long, ByRef dstTransparentIndex As Long)
    ReDim dstColors(0 To 255) As RGBQuad
    If (m_ColorTableCount > 0) Then
        CopyMemoryStrict VarPtr(dstColors(0)), VarPtr(m_ColorTable(0)), 256 * 4
        dstColorCount = m_ColorTableCount
        dstTransparentIndex = m_TransparentIndex
    Else
        dstColorCount = 0
        dstTransparentIndex = -1
    End If
End Sub

'Only valid after LoadPSD has been called
Friend Function GetHResolution() As Single
    GetHResolution = m_ResolutionInfo.riHRes
End Function

'Only valid after LoadPSD has been called
Friend Function GetVResolution() As Single
    GetVResolution = m_ResolutionInfo.riVRes
End Function

'Only valid after LoadPSD has been called
Friend Function GetBitsPerPixel() As Long
    
    Dim bpChannel As Long
    bpChannel = m_Header.BitsPerChannel
    
    'Bits per pixel is a little convoluted to calculate, as it depends on both color mode and color depth
    ' (and optionally, the presence of alpha).
    With m_Header
        If (.ColorMode = cm_Bitmap) Then
            GetBitsPerPixel = 1
        ElseIf (.ColorMode = cm_CMYK) Then
            GetBitsPerPixel = bpChannel * 4
            If Me.HasAlpha() Then GetBitsPerPixel = GetBitsPerPixel + bpChannel
        ElseIf (.ColorMode = cm_Duotone) Then
            GetBitsPerPixel = bpChannel
        ElseIf (.ColorMode = cm_Grayscale) Then
            GetBitsPerPixel = bpChannel
            If Me.HasAlpha() Then GetBitsPerPixel = GetBitsPerPixel * 2
        ElseIf (.ColorMode = cm_Indexed) Then
            GetBitsPerPixel = 8
        ElseIf (.ColorMode = cm_Lab) Then
            GetBitsPerPixel = bpChannel * 3
            If Me.HasAlpha() Then GetBitsPerPixel = GetBitsPerPixel + bpChannel
        
        'Multichannel doesn't have a meaningful analog in PD; treat it as grayscale
        ElseIf (.ColorMode = cm_Multichannel) Then
            GetBitsPerPixel = bpChannel
            If Me.HasAlpha() Then GetBitsPerPixel = GetBitsPerPixel * 2
        ElseIf (.ColorMode = cm_RGB) Then
            GetBitsPerPixel = bpChannel * 3
            If Me.HasAlpha() Then GetBitsPerPixel = GetBitsPerPixel + bpChannel
        Else
            GetBitsPerPixel = 32
        End If
    End With
    
End Function

'Only valid after LoadPSD has been called
Friend Function GetColorMode() As PSD_ColorMode
    GetColorMode = m_Header.ColorMode
End Function

'Only valid after LoadPSD has been called
Friend Function HasAlpha() As Boolean
    Dim i As Long
    For i = 0 To Abs(m_numOfLayers) - 1
        If Not (m_Layers(i) Is Nothing) Then
            If m_Layers(i).HasAlpha Then HasAlpha = True
        End If
    Next i
End Function

'Only valid after LoadPSD has been called
Friend Function HasICCProfile() As Boolean
    HasICCProfile = Not (m_Profile Is Nothing)
End Function

Friend Function GetICCProfile() As pdICCProfile
    Set GetICCProfile = m_Profile
End Function

'Only valid after LoadPSD has been called
Friend Function IsGrayscaleColorMode() As Boolean
    IsGrayscaleColorMode = (m_Header.ColorMode = cm_Grayscale) Or (m_Header.ColorMode = cm_Multichannel)
End Function

Private Function Step5_AssemblePDImage(ByRef srcFile As String, ByRef dstImage As pdImage) As PD_PSDResult

    On Error GoTo InternalVBError
    
    Step5_AssemblePDImage = psd_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "Step5_AssemblePDImage", "filename has changed since original validation!"
        Step5_AssemblePDImage = psd_Failure
        Exit Function
    End If
    
    'Set basic image attributes
    dstImage.Width = m_Header.ImageWidthPx
    dstImage.Height = m_Header.ImageHeightPx
    
    'Use retrieved DPI, if available, but if not, note that PSDs default to 72 ppi
    Dim hDPI As Single, vDPI As Single
    hDPI = Me.GetHResolution()
    vDPI = Me.GetVResolution()
    If (hDPI <= 0!) Then hDPI = 72!
    If (vDPI <= 0!) Then vDPI = 72!
    dstImage.SetDPI hDPI, vDPI
    
    Dim tmpLayer As pdLayer, numUsableLayers As Long, newLayerID As Long
    numUsableLayers = 0
    
    'Layer groups are not currently supported by PD, but we do use group data to enable group
    ' visibility/invisibility when loading groups.  Note that PSDs support nested layer groups,
    ' so a layer can be within e.g. a whole bunch of parent groups with varying depths.
    ' I currently use a formula where if *any* of the current parent layers are invisible,
    ' the current layer gets hidden too.  This produces the most "useful-looking" document
    ' (in lieu of actually supporting groups ourselves, ha).
    Dim curGroupDepth As Long
    curGroupDepth = 0
    
    'As of CS5, maximum nesting depth is 10.  PD covers larger depths than this, but we
    ' initialize our visibility tracker to 10 levels to start. (Source, good as of July 2020:
    ' https://feedback.photoshop.com/photoshop_family/topics/photoshop_please_remove_the_depth_limitation_on_nested_layer_groups_in_photoshop_cs4_and_earlier )
    Const MAX_GROUP_DEPTH As Long = 10
    Dim groupVisibility() As Boolean
    ReDim groupVisibility(0 To MAX_GROUP_DEPTH - 1) As Boolean
    Dim groupNames() As String
    ReDim groupNames(0 To MAX_GROUP_DEPTH - 1) As String
    
    'If any group *above* the current one is invisible, the current layer's visibility flag
    ' gets overriden with the parent group.  This is my preferred behavior as it ensures that
    ' layers belonging to invisible groups are not visible by default in the image PD constructs.
    ' (We update this flag whenever group depth changes.)
    Dim curGroupVisibility As Boolean
    curGroupVisibility = True
    
    'Layers representing group markers do not normally have images associated with them,
    ' *unless* the group has a layer mask.  If it does, we'll need to retrieve it and
    ' (TBD - at present, since we don't support groups, maybe just apply the mask to
    ' all layers within the group...?)
    Dim groupDIB As pdDIB
    
    'Now we need to iterate through layers and convert their data from PSD format to PD format.
    ' Note that we deliberately iterate layers *backward*.  I made this change because we need
    ' to iterate layer group *start* markers before the layers in that group, and Adobe lists
    ' groups as layers in top-down order (e.g. the start of a layer group appears "above" all
    ' members of that group).  This allows us to make sure children layers properly inherit
    ' group visibility state, at least until PD properly supports layer groups.
    Dim i As Long
    For i = Abs(m_numOfLayers) - 1 To 0 Step -1
        
        If m_Layers(i).DoesLayerHaveUsableData() Then
            
            'Multichannel images are handled differently; each channel is displayed in its own layer
            If (m_Header.ColorMode = cm_Multichannel) Then
            
                Dim j As Long
                For j = 0 To m_Layers(i).GetLayerCount - 1
                
                    'Prep a new layer object and initialize it with the image bits we've retrieved
                    newLayerID = dstImage.CreateBlankLayer()
                    Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                    
                    Dim newLayerName As String
                    newLayerName = g_Language.TranslateMessage("%1 (Channel %2)", m_Layers(i).GetLayerName(), m_Layers(i).GetMultiChannelLayerID(j))
                    tmpLayer.InitializeNewLayer PDL_Image, newLayerName, m_Layers(i).GetMultiChannelLayerDIB(j)
                
                    'Fill in any remaining layer properties
                    With m_Layers(i)
                        
                        'Hardcode "darken" blend mode so that the resulting layers composite together
                        ' into a rough estimation of the original result
                        tmpLayer.SetLayerBlendMode BM_Darken
                        tmpLayer.SetLayerOpacity .GetLayerOpacity()
                        tmpLayer.SetLayerOffsetX .GetLayerOffsetX()
                        tmpLayer.SetLayerOffsetY .GetLayerOffsetY()
                        tmpLayer.SetLayerVisibility .GetLayerVisibility()
                        
                    End With
                
                Next j
                
                numUsableLayers = numUsableLayers + 1
                
            'All other color modes are handled normally, with a 1:1 correspondence between
            ' PSD layers and PhotoDemon layers
            Else
            
                'Prep a new layer object and initialize it with the image bits we've retrieved
                newLayerID = dstImage.CreateBlankLayer()
                Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                tmpLayer.InitializeNewLayer PDL_Image, m_Layers(i).GetLayerName(), m_Layers(i).GetLayerDIB()
                
                'Fill in any remaining layer properties
                With m_Layers(i)
                    tmpLayer.SetLayerBlendMode .GetLayerBlendMode()
                    tmpLayer.SetLayerOpacity .GetLayerOpacity()
                    tmpLayer.SetLayerOffsetX .GetLayerOffsetX()
                    tmpLayer.SetLayerOffsetY .GetLayerOffsetY()
                    tmpLayer.SetLayerVisibility .GetLayerVisibility()
                End With
            
                numUsableLayers = numUsableLayers + 1
                
            End If
            
            'If we're inside a layer group, override the current layer visibility with its
            ' parent group visibility.
            If (curGroupDepth > 0) Then
                If (Not curGroupVisibility) Then tmpLayer.SetLayerVisibility False
            End If
        
        'If this layer doesn't contain pixel data (e.g. Not m_Layers(i).DoesLayerHaveUsableData()),
        ' it may be a group marker.  Set group flags accordingly.
        Else
            
            'If pixel data is unusable AND this is not a group marker, we'll use this flag to
            ' add some debug data at the end of this block.
            Dim layerDataRelevant As Boolean
            
            'Layer groups may require a dummy DIB as their backing surface (a blank 1x1 transparent DIB).
            ' This may be replaced with a layer group mask, as relevant.
            Dim dummyDIB As pdDIB
            
            'When a new layer group, increment our group depth counter
            If m_Layers(i).IsLayerGroupStart() Then
                
                layerDataRelevant = True
                
                'Add this group's visibility and name to our running collection.  (We track this info
                ' so we can properly handle complex nested group arrangements.)
                If (curGroupDepth > UBound(groupVisibility)) Then
                    ReDim Preserve groupVisibility(0 To curGroupDepth * 2 - 1) As Boolean
                    ReDim Preserve groupNames(0 To curGroupDepth * 2 - 1) As String
                End If
                
                groupVisibility(curGroupDepth) = m_Layers(i).GetLayerVisibility()
                groupNames(curGroupDepth) = m_Layers(i).GetLayerName()
                
                curGroupVisibility = curGroupVisibility And groupVisibility(curGroupDepth)
                
                'PD can optionally create "dummy" layers to represent the group hierarchy.  This is a
                ' workaround until PD properly supports layer groups.
                If PSD_GROUPS_AS_DUMMY_LAYERS Then
                    
                    'Prep a new "dummy" layer object to represent this group
                    newLayerID = dstImage.CreateBlankLayer()
                    Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                    
                    If (dummyDIB Is Nothing) Then
                        Set dummyDIB = New pdDIB
                        dummyDIB.CreateBlank 1, 1, 32, 0, 0
                        dummyDIB.SetInitialAlphaPremultiplicationState True
                    End If
                    
                    'Photoshop allows layer groups to have their own masks (a nice feature).  If we encountered
                    ' a mask while loading this group, use the mask image instead of the dummy DIB.
                    If (m_Layers(i).GetLayerDIB Is Nothing) Then
                        Set groupDIB = dummyDIB
                    Else
                        If PSD_LOAD_GROUP_MASKS_AS_PIXELS Then
                            Set groupDIB = m_Layers(i).GetLayerDIB
                        Else
                            Set groupDIB = dummyDIB
                        End If
                    End If
                    
                    tmpLayer.InitializeNewLayer PDL_Image, g_Language.TranslateMessage("Group start:") & " " & m_Layers(i).GetLayerName, groupDIB, True
                    
                    'Fill in any remaining layer properties
                    With m_Layers(i)
                        tmpLayer.SetLayerBlendMode .GetLayerBlendMode()
                        tmpLayer.SetLayerOpacity .GetLayerOpacity()
                        tmpLayer.SetLayerVisibility m_Layers(i).GetLayerVisibility()
                        
                        If (Not PSD_LOAD_GROUP_MASKS_AS_PIXELS) Then
                            If (Not m_Layers(i).GetMaskMergedIntoDIB) And m_Layers(i).GetMaskExists And (m_Layers(i).GetMaskUnmergedSize <> 0) Then
                                
                                tmpLayer.SetGenericLayerProperty pgp_MaskExists, True
                                
                                'TODO: also mark the mask as active; this is disabled currently, pending full group support
                                'tmpLayer.SetGenericLayerProperty pgp_MaskActive, True
                                Dim srcMaskRectF As RectF
                                srcMaskRectF = m_Layers(i).GetMaskRectF
                                
                                tmpLayer.GetLayerMask.SetMaskRect srcMaskRectF.Left, srcMaskRectF.Top, srcMaskRectF.Width, srcMaskRectF.Height
                                tmpLayer.GetLayerMask.SetMaskBytes srcMaskRectF.Width, srcMaskRectF.Height, m_Layers(i).GetMaskUnmergedPtr
                                tmpLayer.GetLayerMask.SetOpacityOutsideMask CSng(m_Layers(i).GetMaskDefaultColor) / 2.55!
                                
                                'For now, we also need to flag the mask with a special "came from PSD layer group" marker,
                                ' so that PhotoDemon knows how to handle the mask.
                                tmpLayer.GetLayerMask.SetMaskFlag_PSGroup
                                
                            End If
                        End If
                    
                    End With
                
                    numUsableLayers = numUsableLayers + 1
                
                End If
                
                'Increment nested group counter
                curGroupDepth = curGroupDepth + 1
                
            End If
            
            'Look for the end of the current layer group (if any)
            If m_Layers(i).IsLayerGroupEnd() Then
                
                layerDataRelevant = True
                
                'Decrement depth counter, with a failsafe check for malformed PSDs
                curGroupDepth = curGroupDepth - 1
                If (curGroupDepth < 0) Then
                    curGroupDepth = 0
                    m_Warnings.AddString "malformed layer groups found; curGroupDepth below zero!"
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
                
                'PD can optionally create "dummy" layers to represent the group hierarchy.  This is a
                ' workaround until PD properly supports layer groups.
                If PSD_GROUPS_AS_DUMMY_LAYERS Then
                    
                    'Prep a new "dummy" layer object to represent this group
                    newLayerID = dstImage.CreateBlankLayer()
                    Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                    
                    If (dummyDIB Is Nothing) Then
                        Set dummyDIB = New pdDIB
                        dummyDIB.CreateBlank 1, 1, 32, 0, 0
                    End If
                    
                    'For layer *start* markers, this is where we'd check for a mask and process it as relevant.
                    ' Layer *end* markers do not typically carry meaningful data.  (I can always revisit this if
                    ' a new PSD shows up with contradictory behavior!)
                    tmpLayer.InitializeNewLayer PDL_Image, g_Language.TranslateMessage("Group end:") & " " & groupNames(curGroupDepth), dummyDIB, True
                    
                    'Fill in any remaining layer properties
                    With m_Layers(i)
                        tmpLayer.SetLayerBlendMode .GetLayerBlendMode()
                        tmpLayer.SetLayerOpacity .GetLayerOpacity()
                        tmpLayer.SetLayerVisibility groupVisibility(curGroupDepth)
                    End With
                
                    numUsableLayers = numUsableLayers + 1
                
                End If
                
            End If
            
            If (Not layerDataRelevant) Then
                If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Skipping layer " & i & ": layer data unusable"
            End If
            
        End If
        
    Next i
    
    'We now need to perform one last failsafe check to ensure that at least *one* valid layer exists.
    If (numUsableLayers <= 0) Then
    
        'No useable layers exist.  This typically means the image is comprised of nothing but
        ' vector layers (or effect/fill layers), which PD can't interpret.
        
        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "No usable layers found; using composite image instead..."
        
        'Instead of just bombing out, we're going to try and retrieve the composite image.
        If (m_CompositeImageOffset > 0) Then
            
            'High bit-depth images move the underlying stream object around in unpredictable ways.
            ' Reset the entire stream accordingly.
            If m_HighBitDepthToggle Then
                Set m_Stream = New pdStream
                If Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, m_SourceFilename, , , OptimizeSequentialAccess) Then
                    m_Warnings.AddString "Can't read file; it may be locked or in an inaccessible location."
                    Step5_AssemblePDImage = psd_Failure
                    Exit Function
                End If
            End If
            
            'Move the stream pointer into position, then attempt to load the composite image
            m_Stream.SetPosition m_CompositeImageOffset, FILE_BEGIN
            Step5_AssemblePDImage = Step4B_SingleLayerImage()
            
            'If the composite image loaded successfully, initialize it as the only layer in the target image
            If (Step5_AssemblePDImage < psd_Failure) Then
            
                'Prep a new layer object and initialize it with the image bits we've retrieved
                newLayerID = dstImage.CreateBlankLayer()
                Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                tmpLayer.InitializeNewLayer PDL_Image, g_Language.TranslateMessage("Background"), m_Layers(0).GetLayerDIB()
                
                'Fill in any remaining layer properties
                With tmpLayer
                    .SetLayerBlendMode BM_Normal
                    .SetLayerOpacity 100!
                    .SetLayerOffsetX 0
                    .SetLayerOffsetY 0
                    .SetLayerVisibility True
                End With
            
            End If
            
        End If
    
    'At least one valid layer exists, so finalize the image, then exit.
    Else
        
        'Because layers were loaded in top-down order, we now need to *reverse* the layer stack
        dstImage.ReverseLayerOrder
        
        'Make sure we were able to generate at least one usable layer.
        If (numUsableLayers <= 0) Or (dstImage.GetNumOfLayers <= 0) Then
            m_Warnings.AddString "unable to create any usable layers.  Abandoning load."
            Set m_Stream = Nothing
            Step5_AssemblePDImage = psd_Failure
            Exit Function
        End If
        
    End If
    
    'Look for the "clipping" flag on any loaded layers.  If it is set, we can mimic the clipping results by
    ' merging the alpha channel of any clipped layers with the alpha channel of the lowest non-background
    ' layer in the image (I think?  Photoshop's strategy may be more complicated than this, but I don't
    ' have enough images available to test it well.)
    
    ' TODO!
    
    'Finally, free the underlying stream
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen() Then m_Stream.StopStream True
    End If
    
    Exit Function

InternalVBError:
    InternalError "Step5_AssemblePDImage", "internal VB error #" & Err.Number & ": " & Err.Description
    m_Warnings.AddString "Internal error in step 5, #" & Err.Number & ": " & Err.Description
    Step5_AssemblePDImage = psd_Failure
    
End Function

Private Function Step4B_SingleLayerImage() As PD_PSDResult

    'If a PSD only contains one layer, the layer's data will be saved at the end of the file,
    ' in the "image data" section.  Step4_GatherLayersAndMasks() will defer to us in this case,
    ' so we can manually load the remaining layer data.
    m_numOfLayers = 1
    ReDim m_Layers(0) As pdPSDLayer
    Set m_Layers(0) = New pdPSDLayer
    
    Dim addLayerInfo As pdPSDLayerInfo
    Set addLayerInfo = New pdPSDLayerInfo
    
    'With the pointer now correctly aligned, we can proceed with gathering pixel data
    Step4B_SingleLayerImage = m_Layers(0).NotifySingleLayerImage(m_Stream, m_Warnings, m_PSDisPSB, m_Header.NumChannels, m_Header.ImageWidthPx, m_Header.ImageHeightPx, m_Header.BitsPerChannel, m_Header.ColorMode, addLayerInfo)
    
    'If image data was parsed successfully, proceed with decoding
    If (Step4B_SingleLayerImage < psd_Failure) Then
        If (m_Layers(0).DecodeChannels(m_Warnings, m_PSDisPSB, m_Header.BitsPerChannel, m_Header.ColorMode, Me) < psd_Failure) Then
            Step4B_SingleLayerImage = m_Layers(0).ConstructImage(m_Stream, m_Warnings, m_PSDisPSB, m_Header.BitsPerChannel, m_Header.ColorMode, m_Profile, Me)
        Else
            InternalError "Step4B_SingleLayerImage", "pdPSDLayer.DecodeChannels() failed catastrophically"
            Step4B_SingleLayerImage = psd_Failure
        End If
    End If
    
End Function

Private Function Step4_GatherLayersAndMasks(ByRef srcFile As String) As PD_PSDResult

    On Error GoTo InternalVBError
    
    Step4_GatherLayersAndMasks = psd_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "Step4_GatherLayersAndMasks", "filename has changed since original validation!"
        Step4_GatherLayersAndMasks = psd_Failure
        Exit Function
    End If
    
    'The layer and mask section of a PSD is extremely complicated.  Almost all records are variable-length,
    ' and values within records can also be variable-length.  Individual bits often need to be read to know
    ' how many more bytes to read.  It's all very, very ugly.
    
    'For now, our primary goal is retrieving layer channel data.  Other (optional) attributes can be extracted
    ' in the future.
    
    'First up is the length of the entire layer/mask section.  This is 4-bytes in PSDs and 8-bytes in PSBs.
    ' (While we handle both cases, note that PD is unlikely to work with PSBs on account of being 32-bit.)
    Dim lenOfSection As Long, finalPointerPosition As Long
    lenOfSection = m_Stream.ReadLong_BE()
    If m_PSDisPSB Then lenOfSection = m_Stream.ReadLong_BE()        'TODO: fix PSB approach?
    finalPointerPosition = m_Stream.GetPosition() + lenOfSection
    
    'Note that the end of this segment is the Image Data section, which stores a copy of the merged image
    ' (if one exists).  Flag that position so we can use it to salvage the image if things go horribly awry.
    m_CompositeImageOffset = finalPointerPosition
    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Length of layer and mask section (offset " & CStr(m_Stream.GetPosition - 4) & "): " & lenOfSection
    
    'If a PSD only contains one layer, it can be written at the end of the file (in the
    ' "image data" segment), bypassing this section completely.  To cover this special case,
    ' we hand off control to a separate function (as the data layout differs in subtle,
    ' obnoxious ways).
    If (lenOfSection <= 0) Then
        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Single-layer PSD found.  Retrieving layer data from image segment instead..."
        m_CompositeImageOffset = 0
        Step4_GatherLayersAndMasks = Step4B_SingleLayerImage()
        Exit Function
    End If
    
    'The remainder of the chunk is comprised of three subsections:
    ' 1) Layer info
    ' 2) Global layer mask info
    ' 3) Tagged blocks with miscellaneous layer/mask data (PS4 or later, only)
    
    'The layer info section has its own length marker, and per the spec, it is
    ' "rounded up to a multiple of 2".
    Dim lenOfLayerInfoSection As Long, finalLayerInfoPosition As Long
    lenOfLayerInfoSection = m_Stream.ReadLong_BE()
    If m_PSDisPSB Then lenOfLayerInfoSection = m_Stream.ReadLong_BE()        'TODO: fix PSB approach
    finalLayerInfoPosition = m_Stream.GetPosition() + finalLayerInfoPosition
    
    'The layer info section *can* have zero-length; this is common on 16-bit images, for example,
    ' as the actual layer data is stored in a separate chunk for backward-compatibility reasons.
    ' (Also, see the psd-tools collection of test images - they cover this case.)
    If (lenOfLayerInfoSection = 0) Then
        
        'Layer info section is null.  We now have one of two possibilities:
        ' 1) This is a single-layer image, and PhotoShop has placed the background layer data in the
        '     composite image chunk at the end of the file.  We can grab that, no problem.
        ' 2) This is a high bit-depth image, and PhotoShop has placed the actual layer data inside a
        '     specialized "Lr16" or "Lr32" additional info block (presumably for backward-compatibility).
        '
        ' We can distinguish between these two states using the length of the layer chunk.
        Dim hpdChunkFound As Boolean
        hpdChunkFound = False
        
        If (m_Stream.GetPosition < finalPointerPosition) Then
            
            'Even though there are no layers, there is global mask data and/or additional
            ' "global" layer info blocks (for lack of a better term; the spec is nebulous on
            ' what these blocks may even hold).
            
            'Global mask data gets checked first.  (We don't use this data in PD, so we just
            ' read it using a temp layer.)
            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Checking for global mask data..."
            Dim tmpLayer As pdPSDLayer
            Set tmpLayer = New pdPSDLayer
            Step4_GatherLayersAndMasks = tmpLayer.ParseGlobalLayerMaskInfo(m_Stream, m_Warnings, m_PSDisPSB)
            
            'Finally, if the stream pointer is *still* not at the end of the section, it means optional
            ' additional tagged blocks are present.  Newer versions of Photoshop may write critical
            ' image data here (e.g. pixel data for 16/32-bit channels) so it's crucial to retrieve
            ' these segments.
            Dim addLayerInfo As pdPSDLayerInfo
            Set addLayerInfo = New pdPSDLayerInfo
            If (m_Stream.GetPosition < finalPointerPosition) Then
            
                'Parse and acquire all tagged blocks, then reset the stream pointer to a
                ' known-aligned position.
                addLayerInfo.ParseAdditionalLayerInfo m_Stream, m_Warnings, m_PSDisPSB, finalPointerPosition, True
                m_Stream.SetPosition finalPointerPosition, FILE_BEGIN
                
            End If
            
            'If we found a high bit-depth segment, we want to continue parsing the file, but using that
            ' high bit-depth segment as if it were the actual file contents.
            If addLayerInfo.DoesKeyExist("Lr16") Or addLayerInfo.DoesKeyExist("Lr32") Then
                
                hpdChunkFound = True
                
                'Reset our original stream; we will instead parse ensuing layer data from the
                ' target block.
                Set m_Stream = New pdStream
                
                'Instead, wrap the m_Stream object around the high bit-depth block
                If addLayerInfo.DoesKeyExist("Lr16") Then
                    Set m_Stream = addLayerInfo.GetStreamForKey("Lr16")
                Else
                    Set m_Stream = addLayerInfo.GetStreamForKey("Lr32")
                End If
                
                'Reset any/all pointer trackers to match the new stream length and position.
                lenOfLayerInfoSection = m_Stream.GetStreamSize()
                
            End If
            
            m_HighBitDepthToggle = hpdChunkFound
            
        End If
    
        'If we didn't find a high bit-depth segment, this is just a plain single-layer image.
        ' The single background layer's contents live in the composite image segment.
        If (Not hpdChunkFound) Then
            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Single-layer PSD found.  Activating alternate parse strategy..."
            Step4_GatherLayersAndMasks = Step4B_SingleLayerImage()
            Exit Function
        End If
    
    End If
    
    'By this point, if we still have a non-zero layer info segment, we can parse it normally,
    ' regardless of where the data actually resides.  (Previous steps may have pointed the
    ' stream at an alternate segment of the PSD file without us knowing; that's by design.)
    If (lenOfLayerInfoSection > 0) Then
        
        Dim i As Long
        
        'Next comes a two-byte "layer count", which is actually a signed integer.  Per the spec:
        ' "If it is a negative number, its absolute value is the number of layers and the first
        '  alpha channel contains the transparency data for the merged result."
        Dim numOfLayers As Integer
        numOfLayers = m_Stream.ReadInt_BE()
        
        'Again, not sure if the layer count can be zero-length, but just in case...
        If (numOfLayers <> 0) Then
            
            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction Abs(numOfLayers) & " layers found.  Parsing layer data now..."
            
            'The "layer records" section comes next.  Layer records describe everything needed to perfectly
            ' re-create a Photoshop layer.  As you can imagine, these records are massive, and most of the
            ' information is useful only to Adobe.
            
            'Because these records are so complex, they are parsed by a child class, and that class stores
            ' whatever relevant information we are able to extract.
            
            'Start by initializing a collection of layer child objects
            m_numOfLayers = numOfLayers
            ReDim m_Layers(0 To Abs(m_numOfLayers) - 1) As pdPSDLayer
            
            'We now rely on the child class to handle further processing.  Iterate through each layer
            ' in turn and hand it off to a new child class instance.
            For i = 0 To Abs(m_numOfLayers) - 1
                Set m_Layers(i) = New pdPSDLayer
                If (Not m_Layers(i).ParseLayer(m_Stream, m_Warnings, m_PSDisPSB) < psd_Failure) Then
                    InternalError "Step4_GatherLayersAndMasks", "pdPSDLayer.ParseLayer() failed catastrophically"
                    Step4_GatherLayersAndMasks = psd_Failure
                    Exit Function
                End If
            Next i
            
        '/end null layer count check
        End If
        
        'With each layer loaded, we can now proceed with loading channel data (eek).  Because Adobe stores
        ' channel data in planar order (not interleaved, e.g. RRRGGGBBB instead of RGBRGBRGB) this step
        ' consumes a lot of memory, as we can't easily stream in chunks of the data at a time - instead,
        ' we have to allocate the full memory and populate it as-we-go.
        For i = 0 To Abs(m_numOfLayers) - 1
            If (m_Layers(i).LoadChannels(m_Stream, m_Warnings, m_PSDisPSB, m_Header.BitsPerChannel, m_Header.ColorMode) >= psd_Failure) Then
                InternalError "Step4_GatherLayersAndMasks", "pdPSDLayer.LoadChannels() failed catastrophically"
                Step4_GatherLayersAndMasks = psd_Failure
                Exit Function
            End If
        Next i
        
        'If global mask data is present, parse it now.
        If (Step4_GatherLayersAndMasks < psd_Failure) Then Step4_GatherLayersAndMasks = m_Layers(0).ParseGlobalLayerMaskInfo(m_Stream, m_Warnings, m_PSDisPSB)
        
        'Finally, an additional set of additional layer information (that's not confusing at all)
        ' may follow the global mask.  This section is entirely optional, and the spec is fuzzy about
        ' what these chunks may contain.  We don't have a use for it in PD at present, so we just
        ' skip this segment pending further investigation.
        If (Step4_GatherLayersAndMasks < psd_Failure) Then
        
            'TODO?
        
        End If
        
        'With all channels loaded, we can now ask each layer to decode the channel data into a
        ' format we can actually interpret.  The amount of work involved varies by channel
        ' type (e.g. layer masks differ from color channels) and compression type.
        For i = 0 To Abs(m_numOfLayers) - 1
            
            'To avoid wasting time, skip layers with null sizes - but *only* if the layer is *not* a group.
            ' (Layer group markers have width/height of 0, but they *are* allowed to store masks.)
            
            If (m_Layers(i).GetLayerWidth <= 0) Or (m_Layers(i).GetLayerHeight <= 0) Then
                If (Not m_Layers(i).IsLayerGroupStart) And (Not m_Layers(i).IsLayerGroupEnd) Then GoTo skipLayer
            End If
            
            If (m_Layers(i).DecodeChannels(m_Warnings, m_PSDisPSB, m_Header.BitsPerChannel, m_Header.ColorMode, Me) < psd_Failure) Then
            
                'If the channel decoded successfully, convert it from a bare byte/int/float stream
                ' into usable pixel data. This comprehensive stage translates channel data to
                ' 8-bpc, converts the resulting channel collection from planar to 32-bpp interleaved
                ' RGBA, and color-manages the results.  The end result is a PD-compatible image buffer
                ' that can be directly associated with a pdLayer object.
                If (Not m_Layers(i).ConstructImage(m_Stream, m_Warnings, m_PSDisPSB, m_Header.BitsPerChannel, m_Header.ColorMode, m_Profile, Me, PSD_LOAD_GROUP_MASKS_AS_PIXELS) < psd_Failure) Then
                    InternalError "Step4_GatherLayersAndMasks", "pdPSDLayer.ConstructImage() failed catastrophically"
                    Step4_GatherLayersAndMasks = psd_Failure
                    Exit Function
                End If
                
            Else
                InternalError "Step4_GatherLayersAndMasks", "pdPSDLayer.DecodeChannels() failed catastrophically"
                Step4_GatherLayersAndMasks = psd_Failure
                Exit Function
            End If
skipLayer:
        Next i
    
    'Layer info section is null-length; layer data has already been located using an alternate strategy
    End If
    
    'At the end of this step, we are probably finished with the source stream object.  However, we may need to
    ' access it if PD is unable to recover any layers from the layer data segment (if for example the image
    ' is comprised solely of vector or fill layers that PD can't parse).  As such, leave the base stream open
    ' "just in case"; we'll free it in a later step.
    
    Exit Function

InternalVBError:
    InternalError "Step4_GatherLayersAndMasks", "internal VB error #" & Err.Number & ": " & Err.Description
    If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 4, #" & Err.Number & ": " & Err.Description
    Step4_GatherLayersAndMasks = psd_Failure
    
End Function

'From the spec: "Image resource blocks are the basic building unit of several file formats, including Photoshop's
' native file format, JPEG, and TIFF. Image resources are used to store non-pixel data associated with images,
' such as pen tool paths. They are referred to as resource blocks because they hold data that was stored in the
' Macintosh's resource fork in early versions of Photoshop.
Private Function Step3_GatherImageResources(ByRef srcFile As String, ByRef dstImage As pdImage) As PD_PSDResult

    On Error GoTo InternalVBError
    
    Step3_GatherImageResources = psd_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "Step3_GatherImageResources", "filename has changed since original validation!"
        Step3_GatherImageResources = psd_Failure
        Exit Function
    End If
    
    'Remember that m_Stream has already been forcibly advanced past the file header (step 1)
    ' and color mode data (step 2).  It now points at the image resources section.
    
    'As before, the first marker in this section is the length of the ENTIRE section.  As a failsafe
    ' against parse errors, we want to make a note of the segment end pointer; after iterating through
    ' all resource blocks, we will double-check this value against wherever the pointer naturally
    ' ended up.
    Dim resSectionLength As Long
    resSectionLength = m_Stream.ReadLong_BE()
    m_NumImageResources = 0
    
    If (resSectionLength > 0) Then
        
        'Calculate a final position for this segment; we'll continue iterating blocks until we hit this
        Dim finalPosition As Long
        finalPosition = m_Stream.GetPosition() + resSectionLength
        
        'Initialize our resources collection
        ReDim m_ImageResources(0 To 3) As PSD_ImageResource
        
        Dim resNameLength As Byte, resSize As Long
        
        'In the future, we will likely want to separate and store the various resource blocks we encounter.
        ' Ideally, these blocks could even be preserved if a PSD is saved back out to file during a session.
        ' For this first prototype, however, we simply want to iterate all blocks and ensure that our
        ' pointer math works out correctly.
        Do
            
            Const IMAGE_RESOURCE_ID As String = "8BIM"
            
            'Blocks *must* start with a unique '8BIM' identifier
            If (m_Stream.ReadString_ASCII(4) = IMAGE_RESOURCE_ID) Then
            
                'Prep storage
                If (m_NumImageResources > UBound(m_ImageResources)) Then ReDim Preserve m_ImageResources(0 To m_NumImageResources * 2 - 1) As PSD_ImageResource
                
                'Resource ID is a Photoshop-specific identifier.  There are 100+ possible values here;
                ' we do not currently attempt to validate these.
                m_ImageResources(m_NumImageResources).irID = m_Stream.ReadInt_BE()
                
                'The ID is followed by a Pascal ShortString.  This is a one-byte length value, followed by a
                ' string of chars (current system codepage) of length 0-255.  Adobe forcibly pads these to be
                ' a length with a multiple of two, *including* the length bit (e.g. per the spec, "a null name
                ' consists of two bytes of zero").  In practice, these seem to always be null-strings.
                resNameLength = m_Stream.ReadByte()
                If (resNameLength = 0) Then
                    m_Stream.SetPosition 1, FILE_CURRENT   'Forcibly advance the pointer by 1 more byte
                Else
                    'Retrieve the string, and advance by an additional byte if the string length is even.
                    m_ImageResources(m_NumImageResources).irName = Strings.TrimNull(m_Stream.ReadString_ASCII(resNameLength))
                    If ((resNameLength And 1) = 0) Then m_Stream.SetPosition 1, FILE_CURRENT
                End If
                
                'Next, read the length of the resource data.  Again, note that this is padded to make the
                ' size even, if the resource data itself is *not* even.
                resSize = m_Stream.ReadLong_BE()
                m_ImageResources(m_NumImageResources).irDataLength = resSize
                If PSD_DEBUG_VERBOSE Then
                    With m_ImageResources(m_NumImageResources)
                        PDDebug.LogAction "Image resource: 0x" & Hex$(.irID) & ", " & .irDataLength & " bytes"
                    End With
                End If
                
                'Before touching the actual resource data, calculate a final pointer position.
                ' This is important as...
                '  1) we may not actually retrieve resource data because it's enormous and useless to us, or...
                '  2) we may retrieve only a portion of the data, and...
                '  3) resource data is padded to even numbers, so the useful size of the data may
                '     not be the amount we actually need to move the pointer.
                '
                'By calculating a final position in advance, we can forcibly set the stream pointer to
                ' this value, regardless of what we do with the actual resource data.
                Dim finalResPosition As Long
                finalResPosition = m_Stream.GetPosition() + resSize
                If ((resSize And 1) = 1) Then finalResPosition = finalResPosition + 1
                
                'Retrieve the resource (if its length is non-zero; some 3rd-party PSD writers, like PhotoPea,
                ' write 0-length resources for reasons unknown - see https://github.com/PsdPlugin/PsdPlugin/issues/17)
                If (resSize > 0) Then m_Stream.ReadBytes m_ImageResources(m_NumImageResources).irDataBytes, resSize, True
                
                'Realign the pointer to match the previously calculated "end-of-resource" position.
                m_Stream.SetPosition finalResPosition, FILE_BEGIN
                
                m_NumImageResources = m_NumImageResources + 1
                
            Else
                m_Warnings.AddString "Image resource block starts with invalid identifier!"
            End If
        
        Loop While (m_Stream.GetPosition() < finalPosition)
        
        'In case image resource blocks have added padding, forcibly set the stream position
        ' to our calculated final value.
        m_Stream.SetPosition finalPosition, FILE_BEGIN
        
        'Iterate the completed resource collection, looking for resources useful to PD
        If (m_NumImageResources > 0) Then
            
            'Temporary stream objects are useful for iterating complex resource data streams
            Dim tmpStream As pdStream
            
            Dim i As Long
            For i = 0 To m_NumImageResources - 1
                
                With m_ImageResources(i)
                    
                    Select Case .irID
                    
                        'ICC profiles are critical, especially if the source is CMYK
                        Case rid_IccProfile
                            
                            If (.irDataLength > 0) Then
                            
                                Set m_Profile = New pdICCProfile
                                m_Profile.LoadICCFromPtr .irDataLength, VarPtr(.irDataBytes(0))
                                If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "ICC profile detected (" & CStr(.irDataLength) & " bytes)"
                                
                                'Add the retrieved profile to PD's central cache, and tag the destination image
                                ' to note that the image is color-managed
                                Dim colorProfileHash As String
                                colorProfileHash = ColorManagement.AddProfileToCache(m_Profile)
                                dstImage.SetColorProfile_Original colorProfileHash
                        
                            End If
                            
                        'Resolution info must be preserved
                        Case rid_ResolutionInfo
                        
                            If (.irDataLength > 0) Then
                            
                                Set tmpStream = New pdStream
                                tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , .irDataLength, VarPtr(.irDataBytes(0))
                                
                                'Resolutions are stored as PPI in 32-bit fixed-point, while all other indicators
                                ' are ushorts.
                                m_ResolutionInfo.riHRes = tmpStream.ReadFixed1616_BE()
                                m_ResolutionInfo.riHResUnit = tmpStream.ReadInt_BE()
                                m_ResolutionInfo.riWidthUnit = tmpStream.ReadInt_BE()
                                m_ResolutionInfo.riVRes = tmpStream.ReadFixed1616_BE()
                                m_ResolutionInfo.riVResUnit = tmpStream.ReadInt_BE()
                                m_ResolutionInfo.riHeightUnit = tmpStream.ReadInt_BE()
                                
                                Set tmpStream = Nothing
                                
                            End If
                            
                        'In indexed images, we are interested in both the indexed table color count
                        ' and transparency index, if any
                        Case rid_IndexedColorTableCount
                            If (.irDataLength > 0) Then
                                Set tmpStream = New pdStream
                                tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , .irDataLength, VarPtr(.irDataBytes(0))
                                m_ColorTableCount = tmpStream.ReadInt_BE()
                                Set tmpStream = Nothing
                            End If
                            
                        Case rid_TransparentIndex
                            If (.irDataLength > 0) Then
                                Set tmpStream = New pdStream
                                tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , .irDataLength, VarPtr(.irDataBytes(0))
                                m_TransparentIndex = tmpStream.ReadInt_BE()
                                Set tmpStream = Nothing
                            End If
                            
                        'At import time, "Version Info" is mostly a curiosity (it tells us what program wrote the PSD),
                        ' but because we are a well-behaved engine, we also need to *write* this data at save time.
                        ' PD always imports anything it attempts to write so we can verify correct round-trip behavior.
                        Case rid_VersionInfo
                            If (.irDataLength > 0) Then
                                
                                Set tmpStream = New pdStream
                                tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , .irDataLength, VarPtr(.irDataBytes(0))
                                
                                Dim versionNumber As Long
                                versionNumber = tmpStream.ReadLong_BE()
                                
                                Dim hasRealMergedData As Boolean
                                hasRealMergedData = (tmpStream.ReadByte() <> 0)
                                
                                Dim writerName As String, readerName As String
                                resNameLength = tmpStream.ReadLong_BE()
                                writerName = tmpStream.ReadString_Unicode_BE(resNameLength)
                                resNameLength = tmpStream.ReadLong_BE()
                                readerName = tmpStream.ReadString_Unicode_BE(resNameLength)
                                
                                Dim fileVersion As Long
                                fileVersion = tmpStream.ReadLong_BE()
                                
                                If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "PSD version info is: " & versionNumber & ", " & hasRealMergedData & ", " & fileVersion & ", " & writerName & ", " & readerName
                                
                            End If
                            
                        'PhotoDemon does "support" round-tripping of layer group data.  (It does this using
                        ' dummy layers to represent group start/end markers.)  I have no idea what this
                        ' "Groups Enabled ID" setting does.  In the wild, this value is always zero for
                        ' each layer in the image.  PD writes out a similar set of zeroes when saving,
                        ' so we don't actually need to retrieve this data at present (although the code
                        ' below *will* retrieve it just fine.
'                        Case rid_LayerGroupsEnabledID
'
'                            If (.irDataLength > 0) Then
'
'                                Set tmpStream = New pdStream
'                                tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , .irDataLength, VarPtr(.irDataBytes(0))
'
'                                'I haven't figured out whether this information is actually useful.
'                                ' PD can parse layer group data without this at present, so I've disabled
'                                ' parsing + storage until I have a better grasp of what this data even
'                                ' means.
'                                Dim grpID As Long
'                                For grpID = 0 To .irDataLength - 1
'                                     Debug.Print tmpStream.ReadByte()
'                                Next grpID
'
'                                Set tmpStream = Nothing
'
'                            End If
                            
                        'See above comments on rid_LayerGroupsEnabledID.  This segment doesn't make
                        ' sense to me, and parsing it on a few test PSDs has shown no useful data
                        ' from it.  I've left the parsing code in case files turn up "in the wild"
                        ' where this information actually matters.  (Like that chunk, these values
                        ' always show up as "1" for each layer in the image.  PD writes out a similar
                        ' chunk at save-time, just in case PS needs it.)
'                        Case rid_LayersGroupInfo
'
'                            If (.irDataLength > 0) Then
'
'                                Set tmpStream = New pdStream
'                                tmpStream.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , .irDataLength, VarPtr(.irDataBytes(0))
'
'                                Debug.Print "Group info chunk: " & .irDataLength & " bytes total"
'                                For grpID = 0 To (.irDataLength \ 2) - 1
'                                    Debug.Print tmpStream.ReadIntUnsigned_BE()
'                                Next grpID
'
'                                Set tmpStream = Nothing
'
'                            End If
                            
                    End Select
                    
                End With
            
            Next i
        
        End If
        
    'I don't believe files coming from PD will ever have a zero-length resource segment, but 3rd-party implementations
    ' may not use this section at all.  In that case, the stream pointer will already be pointing at the next segment.
    Else
        Step3_GatherImageResources = psd_Success
    End If
    
    Exit Function

InternalVBError:
    InternalError "Step3_GatherImageResources", "internal VB error #" & Err.Number & ": " & Err.Description
    If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 3, #" & Err.Number & ": " & Err.Description
    Step3_GatherImageResources = psd_Failure
    
End Function

Private Function Step2_RetrieveColorModeData(ByRef srcFile As String) As PD_PSDResult

    On Error GoTo InternalVBError
    
    Step2_RetrieveColorModeData = psd_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "Step2_RetrieveColorModeData", "filename has changed since original validation!"
        Step2_RetrieveColorModeData = psd_Failure
        Exit Function
    End If
    
    'Remember that m_Stream has already been forcibly advanced past the file header in Step 1.
    
    'We now want to retrieve the PSD's color mode data section, if it exists.  Per the spec,
    ' "Only indexed color and duotone have color mode data. For all other modes, this section is just
    '  the 4-byte length field, which is set to zero."
    Dim cmLength As Long
    cmLength = m_Stream.ReadLong_BE()
    
    Dim finalPosition As Long
    finalPosition = m_Stream.GetPosition() + cmLength
    
    If (cmLength > 0) Then
    
        'Ensure the file is indexed or duotone.
        If (m_Header.ColorMode = cm_Indexed) Then
            
            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Color table found; extracting now..."
            
            'Indexed mode images used a fixed-size color table.  From the spec:
            ' "Indexed color images: length is 768; color data contains the color table for the image,
            ' in non-interleaved order."
            
            'Retrieve the color table and assemble it into a usable palette.
            m_ColorTableCount = 256
            m_TransparentIndex = -1
            ReDim m_ColorTable(0 To 255) As RGBQuad
            
            Dim i As Long
            For i = 0 To 255
                m_ColorTable(i).Red = m_Stream.ReadByte()
            Next i
            For i = 0 To 255
                m_ColorTable(i).Green = m_Stream.ReadByte()
            Next i
            For i = 0 To 255
                m_ColorTable(i).Blue = m_Stream.ReadByte()
            Next i
            For i = 0 To 255
                m_ColorTable(i).Alpha = 255
            Next i
            
            'The stream pointer will now be pointing at the start of the next segment.
            Step2_RetrieveColorModeData = psd_Success
        
        ElseIf (m_Header.ColorMode = cm_Duotone) Then
        
            'Duotone data is not actually retrievable.  It is a proprietary Adobe format.
            ' Per the spec, "Duotone images: color data contains the duotone specification (the format of
            ' which is not documented). Other applications that read Photoshop files can treat a duotone
            ' image as a gray image, and just preserve the contents of the duotone information when reading
            ' and writing the file."
            m_Warnings.AddString "Color mode is duotone.  Color table will be ignored."
            Step2_RetrieveColorModeData = psd_Warning
        
        'Other color modes should not write a color table.  Advance the stream pointer accordingly,
        ' but also raise a warning.
        Else
            m_Warnings.AddString "Non-zero color mode segment exists, but image color mode isn't indexed or duotone."
            Step2_RetrieveColorModeData = psd_Warning
        End If
        
        'Forcibly set the stream pointer to the correct position
        m_Stream.SetPosition finalPosition, FILE_BEGIN
    
    'If the data length is zero, the stream already points at the next section.  Carry on!
    Else
        Step2_RetrieveColorModeData = psd_Success
    End If
    
    Exit Function

InternalVBError:
    InternalError "Step2_RetrieveColorModeData", "internal VB error #" & Err.Number & ": " & Err.Description
    If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 2, #" & Err.Number & ": " & Err.Description
    Step2_RetrieveColorModeData = psd_Failure

End Function

Private Function Step1_ValidateHeader(ByRef srcFile As String, Optional ByVal checkExtension As Boolean = False) As PD_PSDResult
    
    On Error GoTo InternalVBError
    
    'If the passed path is zero, assume the caller is loading the PSD from memory.
    If (LenB(srcFile) = 0) Then m_SourceFilename = PSD_LOADED_FROM_MEMORY Else m_SourceFilename = srcFile
    
    Dim okToProceed As PD_PSDResult
    okToProceed = psd_Success
    
    'We always check the file extension.  If the user has *asked* us to check it, we treat extension
    ' mismatches as a failure state.  (Otherwise, it will only raise a warning.)  This step is obviously
    ' skipped when a PSD is loaded directly from memory.
    If (m_SourceFilename <> PSD_LOADED_FROM_MEMORY) Then
        If (Strings.StringsNotEqual(Right$(m_SourceFilename, 3), "psd", True) And Strings.StringsNotEqual(Right$(m_SourceFilename, 3), "psb", True)) Then
            m_Warnings.AddString "File extension doesn't match PSD"
            If checkExtension Then okToProceed = psd_FileNotPSD Else okToProceed = psd_Warning
        End If
    End If
    
    'If all pre-checks passed, open a stream over the source data.
    If (okToProceed < psd_Failure) Then
        Set m_Stream = New pdStream
        If (m_SourceFilename = PSD_LOADED_FROM_MEMORY) Then
            If Not m_Stream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_SourcePtrLen, m_SourcePtr) Then
                m_Warnings.AddString "Couldn't start in-memory stream against passed pointer: " & m_SourcePtr
                okToProceed = psd_Failure
            End If
        Else
            If Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, m_SourceFilename, , , OptimizeSequentialAccess) Then
                m_Warnings.AddString "Can't read file; it may be locked or in an inaccessible location."
                okToProceed = psd_Failure
            End If
        End If
    End If
    
    'The stream is open.  Validate both the PSD's ASCII identifier and its version;
    ' if either fails, abandon ship.
    If (okToProceed < psd_Failure) Then
        
        Dim asciiID As String
        asciiID = m_Stream.ReadString_ASCII(4)
        
        Dim psdVersion As Integer
        psdVersion = m_Stream.ReadInt_BE()
        m_PSDisPSB = (psdVersion = 2)
        
        'Validate the asciiID; this value is immutable
        If (asciiID = "8BPS") Then
            
            'Version can only be 1 (for PSD) or 2 (for PSB)
            If (psdVersion = 1) Or (psdVersion = 2) Then okToProceed = psd_Success Else okToProceed = psd_Failure
            If (okToProceed < psd_Failure) Then
            
                'This appears to be a valid PSD file.  Parse the file header.
                With m_Header
                
                    .Version = psdVersion
                
                    'After the signature and version come 6 reserved bytes
                    m_Stream.SetPosition 6&, FILE_CURRENT
                
                    'Number of channels in the image, including alpha channels, comes next
                    .NumChannels = m_Stream.ReadInt_BE()
                
                    'Height/width come next (IN THAT ORDER, don't ask me why)
                    .ImageHeightPx = m_Stream.ReadLong_BE()
                    .ImageWidthPx = m_Stream.ReadLong_BE()
                
                    'Color depth
                    .BitsPerChannel = m_Stream.ReadInt_BE()
                    
                    'Color mode
                    .ColorMode = m_Stream.ReadInt_BE()
                
                    'Validate all header data; if any of it doesn't conform, reject the file.
                    ' (Validation values come from Adobe's spec, hence their arbitrary nature.)
                    If (.NumChannels < 1) Or (.NumChannels > 56) Then
                        m_Warnings.AddString "Unsupported number of channels: " & CStr(.NumChannels) & ". Load canceled."
                        okToProceed = psd_Failure
                    ElseIf (.ImageHeightPx <= 0) Or (.ImageWidthPx <= 0) Then
                        m_Warnings.AddString "Width and/or height is invalid (" & CStr(.ImageWidthPx) & "x" & CStr(.ImageHeightPx) & "). Load canceled."
                        okToProceed = psd_Failure
                    ElseIf (Not m_PSDisPSB) And ((.ImageHeightPx > 30000) Or (.ImageWidthPx > 30000)) Then
                        m_Warnings.AddString "PSD width and/or height is too large (" & CStr(.ImageWidthPx) & "x" & CStr(.ImageHeightPx) & "). Load canceled."
                        okToProceed = psd_Failure
                    ElseIf m_PSDisPSB And ((.ImageHeightPx > 300000) Or (.ImageWidthPx > 300000)) Then
                        m_Warnings.AddString "PSB width and/or height is too large (" & CStr(.ImageWidthPx) & "x" & CStr(.ImageHeightPx) & "). Load canceled."
                        okToProceed = psd_Failure
                    ElseIf ((.BitsPerChannel <> 1) And (.BitsPerChannel <> 8) And (.BitsPerChannel <> 16) And (.BitsPerChannel <> 32)) Then
                        m_Warnings.AddString "Invalid depth/bpc (" & CStr(.BitsPerChannel) & "). Load canceled."
                        okToProceed = psd_Failure
                    ElseIf ((.ColorMode < cm_Bitmap) Or (.ColorMode > cm_Lab)) Then
                        m_Warnings.AddString "Unknown color mode (" & CStr(.ColorMode) & "). Load canceled."
                        okToProceed = psd_Failure
                    End If
                    
                End With
                
            Else
                m_Warnings.AddString "PSD header provided unknown version #" & CStr(psdVersion) & ". Load canceled."
                okToProceed = psd_FileNotPSD
            End If
            
        Else
            m_Warnings.AddString "PSD header failed basic validation.  (This is not a PSD file.)"
            okToProceed = psd_FileNotPSD
        End If
        
    End If
    
    'Validation complete.  If the file validated successfully, this function guarantees that m_Stream
    ' points at the first byte PAST the valid PSD/PSB header.
    Step1_ValidateHeader = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError "IsFilePSD", "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 1, #" & Err.Number & ": " & Err.Description
    Step1_ValidateHeader = psd_Failure

End Function

Private Function GetColorModeName(ByVal srcMode As PSD_ColorMode) As String
    Select Case srcMode
        Case cm_Bitmap
            GetColorModeName = "Bitmap"
        Case cm_Grayscale
            GetColorModeName = "Grayscale"
        Case cm_Indexed
            GetColorModeName = "Indexed"
        Case cm_RGB
            GetColorModeName = "RGB"
        Case cm_CMYK
            GetColorModeName = "CMYK"
        Case cm_Multichannel
            GetColorModeName = "Multichannel"
        Case cm_Duotone
            GetColorModeName = "Duotone"
        Case cm_Lab
            GetColorModeName = "Lab"
    End Select
End Function

Friend Sub Reset()
    Set m_Profile = Nothing
    Set m_Warnings = New pdStringStack
    m_CompositeImageOffset = 0
End Sub

Private Sub Class_Initialize()
    Set m_Stream = New pdStream
    Me.Reset
End Sub

Private Sub Class_Terminate()
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen() Then m_Stream.StopStream True
    End If
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdPSD." & funcName & "() reported an error on file """ & m_SourceFilename & """: " & errDescription
    Else
        Debug.Print "pdPSD." & funcName & "() reported an error on file """ & m_SourceFilename & """: " & errDescription
    End If
End Sub

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
            PDDebug.LogAction "(" & CStr(i + 1) & ") WARNING: pdPSD reported: " & m_Warnings.GetString(i)
        Next i
    End If
End Sub
