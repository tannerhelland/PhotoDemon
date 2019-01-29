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
'Copyright 2019-2019 by Tanner Helland
'Created: 13/January/19
'Last updated: 29/January/19
'Last update: finalize support for modern 16-bit-per-channel PSDs (which store their image data in a
'             separate "Lr16" block
'
'This class (and its associated pdPSD- child classes) handle Adobe PSD parsing duties.  It is custom-built
' for PhotoDemon, with an emphasis on performance and proper color-management of all imported data.
'
'To my knowledge, this is the only PSD parser (besides Adobe's) that passes the full Apple color management
' test suite (https://developer.apple.com/library/archive/technotes/tn2115/_index.html).  This is a large
' point of pride for the author. ;)
'
'Currently, all color modes except LAB, and all color depths below 32-bits-per-channel are believed to be
' fully covered.  32-bit channels and LAB color mode remain a work-in-progress.
'
'At present, this class only handles PSD importing.  Exporting PSDs is still TODO.
'
'As with all 3rd-party PSD loaders, Photoshop has many features that don't have analogs in PhotoDemon.
' Such features are still extracted by this class, but they will not "appear" in the final loaded image.
' The ongoing goal is to expand support in this class as various PSD features are implemented in PD itself.
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PSD files contain a *lot* of information.  To aid debugging, you can activate "verbose" output; this will
' dump all kinds of diagnostic information to the debug log.  (Note that other PSD classes have their own
' version of this constant.)
Private Const PSD_DEBUG_VERBOSE As Boolean = True

'PSD loading is complicated, and a lot of things can go wrong.  Instead of returning binary "success/fail"
' values, we return specific flags; "warnings" may be recoverable and you can still attempt to load the file.
' "Failure" returns are unrecoverable and processing *must* be abandoned.  (As a convenience, you can treat
' the "warning" and "failure" values as flags; specific warning/failure states in each category will share
' the same high flag bit.)
'
'As I get deeper into this class, I may expand this enum to include more detailed states.
Public Enum PD_PSDResult
    psd_Success = 0
    psd_Warning = 256
    psd_Failure = 65536
    psd_FileNotPSD = 16777217
End Enum

#If False Then
    Private Const psd_Success = 0, psd_Warning = 256, psd_Failure = 65536, psd_FileNotPSD = 16777217
#End If

'PSDs support many different color modes.  Some color modes (e.g. Duotone) are proprietary and not
' publicly documented; we follow Adobe's recommendation and simply load the data as grayscale in that case.
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
    rid_IccProfile = 1039
    rid_IccUntagged = 1041      'Spec is unclear on what this means; I'd need to see it "in the wild" to know how to handle in PD
    rid_IndexedColorTableCount = 1046   'Number of used colors - important!
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

'Standard PSD header.  At load-time, only populated if the target PSD validates
Private Type PSD_Header
    Signature As String * 4
    version As Integer
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
    irName As String            'Pascal string, rarely relevant
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
    keepLoading = Step1_ValidateHeader(srcFile, checkExtension, True)
    
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
        
        'If we made it this far successfully, we have enough data to construct a pdImage object!
        ' Free the source stream object (as we no longer need it and it may have cached a bunch
        ' of memory), assemble the final image, then exit.
        If (keepLoading < psd_Failure) Then
            Set m_Stream = Nothing
            keepLoading = Step5_AssemblePDImage(srcFile, dstImage)
            If PSD_DEBUG_VERBOSE Then If (keepLoading < psd_Failure) Then PDDebug.LogAction "PSD parsing step 5 successful." Else PDDebug.LogAction "PSD parsing step 5 unsuccessful."
        End If
        
        LoadPSD = keepLoading
        
    End If
    
    'Current results may not be usable; always fall back to FreeImage's PSD support until
    ' this class matures.
    If (Me.Warnings_GetCount() > 0) Then
        PDDebug.LogAction "The PSD parser generated one or more warnings.  Here is a full list:"
        Me.Warnings_DumpToDebugger
    End If

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
    
    'Now we need to iterate through layers and convert their data from PSD format to PD format
    Dim i As Long
    For i = 0 To Abs(m_numOfLayers) - 1
        
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
                    tmpLayer.InitializeNewLayer PDL_IMAGE, newLayerName, m_Layers(i).GetMultiChannelLayerDIB(j)
                
                    'Fill in any remaining layer properties
                    With m_Layers(i)
                        
                        'Hardcode "darken" blend mode so that the resulting layers composite together
                        ' into a rough estimation of the original result
                        tmpLayer.SetLayerBlendMode BL_DARKEN
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
                tmpLayer.InitializeNewLayer PDL_IMAGE, m_Layers(i).GetLayerName(), m_Layers(i).GetLayerDIB()
                
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
            
        End If
        
    Next i
    
    'Make sure we were able to generate at least one usable layer.
    If (numUsableLayers <= 0) Or (dstImage.GetNumOfLayers <= 0) Then
        m_Warnings.AddString "unable to create any usable layers.  Abandoning load."
        Step5_AssemblePDImage = psd_Failure
    End If
    
    Exit Function

InternalVBError:
    InternalError "Step5_AssemblePDImage", "internal VB error #" & Err.Number & ": " & Err.Description
    m_Warnings.AddString "Internal error in step 5, #" & Err.Number & ": " & Err.Description
    Step5_AssemblePDImage = psd_Failure
    
End Function

Private Function Step4B_SingleLayerImage(ByVal chkPointerPosition As Long) As PD_PSDResult

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
        If (m_Layers(0).DecodeChannels(m_Stream, m_Warnings, m_PSDisPSB, m_Header.BitsPerChannel, m_Header.ColorMode, m_Profile, Me) < psd_Failure) Then
            Step4B_SingleLayerImage = m_Layers(0).ConstructImage(m_Stream, m_Warnings, m_PSDisPSB, m_Header.BitsPerChannel, m_Header.ColorMode, m_Profile, Me)
        Else
            InternalError "Step4B_SingleLayerImage", "pdPSDLayer.DecodeChannels() failed catastrophically"
            Step4B_SingleLayerImage = psd_Failure
            Exit Function
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
    If m_PSDisPSB Then lenOfSection = m_Stream.ReadLong_BE()        'TODO: fix PSB approach
    finalPointerPosition = m_Stream.GetPosition() + lenOfSection
    
    PDDebug.LogAction "Length of layer and mask section is: " & lenOfSection
    
    'If a PSD only contains one layer, it can be written at the end of the file (in the
    ' "image data" segment), bypassing this section completely.  To cover this special case,
    ' we hand off control to a separate function (as the data layout differs in subtle,
    ' obnoxious ways).
    If (lenOfSection <= 0) Then
        If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Single-layer PSD found.  Retrieving layer data from image segment instead..."
        Step4_GatherLayersAndMasks = Step4B_SingleLayerImage(m_Stream.GetPosition())
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
                
                'Close our original stream; it no longer serves any purpose
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
            
        End If
    
        'If we didn't find a high bit-depth segment, this is just a plain single-layer image.
        ' The single background layer's contents live in the composite image segment.
        If (Not hpdChunkFound) Then
            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Single-layer PSD found.  Activating alternate parse strategy..."
            Step4_GatherLayersAndMasks = Step4B_SingleLayerImage(finalPointerPosition)
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
            If (m_Layers(i).DecodeChannels(m_Stream, m_Warnings, m_PSDisPSB, m_Header.BitsPerChannel, m_Header.ColorMode, m_Profile, Me) < psd_Failure) Then
            
                'If the channel decoded successfully, convert it from a bare byte/int/float stream
                ' into usable pixel data. This comprehensive stage translates channel data to
                ' 8-bpc, converts the resulting channel collection from planar to 32-bpp interleaved
                ' RGBA, and color-manages the results.  The end result is a PD-compatible image buffer
                ' that can be directly associated with a pdLayer object.
                If (Not m_Layers(i).ConstructImage(m_Stream, m_Warnings, m_PSDisPSB, m_Header.BitsPerChannel, m_Header.ColorMode, m_Profile, Me) < psd_Failure) Then
                    InternalError "Step4_GatherLayersAndMasks", "pdPSDLayer.ConstructImage() failed catastrophically"
                    Step4_GatherLayersAndMasks = psd_Failure
                    Exit Function
                End If
                
            Else
                InternalError "Step4_GatherLayersAndMasks", "pdPSDLayer.DecodeChannels() failed catastrophically"
                Step4_GatherLayersAndMasks = psd_Failure
                Exit Function
            End If
        Next i
    
    'Layer info section is null-length; layer data has already been located using an alternate strategy
    End If
    
    'At the end of this step, we are finished with the source stream object.  Free it to conserve resources.
    m_Stream.StopStream True
    
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
                ' consists of two bytes of zero").
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
                        PDDebug.LogAction "Image resource: 0x" & Hex$(.irID) & ", " & .irDataLength & " bytes, name: " & .irName
                    End With
                End If
                
                'Before touching the actual resource data, calculate a final pointer position.
                ' This is important as...
                '  1) we may not actually retrieve resource data because it's useless to us, or
                '  2) we may retrieve only a portion of the data, and
                '  3) resource data is padded to even numbers, so the useful size of the data may
                '     not be the amount we actually need to move the pointer.
                '
                'By calculating a final position in advance, we can forcibly set the stream pointer to
                ' this value, regardless of what we do with the actual resource data.
                Dim finalResPosition As Long
                finalResPosition = m_Stream.GetPosition() + resSize
                If ((resSize And 1) = 1) Then finalResPosition = finalResPosition + 1
                
                'Retrieve the resource
                m_Stream.ReadBytes m_ImageResources(m_NumImageResources).irDataBytes, resSize, True
                
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
                                m_ResolutionInfo.riHRes = tmpStream.ReadFixed32_BE()
                                m_ResolutionInfo.riHResUnit = tmpStream.ReadInt_BE()
                                m_ResolutionInfo.riWidthUnit = tmpStream.ReadInt_BE()
                                m_ResolutionInfo.riVRes = tmpStream.ReadFixed32_BE()
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

Private Function Step1_ValidateHeader(ByRef srcFile As String, Optional ByVal checkExtension As Boolean = False, Optional ByVal cacheFileIfSuccessful As Boolean = False) As PD_PSDResult

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
        If (m_SourceFilename = PSD_LOADED_FROM_MEMORY) Then
            Set m_Stream = New pdStream
            If Not m_Stream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_SourcePtrLen, m_SourcePtr) Then
                m_Warnings.AddString "Couldn't start in-memory stream against passed pointer: " & m_SourcePtr
                okToProceed = psd_Failure
            End If
        Else
            Set m_Stream = New pdStream
            If Not m_Stream.StartStream(PD_SM_FileBacked, PD_SA_ReadOnly, m_SourceFilename, , , OptimizeSequentialAccess) Then
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
            
                'This appears to be a valid PSD file.  If streaming from the HDD, attempt to pre-cache
                ' smaller files in memory to improve load time.  (Obviously, we skip this step when
                ' loading a PSD directly from memory!)
                If (m_SourceFilename <> PSD_LOADED_FROM_MEMORY) And cacheFileIfSuccessful Then
                    
                    'Given the huge variance in potential PSD file sizes, I don't know of a good way to
                    ' "predict" the gains we might get from pre-loading the file (or if this will just
                    ' cause issues due to increased memory pressure).  In lieu of something more complex,
                    ' let's just test file size and use that as our guide - sufficiently small files
                    ' will then "always" pre-cache successfully, while large files get some measure of
                    ' safety against memory-related issues.
                    
                    'We currently use a threshold of 50 mb.  PSDs don't encode particularly efficiently,
                    ' so we're unlikely to encounter a very small PSD that expands to huge numbers
                    ' when expanded.  (And if a file's *that* big, 50mb of RAM won't make a difference.)
                    Const FILE_SIZE_TOO_BIG As Long = 50000000
                    If (Files.FileLenW(m_SourceFilename) < FILE_SIZE_TOO_BIG) Then
                        
                        Dim cachedStream As pdStream
                        Set cachedStream = New pdStream
                        
                        'If pre-caching fails, no problem; we'll just abandon the attempt and
                        ' continue using our original stream.
                        If Files.FileLoadAsPDStream(m_SourceFilename, cachedStream) Then
                            
                            'Replace the master stream object and advance its pointer by six bytes
                            ' (to bypass the fields we've already validated)
                            Set m_Stream = cachedStream
                            m_Stream.SetPosition 6&, FILE_BEGIN
                            If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "PSD looks good; caching source file in memory for improved performance..."
                        
                        End If
                    
                    '/File too big to pre-cache
                    End If
                
                '/PSD is being loaded from pointer, or caller explicitly asked us to *not* pre-cache
                End If
                
                'Confirm our stream is still attached to the PSD data and finish validating the header segment
                If (okToProceed < psd_Failure) Then
                    
                    With m_Header
                    
                        .version = psdVersion
                    
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
                
                End If
                
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
