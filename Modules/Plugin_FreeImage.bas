Attribute VB_Name = "Plugin_FreeImage"
'***************************************************************************
'FreeImage Interface (Advanced)
'Copyright 2012-2026 by Tanner Helland
'Created: 3/September/12
'Last updated: 07/October/21
'Last update: move output message callback here, and rewrite it to use internal PD objects for better perf
'
'This module represents a new - and significantly more comprehensive - approach to loading images via the
' FreeImage libary. It handles a variety of decisions on a per-format basis to ensure optimal load speed
' and quality.
'
'Please note that this module relies heavily on Carsten Klein's FreeImage wrapper for VB (included in this project
' as Outside_FreeImageV3; see that file for license details).  Thank you to Carsten for his work on translating
' myriad FreeImage declares into VB-compatible formats.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type

Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal nScan As Long, ByVal numScans As Long, ByVal ptrSrcBits As Long, ByRef BitsInfo As Any, ByVal wUsage As Long) As Long

'A single FreeImage declare is used here, to supply a callback for errors and other output messages
Private Declare Sub FreeImage_SetOutputMessage Lib "FreeImage" Alias "_FreeImage_SetOutputMessageStdCall@4" (ByVal pCallback As Long)

'DLL handle; if it is zero, FreeImage is not available
Private m_FreeImageHandle As Long

'Additional variables for PD-specific tone-mapping functions
Private m_shoulderStrength As Double, m_linearStrength As Double, m_linearAngle As Double, m_linearWhitePoint As Single
Private m_toeStrength As Double, m_toeNumerator As Double, m_toeDenominator As Double, m_toeAngle As Double

'Cache(s) for post-export image previews.  These objects can be safely freed, as they will be properly initialized on-demand.
Private m_ExportPreviewBytes() As Byte
Private m_ExportPreviewDIB As pdDIB

'FreeImage supports a callback for errors.  We store any returned strings in a normal pdStringStack, for convenience.
Private m_Errors As pdStringStack

'Initialize FreeImage.  The optional "actuallyLoadDLL" parameter allows you to just check if the DLL exists
' (pass FALSE) or fully initialize the library (pass TRUE).
Public Function InitializeFreeImage(Optional ByVal actuallyLoadDLL As Boolean = True) As Boolean
    
    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim fiPath As String
    fiPath = PluginManager.GetPluginPath & "FreeImage.dll"
    
    If actuallyLoadDLL Then
        
        'On successful initialization, PD supplies a callback to FreeImage for detailed error messages.
        If (m_FreeImageHandle = 0) Then
            PDDebug.LogAction "(Note: FreeImage is being loaded for the first time.)"
            m_FreeImageHandle = VBHacks.LoadLib(fiPath)
            If (m_FreeImageHandle <> 0) Then Plugin_FreeImage.InitializeFICallback
        End If
        
        InitializeFreeImage = (m_FreeImageHandle <> 0)
        
        If (Not InitializeFreeImage) Then
            FI_DebugMsg "WARNING!  LoadLibrary failed to load FreeImage.  Last DLL error: " & Err.LastDllError
            FI_DebugMsg "(FYI, the attempted path was: " & fiPath & ")"
        End If
        
    'At startup, we just do a quick check to ensure FreeImage exists - but we don't actually load it yet.
    ' (It will be loaded on-demand if required.)
    Else
        InitializeFreeImage = Files.FileExists(fiPath)
        If (Not InitializeFreeImage) Then FI_DebugMsg "WARNING!  FreeImage missing.  (FYI, the attempted path was: " & fiPath & ")"
    End If
    
End Function

Public Function ReleaseFreeImage() As Boolean
    If (m_FreeImageHandle <> 0) Then VBHacks.FreeLib m_FreeImageHandle
    ReleaseFreeImage = True
End Function

'Load a given file.  If successful, returns a non-zero FreeImage handle.
' Multi-page files will also fill a multipage DIB handle, which must also be freed post-load
' (in addition to the default handle returned by this function).
'
'On success, the target DIB object will also have its OriginalColorSpace member filled.
Private Function FI_LoadImageU(ByVal srcFilename As String, ByVal fileFIF As FREE_IMAGE_FORMAT, ByVal fi_ImportFlags As FREE_IMAGE_LOAD_OPTIONS, ByRef dstDIB As pdDIB, ByRef fi_multi_hDIB As Long, Optional ByVal pageToLoad As Long = 0&, Optional ByVal suppressDebugData As Boolean = False) As Long

    'FreeImage uses separate import behavior for single-page and multi-page files.
    ' As such, we may need to track multiple handles (e.g. a handle to the full image,
    ' and a handle to the current page).  If fi_multi_hDIB is non-zero, this is a multipage image.
    Dim fi_hDIB As Long
    If (pageToLoad <= 0) Then
        FI_DebugMsg "Invoking FreeImage_LoadUInt...", suppressDebugData
        fi_hDIB = FreeImage_LoadUInt(fileFIF, StrPtr(srcFilename), fi_ImportFlags)
    Else
        
        'Multipage support can be finicky, so it reports more debug info than PD usually prefers
        If (fileFIF = PDIF_GIF) Then
            FI_DebugMsg "Importing frame # " & CStr(pageToLoad + 1) & " from animated GIF file...", suppressDebugData
        ElseIf (fileFIF = FIF_ICO) Then
            FI_DebugMsg "Importing icon # " & CStr(pageToLoad + 1) & " from ICO file...", suppressDebugData
        Else
            FI_DebugMsg "Importing page # " & CStr(pageToLoad + 1) & " from multipage TIFF file...", suppressDebugData
        End If
        
        If (fileFIF = PDIF_GIF) Then
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(PDIF_GIF, srcFilename, fiFlags:=fi_ImportFlags Or FILO_GIF_PLAYBACK)
        ElseIf (fileFIF = FIF_ICO) Then
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_ICO, srcFilename, fiFlags:=fi_ImportFlags)
        Else
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(PDIF_TIFF, srcFilename, fiFlags:=fi_ImportFlags)
        End If
        
        fi_hDIB = FreeImage_LockPage(fi_multi_hDIB, pageToLoad)
        
    End If
    
    FI_LoadImageU = fi_hDIB
    
End Function

'Load an image via FreeImage.  It is assumed that the source file has already been vetted for things like "does it exist?"
Public Function FI_LoadImage_V5(ByVal srcFilename As String, ByRef dstDIB As pdDIB, Optional ByVal pageToLoad As Long = 0, Optional ByVal showMessages As Boolean = True, Optional ByRef targetImage As pdImage = Nothing, Optional ByVal suppressDebugData As Boolean = False) As PD_OPERATION_OUTCOME
    
    On Error GoTo FreeImageV5_Error
    
    '****************************************************************************
    ' Make sure FreeImage exists and is usable
    '****************************************************************************
    
    'PD now initializes FreeImage "on-demand", so it won't be available until the first user request for it.
    ' Ensure successful initialization before proceeding.
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    
    If (Not ImageFormats.IsFreeImageEnabled()) Then
        FI_LoadImage_V5 = PD_FAILURE_GENERIC
        Exit Function
    End If
    
    '****************************************************************************
    ' Determine image format
    '****************************************************************************
    
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    FI_DebugMsg "Running filetype heuristics...", suppressDebugData
    
    Dim fileFIF As FREE_IMAGE_FORMAT
    fileFIF = FI_DetermineFiletype(srcFilename, dstDIB)
    
    'If FreeImage doesn't recognize the filetype, abandon the import attempt.
    If (fileFIF = FIF_UNKNOWN) Then
        PDDebug.LogAction "Filetype not supported by FreeImage.  Import abandoned."
        FI_LoadImage_V5 = PD_FAILURE_GENERIC
        Exit Function
    End If
    
    '****************************************************************************
    ' Prepare load flags for this file format
    '****************************************************************************
    
    FI_DebugMsg "Preparing FreeImage import flags...", suppressDebugData
    
    Dim fi_ImportFlags As FREE_IMAGE_LOAD_OPTIONS
    fi_ImportFlags = FI_DetermineImportFlags(srcFilename, fileFIF, Not showMessages)
    
    '****************************************************************************
    ' Load the image into a FreeImage container
    '****************************************************************************
    
    'FreeImage uses separate import behavior for single-page and multi-page files.
    ' As such, we may need to track multiple handles (e.g. a handle to the full image,
    ' and a different handle to the current page).  If fi_multi_hDIB is non-zero,
    ' this is a multipage image.
    Dim fi_hDIB As Long, fi_multi_hDIB As Long
    fi_hDIB = FI_LoadImageU(srcFilename, fileFIF, fi_ImportFlags, dstDIB, fi_multi_hDIB, pageToLoad, suppressDebugData)
    
    'If an empty handle is returned, abandon the import attempt.
    If (fi_hDIB = 0) Then
        FI_DebugMsg "Import via FreeImage failed (blank handle).", suppressDebugData
        FI_LoadImage_V5 = PD_FAILURE_GENERIC
        Exit Function
    End If
    
    '****************************************************************************
    ' Retrieve generic metadata, like color depth and resolution (if available)
    '****************************************************************************
    
    Dim fi_BPP As Long, fi_DataType As FREE_IMAGE_TYPE
    fi_BPP = FreeImage_GetBPP(fi_hDIB)
    fi_DataType = FreeImage_GetImageType(fi_hDIB)
    FI_DebugMsg "Heuristics show image bit-depth: " & fi_BPP & ", pixel type: " & FI_GetImageTypeAsString(fi_DataType), suppressDebugData
    
    dstDIB.SetDPI FreeImage_GetResolutionX(fi_hDIB), FreeImage_GetResolutionY(fi_hDIB)
    dstDIB.SetOriginalColorDepth FreeImage_GetBPP(fi_hDIB)
    
    '****************************************************************************
    ' Retrieve any attached ICC profiles
    '****************************************************************************
    
    'If FreeImage detects a color profile in the image, we want to do several things:
    ' 1) Pull the raw ICC profile bytes into a pdICCprofile object
    ' 2) Add the ICC data to our central ICC profile cache for this session.
    ' 3) As of v7.0, we also want to perform a hard-convert to sRGB, and flag the target DIB
    '    accordingly.  (In the future, we may just tag the DIB against it's existing space.)
    Dim srcColorProfile As pdICCProfile, colorProfileHash As String
    If FreeImage_HasICCProfile(fi_hDIB) Then
        If FI_LoadICCProfile(fi_hDIB, srcColorProfile) Then
            
            'Add the retrieved profile to PD's central cache, and tag the destination image (if any)
            ' to note that this profile is the image's original color space.
            colorProfileHash = ColorManagement.AddProfileToCache(srcColorProfile)
            If (Not targetImage Is Nothing) Then targetImage.SetColorProfile_Original colorProfileHash
            
        End If
    End If
    
    '****************************************************************************
    ' If the image has a palette, retrieve it
    '****************************************************************************
    
    'As of 7.0, we cache 8-bit palettes inside their destination image; this palette can be
    ' exported (via File > Export) or re-used by certain tools.
    If (Not targetImage Is Nothing) And (fi_BPP <= 8) And (fi_DataType = FIT_BITMAP) Then
        Dim srcPalette() As RGBQuad, numOfColors As Long
        If Outside_FreeImageV3.FreeImage_GetPalette_ByTanner(fi_hDIB, srcPalette, numOfColors) Then
            targetImage.SetOriginalPalette srcPalette, numOfColors
            FI_DebugMsg "Image palette cached locally (" & numOfColors & " colors)"
        End If
    End If
    
    '****************************************************************************
    ' Determine color vs grayscale status
    '****************************************************************************
    
    If (Not targetImage Is Nothing) Then
        
        Dim imgLikelyGrayscale As Boolean: imgLikelyGrayscale = False
        
        If (fi_DataType = FIT_BITMAP) Then
            imgLikelyGrayscale = (FreeImage_GetColorType(fi_hDIB) = FIC_MINISBLACK) Or (FreeImage_GetColorType(fi_hDIB) = FIC_MINISWHITE)
        Else
            imgLikelyGrayscale = (fi_DataType = FIT_UINT16) Or (fi_DataType = FIT_INT16) Or (fi_DataType = FIT_UINT32) Or (fi_DataType = FIT_INT32) Or (fi_DataType = FIT_FLOAT) Or (fi_DataType = FIT_DOUBLE)
        End If
        
        targetImage.SetOriginalGrayscale imgLikelyGrayscale
        
    End If
    
    '****************************************************************************
    ' Retrieve alpha channel presence, if any
    '****************************************************************************
    
    If (Not targetImage Is Nothing) Then targetImage.SetOriginalAlpha FreeImage_IsTransparent(fi_hDIB)
    
    '****************************************************************************
    ' Copy/transform the FreeImage object into the destination pdDIB object
    '****************************************************************************
    
    'Converting any arbitrary chunk of image bytes into a valid 24- or 32-bpp image is a non-trivial task.
    ' As such, we split this specialized handling into its own function.
    
    '(Also, I know it seems weird, but the target function needs to run some heuristics on the incoming data to
    ' see if it came from the Windows clipboard.  If it did, we have to apply some special post-processing to
    ' the pixel data to compensate for some weird GDI interop quirks.)
    
    'Note that the result of this transformation *will* be hard-converted to sRGB, if the source image has
    ' a color profile associated with it.  (If it does *not* have an attached profile, the destination DIB
    ' will be marked as "untagged".)
    Dim specialClipboardHandlingRequired As Boolean
    
    '(NOTE: this call may raise additional dialogs, like tone-mapping for untagged HDR images.)
    FI_LoadImage_V5 = FI_GetFIObjectIntoDIB(fi_hDIB, fi_multi_hDIB, fileFIF, fi_DataType, specialClipboardHandlingRequired, srcFilename, dstDIB, pageToLoad, showMessages, targetImage, suppressDebugData)
    If (FI_LoadImage_V5 = PD_SUCCESS) Then
    
        'The FI data now exists inside a pdDIB object, at 24- or 32-bpp.
        
        '****************************************************************************
        ' Release all remaining FreeImage-specific structures and links
        '****************************************************************************
        
        FI_Unload fi_hDIB, fi_multi_hDIB
        FI_DebugMsg "Image load successful.  FreeImage handle released.", suppressDebugData
        
        '****************************************************************************
        ' Finalize alpha values in the target image
        '****************************************************************************
        
        'If this image came from the clipboard, and its alpha state is unknown, we're going to force all alpha values
        ' to 255 to avoid potential driver-specific issues with the PrtScrn key.
        If specialClipboardHandlingRequired Then
            FI_DebugMsg "Image came from the clipboard; finalizing alpha now...", suppressDebugData
            dstDIB.ForceNewAlpha 255
        End If
        
        'Regardless of original bit-depth, the final PhotoDemon image will always be 32-bits, with pre-multiplied alpha.
        dstDIB.SetInitialAlphaPremultiplicationState True
        
        '****************************************************************************
        ' Load complete
        '****************************************************************************
        
        'Confirm this load as successful
        FI_LoadImage_V5 = PD_SUCCESS
    
    'If the source function failed, there's nothing we can do here; the incorrect error code will have already been set,
    ' so we can simply bail.
    End If
    
    'ERROR HANDLING CODE ONLY BEYOND THIS FUNCTION EXIT
    Exit Function
    
    '****************************************************************************
    ' Error handling
    '****************************************************************************
    
FreeImageV5_Error:
    
    FI_DebugMsg "VB-specific error occurred inside FI_LoadImage_V5.  Err #: " & Err.Number & ", " & Err.Description, suppressDebugData
    If showMessages Then Message "Image import failed"
    FI_Unload fi_hDIB, fi_multi_hDIB
    FI_LoadImage_V5 = PD_FAILURE_GENERIC
    
End Function

'Given a valid handle to a FreeImage object (and/or multipage object, as relevant), get the FreeImage object
' into a pdDIB object.  While this sounds simple, it really isn't, primarily because we have to deal with so
' many possible color depths, models (gray or color), alpha-channel encodings, ICC profile behaviors, etc.
'
'Note: this function may raise modal UI dialogs.
'
'RETURNS: PD_SUCCESS if successful; some other code if the load fails.  Review debug messages for additional info.
Private Function FI_GetFIObjectIntoDIB(ByRef fi_hDIB As Long, ByRef fi_multi_hDIB As Long, ByVal fileFIF As FREE_IMAGE_FORMAT, ByVal fi_DataType As FREE_IMAGE_TYPE, ByRef specialClipboardHandlingRequired As Boolean, ByVal srcFilename As String, ByRef dstDIB As pdDIB, Optional ByVal pageToLoad As Long = 0, Optional ByVal showMessages As Boolean = True, Optional ByRef targetImage As pdImage = Nothing, Optional ByVal suppressDebugData As Boolean = False, Optional ByRef multiDibIsDetached As Boolean = False) As PD_OPERATION_OUTCOME
    
    On Error GoTo FiObject_Error
    
    '****************************************************************************
    ' If the image is in an unsupported format, convert it to standard 24 or 32-bpp RGBA
    '****************************************************************************
    
    'As much as possible, we prefer to convert bit-depth using the existing FreeImage handle as the source,
    ' and the target pdDIB object as the destination.  This lets us skip redundant allocations for temporary
    ' FreeImage objects.
    '
    'If the image has successfully been moved into the target pdDIB object, this tracker *must* be set to TRUE.
    ' (Otherwise, a failsafe check at the end of this function will perform an auto-copy.)
    Dim dstDIBFinished As Boolean: dstDIBFinished = False
    
    'When working with a multipage image, we may need to "detach" the current page DIB from its parent multipage handle.
    ' (This happens if an intermediate copy of the FI object is required.)
    '
    'If we detach an individual page DIB from it parent, this variable will note it, so we know to use the standalone
    ' unload function before exiting (instead of the multipage-specific one).
    multiDibIsDetached = False
    
    'Intermediate FreeImage objects may also be required during the transform process
    Dim new_hDIB As Long
    
    'Before proceeding, cache any ICC profiles.  (The original FreeImage handle may be freed as part of
    ' moving image data between color spaces, and we don't want to accidentally lose its color profile.)
    Dim srcIccProfile As pdICCProfile, profileOK As Boolean
    If (fi_hDIB <> 0) Then
        If FreeImage_HasICCProfile(fi_hDIB) Then profileOK = FI_LoadICCProfile(fi_hDIB, srcIccProfile)
    End If
    
    '****************************************************************************
    ' CMYK images are handled first (they require special treatment)
    '****************************************************************************
    
    'Note that all "continue loading" checks start with "If (Not dstDIBFinished)".  When various conditions are met,
    ' this function may attempt to shortcut the load process.  If this occurs, "dstDIBFinished" will be set to TRUE,
    ' allowing subsequent checks to be skipped.
    If (Not dstDIBFinished) And (FreeImage_GetColorType(fi_hDIB) = FIC_CMYK) Then
        
        FI_DebugMsg "CMYK image detected.  Preparing transform into RGB space...", suppressDebugData
        
        'Proper CMYK conversions require an ICC profile.  If this image doesn't have one, it's a pointless image
        ' (it's impossible to construct a "correct" copy since CMYK is device-specific), but we'll of course try
        ' to load it anyway.
        Dim cmykConversionSuccessful As Boolean: cmykConversionSuccessful = False
        If FreeImage_HasICCProfile(fi_hDIB) Then cmykConversionSuccessful = ConvertCMYKFiDIBToRGB(fi_hDIB, dstDIB)
        
        'If the ICC transform worked, free the FreeImage handle and note that the destination image is ready to go!
        If cmykConversionSuccessful Then
            FI_Unload fi_hDIB, fi_multi_hDIB, True, multiDibIsDetached
            dstDIBFinished = True
        
        'If CMYK conversion failed, re-load the image and use FreeImage to apply a generic CMYK -> RGB transform.
        Else
            FI_DebugMsg "ICC-based CMYK transformation failed.  Falling back to default CMYK conversion...", suppressDebugData
            FI_Unload fi_hDIB, fi_multi_hDIB
            fi_hDIB = FreeImage_LoadUInt(fileFIF, StrPtr(srcFilename), FILO_JPEG_ACCURATE Or FILO_JPEG_EXIFROTATE)
        End If
        
    End If
    
    'Between attempted conversions, we typically reset the BPP tracker (as it may have changed due to internal
    ' FreeImage handling)
    Dim fi_BPP As Long
    If (fi_hDIB <> 0) Then fi_BPP = FreeImage_GetBPP(fi_hDIB)
    
    '****************************************************************************
    ' With CMYK images out of the way, deal with high bit-depth images in normal color spaces
    '****************************************************************************
    
    'FIT_BITMAP refers to any image with channel data <= 8 bits per channel.  We want to check for images that do *not*
    ' fit this definition, e.g. images ranging from 16-bpp grayscale images to 128-bpp RGBA images.
    If (Not dstDIBFinished) And (fi_DataType <> FIT_BITMAP) Then
        
        'We have two mechanisms for downsampling a high bit-depth image:
        ' 1) Using an embedded ICC profile (the preferred mechanism)
        ' 2) Using a generic tone-mapping algorithm to estimate conversion parameters
        '
        'If at all possible, we will try to use (1) before (2).  Success is noted by the following variable.
        Dim hdrICCSuccess As Boolean: hdrICCSuccess = False
        
        'If an ICC profile exists, attempt to use it
        If FreeImage_HasICCProfile(fi_hDIB) And ColorManagement.UseEmbeddedICCProfiles() Then
            
            FI_DebugMsg "HDR image identified.  ICC profile found; attempting to convert automatically...", suppressDebugData
            hdrICCSuccess = GenerateICCCorrectedFIDIB(fi_hDIB, dstDIB, dstDIBFinished, new_hDIB)
            
            'Some esoteric color-depths may require us to use a temporary FreeImage handle instead of copying
            ' the color-managed result directly into a pdDIB object.
            If hdrICCSuccess Then
                
                dstDIB.SetColorProfileHash ColorManagement.GetSRGBProfileHash()
                dstDIB.SetColorManagementState cms_ProfileConverted
                
                If (Not dstDIBFinished) And (new_hDIB <> 0) Then
                    FI_Unload fi_hDIB, fi_multi_hDIB, True, multiDibIsDetached
                    fi_hDIB = new_hDIB
                    new_hDIB = 0
                End If
                
            Else
                FI_DebugMsg "ICC transformation unsuccessful; dropping back to tone-mapping...", suppressDebugData
            End If
        
        End If
        
        'If we can't find an ICC profile, we have no choice but to use tone-mapping to generate a 24/32-bpp image
        If (Not hdrICCSuccess) Then
        
            FI_DebugMsg "HDR image identified.  Raising tone-map dialog...", suppressDebugData
            
            'Use the central tone-map handler to apply further tone-mapping
            Dim toneMappingOutcome As PD_OPERATION_OUTCOME
            toneMappingOutcome = RaiseToneMapDialog(fi_hDIB, new_hDIB, (Not showMessages) Or (Macros.GetMacroStatus = MacroBATCH))
            
            'A non-zero return signifies a successful tone-map operation.  Unload our old handle, and proceed with the new handle
            If (toneMappingOutcome = PD_SUCCESS) And (new_hDIB <> 0) Then
                
                'Add a note to the target image that tone-mapping was forcibly applied to the incoming data
                If (Not targetImage Is Nothing) Then targetImage.ImgStorage.AddEntry "Tone-mapping", True
                
                'Replace the primary FI_DIB handle with the new one, then carry on with loading
                If (new_hDIB <> fi_hDIB) Then FI_Unload fi_hDIB, fi_multi_hDIB, True, multiDibIsDetached
                fi_hDIB = new_hDIB
                FI_DebugMsg "Tone mapping complete.", suppressDebugData
                
            'The tone-mapper will return 0 if it failed.  If this happens, we cannot proceed with loading.
            Else
                FI_Unload fi_hDIB, fi_multi_hDIB, True, multiDibIsDetached
                If (toneMappingOutcome <> PD_SUCCESS) Then FI_GetFIObjectIntoDIB = toneMappingOutcome Else FI_GetFIObjectIntoDIB = PD_FAILURE_GENERIC
                FI_DebugMsg "Tone-mapping canceled due to user request or error.  Abandoning image import.", suppressDebugData
                Exit Function
            End If
            
        End If
    
    End If
    
    'Between attempted conversions, we reset the BPP tracker (as it may have changed)
    If (fi_hDIB <> 0) Then fi_BPP = FreeImage_GetBPP(fi_hDIB)
    
    
    '****************************************************************************
    ' If the image is < 32bpp, upsample it to 32bpp
    '****************************************************************************
    
    'The source image should now be in one of two bit-depths:
    ' 1) 32-bpp RGBA
    ' 2) Some bit-depth less than 32-bpp RGBA
    '
    'In the second case, we want to upsample the data to 32-bpp RGBA, adding an opaque alpha channel as necessary.
    ' (In the past, this block only triggered if the BPP was below 24, but I'm now relying on FreeImage to apply
    '  any necessary 24- to 32-bpp conversions as well.)
    If (Not dstDIBFinished) And (fi_BPP < 32) Then
        
        'If the image is grayscale, and it has an ICC profile, we need to apply that prior to continuing.
        ' (Grayscale images have grayscale ICC profiles which the default ICC profile handler can't address.)
        If (fi_BPP = 8) And FreeImage_HasICCProfile(fi_hDIB) And ColorManagement.UseEmbeddedICCProfiles() Then
            
            'In the future, 8-bpp RGB/A conversion could be handled here.
            ' (Note that you need to up-sample the source image prior to conversion, however, as LittleCMS doesn't work with palettes.)
            
            'At present, we only cover grayscale ICC profiles in indexed images
            If ((FreeImage_GetColorType(fi_hDIB) = FIC_MINISBLACK) Or (FreeImage_GetColorType(fi_hDIB) = FIC_MINISWHITE)) Then
                
                FI_DebugMsg "8bpp grayscale image with ICC profile identified.  Applying color management now...", suppressDebugData
                new_hDIB = 0
                
                If GenerateICCCorrectedFIDIB(fi_hDIB, dstDIB, dstDIBFinished, new_hDIB) Then
                    If (Not dstDIBFinished) And (new_hDIB <> 0) Then
                        FI_Unload fi_hDIB, fi_multi_hDIB, True, multiDibIsDetached
                        fi_hDIB = new_hDIB
                        new_hDIB = 0
                    End If
                End If
            End If
            
        End If
        
        If (Not dstDIBFinished) Then
        
            'In the past, we would check for an alpha channel here (something like "fi_hasTransparency = FreeImage_IsTransparent(fi_hDIB)"),
            ' but that is no longer necessary.  We instead rely on FreeImage to convert to 32-bpp regardless of transparency status.
            new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, False)
            
            If (new_hDIB <> fi_hDIB) Then
                FI_Unload fi_hDIB, fi_multi_hDIB, True, multiDibIsDetached
                fi_hDIB = new_hDIB
            Else
                PDDebug.LogAction "WARNING!  FI_GetFIObjectIntoDIB failed to convert a color depth!"
            End If
            
        End If
            
    End If
    
    'By this point, we have loaded the image, and it is guaranteed to be at 32 bit color depth.  Verify it one final time.
    If (fi_hDIB <> 0) Then fi_BPP = FreeImage_GetBPP(fi_hDIB)
    
    
    '****************************************************************************
    ' If the image has an ICC profile but we haven't yet applied it, do so now.
    '****************************************************************************
    
    If (Not dstDIBFinished) And (dstDIB.GetColorManagementState = cms_NoManagement) And profileOK Then
        
        If (Not srcIccProfile Is Nothing) And ColorManagement.UseEmbeddedICCProfiles() Then
        
            FI_DebugMsg "Applying final color management operation...", suppressDebugData
            
            new_hDIB = 0
            If GenerateICCCorrectedFIDIB(fi_hDIB, dstDIB, dstDIBFinished, new_hDIB, srcIccProfile) Then
                If (Not dstDIBFinished) And (new_hDIB <> 0) Then
                    FI_Unload fi_hDIB, fi_multi_hDIB, True, multiDibIsDetached
                    fi_hDIB = new_hDIB
                    new_hDIB = 0
                End If
            End If
            
        End If
        
    End If
    
    'Between attempted conversions, we reset the BPP tracker (as it may have changed)
    If (fi_hDIB <> 0) Then fi_BPP = FreeImage_GetBPP(fi_hDIB)
    
    
    '****************************************************************************
    ' PD's current rendering engine requires pre-multiplied alpha values.  Apply premultiplication now - but ONLY if
    ' the image did not come from the clipboard.  (Clipboard images requires special treatment.)
    '****************************************************************************
    
    Dim tmpClipboardInfo As PD_Clipboard_Info
    specialClipboardHandlingRequired = False
    
    If (Not dstDIBFinished) And (fi_BPP = 32) Then
        
        'If the clipboard is active, this image came from a Paste operation.  It may require extra alpha heuristics.
        If g_Clipboard.IsClipboardOpen Then
        
            'Retrieve a local copy of PD's clipboard info struct.  We're going to analyze it, to see if we need to
            ' run some alpha heuristics (because the clipboard is shit when it comes to handling alpha correctly.)
            tmpClipboardInfo = g_Clipboard.GetClipboardInfo
            
            'If the clipboard image was originally placed on the clipboard as a DDB, a whole variety of driver-specific
            ' issues may be present.
            If (tmpClipboardInfo.pdci_OriginalFormat = CF_BITMAP) Then
            
                'Well, this sucks.  The original owner of this clipboard data (maybe even Windows itself, in the case
                ' of PrtScrn) placed an image on the clipboard in the ancient CF_BITMAP format, which is a DDB with
                ' device-specific coloring.  In the age of 24/32-bit displays, we don't care about color issues so
                ' much, but alpha is whole other mess.  For performance reasons, most display drivers run in 32-bpp
                ' mode, with the alpha values typically ignored.  Unfortunately, some drivers (*cough* INTEL *cough*)
                ' may leave junk in the 4th bytes instead of wiping them clean, preventing us from easily telling
                ' if the source data has alpha values filled intentionally, or by accident.
                
                'Because there is no foolproof way to know if the alpha values are valid, we should probably prompt
                ' the user for feedback on how to proceed.  For now, however, simply wipe the alpha bytes of anything
                ' placed on the clipboard in CF_BITMAP format.
                
                '(The image is still in FreeImage format at this point, so we set a flag and will apply the actual
                ' alpha transform later.)
                specialClipboardHandlingRequired = True
            
            'The image was originally placed on the clipboard as a DIB.  Assume the caller knew what they were doing
            ' with their own alpha bytes, and apply premultiplication now.
            Else
                FreeImage_PreMultiplyWithAlpha fi_hDIB
            End If
        
        'This is a normal image - carry on!
        Else
            FreeImage_PreMultiplyWithAlpha fi_hDIB
        End If
        
    End If
    
    
    '****************************************************************************
    ' Copy the data from the FreeImage object to the target pdDIB object
    '****************************************************************************
    
    'Note that certain code paths may have already populated the pdDIB object.  We only need to perform this step if the image
    ' data still resides inside a FreeImage handle.
    If (Not dstDIBFinished) And (fi_hDIB <> 0) Then
        
        'Get width and height from the file, and create a new DIB to match
        Dim fi_Width As Long, fi_Height As Long
        fi_Width = FreeImage_GetWidth(fi_hDIB)
        fi_Height = FreeImage_GetHeight(fi_hDIB)
        
        'Update Dec '12: certain faulty TIFF files can confuse FreeImage and cause it to report wildly bizarre height and width
        ' values; check for this, and if it happens, abandon the load immediately.  (This is not ideal, because it leaks memory
        ' - but it prevents a hard program crash, so I consider it the lesser of two evils.)
        If (fi_Width > 1000000) Or (fi_Height > 1000000) Then
            FI_GetFIObjectIntoDIB = PD_FAILURE_GENERIC
            Exit Function
        Else
        
            'Our caller may be reusing the same image across multiple loads.  To improve performance, only create a new
            ' object if necessary; otherwise, reuse the previous instance.
            Dim dibReady As Boolean
            If (dstDIB.GetDIBWidth = fi_Width) And (dstDIB.GetDIBHeight = fi_Height) And (dstDIB.GetDIBColorDepth = fi_BPP) Then
                dstDIB.ResetDIB 0
                dibReady = True
            Else
                FI_DebugMsg "Requesting memory for final image transfer...", suppressDebugData
                dibReady = dstDIB.CreateBlank(fi_Width, fi_Height, fi_BPP, 0, 0)
                If dibReady Then FI_DebugMsg "Memory secured.  Finalizing image load...", suppressDebugData
            End If
            
            If dibReady Then
                SetDIBitsToDevice dstDIB.GetDIBDC, 0, 0, fi_Width, fi_Height, 0, 0, 0, fi_Height, FreeImage_GetBits(fi_hDIB), ByVal FreeImage_GetInfo(fi_hDIB), 0&
            Else
                FI_DebugMsg "Import via FreeImage failed (couldn't create DIB).", suppressDebugData
                FI_Unload fi_hDIB, fi_multi_hDIB, True, multiDibIsDetached
                FI_GetFIObjectIntoDIB = PD_FAILURE_GENERIC
                Exit Function
            End If
        End If
        
    End If
    
    'If we made it all the way here, we have successfully moved the original FreeImage object into the destination pdDIB object.
    FI_GetFIObjectIntoDIB = PD_SUCCESS
    
    Exit Function
    
FiObject_Error:
    
    FI_DebugMsg "VB-specific error occurred inside FI_GetFIObjectIntoDIB.  Err #: " & Err.Number & ", " & Err.Description, suppressDebugData
    If showMessages Then Message "Image import failed"
    FI_Unload fi_hDIB, fi_multi_hDIB, True, multiDibIsDetached
    FI_GetFIObjectIntoDIB = PD_FAILURE_GENERIC
    
End Function

'After the first page of a multipage image has been loaded successfully, call this function to load the remaining pages into the
' destination object.
Public Function FinishLoadingMultipageImage(ByRef srcFilename As String, ByRef dstDIB As pdDIB, Optional ByVal numOfPages As Long = 0, Optional ByVal showMessages As Boolean = True, Optional ByRef targetImage As pdImage = Nothing, Optional ByVal suppressDebugData As Boolean = False, Optional ByVal suggestedFilename As String = vbNullString) As PD_OPERATION_OUTCOME

    'Ensure library is available before proceeding
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    'Get a multipage handle to the source file
    Dim fileFIF As FREE_IMAGE_FORMAT
    fileFIF = FI_DetermineFiletype(srcFilename, dstDIB)
    
    Dim fi_ImportFlags As FREE_IMAGE_LOAD_OPTIONS
    fi_ImportFlags = FI_DetermineImportFlags(srcFilename, fileFIF, Not showMessages)
    
    Dim fi_hDIB As Long, fi_multi_hDIB As Long
    If (fileFIF = PDIF_GIF) Then
        fi_multi_hDIB = FreeImage_OpenMultiBitmap(PDIF_GIF, srcFilename, fiFlags:=fi_ImportFlags Or FILO_GIF_PLAYBACK)
    ElseIf (fileFIF = FIF_ICO) Then
        fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_ICO, srcFilename, fiFlags:=fi_ImportFlags)
    Else
        fi_multi_hDIB = FreeImage_OpenMultiBitmap(PDIF_TIFF, srcFilename, fiFlags:=fi_ImportFlags)
    End If
    
    'We are now going to keep that source file open for the duration of the load process.
    Dim fi_BPP As Long, fi_DataType As FREE_IMAGE_TYPE
    Dim specialClipboardHandlingRequired As Boolean, loadSuccess As Boolean
    Dim newLayerID As Long, newLayerName As String
    Dim multiDibIsDetached As Boolean
    
    'Start iterating pages!
    Dim pageToLoad As Long
    For pageToLoad = 1 To numOfPages - 1
        
        Message "Loading page %1 of %2...", CStr(pageToLoad + 1), numOfPages, "DONOTLOG"
        If ((pageToLoad And 7) = 0) Then
            VBHacks.PurgeInputMessages FormMain.hWnd
            VBHacks.DoEvents_SingleHwnd FormMain.hWnd
        End If
        
        'Lock the current page
        fi_hDIB = FreeImage_LockPage(fi_multi_hDIB, pageToLoad)
        If (fi_hDIB <> 0) Then
            
            'Store various bits of file metadata before proceeding
            fi_BPP = FreeImage_GetBPP(fi_hDIB)
            fi_DataType = FreeImage_GetImageType(fi_hDIB)
            dstDIB.SetDPI FreeImage_GetResolutionX(fi_hDIB), FreeImage_GetResolutionY(fi_hDIB)
            dstDIB.SetOriginalColorDepth FreeImage_GetBPP(fi_hDIB)
            
            'Retrieve a matching ICC profile, if any, and add it to the central cache
            Dim tmpProfile As pdICCProfile, profHash As String
            If FreeImage_HasICCProfile(fi_hDIB) Then
                FI_LoadICCProfile fi_hDIB, tmpProfile
                profHash = ColorManagement.AddProfileToCache(tmpProfile)
                dstDIB.SetColorProfileHash profHash
            End If
            
            'Copy/transform the FreeImage object into a guaranteed 24- or 32-bpp destination DIB
            specialClipboardHandlingRequired = False
            loadSuccess = (FI_GetFIObjectIntoDIB(fi_hDIB, fi_multi_hDIB, fileFIF, fi_DataType, specialClipboardHandlingRequired, srcFilename, dstDIB, pageToLoad, showMessages, targetImage, suppressDebugData, multiDibIsDetached) = PD_SUCCESS)
            
            'Regardless of outcome, free ("unlock" in FI parlance) FreeImage's copy of this page
            FI_Unload fi_hDIB, fi_multi_hDIB, True, multiDibIsDetached
            
            If loadSuccess Then
            
                'Make sure the DIB meets new v7.0 requirements (including premultiplied alpha)
                If specialClipboardHandlingRequired Then dstDIB.ForceNewAlpha 255
                dstDIB.SetInitialAlphaPremultiplicationState True
                ImageImporter.ForceTo32bppMode dstDIB
                
                'Create a blank layer in the receiving image, and retrieve a pointer to it
                newLayerID = targetImage.CreateBlankLayer
                newLayerName = Layers.GenerateInitialLayerName(srcFilename, suggestedFilename, True, targetImage, dstDIB, pageToLoad)
                targetImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, newLayerName, dstDIB, True
                
            End If
            
        Else
            PDDebug.LogAction "WARNING!  Failed to lock page #" & pageToLoad
        End If
    
    Next pageToLoad
    
    'Release our original multipage image handle, then exit
    FI_Unload fi_hDIB, fi_multi_hDIB
    FI_DebugMsg "Multipage image load successful.  Original FreeImage handle released.", suppressDebugData
    
    FinishLoadingMultipageImage = PD_SUCCESS

End Function

'Given a path to a file and a destination pdDIB object, detect the file's type and store it inside the target DIB.
' (Knowing the source of a DIB allows us to run better heuristics on various image attributes.)
' On success, returns the detected FIF; on failure, returns FIF_UNKNOWN.  Note that the dstDIB's format may vary
' from the returned format, as part of the translation process between FreeImage format IDs and PhotoDemon format IDs.
Private Function FI_DetermineFiletype(ByVal srcFilename As String, ByRef dstDIB As pdDIB) As FREE_IMAGE_FORMAT

    'While we could manually test our extension against the FreeImage database, it is capable of doing so itself.
    'First, check the file header to see if it matches a known head type
    Dim fileFIF As FREE_IMAGE_FORMAT
    fileFIF = FreeImage_GetFileTypeU(StrPtr(srcFilename))
    
    'For certain filetypes (CUT, MNG, PCD, TARGA and WBMP, according to the FreeImage documentation),
    ' the lack of a reliable header may prevent GetFileType from working.  As a result, double-check
    ' the file using its extension.
    If (fileFIF = FIF_UNKNOWN) Then fileFIF = FreeImage_GetFIFFromFilenameU(StrPtr(srcFilename))
    
    'By this point, if the file still doesn't show up in FreeImage's database, abandon the import attempt.
    If (fileFIF <> FIF_UNKNOWN) Then
        If (Not FreeImage_FIFSupportsReading(fileFIF)) Then fileFIF = FIF_UNKNOWN
    End If
    
    'Store this file format inside the DIB
    Dim internalFIF As PD_IMAGE_FORMAT
    internalFIF = fileFIF
    
    'All pixmap formats are condensed down to PNM, which greatly simplifies internal tracking
    Select Case internalFIF
        Case PDIF_PBM, PDIF_PBMRAW, PDIF_PFM, PDIF_PGM, PDIF_PGMRAW, PDIF_PNM, PDIF_PPM, PDIF_PPMRAW
            internalFIF = PDIF_PNM
    End Select
    
    If (Not dstDIB Is Nothing) Then dstDIB.SetOriginalFormat internalFIF
    
    FI_DetermineFiletype = fileFIF
    
End Function

'Given a path to an incoming file, the file's format, and an optional "use preview" setting (which will grab thumbnails only),
' determine the correct load-time flags for FreeImage.
Private Function FI_DetermineImportFlags(ByVal srcFilename As String, ByVal fileFIF As FREE_IMAGE_FORMAT, Optional ByVal usePreviewIfAvailable As Boolean = False) As FREE_IMAGE_LOAD_OPTIONS

    'Certain filetypes offer import options.  Check the FreeImage type to see if we want to enable any optional flags.
    Dim fi_ImportFlags As FREE_IMAGE_LOAD_OPTIONS
    fi_ImportFlags = 0
    
    Select Case fileFIF
            
        Case FIF_JPEG
            
            'For JPEGs, specify a preference for accuracy and quality over import speed.
            fi_ImportFlags = fi_ImportFlags Or FILO_JPEG_ACCURATE
            
            'The user can modify EXIF auto-rotation behavior
            If ImageImporter.GetImportPref_JPEGOrientation() Then fi_ImportFlags = fi_ImportFlags Or FILO_JPEG_EXIFROTATE
            
            'CMYK files are fully supported
            fi_ImportFlags = fi_ImportFlags Or FILO_JPEG_CMYK
        
        Case FIF_RAW
            
            'If this is not a primary image, RAW format files can load just their thumbnail.  This is significantly faster.
            If usePreviewIfAvailable Then fi_ImportFlags = fi_ImportFlags Or FILO_RAW_PREVIEW
        
        Case FIF_TIFF
            
            'CMYK files are fully supported
            fi_ImportFlags = fi_ImportFlags Or TIFF_CMYK
    
    End Select
        
    FI_DetermineImportFlags = fi_ImportFlags
    
End Function

Private Function FI_LoadICCProfile(ByVal fi_Bitmap As Long, ByRef dstProfile As pdICCProfile) As Boolean
    
    If (FreeImage_GetICCProfileSize(fi_Bitmap) > 0) Then
        
        If (dstProfile Is Nothing) Then Set dstProfile = New pdICCProfile
        
        Dim iccSize As Long, iccPtr As Long
        iccSize = Outside_FreeImageV3.FreeImage_GetICCProfileSize(fi_Bitmap)
        iccPtr = Outside_FreeImageV3.FreeImage_GetICCProfileDataPointer(fi_Bitmap)
        FI_LoadICCProfile = dstProfile.LoadICCFromPtr(iccSize, iccPtr)
        
    Else
        FI_DebugMsg "WARNING!  ICC profile size is invalid (<=0)."
    End If
    
End Function

Private Function FI_GetImageTypeAsString(ByVal fi_DataType As FREE_IMAGE_TYPE) As String

    Select Case fi_DataType
        Case FIT_UNKNOWN
            FI_GetImageTypeAsString = "Unknown"
        Case FIT_BITMAP
            FI_GetImageTypeAsString = "Standard bitmap (1 to 32bpp)"
        Case FIT_UINT16
            FI_GetImageTypeAsString = "HDR Grayscale (Unsigned int)"
        Case FIT_INT16
            FI_GetImageTypeAsString = "HDR Grayscale (Signed int)"
        Case FIT_UINT32
            FI_GetImageTypeAsString = "HDR Grayscale (Unsigned long)"
        Case FIT_INT32
            FI_GetImageTypeAsString = "HDR Grayscale (Signed long)"
        Case FIT_FLOAT
            FI_GetImageTypeAsString = "HDR Grayscale (Float)"
        Case FIT_DOUBLE
            FI_GetImageTypeAsString = "HDR Grayscale (Double)"
        Case FIT_COMPLEX
            FI_GetImageTypeAsString = "Complex (2xDouble)"
        Case FIT_RGB16
            FI_GetImageTypeAsString = "HDR RGB (3xInteger)"
        Case FIT_RGBA16
            FI_GetImageTypeAsString = "HDR RGBA (4xInteger)"
        Case FIT_RGBF
            FI_GetImageTypeAsString = "HDR RGB (3xFloat)"
        Case FIT_RGBAF
            FI_GetImageTypeAsString = "HDR RGBA (4xFloat)"
    End Select

End Function

'Unload a FreeImage handle.  If the handle is to a multipage object, pass that handle, too; this function will automatically switch
' to multipage behavior if the multipage handle is non-zero.
'
'On success, any unloaded handles will be forcibly reset to zero.
Private Sub FI_Unload(ByRef srcFIHandle As Long, Optional ByRef srcFIMultipageHandle As Long = 0, Optional ByVal leaveMultiHandleOpen As Boolean = False, Optional ByRef fiDibIsDetached As Boolean = False)
    
    If ((srcFIMultipageHandle = 0) Or fiDibIsDetached) Then
        If (srcFIHandle <> 0) Then FreeImage_UnloadEx srcFIHandle
        srcFIHandle = 0
    Else
        
        If (srcFIHandle <> 0) Then
            FreeImage_UnlockPage srcFIMultipageHandle, srcFIHandle, False
            srcFIHandle = 0
        End If
        
        'Now comes a weird bit of special handling.  It may be desirable to unlock a page, but leave the base multipage image open.
        ' (When loading a multipage image, this yields much better performance.)  However, we need to note that the resulting
        ' DIB handle is now "detached", meaning we can't use UnlockPage on it in the future.
        If (Not leaveMultiHandleOpen) Then
            FreeImage_CloseMultiBitmap srcFIMultipageHandle
            srcFIMultipageHandle = 0
        Else
            fiDibIsDetached = True
        End If
        
    End If
End Sub

'See if an image file is actually comprised of multiple files (e.g. animated GIFs, multipage TIFs).
' Input: file name to be checked
' Returns: 0 if only one image is found.  Page (or frame) count if multiple images are found.
Public Function IsMultiImage(ByRef srcFilename As String) As Long

    'Ensure library is available before proceeding
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    
    On Error GoTo isMultiImage_Error
    
    'Double-check that FreeImage.dll was located at start-up
    If (Not ImageFormats.IsFreeImageEnabled()) Then
        IsMultiImage = 0
        Exit Function
    End If
        
    'Determine the file type.  (Currently, this feature only works on animated GIFs, multipage TIFFs, and icons.)
    Dim fileFIF As FREE_IMAGE_FORMAT
    fileFIF = FreeImage_GetFileTypeU(StrPtr(srcFilename))
    If (fileFIF = FIF_UNKNOWN) Then fileFIF = FreeImage_GetFIFFromFilenameU(StrPtr(srcFilename))
    
    'If FreeImage can't determine the file type, or if the filetype is not GIF or TIF, return False
    If (Not FreeImage_FIFSupportsReading(fileFIF)) Or ((fileFIF <> PDIF_GIF) And (fileFIF <> PDIF_TIFF) And (fileFIF <> FIF_ICO)) Then
        IsMultiImage = 0
        Exit Function
    End If
    
    'At this point, we are guaranteed that the image is a GIF, TIFF, or icon file.
    ' Open the file using the multipage function
    Dim fi_multi_hDIB As Long
    If (fileFIF = PDIF_GIF) Then
        fi_multi_hDIB = FreeImage_OpenMultiBitmap(PDIF_GIF, srcFilename)
    ElseIf (fileFIF = FIF_ICO) Then
        fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_ICO, srcFilename)
    Else
        fi_multi_hDIB = FreeImage_OpenMultiBitmap(PDIF_TIFF, srcFilename)
    End If
    
    'Get the page count, then close the file
    Dim pageCheck As Long
    pageCheck = FreeImage_GetPageCount(fi_multi_hDIB)
    FreeImage_CloseMultiBitmap fi_multi_hDIB
    
    'Return the page count (which will be zero if only a single page or frame is present)
    IsMultiImage = pageCheck
    
    Exit Function
    
isMultiImage_Error:

    IsMultiImage = 0

End Function

'Given a source FreeImage handle with an attached ICC profile, create a new, ICC-corrected version of the
' image and place it inside the destination DIB if at all possible.  The byref parameter pdDIBIsDestination
' will be set to TRUE if this approach succeeds; if it is set to FALSE, you must use the fallbackFIHandle,
' instead, which will point to a newly allocated FreeImage object.
'
'IMPORTANT NOTE: the source handle *will not be freed*, even if the transformation is successful.  The caller must do this manually.
Private Function GenerateICCCorrectedFIDIB(ByVal srcFIHandle As Long, ByRef dstDIB As pdDIB, ByRef pdDIBIsDestination As Boolean, ByRef fallbackFIHandle As Long, Optional ByRef useThisProfile As pdICCProfile = Nothing) As Boolean
    
    GenerateICCCorrectedFIDIB = False
    pdDIBIsDestination = False
    fallbackFIHandle = 0
    
    'Retrieve the source image's bit-depth and data type.
    Dim fi_BPP As Long
    fi_BPP = FreeImage_GetBPP(srcFIHandle)
    
    Dim fi_DataType As FREE_IMAGE_TYPE
    fi_DataType = FreeImage_GetImageType(srcFIHandle)
    
    'FreeImage provides a bunch of custom identifiers for various grayscale types.  When one of these is found, we can
    ' skip further heuristics.
    Dim isGrayscale As Boolean
    Select Case fi_DataType
    
        Case FIT_DOUBLE, FIT_FLOAT, FIT_INT16, FIT_UINT16, FIT_INT32, FIT_UINT32
            isGrayscale = True
        
        'Note that a lack of identifiers *doesn't necessarily mean* the image is not grayscale.  It simply means the image is
        ' not in a FreeImage-specific grayscale format.  (Some formats, like 16-bit grayscale + 16-bit alpha are not supported
        ' by FreeImage, and will be returned as 64-bpp RGBA instead.)
        Case Else
            isGrayscale = False
    
    End Select
    
    'Check for 8-bpp grayscale images now; they use a separate detection technique
    If (Not isGrayscale) Then
        If (fi_BPP = 8) Then
            If ((FreeImage_GetColorType(srcFIHandle) = FIC_MINISBLACK) Or (FreeImage_GetColorType(srcFIHandle) = FIC_MINISWHITE)) Then isGrayscale = True
        End If
    End If
        
    'Also, check for transparency in the source image.  Color-management will generally ignore alpha values, but we need to
    ' supply a flag telling the ICC engine to mirror alpha bytes to the new DIB copy.
    Dim hasTransparency As Boolean, transparentEntries As Long
    hasTransparency = FreeImage_IsTransparent(srcFIHandle)
    If (Not hasTransparency) Then
    
        transparentEntries = FreeImage_GetTransparencyCount(srcFIHandle)
        hasTransparency = (transparentEntries > 0)
        
        '32-bpp images with a fully opaque alpha channel may return FALSE; this is a stupid FreeImage issue.
        ' Check for such a mismatch, and forcibly mark the data as 32-bpp RGBA.  (Otherwise we will get stride issues when
        ' applying the color management transform.)
        If (fi_BPP = 32) Then
            If ((FreeImage_GetColorType(srcFIHandle) = FIC_RGB) Or (FreeImage_GetColorType(srcFIHandle) = FIC_RGBALPHA)) Then hasTransparency = True
        End If
        
    End If
    
    'Allocate a destination FI DIB object in default BGRA order.  Note that grayscale images specifically use an 8-bpp target;
    ' this is by design, as the ICC engine cannot perform grayscale > RGB expansion.  (Instead, we must perform the ICC transform
    ' in pure grayscale space, *then* translate the result to RGB.)
    '
    'Note also that we still have not addressed the problem where "isGrayscale = True" but FreeImage has mis-detected color.
    ' We will deal with this in a subsequent step.
    Dim targetBitDepth As Long
    If isGrayscale Then
        targetBitDepth = 8
    Else
        If hasTransparency Then targetBitDepth = 32 Else targetBitDepth = 24
    End If
    
    '8-bpp grayscale images will use a FreeImage container instead of a pdDIB.  (pdDIB objects only support 24- and 32-bpp targets.)
    Dim newFIDIB As Long
    If (targetBitDepth = 8) Then
        newFIDIB = FreeImage_Allocate(FreeImage_GetWidth(srcFIHandle), FreeImage_GetHeight(srcFIHandle), targetBitDepth)
    Else
        dstDIB.CreateBlank FreeImage_GetWidth(srcFIHandle), FreeImage_GetHeight(srcFIHandle), 32, 0, 255
    End If
    
    'Extract the embedded ICC profile into a pdICCProfile object
    Dim tmpProfile As pdICCProfile
    If (useThisProfile Is Nothing) Then
        Set tmpProfile = New pdICCProfile
        FI_LoadICCProfile srcFIHandle, tmpProfile
    Else
        Set tmpProfile = useThisProfile
    End If
    
    'We now want to use LittleCMS to perform an immediate ICC correction.
    
    'Start by creating two LCMS profile handles:
    ' 1) a source profile (the in-memory copy of the ICC profile associated with this DIB)
    ' 2) a destination profile (the current PhotoDemon working space)
    Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile
    Set srcProfile = New pdLCMSProfile
    Set dstProfile = New pdLCMSProfile
    
    If srcProfile.CreateFromPDICCObject(tmpProfile) Then
        
        Dim dstProfileSuccess As Long
        If isGrayscale Then
            dstProfileSuccess = dstProfile.CreateGenericGrayscaleProfile
        Else
            dstProfileSuccess = dstProfile.CreateSRGBProfile
        End If
        
        If dstProfileSuccess Then
            
            'DISCLAIMER! Until rendering intent has a dedicated preference, PD defaults to perceptual render intent.
            ' This provides better results on most images, it correctly preserves gamut, and it is the standard
            ' behavior for PostScript workflows.  See http://fieryforums.efi.com/showthread.php/835-Rendering-Intent-Control-for-Embedded-Profiles
            ' Also see: https://developer.mozilla.org/en-US/docs/ICC_color_correction_in_Firefox)
            '
            'For future reference, I've left the code below for retrieving rendering intent from the source profile
            Dim targetRenderingIntent As LCMS_RENDERING_INTENT
            targetRenderingIntent = INTENT_PERCEPTUAL
            'targetRenderingIntent = srcProfile.GetRenderingIntent
            
            'Now, we need to create a transform between the two bit-depths.  This involves mapping the FreeImage bit-depth constants
            ' to compatible LCMS ones.
            Dim srcPixelFormat As LCMS_PIXEL_FORMAT, dstPixelFormat As LCMS_PIXEL_FORMAT
            
            'FreeImage does not natively support grayscale+alpha images.  These will be implicitly mapped to RGBA, so we only
            ' need to check grayscale formats if hasTransparency = FALSE.
            Dim transformImpossible As Boolean: transformImpossible = False
            
            If isGrayscale Then
                
                'Regardless of alpha, we want to map the grayscale data to an 8-bpp target.  (If alpha is present, we will
                ' manually back up the current alpha-bytes, then re-apply them after the ICC transform completes.)
                dstPixelFormat = TYPE_GRAY_8
                
                If (fi_DataType = FIT_DOUBLE) Then
                    srcPixelFormat = TYPE_GRAY_DBL
                ElseIf (fi_DataType = FIT_FLOAT) Then
                    srcPixelFormat = TYPE_GRAY_FLT
                ElseIf (fi_DataType = FIT_INT16) Then
                    srcPixelFormat = TYPE_GRAY_16
                ElseIf (fi_DataType = FIT_UINT16) Then
                    srcPixelFormat = TYPE_GRAY_16
                ElseIf (fi_DataType = FIT_INT32) Then
                    transformImpossible = True
                ElseIf (fi_DataType = FIT_UINT32) Then
                    transformImpossible = True
                Else
                    srcPixelFormat = TYPE_GRAY_8
                End If
                
            Else
                
                'Regardless of source transparency, we *always* map the image to a 32-bpp target
                dstPixelFormat = TYPE_BGRA_8
                    
                If hasTransparency Then
                
                    If (fi_DataType = FIT_BITMAP) Then
                        If (FreeImage_GetRedMask(srcFIHandle) > FreeImage_GetBlueMask(srcFIHandle)) Then
                            srcPixelFormat = TYPE_BGRA_8
                        Else
                            srcPixelFormat = TYPE_RGBA_8
                        End If
                        
                    ElseIf (fi_DataType = FIT_RGBA16) Then
                        If (FreeImage_GetRedMask(srcFIHandle) > FreeImage_GetBlueMask(srcFIHandle)) Then
                            srcPixelFormat = TYPE_BGRA_16
                        Else
                            srcPixelFormat = TYPE_RGBA_16
                        End If
                        
                    'The only other possibility is RGBAF; LittleCMS supports this format, but we'd have to construct our own macro
                    ' to define it.  Just skip it at present.
                    Else
                        transformImpossible = True
                    End If
                    
                Else
                    
                    If (fi_DataType = FIT_BITMAP) Then
                        If (FreeImage_GetRedMask(srcFIHandle) > FreeImage_GetBlueMask(srcFIHandle)) Then
                            srcPixelFormat = TYPE_BGR_8
                        Else
                            srcPixelFormat = TYPE_RGB_8
                        End If
                    ElseIf (fi_DataType = FIT_RGB16) Then
                        If (FreeImage_GetRedMask(srcFIHandle) > FreeImage_GetBlueMask(srcFIHandle)) Then
                            srcPixelFormat = TYPE_BGR_16
                        Else
                            srcPixelFormat = TYPE_RGB_16
                        End If
                        
                    'The only other possibility is RGBF; LittleCMS supports this format, but we'd have to construct our own macro
                    ' to define it.  Just skip it at present.
                    Else
                        transformImpossible = True
                    End If
                    
                End If
            
            End If
            
            'Some color spaces may not be supported; that's okay - we'll use tone-mapping to handle them.
            If (Not transformImpossible) Then
                
                'Create a transform that uses the target DIB as both the source and destination
                Dim cTransform As pdLCMSTransform
                Set cTransform = New pdLCMSTransform
                If cTransform.CreateTwoProfileTransform(srcProfile, dstProfile, srcPixelFormat, dstPixelFormat, targetRenderingIntent) Then
                    
                    'LittleCMS 2.0 allows us to free our source profiles immediately after a transform is created.
                    ' (Note that we don't *need* to do this, nor does this code leak if we don't manually free both
                    '  profiles, but as we're about to do an energy- and memory-intensive operation, it doesn't
                    '  hurt to free the profiles now.)
                    Set srcProfile = Nothing: Set dstProfile = Nothing
                    
                    'At present, grayscale images will be converted into a destination FreeImage handle
                    Dim transformSuccess As Boolean
                    
                    If isGrayscale Then
                        transformSuccess = cTransform.ApplyTransformToArbitraryMemory(FreeImage_GetScanline(srcFIHandle, 0), FreeImage_GetScanline(newFIDIB, 0), FreeImage_GetPitch(srcFIHandle), FreeImage_GetPitch(newFIDIB), FreeImage_GetHeight(srcFIHandle), FreeImage_GetWidth(srcFIHandle))
                    Else
                        transformSuccess = cTransform.ApplyTransformToArbitraryMemory(FreeImage_GetScanline(srcFIHandle, 0), dstDIB.GetDIBScanline(0), FreeImage_GetPitch(srcFIHandle), dstDIB.GetDIBStride, FreeImage_GetHeight(srcFIHandle), FreeImage_GetWidth(srcFIHandle), True)
                    End If
                    
                    If transformSuccess Then
                    
                        FI_DebugMsg "Color-space transformation successful."
                        dstDIB.SetColorManagementState cms_ProfileConverted
                        dstDIB.SetColorProfileHash ColorManagement.GetSRGBProfileHash()
                        GenerateICCCorrectedFIDIB = True
                        
                        'We now need to clarify for the caller where the ICC-transformed data sits.  8-bpp grayscale *without* alpha
                        ' will be stored in a new 8-bpp FreeImage object.  Other formats have likely been placed directly into
                        ' the target pdDIB object (which means the FreeImage loader can skip subsequent steps).
                        If isGrayscale Then
                            pdDIBIsDestination = False
                            fallbackFIHandle = newFIDIB
                            
                        'Non-grayscale images *always* get converted directly into a pdDIB object.
                        Else
                            pdDIBIsDestination = True
                            fallbackFIHandle = 0
                            If (targetBitDepth = 24) Then dstDIB.SetInitialAlphaPremultiplicationState True Else dstDIB.SetAlphaPremultiplication True
                        End If
                        
                    End If
                    
                    'Note that we could free the transform here, but it's unnecessary.  (The pdLCMSTransform class
                    ' is self-freeing upon destruction.)
                    
                Else
                    FI_DebugMsg "WARNING!  Plugin_FreeImage.GenerateICCCorrectedFIDIB failed to create a valid transformation handle!"
                End If
            
            'Impossible transformations return a null handle
            Else
                FI_DebugMsg "WARNING!  Plugin_FreeImage.GenerateICCCorrectedFIDIB is functional, but the source pixel format is incompatible with the current ICC engine."
            End If
        
        Else
            FI_DebugMsg "WARNING!  Plugin_FreeImage.GenerateICCCorrectedFIDIB failed to create a valid destination profile handle."
        End If
    
    Else
        FI_DebugMsg "WARNING!  Plugin_FreeImage.GenerateICCCorrectedFIDIB failed to create a valid source profile handle."
    End If
    
    'If the transformation failed, free our temporarily allocated FreeImage DIB
    If (Not GenerateICCCorrectedFIDIB) And (newFIDIB <> 0) Then FI_Unload newFIDIB

End Function

'Given a source FreeImage handle in CMYK format, and a destination pdDIB that contains a valid ICC profile,
' create a new, ICC-corrected version of the image, in RGB format, and stored inside the destination pdDIB.
'
'IMPORTANT NOTE: the source handle *will not be freed*, even if the transformation is successful.
' The caller must free it manually.
Private Function ConvertCMYKFiDIBToRGB(ByVal srcFIHandle As Long, ByRef dstDIB As pdDIB) As Boolean
    
    'As a failsafe, confirm that the incoming image is CMYK format *and* that it has an ICC profile
    If (FreeImage_GetColorType(srcFIHandle) = FIC_CMYK) And FreeImage_HasICCProfile(srcFIHandle) Then
    
        'Prep the source DIB
        If dstDIB.CreateBlank(FreeImage_GetWidth(srcFIHandle), FreeImage_GetHeight(srcFIHandle), 32, 0, 255) Then
            
            'Extract the ICC profile into a pICCProfile object
            Dim tmpProfile As pdICCProfile
            FI_LoadICCProfile srcFIHandle, tmpProfile
            
            'We now want to use LittleCMS to perform an immediate ICC correction.
            
            'Start by creating two LCMS profile handles:
            ' 1) a source profile (the in-memory copy of the ICC profile associated with this DIB)
            ' 2) a destination profile (the current PhotoDemon working space)
            Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile
            Set srcProfile = New pdLCMSProfile
            Set dstProfile = New pdLCMSProfile
            
            If srcProfile.CreateFromPDICCObject(tmpProfile) Then
                
                If dstProfile.CreateSRGBProfile() Then
                    
                    'DISCLAIMER! Until rendering intent has a dedicated preference, PD defaults to perceptual render intent.
                    ' This provides better results on most images, it correctly preserves gamut, and it is the standard
                    ' behavior for PostScript workflows.  See http://fieryforums.efi.com/showthread.php/835-Rendering-Intent-Control-for-Embedded-Profiles
                    ' Also see: https://developer.mozilla.org/en-US/docs/ICC_color_correction_in_Firefox)
                    '
                    'For future reference, I've left the code below for retrieving rendering intent from the source profile
                    Dim targetRenderingIntent As LCMS_RENDERING_INTENT
                    targetRenderingIntent = INTENT_PERCEPTUAL
                    'targetRenderingIntent = srcProfile.GetRenderingIntent
                    
                    'Now, we need to create a transform between the two bit-depths.  This involves mapping the FreeImage bit-depth constants
                    ' to compatible LCMS ones.
                    Dim srcPixelFormat As LCMS_PIXEL_FORMAT, dstPixelFormat As LCMS_PIXEL_FORMAT
                    If (FreeImage_GetBPP(srcFIHandle) = 64) Then srcPixelFormat = TYPE_CMYK_16 Else srcPixelFormat = TYPE_CMYK_8
                    dstPixelFormat = TYPE_BGRA_8
                    
                    'Create a transform that uses the target DIB as both the source and destination
                    Dim cTransform As pdLCMSTransform
                    Set cTransform = New pdLCMSTransform
                    If cTransform.CreateTwoProfileTransform(srcProfile, dstProfile, srcPixelFormat, dstPixelFormat, targetRenderingIntent) Then
                        
                        'LittleCMS 2.0 allows us to free our source profiles immediately after a transform is created.
                        ' (Note that we don't *need* to do this, nor does this code leak if we don't manually free both
                        '  profiles, but as we're about to do an energy- and memory-intensive operation, it doesn't
                        '  hurt to free the profiles now.)
                        Set srcProfile = Nothing: Set dstProfile = Nothing
                        
                        If cTransform.ApplyTransformToArbitraryMemory(FreeImage_GetScanline(srcFIHandle, 0), dstDIB.GetDIBScanline(0), FreeImage_GetPitch(srcFIHandle), dstDIB.GetDIBStride, FreeImage_GetHeight(srcFIHandle), FreeImage_GetWidth(srcFIHandle), True) Then
                            FI_DebugMsg "ICC profile transformation successful.  New FreeImage handle now lives in the current RGB working space."
                            dstDIB.SetColorManagementState cms_ProfileConverted
                            dstDIB.SetColorProfileHash ColorManagement.GetSRGBProfileHash()
                            dstDIB.SetInitialAlphaPremultiplicationState True
                            ConvertCMYKFiDIBToRGB = True
                        End If
                    
                    'Note that we could free the transform here, but it's unnecessary.  (The pdLCMSTransform class
                    ' is self-freeing upon destruction.)
                    
                    Else
                        FI_DebugMsg "WARNING!  Plugin_FreeImage.ConvertCMYKFiDIBToRGB failed to create a valid transformation handle!"
                    End If
                    
                Else
                    FI_DebugMsg "WARNING!  Plugin_FreeImage.ConvertCMYKFiDIBToRGB failed to create a valid destination profile handle."
                End If
            
            Else
                FI_DebugMsg "WARNING!  Plugin_FreeImage.ConvertCMYKFiDIBToRGB failed to create a valid source profile handle."
            End If
            
        Else
            FI_DebugMsg "WARNING!  Destination DIB could not be allocated - is the source image corrupt?"
        End If
    
    Else
        FI_DebugMsg "WARNING!  Don't call ConvertCMYKFiDIBToRGB() if the source object is not CMYK format!"
    End If

End Function

'Given a FreeImage handle, return a 24 or 32bpp pdDIB object, as relevant.  Note that this function does not modify
' premultiplication status of 32bpp images.  The caller is responsible for applying that (as necessary).
'
'NOTE!  This function requires the FreeImage DIB to already be in 24 or 32bpp format.  It will fail if another bit-depth is used.
'ALSO NOTE!  This function does not set alpha premultiplication.  It's assumed that the caller knows that value in advance.
'ALSO NOTE!  This function does not free the incoming FreeImage handle, by design.
Public Function GetPDDibFromFreeImageHandle(ByVal srcFI_Handle As Long, ByRef dstDIB As pdDIB) As Boolean
    
    'Ensure library is available before proceeding
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    
    'We may need to perform an intermediary color-depth transform, which may change the passed FI handle.
    ' Check this state and do *not* free the intermediary handle if it matches the original one.
    Dim fiHandleBackup As Long
    fiHandleBackup = srcFI_Handle
    
    'Double-check the FreeImage handle's bit depth
    Dim fiBPP As Long
    fiBPP = FreeImage_GetBPP(srcFI_Handle)
    
    If (fiBPP <> 24) And (fiBPP <> 32) Then
        
        'If the DIB is less than 24 bpp, upsample now
        If (fiBPP < 24) Then
            
            'Conversion to higher bit depths is contingent on the presence of an alpha channel
            If FreeImage_IsTransparent(srcFI_Handle) Or (FreeImage_GetTransparentIndex(srcFI_Handle) <> -1) Then
                srcFI_Handle = FreeImage_ConvertColorDepth(srcFI_Handle, FICF_RGB_32BPP, False)
            Else
                srcFI_Handle = FreeImage_ConvertColorDepth(srcFI_Handle, FICF_RGB_24BPP, False)
            End If
            
            'Verify the new bit-depth
            fiBPP = FreeImage_GetBPP(srcFI_Handle)
            If (fiBPP <> 24) And (fiBPP <> 32) Then
                
                'The transform failed (for whatever reason). If a new DIB was created, release it before exiting.
                If (srcFI_Handle <> 0) And (srcFI_Handle <> fiHandleBackup) Then FreeImage_UnloadEx srcFI_Handle
                GetPDDibFromFreeImageHandle = False
                Exit Function
                
            End If
            
        Else
            GetPDDibFromFreeImageHandle = False
            Exit Function
        End If
        
    End If
    
    'Proceed with DIB copying
    Dim fi_Width As Long, fi_Height As Long
    fi_Width = FreeImage_GetWidth(srcFI_Handle)
    fi_Height = FreeImage_GetHeight(srcFI_Handle)
    dstDIB.CreateBlank fi_Width, fi_Height, fiBPP, 0
    SetDIBitsToDevice dstDIB.GetDIBDC, 0, 0, fi_Width, fi_Height, 0, 0, 0, fi_Height, FreeImage_GetBits(srcFI_Handle), ByVal FreeImage_GetInfo(srcFI_Handle), 0&
    
    'If we created a temporary DIB, free it before exiting
    If (srcFI_Handle <> fiHandleBackup) Then FreeImage_UnloadEx srcFI_Handle
    
    GetPDDibFromFreeImageHandle = True
    
End Function

'Paint a FreeImage DIB to an arbitrary clipping rect on some target pdDIB.  This does not free or otherwise modify the source FreeImage object
Public Function PaintFIDibToPDDib(ByRef dstDIB As pdDIB, ByVal fi_Handle As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long) As Boolean
    
    'Ensure library is available before proceeding
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    
    If (Not dstDIB Is Nothing) And (fi_Handle <> 0) Then
        
        Dim bmpInfo As BITMAPINFO
        Outside_FreeImageV3.FreeImage_GetInfoHeaderEx fi_Handle, VarPtr(bmpInfo.bmiHeader)
        If dstDIB.IsDIBTopDown Then bmpInfo.bmiHeader.biHeight = -1 * (bmpInfo.bmiHeader.biHeight)
        
        Dim iHeight As Long: iHeight = Abs(bmpInfo.bmiHeader.biHeight)
        PaintFIDibToPDDib = (SetDIBitsToDevice(dstDIB.GetDIBDC, dstX, dstY, dstWidth, dstHeight, 0, 0, 0, iHeight, FreeImage_GetBits(fi_Handle), bmpInfo, 0&) <> 0)
        
        'When painting from a 24-bpp source to a 32-bpp target, the destination alpha channel will be ignored by GDI.
        ' We must forcibly fill it with opaque alpha values, or the resulting image will retain its existing alpha (typically 0!)
        If (dstDIB.GetDIBColorDepth = 32) And (FreeImage_GetBPP(fi_Handle) = 24) Then dstDIB.ForceNewAlpha 255
        
    Else
        FI_DebugMsg "WARNING!  Destination DIB is empty or FreeImage handle is null.  Cannot proceed with painting."
    End If
    
End Function

'Prior to applying tone-mapping settings, query the user for their preferred behavior.
' If the user doesn't want this dialog raised, this function will silently retrieve the last-used settings
' from the preference file, and proceed with tone-mapping automatically.
' (The silent behavior can also be enforced by setting the noUIMode parameter to TRUE.)
'
'Returns: fills dst_fiHandle with a non-zero FreeImage 24 or 32bpp image handle if successful.  0 if unsuccessful.
'         The function itself will return a PD_OPERATION_OUTCOME value; this is important for determining if the
'         user canceled the dialog.
'
'IMPORTANT NOTE!  If this function fails, further loading of the image must be halted.  PD cannot yet operate
' internally on anything larger than 32bpp, so if tone-mapping fails, we must abandon loading.
' (A failure state can also be triggered by the user canceling the tone-mapping dialog.)
Private Function RaiseToneMapDialog(ByRef fi_Handle As Long, ByRef dst_fiHandle As Long, Optional ByVal noUIMode As Boolean = False) As PD_OPERATION_OUTCOME

    'Ask the user how they want to proceed.  Note that the dialog wrapper automatically handles the case of "do not prompt;
    ' use previous settings."  If that happens, it will retrieve the proper conversion settings for us, and return a dummy
    ' value of OK (as if the dialog were actually raised).
    Dim howToProceed As VbMsgBoxResult, ToneMapSettings As String
    If noUIMode Then
        howToProceed = vbOK
        ToneMapSettings = vbNullString
    Else
        howToProceed = Dialogs.PromptToneMapSettings(fi_Handle, ToneMapSettings)
    End If
    
    'Check for a cancellation state; if encountered, abandon ship now.
    If (howToProceed = vbOK) Then
        
        'The ToneMapSettings string will now contain all the information we need to proceed with the tone-map.
        ' Forward it to the central tone-mapping handler and use its success/fail state for this function as well.
        FI_DebugMsg "Tone-map dialog appears to have been successful; result = " & howToProceed
        If (Not noUIMode) Then Message "Applying tone-mapping..."
        dst_fiHandle = ApplyToneMapping(fi_Handle, ToneMapSettings)
        
        If (dst_fiHandle = 0) Then
            FI_DebugMsg "WARNING!  ApplyToneMapping() failed for reasons unknown."
            RaiseToneMapDialog = PD_FAILURE_GENERIC
        Else
            RaiseToneMapDialog = PD_SUCCESS
        End If
        
    Else
        FI_DebugMsg "Tone-map dialog appears to have been cancelled; result = " & howToProceed
        dst_fiHandle = 0
        RaiseToneMapDialog = PD_FAILURE_USER_CANCELED
    End If

End Function

'Apply tone-mapping to a FreeImage DIB.  All valid FreeImage data types are supported, but for performance reasons, an intermediate cast to
' RGBF or RGBAF may be applied (because VB doesn't provide unsigned Int datatypes).
'
'Returns: a non-zero FreeImage 24 or 32bpp image handle if successful.  0 if unsuccessful.
'
'IMPORTANT NOTE!  This function always releases the incoming FreeImage handle, regardless of success or failure state.  This is
' to ensure proper load behavior (e.g. loading can't continue after a failed conversion, because we've forcibly killed the image handle),
' and to reduce resource usage (as the source handle is likely enormous, and we don't want it sitting around any longer than is
' absolutely necessary).
Public Function ApplyToneMapping(ByRef fi_Handle As Long, ByRef inputSettings As String) As Long
    
    'Ensure library is available before proceeding
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    
    'Retrieve the source image's bit-depth and data type.  These are crucial to successful tone-mapping operations.
    Dim fi_BPP As Long
    fi_BPP = FreeImage_GetBPP(fi_Handle)
    
    Dim fi_DataType As FREE_IMAGE_TYPE
    fi_DataType = FreeImage_GetImageType(fi_Handle)
    
    'Also, check for transparency in the source image.
    Dim hasTransparency As Boolean, transparentEntries As Long
    hasTransparency = FreeImage_IsTransparent(fi_Handle)
    transparentEntries = FreeImage_GetTransparencyCount(fi_Handle)
    If (transparentEntries > 0) Then hasTransparency = True
    
    Dim newHandle As Long, rgbfHandle As Long
    
    'toneMapSettings contains all conversion instructions.  Parse it to determine which tone-map function to use.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString inputSettings
    
    'The first parameter contains the requested tone-mapping operation.
    Select Case cParams.GetLong("method", PDTM_DRAGO)
    
        'Linear map
        Case PDTM_LINEAR
                
            newHandle = fi_Handle
            
            'For performance reasons, I've only written a single RGBF/RGBAF-based linear transform.  If the image is not in one
            ' of these formats, convert it now.
            If ((fi_DataType <> FIT_RGBF) And (fi_DataType <> FIT_RGBAF)) Then
                
                'In the future, a transparency-friendly conversion may become available.  For now, however, transparency
                ' is sacrificed as part of the conversion function (as FreeImage does not provide an RGBAF cast).
                If hasTransparency Then
                    rgbfHandle = FreeImage_ConvertToRGBAF(fi_Handle)
                Else
                    rgbfHandle = FreeImage_ConvertToRGBF(fi_Handle)
                End If
                
                If (rgbfHandle = 0) Then
                    FI_DebugMsg "WARNING!  FreeImage_ConvertToRGBA/F failed for reasons unknown."
                    ApplyToneMapping = 0
                    Exit Function
                Else
                    FI_DebugMsg "FreeImage_ConvertToRGBA/F successful.  Proceeding with manual tone-mapping operation."
                End If
                
                newHandle = rgbfHandle
                
            End If
            
            'At this point, fi_Handle now represents a 32-bpc RGBF (or RGBAF) type FreeImage DIB.  Apply manual tone-mapping now.
            newHandle = ConvertFreeImageRGBFTo24bppDIB(newHandle, cParams.GetLong("normalize", PD_BOOL_TRUE), cParams.GetBool("ignorenegative", PD_BOOL_TRUE), cParams.GetDouble("gamma", 2.2))
            
            'Unload the intermediate RGBF handle as necessary
            If (rgbfHandle <> 0) Then FreeImage_Unload rgbfHandle
            
            ApplyToneMapping = newHandle
            
        'Filmic tone-map; basically a nice S-curve with an emphasis on rich blacks
        Case PDTM_FILMIC
            
            newHandle = fi_Handle
            
            'For performance reasons, I've only written a single RGBF/RGBAF-based linear transform.  If the image is not in one
            ' of these formats, convert it now.
            If (fi_DataType <> FIT_RGBF) And (fi_DataType <> FIT_RGBAF) Then
                
                'In the future, a transparency-friendly conversion may become available.  For now, however, transparency
                ' is sacrificed as part of the conversion function (as FreeImage does not provide an RGBAF cast).
                If hasTransparency Then
                    rgbfHandle = FreeImage_ConvertToRGBAF(fi_Handle)
                Else
                    rgbfHandle = FreeImage_ConvertToRGBF(fi_Handle)
                End If
                
                If (rgbfHandle = 0) Then
                    FI_DebugMsg "WARNING!  FreeImage_ConvertToRGBA/F failed for reasons unknown."
                    ApplyToneMapping = 0
                    Exit Function
                Else
                    FI_DebugMsg "FreeImage_ConvertToRGBA/F successful.  Proceeding with manual tone-mapping operation."
                End If
                
                newHandle = rgbfHandle
                
            End If
            
            'At this point, fi_Handle now represents a 24bpp RGBF type FreeImage DIB.  Apply manual tone-mapping now.
            newHandle = ToneMapFilmic_RGBFTo24bppDIB(newHandle, cParams.GetDouble("gamma", 2.2), cParams.GetDouble("exposure", 2#), , , , , , , cParams.GetDouble("whitepoint", 11.2))
            
            'Unload the intermediate RGBF handle as necessary
            If (rgbfHandle <> 0) Then FreeImage_Unload rgbfHandle
            
            ApplyToneMapping = newHandle
        
        'Adaptive logarithmic map
        Case PDTM_DRAGO
            ApplyToneMapping = FreeImage_TmoDrago03(fi_Handle, cParams.GetDouble("gamma", 2.2), cParams.GetDouble("exposure", 0#))
            
        'Photoreceptor map
        Case PDTM_REINHARD
            ApplyToneMapping = FreeImage_TmoReinhard05Ex(fi_Handle, cParams.GetDouble("intensity", 0#), ByVal 0#, cParams.GetDouble("adaptation", 1#), cParams.GetDouble("colorcorrection", 0#))
        
    End Select
    
End Function

'Perform linear scaling of a 96bpp RGBF image to standard 24bpp.  Note that an intermediate pdDIB object is used for convenience, but the returned
' handle is to a FREEIMAGE DIB.
'
'Returns: a non-zero FreeImage 24bpp image handle if successful.  0 if unsuccessful.
'
'IMPORTANT NOTE: REGARDLESS OF SUCCESS, THIS FUNCTION DOES NOT FREE THE INCOMING fi_Handle PARAMETER.  If the function fails (returns 0),
' I assume the caller still wants the original handle so it can proceed accordingly.  Similarly, because this function is used to render
' tone-mapping previews, it doesn't make sense to free the handle upon success, either.
'
'OTHER IMPORTANT NOTE: it's probably obvious, but the 24bpp handle this function returns (if successful) must also be freed by the caller.
' Forget this, and the function will leak.
Private Function ConvertFreeImageRGBFTo24bppDIB(ByVal fi_Handle As Long, Optional ByVal toNormalize As PD_BOOL = PD_BOOL_AUTO, Optional ByVal ignoreNegative As Boolean = False, Optional ByVal newGamma As Double = 2.2) As Long
    
    'Before doing anything, check the incoming fi_Handle.  For performance reasons, this function only handles RGBF and RGBAF formats.
    ' Other formats are invalid.
    Dim fi_DataType As FREE_IMAGE_TYPE
    fi_DataType = FreeImage_GetImageType(fi_Handle)
    
    If (fi_DataType <> FIT_RGBF) And (fi_DataType <> FIT_RGBAF) Then
        FI_DebugMsg "Tone-mapping request invalid"
        ConvertFreeImageRGBFTo24bppDIB = 0
        Exit Function
    End If
    
    'Here's how this works: basically, we must manually convert the image, one scanline at a time, into 24bpp RGB format.
    ' In the future, it might be nice to provide different conversion algorithms, but for now, linear scaling is assumed.
    ' Some additional options can be set by the caller (like normalization)
    
    'Start by determining if normalization is required for this image.
    Dim mustNormalize As Boolean
    Dim minR As Double, maxR As Double, minG As Double, maxG As Double, minB As Double, maxB As Double
    Dim rDist As Double, gDist As Double, bDist As Double
    
    'The toNormalize input has three possible values: false, true, or "decide for yourself".  In the last case, the image will be scanned,
    ' and normalization will only be enabled if values fall outside the [0, 1] range.  (Files written by PhotoDemon will always be normalized
    ' at write-time, so this technique works well when moving images into and out of PD.)
    If toNormalize = PD_BOOL_AUTO Then
        mustNormalize = IsNormalizeRequired(fi_Handle, minR, maxR, minG, maxG, minB, maxB)
    Else
        mustNormalize = (toNormalize = PD_BOOL_TRUE)
        If mustNormalize Then IsNormalizeRequired fi_Handle, minR, maxR, minG, maxG, minB, maxB
    End If
    
    'I have no idea if normalization is supposed to include negative numbers or not; each high-bit-depth format has its own quirks,
    ' and none are clear on preferred defaults, so I'll leave this as manually settable for now.
    If ignoreNegative Then
        
        rDist = maxR
        gDist = maxG
        bDist = maxB
        
        minR = 0
        minG = 0
        minB = 0
    
    'If negative values are considered valid, calculate a normalization distance between the max and min values of each channel
    Else
    
        rDist = maxR - minR
        gDist = maxG - minG
        bDist = maxB - minB
    
    End If
    
    If (rDist <> 0#) Then rDist = 1# / rDist Else rDist = 0#
    If (gDist <> 0#) Then gDist = 1# / gDist Else gDist = 0#
    If (bDist <> 0#) Then bDist = 1# / bDist Else bDist = 0#
    
    'This Single-type array will consistently be updated to point to the current line of pixels in the image (RGBF format, remember!)
    Dim srcImageData() As Single, srcSA As SafeArray1D
    
    'Create a 24bpp or 32bpp DIB at the same size as the image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    If (fi_DataType = FIT_RGBF) Then
        tmpDIB.CreateBlank FreeImage_GetWidth(fi_Handle), FreeImage_GetHeight(fi_Handle), 24
    Else
        tmpDIB.CreateBlank FreeImage_GetWidth(fi_Handle), FreeImage_GetHeight(fi_Handle), 32
    End If
    
    'Point a byte array at the temporary DIB
    Dim dstImageData() As Byte, tmpSA As SafeArray2D
    tmpDIB.WrapArrayAroundDIB dstImageData, tmpSA
        
    'Iterate through each scanline in the source image, copying it to destination as we go.
    Dim iWidth As Long, iHeight As Long, iHeightInv As Long, iScanWidth As Long
    iWidth = FreeImage_GetWidth(fi_Handle) - 1
    iHeight = FreeImage_GetHeight(fi_Handle) - 1
    iScanWidth = FreeImage_GetPitch(fi_Handle)
    
    Dim pxSize As Long
    If (fi_DataType = FIT_RGBF) Then pxSize = 3 Else pxSize = 4
    
    'Prep any other post-processing adjustments
    Dim gammaCorrection As Double
    gammaCorrection = 1# / newGamma
    
    'Due to the potential math involved in conversion (if gamma and other settings are being toggled), we need a lot of intermediate variables.
    ' Depending on the user's settings, some of these may go unused.
    Dim rSrcF As Double, gSrcF As Double, bSrcF As Double
    Dim rDstF As Double, gDstF As Double, bDstF As Double
    Dim rDstL As Long, gDstL As Long, bDstL As Long
    
    'Alpha is also a possibility, but unlike RGB values, we assume it is always normalized.  This allows us to skip any intermediate processing,
    ' and simply copy the value directly into the destination (after redistributing to the proper range, of course).
    Dim aDstF As Double, aDstL As Long
    
    Dim x As Long, y As Long, xStride As Long
    
    'Point a 1D VB array at the first scanline
    For y = 0 To iHeight
    
        'FreeImage DIBs are stored bottom-up; we invert them during processing
        iHeightInv = iHeight - y
        
        'Update the current scanline pointer
        VBHacks.WrapArrayAroundPtr_Float srcImageData, srcSA, FreeImage_GetScanline(fi_Handle, y), iScanWidth * 4
        
        'Iterate through this line, converting values as we go
        For x = 0 To iWidth
            
            'Retrieve the source values.  This includes an implicit cast to Double, which I've done because some formats support
            ' IEEE constants like NaN or Infinity.  VB doesn't deal with these gracefully, and an implicit cast to Double seems
            ' to reduce unpredictable errors, possibly by giving any range-check code some breathing room.
            xStride = x * pxSize
            rSrcF = CDbl(srcImageData(xStride))
            gSrcF = CDbl(srcImageData(xStride + 1))
            bSrcF = CDbl(srcImageData(xStride + 2))
            If (pxSize = 4) Then aDstF = CDbl(srcImageData(xStride + 3))
            
            'If normalization is required, apply it now
            If mustNormalize Then
                
                'If the caller has requested that we ignore negative values, clamp negative values to zero
                If ignoreNegative Then
                
                    If (rSrcF < 0#) Then rSrcF = 0#
                    If (gSrcF < 0#) Then gSrcF = 0#
                    If (bSrcF < 0#) Then bSrcF = 0#
                
                'If negative values are considered valid, redistribute them on the range [0, Dist[Min, Max]]
                Else
                    rSrcF = rSrcF - minR
                    gSrcF = gSrcF - minG
                    bSrcF = bSrcF - minB
                End If
                
                rDstF = rSrcF * rDist
                gDstF = gSrcF * gDist
                bDstF = bSrcF * bDist
                
            'If an image does not need to be normalized, this step is much easier
            Else
                
                rDstF = rSrcF
                gDstF = gSrcF
                bDstF = bSrcF
                
            End If
            
            'FYI, alpha is always un-normalized
                        
            'Apply gamma now (if any).  Unfortunately, lookup tables aren't an option because we're dealing with floating-point input,
            ' so this step is a little slow due to the exponent operator.
            If (newGamma <> 1#) Then
                If (rDstF > 0#) Then rDstF = rDstF ^ gammaCorrection
                If (gDstF > 0#) Then gDstF = gDstF ^ gammaCorrection
                If (bDstF > 0#) Then bDstF = bDstF ^ gammaCorrection
            End If
            
            'In the future, additional corrections could be applied here.
            
            'Apply failsafe range checks now
            If (rDstF < 0#) Then rDstF = 0#
            If (rDstF > 1#) Then rDstF = 1#
                
            If (gDstF < 0#) Then gDstF = 0#
            If (gDstF > 1#) Then gDstF = 1#
                
            If (bDstF < 0#) Then bDstF = 0#
            If (bDstF > 1#) Then bDstF = 1#
            
            'Handle alpha, if necessary
            If (pxSize = 4) Then
                If (aDstF > 1#) Then aDstF = 1#
                If (aDstF < 0#) Then aDstF = 0#
                aDstL = aDstF * 255
            End If
            
            'Calculate corresponding integer values on the range [0, 255]
            rDstL = Int(rDstF * 255#)
            gDstL = Int(gDstF * 255#)
            bDstL = Int(bDstF * 255#)
                        
            'Copy the final, safe values into the destination
            dstImageData(xStride, iHeightInv) = bDstL
            dstImageData(xStride + 1, iHeightInv) = gDstL
            dstImageData(xStride + 2, iHeightInv) = rDstL
            If (pxSize = 4) Then dstImageData(xStride + 3, iHeightInv) = aDstL
            
        Next x
        
    Next y
    
    'Free our 1D array reference
    VBHacks.UnwrapArrayFromPtr_Float srcImageData
    
    'Point dstImageData() away from the DIB and deallocate it
    tmpDIB.UnwrapArrayFromDIB dstImageData
    
    'Create a FreeImage object from our pdDIB object, then release our pdDIB copy
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.GetDIBDC)
    
    'Success!
    ConvertFreeImageRGBFTo24bppDIB = fi_DIB

End Function

'Perform so-called "Filmic" tone-mapping of a 96bpp RGBF image to standard 24bpp.  Note that an intermediate pdDIB object is used
' for convenience, but the returned handle is to a FREEIMAGE DIB.
'
'Returns: a non-zero FreeImage 24bpp image handle if successful.  0 if unsuccessful.
'
'IMPORTANT NOTE: REGARDLESS OF SUCCESS, THIS FUNCTION DOES NOT FREE THE INCOMING fi_Handle PARAMETER.  If the function fails (returns 0),
' I assume the caller still wants the original handle so it can proceed accordingly.  Similarly, because this function is used to render
' tone-mapping previews, it doesn't make sense to free the handle upon success, either.
'
'OTHER IMPORTANT NOTE: it's probably obvious, but the 24bpp handle this function returns (if successful) must also be freed by the caller.
' Forget this, and the function will leak.
Private Function ToneMapFilmic_RGBFTo24bppDIB(ByVal fi_Handle As Long, Optional ByVal newGamma As Single = 2.2!, Optional ByVal exposureCompensation As Single = 2!, Optional ByVal shoulderStrength As Single = 0.22!, Optional ByVal linearStrength As Single = 0.3!, Optional ByVal linearAngle As Single = 0.1!, Optional ByVal toeStrength As Single = 0.2!, Optional ByVal toeNumerator As Single = 0.01!, Optional ByVal toeDenominator As Single = 0.3!, Optional ByVal linearWhitePoint As Single = 11.2!) As Long
    
    'Before doing anything, check the incoming fi_Handle.  For performance reasons, this function only handles RGBF and RGBAF formats.
    ' Other formats are invalid.
    Dim fi_DataType As FREE_IMAGE_TYPE
    fi_DataType = FreeImage_GetImageType(fi_Handle)
    
    If (fi_DataType <> FIT_RGBF) And (fi_DataType <> FIT_RGBAF) Then
        FI_DebugMsg "Tone-mapping request invalid"
        ToneMapFilmic_RGBFTo24bppDIB = 0
        Exit Function
    End If
    
    'This Single-type array will consistently be updated to point to the current line of pixels in the image (RGBF format, remember!)
    Dim srcImageData() As Single
    Dim srcSA As SafeArray1D
    
    'Create a 24bpp or 32bpp DIB at the same size as the image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    If (fi_DataType = FIT_RGBF) Then
        tmpDIB.CreateBlank FreeImage_GetWidth(fi_Handle), FreeImage_GetHeight(fi_Handle), 24
    Else
        tmpDIB.CreateBlank FreeImage_GetWidth(fi_Handle), FreeImage_GetHeight(fi_Handle), 32
    End If
    
    'Point a byte array at the temporary DIB
    Dim dstImageData() As Byte, tmpSA As SafeArray2D
    tmpDIB.WrapArrayAroundDIB dstImageData, tmpSA
        
    'Iterate through each scanline in the source image, copying it to destination as we go.
    Dim iWidth As Long, iHeight As Long, iHeightInv As Long, iScanWidth As Long
    iWidth = FreeImage_GetWidth(fi_Handle) - 1
    iHeight = FreeImage_GetHeight(fi_Handle) - 1
    iScanWidth = FreeImage_GetPitch(fi_Handle)
    
    Dim pxSize As Long
    If (fi_DataType = FIT_RGBF) Then pxSize = 3 Else pxSize = 4
    
    'Shift the parameter values into module-level variables; this is necessary because the actual filmic tone-map function
    ' is standalone, and we don't want to be passing a crapload of Double-type variables to it for every channel of
    ' every pixel in the (presumably large) image.
    m_shoulderStrength = shoulderStrength
    m_linearStrength = linearStrength
    m_linearAngle = linearAngle
    m_toeStrength = toeStrength
    m_toeNumerator = toeNumerator
    m_toeDenominator = toeDenominator
    m_linearWhitePoint = linearWhitePoint
    m_toeAngle = toeNumerator / toeDenominator
    
    'In advance, calculate the filmic function for the white point
    Dim fWhitePoint As Double
    fWhitePoint = fFilmicTonemap(m_linearWhitePoint)
    If (fWhitePoint <> 0#) Then fWhitePoint = 1# / fWhitePoint
    
    'Prep any other post-processing adjustments
    Dim gammaCorrection As Double
    gammaCorrection = 1# / newGamma
    
    'Due to the potential math involved in conversion (if gamma and other settings are being toggled), we need a lot of intermediate variables.
    ' Depending on the user's settings, some of these may go unused.
    Dim rSrcF As Single, gSrcF As Single, bSrcF As Single
    Dim rDstF As Single, gDstF As Single, bDstF As Single
    Dim rDstL As Long, gDstL As Long, bDstL As Long
    
    'Alpha is also a possibility, but unlike RGB values, we assume it is always normalized.  This allows us to skip any intermediate processing,
    ' and simply copy the value directly into the destination (after redistributing to the proper range, of course).
    Dim aDstF As Double, aDstL As Long
    
    Dim x As Long, y As Long, xStride As Long
    
    For y = 0 To iHeight
    
        'FreeImage DIBs are stored bottom-up; we invert them during processing
        iHeightInv = iHeight - y
    
        'Update our scanline pointer
        VBHacks.WrapArrayAroundPtr_Float srcImageData, srcSA, FreeImage_GetScanline(fi_Handle, y), iScanWidth * 4
        
        'Iterate through this line, converting values as we go
        For x = 0 To iWidth
            
            'Retrieve the source values.
            xStride = x * pxSize
            rSrcF = srcImageData(xStride)
            gSrcF = srcImageData(xStride + 1)
            bSrcF = srcImageData(xStride + 2)
            If (pxSize = 4) Then aDstF = srcImageData(xStride + 3)
            
            'Apply filmic tone-mapping.  See http://fr.slideshare.net/ozlael/hable-john-uncharted2-hdr-lighting for details
            rDstF = fFilmicTonemap(exposureCompensation * rSrcF) * fWhitePoint
            gDstF = fFilmicTonemap(exposureCompensation * gSrcF) * fWhitePoint
            bDstF = fFilmicTonemap(exposureCompensation * bSrcF) * fWhitePoint
                                    
            'Apply gamma now (if any).  Unfortunately, lookup tables aren't an option because we're dealing with floating-point input,
            ' so this step is a little slow due to the exponent operator.
            If (newGamma <> 1!) Then
                If (rDstF > 0!) Then rDstF = rDstF ^ gammaCorrection
                If (gDstF > 0!) Then gDstF = gDstF ^ gammaCorrection
                If (bDstF > 0!) Then bDstF = bDstF ^ gammaCorrection
            End If
            
            'Apply failsafe range checks
            If (rDstF < 0!) Then rDstF = 0!
            If (rDstF > 1!) Then rDstF = 1!
                
            If (gDstF < 0!) Then gDstF = 0!
            If (gDstF > 1!) Then gDstF = 1!
                
            If (bDstF < 0!) Then bDstF = 0!
            If (bDstF > 1!) Then bDstF = 1!
            
            'Handle alpha, if necessary
            If (pxSize = 4) Then
                If (aDstF > 1!) Then aDstF = 1!
                If (aDstF < 0!) Then aDstF = 0!
                aDstL = aDstF * 255
            End If
            
            'Calculate corresponding integer values on the range [0, 255]
            rDstL = Int(rDstF * 255! + 0.5)
            gDstL = Int(gDstF * 255! + 0.5)
            bDstL = Int(bDstF * 255! + 0.5)
                        
            'Copy the final, safe values into the destination
            dstImageData(xStride, iHeightInv) = bDstL
            dstImageData(xStride + 1, iHeightInv) = gDstL
            dstImageData(xStride + 2, iHeightInv) = rDstL
            If (pxSize = 4) Then dstImageData(xStride + 3, iHeightInv) = aDstL
            
        Next x
        
    Next y
    
    'Free our 1D array reference
    VBHacks.UnwrapArrayFromPtr_Float srcImageData
        
    'Point dstImageData() away from the DIB and deallocate it
    tmpDIB.UnwrapArrayFromDIB dstImageData
    
    'Create a FreeImage object from our pdDIB object, then release our pdDIB copy
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.GetDIBDC)
    
    'Success!
    ToneMapFilmic_RGBFTo24bppDIB = fi_DIB

End Function

'Filmic tone-map function
Private Function fFilmicTonemap(ByVal x As Single) As Single
    
    'In advance, calculate the filmic function for the white point
    Dim numFunction As Single, denFunction As Single
    
    numFunction = x * (m_shoulderStrength * x + m_linearStrength * m_linearAngle) + m_toeStrength * m_toeNumerator
    denFunction = x * (m_shoulderStrength * x + m_linearStrength) + m_toeStrength * m_toeDenominator
    
    'Failsafe check for DBZ errors
    If (denFunction > 0!) Then
        fFilmicTonemap = (numFunction / denFunction) - m_toeAngle
    Else
        fFilmicTonemap = 1!
    End If
    
End Function

'Returns TRUE if an RGBF format FreeImage DIB contains values outside the [0, 1] range (meaning normalization is required).
' If normalization is required, the various min and max parameters will be filled for each channel.  It is up to the caller to determine how
' these values are used; this function is only diagnostic.
Private Function IsNormalizeRequired(ByVal fi_Handle As Long, ByRef dstMinR As Double, ByRef dstMaxR As Double, ByRef dstMinG As Double, ByRef dstMaxG As Double, ByRef dstMinB As Double, ByRef dstMaxB As Double) As Boolean
    
    'Before doing anything, check the incoming fi_Handle.  If alpha is present, pixel alignment calculations must be modified.
    Dim fi_DataType As FREE_IMAGE_TYPE
    fi_DataType = FreeImage_GetImageType(fi_Handle)
    
    'Values within the [0, 1] range are considered normal.  Values outside this range are not normal, and normalization is thus required.
    ' Because an image does not have to include 0 or 1 values specifically, we return TRUE exclusively; e.g. if any value falls outside
    ' the [0, 1] range, normalization is required.
    Dim minR As Single, maxR As Single, minG As Single, maxG As Single, minB As Single, maxB As Single
    minR = 1: minG = 1: minB = 1
    maxR = 0: maxG = 0: maxB = 0
    
    'This Single-type array will consistently be updated to point to the current line of pixels in the image (RGBF format, remember!)
    Dim srcImageData() As Single, srcSA As SafeArray1D
    
    'Iterate through each scanline in the source image, checking normalize parameters as we go.
    Dim iWidth As Long, iHeight As Long, iScanWidth As Long
    iWidth = FreeImage_GetWidth(fi_Handle) - 1
    iHeight = FreeImage_GetHeight(fi_Handle) - 1
    iScanWidth = FreeImage_GetPitch(fi_Handle)
    
    Dim pxSize As Long
    If (fi_DataType = FIT_RGBF) Then pxSize = 3 Else pxSize = 4
    
    Dim srcR As Single, srcG As Single, srcB As Single
    Dim x As Long, y As Long, xStride As Long
    
    For y = 0 To iHeight
        
        'Update the scanline pointer
        VBHacks.WrapArrayAroundPtr_Float srcImageData, srcSA, FreeImage_GetScanline(fi_Handle, y), iScanWidth * 4
        
        'Iterate through this line, checking values as we go
        For x = 0 To iWidth
            
            xStride = x * pxSize
            
            srcR = srcImageData(xStride)
            srcG = srcImageData(xStride + 1)
            srcB = srcImageData(xStride + 2)
            
            'Check max/min values independently for each channel
            If (srcR < minR) Then minR = srcR
            If (srcR > maxR) Then maxR = srcR
            
            If (srcG < minG) Then minG = srcG
            If (srcG > maxG) Then maxG = srcG
            
            If (srcB < minB) Then minB = srcB
            If (srcB > maxB) Then maxB = srcB
            
        Next x
        
    Next y
    
    'Free our 1D array reference
    VBHacks.UnwrapArrayFromPtr_Float srcImageData
    
    'Fill min/max RGB values regardless of normalization
    dstMinR = minR
    dstMaxR = maxR
    dstMinG = minG
    dstMaxG = maxG
    dstMinB = minB
    dstMaxB = maxB
    
    'If the max or min lie outside the image, notify the caller that normalization is required on this image
    If (minR < 0#) Or (maxR > 1#) Or (minG < 0#) Or (maxG > 1#) Or (minB < 0#) Or (maxB > 1#) Then
        IsNormalizeRequired = True
    Else
        IsNormalizeRequired = False
    End If
    
End Function

'FreeImage supports a user-defined callback for library errors.  We use this and store errors in a pdStringStack object.
Public Sub InitializeFICallback()

    'Ensure library is available before proceeding
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    
    Set m_Errors = New pdStringStack
    FreeImage_SetOutputMessage AddressOf FreeImage_ErrorHandler
    
End Sub

'If InitializeFICallback(), above, was called, FreeImage will supply output messages (typically errors only)
' to this function.
Private Sub FreeImage_ErrorHandler(ByVal imgFormat As FREE_IMAGE_FORMAT, ByVal ptrMessage As Long)
    
    Dim strErrorMessage As String
    If (ptrMessage <> 0) Then
        strErrorMessage = Strings.TrimNull(Strings.StringFromCharPtr(ptrMessage, False))
    Else
        strErrorMessage = "unknown error (blank *char)"
    End If
    
    If (LenB(strErrorMessage) <> 0) Then
        
        If (m_Errors Is Nothing) Then Set m_Errors = New pdStringStack
        
        'Avoid duplicates when adding error strings.  (A single bad image may return the same problem
        ' multiple times, and we don't want to flood the user with useless crap.)
        If (Not m_Errors.ContainsString(strErrorMessage, True)) Then m_Errors.AddString strErrorMessage
        
    End If
    
    'In nightly builds, it can be helpful to see error messages *immediately* instead of saving a list
    ' and presenting them to the user all at once.
    PDDebug.LogAction "FreeImage reported an internal error: " & strErrorMessage, PDM_External_Lib
    
End Sub

Public Function FreeImageErrorState() As Boolean
    If (Not m_Errors Is Nothing) Then
        FreeImageErrorState = (m_Errors.GetNumOfStrings > 0)
    Else
        FreeImageErrorState = False
    End If
End Function

Public Function GetFreeImageErrors(Optional ByVal eraseListUponReturn As Boolean = True) As String
    
    Const DOUBLE_QUOTES As String = """"
    
    Dim listOfErrors As pdString
    Set listOfErrors = New pdString
    listOfErrors.Append DOUBLE_QUOTES
    
    'To simplify our life, instantiate the error stack if it wasn't already.
    If (m_Errors Is Nothing) Then Set m_Errors = New pdStringStack
    
    'Condense all recorded errors into a single multi-line string
    If (m_Errors.GetNumOfStrings > 0) Then
        Dim i As Long
        For i = 0 To m_Errors.GetNumOfStrings - 1
            If (i < m_Errors.GetNumOfStrings - 1) Then
                listOfErrors.AppendLine m_Errors.GetString(i)
            Else
                listOfErrors.Append m_Errors.GetString(i)
            End If
        Next i
    End If
    
    listOfErrors.Append DOUBLE_QUOTES
    GetFreeImageErrors = listOfErrors.ToString()
    
    If eraseListUponReturn Then m_Errors.ResetStack
    
End Function

'Need a FreeImage object at a specific color depth?  Use this function.
'
'The source DIB will not be modified by this function, but some settings require us to make a copy of the source DIB.
' (Non-standard alpha settings are the primary culprit, as we have to handle those conversions internally.)
'
'Obviously, you must manually free the returned FreeImage handle when you're done with it.
'
'Some combinations of parameters are not valid; for example, alphaState and outputColorDepth must be mixed carefully
' (you cannot set binary or color-based alpha for 32-bpp color mode).  For additional details, please refer to the
' ImageExporter module, which goes over these limitations in detail.
'
'Also, please note that this function does not change alpha premultiplication.  The caller needs to handle this in advance.
'
'Finally, this function does not run heuristics on the incoming image.  For example, if you tell it to create a
' grayscale image, it *will* create a grayscale image, regardless of the input.  As such, you must run any
' "auto-convert to best depth" heuristics *prior* to calling this function!
'
'Returns: a non-zero FI handle if successful; 0 if something goes horribly wrong.
Public Function GetFIDib_SpecificColorMode(ByRef srcDIB As pdDIB, ByVal outputColorDepth As Long, Optional ByVal desiredAlphaState As PD_ALPHA_STATUS = PDAS_ComplicatedAlpha, Optional ByVal currentAlphaState As PD_ALPHA_STATUS = PDAS_ComplicatedAlpha, Optional ByVal alphaCutoffOrColor As Long = 127, Optional ByVal finalBackColor As Long = vbWhite, Optional ByVal forceGrayscale As Boolean = False, Optional ByVal paletteCount As Long = 256, Optional ByVal RGB16bppUse565 As Boolean = True, Optional ByVal doNotUseFIGrayscale As Boolean = False, Optional ByVal quantMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT, Optional ByRef srcPalette As pdPalette = Nothing) As Long
    
    'Ensure library is available before proceeding
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    
    'If FreeImage is not enabled, exit immediately
    If (Not ImageFormats.IsFreeImageEnabled()) Then
        GetFIDib_SpecificColorMode = 0
        Exit Function
    End If
    
    Dim fi_DIB As Long, tmpFIHandle As Long
    
    'Perform a quick check for 32-bpp images with complex alpha; we can return those immediately!
    If (outputColorDepth = 32) And (desiredAlphaState = PDAS_ComplicatedAlpha) And (srcDIB.GetDIBColorDepth = 32) And (Not forceGrayscale) Then
        GetFIDib_SpecificColorMode = FreeImage_CreateFromDC(srcDIB.GetDIBDC)
        Exit Function
    End If
    
    'Before proceeding, we first need to manually correct conditions that FreeImage cannot currently meet.
    ' Most significant among these is the combination of grayscale images + alpha channels; these must be forcibly expanded
    ' to RGBA at a matching bit-depth.
    If forceGrayscale And (desiredAlphaState <> PDAS_NoAlpha) Then
        
        'FreeImage supports the notion of 8-bpp images with a single transparent color.
        If (outputColorDepth <= 8) And ((desiredAlphaState = PDAS_BinaryAlpha) Or (desiredAlphaState = PDAS_NewAlphaFromColor)) Then
            
            'This output is now supported.  We just need to make sure we don't use FreeImage's default grayscale path
            ' (as it can't handle alpha correctly).
            doNotUseFIGrayscale = True
            
        'Other gray + transparency options are not currently supported.
        Else
        
            If (outputColorDepth <= 8) And (desiredAlphaState = PDAS_ComplicatedAlpha) Then
                
                'Make sure we do not use the default FreeImage grayscale path; instead, we'll use PD's custom-managed solution
                doNotUseFIGrayscale = True
            
            Else
                'Expand to full RGBA color depths as necessary.
                If (outputColorDepth = 16) Then
                    outputColorDepth = 64
                ElseIf (outputColorDepth = 32) Then
                    outputColorDepth = 128
                End If
            End If
        
        End If
        
    End If
    
    'Some modifications require us to preprocess the incoming image; because this function cannot modify the
    ' incoming DIB, we must use a temporary copy.
    Dim tmpDIBRequired As Boolean: tmpDIBRequired = False
    
    'Some modifications require us to generate a temporary transparency table.  This byte array contains new alpha
    ' values for a given image.  We do not apply these values until *after* an image has been converted to 8-bpp.
    Dim transparencyTableActive As Boolean: transparencyTableActive = False
    Dim transparencyTableBackup As PD_ALPHA_STATUS: transparencyTableBackup = PDAS_ComplicatedAlpha
    Dim tmpTransparencyTable() As Byte
    
    'The order of operations here is a bit tricky.  First, we need to deal with the problem of specialized alpha modes.
    
    'If the caller does not want alpha in the final image, composite against the given backcolor now
    If (desiredAlphaState = PDAS_NoAlpha) Then
        ResetExportPreviewDIB tmpDIBRequired, srcDIB
        m_ExportPreviewDIB.CompositeBackgroundColor Colors.ExtractRed(finalBackColor), Colors.ExtractGreen(finalBackColor), Colors.ExtractBlue(finalBackColor)
    
    'The color-based alpha mode requires us to leave the image in 32-bpp mode, but force it to use only "0" or "255"
    ' alpha values, with a specified transparent color providing the guide for which pixels get turned transparent.
    ' As part of this process, we will create a new transparency map for the image, but when that map gets applied
    ' to the image varies depending on the output color depth.
    
    'This mode must be handled early, because it requires custom PD code, and subsequent quantization (e.g. 8-bit mode) will
    ' yield incorrect results if we attempt to process the transparent color post-quantization (as multiple pixels
    ' may get forced to the same color value, creating new areas of unwanted transparency!).
    ElseIf (desiredAlphaState = PDAS_NewAlphaFromColor) Then
        
        'Apply new alpha.  (This function will return TRUE if the color match is found, and the resulting image thus
        ' contains some amount of transparency.)
        If DIBs.MakeColorTransparent_Ex(srcDIB, tmpTransparencyTable, alphaCutoffOrColor) Then
            
            'If the output color depth is 32-bpp, apply the new transparency table immediately
            If (outputColorDepth > 8) Then
                ResetExportPreviewDIB tmpDIBRequired, srcDIB
                DIBs.ApplyBinaryTransparencyTable m_ExportPreviewDIB, tmpTransparencyTable, finalBackColor
                currentAlphaState = PDAS_BinaryAlpha
            
            'If the output color depth is 8-bpp, note that we need to re-apply the transparency table *after* quantization
            Else
                transparencyTableActive = True
            End If
        
        'If the MakeColorTransparent_Ex function failed, no color matches are found; this lets us use 24-bpp output.
        Else
            ResetExportPreviewDIB tmpDIBRequired, srcDIB
            m_ExportPreviewDIB.CompositeBackgroundColor Colors.ExtractRed(finalBackColor), Colors.ExtractGreen(finalBackColor), Colors.ExtractBlue(finalBackColor)
            desiredAlphaState = PDAS_NoAlpha
        End If
        
    'Binary alpha values require us to leave the image in 32-bpp mode, but force it to use only "0" or "255" alpha values.
    ' Depending on the output color depth, we may not apply the new alpha until after subsequent steps.
    ElseIf (desiredAlphaState = PDAS_BinaryAlpha) Then
        
        'A cutoff of zero means all pixels are gonna be opaque
        If (alphaCutoffOrColor = 0) Then
            ResetExportPreviewDIB tmpDIBRequired, srcDIB
            m_ExportPreviewDIB.CompositeBackgroundColor Colors.ExtractRed(finalBackColor), Colors.ExtractGreen(finalBackColor), Colors.ExtractBlue(finalBackColor)
            desiredAlphaState = PDAS_NoAlpha
        Else
            
            'If the image doesn't already have binary alpha, apply it now
            If (currentAlphaState <> PDAS_BinaryAlpha) Or (outputColorDepth < 32) Then
                
                DIBs.ApplyAlphaCutoff_Ex srcDIB, tmpTransparencyTable, alphaCutoffOrColor
                
                'If the output color depth is 32-bpp, apply the new transparency table immediately
                If (outputColorDepth > 8) Then
                    ResetExportPreviewDIB tmpDIBRequired, srcDIB
                    DIBs.ApplyBinaryTransparencyTable m_ExportPreviewDIB, tmpTransparencyTable, finalBackColor
                    currentAlphaState = PDAS_BinaryAlpha
                
                'If the output color depth is 8-bpp, note that we need to re-apply the transparency table *after* quantization
                Else
                    transparencyTableActive = True
                End If
                
            End If
            
        End If
        
    End If
    
    'If the caller wants a grayscale image AND an alpha channel, we must apply a grayscale conversion now,
    ' as FreeImage does not internally support the notion of 8-bpp grayscale + alpha.
    '
    '(If the caller does not want alpha in the final image, FreeImage will handle the grayscale conversion
    ' internally, so we can skip this step entirely.)
    If forceGrayscale And ((desiredAlphaState <> PDAS_NoAlpha) Or (paletteCount <> 256) Or (outputColorDepth > 16) Or (doNotUseFIGrayscale)) Then
        ResetExportPreviewDIB tmpDIBRequired, srcDIB
        DIBs.MakeDIBGrayscale m_ExportPreviewDIB, paletteCount
    End If
    
    'Next, figure out scenarios where we can pass FreeImage a 24-bpp image.  This is helpful because FreeImage is
    ' unreliable when working with certain types of 32-bpp data (e.g. downsampling 32-bpp data to 8-bpp).
    Dim reduceTo24bpp As Boolean: reduceTo24bpp = False
    If (desiredAlphaState = PDAS_NoAlpha) Then reduceTo24bpp = True
    If (outputColorDepth = 24) Or (outputColorDepth = 48) Or (outputColorDepth = 96) Then reduceTo24bpp = True
    
    'We will also forcibly reduce the incoming image to 24bpp if it doesn't contain any meaningful alpha values
    If (Not reduceTo24bpp) Then
        If tmpDIBRequired Then
            If (m_ExportPreviewDIB.GetDIBColorDepth = 32) Then reduceTo24bpp = DIBs.IsDIBAlphaBinary(m_ExportPreviewDIB, False)
        Else
            If (srcDIB.GetDIBColorDepth = 32) Then reduceTo24bpp = DIBs.IsDIBAlphaBinary(srcDIB, False)
        End If
    End If
    
    'Finally, binary alpha modes + indexed color modes also require us to perform a 24-bpp reduction now.
    If ((desiredAlphaState = PDAS_BinaryAlpha) Or (desiredAlphaState = PDAS_NewAlphaFromColor)) And (outputColorDepth <= 8) And transparencyTableActive Then
        transparencyTableBackup = desiredAlphaState
        reduceTo24bpp = True
    End If
    
    'If any of the 24-bpp criteria are met, apply a forcible conversion now
    If reduceTo24bpp Then
        
        ResetExportPreviewDIB tmpDIBRequired, srcDIB
        
        'Forcibly remove alpha from the image
        If (outputColorDepth < 32) Or (outputColorDepth = 48) Or (outputColorDepth = 96) Then
            If Not m_ExportPreviewDIB.ConvertTo24bpp(finalBackColor) Then FI_DebugMsg "WARNING!  GetFIDib_SpecificColorMode could not convert the incoming DIB to 24-bpp."
        Else
            m_ExportPreviewDIB.CompositeBackgroundColor Colors.ExtractRed(finalBackColor), Colors.ExtractGreen(finalBackColor), Colors.ExtractBlue(finalBackColor)
        End If
        
        'Reset the target alpha mode, so we can tell FreeImage that alpha handling is not required for this image
        desiredAlphaState = PDAS_NoAlpha
        
    End If
    
    'If binary alpha is in use, we must forcibly reset the desired alpha tracker to our desired mode (which will have
    ' been reset by the 24-bpp reduction section, above)
    If transparencyTableActive Then desiredAlphaState = transparencyTableBackup
    
    'Create a FreeImage handle that points to our source image
    If tmpDIBRequired Then
        fi_DIB = FreeImage_CreateFromDC(m_ExportPreviewDIB.GetDIBDC)
    Else
        fi_DIB = FreeImage_CreateFromDC(srcDIB.GetDIBDC)
    End If
    
    If (fi_DIB = 0) Then FI_DebugMsg "WARNING!  Plugin_FreeImage.GetFIDib_SpecificColorMode() failed to create a valid handle from the incoming image!"
    
    'From this point forward, we must operate *only* on this fi_DIB handle.
    
    '1-bpp is easy; handle it now
    If (outputColorDepth = 1) Then
        tmpFIHandle = FreeImage_Dither(fi_DIB, FID_FS)
        If (tmpFIHandle <> fi_DIB) Then
            FreeImage_Unload fi_DIB
            fi_DIB = tmpFIHandle
        End If
    
    'Non-1-bpp is harder
    Else
        
        'Handle grayscale, non-alpha variants first; they use their own dedicated conversion functions
        If (forceGrayscale And (desiredAlphaState = PDAS_NoAlpha) And (Not doNotUseFIGrayscale)) Then
            fi_DIB = GetGrayscaleFIDib(fi_DIB, outputColorDepth)
        
        'Non-grayscale variants (or grayscale variants + alpha) are more complicated
        Else
        
            'Start with non-alpha color modes.  They are easier to handle.
            ' (Also note that this step will *only* be triggered if forceGrayscale = False; the combination of
            '  "forceGrayscale = True" and "PDAS_NoAlpha" is handled by its own If branch, above.)
            If (desiredAlphaState = PDAS_NoAlpha) Then
                
                'Walk down the list of valid outputs, starting at the low end
                If (outputColorDepth <= 8) Then
                    
                    'If the caller wants us to use a palette that *they* have supplied, we need to honor that now.
                    Dim srcQuads() As RGBQuad, srcColorCount As Long
                    If (Not srcPalette Is Nothing) Then
                        If srcPalette.CopyPaletteToArray(srcQuads) Then srcColorCount = UBound(srcQuads) + 1
                    End If
                    
                    'FreeImage supports a new "lossless" quantization method that is perfect for images that already
                    ' have 256 colors or less.  This method is basically just a hash table, and it lets us avoid
                    ' lossy quantization if at all possible.
                    If (paletteCount = 256) And (srcPalette Is Nothing) Then
                        tmpFIHandle = FreeImage_ColorQuantize(fi_DIB, FIQ_LFPQUANT)
                    Else
                        If (srcPalette Is Nothing) Then
                            tmpFIHandle = FreeImage_ColorQuantizeExInt(fi_DIB, FIQ_LFPQUANT, paletteCount)
                        Else
                            tmpFIHandle = FreeImage_ColorQuantizeExInt(fi_DIB, FIQ_LFPQUANT, srcColorCount, srcColorCount, VarPtr(srcQuads(0)))
                        End If
                    End If
                    
                    '0 means the image has > 256 colors, and must be quantized via lossy means
                    If (tmpFIHandle = 0) Then
                        
                        If (quantMethod = FIQ_LFPQUANT) Then quantMethod = FIQ_WUQUANT
                        
                        'If the caller has specified which palette to use, honor that now
                        If (Not srcPalette Is Nothing) Then
                            tmpFIHandle = FreeImage_ColorQuantizeExInt(fi_DIB, FIQ_NNQUANT, srcColorCount, srcColorCount, VarPtr(srcQuads(0)))
                        Else
                        
                            'If we're going straight to 4-bits, ignore an optimal palette count in favor of a 16-color one.
                            If (outputColorDepth = 4) Then
                                tmpFIHandle = FreeImage_ColorQuantizeExInt(fi_DIB, quantMethod, 16)
                            Else
                                If (paletteCount = 256) Then
                                    tmpFIHandle = FreeImage_ColorQuantize(fi_DIB, quantMethod)
                                Else
                                    tmpFIHandle = FreeImage_ColorQuantizeExInt(fi_DIB, quantMethod, paletteCount)
                                End If
                            End If
                            
                        End If
                        
                    End If
                    
                    If (tmpFIHandle <> fi_DIB) Then
                        
                        If (tmpFIHandle <> 0) Then
                            FreeImage_Unload fi_DIB
                            fi_DIB = tmpFIHandle
                        Else
                            FI_DebugMsg "WARNING!  tmpFIHandle is zero!"
                        End If
                        
                    End If
                    
                    'We now have an 8-bpp image.  Forcibly convert to 4-bpp if necessary.
                    If (outputColorDepth = 4) Then
                        tmpFIHandle = FreeImage_ConvertTo4Bits(fi_DIB)
                        If (tmpFIHandle <> fi_DIB) And (tmpFIHandle <> 0) Then
                            FreeImage_Unload fi_DIB
                            fi_DIB = tmpFIHandle
                        End If
                    End If
                    
                'Some bit-depth > 8
                Else
                
                    '15- and 16- are handled similarly
                    If (outputColorDepth = 15) Or (outputColorDepth = 16) Then
                        
                        If (outputColorDepth = 15) Then
                            tmpFIHandle = FreeImage_ConvertTo16Bits555(fi_DIB)
                        Else
                            If RGB16bppUse565 Then
                                tmpFIHandle = FreeImage_ConvertTo16Bits565(fi_DIB)
                            Else
                                tmpFIHandle = FreeImage_ConvertTo16Bits555(fi_DIB)
                            End If
                        End If
                        
                        If (tmpFIHandle <> fi_DIB) Then
                            FreeImage_Unload fi_DIB
                            fi_DIB = tmpFIHandle
                        End If
                        
                    'Some bit-depth > 16
                    Else
                        
                        '24-bpp doesn't need to be handled, because it is the default entry point for PD images
                        If (outputColorDepth > 24) And (outputColorDepth <> 32) Then
                        
                            'High bit-depth variants are covered last
                            If (outputColorDepth = 48) Then
                                tmpFIHandle = FreeImage_ConvertToRGB16(fi_DIB)
                                If (tmpFIHandle <> fi_DIB) Then
                                    FreeImage_Unload fi_DIB
                                    fi_DIB = tmpFIHandle
                                End If
                                
                            '96-bpp is the only other possibility
                            Else
                                tmpFIHandle = FreeImage_ConvertToRGBF(fi_DIB)
                                If (tmpFIHandle <> fi_DIB) Then
                                    FreeImage_Unload fi_DIB
                                    fi_DIB = tmpFIHandle
                                End If
                            End If
                        
                        End If
                        
                    End If
                
                End If
            
            'The image contains alpha, and the caller wants alpha in the final image.
            
            '(Note also that forceGrayscale may or may not be TRUE, but a grayscale conversion will have been applied by a
            ' previous step, so we can safely ignore its value here.  This is necessary because FreeImage does not internally
            ' support the concept of gray+alpha images - they must be expanded to RGBA.)
            Else
                
                'Skip 32-bpp outputs, as the image will already be in that depth by default!
                If (outputColorDepth <> 32) Then
                
                    '(FYI: < 32-bpp + alpha is the ugliest conversion we handle)
                    If (outputColorDepth < 32) Then
                        
                        Dim paletteCheck(0 To 255) As Byte
                        
                        Dim fiPixels() As Byte, srcSA As SafeArray1D
                        
                        Dim iWidth As Long, iHeight As Long, iScanWidth As Long
                        Dim x As Long, y As Long, transparentIndex As Long
                        
                        'PNG is the output format that gives us the most grief here, because it supports so many different
                        ' transparency formats.  We have to manually work around formats not supported by FreeImage,
                        ' which means an unpleasant amount of custom code.
                        
                        'First, we start by getting the image into 8-bpp color mode.  How we do this varies by transparency type.
                        ' 1) Images with full transparency need to be quantized, then converted to back to 32-bpp mode.
                        '    We will manually plug-in the correct alpha bytes post-quantization.
                        ' 2) Images with binary transparency need to be quantized to 255 colors or less.  The image can stay
                        '    in 8-bpp mode; we will fill the first empty palette index with transparency, and update the
                        '    image accordingly.
                        
                        'Full transparency is desired in the final image
                        If (desiredAlphaState = PDAS_ComplicatedAlpha) Then
                            
                            'Start by backing up the image's current transparency data.
                            DIBs.RetrieveTransparencyTable srcDIB, tmpTransparencyTable
                            
                            'Fix premultiplication
                            ResetExportPreviewDIB tmpDIBRequired, srcDIB
                            Dim resetAlphaPremultiplication As Boolean: resetAlphaPremultiplication = False
                            If m_ExportPreviewDIB.GetAlphaPremultiplication Then
                                resetAlphaPremultiplication = True
                                m_ExportPreviewDIB.ConvertTo24bpp finalBackColor
                            End If
                            
                            FreeImage_Unload fi_DIB
                            fi_DIB = FreeImage_CreateFromDC(m_ExportPreviewDIB.GetDIBDC)
                            
                            'Quantize the image (using lossless means, if possible)
                            tmpFIHandle = FreeImage_ColorQuantizeExInt(fi_DIB, FIQ_LFPQUANT, paletteCount)
                            
                            '0 means the image has > 255 colors, and must be quantized via lossy means
                            If (quantMethod = FIQ_LFPQUANT) Then quantMethod = FIQ_WUQUANT
                            If (tmpFIHandle = 0) Then tmpFIHandle = FreeImage_ColorQuantizeExInt(fi_DIB, quantMethod, paletteCount)
                            
                            'Regardless of what quantization method was applied, update our pointer to point at the new
                            ' 8-bpp copy of the source image.
                            If (tmpFIHandle <> fi_DIB) Then
                                If (tmpFIHandle <> 0) Then
                                    FreeImage_Unload fi_DIB
                                    fi_DIB = tmpFIHandle
                                Else
                                    FI_DebugMsg "WARNING!  FreeImage failed to quantize the original fi_DIB into a valid 8-bpp version."
                                End If
                            End If
                            
                            'fi_DIB now points at an 8-bpp image.  Upsample it to 32-bpp.
                            tmpFIHandle = FreeImage_ConvertTo32Bits(fi_DIB)
                            If (tmpFIHandle <> fi_DIB) Then
                                If (tmpFIHandle <> 0) Then
                                    FreeImage_Unload fi_DIB
                                    fi_DIB = tmpFIHandle
                                Else
                                    FI_DebugMsg "WARNING!  FreeImage failed to convert the quantized fi_DIB into a valid 32-bpp version."
                                End If
                            End If
                            
                            'Next, we need to copy our 32-bpp data over FreeImage's 32-bpp data.
                            iWidth = FreeImage_GetWidth(fi_DIB) - 1
                            iHeight = FreeImage_GetHeight(fi_DIB) - 1
                            iScanWidth = FreeImage_GetPitch(fi_DIB)
                            
                            For y = 0 To iHeight
                                
                                'Point a 1D VB array at this scanline
                                VBHacks.WrapArrayAroundPtr_Byte fiPixels, srcSA, FreeImage_GetScanline(fi_DIB, y), iScanWidth
                                
                                'Iterate through this line, copying over new transparency indices as we go
                                For x = 0 To iWidth
                                    fiPixels(x * 4 + 3) = tmpTransparencyTable(x, iHeight - y)
                                Next x
                                
                            Next y
                            
                            'Free our 1D array reference
                            VBHacks.UnwrapArrayFromPtr_Byte fiPixels
                            
                            'We now have a 32-bpp image with quantized RGB values, but intact A values.
                            If resetAlphaPremultiplication Then FreeImage_PreMultiplyWithAlpha fi_DIB
                            
                        'Binary transparency is desired in the final image
                        Else
                            
                            'FreeImage supports a new "lossless" quantization method that is perfect for images that already
                            ' have 255 colors or less.  This method is basically just a hash table, and it lets us avoid
                            ' lossy quantization if at all possible.
                            ' (Note that we must forcibly request 255 colors; one color is reserved for the transparent index.)
                            If (paletteCount >= 256) Then paletteCount = 255
                            tmpFIHandle = FreeImage_ColorQuantizeExInt(fi_DIB, FIQ_LFPQUANT, paletteCount)
                            
                            '0 means the image has > 255 colors, and must be quantized via lossy means
                            If (tmpFIHandle = 0) Then
                                If (quantMethod = FIQ_LFPQUANT) Then quantMethod = FIQ_WUQUANT
                                tmpFIHandle = FreeImage_ColorQuantizeExInt(fi_DIB, quantMethod, paletteCount)
                            End If
                            
                            'Regardless of what quantization method was applied, update our pointer to point at the new
                            ' 8-bpp copy of the source image.
                            If (tmpFIHandle <> fi_DIB) Then
                                If (tmpFIHandle <> 0) Then
                                    FreeImage_Unload fi_DIB
                                    fi_DIB = tmpFIHandle
                                Else
                                    FI_DebugMsg "WARNING!  FreeImage failed to quantize the original fi_DIB into a valid 8-bpp version."
                                End If
                            End If
                            
                            'fi_DIB now points at an 8-bpp image.
                            
                            'Next, we need to create our transparent index in the palette.  FreeImage won't reliably tell
                            ' us how much of a palette is currently in use (ugh), so instead, we must manually scan the
                            ' image, looking for unused palette entries.
                            iWidth = FreeImage_GetWidth(fi_DIB) - 1
                            iHeight = FreeImage_GetHeight(fi_DIB) - 1
                            iScanWidth = FreeImage_GetPitch(fi_DIB)
                            
                            For y = 0 To iHeight
                                
                                'Point a 1D VB array at this scanline
                                VBHacks.WrapArrayAroundPtr_Byte fiPixels, srcSA, FreeImage_GetScanline(fi_DIB, y), iScanWidth
                                
                                'Iterate through this line, checking values as we go
                                For x = 0 To iWidth
                                    paletteCheck(fiPixels(x)) = 1
                                Next x
                                
                            Next y
                            
                            'Free our 1D array reference
                            VBHacks.UnwrapArrayFromPtr_Byte fiPixels
                            
                            'Scan through the palette array, looking for the first 0 entry (which means that value was not
                            ' found in the source image).
                            transparentIndex = -1
                            For x = 0 To 255
                                If paletteCheck(x) = 0 Then
                                    transparentIndex = x
                                    Exit For
                                End If
                            Next x
                            
                            'It shouldn't be possible for a 256-entry palette to exist, but if it does, we have no choice
                            ' but to "steal" a palette index for transparency.
                            If (transparentIndex = -1) Then
                                FI_DebugMsg "WARNING!  FreeImage returned a full palette, so transparency will have to steal an existing entry!"
                                transparentIndex = 255
                            End If
                            
                            'Tell FreeImage which palette index we want to use for transparency
                            FreeImage_SetTransparentIndex fi_DIB, transparentIndex
                            
                            'Now that we have a transparent index, we need to update the target image to be transparent
                            ' in all the right locations.  Use our previously generated transparency table for this.
                            For y = 0 To iHeight
                                
                                'Point a 1D VB array at this scanline
                                VBHacks.WrapArrayAroundPtr_Byte fiPixels, srcSA, FreeImage_GetScanline(fi_DIB, y), iScanWidth
                                
                                'The FreeImage DIB will be upside-down at this point
                                For x = 0 To iWidth
                                    If tmpTransparencyTable(x, iHeight - y) = 0 Then fiPixels(x) = transparentIndex
                                Next x
                                
                            Next y
                            
                            VBHacks.UnwrapArrayFromPtr_Byte fiPixels
                            
                            'We now have an < 8-bpp image with a transparent index correctly marked.  Whew!
                            
                        End If
                    
                    'Output is > 32-bpp with transparency
                    Else
                        
                        '64-bpp is 16-bits per channel RGBA
                        If (outputColorDepth = 64) Then
                            tmpFIHandle = FreeImage_ConvertToRGBA16(fi_DIB)
                            If (tmpFIHandle <> fi_DIB) Then
                                FreeImage_Unload fi_DIB
                                fi_DIB = tmpFIHandle
                            End If
                            
                        '128-bpp is the only other possibility (32-bits per channel RGBA, specifically)
                        Else
                            tmpFIHandle = FreeImage_ConvertToRGBAF(fi_DIB)
                            If (tmpFIHandle <> fi_DIB) Then
                                FreeImage_Unload fi_DIB
                                fi_DIB = tmpFIHandle
                            End If
                        End If
                        
                    End If
                
                'End 32-bpp requests
                End If
            'End alpha vs non-alpha
            End If
        'End grayscale vs non-grayscale
        End If
    'End 1-bpp vs > 1-bpp
    End If
    
    GetFIDib_SpecificColorMode = fi_DIB
    
End Function

Private Sub ResetExportPreviewDIB(ByRef trackerBool As Boolean, ByRef srcDIB As pdDIB)
    If (Not trackerBool) Then
        If (m_ExportPreviewDIB Is Nothing) Then Set m_ExportPreviewDIB = New pdDIB
        m_ExportPreviewDIB.CreateFromExistingDIB srcDIB
        trackerBool = True
    End If
End Sub

'Convert an incoming FreeImage handle to a grayscale FI variant.  The source handle will be unloaded as necessary.
Private Function GetGrayscaleFIDib(ByVal fi_DIB As Long, ByVal outputColorDepth As Long) As Long
    
    Dim tmpFIHandle As Long
    
    'Output color depth is important here.  16-bpp and 32-bpp grayscale are actually high bit-depth modes!
    If (outputColorDepth <= 8) Then
        
        'Create an 8-bpp palette
        tmpFIHandle = FreeImage_ConvertToGreyscale(fi_DIB)
        If (tmpFIHandle <> fi_DIB) And (tmpFIHandle <> 0) Then
            FreeImage_Unload fi_DIB
            fi_DIB = tmpFIHandle
        End If
        
        'If the caller wants a 4-bpp palette, do that now.
        If (outputColorDepth = 4) And (tmpFIHandle <> 0) Then
            tmpFIHandle = FreeImage_ConvertTo4Bits(fi_DIB)
                If (tmpFIHandle <> fi_DIB) Then
                FreeImage_Unload fi_DIB
                fi_DIB = tmpFIHandle
            End If
        End If
    
    'Forcing to grayscale and using an outputColorDepth > 8 means you want a high bit-depth copy!
    Else
    
        '32-bpp
        If (outputColorDepth = 32) Then
            tmpFIHandle = FreeImage_ConvertToFloat(fi_DIB)
            If (tmpFIHandle <> fi_DIB) And (tmpFIHandle <> 0) Then
                FreeImage_Unload fi_DIB
                fi_DIB = tmpFIHandle
            End If
        
        'Output color-depth must be 16; any other values are invalid
        Else
            tmpFIHandle = FreeImage_ConvertToUINT16(fi_DIB)
            If (tmpFIHandle <> fi_DIB) And (tmpFIHandle <> 0) Then
                FreeImage_Unload fi_DIB
                fi_DIB = tmpFIHandle
            End If
        End If
        
    End If
    
    GetGrayscaleFIDib = fi_DIB
    
End Function

'Given a source FreeImage handle and FI format, fill a destination DIB with a post-"exported-to-that-format" version of the image.
' This is used to generate the "live previews" used in various "export to lossy format" dialogs.
'
'(Note that you could technically pass a bare DIB to this function, but because different dialogs provide varying levels of control
' over the source image, it's often easier to let the caller handle that step.  That way, they can cache a FI handle in the most
' relevant color depth, shaving previous ms off the actual export+import step.)
Public Function GetExportPreview(ByRef srcFI_Handle As Long, ByRef dstDIB As pdDIB, ByVal dstFormat As PD_IMAGE_FORMAT, Optional ByVal fi_SaveFlags As Long = 0, Optional ByVal fi_LoadFlags As Long = 0) As Boolean
    
    'Ensure library is available before proceeding
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    
    Dim fi_Size As Long
    If FreeImage_SaveToMemoryEx(dstFormat, srcFI_Handle, m_ExportPreviewBytes, fi_SaveFlags, fi_Size) Then
        
        Dim fi_DIB As Long
        fi_DIB = FreeImage_LoadFromMemoryEx(VarPtr(m_ExportPreviewBytes(0)), fi_Size, fi_LoadFlags, dstFormat)
        
        If (fi_DIB <> 0) Then
        
            'Because we're going to do a fast copy operation, we need to flip the FreeImage DIB to match DIB orientation
            FreeImage_FlipVertically fi_DIB
            
            'If a format requires special handling, trigger it here
            If (dstFormat = PDIF_PBM) Or (dstFormat = PDIF_PBMRAW) And (FreeImage_GetBPP(fi_DIB) = 1) Then
                FreeImage_Invert fi_DIB
            End If
            
            'Convert the incoming DIB to a 24-bpp or 32-bpp representation
            If (FreeImage_GetBPP(fi_DIB) <> 24) And (FreeImage_GetBPP(fi_DIB) <> 32) Then
                
                Dim newFI_Handle As Long
                If FreeImage_IsTransparent(fi_DIB) Or (FreeImage_GetTransparentIndex(fi_DIB) <> -1) Then
                    newFI_Handle = FreeImage_ConvertColorDepth(fi_DIB, FICF_RGB_32BPP, False)
                Else
                    newFI_Handle = FreeImage_ConvertColorDepth(fi_DIB, FICF_RGB_24BPP, False)
                End If
                
                If (newFI_Handle <> fi_DIB) Then
                    FreeImage_Unload fi_DIB
                    fi_DIB = newFI_Handle
                End If
                
            End If
            
            'Copy the DIB into a PD DIB object
            If Not Plugin_FreeImage.PaintFIDibToPDDib(dstDIB, fi_DIB, 0, 0, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight) Then
                FI_DebugMsg "WARNING!  Plugin_FreeImage.PaintFIDibToPDDib failed for unknown reasons."
            End If
            
            FreeImage_Unload fi_DIB
            GetExportPreview = True
        Else
            FI_DebugMsg "WARNING!  Plugin_FreeImage.GetExportPreview failed to generate a valid fi_Handle."
            GetExportPreview = False
        End If
        
    Else
        FI_DebugMsg "WARNING!  Plugin_FreeImage.GetExportPreview failed to save the requested handle to an array."
        GetExportPreview = False
    End If
    
End Function

'PD uses a persistent cache for generating post-export preview images.  This costs several MB of memory but greatly improves
' responsiveness of export dialogs.  When such a dialog is unloaded, you can call this function to forcibly reclaim the memory
' associated with that cache.
Public Sub ReleasePreviewCache(Optional ByVal unloadThisFIHandleToo As Long = 0)
    
    'Ensure library is available before proceeding
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    
    Erase m_ExportPreviewBytes
    Set m_ExportPreviewDIB = Nothing
    If (unloadThisFIHandleToo <> 0) Then ReleaseFreeImageObject unloadThisFIHandleToo
    
End Sub

Public Sub ReleaseFreeImageObject(ByVal srcFIHandle As Long)
    If (m_FreeImageHandle = 0) Then InitializeFreeImage True
    FreeImage_Unload srcFIHandle
End Sub

Private Sub FI_DebugMsg(ByVal srcDebugMsg As String, Optional ByVal suppressDebugData As Boolean = False)
    If (Not suppressDebugData) Then PDDebug.LogAction srcDebugMsg, PDM_External_Lib
End Sub
