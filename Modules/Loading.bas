Attribute VB_Name = "Loading"
'***************************************************************************
'General-purpose image and data import interface
'Copyright 2001-2026 by Tanner Helland
'Created: 4/15/01
'Last updated: 18/November/25
'Last update: expand QuickLoad support for recent format support enhancements
'
'This module provides high-level "load" functionality for getting image files into PD.
' There are a number of different ways to do this; for example, loading a user-facing image
' is a horrifically complex affair, with lots of messy work involved in metadata parsing,
' UI prep, Undo/Redo stuff, and more.  Conversely, loading an image file as a resource or
' internal image can bypass many of those steps.
'
'Note that these high-level functions call into a number of lower-level functions inside
' the ImageImporter module, and potentially various third-party libraries.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'If an image load was initiated as part of a multi-image import (like the user dragging a million photos
' from an Explorer window), this will be set to TRUE.  While TRUE, as many import dialogs as possible
' need to be suspended until the *end* of the import process.
Private m_MultiImageLoadActive As Boolean

'This function is used for loading a user-facing image (vs loading an internal PD image).  Loading a user-facing image involves
' a large amount of extra work (like metadata parsing) which we simply don't care about when loading internal resources.
'
'Note that this function will use one of several backends to load a given image; different filetypes are preferentially handled by
' different means, so portions of this function may call into external DLLs for parts of its functionality.  (The interaction between
' this function and various plugins is complex; I recommend studying the separate ImageImporter module for details.)
'
'INPUTS:
' 1) srcFile
'    Fully qualified, absolute path to the source image.  Unicode is supported.
' 2) [optional] suggestedFilename
'    If loading an image from a temp file (e.g. clipboard, scanner), this value will be used in two places:as the window caption
'    (prior to first-save) and as the suggested filename at first-save. As such, make it user-friendly, e.g. "Clipboard image".
'    If this parameter is *not* supplied, the image's current filename will automatically be used.
' 3) [optional] addToRecentFiles
'    When a file loads successfully, we typically add it to the File > Recent Files list.  Some load operations, like
'    "Add new layer from file" (or restoring a file from Autosave) don't easily fit into this paradigm. This value tells the
'    load engine to skip the "add to recent files" step.
' 4) [optional] suspendWarnings
'    At times, the caller may not want to have UI warnings raised for malformed or invalid files.  Batch processing and
'    multi-image load are two examples.  When suspendWarnings = vbYES, any user-facing messages related to bad files are
'    suppressed.  (Note that the warnings can still be retrieved from debug logs, however.)  This value is passed ByRef
'    so that is suspendWarnings is vbNO, the caller can handle vbCancel results (if desired) from raised message boxes.
' 5) [optional] handleUIDisabling
'    By default, this function takes control of PhotoDemon's UI and disables anything interactable while the load process occurs.
'    Some specialized load functions (like batch processing) already assume specialized control, and will not want the load
'    process to re-enable everything when the load completes.
' 6) [optional] overrideParameters
'    During a batch process, normal image import processes (like displaying optional prompts) may be suspended.  This param
'    string can contain custom parameters that can be blindly forwarded to subsequent import operations.  (Basically, it's a
'    catch-all for future improvements and modifications.)
' 7) [optional] numCanceledImports
'    Each time the user cancels an image via import dialog (like a PDF or SVG size dialog), this is incremented by one.
'    (If you want it reset between calls, reset it manually; this function only increments it.)
Public Function LoadFileAsNewImage(ByRef srcFile As String, Optional ByVal suggestedFilename As String = vbNullString, Optional ByVal addToRecentFiles As Boolean = True, Optional ByRef suspendWarnings As VbMsgBoxResult = vbNo, Optional ByVal handleUIDisabling As Boolean = True, Optional ByVal overrideParameters As String = vbNullString, Optional ByRef numCanceledImports As Long = 0) As Boolean
    
    '*** AND NOW, AN IMPORTANT MESSAGE ABOUT DOEVENTS ***
    
    'Normally, PhotoDemon avoids calling DoEvents for all the obvious reasons.
    ' This function is an exception to that rule.
    
    'While this function stays busy loading the image in question, the ExifTool plugin runs asynchronously,
    ' parsing image metadata and forwarding the results to a pdPipeAsync instance on PD's primary form.
    ' By using DoEvents throughout this function (specifically, a custom-built version that only allows
    ' timer messages through, named "VBHacks.DoEventsTimersOnly"), we periodically yield control to that
    ' pdPipeAsync instance, which allows it to clear stdout so ExifTool can continue pushing metadata through.
    ' (If we don't do this, ExifTool will freeze when stdout fills its buffer, which is not just possible but
    ' *probable*, given how much metadata the average photo file contains.)
    
    'That said, please note that a LOT of precautions have been taken to make sure DoEvents doesn't cause reentry
    ' and other issues.  Do *not* mimic this behavior in your own code unless you understand the repercussions!
    
    '*** END MESSAGE ***
    
    'Image loading is a place where many things can go wrong - bad files, corrupt formats, heavy RAM usage,
    ' incompatible color formats, and about a bazillion other problems.  As such, this function dumps a *lot* of
    ' information to the debug log.
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    PDDebug.LogAction "Image load requested for """ & Files.FileGetName(srcFile) & """.  Baseline memory reading:"
    PDDebug.LogAction vbNullString, PDM_Mem_Report
    
    'Note the caller's desire to suspend pop-up warnings for things like missing or broken files.  (Batch processes
    ' request this, for example.)  When this is set to vbNO, we'll pester the user with message boxes on critical errors.
    Dim arePopupsAllowed As Boolean
    arePopupsAllowed = (suspendWarnings = vbNo)
    
    'Before doing anything else, purge the input queue to ensure no stray key or mouse events are "left over"
    ' from things like a common dialog interaction.  (Refer to the DoEvents warning, above, about the precautions
    ' this function takes to ensure no message loop funny business.)
    VBHacks.PurgeInputMessages FormMain.hWnd
    
    'Display a busy cursor
    If handleUIDisabling Then
        Message "Loading image..."
        Processor.MarkProgramBusyState True, True
    End If
    
    '*************************************************************************************************************************************
    ' Prepare all variables related to image loading
    '*************************************************************************************************************************************
    
    'Normally, an unsuccessful load just causes the function to exit prematurely, but sometimes we can't detect an unsuccessful load
    ' until deep into the load process.  When this happens, we may need to roll-back things like memory allocations, so we check success
    ' state quite a few times throughout the function.
    Dim loadSuccessful As Boolean: loadSuccessful = False
    
    'Some behavior varies based on the image decoding engine used.  PD uses a fairly complex cascading system for image decoders;
    ' if one fails, we continue trying alternates until either the load succeeds, or all known decoders have been exhausted.
    Dim decoderUsed As PD_ImageDecoder: decoderUsed = id_Failure
    
    'Some image formats (like TIFF, animated GIF, icons) support the notion of "multiple pages".  PD can detect such images,
    ' and depending on user input, handle the file a few different ways.
    Dim imageHasMultiplePages As Boolean: imageHasMultiplePages = False
    Dim numOfPages As Long: numOfPages = 0
    
    'We now have a few tedious checks to perform: like making sure the file actually exists!
    If (Not Files.FileExists(srcFile)) Then
        If handleUIDisabling Then Processor.MarkProgramBusyState False, True
        If arePopupsAllowed Then
            Message "Warning - file not found: %1", srcFile
            PDMsgBox "Unfortunately, the image '%1' could not be found." & vbCrLf & vbCrLf & "If this image was originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbExclamation Or vbOKOnly, "File not found", srcFile
        End If
        LoadFileAsNewImage = False
        Exit Function
    End If
    
    If (Not Files.FileTestAccess_Read(srcFile)) Then
        If handleUIDisabling Then Processor.MarkProgramBusyState False, True
        If arePopupsAllowed Then
            Message "Warning - file locked: %1", srcFile
            PDMsgBox "Unfortunately, the file '%1' is currently locked by another program on this PC." & vbCrLf & vbCrLf & "Please close this file in any other running programs, then try again.", vbExclamation Or vbOKOnly, "File locked", srcFile
        End If
        LoadFileAsNewImage = False
        Exit Function
    End If
    
    'Now we get into the meat-and-potatoes portion of this sub.  Main segments are labeled by large, asterisk-separated bars.
    ' These segments generally describe a group of tasks with related purpose, and many of these tasks branch out into other modules.
    
    '*************************************************************************************************************************************
    ' If memory usage is a concern, try to suspend some images to disk before proceeding
    '*************************************************************************************************************************************
    
    'A central memory manager handles this operation for us - we just need to notify it to act.
    ' Currently, PD wants to ensure it has about ~100 MB free for a newly loaded image.  That provides
    ' ~20 megapixels worth of space (80 MB) plus a little overhead for the decoding process.
    '
    'If that much memory, plus whatever PD currently has allocated, exceeds 80% of *available* memory,
    ' PD will try to suspend other open images to disk in an attempt to maximize available space.
    LargeAllocationIncoming 100
    
    '*************************************************************************************************************************************
    ' If the image being loaded is a primary image (e.g. one opened normally), prepare a blank pdImage object to receive it
    '*************************************************************************************************************************************
    
    'To prevent re-entry problems, forcibly disable the main form before proceeding further.  Note that any criteria that result in
    ' a premature exit from this function *MUST* reenable the form manually!
    If handleUIDisabling Then FormMain.Enabled = False
    
    'PD has a three-tiered management system for images:
    ' 1) pdImage object: the main object, which holds a stack of one or more layers, and a bunch of image-level data (like filename)
    ' 2) pdLayer object: a layer object which holds a stack of one or more DIBs, and a bunch of layer-level data (like blendmode)
    ' 3) pdDIB object: eventually this will be retitled as pdSurface, as it may not be a DIB, but at present, a single grid of pixels
    
    'Different parts of the load process interact with different levels of our target pdImage object.  If loading a PDI file
    ' (PhotoDemon's native format), multiple layers and DIBs will be loaded and processed for a singular pdImage object.
    
    'Anyway, in the future, I'd like to avoid referencing the pdImages collection directly, and instead use helper functions.
    ' To facilitate this switch, I've written this function to use generic "targetImage" and "targetDIB" objects.  (targetLayer isn't
    ' as important, as most image files only consist of a single default layer inside targetImage.)
    
    'Retrieve an empty, default pdImage object.  Note that this object does not yet exist inside the main pdImages collection,
    ' so we cannot refer to it by ordinal.
    Dim targetImage As pdImage
    PDImages.GetDefaultPDImageObject targetImage
    
    'Normally, we don't assign an ID value to an image until we actually add it to the central
    ' pdImages collection.  However, some tasks (like retrieving metadata asynchronously) require
    ' an ID so we can synchronize incoming data post-load.  Give the target image a provisional
    ' image ID; this ID will become its formal ID only if it loads successfully.
    targetImage.imageID = PDImages.GetProvisionalImageID()
    
    'Next, create a blank target DIB.  Image loaders need a place to stick their decoded image data, and we'll use this
    ' same target DIB regardless of actual parser.
    Dim targetDIB As pdDIB
    Set targetDIB = New pdDIB
    
    '*************************************************************************************************************************************
    ' Make a best guess at the incoming image's format
    '*************************************************************************************************************************************
    
    PDDebug.LogAction "Determining filetype..."
    targetImage.SetOriginalFileFormat PDIF_UNKNOWN
    
    Dim srcFileExtension As String
    srcFileExtension = UCase$(Files.FileGetExtension(srcFile))
    
    Dim internalFormatID As PD_IMAGE_FORMAT
    internalFormatID = CheckForInternalFiles(srcFileExtension)
    
    'Files with a PD-specific format have now been specially marked, while generic files (JPEG, PNG, etc) have not.
    
    
    '*************************************************************************************************************************************
    ' Split handling into two groups: internal PD formats vs generic external formats
    '*************************************************************************************************************************************
    
    'Image load performance is critical to a good user experience.  Profile it and report timings
    ' in the debug log.
    Dim justImageLoadTime As Currency
    VBHacks.GetHighResTime justImageLoadTime
    
    Dim freeImage_Return As PD_OPERATION_OUTCOME
    
    'In recent years, I've tried to support more vector formats in PD.  These formats often require an import dialog
    ' (where the user can control rasterization settings).  If the user cancels these import dialogs, we don't want
    ' to pester them with error dialogs.  Formats that support a user dialog will set this value as necessary.
    Dim userCanceledImportDialog As Boolean
    userCanceledImportDialog = False
    
    If (internalFormatID = PDIF_UNKNOWN) Then
    
        'Note that some formats may raise additional dialogs (e.g. tone-mapping HDR/RAW images, selecting pages
        ' from a PDF), so the loader does not return binary pass/fail state.
        '
        'If the function fails due to user cancellation, we should suppress subsequent error message boxes.
        loadSuccessful = ImageImporter.CascadeLoadGenericImage(srcFile, targetImage, targetDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages, overrideParameters, userCanceledImportDialog, suspendWarnings)
        
        '*************************************************************************************************************************************
        ' If the ExifTool plugin is available and this is a non-PD-specific file, initiate a separate thread for metadata extraction
        '*************************************************************************************************************************************
        If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) And (internalFormatID <> PDIF_PDI) And (internalFormatID <> PDIF_RAWBUFFER) And (Not userCanceledImportDialog) Then
            PDDebug.LogAction "Starting separate metadata extraction thread..."
            ExifTool.StartMetadataProcessing srcFile, targetImage
        End If
    
    'PD-specific files use their own load function, which bypasses a lot of tedious format-detection heuristics
    Else
    
        loadSuccessful = ImageImporter.CascadeLoadInternalImage(internalFormatID, srcFile, targetImage, targetDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
        If (Not loadSuccessful) Then PDDebug.LogAction "WARNING!  LoadFileAsNewImage failed on an internal file; all engines failed to handle " & srcFile & " correctly."
        
    End If
    
    'Eventually, the user may choose to save this image in a new format, but for now, the original and current formats are identical
    targetImage.SetCurrentFileFormat targetImage.GetOriginalFileFormat
    
    PDDebug.LogAction "Format-specific parsing complete.  Running a few failsafe checks on the new pdImage object..."
    
    'Because ExifTool is sending us data in the background, we must periodically yield for metadata piping.
    If (ExifTool.IsMetadataPipeActive) Then VBHacks.DoEventsTimersOnly
    
    
    '*************************************************************************************************************************************
    ' Run a few failsafe checks to confirm that the image data was loaded successfully
    '*************************************************************************************************************************************
    
    If loadSuccessful And (targetDIB.GetDIBWidth > 0) And (targetDIB.GetDIBHeight > 0) And (Not targetImage Is Nothing) Then
        
        PDDebug.LogAction "Debug note: image load appeared to be successful.  Summary forthcoming."
        
        '*************************************************************************************************************************************
        ' If the loaded image was in PDI format (PhotoDemon's internal format), skip a number of additional processing steps.
        '*************************************************************************************************************************************
        
        If (internalFormatID <> PDIF_PDI) Then
            
            'While inside this section of the load process, you'll notice a consistent trend regarding DOEVENTS.
            ' If you haven't already, this is a good time to scroll to the top of this function and read the IMPORTANT NOTE!
            
            '*************************************************************************************************************************************
            ' If the incoming image is 24bpp, convert it to 32bpp.  (PD assumes an available alpha channel for all layers.)
            '*************************************************************************************************************************************
            
            If ImageImporter.ForceTo32bppMode(targetDIB) Then VBHacks.DoEventsTimersOnly
            
            '*************************************************************************************************************************************
            ' The target DIB has been loaded successfully, so copy its contents into the main layer of the targetImage
            '*************************************************************************************************************************************
            
            'If the source file was already designed as a multi-layer format (e.g. PSD, OpenRaster, etc),
            ' this step is unnecessary.
            Dim layersAlreadyLoaded As Boolean: layersAlreadyLoaded = False
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_CBZ)
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_HEIF)
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_ICO)
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_JXL)
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_MBM)
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_ORA)
            layersAlreadyLoaded = layersAlreadyLoaded Or ((targetImage.GetCurrentFileFormat = PDIF_PCX) And (decoderUsed = id_PCXParser))
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_PDF)
            layersAlreadyLoaded = layersAlreadyLoaded Or ((targetImage.GetCurrentFileFormat = PDIF_PSD) And (decoderUsed = id_PSDParser))
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_PSP)
            layersAlreadyLoaded = layersAlreadyLoaded Or ((targetImage.GetCurrentFileFormat = PDIF_WEBP) And (decoderUsed = id_libwebp))
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_XCF)
            
            If (Not layersAlreadyLoaded) Then
                
                'Besides a source DIB, the "add new layer" function also wants a name for the new layer.  Create one now.
                Dim newLayerName As String
                newLayerName = Layers.GenerateInitialLayerName(srcFile, suggestedFilename, imageHasMultiplePages, targetImage, targetDIB)
                
                'Create the new layer in the target image, and pass our created name to it
                Dim newLayerID As Long
                newLayerID = targetImage.CreateBlankLayer
                targetImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, newLayerName, targetDIB, imageHasMultiplePages
                targetImage.UpdateSize
                
            End If
            
            If (ExifTool.IsMetadataPipeActive) Then VBHacks.DoEventsTimersOnly
            
        '/End specialized handling for non-PDI files
        End If
        
        'Any remaining attributes of interest should be stored in the target image now
        targetImage.ImgStorage.AddEntry "OriginalFileSize", Files.FileLenW(srcFile)
        
        'We've now completed the bulk of the image load process.  In nightly builds, dump a bunch of image-related data out to file;
        ' such data is invaluable when tracking down bugs.
        PDDebug.LogAction "~ Summary of image """ & Files.FileGetName(srcFile) & """ follows ~", , True
        PDDebug.LogAction vbTab & "Image ID: " & targetImage.imageID, , True
        
        PDDebug.LogAction vbTab & "Load engine: " & GetDecoderName(decoderUsed), , True
        PDDebug.LogAction vbTab & "Detected format: " & ImageFormats.GetInputFormatDescription(ImageFormats.GetIndexOfInputPDIF(targetImage.GetOriginalFileFormat)), , True
        PDDebug.LogAction vbTab & "Image dimensions: " & targetImage.Width & "x" & targetImage.Height, , True
        PDDebug.LogAction vbTab & "Image size (original file): " & Format$(targetImage.ImgStorage.GetEntry_Long("OriginalFileSize"), "#,#") & " Bytes", , True
        PDDebug.LogAction vbTab & "Image size (as loaded, approximate): " & Format$(targetImage.EstimateRAMUsage, "#,#") & " Bytes", , True
        PDDebug.LogAction vbTab & "Original color depth: " & targetImage.GetOriginalColorDepth, , True
        PDDebug.LogAction vbTab & "ICC profile embedded: " & (LenB(targetImage.GetColorProfile_Original) <> 0), , True
        PDDebug.LogAction vbTab & "Multiple pages embedded: " & CStr(imageHasMultiplePages) & " (" & numOfPages & ")", , True
        PDDebug.LogAction vbTab & "Number of layers: " & targetImage.GetNumOfLayers, , True
        PDDebug.LogAction vbTab & "Time to load: " & VBHacks.GetTimeDiffNowAsString(justImageLoadTime), , True
        PDDebug.LogAction "~ End of image summary ~", , True
        
        '*************************************************************************************************************************************
        ' Generate all relevant pdImage attributes tied to the source file (like the image's name and save state)
        '*************************************************************************************************************************************
        
        'First, see if this image is being restored from PD's "autosave" engine.  Autosaved images require special handling, because their
        ' state must be reconstructed from whatever bits we can dredge up from the temp file.
        If (srcFileExtension = "PDTMP") Then
            ImageImporter.SyncRecoveredAutosaveImage srcFile, targetImage
        Else
            ImageImporter.GenerateExtraPDImageAttributes srcFile, targetImage, suggestedFilename
        End If
        
        'Because ExifTool is sending us data in the background, we periodically yield for metadata piping.
        If (ExifTool.IsMetadataPipeActive) Then VBHacks.DoEventsTimersOnly
            
            
        '*************************************************************************************************************************************
        ' If this is a primary image, update all relevant UI elements (image size display, custom form icon, etc)
        '*************************************************************************************************************************************
        
        PDDebug.LogAction "Finalizing image details..."
        
        'The finalized pdImage object is finally worthy of being added to the central PD collection.
        ' (Note that this function will automatically update PDImages.GetActiveImageID() to point
        ' at the new image.)
        PDImages.AddImageToCentralCollection targetImage
        
        'The UI needs a *lot* of changes to reflect the state of the newly loaded image
        ImageImporter.ApplyPostLoadUIChanges srcFile, targetImage, addToRecentFiles
        
        'Because ExifTool is sending us data in the background, we periodically yield for metadata piping.
        If (ExifTool.IsMetadataPipeActive) Then VBHacks.DoEventsTimersOnly
        
            
        '*************************************************************************************************************************************
        ' If the just-loaded image was in a multipage format (e.g. multipage TIFF),
        '  perform a few extra checks.
        '*************************************************************************************************************************************
        
        'Before continuing on to the next image (if any), see if the just-loaded image contains multiple
        ' pages within the file. If it does, load each page into its own layer.
        '
        'NOTE: some multipage formats (like PSD, ORA, ICO, etc) load all pages/frames in the initial
        ' load function.  This "separate multipage loader function" approach primarily exists for
        ' legacy functions where a 3rd-party library is responsible for parsing the extra pages.
        If imageHasMultiplePages And ((targetImage.GetOriginalFileFormat = PDIF_TIFF) Or (targetImage.GetOriginalFileFormat = PDIF_GIF) Or (targetImage.GetOriginalFileFormat = PDIF_PNG) Or (targetImage.GetOriginalFileFormat = PDIF_AVIF) Or (targetImage.GetOriginalFileFormat = PDIF_DDS)) Then
            
            'Add a flag to this pdImage object noting that the multipage loading path *was* utilized.
            targetImage.ImgStorage.AddEntry "MultipageImportActive", True
            
            'The actual load process now varies by import engine.  PD can use both FreeImage and GDI+
            ' to import certain types of multipage images (e.g. TIFF).
            If (decoderUsed = id_FreeImage) Then
            
                'We now have several options for loading the remaining pages in this file.
                
                'For most images, the easiest path would be to keep calling the standard FI_LoadImage function(), passing it updated
                ' page numbers as we go.  This ensures that all the usual fallbacks and detailed edge-case handling (like ICC profiles
                ' that vary by page) are handled correctly.
                '
                'However, it also means that the source file is loaded/unloaded on each frame, because the FreeImage load function
                ' was never meant to be used like this.  This isn't a problem for images with a few pages, but if the image is large
                ' and/or if it has tons of frames (like a length animated GIF), we could be here awhile.
                '
                'As of 7.0, a better solution exists: ask FreeImage to cache the source file, and keep it cached until all frames
                ' have been loaded.  This is *way* faster, and it also lets us bypass a bunch of per-file validation checks
                ' (since we already know the source file is okay).
                loadSuccessful = Plugin_FreeImage.FinishLoadingMultipageImage(srcFile, targetDIB, numOfPages, , targetImage, , suggestedFilename)
            
            'GDI+ path
            ElseIf (decoderUsed = id_GDIPlus) Then
            
                'If we implement a load-time dialog in the future, and the user (for whatever reason)
                ' doesn't want all pages loaded, call this function to free cached multipage handles.
                ' (Note that you'll need to uncomment the corresponding function in the GDI_Plus
                ' module, too.)
                'GDI_Plus.MultiPageDataNotWanted
                
                'Otherwise, assume they want all pages loaded
                loadSuccessful = GDI_Plus.ContinueLoadingMultipageImage(srcFile, targetDIB, numOfPages, , targetImage, , suggestedFilename)
            
            'Internal multipage loader; this is used for animated PNG files
            Else
                If (targetImage.GetOriginalFileFormat = PDIF_PNG) Or (targetImage.GetOriginalFileFormat = PDIF_AVIF) Or (targetImage.GetOriginalFileFormat = PDIF_DDS) Then loadSuccessful = ImageImporter.LoadRemainingPNGFrames(targetImage)
            End If
            
            'As a convenience, make all but the first page/frame/icon invisible when the source is a GIF or PNG.
            ' (TIFFs don't typically require this, as all pages tend to be the same size.  Note that an exception
            '  to this is PSDs exported as multipage TIFFs via Photoshop - but in that case, we *still* want to
            '  make all pages visible by default)
            If (targetImage.GetNumOfLayers > 1) And (targetImage.GetOriginalFileFormat <> PDIF_TIFF) Then
                
                Dim pageTracker As Long
                For pageTracker = 1 To targetImage.GetNumOfLayers - 1
                    targetImage.GetLayerByIndex(pageTracker).SetLayerVisibility False
                Next pageTracker
                
                targetImage.SetActiveLayerByIndex 0
                
                'Also tag the image as being animated; we use this to activate some contextual UI bits
                targetImage.SetAnimated True
                
            End If
            
            'With all pages/frames/icons successfully loaded, redraw the main viewport
            Viewport.Stage1_InitializeBuffer targetImage, FormMain.MainCanvas(0), VSR_ResetToZero
            
        'Add a flag to this pdImage object noting that the multipage loading path was *not* utilized.
        Else
            targetImage.ImgStorage.AddEntry "MultipageImportActive", False
        End If
            
        '*************************************************************************************************************************************
        ' Hopefully metadata processing has finished, but if it hasn't, start a timer on the main form, which will wait for it to complete.
        '*************************************************************************************************************************************
        
        'Ask the metadata handler if it has finished parsing the image
        If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) And (decoderUsed <> id_PDIParser) Then
            
            'Some tools may have already stopped to load metadata
            If (Not targetImage.ImgMetadata.HasMetadata) Then
            
                If ExifTool.IsMetadataFinished Then
                
                    PDDebug.LogAction "Metadata retrieved successfully."
                    targetImage.ImgMetadata.LoadAllMetadata ExifTool.RetrieveMetadataString, targetImage.imageID
                    
                    'Because metadata already finished processing, retrieve any specific metadata-related entries
                    ' that may be useful to initial display of the image, like image resolution
                    Dim xResolution As Double, yResolution As Double
                    If targetImage.ImgMetadata.GetResolution(xResolution, yResolution) Then targetImage.SetDPI xResolution, yResolution
                    
                Else
                    PDDebug.LogAction "Metadata parsing hasn't finished; switching to asynchronous wait mode..."
                    FormMain.StartMetadataTimer
                End If
            
            End If
        
        End If
            
            
        '*************************************************************************************************************************************
        ' As of 2014, the new Undo/Redo engine requires a base pdImage copy as the starting point for Undo/Redo diffs.
        '*************************************************************************************************************************************
        
        'If this is a primary image, force an immediate Undo/Redo write to file.  This serves multiple purposes: it is our
        ' baseline for calculating future Undo/Redo diffs, and it can be used to recover the original file if something
        ' goes wrong before the user performs a manual save (e.g. AutoSave).
        '
        '(Note that all Undo behavior is disabled during batch processing, to improve performance, so we can skip this step.)
        If (Macros.GetMacroStatus <> MacroBATCH) Then
            
            Dim autoSaveTime As Currency
            VBHacks.GetHighResTime autoSaveTime
            PDDebug.LogAction "Creating initial auto-save entry (this may take a moment)..."
            
            Dim tmpProcCall As PD_ProcessCall
            With tmpProcCall
                .pcID = g_Language.TranslateMessage("Original image")
                .pcParameters = vbNullString
                .pcUndoType = UNDO_Everything
                .pcRaiseDialog = False
                .pcRecorded = True
            End With
            
            targetImage.UndoManager.CreateUndoData tmpProcCall
            PDDebug.LogAction "Initial auto-save creation took " & VBHacks.GetTimeDiffNowAsString(autoSaveTime)
            
        End If
            
            
        '*************************************************************************************************************************************
        ' Image loaded successfully.  Carry on.
        '*************************************************************************************************************************************
        
        loadSuccessful = True
        
        'In debug mode, note the new memory baseline, post-load
        PDDebug.LogAction "New memory report after loading image """ & Files.FileGetName(srcFile) & """:"
        PDDebug.LogAction vbNullString, PDM_Mem_Report
            
        'Also report an estimated memory delta, based on the pdImage object's self-reported memory usage.
        ' This provides a nice baseline for making sure PD's memory usage isn't out of whack for a given image.
        PDDebug.LogAction "(FYI, expected delta was approximately " & Format$(targetImage.EstimateRAMUsage() / 1000#, "#,#") & " K)"
        
    'This ELSE block is hit when the image fails post-load verification checks.  Treat the load as unsuccessful.
    Else
    
        loadSuccessful = False
        
        'Deactivate the (now useless) pdImage and pdDIB objects, which will forcibly unload whatever resources they may have claimed
        If (Not targetDIB Is Nothing) Then Set targetDIB = Nothing
        
        If (Not targetImage Is Nothing) Then
            targetImage.FreeAllImageResources
            Set targetImage = Nothing
        End If
        
    End If
    
    '*************************************************************************************************************************************
    ' As all images have now loaded, re-enable the main form
    '*************************************************************************************************************************************
    
    'Purge any input events that may have occurred during the load process
    VBHacks.PurgeInputMessages FormMain.hWnd
    
    'Synchronize all interface elements to match the newly loaded image(s)
    If handleUIDisabling Then Interface.SyncInterfaceToCurrentImage
    
    'Synchronize any non-destructive settings to the currently active layer
    If (handleUIDisabling And loadSuccessful) Then
        Processor.SyncAllGenericLayerProperties PDImages.GetActiveImage.GetActiveLayer
        Processor.SyncAllTextLayerProperties PDImages.GetActiveImage.GetActiveLayer
    End If
    
    '*************************************************************************************************************************************
    ' Before finishing, display any relevant load problems (missing files, invalid formats, etc)
    '*************************************************************************************************************************************
    
    'Restore the screen cursor if necessary, then set focus to the canvas
    If handleUIDisabling Then Processor.MarkProgramBusyState False, True, (PDImages.GetNumOpenImages > 1)
    If (Macros.GetMacroStatus <> MacroBATCH) Then FormMain.MainCanvas(0).SetFocusToCanvasView
    
    'Report success/failure back to the user
    LoadFileAsNewImage = loadSuccessful And (Not targetImage Is Nothing)
    
    'NEW IN 2025: look for mismatches between file extension and file type in the source file.
    ' If this happens, warn the user and offer to rename the underlying file with a correct extension.
    ' (Like anything else that raises a modal dialog, this check is disabled during batch processes.)
    If LoadFileAsNewImage And (Macros.GetMacroStatus <> MacroBATCH) And arePopupsAllowed Then
        
        'Ignore images that didn't originate from disk
        If (LenB(suggestedFilename) = 0) Then
        
        'Ignore image originating from temp files (common when e.g. loading from a .zip file)
        If (targetImage.GetOriginalFileFormat <> PDIF_UNKNOWN) And (targetImage.GetOriginalFileFormat <> PDIF_RAWBUFFER) _
            And (targetImage.GetOriginalFileFormat <> PDIF_TMPFILE) Then
    
            'The file appears to be a normal on-disk image file.
            
            'See what file format is expected for this particular extension.
            Dim expectedFormatForExtension As PD_IMAGE_FORMAT
            expectedFormatForExtension = ImageFormats.IsExtensionOkayForAnyPDIF(Files.FileGetExtension(srcFile))
            
            'Compare the expected extension to the one we got.
            Dim warnUser As Boolean
            warnUser = (expectedFormatForExtension <> targetImage.GetOriginalFileFormat)
            
            'Hmmm, the file's contents don't match its format!  We don't want to mess with images
            ' in a custom format, but if the image file has an extension that *is* associated with another format
            ' (like e.g. a JPEG image with a PNG extension), we *do* want to warn the user.
            If warnUser And (expectedFormatForExtension <> PDIF_UNKNOWN) Then
                
                PDDebug.LogAction "WARNING: bad file extension found.  If the correctly named version doesn't exist, we'll offer to rename..."
                
                Dim correctExtension As String
                correctExtension = ImageFormats.GetExtensionFromPDIF(expectedFormatForExtension)
                
                Dim renamedFilename As String
                renamedFilename = Files.FileGetPath(srcFile) & Files.FileGetName(srcFile, True) & "." & ImageFormats.GetExtensionFromPDIF(targetImage.GetOriginalFileFormat)
                
                If (Not Files.FileExists(renamedFilename)) Then
                    
                    Dim msgBadExtension As pdString
                    Set msgBadExtension = New pdString
                    msgBadExtension.AppendLine g_Language.TranslateMessage("This file has the extension ""%1"", but it is actually in ""%2"" format.", Files.FileGetExtension(srcFile), ImageFormats.GetExtensionFromPDIF(targetImage.GetOriginalFileFormat))
                    msgBadExtension.AppendLineBreak
                    msgBadExtension.AppendLine g_Language.TranslateMessage("May PhotoDemon rename the file with a correct extension?")
                    msgBadExtension.AppendLineBreak
                    msgBadExtension.AppendLine g_Language.TranslateMessage("(If you choose ""Yes"", the file will be renamed to ""%1"")", renamedFilename)
                    
                    Dim renameResult As VbMsgBoxResult
                    renameResult = PDMsgBox(msgBadExtension.ToString(), vbExclamation Or vbYesNoCancel Or vbApplicationModal, "Bad file extension")
                    userCanceledImportDialog = (renameResult = vbCancel)
                    
                    If (renameResult = vbYes) Then
                        
                        'The user wants us to rename the file.  This requires us to "rewind" some choices made
                        ' earlier in the load process, and update various tracking bits (like MRU menus) accordingly.
                        
                        'Before doing anything, don't proceed until metadata has finished loading.
                        If (Not targetImage.ImgMetadata.HasMetadata) Then
                            
                            'Let metadata process in a separate thread, because if we try to rename the file while Exiftool is
                            ' touching it, either (or both) operations will fail.
                            '
                            '(We don't want this to go on forever, though, so after a few seconds, attempt anyway -
                            ' something may have gone wrong on ExifTool's end.)
                            Dim totalTimeElapsed As Long
                            Do While (Not ExifTool.IsMetadataFinished) And (totalTimeElapsed < 2000)
                                VBHacks.SleepAPI 200
                                totalTimeElapsed = totalTimeElapsed + 200
                            Loop
                            
                        End If
                        
                        'Attempt the rename.
                        If Files.FileMove(srcFile, renamedFilename, False) Then
                            
                            'We replaced the file without trouble.  Now we need to update a bunch of internal stuff
                            ' that's no longer relevant (like MRUs).
                            targetImage.ImgStorage.AddEntry "CurrentLocationOnDisk", renamedFilename
                            targetImage.ImgStorage.AddEntry "OriginalFileName", Files.FileGetName(renamedFilename, True)
                            targetImage.ImgStorage.AddEntry "OriginalFileExtension", Files.FileGetExtension(renamedFilename)
                            If addToRecentFiles Then g_RecentFiles.ReplaceExistingEntry srcFile, renamedFilename
                            If handleUIDisabling Then Interface.SyncInterfaceToCurrentImage
                            
                        Else
                            PDDebug.LogAction "WARNING: file copy failed."
                            If handleUIDisabling Then Processor.MarkProgramBusyState False, True
                            Message "Warning - file locked: %1", srcFile
                            PDMsgBox "Unfortunately, the file '%1' is currently locked by another program on this PC." & vbCrLf & vbCrLf & "Please close this file in any other running programs, then try again.", vbExclamation Or vbOKOnly, "File locked", srcFile
                        End If
                        
                    '/end user answered "yes" to rename
                    End If
                
                '/target file (with correct extension) already exists in that folder!
                Else
                    PDDebug.LogAction "WARNING: file with correct extension already exists!"
                End If
            
            '/end file format is OK for this extension
            End If
            
        End If  '/end image was temp file
        End If  '/end image didn't originate on disk
    End If  '/end in batch process, or file didn't load correctly anyway
    
    'If any of the import dialogs were outright canceled, relay this to the caller via the ByRef suspendWarnings param
    If userCanceledImportDialog Then suspendWarnings = vbCancel
    
    'Activate the new image (if loading was successful) and exit
    If LoadFileAsNewImage Then
        If handleUIDisabling Then CanvasManager.ActivatePDImage PDImages.GetActiveImageID(), "LoadFileAsNewImage", newImageJustLoaded:=True
        Message "Image loaded successfully."
    Else
        If userCanceledImportDialog Then numCanceledImports = numCanceledImports + 1
        If (Macros.GetMacroStatus <> MacroBATCH) And arePopupsAllowed And (freeImage_Return <> PD_FAILURE_USER_CANCELED) Then
            If userCanceledImportDialog Then
                Message "Action canceled."
            Else
                Message "Failed to load %1", srcFile
                If (Not m_MultiImageLoadActive) Then
                    Dim tmpFileList As pdStringStack: Set tmpFileList = New pdStringStack
                    tmpFileList.AddString Files.FileGetName(srcFile)
                    ShowFailedLoadMsgBox tmpFileList
                End If
            End If
        End If
    End If
    
    PDDebug.LogAction "Image loaded in " & Format$(VBHacks.GetTimerDifferenceNow(startTime) * 1000, "#0") & " ms"
        
End Function

'Quick and dirty function for loading an image file to a pdDIB object (*NOT* a pdImage object).
'
'Per the name, this function provides an absolutely barebones approach to getting image data into a useable RGBA surface.
' It explicitly assumes that the calling function handles any work above and beyond this.
' (Note that things like metadata are not processed *at all.*)
'
'Similarly, format sorting is *primarily* handled by file extension.  This function doesn't interrogate
' format details as aggressively as PD's central LoadFileAsNewImage(), so mislabeled file extensions
' may result in a file not being loaded.  This is by design.
'
'That said, format ID and decoder sorting is still applied intelligently, so this function should reflect
' PD's full capacity for image format support.  Importantly, however, multi-page files get squashed into
' single-frame composites, by design.
'
'The function will return TRUE if successful; detailed load information is not available past that.
Public Function QuickLoadImageToDIB(ByVal imagePath As String, ByRef targetDIB As pdDIB, Optional ByVal applyUIChanges As Boolean = True, Optional ByVal displayMessagesToUser As Boolean = True, Optional ByVal suppressDebugData As Boolean = False) As Boolean
    
    Dim loadSuccessful As Boolean: loadSuccessful = False
    
    'Even though this function is designed to operate as quickly as possible, some images may take a long time to load.
    If applyUIChanges Then
        Processor.MarkProgramBusyState True, True
    End If
    
    'Before attempting to load an image, make sure it exists
    If (Not Files.FileExists(imagePath)) Then
        PDDebug.LogAction "QuickLoadImageToDIB error - file not found:" & imagePath
        If displayMessagesToUser Then PDMsgBox "Unfortunately, the image '%1' could not be found." & vbCrLf & vbCrLf & "If this image was originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbExclamation Or vbOKOnly, "File not found", imagePath
        QuickLoadImageToDIB = False
        If applyUIChanges Then Processor.MarkProgramBusyState False, True
        Exit Function
    End If
    
    'Prepare a dummy pdImage object, which some external functions may require
    Dim tmpPDImage As pdImage
    Set tmpPDImage = New pdImage
    
    'Start by stripping the extension from the file path
    Dim fileExtension As String
    fileExtension = UCase$(Files.FileGetExtension(imagePath))
    loadSuccessful = False
    
    'For weird and/or esoteric formats, we'll throw 'em at FreeImage and see what happens
    Dim freeImageReturn As PD_OPERATION_OUTCOME
    freeImageReturn = PD_FAILURE_GENERIC
    
    'Depending on the file's extension, load the image using the most appropriate image decoding routine
    Select Case fileExtension
    
        'PhotoDemon's custom file format must be handled specially (as obviously, FreeImage and GDI+ won't handle it!)
        Case "PDI"
        
            'PDI images require zstd, and are only loaded via a custom routine (obviously, since they are PhotoDemon's native format)
            loadSuccessful = LoadPDI_Normal(imagePath, targetDIB, tmpPDImage)
            
            'Retrieve a copy of the fully composited image
            tmpPDImage.GetCompositedImage targetDIB
            
        'TMPDIB files are raw pdDIB objects dumped directly to file.  In some cases, this is faster and
        ' easier for PD than wrapping the pdDIB object inside a pdPackage layer (especially if this function
        ' is going to be used, since we're just going to decode the saved file into a pdDIB anyway).
        Case "TMPDIB", "PDTMPDIB"
            loadSuccessful = LoadRawImageBuffer(imagePath, targetDIB, tmpPDImage)
            
        'TMP files are internal PD temp files generated from a wide variety of use-cases (Clipboard is one example).  These are
        ' typically in BMP format, but this is not contractual.  A standard cascade of load functions is used.
        Case "TMP"
            If ImageFormats.IsFreeImageEnabled() Then loadSuccessful = (FI_LoadImage_V5(imagePath, targetDIB, 0, False, Nothing, suppressDebugData) = PD_SUCCESS)
            If (Not loadSuccessful) Then loadSuccessful = LoadGDIPlusImage(imagePath, targetDIB, tmpPDImage)
            If (Not loadSuccessful) Then loadSuccessful = LoadRawImageBuffer(imagePath, targetDIB, tmpPDImage)
            
        'PDTMP files are custom PD-format files saved ONLY during Undo/Redo or Autosaving.
        ' As such, they have weirdly specific parsing criteria inside PD's central load function,
        ' but for quick-loading, we can simply grab the raw image buffer inside 'em.
        Case "PDTMP"
            loadSuccessful = LoadRawImageBuffer(imagePath, targetDIB, tmpPDImage)
        
        'Internal decoders follow
        Case "CBZ"
            Dim cCBZ As pdCBZ
            Set cCBZ = New pdCBZ
            If cCBZ.IsFileCBZ(imagePath) Then loadSuccessful = cCBZ.LoadCBZ(imagePath, tmpPDImage)
            If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
        
        Case "DDS"
            If Plugin_DDS.IsDirectXTexAvailable() Then
                loadSuccessful = ImageImporter.LoadDDS(imagePath, tmpPDImage, targetDIB, False, 1, False)
            Else
                freeImageReturn = FI_LoadImage_V5(imagePath, targetDIB, 0, False, Nothing, suppressDebugData)
                loadSuccessful = (freeImageReturn = PD_SUCCESS)
            End If
        
        Case "HGT"
            Dim cHGT As pdHGT
            Set cHGT = New pdHGT
            If cHGT.IsFileHGT(imagePath) Then loadSuccessful = cHGT.LoadHGT_FromFile(imagePath, tmpPDImage, targetDIB)
        
        'Icons are weird because we need to grab a specific frame, but which one?  Just grab the first for now;
        ' if the user wants a specific one, they'd need to load the file individually and manually grab a layer.
        Case "ICO"
            Dim cIconReader As pdICO
            Set cIconReader = New pdICO
            If cIconReader.IsFileICO(imagePath, True) Then
                loadSuccessful = (cIconReader.LoadICO(imagePath, tmpPDImage, targetDIB) < ico_Failure)
                If loadSuccessful And (Not tmpPDImage Is Nothing) Then
                    Set targetDIB = New pdDIB
                    targetDIB.CreateFromExistingDIB tmpPDImage.GetActiveLayer.GetLayerDIB
                End If
            End If
        
        Case "JLS"
            loadSuccessful = Plugin_CharLS.LoadJLS(imagePath, tmpPDImage, targetDIB)
        
        Case "JP2", "J2K", "JPT", "J2C", "JPC", "JPX", "JPF", "JPH"
            loadSuccessful = Plugin_OpenJPEG.LoadJP2(imagePath, tmpPDImage, targetDIB)
        
        Case "JXL"
            If Plugin_jxl.IsFileJXL_NoExternalLibrary(imagePath) Then
                If Plugin_jxl.IsJXLImportAvailable() Then
                    loadSuccessful = Plugin_jxl.LoadJXL(imagePath, tmpPDImage, targetDIB)
                    If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
                End If
            End If
        
        Case "MBM", "MBW", "MCL", "AIF", "ABW", "ACL"
            Dim cMBM As pdMBM
            Set cMBM = New pdMBM
            If cMBM.IsFileMBM(imagePath) Then loadSuccessful = cMBM.LoadMBM_FromFile(imagePath, tmpPDImage, targetDIB)
            If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
        
        Case "ORA"
            Dim cORA As pdOpenRaster
            Set cORA = New pdOpenRaster
            If cORA.IsFileORA(imagePath) Then loadSuccessful = cORA.LoadORA(imagePath, tmpPDImage)
            If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
        
        Case "PCC", "PCX", "DCX"
            Dim cPCX As pdPCX
            Set cPCX = New pdPCX
            If cPCX.IsFilePCX(imagePath, False, True) Then loadSuccessful = cPCX.LoadPCX_FromFile(imagePath, tmpPDImage, targetDIB)
            Set cPCX = Nothing
            If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
        
        Case "PDF"
            If Plugin_PDF.IsFileLikelyPDF(imagePath) Then
                loadSuccessful = ImageImporter.LoadPDF(imagePath, tmpPDImage, targetDIB, True)
                If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
            End If
            
        Case "PNG", "APNG"
            Dim cPNG As pdPNG
            Set cPNG = New pdPNG
            loadSuccessful = (cPNG.LoadPNG_Simple(imagePath, tmpPDImage, targetDIB) < png_Failure)
            If (Not targetDIB.GetAlphaPremultiplication) Then targetDIB.SetAlphaPremultiplication True
            
        Case "PSD", "PSB"
            Dim cPSD As pdPSD
            Set cPSD = New pdPSD
            If cPSD.IsFilePSD(imagePath) Then loadSuccessful = (cPSD.LoadPSD(imagePath, tmpPDImage, targetDIB) < psd_Failure)
            If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
        
        Case "PSP", "PSPIMAGE", "TUB", "PSPTUBE", "PFR", "PSPFRAME", "MSK", "PSPMASK", "PSPBRUSH"
            Dim cPSP As pdPSP
            Set cPSP = New pdPSP
            If cPSP.IsFilePSP(imagePath) Then loadSuccessful = (cPSP.LoadPSP(imagePath, tmpPDImage, targetDIB) < psp_Failure)
            If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
        
        Case "QOI"
            Dim cQOI As pdQOI
            Set cQOI = New pdQOI
            If cQOI.IsFileQOI(imagePath, False, True) Then loadSuccessful = cQOI.LoadQOI_FromFile(imagePath, tmpPDImage, targetDIB)
            If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
        
        Case "SVG", "SVGZ"
            If Plugin_resvg.IsResvgEnabled() Then
                If Plugin_resvg.IsFileSVGCandidate(imagePath) Then loadSuccessful = Plugin_resvg.LoadSVG_FromFile(imagePath, tmpPDImage, targetDIB, True)
            End If
        
        Case "WBMP", "WBM", "WAP"
            Dim cWbmp As pdWBMP
            Set cWbmp = New pdWBMP
            If cWbmp.IsFileWBMP(imagePath) Then loadSuccessful = cWbmp.LoadWBMP_FromFile(imagePath, tmpPDImage, targetDIB)
        
        Case "WEBP"
            If Plugin_WebP.IsWebPEnabled() Then
                If Plugin_WebP.IsWebP(imagePath) Then
                    Dim cWebP As pdWebP
                    Set cWebP = New pdWebP
                    loadSuccessful = cWebP.QuickLoadWebP_FromFile(imagePath, tmpPDImage, targetDIB)
                    If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
                End If
            End If
        
        Case "XBM"
            Dim cXbm As pdXBM
            Set cXbm = New pdXBM
            If cXbm.IsFileXBM(imagePath) Then loadSuccessful = cXbm.LoadXBM_FromFile(imagePath, tmpPDImage, targetDIB)
            
        Case "XCF"
            Dim cXCF As pdXCF
            Set cXCF = New pdXCF
            If cXCF.IsFileXCF(imagePath) Then loadSuccessful = cXCF.LoadXCF_FromFile(imagePath, tmpPDImage, targetDIB)
            If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
        
        'HEIF support was added in v2024.8
        Case "HEIF", "HEIFS", "HEIC", "HEICS", "HIF"
            If Plugin_Heif.IsFileHeif(imagePath) Then loadSuccessful = Plugin_Heif.LoadHeifImage(imagePath, tmpPDImage, targetDIB, True, True)
            If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
            
        'AVIF support was added in v9.0
        Case "AVCI", "AVCS", "AVIF", "AVIFS"
            loadSuccessful = Plugin_AVIF.QuickLoadPotentialAVIFToDIB(imagePath, targetDIB, tmpPDImage)
            
        'All other formats follow a set pattern: try to load them via FreeImage (if it's available), then GDI+, then finally
        ' VB's internal LoadPicture function.
        Case Else
            
            'FreeImage's TIFF support (via libTIFF?) is wonky.  It's prone to bad crashes and inexplicable
            ' memory issues (including allocation failures on normal-sized images), so for TIFFs we want to
            ' try GDI+ before trying FreeImage.  (PD's GDI+ image loader was heavily restructured in v8.0 to
            ' support things like multi-page import, so this strategy wasn't viable until then.)
            Dim tryGDIPlusFirst As Boolean
            tryGDIPlusFirst = (fileExtension = "TIF") Or (fileExtension = "TIFF")
            
            'On modern Windows builds (8+) FreeImage is markedly slower than GDI+ at loading JPEG images,
            ' so let's also default to GDI+ for JPEGs.
            tryGDIPlusFirst = tryGDIPlusFirst Or (fileExtension = "JPG") Or (fileExtension = "JPEG") Or (fileExtension = "JPE")
            
            'Animated GIFs are supported by both engines, but GDI+ is faster
            tryGDIPlusFirst = tryGDIPlusFirst Or (fileExtension = "GIF")
            
            'If GDI+ is preferable, attempt it now
            If tryGDIPlusFirst Then loadSuccessful = LoadGDIPlusImage(imagePath, targetDIB, tmpPDImage, nonInteractiveMode:=True)
            
            'If GDI+ failed, proceed with FreeImage
            If (Not loadSuccessful) And ImageFormats.IsFreeImageEnabled() Then
                freeImageReturn = FI_LoadImage_V5(imagePath, targetDIB, 0, False, Nothing, suppressDebugData)
                loadSuccessful = (freeImageReturn = PD_SUCCESS)
                
                'If FreeImage failed and we haven't tried GDI+ yet, try it now
                If (Not loadSuccessful) And (Not tryGDIPlusFirst) Then loadSuccessful = LoadGDIPlusImage(imagePath, targetDIB, tmpPDImage, nonInteractiveMode:=True)
                
            End If
            
            'As a final resort, attempt WIC.  (Win 10+ may provide HEIF codecs or other user-installed ones.)
            If (Not loadSuccessful) And WIC.IsWICAvailable() Then loadSuccessful = LoadHEIF(imagePath, tmpPDImage, targetDIB)
            
    End Select
    
    'Sometimes, our image load functions will think the image loaded correctly, but they will return a blank image.  Check for
    ' non-zero width and height before continuing.
    If (targetDIB Is Nothing) Then
        loadSuccessful = False
    Else
        If (targetDIB.GetDIBWidth = 0) Or (targetDIB.GetDIBHeight = 0) Then loadSuccessful = False
    End If
    
    If (Not loadSuccessful) Then
        
        'Only display an error dialog if the import wasn't canceled by the user
        If displayMessagesToUser Then
            If (freeImageReturn <> PD_FAILURE_USER_CANCELED) Then
                Message "Failed to load %1", imagePath
                PDMsgBox "Unfortunately, PhotoDemon was unable to load the following image:" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save this image in a generic format (such as JPEG or PNG) before loading it.  Thanks!", vbExclamation Or vbOKOnly, "Image import failed", imagePath
            Else
                Message "Layer import canceled."
            End If
        End If
        
        'Deactivate the (now useless) DIB and parent object
        If (Not tmpPDImage Is Nothing) Then
            tmpPDImage.FreeAllImageResources
            Set tmpPDImage = Nothing
        End If
        
        If (Not targetDIB Is Nothing) Then Set targetDIB = Nothing
        
        'Re-enable the main interface
        If applyUIChanges Then Processor.MarkProgramBusyState False, True
        
        'Exit with failure status
        QuickLoadImageToDIB = False
        
        Exit Function
        
    End If
    
    'Restore the main interface
    If applyUIChanges Then Processor.MarkProgramBusyState False, True

    'If we made it all the way here, the image file was loaded successfully!
    QuickLoadImageToDIB = True

End Function

'Given a source filename's extension, return the estimated filetype (as an FIF_ constant) if the image format is specific to PD.
' This lets us quickly redirect PD-specific files to our own internal functions.
Private Function CheckForInternalFiles(ByRef srcFileExtension As String) As PD_IMAGE_FORMAT
    
    CheckForInternalFiles = PDIF_UNKNOWN
    
    Select Case srcFileExtension
    
        'Well-formatted PDI files
        Case "PDI", "PDTMP"
            CheckForInternalFiles = PDIF_PDI
            
        'TMPDIB files are raw pdDIB objects dumped directly to file.  In some cases, this is faster and easier for PD than wrapping
        ' the pdDIB object inside a pdPackage layer (e.g. during clipboard interactions, since we start with a raw pdDIB object
        ' after selections and such are applied to the base layer/image, so we may as well just use the raw pdDIB data we've cached).
        Case "TMPDIB", "PDTMPDIB"
            CheckForInternalFiles = PDIF_RAWBUFFER
            
        'Straight TMP files are internal files (BMP, typically) used by PhotoDemon.
        ' In some cases these come from 3rd-party libraries (e.g. EZTWAIN) so their format
        ' is not necessarily guaranteed in advance.
        Case "TMP"
            CheckForInternalFiles = PDIF_TMPFILE
            
    End Select
    
    'Any other formats will be dealt with by PD's standard cascade of load functions.

End Function

Private Function GetDecoderName(ByVal srcDecoder As PD_ImageDecoder) As String
    Select Case srcDecoder
        Case id_GDIPlus
            GetDecoderName = "GDI+"
        Case id_FreeImage
            GetDecoderName = "FreeImage"
        Case id_CBZParser
            GetDecoderName = "Internal CBZ parser"
        Case id_HDRParser
            GetDecoderName = "Internal HDR parser"
        Case id_HGTParser
            GetDecoderName = "Internal HGT parser"
        Case id_ICOParser
            GetDecoderName = "Internal ICO parser"
        Case id_PDIParser
            GetDecoderName = "Internal PDI parser"
        Case id_MBMParser
            GetDecoderName = "Internal MBM parser"
        Case id_ORAParser
            GetDecoderName = "Internal OpenRaster parser"
        Case id_PCXParser
            GetDecoderName = "Internal PCX parser"
        Case id_PNGParser
            GetDecoderName = "Internal PNG parser"
        Case id_PSDParser
            GetDecoderName = "Internal PSD parser"
        Case id_PSPParser
            GetDecoderName = "Internal PaintShop Pro parser"
        Case id_QOIParser
            GetDecoderName = "Internal QOI parser"
        Case id_WBMPParser
            GetDecoderName = "Internal WBMP parser"
        Case id_XBMParser
            GetDecoderName = "Internal XBM parser"
        Case id_XCFParser
            GetDecoderName = "Internal XCF parser"
        Case id_WIC
            GetDecoderName = "Windows Imaging Component"
        Case id_CharLS
            GetDecoderName = "CharLS"
        Case id_libavif
            GetDecoderName = "libavif"
        Case id_libwebp
            GetDecoderName = "libwebp"
        Case id_resvg
            GetDecoderName = "resvg"
        Case id_libjxl
            GetDecoderName = "libjxl"
        Case id_pdfium
            GetDecoderName = "pdfium"
        Case id_libheif
            GetDecoderName = "libheif"
        Case id_DirectXTex
            GetDecoderName = "DirectXTex"
        Case id_OpenJPEG
            GetDecoderName = "OpenJPEG"
        Case Else
            GetDecoderName = "unknown?!"
    End Select
End Function

'Want to load a whole bunch of image sources at once?  Use this function to do so.
' While helpful, note that it comes with some caveats:
' 1) The only supported sources are absolute filenames.
' 2) You lose the ability to assign custom titles to incoming images.  Titles will be auto-assigned based on their filenames.
' 3) You won't receive detailed success/failure information on each file.  Instead, this function will return TRUE if it loaded
'    at least one image successfully.
'
'If you want per-file success/fail results, call LoadFileAsNewImage manually from your own loop.
Public Function LoadMultipleImageFiles(ByRef srcList As pdStringStack, Optional ByVal updateRecentFileList As Boolean = True, Optional ByRef numCanceledImports As Long = 0) As Boolean

    If (Not srcList Is Nothing) Then
        
        m_MultiImageLoadActive = True
        
        'A lot can go wrong when loading image files.  This function will track failures and notify the user post-load.
        Dim numFailures As Long, numSuccesses As Long
        Dim brokenFiles As pdStringStack
        Set brokenFiles = New pdStringStack
        
        'The user may receive import dialogs for some formats.  This value will track the number of canceled import dialogs.
        ' If this value matches the number of failed imports, nothing actually went wrong during the import - the user just
        ' canceled 1+ loads.
        numCanceledImports = 0
        
        Processor.MarkProgramBusyState True, True
        
        Dim tmpFilename As String
        Do While srcList.PopString(tmpFilename)
            
            'The command-line may include other switches besides just filenames.  Ensure target file
            ' exists before forwarding it to the loader.
            If Files.FileExists(tmpFilename) Then
                
                Message "Importing %1...", Files.FileGetName(tmpFilename)
                
                Dim loadResult As VbMsgBoxResult
                loadResult = vbNo
                
                'Proceed with the load, and track successes/failures separately
                If Loading.LoadFileAsNewImage(tmpFilename, vbNullString, updateRecentFileList, loadResult, False, vbNullString, numCanceledImports) Then
                    numSuccesses = numSuccesses + 1
                Else
                    If (LenB(tmpFilename) <> 0) Then
                        numFailures = numFailures + 1
                        brokenFiles.AddString Files.FileGetName(tmpFilename, False)
                    End If
                End If
                
                If (loadResult = vbCancel) Then Exit Do
                
            End If
            
        Loop
        
        'Make sure we loaded at least one image from the original list
        If ((numSuccesses + numFailures) > 1) Or (numFailures > 0) Then
            Message "%1 of %2 images loaded successfully", numSuccesses, numSuccesses + numFailures
        Else
            Message vbNullString
        End If
        
        LoadMultipleImageFiles = (numSuccesses > 0)
        
        'Manually activate the last-loaded image
        Dim imgStack As pdStack
        If PDImages.GetListOfActiveImageIDs(imgStack) Then
            CanvasManager.ActivatePDImage imgStack.GetInt(imgStack.GetNumOfInts - 1), "LoadMultipleImageFiles", , , True
        End If
        
        'Synchronize everything to all open images
        SyncInterfaceToCurrentImage
        Processor.MarkProgramBusyState False, True, (PDImages.GetNumOpenImages() > 1)
        
        'Free the shared compression buffer (which may have been used to "suspend" previous images as we went)
        UIImages.FreeSharedCompressBuffer
        
        m_MultiImageLoadActive = False
        
        'Even if returning TRUE, we still want to notify the user of any failed files
        If (numFailures > 0) And (numCanceledImports = 0) Then
            ShowFailedLoadMsgBox brokenFiles
        End If
        
    Else
        LoadMultipleImageFiles = False
    End If

End Function

'Display a message box with explanation for one or more failed-to-load files
Private Sub ShowFailedLoadMsgBox(ByRef srcFilesBroken As pdStringStack)
    
    'Failsafe only
    If (srcFilesBroken Is Nothing) Then Exit Sub
    If (srcFilesBroken.GetNumOfStrings <= 0) Then Exit Sub
    
    'Assemble the list of broken files into a list of filenames and/or plugin error messages,
    ' to help the user understand what may have gone wrong.
    Dim listOfFiles As pdString
    Set listOfFiles = New pdString
    
    'Retrieve any third-party library errors from the plugin manager
    Dim tplNames As pdStringStack, tplMsgs As pdStringStack, tplFilenames As pdStringStack
    PluginManager.GetErrorPluginStacks tplNames, tplMsgs, tplFilenames
    
    Dim i As Long, idxMatch As Long
    For i = 0 To srcFilesBroken.GetNumOfStrings() - 1
        
        'Tahoma on XP doesn't have the same unicode range guarantees as Vista+
        If OS.IsWin7OrLater Then
            listOfFiles.Append ChrW$(&H2022)
        Else
            listOfFiles.Append "-"
        End If
        
        listOfFiles.Append Space$(2)
        listOfFiles.AppendLine srcFilesBroken.GetString(i)
            
        'If a load process supplied a reason for the error, append it now
        idxMatch = tplFilenames.ContainsString(srcFilesBroken.GetString(i), True)
        If (idxMatch >= 0) Then
            listOfFiles.Append Space$(4) & "("
            listOfFiles.Append g_Language.TranslateMessage("A third-party library (%1) reported the following error:", tplNames.GetString(idxMatch))
            listOfFiles.Append " "
            listOfFiles.Append Strings.ForceSingleLine(Trim$(tplMsgs.GetString(idxMatch)))
            listOfFiles.AppendLine ")"
        End If
        
    Next i
    
    PDMsgBox "Unfortunately, PhotoDemon was unable to load the following image(s):" & vbCrLf & vbCrLf & "%1" & vbCrLf & "Please verify that these image(s) exist, and that they use a supported image format (like JPEG or PNG).  Thanks!", vbExclamation Or vbOKOnly, "Some images were not loaded", listOfFiles.ToString()
    
End Sub

'Load an image via drag/drop on an individual control.
' Optionally, a target x/y can be passed (this is really only useful for dropping on an existing canvas;
' the values will be ignored otherwise).
Public Function LoadFromDragDrop(ByRef Data As DataObject, ByRef Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, Optional ByRef x As Single = 0!, Optional ByRef y As Single = 0!) As Boolean

    'Make sure the main window is available (e.g. a modal dialog hasn't stolen focus)
    If (Not g_AllowDragAndDrop) Then Exit Function
    
    'Use the external function (in the clipboard handler, as the code is roughly identical to clipboard pasting)
    ' to load the OLE source.
    Dim dropAsNewLayer As VbMsgBoxResult
    dropAsNewLayer = Dialogs.PromptDropAsNewLayer()
    
    If (dropAsNewLayer <> vbCancel) Then
        If (dropAsNewLayer = vbNo) And ((x <> 0!) Or (y <> 0!)) Then
            LoadFromDragDrop = g_Clipboard.LoadImageFromDragDrop(Data, Effect, (dropAsNewLayer = vbNo), Int(x + 0.5!), Int(y + 0.5!))
        Else
            LoadFromDragDrop = g_Clipboard.LoadImageFromDragDrop(Data, Effect, (dropAsNewLayer = vbNo))
        End If
    End If
    
End Function

'I don't know if it makes sense here, but because this module handles loading via drag+drop,
' I've also placed a helper function here for handling cursor update on a drag/over event.
Public Function HelperForDragOver(ByRef Data As DataObject, ByRef Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single, ByRef State As Integer) As Boolean
    
    'PD supports a lot of potential drop sources these days.
    ' These values are defined and addressed by the main clipboard handler, because Drag/Drop and clipboard actions
    ' share a lot of code.
    If g_Clipboard.IsObjectDragDroppable(Data) And g_AllowDragAndDrop Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If

End Function

'This routine sets the message on the splash screen (used only when the program is first started)
Public Sub LogStartupEvent(ByRef sMsg As String)
    
    Static loadProgress As Long
    
    'In debug mode, mirror message output to PD's central Debugger
    PDDebug.LogAction sMsg, PDM_User_Message
    
    'Load messages are translatable, but we don't want to translate them if the translation object isn't ready yet
    If (Not g_Language Is Nothing) Then
        If g_Language.ReadyToTranslate Then
            If g_Language.TranslationActive Then sMsg = g_Language.TranslateMessage(sMsg)
        End If
    End If
    
    'We no longer display the actual loading text to the user; instead, they just get a small progress bar update.
    If FormSplash.Visible Then FormSplash.UpdateLoadProgress loadProgress
    
    loadProgress = loadProgress + 1
    
End Sub

'Make a copy of the current image.  Thanks to PSC user "Achmad Junus" for this suggestion.
Public Sub DuplicateCurrentImage()
    
    Message "Duplicating current image..."
    
    'Ask the currently active image to write itself out to file
    Dim tmpDuplicationFile As String
    tmpDuplicationFile = UserPrefs.GetTempPath & "PDDuplicate.pdi"
    Saving.SavePDI_Image PDImages.GetActiveImage(), tmpDuplicationFile, True, cf_Lz4, cf_Lz4, False
    
    'We can now use the standard image load routine to import the temporary file
    Dim sTitle As String
    sTitle = PDImages.GetActiveImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    If (LenB(sTitle) = 0) Then sTitle = g_Language.TranslateMessage("[untitled image]")
    sTitle = sTitle & " - " & g_Language.TranslateMessage("Copy")
    
    Loading.LoadFileAsNewImage tmpDuplicationFile, sTitle, False
                    
    'Be polite and remove the temporary file
    Files.FileDeleteIfExists tmpDuplicationFile
    
    Message "Image duplication complete."
    
End Sub

'When testing, it can be helpful to load *all* entries from the recent files menu.
Public Sub LoadAllRecentFiles()
    
    If (g_RecentFiles.GetNumOfItems > 0) Then
        
        Dim listOfFiles As pdStringStack
        Set listOfFiles = New pdStringStack
        
        Dim i As Long
        For i = 0 To g_RecentFiles.GetNumOfItems() - 1
            listOfFiles.AddString g_RecentFiles.GetFullPath(i)
        Next i
        
        Loading.LoadMultipleImageFiles listOfFiles, True
        
    End If
        
End Sub
