Attribute VB_Name = "Loading"
'***************************************************************************
'General-purpose image and data import interface
'Copyright 2001-2021 by Tanner Helland
'Created: 4/15/01
'Last updated: 18/June/21
'Last update: expand the internal QuickLoadToDIB function to cover all the esoteric formats PD now supports via internal decoders
'
'This module provides high-level "load" functionality for getting image files into PD.  There are a number of different ways to do this;
' for example, loading a user-facing image file is a horrifically complex affair, with lots of messy work involved in metadata parsing,
' UI prep, Undo/Redo stuff, and more.  Conversely, loading an image file as a resource or internal image can bypass a lot of those steps.
'
'Note that these high-level functions call into a number of lower-level functions inside the ImageImporter module, and potentially various
' plugin-specific interfaces (e.g. FreeImage).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This function is used for loading a user-facing image (vs loading an internal PD image).  Loading a user-facing image involves
' a large amount of extra work (like metadata parsing) which we simply don't care about when loading internal resources.
'
'Note that this function will use one of several backends to load a given image; different filetypes are preferentially handled by
' different means, so portions of this function may call into external DLLs for parts of its functionality.  (The interaction between
' this function and various plugins is complex; I recommend studying the separate ImageImporter module for details.)
'
'INPUTS:
' 1) srcFile: fully qualified, absolute path to the source image.  Unicode is fully supported.
' 2) [optional] suggestedFilename: if loading an image from a temp file (e.g. clipboard, scanner), this value will be used in two places:
'                                  as the image's window caption, prior to first-save, and as the suggested filename at first-save.  As such,
'                                  make it user-friendly, e.g. "Clipboard image".
'                                  If this parameter is *not* supplied, the image's current filename will automatically be used.
' 3) [optional] addToRecentFiles: when a file loads successfully, we typically add it to the File > Recent Files list.  Some load operations,
'                                 like "Add new layer from file", or restoring a file from Autosave, don't easily fit into this paradigm.
'                                 This value tells the load engine to skip the "add to recent files" step.
' 4) [optional] suspendWarnings: at times, the caller may not want to have UI warnings raised for malformed or invalid files.  Batch processing
'                                and multi-image load are two examples.  If suspendWarnings = TRUE, any user-facing messages related to
'                                bad files will be suppressed.  (Note that the warnings can still be retrieved from debug logs, however.)
Public Function LoadFileAsNewImage(ByRef srcFile As String, Optional ByVal suggestedFilename As String = vbNullString, Optional ByVal addToRecentFiles As Boolean = True, Optional ByVal suspendWarnings As Boolean = False, Optional ByVal handleUIDisabling As Boolean = True) As Boolean
    
    '*** AND NOW, AN IMPORTANT MESSAGE ABOUT DOEVENTS ***
    
    'Normally, PD avoids DoEvents for all the obvious reasons.  This function is a stark exception to that rule.
    ' Why?
    
    'While this function stays busy loading the image in question, the ExifTool plugin runs asynchronously,
    ' parsing image metadata and forwarding the results to a pdAsyncPipe instance on PD's primary form.
    ' By using DoEvents throughout this function, we periodically yield control to that pdAsyncPipe instance,
    ' which allows it to clear stdout so ExifTool can continue pushing metadata through.  (If we don't do this,
    ' ExifTool will freeze when stdout fills its buffer, which is not just possible but *probable*, given how
    ' much metadata the average JPEG contains.)
    
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
        If (Not suspendWarnings) Then
            Message "Warning - file not found: %1", srcFile
            PDMsgBox "Unfortunately, the image '%1' could not be found." & vbCrLf & vbCrLf & "If this image was originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbExclamation Or vbOKOnly, "File not found", srcFile
        End If
        LoadFileAsNewImage = False
        Exit Function
    End If
    
    If (Not Files.FileTestAccess_Read(srcFile)) Then
        If handleUIDisabling Then Processor.MarkProgramBusyState False, True
        If (Not suspendWarnings) Then
            Message "Warning - file locked: %1", srcFile
            PDMsgBox "Unfortunately, the file '%1' is currently locked by another program on this PC." & vbCrLf & vbCrLf & "Please close this file in any other running programs, then try again.", vbExclamation Or vbOKOnly, "File locked", srcFile
        End If
        LoadFileAsNewImage = False
        Exit Function
    End If
    
    'Now we get into the meat-and-potatoes portion of this sub.  Main segments are labeled by large, asterisk-separated bars.
    ' These segments generally describe a group of tasks with related purpose, and many of these tasks branch out into other modules.
    
        
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
    
    'Normally, we don't assign an ID value to an image until we actually add it to the master pdImages collection.  However, some tasks
    ' (like retrieving metadata asynchronously) require an ID so we can synchronize incoming data post-load.  Give the target image
    ' a provisional image ID; this ID will become its formal ID only if it's loaded successfully.
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
    
    If (internalFormatID = PDIF_UNKNOWN) Then
    
        'Note that FreeImage may raise additional dialogs (e.g. for HDR/RAW images), so it does not return a binary pass/fail.
        ' If the function fails due to user cancellation, we will suppress subsequent error message boxes.
        loadSuccessful = ImageImporter.CascadeLoadGenericImage(srcFile, targetImage, targetDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
        
        '*************************************************************************************************************************************
        ' If the ExifTool plugin is available and this is a non-PD-specific file, initiate a separate thread for metadata extraction
        '*************************************************************************************************************************************
        If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) And (internalFormatID <> PDIF_PDI) And (internalFormatID <> PDIF_RAWBUFFER) Then
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
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_ICO)
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_MBM)
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_ORA)
            layersAlreadyLoaded = layersAlreadyLoaded Or ((targetImage.GetCurrentFileFormat = PDIF_PSD) And (decoderUsed = id_PSDParser))
            layersAlreadyLoaded = layersAlreadyLoaded Or (targetImage.GetCurrentFileFormat = PDIF_PSP)
            layersAlreadyLoaded = layersAlreadyLoaded Or ((targetImage.GetCurrentFileFormat = PDIF_WEBP) And (decoderUsed = id_libwebp))
            
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
        
        'The finalized pdImage object is finally worthy of being added to the master PD collection.  Note that this function will
        ' automatically update PDImages.GetActiveImageID() to point to the new image.
        PDImages.AddImageToMasterCollection targetImage
        
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
        
        If imageHasMultiplePages And ((targetImage.GetOriginalFileFormat = PDIF_TIFF) Or (targetImage.GetOriginalFileFormat = PDIF_GIF) Or (targetImage.GetOriginalFileFormat = PDIF_PNG) Or (targetImage.GetOriginalFileFormat = PDIF_AVIF)) Then
            
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
                If (targetImage.GetOriginalFileFormat = PDIF_PNG) Or (targetImage.GetOriginalFileFormat = PDIF_AVIF) Then loadSuccessful = ImageImporter.LoadRemainingPNGFrames(targetImage)
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
        PDDebug.LogAction "(FYI, expected delta was approximately " & Format$(targetImage.EstimateRAMUsage \ 1000, "#,#") & " K)"
        
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
    
    'Restore the screen cursor if necessary
    If handleUIDisabling Then Processor.MarkProgramBusyState False, True, (PDImages.GetNumOpenImages > 1)
    
    'Report success/failure back to the user
    LoadFileAsNewImage = (loadSuccessful And (Not targetImage Is Nothing))
    
    'Activate the new image (if loading was successful) and exit
    If LoadFileAsNewImage Then
        If handleUIDisabling Then CanvasManager.ActivatePDImage PDImages.GetActiveImageID(), "LoadFileAsNewImage", , , True
        Message "Image loaded successfully."
    Else
        If (Macros.GetMacroStatus <> MacroBATCH) And (Not suspendWarnings) And (freeImage_Return <> PD_FAILURE_USER_CANCELED) Then
            Message "Failed to load %1", srcFile
            PDMsgBox "Unfortunately, PhotoDemon was unable to load the following image:" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save this image in a generic format (such as JPEG or PNG) before loading it.  Thanks!", vbExclamation Or vbOKOnly, "Image import failed", srcFile
        End If
    End If
    
    PDDebug.LogAction "Image loaded in " & Format$(VBHacks.GetTimerDifferenceNow(startTime) * 1000, "#0") & " ms"
        
End Function

'Quick and dirty function for loading an image file to a pdDIB object (*NOT* a pdImage object).
'
'This function provides none of the extra scans or features that the more advanced LoadFileAsNewImage() does;
' instead, it assumes that the calling function will handle any extra work.
' (Note that things like metadata will not be processed *at all* for the image file.)
'
'That said, internal decoders and FreeImage/GDI+ are still used intelligently, so this function should reflect
' PD's full capacity for image format support.  Importantly, however, multi-page files will be squashed into
' single-frame composite forms, by design.
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
            If ImageFormats.IsFreeImageEnabled() Then loadSuccessful = (FI_LoadImage_V5(imagePath, targetDIB, , False, , suppressDebugData) = PD_SUCCESS)
            If (Not loadSuccessful) Then loadSuccessful = LoadGDIPlusImage(imagePath, targetDIB, tmpPDImage)
            If (Not loadSuccessful) Then loadSuccessful = LoadRawImageBuffer(imagePath, targetDIB, tmpPDImage)
            
        'PDTMP files are custom PD-format files saved ONLY during Undo/Redo or Autosaving.  As such, they have some weirdly specific
        ' parsing criteria during the master load function, but for quick-loading, we can simply grab the raw image buffer portion.
        Case "PDTMP"
            loadSuccessful = LoadRawImageBuffer(imagePath, targetDIB, tmpPDImage)
        
        'Internal decoders follow
        Case "CBZ"
            Dim cCBZ As pdCBZ
            Set cCBZ = New pdCBZ
            If cCBZ.IsFileCBZ(imagePath) Then loadSuccessful = cCBZ.LoadCBZ(imagePath, tmpPDImage)
            If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
            
        Case "JLS"
            loadSuccessful = Plugin_CharLS.LoadJLS(imagePath, tmpPDImage, targetDIB)
        
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
            
        Case "PNG"
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
        
        'AVIF support was provisionally added in v9.0.  Loading requires 64-bit Windows and manual
        ' copying of the official libavif exe binaries (for example,
        ' https://github.com/AOMediaCodec/libavif/releases/tag/v0.9.0)
        '...into the /App/PhotoDemon/Plugins subfolder.
        Case "HEIF", "HEIFS", "HEIC", "HEICS", "AVCI", "AVCS", "AVIF", "AVIFS"
            If Plugin_AVIF.IsAVIFImportAvailable() Then
            
                Dim tmpFile As String, intermediaryPDIF As PD_IMAGE_FORMAT
                loadSuccessful = Plugin_AVIF.ConvertAVIFtoStandardImage(imagePath, tmpFile, intermediaryPDIF)
                
                If loadSuccessful Then
                    If (intermediaryPDIF = PDIF_PNG) Then
                        Set cPNG = New pdPNG
                        loadSuccessful = (cPNG.LoadPNG_Simple(tmpFile, tmpPDImage, targetDIB) < png_Failure)
                    Else
                        loadSuccessful = LoadGDIPlusImage(tmpFile, targetDIB, tmpPDImage)
                    End If
                End If
                
                Files.FileDeleteIfExists tmpFile
                If (Not targetDIB.GetAlphaPremultiplication) Then targetDIB.SetAlphaPremultiplication True
                If loadSuccessful Then tmpPDImage.GetCompositedImage targetDIB, True
                
            End If
        
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
            If tryGDIPlusFirst Then loadSuccessful = LoadGDIPlusImage(imagePath, targetDIB, tmpPDImage)
            
            'If GDI+ failed, proceed with FreeImage
            If (Not loadSuccessful) And ImageFormats.IsFreeImageEnabled() Then
                
                Dim freeImageReturn As PD_OPERATION_OUTCOME
                freeImageReturn = PD_FAILURE_GENERIC
                freeImageReturn = FI_LoadImage_V5(imagePath, targetDIB, 0, False, , suppressDebugData)
                loadSuccessful = (freeImageReturn = PD_SUCCESS)
                
                'If FreeImage failed and we haven't tried GDI+ yet, try it now
                If (Not loadSuccessful) And (Not tryGDIPlusFirst) Then loadSuccessful = LoadGDIPlusImage(imagePath, targetDIB, tmpPDImage)
                
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
            GetDecoderName = "FreeImage plugin"
        Case id_CBZParser
            GetDecoderName = "Internal CBZ parser"
        Case id_ICOParser
            GetDecoderName = "Internal ICO parser"
        Case id_PDIParser
            GetDecoderName = "Internal PDI parser"
        Case id_MBMParser
            GetDecoderName = "Internal MBM parser"
        Case id_ORAParser
            GetDecoderName = "Internal OpenRaster parser"
        Case id_PNGParser
            GetDecoderName = "Internal PNG parser"
        Case id_PSDParser
            GetDecoderName = "Internal PSD parser"
        Case id_PSPParser
            GetDecoderName = "Internal PaintShop Pro parser"
        Case id_WIC
            GetDecoderName = "Windows Imaging Component"
        Case id_CharLS
            GetDecoderName = "CharLS plugin"
        Case id_libavif
            GetDecoderName = "libavif plugin"
        Case id_libwebp
            GetDecoderName = "libwebp plugin"
        Case Else
            GetDecoderName = "unknown?!"
    End Select
End Function

'Want to load a whole bunch of image sources at once?  Use this function to do so.  While helpful, note that it comes with some caveats:
' 1) The only supported sources are absolute filenames.
' 2) You lose the ability to assign custom titles to incoming images.  Titles will be auto-assigned based on their filenames.
' 3) You won't receive detailed success/failure information on each file.  Instead, this function will return TRUE if it was able to load
'    at least one image successfully.  If you want per-file success/fail results, call LoadFileAsNewImage manually from your own loop.
Public Function LoadMultipleImageFiles(ByRef srcList As pdStringStack, Optional ByVal updateRecentFileList As Boolean = True) As Boolean

    If (Not srcList Is Nothing) Then
        
        'A lot can go wrong when loading image files.  This function will track failures and notify the user post-load.
        Dim numFailures As Long, numSuccesses As Long
        Dim brokenFiles As String
        
        Processor.MarkProgramBusyState True, True
        
        Dim tmpFilename As String
        Do While srcList.PopString(tmpFilename)
            
            'The command-line may include other switches besides just filenames.  Ensure target file
            ' exists before forwarding it to the loader.
            If Files.FileExists(tmpFilename) Then
                
                If LoadFileAsNewImage(tmpFilename, , updateRecentFileList, True, False) Then
                    numSuccesses = numSuccesses + 1
                Else
                    If (LenB(tmpFilename) <> 0) Then
                        numFailures = numFailures + 1
                        brokenFiles = brokenFiles & Files.FileGetName(tmpFilename) & vbCrLf
                    End If
                End If
                
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
            CanvasManager.ActivatePDImage imgStack.GetInt(imgStack.GetNumOfInts - 1), "LoadFileAsNewImage", , , True
        End If
        
        'Synchronize everything to all open images
        SyncInterfaceToCurrentImage
        Processor.MarkProgramBusyState False, True, (PDImages.GetNumOpenImages() > 1)
        
        'Even if returning TRUE, we still want to notify the user of any failed files
        If (numFailures > 0) Then
            PDMsgBox "Unfortunately, PhotoDemon was unable to load the following image(s):" & vbCrLf & vbCrLf & "%1" & vbCrLf & "Please verify that these image(s) exist, and that they use a supported image format (like JPEG or PNG).  Thanks!", vbExclamation Or vbOKOnly, "Some images were not loaded", brokenFiles
        End If
        
    Else
        LoadMultipleImageFiles = False
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
