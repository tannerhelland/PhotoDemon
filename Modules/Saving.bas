Attribute VB_Name = "Saving"
'***************************************************************************
'File Saving Interface
'Copyright 2001-2022 by Tanner Helland
'Created: 4/15/01
'Last updated: 25/February/22
'Last update: add QOI export
'
'This module handles high-level image export duties.  Low-level export functions
' are generally located in the ImageExport module; see there for per-format details.
'
'The most important function here is PhotoDemon_SaveImage at the top of the module.
' This function is responsible for a multitude of decision-making related to exporting
' an image, including tasks like raising format-specific save dialogs, determining what
' color-depth to use, and various post-save housekeeping (like MRU updates).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_PDIWriter As pdPackageChunky

'When a Save request is invoked, call this function to determine if Save As is needed instead.  (Several factors can
' affect whether Save is okay; for example, if an image has never been saved before, we must raise a dialog to ask
' for a save location and filename.)
Public Function IsCommonDialogRequired(ByRef srcImage As pdImage) As Boolean
    
    'At present, this heuristic is pretty simple: if the image hasn't been saved to disk before, require a Save As instead.
    IsCommonDialogRequired = (LenB(srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) = 0)

End Function

'This routine will blindly save the composited layer contents (from the pdImage object specified by srcPDImage) to dstPath.
' It is up to the calling routine to make sure this is what is wanted. (Note: this routine will erase any existing image
' at dstPath, so BE VERY CAREFUL with what you send here!)
'
'INPUTS:
'   1) pdImage to be saved
'   2) Destination file path
'   3) Optional: whether to force display of an "additional save options" dialog (JPEG quality, etc).  Save As commands
'      forcibly set this to TRUE, so that the user can input new export settings.
Public Function PhotoDemon_SaveImage(ByRef srcImage As pdImage, ByVal dstPath As String, Optional ByVal forceOptionsDialog As Boolean = False) As Boolean
    
    'There are a few different ways the save process can "fail":
    ' 1) a save dialog with extra options is required, and the user cancels it
    ' 2) file-system errors (folder not writable, not enough free space, etc)
    ' 3) save engine errors (e.g. FreeImage explodes mid-save)
    
    'These have varying degrees of severity, but I mention this in advance because a number of post-save behaviors (like updating
    ' the Recent Files list) are abandoned under *any* of these occurrences.  As such, a lot of this function postpones various
    ' tasks until after all possible failure states have been dealt with.
    Dim saveSuccessful As Boolean: saveSuccessful = False
    
    'The caller must tell us which format they want us to use.
    Dim saveFormat As PD_IMAGE_FORMAT
    saveFormat = srcImage.GetCurrentFileFormat
    
    Dim dictEntry As String
    
    'The first major task this function deals with is save prompts.  The formula for showing these is hierarchical:
    
    ' 0) SPECIAL STEP: if we are in the midst of a batch process, *never* display a dialog.
    ' 1) If the caller has forcibly requested an options dialog (e.g. "Save As"), display a dialog.
    ' 2) If the caller hasn't forcibly requested a dialog...
        '3) See if this output format even supports dialogs.  If it doesn't, proceed with saving.
        '4) If this output format does support a dialog...
            '5) If the user has already seen a dialog for this format, don't show one again
            '6) If the user hasn't already seen a dialog for this format, it's time to show them one!
    
    'We'll deal with each of these in turn.
    Dim needToDisplayDialog As Boolean: needToDisplayDialog = forceOptionsDialog
    
    'Make sure we're not in the midst of a batch process operation
    If (Macros.GetMacroStatus <> MacroBATCH) Then
        
        'See if this format even supports dialogs...
        If ImageFormats.IsExportDialogSupported(saveFormat) Then
        
            'If the caller did *not* specifically request a dialog, run some heuristics to see if we need one anyway
            ' (e.g. if this the first time saving a JPEG file, we need to query the user for a Quality value)
            If (Not forceOptionsDialog) Then
            
                'See if the user has already seen this dialog...
                dictEntry = GetExportDictFlag("HasSeenExportDialog", saveFormat, srcImage)
                needToDisplayDialog = Not srcImage.ImgStorage.GetEntry_Boolean(dictEntry, False)
                
                'If the user has seen a dialog, we'll perform one last failsafe check.  Make sure that the
                ' exported format's parameter string exists; if it doesn't, we need to prompt them again.
                ' (This ensures that the user sees at least *1* save settings dialog per session, per format.)
                dictEntry = GetExportDictFlag("ExportParams", saveFormat, srcImage)
                If (Not needToDisplayDialog) And (LenB(srcImage.ImgStorage.GetEntry_String(dictEntry, vbNullString)) = 0) Then
                    PDDebug.LogAction "WARNING!  PhotoDemon_SaveImage found an image where HasSeenExportDialog = TRUE, but ExportParams = null.  Fix this!"
                    needToDisplayDialog = True
                End If
                
            End If
        
        'If this format doesn't support an export dialog, forcibly reset the forceOptionsDialog parameter to match
        Else
            needToDisplayDialog = False
        End If
        
    Else
        needToDisplayDialog = False
    End If
    
    'All export dialogs fulfill the same purpose: they fill an XML string with a list of key+value pairs detailing setting relevant
    ' to that format.  This XML string is then passed to the respective save function, which applies the settings as relevant.
    
    'Upon a successful save, we cache that format-specific parameter string inside the parent image; the same settings are then
    ' reused on subsequent saves, instead of re-prompting the user.
    
    'It is now time to retrieve said parameter string, either from a dialog, or from the pdImage settings dictionary.
    Dim saveParameters As String, metadataParameters As String
    If needToDisplayDialog Then
        
        'After a successful dialog invocation, immediately save the metadata parameters to the parent pdImage object.
        ' ExifTool will handle those settings separately, independent of the format-specific export engine.
        Dim useAnimationDialog As Boolean
        If (Not srcImage Is Nothing) Then useAnimationDialog = srcImage.IsAnimated Else useAnimationDialog = False
        If Saving.GetExportParamsFromDialog(srcImage, saveFormat, saveParameters, metadataParameters, useAnimationDialog) Then
            srcImage.ImgStorage.AddEntry "MetadataSettings", metadataParameters
            
        'If the user cancels the dialog, exit immediately
        Else
            Message "Save canceled."
            PhotoDemon_SaveImage = False
            Exit Function
        End If
        
    Else
        dictEntry = GetExportDictFlag("ExportParams", saveFormat, srcImage)
        saveParameters = srcImage.ImgStorage.GetEntry_String(dictEntry, vbNullString)
        metadataParameters = srcImage.ImgStorage.GetEntry_String("MetadataSettings", vbNullString)
    End If
    
    'Before proceeding with the save, check for some file-level errors that may cause problems.
    
    'If the file already exists, ensure we have write+delete access
    If (Not Files.FileTestAccess_Write(dstPath)) Then
        Message "Warning - file locked: %1", dstPath
        PDMsgBox "Unfortunately, the file '%1' is currently locked by another program on this PC." & vbCrLf & vbCrLf & "Please close this file in any other running programs, then try again.", vbExclamation Or vbOKOnly, "File locked", dstPath
        PhotoDemon_SaveImage = False
        Exit Function
    End If
    
    'As saving can be somewhat lengthy for large images and/or complex formats, lock the UI now.  Note that we *must* call
    ' the "EndSaveProcess" function to release the UI lock.
    BeginSaveProcess
    Message "Saving %1 file...", UCase$(ImageFormats.GetExtensionFromPDIF(saveFormat))
    
    'If the image is being saved to a layered format (like multipage TIFF), various parts of the export engine may
    ' want to inject useful information into the finished file (e.g. ExifTool can append things like page names).
    ' Mark the outgoing file now.
    MarkMultipageExportStatus srcImage, saveFormat, saveParameters, metadataParameters
    
    'With all save parameters collected, we can offload the rest of the save process to per-format save functions.
    saveSuccessful = Saving.ExportToSpecificFormat(srcImage, dstPath, saveFormat, saveParameters, metadataParameters)
    If saveSuccessful Then
        
        'The file was saved successfully!  Copy the save parameters into the parent pdImage object; subsequent "save" actions
        ' can use these instead of querying the user again.
        dictEntry = GetExportDictFlag("ExportParams", saveFormat, srcImage)
        srcImage.ImgStorage.AddEntry dictEntry, saveParameters
        
        'If a dialog was displayed, note that as well
        If needToDisplayDialog Then
            dictEntry = GetExportDictFlag("HasSeenExportDialog", saveFormat, srcImage)
            srcImage.ImgStorage.AddEntry dictEntry, True
        End If
        
        'Similarly, remember the file's location and selected name for future saves
        srcImage.ImgStorage.AddEntry "CurrentLocationOnDisk", dstPath
        srcImage.ImgStorage.AddEntry "OriginalFileName", Files.FileGetName(dstPath, True)
        srcImage.ImgStorage.AddEntry "OriginalFileExtension", Files.FileGetExtension(dstPath)
        
        'Update the parent image's save state.
        If (saveFormat = PDIF_PDI) Then srcImage.SetSaveState True, pdSE_SavePDI Else srcImage.SetSaveState True, pdSE_SaveFlat
        
        'If the file was successfully written, we can now embed any additional metadata.
        ' (Note: I don't like embedding metadata in a separate step, but that's a necessary evil of routing all metadata handling
        ' through an external plugin.  Exiftool requires an existant file to be used as a target, and an existant metadata file
        ' to be used as its source.  It cannot operate purely in-memory - but hey, that's why it's asynchronous!)
        If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) And (Not srcImage.ImgMetadata Is Nothing) Then
            
            'Some export formats aren't supported by ExifTool; we don't even attempt to write metadata on such images
            If ImageFormats.IsExifToolRelevant(saveFormat) Then srcImage.ImgMetadata.WriteAllMetadata dstPath, srcImage
            
        End If
        
        'With all save work complete, we can now update various UI bits to reflect the new image.  Note that these changes are
        ' only applied if we are *not* in the midst  of a batch conversion.
        If (Macros.GetMacroStatus <> MacroBATCH) Then
            g_RecentFiles.AddFileToList dstPath, srcImage
            Interface.SyncInterfaceToCurrentImage
            Interface.NotifyImageChanged PDImages.GetActiveImageID()
            Menus.UpdateSpecialMenu_WindowsOpen
        End If
        
        'At this point, it's safe to re-enable the main form and restore the default cursor
        EndSaveProcess
        
        Message "Save complete."
    
    'If something went wrong during the save process, the exporter likely provided its own error report.  Attempt to assemble
    ' a meaningful message for the user.
    Else
    
        Message "Save canceled."
        EndSaveProcess
        
        'If FreeImage failed, it should have provided detailed information on the problem.  Present it to the user, in hopes that
        ' they might use it to rectify the situation (or least notify us of what went wrong!)
        If Plugin_FreeImage.FreeImageErrorState Then
            
            Dim fiErrorList As String
            fiErrorList = Plugin_FreeImage.GetFreeImageErrors
            
            'Display the error message
            PDMsgBox "An error occurred when attempting to save this image.  The FreeImage plugin reported the following error details: " & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "In the meantime, please try saving the image to an alternate format.  You can also let the PhotoDemon developers know about this via the Help > Submit Bug Report menu.", vbCritical Or vbOKOnly, "Error", fiErrorList
            
        Else
            PDMsgBox "An unspecified error occurred when attempting to save this image.  Please try saving the image to an alternate format." & vbCrLf & vbCrLf & "If the problem persists, please report it to the PhotoDemon developers via photodemon.org/contact", vbCritical Or vbOKOnly, "Error"
        End If
        
    End If
    
    PhotoDemon_SaveImage = saveSuccessful
    
End Function

'This _BatchSave() function is a shortened, accelerated version of the full _SaveImage() function above.
' It should *only* be used during Batch Process operations, where there is no possibility of user interaction.
' Note that the input parameters are different, as the batch processor requires the user to set most export
' settings in advance (since we can't raise export dialogs mid-batch).
Public Function PhotoDemon_BatchSaveImage(ByRef srcImage As pdImage, ByVal dstPath As String, ByVal saveFormat As PD_IMAGE_FORMAT, Optional ByVal saveParameters As String = vbNullString, Optional ByVal metadataParameters As String = vbNullString) As Boolean
    
    'The important thing to note about this function is that it *requires* the image to be immediately unloaded
    ' after the save operation finishes.  To improve performance, the source pdImage object is not updated against
    ' any changes incurred by the save operation, so that object *will* be "corrupted" after a save operation occurs.
    ' (Note also that things like failed saves cannot raise any modal dialogs, so the only notification of failure
    ' is the return value of this function.)
    Dim saveSuccessful As Boolean: saveSuccessful = False
    
    'If the image is being saved to a layered format (like multipage TIFF), various parts of the export engine may
    ' want to inject useful information into the finished file (e.g. ExifTool can append things like page names).
    ' Mark the outgoing file now.
    srcImage.ImgStorage.AddEntry "MetadataSettings", metadataParameters
    MarkMultipageExportStatus srcImage, saveFormat, saveParameters, metadataParameters
    
    'With all save parameters collected, we can offload the rest of the save process to per-format save functions.
    saveSuccessful = Saving.ExportToSpecificFormat(srcImage, dstPath, saveFormat, saveParameters, metadataParameters)
    
    If saveSuccessful Then
        
        'If the file was successfully written, we can now embed any additional metadata.
        ' (Note: I don't like embedding metadata in a separate step, but that's a necessary evil of routing all metadata handling
        ' through an external plugin.  Exiftool requires an existant file to be used as a target, and an existant metadata file
        ' to be used as its source.  It cannot operate purely in-memory - but hey, that's why it's asynchronous!)
        If (PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) And (Not srcImage.ImgMetadata Is Nothing) And (saveFormat <> PDIF_PDI)) Then
            
            'Sometimes, PD may process images faster than ExifTool can parse the source file's metadata.
            ' Check for this, and pause until metadata processing catches up.
            If ExifTool.IsMetadataPipeActive Then
                
                PDDebug.LogAction "Pausing batch process so that metadata processing can catch up..."
                
                Do While ExifTool.IsMetadataPipeActive
                    VBHacks.SleepAPI 50
                    DoEvents
                Loop
                
                PDDebug.LogAction "Metadata processing caught up; proceeding with batch operation..."
                
            End If
            
            srcImage.ImgMetadata.WriteAllMetadata dstPath, srcImage
            
            Do While ExifTool.IsVerificationModeActive
                VBHacks.SleepAPI 50
                DoEvents
            Loop
            
        End If
        
    End If
    
    PhotoDemon_BatchSaveImage = saveSuccessful
    
End Function

Private Function GetExportDictFlag(ByRef categoryName As String, ByVal dstFileFormat As PD_IMAGE_FORMAT, ByRef srcImage As pdImage) As String
    GetExportDictFlag = categoryName & ImageFormats.GetExtensionFromPDIF(dstFileFormat)
    If ImageFormats.IsExportDialogSupported(dstFileFormat) Then
        If srcImage.IsAnimated Then GetExportDictFlag = GetExportDictFlag & "-animated"
    End If
End Function

Private Sub MarkMultipageExportStatus(ByRef srcImage As pdImage, ByVal outputPDIF As PD_IMAGE_FORMAT, Optional ByVal saveParameters As String = vbNullString, Optional ByVal metadataParameters As String = vbNullString)
    
    Dim saveIsMultipage As Boolean: saveIsMultipage = False
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString saveParameters
    
    'TIFF is currently the only image format that supports multipage export as an option.
    ' (For all other formats, it is handled automatically, e.g. animated GIFs are rerouted to the
    ' animation exporter, PSDs are written as multi-layer files, etc.)
    If (outputPDIF = PDIF_TIFF) Then
    
        'The format parameter string contains the multipage indicator, if any.  (Default is to write a single-page TIFF.)
        If cParams.GetBool("tiff-multipage", False) Then saveIsMultipage = True
        
    End If
    
    'If the outgoing image is multipage, add a special dictionary entry that other functions can easily test.
    srcImage.ImgStorage.AddEntry "MultipageExportActive", saveIsMultipage
    
End Sub

'Given a source image, a desired export format, and a destination string, fill the destination string with format-specific parameters
' returned from the associated format-specific dialog.
'
'Returns: TRUE if dialog was closed via OK button; FALSE otherwise.
Public Function GetExportParamsFromDialog(ByRef srcImage As pdImage, ByVal outputPDIF As PD_IMAGE_FORMAT, ByRef dstParamString As String, ByRef dstMetadataString As String, Optional ByVal displayAnimationVersion As Boolean = False) As Boolean
    
    'As a failsafe, make sure the requested format even *has* an export dialog!
    If ImageFormats.IsExportDialogSupported(outputPDIF) Then
        
        Select Case outputPDIF
            
            Case PDIF_AVIF
                GetExportParamsFromDialog = (Dialogs.PromptAVIFSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_BMP
                GetExportParamsFromDialog = (Dialogs.PromptBMPSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_GIF
                If displayAnimationVersion Then
                    GetExportParamsFromDialog = (Dialogs.PromptExportAnimatedGIF(srcImage, dstParamString, dstMetadataString) = vbOK)
                Else
                    GetExportParamsFromDialog = (Dialogs.PromptGIFSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
                End If
            
            Case PDIF_ICO
                GetExportParamsFromDialog = (Dialogs.PromptICOSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_JP2
                GetExportParamsFromDialog = (Dialogs.PromptJP2Settings(srcImage, dstParamString, dstMetadataString) = vbOK)
                
            Case PDIF_JPEG
                GetExportParamsFromDialog = (Dialogs.PromptJPEGSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
                
            Case PDIF_JXR
                GetExportParamsFromDialog = (Dialogs.PromptJXRSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
        
            Case PDIF_PNG
                If displayAnimationVersion Then
                    GetExportParamsFromDialog = (Dialogs.PromptExportAnimatedPNG(srcImage, dstParamString, dstMetadataString) = vbOK)
                Else
                    GetExportParamsFromDialog = (Dialogs.PromptPNGSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
                End If
                
            Case PDIF_PNM
                GetExportParamsFromDialog = (Dialogs.PromptPNMSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_PSD
                GetExportParamsFromDialog = (Dialogs.PromptPSDSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_PSP
                GetExportParamsFromDialog = (Dialogs.PromptPSPSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_TIFF
                GetExportParamsFromDialog = (Dialogs.PromptTIFFSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_WEBP
                If displayAnimationVersion Then
                    GetExportParamsFromDialog = (Dialogs.PromptExportAnimatedWebP(srcImage, dstParamString, dstMetadataString) = vbOK)
                Else
                    GetExportParamsFromDialog = (Dialogs.PromptWebPSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
                End If
                
        End Select
        
    Else
        GetExportParamsFromDialog = False
        dstParamString = vbNullString
    End If
        
End Function

'Already have a save parameter string assembled?  Call this function to export directly to a given format, with no UI prompts.
' (I *DO NOT* recommend calling this function directly.  PD only uses it from within the main _SaveImage function, which also applies
'  a number of failsafe checks against things like path accessibility and format compatibility.)
Private Function ExportToSpecificFormat(ByRef srcImage As pdImage, ByRef dstPath As String, ByVal outputPDIF As PD_IMAGE_FORMAT, Optional ByVal saveParameters As String = vbNullString, Optional ByVal metadataParameters As String = vbNullString) As Boolean
    
    If (srcImage Is Nothing) Then Exit Function
    
    'Generate perf reports on export; this is useful for regression testing
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    Select Case outputPDIF
        
        Case PDIF_AVIF
            ExportToSpecificFormat = ImageExporter.ExportAVIF(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_BMP
            ExportToSpecificFormat = ImageExporter.ExportBMP(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_GIF
            If srcImage.IsAnimated Then
                ExportToSpecificFormat = ImageExporter.ExportGIF_Animated(srcImage, dstPath, saveParameters, metadataParameters)
            Else
                ExportToSpecificFormat = ImageExporter.ExportGIF(srcImage, dstPath, saveParameters, metadataParameters)
            End If
            
        Case PDIF_HDR
            ExportToSpecificFormat = ImageExporter.ExportHDR(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_ICO
            ExportToSpecificFormat = ImageExporter.ExportICO(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_JP2
            ExportToSpecificFormat = ImageExporter.ExportJP2(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_JPEG
            ExportToSpecificFormat = ImageExporter.ExportJPEG(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_JXR
            ExportToSpecificFormat = ImageExporter.ExportJXR(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_ORA
            ExportToSpecificFormat = ImageExporter.ExportORA(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_PDI
            ExportToSpecificFormat = Saving.SavePDI_Image(srcImage, dstPath, False, cf_Zstd, cf_Zstd, False, True, Compression.GetDefaultCompressionLevel(cf_Zstd))
                        
        Case PDIF_PNG
            If srcImage.IsAnimated Then
                ExportToSpecificFormat = ImageExporter.ExportPNG_Animated(srcImage, dstPath, saveParameters, metadataParameters)
            Else
                ExportToSpecificFormat = ImageExporter.ExportPNG(srcImage, dstPath, saveParameters, metadataParameters)
            End If
        
        Case PDIF_PNM
            ExportToSpecificFormat = ImageExporter.ExportPNM(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_PSD
            ExportToSpecificFormat = ImageExporter.ExportPSD(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_PSP
            ExportToSpecificFormat = ImageExporter.ExportPSP(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_QOI
            ExportToSpecificFormat = ImageExporter.ExportQOI(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_TARGA
            ExportToSpecificFormat = ImageExporter.ExportTGA(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_TIFF
            ExportToSpecificFormat = ImageExporter.ExportTIFF(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_WEBP
            If srcImage.IsAnimated Then
                ExportToSpecificFormat = ImageExporter.ExportWebP_Animated(srcImage, dstPath, saveParameters, metadataParameters)
            Else
                ExportToSpecificFormat = ImageExporter.ExportWebP(srcImage, dstPath, saveParameters, metadataParameters)
            End If
            
        Case Else
            Message "Output format not recognized.  Save aborted.  Please use the Help -> Submit Bug Report menu item to report this incident."
            ExportToSpecificFormat = False
            
    End Select
    
    If ExportToSpecificFormat Then PDDebug.LogAction "Image export took " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Function

'Save a PDI file ("PhotoDemon Image", e.g. our native format)
' FUTURE TODO:
'  - It might be nice to store a copy of the fully composited image in the file, to simplify the work other software
'    has to do.  That said, this inevitably increases both export time and file size - and at present, PD isn't used
'    widely enough to warrant those trade-offs.
Public Function SavePDI_Image(ByRef srcPDImage As pdImage, ByRef dstFileAndPath As String, Optional ByVal suppressMessages As Boolean = False, Optional ByVal compressHeaders As PD_CompressionFormat = cf_Zstd, Optional ByVal compressLayers As PD_CompressionFormat = cf_Zstd, Optional ByVal writeHeaderOnlyFile As Boolean = False, Optional ByVal includeMetadata As Boolean = False, Optional ByVal compressionLevel As Long = -1, Optional ByVal srcIsUndo As Boolean = False, Optional ByRef dstUndoFileSize As Long) As Boolean
    
    On Error GoTo SavePDIError
    
    'Perform a few failsafe checks
    If (srcPDImage Is Nothing) Then Exit Function
    If (LenB(dstFileAndPath) = 0) Then Exit Function
    
    'Want to time this function?  Here's your chance:
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    Dim sFileType As String
    sFileType = "PDI"
    
    If (Not suppressMessages) Then Message "Saving %1 image...", sFileType
    
    'First things first: create a pdPackage instance and initialize it on the target file.
    ' It will handle the messy business of compressing various data bits into a running stream.
    ' (An important difference from past PDI writers is that we don't need to know the number
    ' of nodes or anything else in advance.  We don't even need to estimate a file size, as the
    ' memory-mapped file interface will silently handle that for us.)
    Dim pdiWriter As pdPackageChunky
    Set pdiWriter = New pdPackageChunky
    pdiWriter.StartNewPackage_File dstFileAndPath, False, srcPDImage.EstimateRAMUsage, "PDIF"
    
    'The first node we'll add is a standard pdImage header, in XML format.
    
    'Retrieve the layer header (in XML format), then write the XML stream to the package
    Dim dataString As String, dataUTF8() As Byte, utf8Len As Long
    dataString = srcPDImage.GetHeaderAsXML()
    Strings.UTF8FromStrPtr StrPtr(dataString), Len(dataString), dataUTF8, utf8Len
    pdiWriter.AddChunk_WholeFromPtr "IHDR", VarPtr(dataUTF8(0)), utf8Len, compressHeaders
    
    'Next, we will add each pdLayer object to the stream.  This is done in two steps:
    ' 1) First, obtain the layer header in XML format and write it out
    ' 2) Second, obtain any layer-specific data (DIB for raster layers, XML for vector layers) and write it out
    Dim i As Long
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        
        'Retrieve the layer header and add it to the stream.
        WriteLayerHeaderToPackage pdiWriter, srcPDImage.GetLayerByIndex(i), dataString, dataUTF8, compressHeaders
        
        'If this is not a header-only file, retrieve the layer's data (BGRA bytes for raster layers, XML for vector layers)
        ' and add it to the stream.
        If (Not writeHeaderOnlyFile) Then WriteLayerDataToPackage pdiWriter, srcPDImage.GetLayerByIndex(i), compressLayers, compressionLevel
        
    Next i
    
    'Next, if the "write metadata" flag has been set, and this image has metadata, add a metadata entry to the file.
    If includeMetadata And (Not srcPDImage.ImgMetadata Is Nothing) Then
        
        If srcPDImage.ImgMetadata.HasMetadata Then
            
            Dim mdStartTime As Currency
            VBHacks.GetHighResTime mdStartTime
            
            'To avoid unnecessary string copies, we write the (potentially large) original metadata string directly
            ' from its source pointer.
            Dim mdPtr As Long, mdLen As Long
            srcPDImage.ImgMetadata.GetOriginalXMLMetadataStrPtrAndLen mdPtr, mdLen
            
            If (mdLen <> 0) Then
            
                If Strings.UTF8FromStrPtr(mdPtr, mdLen, dataUTF8, utf8Len) Then
                    pdiWriter.AddChunk_WholeFromPtr "MDET", VarPtr(dataUTF8(0)), utf8Len, compressHeaders
                End If
                
                'Unfortunately, there's no similarly fast way to handle our already-parsed (and potentially modified
                ' by the user) metadata collection.  At present, we manually serialize it to a string and just
                ' write that <sigh>.
                dataString = srcPDImage.ImgMetadata.GetSerializedXMLData()
                If Strings.UTF8FromStrPtr(StrPtr(dataString), Len(dataString), dataUTF8, utf8Len) Then
                    pdiWriter.AddChunk_WholeFromPtr "MDPD", VarPtr(dataUTF8(0)), utf8Len, compressHeaders
                End If
                
                PDDebug.LogAction "Note: metadata writes took " & VBHacks.GetTimeDiffNowAsString(mdStartTime)
                
            Else
                Debug.Print "FYI, metadata string data is reported as zero-length; abandoning write"
            End If
            
        End If
        
    End If
    
    'That's all there is to it!  Write the completed pdPackage out to file.
    dstUndoFileSize = pdiWriter.GetPackageSize() + 8    '+8 for the final chunk in the file, which isn't written yet
    SavePDI_Image = pdiWriter.FinishPackage()
    
    'Report timing on debug builds
    If SavePDI_Image Then
        PDDebug.LogAction "Saved PDI file in " & CStr(VBHacks.GetTimerDifferenceNow(startTime) * 1000) & " ms."
    Else
        PDDebug.LogAction "WARNING!  SavePDI_Image failed after " & CStr(VBHacks.GetTimerDifferenceNow(startTime) * 1000) & " ms."
    End If
    
    If (Not suppressMessages) Then Message "Save complete."
    
    Exit Function
    
SavePDIError:
    PDDebug.LogAction "An error occurred in SavePDI_Image: " & Err.Number & " - " & Err.Description
    SavePDI_Image = False
    
End Function

Private Function SavePDI_SingleLayer(ByRef srcLayer As pdLayer, ByRef pdiPath As String, Optional ByVal compressHeaders As PD_CompressionFormat = cf_Zstd, Optional ByVal compressLayers As PD_CompressionFormat = cf_Zstd, Optional ByVal writeHeaderOnlyFile As Boolean = False, Optional ByVal compressionLevel As Long = -1, Optional ByRef dstUndoFileSize As Long) As Boolean

    On Error GoTo SavePDLayerError
    
    'Perform a few failsafe checks
    If (srcLayer Is Nothing) Then Exit Function
    If (srcLayer.GetLayerDIB Is Nothing) Then Exit Function
    If (LenB(pdiPath) = 0) Then Exit Function
    
    'Enable for detailed profiling
    Const REPORT_LAYER_SAVE_TIMING As Boolean = False
    Dim startTime As Currency
    If REPORT_LAYER_SAVE_TIMING Then VBHacks.GetHighResTime startTime
    
    Dim sFileType As String
    sFileType = "PDI"
    
    'First things first: create a pdPackage instance.  It handles the messy business of assembling
    ' the layer file (including all compression tasks).
    If (m_PDIWriter Is Nothing) Then Set m_PDIWriter = New pdPackageChunky
    
    'Unlike an actual PDI file, which stores a whole bunch of data, layer temp files only store
    ' two pieces of data: the layer header, and the DIB bytestream.
    m_PDIWriter.StartNewPackage_File pdiPath, False, , "UNDO"
    
    If REPORT_LAYER_SAVE_TIMING Then
        PDDebug.LogAction "Time required for allocate: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Retrieve the layer header (in XML format), then write the XML stream to the package
    Dim dataString As String, dataUTF8() As Byte
    WriteLayerHeaderToPackage m_PDIWriter, srcLayer, dataString, dataUTF8, compressHeaders
    
    If REPORT_LAYER_SAVE_TIMING Then
        PDDebug.LogAction "Time required for layer header: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'If this is not a header-only request, retrieve the layer DIB (as a byte array), then copy the array
    ' into the pdPackage instance
    If (Not writeHeaderOnlyFile) Then WriteLayerDataToPackage m_PDIWriter, srcLayer, compressLayers, compressionLevel
    
    If REPORT_LAYER_SAVE_TIMING Then PDDebug.LogAction "Time required for layer contents: " & VBHacks.GetTimeDiffNowAsString(startTime)
    
    'Report our finished package size to the caller
    dstUndoFileSize = m_PDIWriter.GetPackageSize()
    
    'That's everything!  Just remember to finalize the package before exiting.
    SavePDI_SingleLayer = m_PDIWriter.FinishPackage()
    If (Not SavePDI_SingleLayer) Then PDDebug.LogAction "WARNING!  SavingSavePDI_SingleLayer received a failure status from pdiWriter.WritePackageToFile()"
    
    Exit Function
    
SavePDLayerError:
    PDDebug.LogAction "WARNING!  Saving.SavePDI_SingleLayer failed with error #" & Err.Number & ", " & Err.Description
    SavePDI_SingleLayer = False
End Function

'Private function to dump a given pdLayer object's header to a running pdStream instance.
' This function is called by both save-image and save-layer functions; it is expected that these will
' always use the same format going forward.
Private Function WriteLayerHeaderToPackage(ByRef dstPackage As pdPackageChunky, ByRef srcLayer As pdLayer, ByRef dataString As String, ByRef dataUTF8() As Byte, Optional ByVal compressHeaders As PD_CompressionFormat = cf_Zstd, Optional ByVal compressionLevel As Long = -1) As Boolean
    Dim utf8Len As Long
    dataString = srcLayer.GetLayerHeaderAsXML()
    Strings.UTF8FromStrPtr StrPtr(dataString), Len(dataString), dataUTF8, utf8Len
    dstPackage.AddChunk_WholeFromPtr "LHDR", VarPtr(dataUTF8(0)), utf8Len, compressHeaders
    WriteLayerHeaderToPackage = True
End Function

'Private function to dump a given pdLayer object's data to a running pdStream instance.
' Raster and vector layers can both be passed.
Private Function WriteLayerDataToPackage(ByRef dstPackage As pdPackageChunky, ByRef srcLayer As pdLayer, Optional ByVal compressData As PD_CompressionFormat = cf_Zstd, Optional ByVal compressionLevel As Long = -1) As Boolean

    'Image layers save their pixel data as a raw byte stream
    If srcLayer.IsLayerRaster Then
    
        Dim layerDIBPointer As Long, layerDIBLength As Long
        srcLayer.GetLayerDIB.RetrieveDIBPointerAndSize layerDIBPointer, layerDIBLength
        dstPackage.AddChunk_WholeFromPtr "LDAT", layerDIBPointer, layerDIBLength, compressData, compressionLevel
        WriteLayerDataToPackage = True
        
    'Text (and other vector layers) save their vector contents in XML format
    ElseIf srcLayer.IsLayerVector Then
        
        Dim dataString As String, dataUTF8() As Byte, utf8Len As Long
        dataString = srcLayer.GetVectorDataAsXML()
        Strings.UTF8FromStrPtr StrPtr(dataString), Len(dataString), dataUTF8, utf8Len
        dstPackage.AddChunk_WholeFromPtr "LDAT", VarPtr(dataUTF8(0)), utf8Len, compressData, compressionLevel
        WriteLayerDataToPackage = True
        
    'Other layer types are not currently supported
    Else
        PDDebug.LogAction "WARNING!  WriterLayerDataToStream was passed a layer of unknown type."
        WriteLayerDataToPackage = False
    End If
    
End Function

'Save a new Undo/Redo entry to file.  This function is only called by the createUndoData function in the pdUndo class.
' For the most part, this function simply wraps other save functions; however, certain odd types of Undo diff files (e.g. layer headers)
' may be directly processed and saved by this function.
'
'Note that this function interacts closely with the matching LoadUndo function in the Loading module.  Any novel Undo diff types added
' here must also be mirrored there.
Public Function SaveUndoData(ByRef srcPDImage As pdImage, ByRef dstUndoFilename As String, ByVal processType As PD_UndoType, Optional ByVal targetLayerID As Long = -1, Optional ByVal compressionHint As Long = -1, Optional ByRef dstUndoFileSize As Long) As Boolean
    
    Dim timeAtUndoStart As Currency
    VBHacks.GetHighResTime timeAtUndoStart
    
    'As of v7.0, PD has multiple compression engines available.  These engines are not exposed to the user.  We use LZ4 by default,
    ' as it is far and away the fastest at both compression and decompression (while compressing at marginally worse ratios).
    ' Note that if the user selects increasingly better compression results, we will silently switch to zstd instead.
    Dim undoCmpEngine As PD_CompressionFormat, undoCmpLevel As Long
    If (g_UndoCompressionLevel = 0) Then
        undoCmpEngine = cf_None
        undoCmpLevel = 0
    
    'At level 1 (the current PD default), use LZ4 compression at default strength.  (Remember that LZ4's compression level do not
    ' improve as the level goes up - the algorithm's *performance* improves as the level goes up.)
    ElseIf (g_UndoCompressionLevel = 1) Then
        undoCmpEngine = cf_Lz4
        undoCmpLevel = compressionHint
    
    'For all higher levels, use zstd, and reset the compression level to start at 1 (so a g_UndoCompressionLevel of 2 uses zstd at
    ' its default compression strength of level 1).
    Else
        undoCmpEngine = cf_Zstd
        undoCmpLevel = g_UndoCompressionLevel - 1
    End If
    
    Dim undoSuccess As Boolean
    
    'What kind of Undo data we save is determined by the current processType.
    Select Case processType
    
        'EVERYTHING, meaning a full copy of the pdImage stack and any selection data
        Case UNDO_Everything
            Dim tmpFileSizeCheck As Long
            undoSuccess = Saving.SavePDI_Image(srcPDImage, dstUndoFilename, True, cf_Lz4, undoCmpEngine, False, True, undoCmpLevel, True, dstUndoFileSize)
            srcPDImage.MainSelection.WriteSelectionToFile dstUndoFilename & ".selection", undoCmpEngine, undoCmpLevel, undoCmpEngine, undoCmpLevel, tmpFileSizeCheck
            dstUndoFileSize = dstUndoFileSize + tmpFileSizeCheck
            
        'A full copy of the pdImage stack
        Case UNDO_Image, UNDO_Image_VectorSafe
            undoSuccess = Saving.SavePDI_Image(srcPDImage, dstUndoFilename, True, cf_Lz4, undoCmpEngine, False, True, undoCmpLevel, True, dstUndoFileSize)
        
        'A full copy of the pdImage stack, *without any layer DIB data*
        Case UNDO_ImageHeader
            undoSuccess = Saving.SavePDI_Image(srcPDImage, dstUndoFilename, True, undoCmpEngine, cf_None, True, True, undoCmpLevel, True, dstUndoFileSize)
        
        'Layer data only (full layer header + full layer DIB).
        Case UNDO_Layer, UNDO_Layer_VectorSafe
            undoSuccess = Saving.SavePDI_SingleLayer(srcPDImage.GetLayerByID(targetLayerID), dstUndoFilename & ".layer", cf_Zstd, undoCmpEngine, False, undoCmpLevel, dstUndoFileSize)
        
        'Layer header data only (e.g. DO NOT WRITE OUT THE LAYER DIB)
        Case UNDO_LayerHeader
            undoSuccess = Saving.SavePDI_SingleLayer(srcPDImage.GetLayerByID(targetLayerID), dstUndoFilename & ".layer", undoCmpEngine, cf_None, True, undoCmpLevel, dstUndoFileSize)
            
        'Selection data only
        Case UNDO_Selection
            undoSuccess = srcPDImage.MainSelection.WriteSelectionToFile(dstUndoFilename & ".selection", undoCmpEngine, undoCmpLevel, undoCmpEngine, undoCmpLevel)
            
        'Anything else (this should never happen, but good to have a failsafe)
        Case Else
            PDDebug.LogAction "Unknown Undo data write requested - is it possible to avoid this request entirely??"
            undoSuccess = Saving.SavePDI_Image(srcPDImage, dstUndoFilename, True, cf_Lz4, undoCmpEngine, False, , undoCmpLevel, True, dstUndoFileSize)
        
    End Select
    
    SaveUndoData = undoSuccess
    
    If (Not SaveUndoData) Then PDDebug.LogAction "SaveUndoData returned failure; cause unknown."
    'Want to test undo timing?  Uncomment the line below
    'pdDebug.LogAction "Undo file creation took: " & Format$(VBHacks.GetTimerDifferenceNow(timeAtUndoStart) * 1000, "#0.00") & " ms"
    
End Function

'Quickly save a DIB to file in PNG format.  At present, this is only used when forwarding image data
' to the Windows Photo Printer object.  (All internal quick-saves use PD-specific formats, which are
' much faster to read/write.)
Public Function QuickSaveDIBAsPNG(ByRef dstFilename As String, ByRef srcDIB As pdDIB, Optional ByVal forceTo24bppRGB As Boolean = False, Optional ByVal dontCompress As Boolean = False) As Boolean

    'Perform a few failsafe checks
    If (srcDIB Is Nothing) Then
        QuickSaveDIBAsPNG = False
        PDDebug.LogAction "Can't save null PNG!"
        Exit Function
    End If
    
    If (srcDIB.GetDIBWidth = 0) Or (srcDIB.GetDIBHeight = 0) Then
        QuickSaveDIBAsPNG = False
        PDDebug.LogAction "Can't save zero-width/height PNG!"
        Exit Function
    End If
    
    'PD exclusively uses premultiplied alpha for internal DIBs (unless image processing math dictates otherwise).
    ' Saved files always use non-premultiplied alpha.  If the source image is premultiplied, we want to create a
    ' temporary non-premultiplied copy.
    Dim alphaWasChanged As Boolean
    If srcDIB.GetAlphaPremultiplication Then
        srcDIB.SetAlphaPremultiplication False
        alphaWasChanged = True
    End If
    
    'Sometimes compression isn't necessary, which makes this step extremely fast
    Dim compressLevel As Long
    If dontCompress Then compressLevel = 0 Else compressLevel = 3
    
    Dim outColorType As PD_PNGColorType
    If forceTo24bppRGB Then outColorType = png_Truecolor Else outColorType = png_TruecolorAlpha
    
    Dim cPNG As pdPNG
    Set cPNG = New pdPNG
    QuickSaveDIBAsPNG = (cPNG.SavePNG_ToFile(dstFilename, srcDIB, Nothing, outColorType, 8, compressLevel) < png_Failure)
    
    If (Not QuickSaveDIBAsPNG) Then PDDebug.LogAction "Saving.QuickSaveDIBAsPNG failed (pdPNG couldn't write the file?)."
    
    If alphaWasChanged Then srcDIB.SetAlphaPremultiplication True
    
End Function

'PhotoDemon can currently export animated GIF, PNG, and WebP images.  These fromats all have slight subtleties
' in how we prep frames prior to export, but you can call this universal function to handle all those details
' for you.  Note that you *must* pass a correct format ID as the first parameter, and a reference to the pdImage
' object you want saved.
Public Function Export_Animation(ByVal dstFormat As PD_IMAGE_FORMAT, ByRef srcImage As pdImage) As Boolean

    Export_Animation = False
    If (srcImage Is Nothing) Then Exit Function
    
    'Before proceeding, make sure the image has multiple frames.  If it doesn't, we only need to save a static image.
    If (srcImage.GetNumOfLayers <= 1) Then
        If Not PromptSingleFrameSave() Then Exit Function
    End If
    
    'Reuse the user's current "save image" path for the export
    Dim cdInitialFolder As String
    cdInitialFolder = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
    
    'Suggest a default file name.  (At present, we just reuse the current image's name.)
    Dim dstFile As String
    dstFile = srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    If (LenB(dstFile) = 0) Then dstFile = g_Language.TranslateMessage("New image")
    dstFile = cdInitialFolder & dstFile
    
    Dim cdTitle As String
    Select Case dstFormat
        Case PDIF_GIF
            cdTitle = g_Language.TranslateMessage("Export animated GIF")
        Case PDIF_PNG
            cdTitle = g_Language.TranslateMessage("Export animated PNG")
        Case PDIF_WEBP
            cdTitle = g_Language.TranslateMessage("Export animated WebP")
        Case Else
            PDDebug.LogAction "WARNING! Saving.Export_Animation received bad PDIF: " & dstFormat
            Exit Function
    End Select
    
    'Start by prompting the user for an export path
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    Dim saveSuccess As Boolean
    Select Case dstFormat
        Case PDIF_GIF
            saveSuccess = saveDialog.GetSaveFileName(dstFile, , True, "GIF - Graphics Interchange Format (*.gif)|*.gif", , cdInitialFolder, cdTitle, ".gif", FormMain.hWnd)
        Case PDIF_PNG
            saveSuccess = saveDialog.GetSaveFileName(dstFile, , True, "APNG/PNG - Animated Portable Network Graphics (*.apng, *.png)|*.apng;*.png", , cdInitialFolder, cdTitle, ".apng", FormMain.hWnd)
        Case PDIF_WEBP
            saveSuccess = saveDialog.GetSaveFileName(dstFile, , True, "WEBP - Google WebP (*.webp)|*.webp", , cdInitialFolder, cdTitle, ".webp", FormMain.hWnd)
    End Select
    
    If saveSuccess Then
        
        'The user supplied a path.
        
        'Before proceeding with the save, check for some file-level errors that may cause problems.
        
        'If the file already exists, ensure we have write+delete access
        If (Not Files.FileTestAccess_Write(dstFile)) Then
            Message "Warning - file locked: %1", dstFile
            PDMsgBox "Unfortunately, the file '%1' is currently locked by another program on this PC." & vbCrLf & vbCrLf & "Please close this file in any other running programs, then try again.", vbExclamation Or vbOKOnly, "File locked", dstFile
            Export_Animation = False
            Exit Function
        End If
        
        'Update the stored last-save-folder value
        UserPrefs.SetPref_String "Paths", "Save Image", Files.FileGetPath(dstFile)
        UserPrefs.SetPref_Boolean "Saving", "Has Saved A File", True
        
        'Next, retrieve export settings
        ' (Batch processor behavior is currently TBD - but note that the export menu is *not* currently supported
        ' by the batch processor.)
        Dim formatParams As String, metadataParams As String
        Dim promptResult As VbMsgBoxResult
        
        Select Case dstFormat
            Case PDIF_GIF
                promptResult = Dialogs.PromptExportAnimatedGIF(srcImage, formatParams, metadataParams)
            Case PDIF_PNG
                promptResult = Dialogs.PromptExportAnimatedPNG(srcImage, formatParams, metadataParams)
            Case PDIF_WEBP
                promptResult = Dialogs.PromptExportAnimatedWebP(srcImage, formatParams, metadataParams)
        End Select
        
        If (promptResult <> vbOK) Then
            Export_Animation = False
            Exit Function
        End If
        
        'Lock the UI
        Saving.BeginSaveProcess
        
        'Perform the actual save
        Dim saveResult As Boolean
        Select Case dstFormat
            Case PDIF_GIF
                saveResult = ImageExporter.ExportGIF_Animated(srcImage, dstFile, formatParams, metadataParams)
            Case PDIF_PNG
                saveResult = ImageExporter.ExportPNG_Animated(srcImage, dstFile, formatParams, metadataParams)
            Case PDIF_WEBP
                saveResult = ImageExporter.ExportWebP_Animated(srcImage, dstFile, formatParams, metadataParams)
        End Select
        
        If saveResult Then
        
            'If the file was successfully written, we can now embed any additional metadata.
            ' (Note: I don't like embedding metadata in a separate step, but that's a necessary evil of routing all metadata handling
            ' through an external plugin.  Exiftool requires an existant file to be used as a target, and an existant metadata file
            ' to be used as its source.  It cannot operate purely in-memory - but hey, that's why it's asynchronous!)
            If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) And (Not srcImage.ImgMetadata Is Nothing) Then
                
                'Some export formats aren't supported by ExifTool; we don't even attempt to write metadata on such images
                If ImageFormats.IsExifToolRelevant(dstFormat) Then srcImage.ImgMetadata.WriteAllMetadata dstFile, srcImage
                
            End If
            
            'With all save work complete, we can now update various UI bits to reflect the new image.  Note that these changes are
            ' only applied if we are *not* in the midst  of a batch conversion.
            If (Macros.GetMacroStatus <> MacroBATCH) Then
                g_RecentFiles.AddFileToList dstFile, srcImage
                Interface.SyncInterfaceToCurrentImage
                Interface.NotifyImageChanged PDImages.GetActiveImageID()
            End If
            
        End If
        
        'Free the UI
        Saving.EndSaveProcess
        Message "Save complete."
        
        'Animated GIFs do not use our own encoder; instead, they rely on the 3rd-party FreeImage library.
        ' If FreeImage fails, it should provide detailed information on any errors encountered.  Present these
        ' to the user, in hopes that they might find it useful (or least pass it on to me!)
        If (dstFormat = PDIF_GIF) Then
            If (Not saveResult) Then
                If Plugin_FreeImage.FreeImageErrorState Then
                    Dim fiErrorList As String
                    fiErrorList = Plugin_FreeImage.GetFreeImageErrors
                    PDMsgBox "An error occurred when attempting to save this image.  The FreeImage plugin reported the following error details: " & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "In the meantime, please try saving the image to an alternate format.  You can also let the PhotoDemon developers know about this via the Help > Submit Bug Report menu.", vbCritical Or vbOKOnly, "Error", fiErrorList
                Else
                    PDMsgBox "An unspecified error occurred when attempting to save this image.  Please try saving the image to an alternate format." & vbCrLf & vbCrLf & "If the problem persists, please report it to the PhotoDemon developers via photodemon.org/contact", vbCritical Or vbOKOnly, "Error"
                End If
            End If
        End If
        
    Else
        Export_Animation = False
    End If
    
End Function

'Save the current pdImage's list of edits to a standalone 3D lut file.
Public Function SaveColorLookupToFile(ByRef srcImage As pdImage) As Boolean
    
    'Failsafe checks
    If (srcImage Is Nothing) Then Exit Function
    
    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    'Determine an initial folder.  This is easy - just grab the last "3dlut" path from the preferences file.
    Dim initialSaveFolder As String
    initialSaveFolder = UserPrefs.GetLUTPath()
    
    'Build common dialog filter lists
    Dim cdFilter As pdString, cdFilterExtensions As pdString
    Set cdFilter = New pdString
    Set cdFilterExtensions = New pdString
    
    cdFilter.Append "Adobe / IRIDAS (.cube)|*.cube|"
    cdFilterExtensions.Append "cube|"
    cdFilter.Append "Adobe SpeedGrade (.look)|*.look|"
    cdFilterExtensions.Append "look|"
    cdFilter.Append "Autodesk Lustre (.3dl)|*.3dl"
    cdFilterExtensions.Append "3dl"
    
    'TODO: look?  icc?
    
    'Default to 3dl pending further testing (note common-dialog indices are 1-based)
    Dim cdIndex As Long
    cdIndex = 2
    
    'Suggest a file name.  At present, we just reuse the current image's name.
    Dim dstFilename As String
    dstFilename = srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    If (LenB(dstFilename) = 0) Then dstFilename = g_Language.TranslateMessage("Color lookup")
    dstFilename = initialSaveFolder & dstFilename
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Export color lookup")
    
    'Prep a common dialog interface
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    If saveDialog.GetSaveFileName(dstFilename, , True, cdFilter.ToString(), cdIndex, UserPrefs.GetColorProfilePath, cdTitle, cdFilterExtensions.ToString(), GetModalOwner().hWnd) Then
        
        Saving.BeginSaveProcess
        
        'Update preferences
        UserPrefs.SetLUTPath Files.FileGetPath(dstFilename)
        
        'Convert common-dialog index into a human-readable string
        Dim targetLutFormat As String
        Select Case cdIndex
            Case 1
                targetLutFormat = "cube"
            Case 2
                targetLutFormat = "look"
            Case 3
                targetLutFormat = "3dl"
        End Select
        
        'Retrieve an original, unmodified copy of the current layer
        Dim idLayer As Long
        idLayer = PDImages.GetActiveImage.GetActiveLayerID
        
        Dim origDIB As pdDIB
        If (Not PDImages.GetActiveImage.UndoManager.GetOriginalLayer_FromUndo(origDIB, idLayer)) Then
            
            'If no changes have been made to the current image, the above function will return FALSE.
            ' In this case, we can just retrieve the current layer as-is (because it's unmodified).
            Set origDIB = New pdDIB
            origDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveDIB
            
        End If
        
        'Grab a soft link to the active layer
        Dim curDIB As pdDIB
        Set curDIB = PDImages.GetActiveImage.GetActiveDIB
        
        'Ensure DIB sizes match (and resize as necessary)
        If (origDIB.GetDIBWidth <> curDIB.GetDIBWidth) Or (origDIB.GetDIBHeight <> curDIB.GetDIBHeight) Then
            
            'Resize the original DIB to match the current DIB size
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.CreateBlank curDIB.GetDIBWidth, curDIB.GetDIBHeight, 32, 0, 0
            GDI_Plus.GDIPlus_StretchBlt tmpDIB, 0, 0, curDIB.GetDIBWidth, curDIB.GetDIBHeight, origDIB, 0, 0, origDIB.GetDIBWidth, origDIB.GetDIBHeight, 1!, GP_IM_HighQualityBilinear, dstCopyIsOkay:=True
            Set origDIB = tmpDIB
            
        End If
        
        ' TODO: get lut size from user
        Const LUT_MAX_COUNT As Long = 17
        
        'Build a LUT that describes all changes to the current layer (this is the longest part to process)
        Dim cExport As pdLUT3D
        Set cExport = New pdLUT3D
        If cExport.BuildLUTFromTwoDIBs(origDIB, curDIB, LUT_MAX_COUNT, True) Then
            
            Message "Saving file..."
            ProgressBars.SetProgBarVal ProgressBars.GetProgBarMax
            
            'Export said LUT to desired format
            Select Case targetLutFormat
                Case "cube"
                    SaveColorLookupToFile = cExport.SaveLUTToFile_Cube(dstFilename, vbNullString, vbNullString)
                Case "look"
                    SaveColorLookupToFile = cExport.SaveLUTToFile_look(dstFilename, vbNullString, vbNullString)
                Case "3dl"
                    SaveColorLookupToFile = cExport.SaveLUTToFile_3dl(dstFilename, vbNullString, vbNullString)
            End Select
            
            ProgressBars.ReleaseProgressBar
            Message "Save complete."
            
        'Unspecified error?
        Else
            Debug.Print "fail?"
        End If
        
        Saving.EndSaveProcess
        
    End If
    
    'Re-enable user input regardless of save success/fail behavior
    Interface.EnableUserInput
    
End Function

'If the current image only has one frame of animation, we can still save it, but the image (obviously) won't animate.
' Call this function to ask the user if they still want to proceed.
'RETURNS: TRUE if the user still wants to proceed, FALSE if they do not.
Private Function PromptSingleFrameSave() As Boolean

    Dim msgText As pdString
    Set msgText = New pdString
    
    msgText.AppendLine g_Language.TranslateMessage("This is a still image (only one frame of animation).")
    msgText.AppendLineBreak
    msgText.Append g_Language.TranslateMessage("You may proceed, but the image will be saved as a static image, not an animated one.")
    
    Dim msgResult As VbMsgBoxResult
    msgResult = PDMsgBox(msgText.ToString(), vbOKCancel Or vbApplicationModal Or vbExclamation, "Export animation")
    If (msgResult = vbCancel) Then Exit Function
    
End Function

'Some image formats can take a long time to write, especially if the image is large.  As a failsafe, call this function prior to
' initiating a save request.  Just make sure to call the counterpart function when saving completes (or if saving fails); otherwise, the
' main form will be disabled!
Public Sub BeginSaveProcess()
    Processor.MarkProgramBusyState True, True
End Sub

Public Sub EndSaveProcess()
    Processor.MarkProgramBusyState False, True
End Sub

'Want to free up memory?  Call this function to release all export caches.
Public Sub FreeUpMemory()
    Set m_PDIWriter = Nothing
End Sub
