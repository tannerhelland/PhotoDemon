Attribute VB_Name = "Saving"
'***************************************************************************
'File Saving Interface
'Copyright 2001-2018 by Tanner Helland
'Created: 4/15/01
'Last updated: 08/March/16
'Last update: refactor various bits of save-related code to make PD's primary save functions much more versatile.
'
'Module responsible for all image saving, with the exception of the GDI+ image save function (which has been left in
' the GDI+ module for consistency's sake).  Export functions are sorted by file type, and most serve as relatively
' lightweight wrappers corresponding functions in the FreeImage plugin.
'
'The most important sub is PhotoDemon_SaveImage at the top of the module.  This sub is responsible for a multitude of
' decision-making related to saving an image, including tasks like raising format-specific save dialogs, determining
' what color-depth to use, and requesting MRU updates post-save.  Note that the raising of export dialogs can be
' manually controlled by the forceOptionsDialog parameter.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'To improve Undo/Redo performance, a persistent Undo writer is used.  (To free up memory, you can release this class;
' it will automatically be re-created, as necessary.)
Private m_PdiWriter As pdPackager

'When a Save request is invoked, call this function to determine if Save As is needed instead.  (Several factors can
' affect whether Save is okay; for example, if an image has never been saved before, we must raise a dialog to ask
' for a save location and filename.)
Public Function IsCommonDialogRequired(ByRef srcImage As pdImage) As Boolean
    
    'At present, this heuristic is pretty simple: if the image hasn't been saved to disk before, require a Save As instead.
    If Len(srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) = 0 Then
        IsCommonDialogRequired = True
    Else
        IsCommonDialogRequired = False
    End If

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
    
    'Retrieve a string representation as well; settings related to this format may be stored inside the pdImage's settings dictionary
    Dim saveExtension As String
    saveExtension = UCase$(g_ImageFormats.GetExtensionFromPDIF(saveFormat))
    
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
        If g_ImageFormats.IsExportDialogSupported(saveFormat) Then
        
            'If the caller did *not* specifically request a dialog, run some heuristics to see if we need one anyway
            ' (e.g. if this the first time saving a JPEG file, we need to query the user for a Quality value)
            If (Not forceOptionsDialog) Then
            
                'See if the user has already seen this dialog...
                dictEntry = "HasSeenExportDialog" & saveExtension
                needToDisplayDialog = Not srcImage.ImgStorage.GetEntry_Boolean(dictEntry, False)
                
                'If the user has seen a dialog, we'll perform one last failsafe check.  Make sure that the exported format's
                ' parameter string exists; if it doesn't, we need to prompt them again.
                dictEntry = "ExportParams" & saveExtension
                If (Not needToDisplayDialog) And (Len(srcImage.ImgStorage.GetEntry_String(dictEntry, vbNullString)) = 0) Then
                    pdDebug.LogAction "WARNING!  PhotoDemon_SaveImage found an image where HasSeenExportDialog = TRUE, but ExportParams = null.  Fix this!"
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
        If Saving.GetExportParamsFromDialog(srcImage, saveFormat, saveParameters, metadataParameters) Then
            srcImage.ImgStorage.AddEntry "MetadataSettings", metadataParameters
            
        'If the user cancels the dialog, exit immediately
        Else
            Message "Save canceled."
            PhotoDemon_SaveImage = False
            Exit Function
        End If
        
    Else
        dictEntry = "ExportParams" & saveExtension
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
    Message "Saving %1 file...", saveExtension
    
    'If the image is being saved to a layered format (like multipage TIFF), various parts of the export engine may
    ' want to inject useful information into the finished file (e.g. ExifTool can append things like page names).
    ' Mark the outgoing file now.
    MarkMultipageExportStatus srcImage, saveFormat, saveParameters, metadataParameters
    
    'With all save parameters collected, we can offload the rest of the save process to per-format save functions.
    saveSuccessful = Saving.ExportToSpecificFormat(srcImage, dstPath, saveFormat, saveParameters, metadataParameters)
    If saveSuccessful Then
        
        'The file was saved successfully!  Copy the save parameters into the parent pdImage object; subsequent "save" actions
        ' can use these instead of querying the user again.
        dictEntry = "ExportParams" & saveExtension
        srcImage.ImgStorage.AddEntry dictEntry, saveParameters
        
        'If a dialog was displayed, note that as well
        If (needToDisplayDialog) Then
            dictEntry = "HasSeenExportDialog" & saveExtension
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
        If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) And (Not srcImage.ImgMetadata Is Nothing) And (Not (saveFormat = PDIF_PDI)) Then
            srcImage.ImgMetadata.WriteAllMetadata dstPath, srcImage
        End If
        
        'With all save work complete, we can now update various UI bits to reflect the new image.  Note that these changes are
        ' only applied if we are *not* in the midst  of a batch conversion.
        If (Macros.GetMacroStatus <> MacroBATCH) Then
            g_RecentFiles.AddFileToList dstPath, srcImage
            Interface.SyncInterfaceToCurrentImage
            Interface.NotifyImageChanged g_CurrentImage
        End If
        
        'At this point, it's safe to re-enable the main form and restore the default cursor
        EndSaveProcess
        
        Message "Save complete."
    
    'If something went wrong during the save process, the exporter likely provided its own error report.  Attempt to assemble
    ' a meaningful message for the user.
    Else
    
        Message "Save canceled."
        
        'If FreeImage failed, it should have provided detailed information on the problem.  Present it to the user, in hopes that
        ' they might use it to rectify the situation (or least notify us of what went wrong!)
        If Plugin_FreeImage.FreeImageErrorState Then
            
            Dim fiErrorList As String
            fiErrorList = Plugin_FreeImage.GetFreeImageErrors
            
            'Display the error message
            EndSaveProcess
            PDMsgBox "An error occurred when attempting to save this image.  The FreeImage plugin reported the following error details: " & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "In the meantime, please try saving the image to an alternate format.  You can also let the PhotoDemon developers know about this via the Help > Submit Bug Report menu.", vbCritical Or vbOKOnly, "Image save error", fiErrorList
            
        Else
            EndSaveProcess
            PDMsgBox "An unspecified error occurred when attempting to save this image.  Please try saving the image to an alternate format." & vbCrLf & vbCrLf & "If the problem persists, please report it to the PhotoDemon developers via photodemon.org/contact", vbCritical Or vbOKOnly, "Image save error"
        End If
        
    End If
    
    PhotoDemon_SaveImage = saveSuccessful
    
End Function

'This _BatchSave() function is a shortened, accelerated version of the full _SaveImage() function above.
' It should *only* be used during Batch Process operations, where there is no possibility of user interaction.
' Note that the input parameters are different, as the batch processor requires the user to set most export
' settings in advance (since we can't raise export dialogs mid-batch).
Public Function PhotoDemon_BatchSaveImage(ByRef srcImage As pdImage, ByVal dstPath As String, ByVal saveFormat As PD_IMAGE_FORMAT, Optional ByVal saveParameters As String = vbNullString, Optional ByVal metadataParameters As String = vbNullString)
    
    'The important thing to note about this function is that it *requires* the image to be immediately unloaded
    ' after the save operation finishes.  To improve performance, the source pdImage object is not updated against
    ' any changes incurred by the save operation, so that object *will* be "corrupted" after a save operation occurs.
    ' (Note also that things like failed saves cannot raise any modal dialogs, so the only notification of failure
    ' is the return value of this function.)
    Dim saveSuccessful As Boolean: saveSuccessful = False
    
    'As saving can be somewhat lengthy for large images and/or complex formats, lock the UI now.  Note that we *must* call
    ' the "EndSaveProcess" function to release the UI lock.
    'BeginSaveProcess
    'Message "Saving %1 file...", saveExtension
    
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
        If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) And (Not srcImage.ImgMetadata Is Nothing) And (Not (saveFormat = PDIF_PDI)) Then
            
            'Sometimes, PD may process images faster than ExifTool can parse the source file's metadata.
            ' Check for this, and pause until metadata processing catches up.
            If ExifTool.IsMetadataPipeActive Then
                
                pdDebug.LogAction "Pausing batch process so that metadata processing can catch up..."
                
                Do While ExifTool.IsMetadataPipeActive
                    Sleep 50
                    DoEvents
                Loop
                
                pdDebug.LogAction "Metadata processing caught up; proceeding with batch operation..."
                
            End If
            
            srcImage.ImgMetadata.WriteAllMetadata dstPath, srcImage
            
            Do While ExifTool.IsVerificationModeActive
                Sleep 50
                DoEvents
            Loop
            
        End If
        
    End If
    
    PhotoDemon_BatchSaveImage = saveSuccessful
    
End Function

Private Sub MarkMultipageExportStatus(ByRef srcImage As pdImage, ByVal outputPDIF As PD_IMAGE_FORMAT, Optional ByVal saveParameters As String = vbNullString, Optional ByVal metadataParameters As String = vbNullString)
    
    Dim saveIsMultipage As Boolean: saveIsMultipage = False
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString saveParameters
    
    'TIFF is currently the only image format that supports multipage export
    If (outputPDIF = PDIF_TIFF) Then
    
        'The format parameter string contains the multipage indicator, if any.  (Default is to write a single-page TIFF.)
        If cParams.GetBool("TIFFMultipage", False) Then saveIsMultipage = True
        
    End If
    
    'If the outgoing image is multipage, add a special dictionary entry that other functions can easily test.
    srcImage.ImgStorage.AddEntry "MultipageExportActive", saveIsMultipage
    
End Sub

'Given a source image, a desired export format, and a destination string, fill the destination string with format-specific parameters
' returned from the associated format-specific dialog.
'
'Returns: TRUE if dialog was closed via OK button; FALSE otherwise.
Public Function GetExportParamsFromDialog(ByRef srcImage As pdImage, ByVal outputPDIF As PD_IMAGE_FORMAT, ByRef dstParamString As String, ByRef dstMetadataString As String) As Boolean
    
    'As a failsafe, make sure the requested format even *has* an export dialog!
    If g_ImageFormats.IsExportDialogSupported(outputPDIF) Then
        
        Select Case outputPDIF
            
            Case PDIF_BMP
                GetExportParamsFromDialog = (DialogManager.PromptBMPSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_GIF
                GetExportParamsFromDialog = (DialogManager.PromptGIFSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_JP2
                GetExportParamsFromDialog = (DialogManager.PromptJP2Settings(srcImage, dstParamString, dstMetadataString) = vbOK)
                
            Case PDIF_JPEG
                GetExportParamsFromDialog = (DialogManager.PromptJPEGSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
                
            Case PDIF_JXR
                GetExportParamsFromDialog = (DialogManager.PromptJXRSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
        
            Case PDIF_PNG
                GetExportParamsFromDialog = (DialogManager.PromptPNGSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
                
            Case PDIF_PNM
                GetExportParamsFromDialog = (DialogManager.PromptPNMSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
                
            Case PDIF_TIFF
                GetExportParamsFromDialog = (DialogManager.PromptTIFFSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_WEBP
                GetExportParamsFromDialog = (DialogManager.PromptWebPSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
                
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

    'As a convenience, load the current set of parameters into an XML parser; some formats use this data to select an
    ' appropriate export engine (if multiples are available, e.g. both FreeImage and GDI+).
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString saveParameters
    
    Select Case outputPDIF
        
        Case PDIF_BMP
            ExportToSpecificFormat = ImageExporter.ExportBMP(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_GIF
            ExportToSpecificFormat = ImageExporter.ExportGIF(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_HDR
            ExportToSpecificFormat = ImageExporter.ExportHDR(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_JP2
            ExportToSpecificFormat = ImageExporter.ExportJP2(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_JPEG
            ExportToSpecificFormat = ImageExporter.ExportJPEG(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_JXR
            ExportToSpecificFormat = ImageExporter.ExportJXR(srcImage, dstPath, saveParameters, metadataParameters)
        
        'Note: if one or more compression libraries are missing, PDI export is not guaranteed to work.
        Case PDIF_PDI
            ExportToSpecificFormat = SavePhotoDemonImage(srcImage, dstPath, False, PD_CE_Zstd, PD_CE_Zstd, False, True, 5)
                        
        Case PDIF_PNG
            ExportToSpecificFormat = ImageExporter.ExportPNG(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_PNM
            ExportToSpecificFormat = ImageExporter.ExportPNM(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_PSD
            ExportToSpecificFormat = ImageExporter.ExportPSD(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_TARGA
            ExportToSpecificFormat = ImageExporter.ExportTGA(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_TIFF
            ExportToSpecificFormat = ImageExporter.ExportTIFF(srcImage, dstPath, saveParameters, metadataParameters)
        
        Case PDIF_WEBP
            ExportToSpecificFormat = ImageExporter.ExportWebP(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case Else
            Message "Output format not recognized.  Save aborted.  Please use the Help -> Submit Bug Report menu item to report this incident."
            ExportToSpecificFormat = False
            
    End Select

End Function

'Save the current image to PhotoDemon's native PDI format
' TODO:
'  - Add support for storing a PNG copy of the fully composited image, preferably in the data chunk of the first node.
'  - Any number of other options might be helpful (e.g. password encryption, etc).  I should probably add a page about the PDI
'    format to the help documentation, where various ideas for future additions could be tracked.
Public Function SavePhotoDemonImage(ByRef srcPDImage As pdImage, ByVal pdiPath As String, Optional ByVal suppressMessages As Boolean = False, Optional ByVal compressHeaders As PD_CompressionEngine = PD_CE_Zstd, Optional ByVal compressLayers As PD_CompressionEngine = PD_CE_Zstd, Optional ByVal writeHeaderOnlyFile As Boolean = False, Optional ByVal includeMetadata As Boolean = False, Optional ByVal compressionLevel As Long = -1, Optional ByVal secondPassDirectoryCompression As PD_CompressionEngine = PD_CE_NoCompression, Optional ByVal srcIsUndo As Boolean = False) As Boolean
    
    On Error GoTo SavePDIError
    
    'Perform a few failsafe checks
    If (srcPDImage Is Nothing) Then Exit Function
    If (LenB(pdiPath) = 0) Then Exit Function
    
    'Want to time this function?  Here's your chance:
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    Dim sFileType As String
    sFileType = "PDI"
    
    If (Not suppressMessages) Then Message "Saving %1 image...", sFileType
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of compressing individual layers,
    ' and storing everything to a running byte stream.
    Dim pdiWriter As pdPackager
    Set pdiWriter = New pdPackager
    
    'When creating the actual package, we specify numOfLayers + 1 nodes.  The +1 is for the pdImage header itself, which
    ' gets its own node, separate from the individual layer nodes.
    pdiWriter.PrepareNewPackage srcPDImage.GetNumOfLayers + 1, PD_IMAGE_IDENTIFIER, srcPDImage.EstimateRAMUsage, PD_SM_FileBacked, pdiPath
        
    'The first node we'll add is the pdImage header, in XML format.
    Dim nodeIndex As Long
    nodeIndex = pdiWriter.AddNode("pdImage Header", -1, 0)
    
    Dim dataString As String
    srcPDImage.WriteExternalData dataString, True
    
    pdiWriter.AddNodeDataFromString nodeIndex, True, dataString, compressHeaders
    
    'The pdImage header only requires one of the two buffers in its node; the other can be happily left blank.
    
    'Next, we will add each pdLayer object to the stream.  This is done in two steps:
    ' 1) First, obtain the layer header in XML format and write it out
    ' 2) Second, obtain any layer-specific data (DIB for raster layers, XML for vector layers) and write it out
    Dim layerXMLHeader As String, layerXMLData As String
    Dim layerDIBPointer As Long, layerDIBLength As Long
    
    Dim i As Long
    For i = 0 To srcPDImage.GetNumOfLayers - 1
    
        'Create a new node for this layer.  Note that the index is stored directly in the node name ("pdLayer (n)")
        ' while the layerID is stored as the nodeID.
        nodeIndex = pdiWriter.AddNode("pdLayer " & i, srcPDImage.GetLayerByIndex(i).GetLayerID, 1)
        
        'Retrieve the layer header and add it to the header section of this node.
        ' (Note: compression level of text data, like layer headers, is not controlled by the user.  For short strings like
        '        these headers, there is no meaningful gain from higher compression settings, but higher settings kills
        '        performance, so we stick with the default recommended zLib compression level.)
        layerXMLHeader = srcPDImage.GetLayerByIndex(i).GetLayerHeaderAsXML(True)
        pdiWriter.AddNodeDataFromString nodeIndex, True, layerXMLHeader, compressHeaders
        
        'If this is not a header-only file, retrieve any layer-type-specific data and add it to the data section of this node
        ' (Note: the user's compression setting *is* used for this data section, as it can be quite large for raster layers
        '        as we have to store a raw stream of the DIB contents.)
        If (Not writeHeaderOnlyFile) Then
        
            'Specific handling varies by layer type
            
            'Image layers save their raster contents as a raw byte stream
            If srcPDImage.GetLayerByIndex(i).IsLayerRaster Then
                
                'Debug.Print "Writing layer index " & i & " out to file as RASTER layer."
                srcPDImage.GetLayerByIndex(i).layerDIB.RetrieveDIBPointerAndSize layerDIBPointer, layerDIBLength
                pdiWriter.AddNodeDataFromPointer nodeIndex, False, layerDIBPointer, layerDIBLength, compressLayers, compressionLevel
                
            'Text (and other vector layers) save their vector contents in XML format
            ElseIf srcPDImage.GetLayerByIndex(i).IsLayerVector Then
                
                'Debug.Print "Writing layer index " & i & " out to file as VECTOR layer."
                layerXMLData = srcPDImage.GetLayerByIndex(i).GetVectorDataAsXML(True)
                pdiWriter.AddNodeDataFromString nodeIndex, False, layerXMLData, compressLayers, compressionLevel
            
            'No other layer types are currently supported
            Else
                Debug.Print "WARNING!  SavePhotoDemonImage can't save the layer at index " & i
                
            End If
            
        End If
    
    Next i
    
    'Next, if the "write metadata" flag has been set, and the image has metadata, add a metadata entry to the file.
    If includeMetadata And (Not srcPDImage.ImgMetadata Is Nothing) Then
        
        Dim mdStartTime As Currency
        VBHacks.GetHighResTime mdStartTime
        
        If srcPDImage.ImgMetadata.HasMetadata Then
            
            'To avoid unnecessary string copies, we write the (potentially large) original metadata string directly
            ' from its source pointer.
            Dim mdPtr As Long, mdLen As Long
            srcPDImage.ImgMetadata.GetOriginalXMLMetadataStrPtrAndLen mdPtr, mdLen
            
            If (mdLen > 0) Then
            
                nodeIndex = pdiWriter.AddNode("pdMetadata_Raw", -1, 2)
                pdiWriter.AddNodeDataFromPointer nodeIndex, True, mdPtr, mdLen * 2, compressHeaders
                'pdiWriter.AddNodeDataFromString nodeIndex, False, srcPDImage.ImgMetadata.GetOriginalXMLMetadataString(), compressHeaders
                'Unfortunately, there's no good way to do this for our already-parsed metadata collection...
                pdiWriter.AddNodeDataFromString nodeIndex, False, srcPDImage.ImgMetadata.GetSerializedXMLData(), compressHeaders
                
                pdDebug.LogAction "Note: metadata writes took " & VBHacks.GetTimeDiffNowAsString(mdStartTime)
                
            Else
                Debug.Print "FYI, metadata string data is reported as zero-length; abandoning write"
            End If
            
        End If
        
    End If
    
    'That's all there is to it!  Write the completed pdPackage out to file.
    SavePhotoDemonImage = pdiWriter.WritePackageToFile(pdiPath, secondPassDirectoryCompression, srcIsUndo)
    
    'Report timing on debug builds
    If SavePhotoDemonImage Then
        pdDebug.LogAction "Saved PDI file in " & CStr(VBHacks.GetTimerDifferenceNow(startTime) * 1000) & " ms."
    Else
        pdDebug.LogAction "WARNING!  SavePhotoDemonImage failed after " & CStr(VBHacks.GetTimerDifferenceNow(startTime) * 1000) & " ms."
    End If
    
    If (Not suppressMessages) Then Message "%1 save complete.", sFileType
    
    Exit Function
    
SavePDIError:

    SavePhotoDemonImage = False
    
End Function

'Save the requested layer to a variant of PhotoDemon's native PDI format.  Because this function is internal (it is used by the
' Undo/Redo engine only), it is not as fleshed-out as the actual SavePhotoDemonImage function.
Public Function SavePhotoDemonLayer(ByRef srcLayer As pdLayer, ByVal pdiPath As String, Optional ByVal suppressMessages As Boolean = False, Optional ByVal compressHeaders As PD_CompressionEngine = PD_CE_Zstd, Optional ByVal compressLayers As PD_CompressionEngine = PD_CE_Zstd, Optional ByVal writeHeaderOnlyFile As Boolean = False, Optional ByVal compressionLevel As Long = -1, Optional ByVal srcIsUndo As Boolean = False) As Boolean
    
    On Error GoTo SavePDLayerError
    
    'Perform a few failsafe checks
    If (srcLayer Is Nothing) Then Exit Function
    If (srcLayer.layerDIB Is Nothing) Then Exit Function
    If (Len(pdiPath) = 0) Then Exit Function
    
    Dim sFileType As String
    sFileType = "PDI"
    
    If (Not suppressMessages) Then Message "Saving %1 layer...", sFileType
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of assembling the layer file.
    ' (Note that we reuse the same pdPackager instance throughout this module; this spares us from constant reallocating
    ' memory for package assembly.)
    If (m_PdiWriter Is Nothing) Then Set m_PdiWriter = New pdPackager
    
    'Unlike an actual PDI file, which stores a whole bunch of images, these temp layer files only have two pieces of data:
    ' the layer header, and the DIB bytestream.  Thus, we know there will only be 1 node required.
    ' (NOTE: if memory consumption ever becomes an issue, you can write Undo/Redo data as a "file-backed" package.  In that
    '  mode, data is streamed immediately out to file, instead of coalescing the package in memory first.  This makes it
    '  quite a bit slower, but the memory burden is significantly reduced.)
    m_PdiWriter.PrepareNewPackage 1, PD_LAYER_IDENTIFIER, srcLayer.EstimateRAMUsage   ', PD_SM_FileBacked, pdiPath
    
    'The first (and only) node we'll add is the specific pdLayer header and DIB data.
    ' To help us reconstruct the node later, we also note the current layer's ID (stored as the node ID)
    '  and the current layer's index (stored as the node type).
    
    'Start by creating the node entry; if successful, this will return the index of the node, which we can use
    ' to supply the actual header and DIB data.
    Dim nodeIndex As Long
    nodeIndex = m_PdiWriter.AddNode("pdLayer", srcLayer.GetLayerID, pdImages(g_CurrentImage).GetLayerIndexFromID(srcLayer.GetLayerID))
    
    'Retrieve the layer header (in XML format), then write the XML stream to the pdPackage instance
    Dim dataString As String
    dataString = srcLayer.GetLayerHeaderAsXML(True)
    
    m_PdiWriter.AddNodeDataFromString nodeIndex, True, dataString, compressHeaders
    
    'If this is not a header-only request, retrieve the layer DIB (as a byte array), then copy the array
    ' into the pdPackage instance
    If (Not writeHeaderOnlyFile) Then
        
        'Specific handling varies by layer type
        
        'Image layers save their raster contents as a raw byte stream
        If srcLayer.IsLayerRaster Then
        
            Dim layerDIBPointer As Long, layerDIBLength As Long
            srcLayer.layerDIB.RetrieveDIBPointerAndSize layerDIBPointer, layerDIBLength
            m_PdiWriter.AddNodeDataFromPointer nodeIndex, False, layerDIBPointer, layerDIBLength, compressLayers, compressionLevel
        
        'Text (and other vector layers) save their vector contents in XML format
        ElseIf srcLayer.IsLayerVector Then
            
            dataString = srcLayer.GetVectorDataAsXML(True)
            m_PdiWriter.AddNodeDataFromString nodeIndex, False, dataString, compressLayers, compressionLevel
        
        'Other layer types are not currently supported
        Else
            Debug.Print "WARNING!  SavePhotoDemonLayer was passed a layer of unknown or unsupported type."
        End If
        
    End If
    
    'That's all there is to it!  Write the completed pdPackage out to file.
    SavePhotoDemonLayer = m_PdiWriter.WritePackageToFile(pdiPath, , srcIsUndo)
    If (Not SavePhotoDemonLayer) Then
        pdDebug.LogAction "WARNING!  SavingSavePhotoDemonLayer received a failure status from pdiWriter.WritePackageToFile()"
    End If
    
    Exit Function
    
SavePDLayerError:
    pdDebug.LogAction "WARNING!  Saving.SavePhotoDemonLayer failed with error #" & Err.Number & ", " & Err.Description
    SavePhotoDemonLayer = False
    
End Function

'Save a new Undo/Redo entry to file.  This function is only called by the createUndoData function in the pdUndo class.
' For the most part, this function simply wraps other save functions; however, certain odd types of Undo diff files (e.g. layer headers)
' may be directly processed and saved by this function.
'
'Note that this function interacts closely with the matching LoadUndo function in the Loading module.  Any novel Undo diff types added
' here must also be mirrored there.
Public Function SaveUndoData(ByRef srcPDImage As pdImage, ByRef dstUndoFilename As String, ByVal processType As PD_UndoType, Optional ByVal targetLayerID As Long = -1, Optional ByVal compressionHint As Long = -1) As Boolean
    
    Dim timeAtUndoStart As Currency
    VBHacks.GetHighResTime timeAtUndoStart
    
    'As of v7.0, PD has multiple compression engines available.  These engines are not exposed to the user.  We use LZ4 by default,
    ' as it is far and away the fastest at both compression and decompression (while compressing at marginally worse ratios).
    ' Note that if the user selects increasingly better compression results, we will silently switch to zstd instead.
    Dim undoCmpEngine As PD_CompressionEngine, undoCmpLevel As Long
    If (g_UndoCompressionLevel = 0) Then
        undoCmpEngine = PD_CE_NoCompression
        undoCmpLevel = 0
    
    'At level 1 (the current PD default), use LZ4 compression at default strength.  (Remember that LZ4's compression level do not
    ' improve as the level goes up - the algorithm's *performance* improves as the level goes up.)
    ElseIf (g_UndoCompressionLevel = 1) Then
        undoCmpEngine = PD_CE_Lz4
        undoCmpLevel = compressionHint
    
    'For all higher levels, use zstd, and reset the compression level to start at 1 (so a g_UndoCompressionLevel of 2 uses zstd at
    ' its default compression strength of level 1).
    Else
        undoCmpEngine = PD_CE_Zstd
        undoCmpLevel = g_UndoCompressionLevel - 1
    End If
    
    Dim undoSuccess As Boolean
    
    'What kind of Undo data we save is determined by the current processType.
    Select Case processType
    
        'EVERYTHING, meaning a full copy of the pdImage stack and any selection data
        Case UNDO_Everything
            undoSuccess = Saving.SavePhotoDemonImage(srcPDImage, dstUndoFilename, True, PD_CE_Lz4, undoCmpEngine, False, True, undoCmpLevel, , True)
            srcPDImage.MainSelection.WriteSelectionToFile dstUndoFilename & ".selection", undoCmpEngine, undoCmpLevel, undoCmpEngine, undoCmpLevel
            
        'A full copy of the pdImage stack
        Case UNDO_Image, UNDO_Image_VectorSafe
            undoSuccess = Saving.SavePhotoDemonImage(srcPDImage, dstUndoFilename, True, PD_CE_Lz4, undoCmpEngine, False, True, undoCmpLevel, , True)
        
        'A full copy of the pdImage stack, *without any layer DIB data*
        Case UNDO_ImageHeader
            undoSuccess = Saving.SavePhotoDemonImage(srcPDImage, dstUndoFilename, True, undoCmpEngine, PD_CE_NoCompression, True, True, undoCmpLevel, , True)
        
        'Layer data only (full layer header + full layer DIB).
        Case UNDO_Layer, UNDO_Layer_VectorSafe
            undoSuccess = Saving.SavePhotoDemonLayer(srcPDImage.GetLayerByID(targetLayerID), dstUndoFilename & ".layer", True, PD_CE_Lz4, undoCmpEngine, False, undoCmpLevel, True)
        
        'Layer header data only (e.g. DO NOT WRITE OUT THE LAYER DIB)
        Case UNDO_LayerHeader
            undoSuccess = Saving.SavePhotoDemonLayer(srcPDImage.GetLayerByID(targetLayerID), dstUndoFilename & ".layer", True, undoCmpEngine, PD_CE_NoCompression, True, undoCmpLevel, True)
            
        'Selection data only
        Case UNDO_Selection
            undoSuccess = srcPDImage.MainSelection.WriteSelectionToFile(dstUndoFilename & ".selection", undoCmpEngine, undoCmpLevel, undoCmpEngine, undoCmpLevel)
            
        'Anything else (this should never happen, but good to have a failsafe)
        Case Else
            pdDebug.LogAction "Unknown Undo data write requested - is it possible to avoid this request entirely??"
            undoSuccess = Saving.SavePhotoDemonImage(srcPDImage, dstUndoFilename, True, PD_CE_Lz4, undoCmpEngine, False, , undoCmpLevel, , True)
        
    End Select
    
    SaveUndoData = undoSuccess
    
    If (Not SaveUndoData) Then pdDebug.LogAction "SaveUndoData returned failure; cause unknown."
    'Want to test undo timing?  Uncomment the line below
    'pdDebug.LogAction "Undo file creation took: " & Format$(VBHacks.GetTimerDifferenceNow(timeAtUndoStart) * 1000, "####0.00") & " ms"
    
End Function

'Quickly save a DIB to file in PNG format.  At present, this is only used when forwarding image data
' to the Windows Photo Printer object.  (All internal quick-saves use PD-specific formats, which are
' much faster to read/write.)
Public Function QuickSaveDIBAsPNG(ByVal dstFilename As String, ByRef srcDIB As pdDIB) As Boolean

    'Perform a few failsafe checks
    If (srcDIB Is Nothing) Then
        QuickSaveDIBAsPNG = False
        Exit Function
    End If
    
    If (srcDIB.GetDIBWidth = 0) Or (srcDIB.GetDIBHeight = 0) Then
        QuickSaveDIBAsPNG = False
        Exit Function
    End If

    'If FreeImage is available, use it to save the PNG; otherwise, fall back to GDI+
    If g_ImageFormats.FreeImageEnabled Then
        
        'PD exclusively uses premultiplied alpha for internal DIBs (unless image processing math dictates otherwise).
        ' Saved files always use non-premultiplied alpha.  If the source image is premultiplied, we want to create a
        ' temporary non-premultiplied copy.
        Dim alphaWasChanged As Boolean
        If srcDIB.GetAlphaPremultiplication Then
            srcDIB.SetAlphaPremultiplication False
            alphaWasChanged = True
        End If
        
        'Convert the temporary DIB to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(srcDIB.GetDIBDC)
    
        'Use that handle to save the image to PNG format
        If (fi_DIB <> 0) Then
            Dim fi_Check As Long
            
            'Output the PNG file at the proper color depth
            Dim fi_OutputColorDepth As FREE_IMAGE_COLOR_DEPTH
            If (srcDIB.GetDIBColorDepth = 24) Then fi_OutputColorDepth = FICD_24BPP Else fi_OutputColorDepth = FICD_32BPP
            
            'Ask FreeImage to write the thumbnail out to file
            fi_Check = FreeImage_SaveEx(fi_DIB, dstFilename, PDIF_PNG, FISO_PNG_Z_BEST_SPEED, fi_OutputColorDepth, , , , , True)
            If (Not fi_Check) Then pdDebug.LogAction "Saving.QuickSaveDIBAsPNG via FreeImage failed (FreeImage_SaveEx silent fail)."
            
        Else
            pdDebug.LogAction "Saving.QuickSaveDIBAsPNG via FreeImage failed (blank handle)."
        End If
        
        If alphaWasChanged Then srcDIB.SetAlphaPremultiplication True
        
    'FreeImage is not available; try to use GDI+ to save a PNG thumbnail
    Else
        If (Not GDIPlusQuickSavePNG(dstFilename, srcDIB)) Then pdDebug.LogAction "Saving.QuickSaveDIBAsPNG via GDI+ failed (unspecified GDI+ error)."
    End If

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
    Set m_PdiWriter = Nothing
End Sub
