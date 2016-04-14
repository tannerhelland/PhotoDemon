Attribute VB_Name = "Saving"
'***************************************************************************
'File Saving Interface
'Copyright 2001-2016 by Tanner Helland
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

'When a Save request is invoked, call this function to determine if Save As is needed instead.  (Several factors can
' affect whether Save is okay; for example, if an image has never been saved before, we must raise a dialog to ask
' for a save location and filename.)
Public Function IsCommonDialogRequired(ByRef srcImage As pdImage) As Boolean
    
    'At present, this heuristic is pretty simple: if the image hasn't been saved to disk before, require a Save As instead.
    If Len(srcImage.imgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) = 0 Then
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
'   3) Optional: whether to force display of an "additional save options" dialog (JPEG quality, etc)
'   4) Optional: a string of relevant save parameters.  This is only used during batch conversion
Public Function PhotoDemon_SaveImage(ByRef srcImage As pdImage, ByVal dstPath As String, Optional ByVal forceOptionsDialog As Boolean = False) As Boolean
    
    'There are a few different ways the save process can "fail":
    ' 1) a save dialog with extra options is required, and the user cancels it
    ' 2) file-system errors (folder not writable, not enough free space, etc)
    ' 3) save engine errors (e.g. FreeImage explodes mid-save)
    
    'These have varying degrees of severity, but I mention this in advance because a number of post-save behaviors (like updating
    ' the Recent Files list) are abandoned under *any* of these occurrences.  As such, a lot of this function postpones various
    ' tasks until after all possible failure states have been dealt with.
    Dim saveSuccessful As Boolean: saveSuccessful = False
    
    'The caller must tell us which format they want us to use.  This value is stored in the .currentFileFormat property of the pdImage object.
    Dim saveFormat As PHOTODEMON_IMAGE_FORMAT
    saveFormat = srcImage.currentFileFormat
    
    'Retrieve a string representation as well; settings related to this format may be stored inside the pdImage's settings dictionary
    Dim saveExtension As String
    saveExtension = UCase$(g_ImageFormats.GetExtensionFromPDIF(saveFormat))
    
    Dim dictEntry As String
    
    'The first major task this function deals with is save prompts.  The formula for showing these is hierarchical:
    
    ' 0) SPECIAL STEP: if we are in the midst of a batch process, *never* display a dialog.
    ' 1) If the caller has forcibly requested an options dialog (as "Save As" does), display a dialog.
    ' 2) If the caller hasn't forcibly requested a dialog...
        '3) See if this output format even supports dialogs.  If it doesn't, proceed with saving.
        '4) If this output format does support a dialog...
            '5) If the user has already seen a dialog for this format, don't show one again
            '6) If the user hasn't already seen a dialog for this format, it's time to show them one!
    
    'We'll deal with each of these in turn.
    Dim needToDisplayDialog As Boolean: needToDisplayDialog = forceOptionsDialog
    
    'Make sure we're not in the midst of a batch process operation
    If (MacroStatus <> MacroBATCH) Then
        
        'See if this format even supports dialogs...
        If g_ImageFormats.IsExportDialogSupported(saveFormat) Then
        
            'If the caller did *not* specifically request a dialog, run some heuristics to see if we need one anyway
            ' (e.g. if this the first time saving a JPEG file, we need to query the user for a Quality value)
            If (Not forceOptionsDialog) Then
            
                'See if the user has already seen this dialog...
                dictEntry = "HasSeenExportDialog" & saveExtension
                needToDisplayDialog = Not srcImage.imgStorage.GetEntry_Boolean(dictEntry, False)
                
                'If the user has seen a dialog, we'll perform one last failsafe check.  Make sure that the exported format's
                ' parameter string exists; if it doesn't, we need to prompt them again.
                dictEntry = "ExportParams" & saveExtension
                If (Not needToDisplayDialog) And (Len(srcImage.imgStorage.GetEntry_String(dictEntry, vbNullString)) = 0) Then
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "WARNING!  PhotoDemon_SaveImage found an image where HasSeenExportDialog = TRUE, but ExportParams = null.  Fix this!"
                    #End If
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
            srcImage.imgStorage.AddEntry "MetadataSettings", metadataParameters
        
        'If the user cancels the dialog, exit immediately
        Else
            Message "Save canceled."
            PhotoDemon_SaveImage = False
            Exit Function
        End If
        
    Else
        dictEntry = "ExportParams" & saveExtension
        saveParameters = srcImage.imgStorage.GetEntry_String(dictEntry, vbNullString)
        metadataParameters = srcImage.imgStorage.GetEntry_String("MetadataSettings", vbNullString)
    End If
    
    'As saving can be somewhat lengthy for large images and/or complex formats, lock the UI now.  Note that we *must* call
    ' the "EndSaveProcess" function to release the UI lock.
    BeginSaveProcess
    Message "Saving %1 file...", saveExtension
    
    'With all save parameters collected, we can offload the rest of the save process to per-format save functions.
    saveSuccessful = ExportToSpecificFormat(srcImage, dstPath, saveFormat, saveParameters, metadataParameters)
    If saveSuccessful Then
        
        'The file was saved successfully!  Copy the save parameters into the parent pdImage object; subsequent "save" actions
        ' can use these instead of querying the user again.
        dictEntry = "ExportParams" & saveExtension
        srcImage.imgStorage.AddEntry dictEntry, saveParameters
        
        'Similarly, remember the file's location and selected name for future saves
        Dim cFile As pdFSO
        Set cFile = New pdFSO
        srcImage.imgStorage.AddEntry "CurrentLocationOnDisk", dstPath
        srcImage.imgStorage.AddEntry "OriginalFileName", cFile.GetFilename(dstPath, True)
        srcImage.imgStorage.AddEntry "OriginalFileExtension", cFile.GetFileExtension(dstPath)
        
        'Update the parent image's save state.
        If saveFormat = PDIF_PDI Then srcImage.SetSaveState True, pdSE_SavePDI Else srcImage.SetSaveState True, pdSE_SaveFlat
        
        'If the file was successfully written, we can now embed any additional metadata.
        ' (Note: I don't like embedding metadata in a separate step, but that's a necessary evil of routing all metadata handling
        ' through an external plugin.  Exiftool requires an existant file to be used as a target, and an existant metadata file
        ' to be used as its source.  It cannot operate purely in-memory - but hey, that's why it's asynchronous!)
        If g_ExifToolEnabled Then
            
            'Only attempt to export metadata if ExifTool was able to successfully cache and parse metadata prior to saving
            If Not (srcImage.imgMetadata Is Nothing) Then
                If srcImage.imgMetadata.HasMetadata Then
                    srcImage.imgMetadata.WriteAllMetadata dstPath, srcImage
                Else
                    Message "No metadata to export.  Continuing save..."
                End If
            End If
            
        End If
        
        'With all save work complete, we can now update various UI bits to reflect the new image.  Note that these changes are
        ' only applied if we are *not* in the midst  of a batch conversion.
        If (MacroStatus <> MacroBATCH) Then
            g_RecentFiles.MRU_AddNewFile dstPath, srcImage
            SyncInterfaceToCurrentImage
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
            PDMsgBox "An error occurred when attempting to save this image.  The FreeImage plugin reported the following error details: " & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "In the meantime, please try saving the image to an alternate format.  You can also let the PhotoDemon developers know about this via the Help > Submit Bug Report menu.", vbCritical Or vbApplicationModal Or vbOKOnly, "Image save error", fiErrorList
            
        Else
            EndSaveProcess
            PDMsgBox "An unspecified error occurred when attempting to save this image.  Please try saving the image to an alternate format." & vbCrLf & vbCrLf & "If the problem persists, please report it to the PhotoDemon developers via photodemon.org/contact", vbCritical Or vbApplicationModal Or vbOKOnly, "Image save error"
        End If
        
    End If
    
    PhotoDemon_SaveImage = saveSuccessful
    
End Function

'Given a source image, a desired export format, and a destination string, fill the destination string with format-specific parameters
' returned from the associated format-specific dialog.
'
'Returns: TRUE if dialog was closed via OK button; FALSE otherwise.
Public Function GetExportParamsFromDialog(ByRef srcImage As pdImage, ByVal outputPDIF As PHOTODEMON_IMAGE_FORMAT, ByRef dstParamString As String, ByRef dstMetadataString As String) As Boolean
    
    'As a failsafe, make sure the requested format even *has* an export dialog!
    If g_ImageFormats.IsExportDialogSupported(outputPDIF) Then
        
        Select Case outputPDIF
            
            Case PDIF_BMP
                GetExportParamsFromDialog = CBool(Dialog_Handler.PromptBMPSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_GIF
                GetExportParamsFromDialog = CBool(Dialog_Handler.PromptGIFSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
            
            Case PDIF_JPEG
                GetExportParamsFromDialog = CBool(Dialog_Handler.PromptJPEGSettings(srcImage, dstParamString, dstMetadataString) = vbOK)
                
            Case PDIF_JP2
                GetExportParamsFromDialog = CBool(Dialog_Handler.PromptJP2Settings(srcImage, dstParamString) = vbOK)
                
            Case PDIF_WEBP
                GetExportParamsFromDialog = CBool(Dialog_Handler.PromptWebPSettings(srcImage, dstParamString) = vbOK)
                
            Case PDIF_JXR
                GetExportParamsFromDialog = CBool(Dialog_Handler.PromptJXRSettings(srcImage, dstParamString) = vbOK)
        
        End Select
        
    Else
        GetExportParamsFromDialog = False
        dstParamString = vbNullString
    End If
        
End Function

'Already have a save parameter string assembled?  Call this function to export directly to a given format, with no UI prompts.
' (I *DO NOT* recommend calling this function directly.  PD only uses it from within the main _SaveImage function, which also applies
'  a number of failsafe checks against things like path accessibility and format compatibility.)
Private Function ExportToSpecificFormat(ByRef srcImage As pdImage, ByRef dstPath As String, ByVal outputPDIF As PHOTODEMON_IMAGE_FORMAT, Optional ByVal saveParameters As String = vbNullString, Optional ByVal metadataParameters As String = vbNullString) As Boolean

    'As a convenience, load the current set of parameters into an XML parser; some formats use this data to select an
    ' appropriate export engine (if multiples are available, e.g. both FreeImage and GDI+).
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString saveParameters
    
    Select Case outputPDIF
        
        Case PDIF_BMP
            ExportToSpecificFormat = ImageExporter.ExportBMP(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_JPEG
            ExportToSpecificFormat = ImageExporter.ExportJPEG(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_PDI
            If g_ZLibEnabled Then
                ExportToSpecificFormat = SavePhotoDemonImage(srcImage, dstPath, , , , , , True)
            Else
                ExportToSpecificFormat = False
            End If
        
        'GIFs are preferentially exported by FreeImage, then GDI+ (if available).  I don't know how to control the algorithm
        ' GDI+ uses for 8-bpp color reduction, so the results of its encoder are likely to be poor.
        Case PDIF_GIF
            ExportToSpecificFormat = ImageExporter.ExportGIF(srcImage, dstPath, saveParameters, metadataParameters)
            
        Case PDIF_PNG
            If g_ImageFormats.FreeImageEnabled Then
                ExportToSpecificFormat = SavePNGImage(srcImage, dstPath, , saveParameters)
            ElseIf g_ImageFormats.GDIPlusEnabled Then
                ExportToSpecificFormat = GDIPlusSavePicture(srcImage, dstPath, ImagePNG, 32)
            Else
                ExportToSpecificFormat = False
            End If
            
        Case PDIF_PPM
            ExportToSpecificFormat = SavePPMImage(srcImage, dstPath, saveParameters)
            
        Case PDIF_TARGA
            ExportToSpecificFormat = SaveTGAImage(srcImage, dstPath, , saveParameters)
            
        Case PDIF_JP2
            ExportToSpecificFormat = SaveJP2Image(srcImage, dstPath, , saveParameters)
            
        'TIFFs are preferentially exported by FreeImage, then GDI+ (if available)
        Case PDIF_TIFF
            If g_ImageFormats.FreeImageEnabled Then
                ExportToSpecificFormat = SaveTIFImage(srcImage, dstPath, , saveParameters)
            ElseIf g_ImageFormats.GDIPlusEnabled Then
                ExportToSpecificFormat = GDIPlusSavePicture(srcImage, dstPath, ImageTIFF, 32)
            Else
                ExportToSpecificFormat = False
            End If
        
        Case PDIF_WEBP
            ExportToSpecificFormat = SaveWebPImage(srcImage, dstPath, , saveParameters)
        
        Case PDIF_JXR
            ExportToSpecificFormat = SaveJXRImage(srcImage, dstPath, , saveParameters)
            
        Case PDIF_HDR
            ExportToSpecificFormat = ImageExporter.ExportHDR(srcImage, dstPath, saveParameters)
        
        Case PDIF_PSD
            ExportToSpecificFormat = ImageExporter.ExportPSD(srcImage, dstPath, saveParameters)
        
        Case Else
            Message "Output format not recognized.  Save aborted.  Please use the Help -> Submit Bug Report menu item to report this incident."
            ExportToSpecificFormat = False
            
    End Select

End Function


'Save the current image to PhotoDemon's native PDI format
' TODO:
'  - Add support for storing a PNG copy of the fully composited image, preferably in the data chunk of the first node.
'  - Figure out a good way to store metadata; the problem is not so much storing the metadata itself, but storing any user edits.
'    I have postponed this until I get metadata editing working more fully.  (NOTE: metadata is now stored correctly, but the
'    user edit aspect remains to be dealt with.)
'  - User-settable options for compression.  Some users may prefer extremely tight compression, at a trade-off of slower
'    image saves.  Similarly, compressing layers in PNG format instead of as a blind zLib stream would probably yield better
'    results (at a trade-off to performance).  (NOTE: these features are now supported by the function, but they are not currently
'    exposed to the user.)
'  - Any number of other options might be helpful (e.g. password encryption, etc).  I should probably add a page about the PDI
'    format to the help documentation, where various ideas for future additions could be tracked.
Public Function SavePhotoDemonImage(ByRef srcPDImage As pdImage, ByVal PDIPath As String, Optional ByVal suppressMessages As Boolean = False, Optional ByVal compressHeaders As Boolean = True, Optional ByVal compressLayers As Boolean = True, Optional ByVal embedChecksums As Boolean = True, Optional ByVal writeHeaderOnlyFile As Boolean = False, Optional ByVal WriteMetadata As Boolean = False, Optional ByVal compressionLevel As Long = -1, Optional ByVal secondPassDirectoryCompression As Boolean = False, Optional ByVal secondPassDataCompression As Boolean = False, Optional ByVal srcIsUndo As Boolean = False) As Boolean
    
    On Error GoTo SavePDIError
    
    'Perform a few failsafe checks
    If srcPDImage Is Nothing Then Exit Function
    If Len(PDIPath) = 0 Then Exit Function
    
    Dim sFileType As String
    sFileType = "PDI"
    
    If Not suppressMessages Then Message "Saving %1 image...", sFileType
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of compressing individual layers,
    ' and storing everything to a running byte stream.
    Dim pdiWriter As pdPackager
    Set pdiWriter = New pdPackager
    pdiWriter.init_ZLib "", True, g_ZLibEnabled
    
    'When creating the actual package, we specify numOfLayers + 1 nodes.  The +1 is for the pdImage header itself, which
    ' gets its own node, separate from the individual layer nodes.
    pdiWriter.prepareNewPackage srcPDImage.GetNumOfLayers + 1, PD_IMAGE_IDENTIFIER, srcPDImage.estimateRAMUsage
        
    'The first node we'll add is the pdImage header, in XML format.
    Dim nodeIndex As Long
    nodeIndex = pdiWriter.addNode("pdImage Header", -1, 0)
    
    Dim dataString As String
    srcPDImage.WriteExternalData dataString, True
    
    pdiWriter.addNodeDataFromString nodeIndex, True, dataString, compressHeaders, , embedChecksums
    
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
        nodeIndex = pdiWriter.addNode("pdLayer " & i, srcPDImage.GetLayerByIndex(i).getLayerID, 1)
        
        'Retrieve the layer header and add it to the header section of this node.
        ' (Note: compression level of text data, like layer headers, is not controlled by the user.  For short strings like
        '        these headers, there is no meaningful gain from higher compression settings, but higher settings kills
        '        performance, so we stick with the default recommended zLib compression level.)
        layerXMLHeader = srcPDImage.GetLayerByIndex(i).getLayerHeaderAsXML(True)
        pdiWriter.addNodeDataFromString nodeIndex, True, layerXMLHeader, compressHeaders, , embedChecksums
        
        'If this is not a header-only file, retrieve any layer-type-specific data and add it to the data section of this node
        ' (Note: the user's compression setting *is* used for this data section, as it can be quite large for raster layers
        '        as we have to store a raw stream of the DIB contents.)
        If Not writeHeaderOnlyFile Then
        
            'Specific handling varies by layer type
            
            'Image layers save their raster contents as a raw byte stream
            If srcPDImage.GetLayerByIndex(i).isLayerRaster Then
                
                Debug.Print "Writing layer index " & i & " out to file as RASTER layer."
                srcPDImage.GetLayerByIndex(i).layerDIB.retrieveDIBPointerAndSize layerDIBPointer, layerDIBLength
                pdiWriter.addNodeDataFromPointer nodeIndex, False, layerDIBPointer, layerDIBLength, compressLayers, compressionLevel, embedChecksums
                
            'Text (and other vector layers) save their vector contents in XML format
            ElseIf srcPDImage.GetLayerByIndex(i).isLayerVector Then
                
                Debug.Print "Writing layer index " & i & " out to file as VECTOR layer."
                layerXMLData = srcPDImage.GetLayerByIndex(i).getVectorDataAsXML(True)
                pdiWriter.addNodeDataFromString nodeIndex, False, layerXMLData, compressLayers, compressionLevel, embedChecksums
            
            'No other layer types are currently supported
            Else
                Debug.Print "WARNING!  SavePhotoDemonImage can't save the layer at index " & i
                
            End If
            
        End If
    
    Next i
    
    'Next, if the "write metadata" flag has been set, and the image has metadata, add a metadata entry to the file.
    If (Not writeHeaderOnlyFile) And WriteMetadata And Not (srcPDImage.imgMetadata Is Nothing) Then
    
        If srcPDImage.imgMetadata.HasMetadata Then
            nodeIndex = pdiWriter.addNode("pdMetadata_Raw", -1, 2)
            pdiWriter.addNodeDataFromString nodeIndex, True, srcPDImage.imgMetadata.GetOriginalXMLMetadataString, compressHeaders, , embedChecksums
            pdiWriter.addNodeDataFromString nodeIndex, False, srcPDImage.imgMetadata.GetSerializedXMLData, compressHeaders, , embedChecksums
        End If
    
    End If
    
    'That's all there is to it!  Write the completed pdPackage out to file.
    SavePhotoDemonImage = pdiWriter.writePackageToFile(PDIPath, secondPassDirectoryCompression, secondPassDataCompression, srcIsUndo)
    
    If Not suppressMessages Then Message "%1 save complete.", sFileType
    
    Exit Function
    
SavePDIError:

    SavePhotoDemonImage = False
    
End Function

'Save the requested layer to a variant of PhotoDemon's native PDI format.  Because this function is internal (it is used by the
' Undo/Redo engine only), it is not as fleshed-out as the actual SavePhotoDemonImage function.
Public Function SavePhotoDemonLayer(ByRef srcLayer As pdLayer, ByVal PDIPath As String, Optional ByVal suppressMessages As Boolean = False, Optional ByVal compressHeaders As Boolean = True, Optional ByVal compressLayers As Boolean = True, Optional ByVal embedChecksums As Boolean = True, Optional ByVal writeHeaderOnlyFile As Boolean = False, Optional ByVal compressionLevel As Long = -1, Optional ByVal srcIsUndo As Boolean = False) As Boolean
    
    On Error GoTo SavePDLayerError
    
    'Perform a few failsafe checks
    If srcLayer Is Nothing Then Exit Function
    If srcLayer.layerDIB Is Nothing Then Exit Function
    If Len(PDIPath) = 0 Then Exit Function
    
    Dim sFileType As String
    sFileType = "PDI"
    
    If Not suppressMessages Then Message "Saving %1 layer...", sFileType
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of assembling the layer file.
    Dim pdiWriter As pdPackager
    Set pdiWriter = New pdPackager
    pdiWriter.init_ZLib "", True, g_ZLibEnabled
    
    'Unlike an actual PDI file, which stores a whole bunch of images, these temp layer files only have two pieces of data:
    ' the layer header, and the DIB bytestream.  Thus, we know there will only be 1 node required.
    pdiWriter.prepareNewPackage 1, PD_LAYER_IDENTIFIER, srcLayer.estimateRAMUsage
        
    'The first (and only) node we'll add is the specific pdLayer header and DIB data.
    ' To help us reconstruct the node later, we also note the current layer's ID (stored as the node ID)
    '  and the current layer's index (stored as the node type).
    
    'Start by creating the node entry; if successful, this will return the index of the node, which we can use
    ' to supply the actual header and DIB data.
    Dim nodeIndex As Long
    nodeIndex = pdiWriter.addNode("pdLayer", srcLayer.getLayerID, pdImages(g_CurrentImage).GetLayerIndexFromID(srcLayer.getLayerID))
    
    'Retrieve the layer header (in XML format), then write the XML stream to the pdPackage instance
    Dim dataString As String
    dataString = srcLayer.getLayerHeaderAsXML(True)
    
    pdiWriter.addNodeDataFromString nodeIndex, True, dataString, compressHeaders, , embedChecksums
    
    'If this is not a header-only request, retrieve the layer DIB (as a byte array), then copy the array
    ' into the pdPackage instance
    If Not writeHeaderOnlyFile Then
        
        'Specific handling varies by layer type
        
        'Image layers save their raster contents as a raw byte stream
        If srcLayer.isLayerRaster Then
        
            Dim layerDIBPointer As Long, layerDIBLength As Long
            srcLayer.layerDIB.retrieveDIBPointerAndSize layerDIBPointer, layerDIBLength
            pdiWriter.addNodeDataFromPointer nodeIndex, False, layerDIBPointer, layerDIBLength, compressLayers, compressionLevel, embedChecksums
        
        'Text (and other vector layers) save their vector contents in XML format
        ElseIf srcLayer.isLayerVector Then
            
            dataString = srcLayer.getVectorDataAsXML(True)
            pdiWriter.addNodeDataFromString nodeIndex, False, dataString, compressLayers, compressionLevel, embedChecksums
        
        'Other layer types are not currently supported
        Else
            Debug.Print "WARNING!  SavePhotoDemonLayer was passed a layer of unknown or unsupported type."
        End If
        
    End If
    
    'That's all there is to it!  Write the completed pdPackage out to file.
    SavePhotoDemonLayer = pdiWriter.writePackageToFile(PDIPath, , , srcIsUndo)
    
    If Not suppressMessages Then Message "%1 save complete.", sFileType
    
    Exit Function
    
SavePDLayerError:

    SavePhotoDemonLayer = False
    
End Function

'Save a PNG (Portable Network Graphic) file.  GDI+ can also do this.  Note that this function is enormous and quite complicated,
' owing to the many interactions between color spaces, bit-depth, custom PNG features (like compression quality), and the
' availability of plugins like PNGQuant that further compress saved PNG files.
Public Function SavePNGImage(ByRef srcPDImage As pdImage, ByVal PNGPath As String, Optional ByVal outputColorDepth As Long = 32, Optional ByVal pngParams As String = "") As Boolean

    On Error GoTo SavePNGError

    'Parse all possible PNG parameters
    ' (At present, three are possible: compression level, interlacing, BKGD chunk preservation (background color)
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(pngParams) <> 0 Then cParams.SetParamString pngParams
    Dim pngCompressionLevel As Long
    pngCompressionLevel = cParams.GetLong(1, 9)
    Dim pngUseInterlacing As Boolean, pngPreserveBKGD As Boolean
    pngUseInterlacing = cParams.GetBool(2, False)
    pngPreserveBKGD = cParams.GetBool(3, False)

    Dim sFileType As String
    sFileType = "PNG"

    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SavePNGImage = False
        Exit Function
    End If
    
    'Before doing anything else, make a special note of the outputColorDepth.  If it is 8bpp, we will use PNGQuant to help with the save.
    Dim output8BPP As Boolean
    If outputColorDepth <= 8 Then output8BPP = True Else output8BPP = False
        
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.GetCompositedImage tmpDIB, False
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth <= 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'PhotoDemon provides PNGQuant plugin support.  PNGQuant can render extremely high-quality 8bpp PNG files
        ' with "full" transparency.  If the pngquant executable is available, the export process of 8bpp PNGs is
        ' a bit different.
        
        'Before we can send stuff off to PNGQuant, however, we need to see if the image has more than 256 colors.
        ' If it doesn't, we can save the file without PNGQuant's help.
        
        'Check to see if the current image had its colors counted before coming here.  If not, count it.
        Dim numColors As Long
        If g_LastImageScanned <> srcPDImage.imageID Then
            numColors = GetQuickColorCount(tmpDIB, srcPDImage.imageID)
        Else
            numColors = g_LastColorCount
        End If
        
        'PNGQuant can handle all types of transparency for us, but if it doesn't exist, we must rely on our own routines.
        If (Not g_ImageFormats.pngQuantEnabled) Then
        
            'Does this DIB contain binary transparency?  If so, mark all transparent pixels with magic magenta.
            If DIB_Handler.IsDIBAlphaBinary(tmpDIB) Then
                tmpDIB.ApplyAlphaCutoff
            Else
            
                'If we are in the midst of a batch conversion, we don't want to bother the user with alpha dialogs.
                ' Thus, use a default cut-off of 127 and continue on.
                If (MacroStatus = MacroBATCH) Then
                    tmpDIB.ApplyAlphaCutoff
                
                'We're not in a batch conversion, so ask the user which cut-off they would like to use.
                Else
            
                    Dim alphaCheck As VbMsgBoxResult
                    alphaCheck = PromptAlphaCutoff(tmpDIB)
                    
                    'If the alpha dialog is canceled, abandon the entire save
                    If alphaCheck = vbCancel Then
                    
                        tmpDIB.eraseDIB
                        Set tmpDIB = Nothing
                        SavePNGImage = False
                        Exit Function
                    
                    'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                    Else
                        tmpDIB.ApplyAlphaCutoff g_AlphaCutoff, , g_AlphaCompositeColor
                    End If
                
                End If
                
            End If
            
        'If PNGQuant is available, force the output to 32bpp.  PNGQuant will take care of the actual 32bpp -> 8bpp reduction.
        Else
            outputColorDepth = 32
        End If
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth < 32) Then tmpDIB.CompositeBackgroundColor 255, 255, 255
    
        'Also, if PNGquant is enabled, use it for the transformation - and note that we need to reset the
        ' first PNG save (pre-PNGQuant) color depth to 24bpp
        If (tmpDIB.getDIBColorDepth = 24) And (outputColorDepth = 8) And g_ImageFormats.pngQuantEnabled Then outputColorDepth = 24
    
    End If
    
    Message "Writing %1 file...", sFileType
    
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    'If the image is being reduced from some higher bit-depth to 1bpp, manually force a conversion with dithering
    If outputColorDepth = 1 Then fi_DIB = FreeImage_Dither(fi_DIB, FID_FS)
    
    'If the image contains alpha and PNGquant is not available, we need to manually convert the FreeImage copy of the image to 8bpp.
    ' Then we need to apply alpha using the cut-off established earlier in this section.
    If handleAlpha And (Not g_ImageFormats.pngQuantEnabled) Then
    
        fi_DIB = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_WUQUANT, True)
        
        'We now need to find the palette index of a known transparent pixel
        Dim transpX As Long, transpY As Long
        tmpDIB.getTransparentLocation transpX, transpY
        
        Dim palIndex As Byte
        FreeImage_GetPixelIndex fi_DIB, transpX, transpY, palIndex
        
        'Request that FreeImage set that palette entry as the transparent index
        FreeImage_SetTransparentIndex fi_DIB, palIndex
        
        'Finally, because some software may not display the transparency correctly, we need to set that
        ' palette index to its original value.  To do that, we must make a copy of the palette and update
        ' the transparency index accordingly.
        Dim fi_Palette() As Long
        fi_Palette = FreeImage_GetPaletteExLong(fi_DIB)
        
        fi_Palette(palIndex) = tmpDIB.GetOriginalTransparentColor()
        
    End If
        
    'Use that handle to save the image to PNG format
    If fi_DIB <> 0 Then
    
        'Embed a background color if available, and the user has requested it.
        If pngPreserveBKGD And srcPDImage.imgStorage.DoesKeyExist("pngBackgroundColor") Then
            
            Dim rQuad As RGBQUAD
            rQuad.Red = ExtractR(srcPDImage.imgStorage.GetEntry_Long("pngBackgroundColor"))
            rQuad.Green = ExtractG(srcPDImage.imgStorage.GetEntry_Long("pngBackgroundColor"))
            rQuad.Blue = ExtractB(srcPDImage.imgStorage.GetEntry_Long("pngBackgroundColor"))
            FreeImage_SetBackgroundColor fi_DIB, rQuad
        
        End If
    
        'Finally, prepare some PNG save flags based on the parameters we were passed
        Dim PNGFlags As Long
        
        'Compression level (1 to 9, but FreeImage also has a "no compression" option with a unique flag)
        PNGFlags = pngCompressionLevel
        If PNGFlags = 0 Then PNGFlags = PNG_Z_NO_COMPRESSION
        
        'Interlacing
        If pngUseInterlacing Then PNGFlags = (PNGFlags Or PNG_INTERLACED)
    
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, PNGPath, PDIF_PNG, PNGFlags, outputColorDepth, , , , , True)
        
        If Not fi_Check Then
            
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            SavePNGImage = False
            Exit Function
        
        'Save was successful.  If PNGQuant is being used to help with the 8bpp reduction, activate it now.
        Else
            
            If g_ImageFormats.pngQuantEnabled And output8BPP Then
            
                'Build a full shell path for the pngquant operation
                Dim shellPath As String
                shellPath = g_PluginPath & "pngquant.exe "
                
                'Force overwrite if a file with that name already exists
                shellPath = shellPath & "-f "
                
                'Display verbose status messages (consider removing this for production builds)
                shellPath = shellPath & "-v "
                
                'Request the addition of a custom "-8bpp.png" extension; without this, PNGquant will use its own extension
                ' (-fs8.png or -or8.png, depending on the use of dithering)
                shellPath = shellPath & "--ext -8bpp.png "
                
                'Now, add options that the user may have specified.
                
                'Dithering override
                If tmpDIB.getDIBColorDepth = 32 Then
                    If Not g_UserPreferences.GetPref_Boolean("Plugins", "PNGQuant Dithering", True) Then
                        shellPath = shellPath & "--nofs "
                    End If
                End If
        
                'Improved IE6 compatibility
                If g_UserPreferences.GetPref_Boolean("Plugins", "PNGQuant IE6 Compatibility", False) Then
                    shellPath = shellPath & "--iebug "
                End If
        
                'Performance
                shellPath = shellPath & "--speed " & g_UserPreferences.GetPref_Long("Plugins", "PNGQuant Performance", 3) & " "
                
                'Append the name of the current image
                shellPath = shellPath & """" & PNGPath & """"
                
                'Use PNGQuant to create a new file
                Message "Using the PNGQuant plugin to write a high-quality 8bpp PNG file.  This may take a moment..."
                
                'Before launching the shell, launch a single DoEvents.  This gives us some leeway before Windows marks the program
                ' as unresponsive...
                DoEvents
                
                Dim shellCheck As Boolean
                shellCheck = ShellAndWait(shellPath, vbMinimizedNoFocus)
            
                'If the shell was successful and the image was created successfully, overwrite the original 32bpp save
                ' (from FreeImage) with the new 8bpp one (from PNGQuant)
                If shellCheck Then
                
                    Message "PNGQuant transformation complete.  Verifying output..."
                
                    'If successful, PNGQuant created a new file with the name "filename-8bpp.png".  We need to rename that file
                    ' to whatever name the user originally supplied - but only if the 8bpp transformation was successful!
                    Dim srcFile As String
                    srcFile = PNGPath
                    StripOffExtension srcFile
                    srcFile = srcFile & "-8bpp.png"
                    
                    'Make sure both FreeImage and PNGQuant were able to generate valid files, then rewrite the FreeImage one
                    ' with the PNGQuant one.
                    Dim cFile As pdFSO
                    Set cFile = New pdFSO
                    
                    If cFile.FileExist(srcFile) And cFile.FileExist(PNGPath) Then
                        cFile.KillFile PNGPath
                        cFile.CopyFile srcFile, PNGPath
                        cFile.KillFile srcFile
                    Else
                        Message "PNGQuant could not write file.  Saving 32bpp image instead..."
                    End If
                    
                Else
                    Message "PNGQuant could not write file.  Saving 32bpp image instead..."
                End If
            
            End If
            
            Message "%1 save complete.", sFileType
            
        End If
        
    Else
        
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SavePNGImage = False
        Exit Function
        
    End If
    
    SavePNGImage = True
    Exit Function
    
SavePNGError:

    SavePNGImage = False
    
End Function

'Save a PPM (Portable Pixmap) image
Public Function SavePPMImage(ByRef srcPDImage As pdImage, ByVal PPMPath As String, Optional ByVal ppmParams As String = "") As Boolean

    On Error GoTo SavePPMError

    'Parse all possible PPM parameters (at present there is only one possible parameter, which sets RAW vs ASCII encoding)
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(ppmParams) <> 0 Then cParams.SetParamString ppmParams
    Dim ppmFormat As Long
    ppmFormat = cParams.GetLong(1, 0)

    Dim sFileType As String
    sFileType = "PPM"

    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SavePPMImage = False
        Exit Function
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'Based on the user's preference, select binary or ASCII encoding for the PPM file
    Dim ppm_Encoding As FREE_IMAGE_SAVE_OPTIONS
    If ppmFormat = 0 Then ppm_Encoding = FISO_PNM_SAVE_RAW Else ppm_Encoding = FISO_PNM_SAVE_ASCII
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.GetCompositedImage tmpDIB, False
    
    'PPM only supports 24bpp
    If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.convertTo24bpp
        
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
        
    'Use that handle to save the image to PPM format (ASCII)
    If fi_DIB <> 0 Then
    
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, PPMPath, PDIF_PPM, ppm_Encoding, FICD_24BPP, , , , , True)
        
        If fi_Check Then
            Message "%1 save complete.", sFileType
        Else
        
            Message "PPM save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
            SavePPMImage = False
            Exit Function
            
        End If
        
    Else
    
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SavePPMImage = False
        Exit Function
        
    End If
    
    SavePPMImage = True
    Exit Function
        
SavePPMError:

    SavePPMImage = False
        
End Function

'Save to Targa (TGA) format.
Public Function SaveTGAImage(ByRef srcPDImage As pdImage, ByVal TGAPath As String, Optional ByVal outputColorDepth As Long = 32, Optional ByVal tgaParams As String = "") As Boolean
    
    On Error GoTo SaveTGAError
    
    'Parse all possible TGA parameters (at present there is only one possible parameter, which specifies RLE compression)
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(tgaParams) <> 0 Then cParams.SetParamString tgaParams
    Dim TGACompression As Boolean
    TGACompression = cParams.GetBool(1, False)
    
    Dim sFileType  As String
    sFileType = "TGA"
    
    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveTGAImage = False
        Exit Function
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.GetCompositedImage tmpDIB, False
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'Does this DIB contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If DIB_Handler.IsDIBAlphaBinary(tmpDIB) Then
            tmpDIB.ApplyAlphaCutoff
        Else
        
            'If we are in the midst of a batch conversion, we don't want to bother the user with alpha dialogs.
            ' Thus, use a default cut-off of 127 and continue on.
            If MacroStatus = MacroBATCH Then
                tmpDIB.ApplyAlphaCutoff
            
            'We're not in a batch conversion, so ask the user which cut-off they would like to use.
            Else
            
                Dim alphaCheck As VbMsgBoxResult
                alphaCheck = PromptAlphaCutoff(tmpDIB)
                
                'If the alpha dialog is canceled, abandon the entire save
                If alphaCheck = vbCancel Then
                
                    tmpDIB.eraseDIB
                    Set tmpDIB = Nothing
                    SaveTGAImage = False
                    Exit Function
                
                'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                Else
                    tmpDIB.ApplyAlphaCutoff g_AlphaCutoff, , g_AlphaCompositeColor
                End If
            End If
            
        End If
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth < 32) Then tmpDIB.CompositeBackgroundColor 255, 255, 255
    
    End If
    
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    'If the image contains alpha, we need to convert the FreeImage copy of the image to 8bpp
    If handleAlpha Then
    
        fi_DIB = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_NNQUANT, True)
        
        'We now need to find the palette index of a known transparent pixel
        Dim transpX As Long, transpY As Long
        tmpDIB.getTransparentLocation transpX, transpY
        
        Dim palIndex As Byte
        FreeImage_GetPixelIndex fi_DIB, transpX, transpY, palIndex
        
        'Request that FreeImage set that palette entry as the transparent index
        FreeImage_SetTransparentIndex fi_DIB, palIndex
        
        'Finally, because some software may not display the transparency correctly, we need to set that
        ' palette index to some normal color instead of bright magenta.  To do that, we must make a
        ' copy of the palette and update the transparency index accordingly.
        Dim fi_Palette() As Long
        fi_Palette = FreeImage_GetPaletteExLong(fi_DIB)
        
        fi_Palette(palIndex) = tmpDIB.GetOriginalTransparentColor()
        
    End If
    
    'Finally, prepare a TGA save flag.  If the user has requested RLE encoding, pass that along to FreeImage.
    Dim TGAflags As Long
    TGAflags = TARGA_DEFAULT
            
    If TGACompression Then TGAflags = TARGA_SAVE_RLE
    
    'Use that handle to save the image to TGA format
    If fi_DIB <> 0 Then
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TGAPath, PDIF_TARGA, TGAflags, outputColorDepth, , , , , True)
        
        If fi_Check Then
            Message "%1 save complete.", sFileType
        Else
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            SaveTGAImage = False
            Exit Function
        End If
        
    Else
    
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SaveTGAImage = False
        Exit Function
        
    End If
    
    SaveTGAImage = True
    Exit Function
    
SaveTGAError:

    SaveTGAImage = False

End Function

'Save a TIFF (Tagged Image File Format) image via FreeImage.  GDI+ can also do this.
Public Function SaveTIFImage(ByRef srcPDImage As pdImage, ByVal TIFPath As String, Optional ByVal outputColorDepth As Long = 32, Optional ByVal tiffParams As String = "") As Boolean
    
    On Error GoTo SaveTIFError
    
    'Parse all possible TIFF parameters
    ' (At present, two are possible: one for compression type, and another for CMYK encoding)
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(tiffParams) <> 0 Then cParams.SetParamString tiffParams
    Dim tiffEncoding As Long
    tiffEncoding = cParams.GetLong(1, 0)
    Dim tiffUseCMYK As Boolean
    tiffUseCMYK = cParams.GetBool(2, False)
    
    Dim sFileType As String
    sFileType = "TIFF"
    
    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        
        PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveTIFImage = False
        Exit Function
        
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'TIFFs have some unique considerations regarding compression techniques.  If a color-depth-specific compression
    ' technique has been requested, modify the output depth accordingly.
    Select Case g_UserPreferences.GetPref_Long("File Formats", "TIFF Compression", 0)
        
        'JPEG compression
        Case 6
            outputColorDepth = 24
        
        'CCITT Group 3
        Case 7
            outputColorDepth = 1
        
        'CCITT Group 4
        Case 8
            outputColorDepth = 1
            
    End Select
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.GetCompositedImage tmpDIB, False
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'Does this DIB contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If DIB_Handler.IsDIBAlphaBinary(tmpDIB) Then
            tmpDIB.ApplyAlphaCutoff
        Else
            
            'If we are in the midst of a batch conversion, we don't want to bother the user with alpha dialogs.
            ' Thus, use a default cut-off of 127 and continue on.
            If MacroStatus = MacroBATCH Then
                tmpDIB.ApplyAlphaCutoff
            
            'We're not in a batch conversion, so ask the user which cut-off they would like to use.
            Else
                Dim alphaCheck As VbMsgBoxResult
                alphaCheck = PromptAlphaCutoff(tmpDIB)
                
                'If the alpha dialog is canceled, abandon the entire save
                If alphaCheck = vbCancel Then
                
                    tmpDIB.eraseDIB
                    Set tmpDIB = Nothing
                    SaveTIFImage = False
                    Exit Function
                
                'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                Else
                    tmpDIB.ApplyAlphaCutoff g_AlphaCutoff, , g_AlphaCompositeColor
                End If
            End If
            
        End If
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth < 32) Then tmpDIB.CompositeBackgroundColor 255, 255, 255
    
    End If
    
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    'If the image is being reduced from some higher bit-depth to 1bpp, manually force a conversion with dithering
    If outputColorDepth = 1 Then fi_DIB = FreeImage_Dither(fi_DIB, FID_FS)
    
    'If the image contains alpha, we need to convert the FreeImage copy of the image to 8bpp
    If handleAlpha Then
    
        fi_DIB = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_NNQUANT, True)
        
        'We now need to find the palette index of a known transparent pixel
        Dim transpX As Long, transpY As Long
        tmpDIB.getTransparentLocation transpX, transpY
        
        Dim palIndex As Byte
        FreeImage_GetPixelIndex fi_DIB, transpX, transpY, palIndex
        
        'Request that FreeImage set that palette entry as the transparent index
        FreeImage_SetTransparentIndex fi_DIB, palIndex
        
        'Finally, because some software may not display the transparency correctly, we need to set that
        ' palette index to some normal color instead of bright magenta.  To do that, we must make a
        ' copy of the palette and update the transparency index accordingly.
        Dim fi_Palette() As Long
        fi_Palette = FreeImage_GetPaletteExLong(fi_DIB)
        
        fi_Palette(palIndex) = tmpDIB.GetOriginalTransparentColor()
        
    End If
    
    'Use that handle to save the image to TIFF format
    If fi_DIB <> 0 Then
        
        'Prepare TIFF export flags based on the user's preferences
        Dim TIFFFlags As Long
        
        Select Case tiffEncoding
        
            'Default settings (LZW for > 1bpp, CCITT Group 4 fax encoding for 1bpp)
            Case 0
                TIFFFlags = TIFF_DEFAULT
                
            'No compression
            Case 1
                TIFFFlags = TIFF_NONE
            
            'Macintosh Packbits (RLE)
            Case 2
                TIFFFlags = TIFF_PACKBITS
            
            'Proper deflate (Adobe-style)
            Case 3
                TIFFFlags = TIFF_ADOBE_DEFLATE
            
            'Obsolete deflate (PKZIP or zLib-style)
            Case 4
                TIFFFlags = TIFF_DEFLATE
            
            'LZW
            Case 5
                TIFFFlags = TIFF_LZW
                
            'JPEG
            Case 6
                TIFFFlags = TIFF_JPEG
            
            'Fax Group 3
            Case 7
                TIFFFlags = TIFF_CCITTFAX3
            
            'Fax Group 4
            Case 8
                TIFFFlags = TIFF_CCITTFAX4
                
        End Select
        
        'If the user has requested CMYK encoding of TIFF files, add that flag and convert the image to 32bpp CMYK
        If (outputColorDepth = 24) And tiffUseCMYK Then
        
            outputColorDepth = 32
            TIFFFlags = (TIFFFlags Or TIFF_CMYK)
            FreeImage_UnloadEx fi_DIB
            
            Dim tmpCMYKDIB As pdDIB
            Set tmpCMYKDIB = New pdDIB
            
            DIB_Handler.createCMYKDIB tmpDIB, tmpCMYKDIB
            fi_DIB = FreeImage_CreateFromDC(tmpCMYKDIB.getDIBDC)
            
            'Release our temporary DIB
            Set tmpCMYKDIB = Nothing
            
        End If
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TIFPath, PDIF_TIFF, TIFFFlags, outputColorDepth, , , , , True)
        
        If fi_Check Then
            Message "%1 save complete.", sFileType
        Else
            
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            SaveTIFImage = False
            Exit Function
            
        End If
        
    Else
    
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SaveTIFImage = False
        Exit Function
        
    End If
    
    SaveTIFImage = True
    Exit Function
    
SaveTIFError:

    SaveTIFImage = False
        
End Function

'Save to JPEG-2000 format using the FreeImage library.
Public Function SaveJP2Image(ByRef srcPDImage As pdImage, ByVal jp2Path As String, Optional ByVal outputColorDepth As Long = 32, Optional ByVal jp2Params As String = "") As Boolean
    
    On Error GoTo SaveJP2Error
    
    'Parse all possible JPEG-2000 params
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString jp2Params
    
    Dim JP2Quality As Long
    JP2Quality = cParams.GetLong("JP2Quality", 1)
    
    Dim sFileType As String
    sFileType = "JPEG-2000"
    
    'Make sure we found the plug-in when we loaded the program
    If (Not g_ImageFormats.FreeImageEnabled) Then
        PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveJP2Image = False
        Exit Function
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.GetCompositedImage tmpDIB, False
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (tmpDIB.getDIBColorDepth = 32) Then tmpDIB.convertTo24bpp
    
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    'Use that handle to save the image to JPEG format
    If fi_DIB <> 0 Then
                
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, jp2Path, PDIF_JP2, JP2Quality, outputColorDepth, , , , , True)
        
        If fi_Check Then
            Message "%1 save complete.", sFileType
        Else
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            SaveJP2Image = False
            Exit Function
        End If
        
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SaveJP2Image = False
        Exit Function
    End If
    
    SaveJP2Image = True
    Exit Function
    
SaveJP2Error:

    SaveJP2Image = False
    
End Function

'Save to JPEG XR format using the FreeImage library.
Public Function SaveJXRImage(ByRef srcPDImage As pdImage, ByVal jxrPath As String, Optional ByVal outputColorDepth As Long = 32, Optional ByVal jxrParams As String = "") As Boolean
    
    On Error GoTo SaveJXRError
    
    'Parse all possible JXR params
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString jxrParams
    
    Dim jxrFlags As Long
    jxrFlags = cParams.GetLong("JXRQuality", 0)
    If cParams.GetBool("JXRProgressive", False) Then jxrFlags = jxrFlags Or JXR_PROGRESSIVE
    
    Dim sFileType As String
    sFileType = "JPEG XR"
    
    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveJXRImage = False
        Exit Function
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.GetCompositedImage tmpDIB, False
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (tmpDIB.getDIBColorDepth = 32) Then tmpDIB.convertTo24bpp
    
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    'Use that handle to save the image to JPEG XR format
    If fi_DIB <> 0 Then
                
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, jxrPath, PDIF_JXR, jxrFlags, outputColorDepth, , , , , True)
        
        If fi_Check Then
            Message "%1 save complete.", sFileType
        Else
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            SaveJXRImage = False
            Exit Function
        End If
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SaveJXRImage = False
        Exit Function
    End If
    
    SaveJXRImage = True
    Exit Function
    
SaveJXRError:
    SaveJXRImage = False
    
End Function

'Save to WebP format using the FreeImage library.
Public Function SaveWebPImage(ByRef srcPDImage As pdImage, ByVal WebPPath As String, Optional ByVal outputColorDepth As Long = 32, Optional ByVal WebPParams As String = "") As Boolean
    
    On Error GoTo SaveWebPError
    
    'Parse all possible WebP params
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString WebPParams
    
    Dim WebPQuality As Long
    WebPQuality = cParams.GetLong("WebPQuality", 0)
    
    Dim sFileType As String
    sFileType = "WebP"
    
    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveWebPImage = False
        Exit Function
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.GetCompositedImage tmpDIB, False
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (tmpDIB.getDIBColorDepth = 32) Then tmpDIB.convertTo24bpp
    
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    'Use that handle to save the image to WebP format
    If fi_DIB <> 0 Then
                
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, WebPPath, PDIF_WEBP, WebPQuality, outputColorDepth, , , , , True)
        
        If fi_Check Then
            Message "%1 save complete.", sFileType
        Else
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            SaveWebPImage = False
            Exit Function
        End If
        
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SaveWebPImage = False
        Exit Function
    End If
    
    SaveWebPImage = True
    Exit Function
    
SaveWebPError:
    SaveWebPImage = False
    
End Function

'Given a source and destination DIB reference, fill the destination with a post-JPEG-compression of the original.  This
' is used to generate the live preview used in PhotoDemon's "export JPEG" dialog.
Public Sub FillDIBWithJPEGVersion(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal jpegQuality As Long, Optional ByVal jpegSubsample As Long = JPEG_SUBSAMPLING_422)

    'srcDIB may be 32bpp.  Convert it to 24bpp if necessary.
    If srcDIB.getDIBColorDepth = 32 Then srcDIB.convertTo24bpp

    'Pass the DIB to FreeImage, which will make a copy for itself.
    Dim fi_DIB As Long
    fi_DIB = Plugin_FreeImage.GetFIHandleFromPDDib_NoCopy(srcDIB)
    
    'Prepare matching flags for FreeImage's JPEG encoder
    Dim jpegFlags As Long
    jpegFlags = jpegQuality Or jpegSubsample
        
    'Now comes the actual JPEG conversion, which is handled exclusively by FreeImage.  Basically, we ask it to save
    ' the image in JPEG format to a byte array; we then hand that byte array back to it and request a decompression.
    Dim jpegArray() As Byte
    Dim fi_Check As Long
    fi_Check = FreeImage_SaveToMemoryEx(PDIF_JPEG, fi_DIB, jpegArray, jpegFlags, True)
    
    fi_DIB = FreeImage_LoadFromMemoryEx(jpegArray, FILO_JPEG_FAST)
    
    'Copy the newly decompressed JPEG into the destination pdDIB object.
    Plugin_FreeImage.PaintFIDibToPDDib dstDIB, fi_DIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight
    
    'Release the FreeImage copy of the DIB
    FreeImage_Unload fi_DIB
    
End Sub

'Given a source image and a desired JPEG perception quality, test various JPEG quality values until an ideal one is found
Public Function FindQualityForDesiredJPEGPerception(ByRef srcImage As pdDIB, ByVal desiredPerception As jpegAutoQualityMode, Optional ByVal useHighQualityColorMatching As Boolean = False) As Long

    'If desiredPerception is 0 ("do not use auto check"), exit
    If desiredPerception = doNotUseAutoQuality Then
        FindQualityForDesiredJPEGPerception = 0
        Exit Function
    End If
    
    'Based on the requested desiredPerception, calculate an RMSD to aim for
    Dim targetRMSD As Double
    
    Select Case desiredPerception
    
        Case noDifference
            If useHighQualityColorMatching Then
                targetRMSD = 2.2
            Else
                targetRMSD = 4.4
            End If
        
        Case tinyDifference
            If useHighQualityColorMatching Then
                targetRMSD = 4#
            Else
                targetRMSD = 8
            End If
        
        Case minorDifference
            If useHighQualityColorMatching Then
                targetRMSD = 5.5
            Else
                targetRMSD = 11
            End If
        
        Case majorDifference
            If useHighQualityColorMatching Then
                targetRMSD = 7
            Else
                targetRMSD = 14
            End If
    
    End Select
        
    'During high-quality color matching, start by converting the source image into a Single-type L*a*b* array.  We only do this once
    ' to improve performance.
    Dim srcImageData() As Single, dstImageData() As Single
    If useHighQualityColorMatching Then convertEntireDIBToLabColor srcImage, srcImageData
    
    Dim curJPEGQuality As Long
    curJPEGQuality = 90
    
    Dim rmsdCheck As Double
    
    Dim rmsdExceeded As Boolean
    rmsdExceeded = False
    
    Dim tmpJPEGImage As pdDIB
    Set tmpJPEGImage = New pdDIB
    
    'Loop through successively smaller quality values (in series of 10) until the target RMSD is exceeded
    Do
    
        'Retrieve a copy of the original image at the current JPEG quality
        tmpJPEGImage.createFromExistingDIB srcImage
        FillDIBWithJPEGVersion tmpJPEGImage, tmpJPEGImage, curJPEGQuality
        
        'Here is where high-quality and low-quality color-matching diverge.
        If useHighQualityColorMatching Then
        
            'Convert the JPEG-ified DIB to the L*a*b* color space
            convertEntireDIBToLabColor tmpJPEGImage, dstImageData
            
            'Retrieve a mean RMSD for the two images
            rmsdCheck = FindMeanRMSDForTwoArrays(srcImageData, dstImageData, srcImage.getDIBWidth - 1, srcImage.getDIBHeight - 1)
            
        Else
        
            rmsdCheck = FindMeanRMSDForTwoDIBs(srcImage, tmpJPEGImage)
        
        End If
        
        'If the rmsdCheck passes, reduce the JPEG threshold and try again
        If rmsdCheck < targetRMSD Then
            curJPEGQuality = curJPEGQuality - 10
            If curJPEGQuality <= 0 Then rmsdExceeded = True
        Else
            rmsdExceeded = True
        End If
    
    Loop While Not rmsdExceeded
    
    'We now have the nearest acceptable JPEG quality as a multiple of 10.  Drill down further to obtain an exact JPEG quality.
    rmsdExceeded = False
    
    Dim firstJpegCheck As Long
    firstJpegCheck = curJPEGQuality
    
    curJPEGQuality = curJPEGQuality + 9
    
    'Loop through successively smaller quality values (in series of 1) until the target RMSD is exceeded
    Do
    
        'Retrieve a copy of the original image at the current JPEG quality
        tmpJPEGImage.createFromExistingDIB srcImage
        FillDIBWithJPEGVersion tmpJPEGImage, tmpJPEGImage, curJPEGQuality
        
        'Here is where high-quality and low-quality color-matching diverge.
        If useHighQualityColorMatching Then
        
            'Convert the JPEG-ified DIB to the L*a*b* color space
            convertEntireDIBToLabColor tmpJPEGImage, dstImageData
            
            'Retrieve a mean RMSD for the two images
            rmsdCheck = FindMeanRMSDForTwoArrays(srcImageData, dstImageData, srcImage.getDIBWidth - 1, srcImage.getDIBHeight - 1)
            
        Else
        
            rmsdCheck = FindMeanRMSDForTwoDIBs(srcImage, tmpJPEGImage)
        
        End If
        
        'If the rmsdCheck passes, reduce the JPEG threshold and try again
        If rmsdCheck < targetRMSD Then
            curJPEGQuality = curJPEGQuality - 1
            If curJPEGQuality <= firstJpegCheck Then rmsdExceeded = True
        Else
            rmsdExceeded = True
        End If
    
    Loop While Not rmsdExceeded
    
    curJPEGQuality = curJPEGQuality + 1
    If curJPEGQuality = 100 Then curJPEGQuality = 99
    
    'We now have a quality value!  Return it.
    FindQualityForDesiredJPEGPerception = curJPEGQuality

End Function

'This function takes two 24bpp DIBs and compares them, returning a single mean RMSD.
Public Function FindMeanRMSDForTwoDIBs(ByRef srcDib1 As pdDIB, ByRef srcDib2 As pdDIB) As Double

    Dim totalRMSD As Double
    totalRMSD = 0

    Dim x As Long, y As Long, quickX As Long
    
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    
    'Acquire pointers to both DIB arrays
    Dim tmpSA1 As SAFEARRAY2D, tmpSA2 As SAFEARRAY2D
    
    Dim srcArray1() As Byte, srcArray2() As Byte
    
    PrepSafeArray tmpSA1, srcDib1
    PrepSafeArray tmpSA2, srcDib2
    
    CopyMemory ByVal VarPtrArray(srcArray1()), VarPtr(tmpSA1), 4
    CopyMemory ByVal VarPtrArray(srcArray2()), VarPtr(tmpSA2), 4
    
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = srcDib1.getDIBWidth
    imgHeight = srcDib2.getDIBHeight
    
    For x = 0 To imgWidth - 1
        quickX = x * 3
    For y = 0 To imgHeight - 1
    
        'Retrieve both sets of L*a*b* coordinates
        r1 = srcArray1(quickX, y)
        g1 = srcArray1(quickX + 1, y)
        b1 = srcArray1(quickX + 2, y)
        
        r2 = srcArray2(quickX, y)
        g2 = srcArray2(quickX + 1, y)
        b2 = srcArray2(quickX + 2, y)
        
        r1 = (r2 - r1) * (r2 - r1)
        g1 = (g2 - g1) * (g2 - g1)
        b1 = (b2 - b1) * (b2 - b1)
        
        'Calculate an RMSD
        totalRMSD = totalRMSD + Sqr(r1 + g1 + b1)
    
    Next y
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcArray1), 0&, 4
    CopyMemory ByVal VarPtrArray(srcArray2), 0&, 4
    
    'Divide the total RMSD by the number of pixels in the image, then exit
    FindMeanRMSDForTwoDIBs = totalRMSD / (imgWidth * imgHeight)

End Function

'This function assumes two 24bpp DIBs have been pre-converted to Single-type L*a*b* arrays.  Use the L*a*b* data to return
' a mean RMSD for the two images.
Public Function FindMeanRMSDForTwoArrays(ByRef srcArray1() As Single, ByRef srcArray2() As Single, ByVal imgWidth As Long, ByVal imgHeight As Long) As Double

    Dim totalRMSD As Double
    totalRMSD = 0

    Dim x As Long, y As Long, quickX As Long
    
    Dim LabL1 As Double, LabA1 As Double, LabB1 As Double
    Dim labL2 As Double, labA2 As Double, labB2 As Double
    
    For x = 0 To imgWidth - 1
        quickX = x * 3
    For y = 0 To imgHeight - 1
    
        'Retrieve both sets of L*a*b* coordinates
        LabL1 = srcArray1(quickX, y)
        LabA1 = srcArray1(quickX + 1, y)
        LabB1 = srcArray1(quickX + 2, y)
        
        labL2 = srcArray2(quickX, y)
        labA2 = srcArray2(quickX + 1, y)
        labB2 = srcArray2(quickX + 2, y)
        
        'Calculate an RMSD
        totalRMSD = totalRMSD + distanceThreeDimensions(LabL1, LabA1, LabB1, labL2, labA2, labB2)
    
    Next y
    Next x
    
    'Divide the total RMSD by the number of pixels in the image, then exit
    FindMeanRMSDForTwoArrays = totalRMSD / (imgWidth * imgHeight)

End Function

'Given a source and destination DIB reference, fill the destination with a post-JPEG-2000-compression of the original.  This
' is used to generate the live preview used in PhotoDemon's "export JPEG-2000" dialog.
Public Sub FillDIBWithJP2Version(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal JP2Quality As Long)
    
    'Pass the DIB to FreeImage, which will make a copy for itself.
    Dim fi_DIB As Long
    fi_DIB = Plugin_FreeImage.GetFIHandleFromPDDib_NoCopy(srcDIB)
        
    'Now comes the actual JPEG-2000 conversion, which is handled exclusively by FreeImage.  Basically, we ask it to save
    ' the image in JPEG-2000 format to a byte array; we then hand that byte array back to it and request a decompression.
    Dim jp2Array() As Byte
    Dim fi_Check As Long
    fi_Check = FreeImage_SaveToMemoryEx(PDIF_JP2, fi_DIB, jp2Array, JP2Quality, True)
    
    fi_DIB = FreeImage_LoadFromMemoryEx(jp2Array, 0, , PDIF_JP2)
    
    'Copy the newly decompressed JPEG-2000 into the destination pdDIB object.
    Plugin_FreeImage.PaintFIDibToPDDib dstDIB, fi_DIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight
    
    'Release the FreeImage copy of the DIB.
    FreeImage_Unload fi_DIB
    
End Sub

'Given a source and destination DIB reference, fill the destination with a post-WebP-compression of the original.  This
' is used to generate the live preview used in PhotoDemon's "export WebP" dialog.
Public Sub FillDIBWithWebPVersion(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal WebPQuality As Long)
    
    'Pass the DIB to FreeImage, which will make a copy for itself.
    Dim fi_DIB As Long
    fi_DIB = Plugin_FreeImage.GetFIHandleFromPDDib_NoCopy(srcDIB)
        
    'Now comes the actual WebP conversion, which is handled exclusively by FreeImage.  Basically, we ask it to save
    ' the image in WebP format to a byte array; we then hand that byte array back to it and request a decompression.
    Dim webPArray() As Byte
    Dim fi_Check As Long
    fi_Check = FreeImage_SaveToMemoryEx(PDIF_WEBP, fi_DIB, webPArray, WebPQuality, True)
    
    fi_DIB = FreeImage_LoadFromMemoryEx(webPArray, , , PDIF_WEBP)
    
    'Random fact: the WebP encoder will automatically downsample 32-bit images with pointless alpha channels to 24-bit.  This causes problems when
    ' we try to preview WebP files prior to encoding, as it may randomly change the bit-depth on us.  Check for this case, and recreate the target
    ' DIB as necessary.
    If FreeImage_GetBPP(fi_DIB) <> dstDIB.getDIBColorDepth Then dstDIB.createBlank dstDIB.getDIBWidth, dstDIB.getDIBHeight, FreeImage_GetBPP(fi_DIB)
        
    'Copy the newly decompressed image into the destination pdDIB object.
    Plugin_FreeImage.PaintFIDibToPDDib dstDIB, fi_DIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight
    
    'Release the FreeImage copy of the DIB.
    FreeImage_Unload fi_DIB
    
End Sub

'Given a source and destination DIB reference, fill the destination with a post-JXR-compression of the original.  This
' is used to generate the live preview used in PhotoDemon's "export JXR" dialog.
Public Sub FillDIBWithJXRVersion(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal jxrQuality As Long)
    
    'Pass the DIB to FreeImage, which will make a copy for itself.
    Dim fi_DIB As Long
    fi_DIB = Plugin_FreeImage.GetFIHandleFromPDDib_NoCopy(srcDIB)
        
    'Now comes the actual JPEG XR conversion, which is handled exclusively by FreeImage.  Basically, we ask it to save
    ' the image in JPEG XR format to a byte array; we then hand that byte array back to it and request a decompression.
    Dim jxrArray() As Byte
    Dim fi_Check As Boolean
    fi_Check = FreeImage_SaveToMemoryEx(PDIF_JXR, fi_DIB, jxrArray, jxrQuality, True)
    Debug.Print "JXR live previews have been problematic; size of returned array is: " & UBound(jxrArray)
    
    If fi_Check Then
    
        fi_DIB = FreeImage_LoadFromMemoryEx(jxrArray, 0, UBound(jxrArray) + 1, PDIF_JXR, VarPtr(jxrArray(0)))
        
        'Copy the newly decompressed image into the destination pdDIB object.
        If fi_DIB <> 0 Then
            Plugin_FreeImage.PaintFIDibToPDDib dstDIB, fi_DIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight
        Else
            Debug.Print "Failed to load JXR from memory; FreeImage didn't return a DIB from FreeImage_LoadFromMemoryEx()"
        End If
    
    Else
        Debug.Print "Failed to save JXR to memory; FreeImage returned FALSE for FreeImage_SaveToMemoryEx()"
    End If
    
    'Release the FreeImage copy of the DIB.
    If fi_DIB <> 0 Then FreeImage_Unload fi_DIB
    
End Sub

'Save a new Undo/Redo entry to file.  This function is only called by the createUndoData function in the pdUndo class.
' For the most part, this function simply wraps other save functions; however, certain odd types of Undo diff files (e.g. layer headers)
' may be directly processed and saved by this function.
'
'Note that this function interacts closely with the matching LoadUndo function in the Loading module.  Any novel Undo diff types added
' here must also be mirrored there.
Public Function SaveUndoData(ByRef srcPDImage As pdImage, ByRef dstUndoFilename As String, ByVal processType As PD_UNDO_TYPE, Optional ByVal targetLayerID As Long = -1) As Boolean
    
    #If DEBUGMODE = 1 Then
        Dim timeAtUndoStart As Double
        timeAtUndoStart = Timer
    #End If
    
    'What kind of Undo data we save is determined by the current processType.
    Select Case processType
    
        'EVERYTHING, meaning a full copy of the pdImage stack and any selection data
        Case UNDO_EVERYTHING
            Saving.SavePhotoDemonImage srcPDImage, dstUndoFilename, True, True, IIf(g_UndoCompressionLevel = 0, False, True), False, False, False, IIf(g_UndoCompressionLevel = 0, -1, g_UndoCompressionLevel), , , True
            srcPDImage.mainSelection.writeSelectionToFile dstUndoFilename & ".selection"
            
        'A full copy of the pdImage stack
        Case UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE
            Saving.SavePhotoDemonImage srcPDImage, dstUndoFilename, True, True, IIf(g_UndoCompressionLevel = 0, False, True), False, False, False, IIf(g_UndoCompressionLevel = 0, -1, g_UndoCompressionLevel), , , True
        
        'A full copy of the pdImage stack, *without any layer DIB data*
        Case UNDO_IMAGEHEADER
            Saving.SavePhotoDemonImage srcPDImage, dstUndoFilename, True, IIf(g_UndoCompressionLevel = 0, False, True), False, False, True, , , , , True
        
        'Layer data only (full layer header + full layer DIB).
        Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE
            Saving.SavePhotoDemonLayer srcPDImage.GetLayerByID(targetLayerID), dstUndoFilename & ".layer", True, True, IIf(g_UndoCompressionLevel = 0, False, True), False, False, IIf(g_UndoCompressionLevel = 0, -1, g_UndoCompressionLevel), True
        
        'Layer header data only (e.g. DO NOT WRITE OUT THE LAYER DIB)
        Case UNDO_LAYERHEADER
            Saving.SavePhotoDemonLayer srcPDImage.GetLayerByID(targetLayerID), dstUndoFilename & ".layer", True, True, False, False, True, , True
            
        'Selection data only
        Case UNDO_SELECTION
            srcPDImage.mainSelection.writeSelectionToFile dstUndoFilename & ".selection"
            
        'Anything else (this should never happen, but good to have a failsafe)
        Case Else
            Debug.Print "Unknown Undo data write requested - is it possible to avoid this request entirely??"
            Saving.SavePhotoDemonImage srcPDImage, dstUndoFilename, True, False, False, False, , , , , , True
        
    End Select
    
    #If DEBUGMODE = 1 Then
        'Want to test undo timing?  Uncomment the line below
        Debug.Print "Time take for Undo file creation: " & Format$(Timer - timeAtUndoStart, "#0.####") & " seconds"
    #End If
    
End Function

'Quickly save a DIB to file in PNG format.  Things like PD's Recent File manager use this function to quickly write DIBs out to file.
Public Function QuickSaveDIBAsPNG(ByVal dstFilename As String, ByRef srcDIB As pdDIB) As Boolean

    'Perform a few failsafe checks
    If (srcDIB Is Nothing) Then
        QuickSaveDIBAsPNG = False
        Exit Function
    End If
    
    If (srcDIB.getDIBWidth = 0) Or (srcDIB.getDIBHeight = 0) Then
        QuickSaveDIBAsPNG = False
        Exit Function
    End If

    'If FreeImage is available, use it to save the PNG; otherwise, fall back to GDI+
    If g_ImageFormats.FreeImageEnabled Then
        
        'PD exclusively uses premultiplied alpha for internal DIBs (unless image processing math dictates otherwise).
        ' Saved files always use non-premultiplied alpha.  If the source image is premultiplied, we want to create a
        ' temporary non-premultiplied copy.
        Dim alphaWasChanged As Boolean
        If srcDIB.getAlphaPremultiplication Then
            srcDIB.SetAlphaPremultiplication False
            alphaWasChanged = True
        End If
        
        'Convert the temporary DIB to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIHandleFromPDDib_NoCopy(srcDIB, True)
    
        'Use that handle to save the image to PNG format
        If fi_DIB <> 0 Then
            Dim fi_Check As Long
            
            'Output the PNG file at the proper color depth
            Dim fi_OutputColorDepth As FREE_IMAGE_COLOR_DEPTH
            If srcDIB.getDIBColorDepth = 24 Then
                fi_OutputColorDepth = FICD_24BPP
            Else
                fi_OutputColorDepth = FICD_32BPP
            End If
            
            'Ask FreeImage to write the thumbnail out to file
            fi_Check = FreeImage_SaveEx(fi_DIB, dstFilename, PDIF_PNG, FISO_PNG_Z_BEST_SPEED, fi_OutputColorDepth, , , , , True)
            If Not fi_Check Then Message "Thumbnail save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
            
        Else
            Message "Thumbnail save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
        End If
        
        If alphaWasChanged Then srcDIB.SetAlphaPremultiplication True
        
    'FreeImage is not available; try to use GDI+ to save a PNG thumbnail
    Else
        
        If Not GDIPlusQuickSavePNG(dstFilename, srcDIB) Then Message "Thumbnail save failed (unspecified GDI+ error)."
        
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
