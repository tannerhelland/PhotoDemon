Attribute VB_Name = "Saving"
'***************************************************************************
'File Saving Interface
'Copyright ©2001-2013 by Tanner Helland
'Created: 4/15/01
'Last updated: 20/September/13
'Last update: Implemented full support for metadata caching.  If metadata is not loaded at image load-time, the save function will now
'              detect this and cache any relevant metadata before overwriting a file.
'
'Module responsible for all image saving, with the exception of the GDI+ image save function (which has been left in the GDI+ module
' for consistency's sake).  Export functions are sorted by file type, and most serve as relatively lightweight wrappers to corresponding
' functions in the FreeImage plugin.
'
'The most important sub is PhotoDemon_SaveImage at the top of the module.  This sub is responsible for doing a multitude of handling before
' and after saving an image, including items like selecting correct color depth prior to save, and requesting MRU updates post-save.  If
' passed the loadRelevantForm parameter, it will also prompt the user for any format-specific settings (such as JPEG quality).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'This routine will blindly save the mainLayer contents (from the pdImage object specified by srcPDImage) to dstPath.
' It is up to the calling routine to make sure this is what is wanted. (Note: this routine will erase any existing image
' at dstPath, so BE VERY CAREFUL with what you send here!)
'
'INPUTS:
'   1) pdImage to be saved
'   2) Destination file path
'   3) Optional: imageID (if provided, the function can write information about the save to the relevant object in the pdImages array - this primarily exists for legacy reasons)
'   4) Optional: whether to display a form for the user to input additional save options (JPEG quality, etc)
'   5) Optional: a string of relevant save parameters.  If this is not provided, relevant parameters will be loaded from the preferences file.
Public Function PhotoDemon_SaveImage(ByRef srcPDImage As pdImage, ByVal dstPath As String, Optional ByVal imageID As Long = -1, Optional ByVal loadRelevantForm As Boolean = False, Optional ByVal saveParamString As String = "", Optional ByVal forceColorDepthMethod As Long = -1) As Boolean
    
    'Only update the MRU list if 1) no form is shown (because the user may cancel it), 2) a form was shown and the user
    ' successfully navigated it, and 3) no errors occured during the export process.  By default, this is set to "do not update."
    Dim updateMRU As Boolean
    updateMRU = False
    
    'Start by determining the output format for this image (which was set either by a "Save As" common dialog box,
    ' or by copying the image's original format - or, if in the midst of a batch process, by the user via the batch wizard).
    Dim saveFormat As Long
    saveFormat = srcPDImage.currentFileFormat


    '****************************************************************************************************
    ' Determine exported color depth
    '****************************************************************************************************

    'The user is allowed to set a persistent preference for output color depth.  This setting affects a "color depth"
    ' parameter that will be sent to the various format-specific save file routines.  The available preferences are:
    ' 0) Mimic the file's original color depth (if available; this may not always be possible, e.g. saving a 32bpp PNG as JPEG)
    ' 1) Count the number of colors used, and save the file based on that (again, if possible)
    ' 2) Prompt the user for their desired export color depth
    '
    'Batch processing allows the user to overwrite their default preference with a specific preference for that batch process;
    ' if this occurs, the "forceColorDepthMethod" is utilized.
    Dim outputColorDepth As Long
    
    'Finally, note that JPEG exporting, on account of it being somewhat specialized, ignores this step completely.
    ' The JPEG routine will do its own scan for grayscale/color and save the proper format automatically.
    If saveFormat <> FIF_JPEG Then
    
        Dim colorDepthMode As Long
        If forceColorDepthMethod = -1 Then
            colorDepthMode = g_UserPreferences.GetPref_Long("Saving", "Outgoing Color Depth", 1)
        Else
            colorDepthMode = forceColorDepthMethod
        End If
    
        Select Case colorDepthMode
        
            'Maintain the file's original color depth (if possible)
            Case 0
                
                'Check to see if this format supports the image's original color depth
                If g_ImageFormats.isColorDepthSupported(saveFormat, srcPDImage.originalColorDepth) Then
                    
                    'If it IS supported, set the original color depth as the output color depth for this save
                    outputColorDepth = srcPDImage.originalColorDepth
                    Message "Original color depth of %1 bpp is supported by this format.  Proceeding with save...", outputColorDepth
                
                'If it IS NOT supported, we need to find the closest available color depth for this format.
                Else
                    outputColorDepth = g_ImageFormats.getClosestColorDepth(saveFormat, srcPDImage.originalColorDepth)
                    Message "Original color depth of %1 bpp is not supported by this format.  Proceeding to save as %2 bpp...", srcPDImage.originalColorDepth, outputColorDepth
                
                End If
            
            'Count colors used
            Case 1
            
                'Count the number of colors in the image.  (The function will automatically cease if it hits 257 colors,
                ' as anything above 256 colors is treated as 24bpp.)
                Dim colorCountCheck As Long
                Message "Counting image colors to determine optimal exported color depth..."
                If imageID <> -1 Then
                    colorCountCheck = getQuickColorCount(srcPDImage, imageID)
                Else
                    colorCountCheck = getQuickColorCount(srcPDImage)
                End If
                
                'If 256 or less colors were found in the image, mark it as 8bpp.  Otherwise, mark it as 24 or 32bpp.
                outputColorDepth = getColorDepthFromColorCount(colorCountCheck, srcPDImage.mainLayer)
                
                'A special case arises when an image has <= 256 colors, but a non-binary alpha channel.  PNG allows for
                ' this, but other formats do not.  Because even the PNG transformation is not lossless, set these types of
                ' images to be exported as 32bpp.
                If (outputColorDepth <= 8) And (srcPDImage.mainLayer.getLayerColorDepth = 32) Then
                    If (Not srcPDImage.mainLayer.isAlphaBinary) Then outputColorDepth = 32
                End If
                
                Message "Color count successful (%1 bpp recommended)", outputColorDepth
                
                'As with case 0, we now need to see if this format supports the suggested color depth
                If g_ImageFormats.isColorDepthSupported(saveFormat, outputColorDepth) Then
                    
                    'If it IS supported, set the original color depth as the output color depth for this save
                    Message "Recommended color depth of %1 bpp is supported by this format.  Proceeding with save...", outputColorDepth
                
                'If it IS NOT supported, we need to find the closest available color depth for this format.
                Else
                    outputColorDepth = g_ImageFormats.getClosestColorDepth(saveFormat, outputColorDepth)
                    Message "Recommended color depth of %1 bpp is not supported by this format.  Proceeding to save as %2 bpp...", srcPDImage.originalColorDepth, outputColorDepth
                
                End If
            
            'Prompt the user (but only if necessary)
            Case 2
            
                'First, check to see if the save format in question supports multiple color depths
                If g_ImageFormats.doesFIFSupportMultipleColorDepths(saveFormat) Then
                    
                    'If it does, provide the user with a prompt to choose whatever color depth they'd like
                    Dim dCheck As VbMsgBoxResult
                    dCheck = promptColorDepth(saveFormat)
                    
                    If dCheck = vbOK Then
                        outputColorDepth = g_ColorDepth
                    Else
                        PhotoDemon_SaveImage = False
                        Message "Save canceled."
                        Exit Function
                    End If
                
                'If this format only supports a single output color depth, don't bother the user with a prompt
                Else
            
                    outputColorDepth = g_ImageFormats.getClosestColorDepth(saveFormat, srcPDImage.originalColorDepth)
            
                End If
                
            'A color depth has been explicitly specified by the forceColorDepthMethod parameter.  We can find the color depth
            ' by subtracting 16 from the parameter value.
            Case Else
            
                outputColorDepth = forceColorDepthMethod - 16
                
                'As a failsafe, make sure this format supports the suggested color depth
                If g_ImageFormats.isColorDepthSupported(saveFormat, outputColorDepth) Then
                    
                    'If it IS supported, set the original color depth as the output color depth for this save
                    Message "Requested color depth of %1 bpp is supported by this format.  Proceeding with save...", outputColorDepth
                
                'If it IS NOT supported, we need to find the closest available color depth for this format.
                Else
                    outputColorDepth = g_ImageFormats.getClosestColorDepth(saveFormat, outputColorDepth)
                    Message "Requested color depth of %1 bpp is not supported by this format.  Proceeding to save as %2 bpp...", srcPDImage.originalColorDepth, outputColorDepth
                
                End If
            
        End Select
    
    End If
    
    
    '****************************************************************************************************
    ' If additional save parameters have been provided, we may need to access them prior to saving.  Parse them now.
    '****************************************************************************************************
    
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString saveParamString
    
    'NOTE: If explicit save parameters were not provided, relevant values will be loaded from the preferences file
    '      on a per-format basis.
        
    
    '****************************************************************************************************
    ' Before saving the image (especially in the case of an overwrite), we need to cache the source image's metadata.
    '****************************************************************************************************
    
    'If metadata export is enabled, cache the metadata now
    If g_UserPreferences.GetPref_Long("Saving", "Metadata Export", 1) <> 3 Then
    
        Message "Caching current image's metadata..."
        
        'PhotoDemon stores metadata in two possible ways: a full binary copy (necessary for "preserve all metadata regardless of relevance")
        ' or a full XML copy (necessary for all other metadata options).  Check which type the user wants us to write, and make sure we
        ' have a copy of the image's metadata in that format.  If we don't, cache it now.
        If g_UserPreferences.GetPref_Long("Saving", "Metadata Export", 1) = 0 Then
        
            'Binary metadata is requested.  Cache it now (if necessary).
            If Not srcPDImage.imgMetadata.hasBinaryMetadata Then
                srcPDImage.imgMetadata.quickCacheMetadata srcPDImage.locationOnDisk
            End If
        
        Else
        
            'XML metadata is requested.  Cache it now (if necessary).
            If Not srcPDImage.imgMetadata.hasXMLMetadata Then
                srcPDImage.imgMetadata.loadAllMetadata srcPDImage.locationOnDisk, srcPDImage.originalFileFormat
            End If
        
        End If
        
    End If
    
    'If available, metadata has now been cached to memory.  This means we can delete or overwrite the source file without
    ' losing its metadata contents.
    
    
    '****************************************************************************************************
    ' Based on the requested file type and color depth, call the appropriate save function
    '****************************************************************************************************
        
    Select Case saveFormat
        
        'JPEG
        Case FIF_JPEG
        
            'JPEG files may need to display a dialog box so the user can set compression quality
            If loadRelevantForm Then
                
                Dim gotSettings As VbMsgBoxResult
                gotSettings = promptJPEGSettings
                
                'If the dialog was canceled, note it.  Otherwise, remember that the user has seen the JPEG save screen at least once.
                If gotSettings = vbOK Then
                    srcPDImage.hasSeenJPEGPrompt = True
                    PhotoDemon_SaveImage = True
                Else
                    PhotoDemon_SaveImage = False
                    Message "Save canceled."
                    Exit Function
                End If
                
                'If the user clicked OK, replace the function's save parameters with the ones set by the user
                cParams.setParamString CStr(g_JPEGQuality)
                cParams.setParamString cParams.getParamString & "|" & g_JPEGFlags
                cParams.setParamString cParams.getParamString & "|" & g_JPEGThumbnail
                
            End If
            
            'Store these JPEG settings in the image object so we don't have to pester the user for it if they save again
            srcPDImage.saveParameters = cParams.getParamString
                            
            'I implement two separate save functions for JPEG images: FreeImage and GDI+.  GDI+ does not need to make a copy
            ' of the image before saving it - which makes it much faster - but FreeImage provides a number of additional
            ' parameters, like optimization, thumbnail embedding, and custom subsampling.  If no optional parameters are in use
            ' (or if FreeImage is unavailable), use GDI+.  Otherwise, use FreeImage.
            If g_ImageFormats.FreeImageEnabled And (cParams.doesParamExist(2) Or cParams.doesParamExist(3)) Then
                updateMRU = SaveJPEGImage(srcPDImage, dstPath, cParams.getParamString)
            ElseIf g_ImageFormats.GDIPlusEnabled Then
                updateMRU = GDIPlusSavePicture(srcPDImage, dstPath, ImageJPEG, 24, cParams.GetLong(1, 92))
            Else
                Message "No %1 encoder found. Save aborted.", "JPEG"
                PhotoDemon_SaveImage = False
                Exit Function
            End If
            
            
        'PDI, PhotoDemon's internal format
        Case 100
            If g_ZLibEnabled Then
                updateMRU = SavePhotoDemonImage(srcPDImage, dstPath)
            Else
            'If zLib doesn't exist...
                pdMsgBox "The zLib compression library (zlibwapi.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable PDI saving, please allow %1 to download plugin updates by going to the Tools -> Options menu, and selecting the 'offer to download core plugins' check box.", vbExclamation + vbOKOnly + vbApplicationModal, " PDI Interface Error", PROGRAMNAME
                Message "No %1 encoder found. Save aborted.", "PDI"
            End If
        
        'GIF
        Case FIF_GIF
        
            'GIFs are preferentially exported by FreeImage, then GDI+ (if available)
            If g_ImageFormats.FreeImageEnabled Then
                If Not cParams.doesParamExist(1) Then
                    updateMRU = SaveGIFImage(srcPDImage, dstPath)
                Else
                    updateMRU = SaveGIFImage(srcPDImage, dstPath, cParams.GetLong(1))
                End If
            ElseIf g_ImageFormats.GDIPlusEnabled Then
                updateMRU = GDIPlusSavePicture(srcPDImage, dstPath, ImageGIF, 8)
            Else
                Message "No %1 encoder found. Save aborted.", "GIF"
                PhotoDemon_SaveImage = False
                Exit Function
            End If
            
        'PNG
        Case FIF_PNG
            
            'PNGs support a number of specialized parameters.  If we weren't passed any, retrieve corresponding values from the preferences
            ' file (specifically, the parameters include: PNG compression level (0-9), interlacing (bool), BKGD preservation (bool).)
            If Not cParams.doesParamExist(1) Then
                cParams.setParamString CStr(g_UserPreferences.GetPref_Long("File Formats", "PNG Compression", 9))
                cParams.setParamString cParams.getParamString() & "|" & CStr(g_UserPreferences.GetPref_Boolean("File Formats", "PNG Interlacing", False))
                cParams.setParamString cParams.getParamString() & "|" & CStr(g_UserPreferences.GetPref_Boolean("File Formats", "PNG Background Color", True))
            End If
            
            'PNGs are preferentially exported by FreeImage, then GDI+ (if available)
            If g_ImageFormats.FreeImageEnabled Then
                updateMRU = SavePNGImage(srcPDImage, dstPath, outputColorDepth, cParams.getParamString)
            ElseIf g_ImageFormats.GDIPlusEnabled Then
                updateMRU = GDIPlusSavePicture(srcPDImage, dstPath, ImagePNG, outputColorDepth)
            Else
                Message "No %1 encoder found. Save aborted.", "PNG"
                PhotoDemon_SaveImage = False
                Exit Function
            End If
            
        'PPM
        Case FIF_PPM
            If Not cParams.doesParamExist(1) Then cParams.setParamString CStr(g_UserPreferences.GetPref_Long("File Formats", "PPM Export Format", 0))
            updateMRU = SavePPMImage(srcPDImage, dstPath, cParams.getParamString)
                
        'TGA
        Case FIF_TARGA
            If Not cParams.doesParamExist(1) Then cParams.setParamString CStr(g_UserPreferences.GetPref_Boolean("File Formats", "TGA RLE", False))
            updateMRU = SaveTGAImage(srcPDImage, dstPath, outputColorDepth, cParams.getParamString)
            
        'JPEG-2000
        Case FIF_JP2
        
            If loadRelevantForm = True Then
                
                Dim gotJP2Settings As VbMsgBoxResult
                gotJP2Settings = promptJP2Settings()
                
                'If the dialog was canceled, note it.  Otherwise, remember that the user has seen the JPEG save screen at least once.
                If gotJP2Settings = vbOK Then
                    srcPDImage.hasSeenJP2Prompt = True
                    PhotoDemon_SaveImage = True
                Else
                    PhotoDemon_SaveImage = False
                    Message "Save canceled."
                    Exit Function
                End If
                
                'If the user clicked OK, replace the functions save parameters with the ones set by the user
                cParams.setParamString CStr(g_JP2Compression)
                
            End If
            
            'Store the JPEG-2000 quality in the image object so we don't have to pester the user for it if they save again
            srcPDImage.saveParameters = cParams.getParamString
        
            updateMRU = SaveJP2Image(srcPDImage, dstPath, outputColorDepth, cParams.getParamString)
            
        'TIFF
        Case FIF_TIFF
            
            'TIFFs use two parameters - compression type, and CMYK encoding (true/false)
            If Not cParams.doesParamExist(1) Then
                cParams.setParamString CStr(g_UserPreferences.GetPref_Long("File Formats", "TIFF Compression", 0)) & "|" & CStr(g_UserPreferences.GetPref_Boolean("File Formats", "TIFF CMYK", False))
            End If
            
            'TIFFs are preferentially exported by FreeImage, then GDI+ (if available)
            If g_ImageFormats.FreeImageEnabled Then
                updateMRU = SaveTIFImage(srcPDImage, dstPath, outputColorDepth, cParams.getParamString)
            ElseIf g_ImageFormats.GDIPlusEnabled Then
                updateMRU = GDIPlusSavePicture(srcPDImage, dstPath, ImageTIFF, outputColorDepth)
            Else
                Message "No %1 encoder found. Save aborted.", "TIFF"
                PhotoDemon_SaveImage = False
                Exit Function
            End If
        
        'Anything else must be a bitmap
        Case FIF_BMP
            
            'If the user has not provided explicit BMP parameters, load their default values from the preferences file
            If Not cParams.doesParamExist(1) Then cParams.setParamString CStr(g_UserPreferences.GetPref_Boolean("File Formats", "Bitmap RLE", False))
            updateMRU = SaveBMP(srcPDImage, dstPath, outputColorDepth, cParams.getParamString)
            
        Case Else
            Message "Output format not recognized.  Save aborted.  Please use the Help -> Submit Bug Report menu item to report this incident."
            PhotoDemon_SaveImage = False
            Exit Function
        
    End Select
    
    '****************************************************************************************************
    ' If the file was successfully written, we can now embed any additional metadata.
    '****************************************************************************************************
    
    'Note: I don't like embedding metadata in a separate step, but that's a necessary evil of routing all metadata handling
    ' through an external plugin.
    
    'Note that updateMRU is used to track save file success, so it will only be TRUE if the image file was written successfully.
    ' If the file was not written successfully, abandon any attempts at metadata embedding.
    If updateMRU And g_ExifToolEnabled Then
        
        'Only attempt to export metadata if ExifTool was able to successfully parse or cache metadata prior to saving
        If srcPDImage.imgMetadata.hasXMLMetadata Or srcPDImage.imgMetadata.hasBinaryMetadata Then
            updateMRU = srcPDImage.imgMetadata.writeAllMetadata(dstPath, g_UserPreferences.GetPref_Long("Saving", "Metadata Export", 1), srcPDImage)
        Else
            Message "No metadata to export.  Continuing save..."
        End If
                
    End If
    
    'UpdateMRU should only be true if the save was successful
    If updateMRU Then
    
        'Additionally, only add this MRU to the list (and generate an accompanying icon) if we are not in the midst
        ' of a batch conversion.
        If MacroStatus <> MacroBATCH Then
        
            'Add this file to the MRU list
            MRU_AddNewFile dstPath, srcPDImage
        
            'Remember the file's location for future saves
            srcPDImage.locationOnDisk = dstPath
            
            'Remember the file's filename
            Dim tmpFilename As String
            tmpFilename = dstPath
            StripFilename tmpFilename
            srcPDImage.originalFileNameAndExtension = tmpFilename
            StripOffExtension tmpFilename
            srcPDImage.originalFileName = tmpFilename
            
            'Mark this file as having been saved
            srcPDImage.setSaveState True
            
            PhotoDemon_SaveImage = True
            
            'Update the interface to match the newly saved image (e.g. disable the Save button)
            syncInterfaceToCurrentImage
            
            'Notify the thumbnail window that this image has been updated (so it can show/hide the save icon)
            If Not srcPDImage.forInternalUseOnly Then toolbar_ImageTabs.notifyUpdatedImage srcPDImage.imageID
            
        End If
    
    Else
        
        'If we aren't updating the MRU, something went wrong.  Display that the save was canceled and exit.
        Message "Save canceled."
        PhotoDemon_SaveImage = False
        Exit Function
        
    End If

    Message "Save complete."

End Function


'Save the current image to BMP format
Public Function SaveBMP(ByRef srcPDImage As pdImage, ByVal BMPPath As String, ByVal outputColorDepth As Long, Optional ByVal bmpParams As String = "") As Boolean
    
    On Error GoTo SaveBMPError
    
    'Parse all possible BMP parameters (at present there is only one possible parameter, which specifies RLE compression for 8bpp images)
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(bmpParams) > 0 Then cParams.setParamString bmpParams
    Dim BMPCompression As Boolean
    BMPCompression = cParams.GetBool(1, False)
    
    Dim sFileType As String
    sFileType = "BMP"
    
    'If the output color depth is 24 or 32bpp, or if both GDI+ and FreeImage are missing, use our own internal methods
    ' to save a BMP file
    If (outputColorDepth = 24) Or (outputColorDepth = 32) Or ((Not g_ImageFormats.GDIPlusEnabled) And (Not g_ImageFormats.FreeImageEnabled)) Then
    
        Message "Saving %1 file...", sFileType
    
        'The layer class is capable of doing this without any outside help.
        srcPDImage.mainLayer.writeToBitmapFile BMPPath
    
        Message "%1 save complete.", sFileType
        
    'If some other color depth is specified, use FreeImage or GDI+ to write the file
    Else
    
        If g_ImageFormats.FreeImageEnabled Then
            
            'Load FreeImage into memory
            Dim hLib As Long
            hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
            
            Message "Preparing %1 image...", sFileType
            
            'Copy the image into a temporary layer
            Dim tmpLayer As pdLayer
            Set tmpLayer = New pdLayer
            tmpLayer.createFromExistingLayer srcPDImage.mainLayer
            
            'If the output color depth is 24 but the current image is 32, composite the image against a white background
            If (outputColorDepth < 32) And (srcPDImage.mainLayer.getLayerColorDepth = 32) Then tmpLayer.convertTo24bpp
            
            'Convert our current layer to a FreeImage-type DIB
            Dim fi_DIB As Long
            fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
                        
            'If the image is being reduced from some higher bit-depth to 1bpp, manually force a conversion with dithering
            If outputColorDepth = 1 Then fi_DIB = FreeImage_Dither(fi_DIB, FID_FS)
            
            'Finally, prepare some BMP save flags.  If the user has requested RLE encoding, and this image is <= 8bpp,
            ' request RLE encoding from FreeImage.
            Dim BMPflags As Long
            BMPflags = BMP_DEFAULT
            
            If outputColorDepth = 8 And BMPCompression Then BMPflags = BMP_SAVE_RLE
            
            'Use that handle to save the image to BMP format, with required color conversion based on the outgoing color depth
            If fi_DIB <> 0 Then
                Dim fi_Check As Long
                fi_Check = FreeImage_SaveEx(fi_DIB, BMPPath, FIF_BMP, BMPflags, outputColorDepth, , , , , True)
                If fi_Check = False Then
                    Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
                    FreeLibrary hLib
                    SaveBMP = False
                    Exit Function
                Else
                    Message "%1 save complete.", sFileType
                End If
            Else
                Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
                SaveBMP = False
                FreeLibrary hLib
                Exit Function
            End If
    
            'Release FreeImage from memory
            FreeLibrary hLib
            
        Else
            GDIPlusSavePicture srcPDImage, BMPPath, ImageBMP, outputColorDepth
        End If
    
    End If
    
    SaveBMP = True
    Exit Function
    
SaveBMPError:

    If hLib <> 0 Then FreeLibrary hLib

    SaveBMP = False
    
End Function

'Save the current image to PhotoDemon's native PDI format
Public Function SavePhotoDemonImage(ByRef srcPDImage As pdImage, ByVal PDIPath As String) As Boolean
    
    On Error GoTo SavePDIError
    
    Dim sFileType As String
    sFileType = "PDI"
    
    Message "Saving %1 image...", sFileType

    'First, have the layer write itself to file in BMP format
    srcPDImage.mainLayer.writeToBitmapFile PDIPath
    
    'Then compress the file using zLib
    CompressFile PDIPath
    
    Message "%1 save complete.", sFileType
    
    SavePhotoDemonImage = True
    Exit Function
    
SavePDIError:

    SavePhotoDemonImage = False
    
End Function

'Save a GIF (Graphics Interchange Format) image.  GDI+ can also do this.
Public Function SaveGIFImage(ByRef srcPDImage As pdImage, ByVal GIFPath As String, Optional ByVal forceAlphaConvert As Long = -1) As Boolean

    On Error GoTo SaveGIFError
    
    Dim sFileType As String
    sFileType = "GIF"

    'Make sure we found the plug-in when we loaded the program
    If g_ImageFormats.FreeImageEnabled = False Then
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveGIFImage = False
        Exit Function
    End If
    
    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
    
    Message "Preparing %1 image...", sFileType
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer srcPDImage.mainLayer
    
    'If the current image is 32bpp, we will need to apply some additional actions to the image to prepare the
    ' transparency.  Mark a bool value, because we will reference it in multiple places throughout the save function.
    Dim handleAlpha As Boolean
    If srcPDImage.mainLayer.getLayerColorDepth = 32 Then handleAlpha = True Else handleAlpha = False
    
    'If the current image contains transparency, we need to modify it in order to retain the alpha channel.
    If handleAlpha Then
    
        'Does this layer contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If tmpLayer.isAlphaBinary Then
            tmpLayer.applyAlphaCutoff
        Else
            If forceAlphaConvert = -1 Then
                Dim alphaCheck As VbMsgBoxResult
                alphaCheck = promptAlphaCutoff(tmpLayer)
                
                'If the alpha dialog is canceled, abandon the entire save
                If alphaCheck = vbCancel Then
                
                    tmpLayer.eraseLayer
                    Set tmpLayer = Nothing
                    SaveGIFImage = False
                    Exit Function
                
                'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                Else
                    tmpLayer.applyAlphaCutoff g_AlphaCutoff
                End If
            Else
                tmpLayer.applyAlphaCutoff forceAlphaConvert
            End If
            
        End If
    
    End If
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'If the image contains alpha, we need to convert the FreeImage copy of the image to 8bpp
    If handleAlpha Then
        fi_DIB = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_NNQUANT, True)
        
        'We now need to find the palette index of a known transparent pixel
        Dim transpX As Long, transpY As Long
        tmpLayer.getTransparentLocation transpX, transpY
        
        Dim palIndex As Byte
        FreeImage_GetPixelIndex fi_DIB, transpX, transpY, palIndex
        
        'Request that FreeImage set that palette entry as the transparent index
        FreeImage_SetTransparentIndex fi_DIB, palIndex
        
        'Finally, because some software may not display the transparency correctly, we need to set that
        ' palette index to some normal color instead of bright magenta.  To do that, we must make a
        ' copy of the palette and update the transparency index accordingly.
        Dim fi_Palette() As Long
        fi_Palette = FreeImage_GetPaletteExLong(fi_DIB)
        
        fi_Palette(palIndex) = tmpLayer.getOriginalTransparentColor()
        
    End If
    
    'Use that handle to save the image to GIF format, with required 8bpp (256 color) conversion
    If fi_DIB <> 0 Then
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, GIFPath, FIF_GIF, , FICD_8BPP, , , , , True)
        If fi_Check = False Then
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            FreeLibrary hLib
            SaveGIFImage = False
            Exit Function
        Else
            Message "%1 save complete.", sFileType
        End If
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SaveGIFImage = False
        FreeLibrary hLib
        Exit Function
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
    SaveGIFImage = True
    Exit Function
    
SaveGIFError:

    If hLib <> 0 Then FreeLibrary hLib

    SaveGIFImage = False
    
End Function

'Save a PNG (Portable Network Graphic) file.  GDI+ can also do this.
Public Function SavePNGImage(ByRef srcPDImage As pdImage, ByVal PNGPath As String, ByVal outputColorDepth As Long, Optional ByVal pngParams As String = "") As Boolean

    On Error GoTo SavePNGError

    'Parse all possible PNG parameters
    ' (At present, three are possible: compression level, interlacing, BKGD chunk preservation (background color)
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(pngParams) > 0 Then cParams.setParamString pngParams
    Dim pngCompressionLevel As Long
    pngCompressionLevel = cParams.GetLong(1, 9)
    Dim pngUseInterlacing As Boolean, pngPreserveBKGD As Boolean
    pngUseInterlacing = cParams.GetBool(2, False)
    pngPreserveBKGD = cParams.GetBool(3, False)

    Dim sFileType As String
    sFileType = "PNG"

    'Make sure we found the plug-in when we loaded the program
    If g_ImageFormats.FreeImageEnabled = False Then
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SavePNGImage = False
        Exit Function
    End If
    
    'Before doing anything else, make a special note of the outputColorDepth.  If it is 8bpp, we will use pngnq-s9 to help with the save.
    Dim output8BPP As Boolean
    If outputColorDepth = 8 Then output8BPP = True Else output8BPP = False
    
    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
    
    Message "Preparing %1 image...", sFileType
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer srcPDImage.mainLayer
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (srcPDImage.mainLayer.getLayerColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'PhotoDemon now offers pngnq support via a plugin.  It can be used to render extremely high-quality 8bpp PNG files
        ' with "full" transparency.  If the pngnq-s9 executable is available, the export process is a bit different.
        
        'Before we can send stuff off to pngnq, however, we need to see if the image has more than 256 colors.  If it
        ' doesn't, we can save the file without pngnq's help.
        
        'Check to see if the current image had its colors counted before coming here.  If not, count it.
        Dim numColors As Long
        If g_LastImageScanned <> srcPDImage.imageID Then
            numColors = getQuickColorCount(srcPDImage, srcPDImage.imageID)
        Else
            numColors = g_LastColorCount
        End If
        
        'Pngnq can handle all types of transparency for us.  If pngnq cannot be found, we must rely on our own routines.
        If Not g_ImageFormats.pngnqEnabled Then
        
            'Does this layer contain binary transparency?  If so, mark all transparent pixels with magic magenta.
            If tmpLayer.isAlphaBinary Then
                tmpLayer.applyAlphaCutoff
            Else
            
                'If we are in the midst of a batch conversion, we don't want to bother the user with alpha dialogs.
                ' Thus, use a default cut-off of 127 and continue on.
                If MacroStatus = MacroBATCH Then
                    tmpLayer.applyAlphaCutoff
                
                'We're not in a batch conversion, so ask the user which cut-off they would like to use.
                Else
            
                    Dim alphaCheck As VbMsgBoxResult
                    alphaCheck = promptAlphaCutoff(tmpLayer)
                    
                    'If the alpha dialog is canceled, abandon the entire save
                    If alphaCheck = vbCancel Then
                    
                        tmpLayer.eraseLayer
                        Set tmpLayer = Nothing
                        SavePNGImage = False
                        Exit Function
                    
                    'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                    Else
                        tmpLayer.applyAlphaCutoff g_AlphaCutoff
                    End If
                
                End If
                
            End If
            
        'If pngnq is available, force the output to 32bpp.  Pngnq will take care of the actual 8bpp reduction.
        Else
            outputColorDepth = 32
        End If
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (srcPDImage.mainLayer.getLayerColorDepth = 32) And (outputColorDepth < 32) Then tmpLayer.compositeBackgroundColor 255, 255, 255
    
        'Also, if pngnq is enabled, we will use that for the transformation - so we need to reset the outgoing color depth to 24bpp
        If (srcPDImage.mainLayer.getLayerColorDepth = 24) And (outputColorDepth = 8) And g_ImageFormats.pngnqEnabled Then outputColorDepth = 24
    
    End If
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'If the image is being reduced from some higher bit-depth to 1bpp, manually force a conversion with dithering
    If outputColorDepth = 1 Then fi_DIB = FreeImage_Dither(fi_DIB, FID_FS)
    
    'If the image contains alpha and pngnq is not available, we need to manually convert the FreeImage copy of the image to 8bpp.
    ' Then we need to apply alpha using the cut-off established earlier in this section.
    If handleAlpha And (Not g_ImageFormats.pngnqEnabled) Then
        fi_DIB = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_NNQUANT, True)
        
        'We now need to find the palette index of a known transparent pixel
        Dim transpX As Long, transpY As Long
        tmpLayer.getTransparentLocation transpX, transpY
        
        Dim palIndex As Byte
        FreeImage_GetPixelIndex fi_DIB, transpX, transpY, palIndex
        
        'Request that FreeImage set that palette entry as the transparent index
        FreeImage_SetTransparentIndex fi_DIB, palIndex
        
        'Finally, because some software may not display the transparency correctly, we need to set that
        ' palette index to its original value.  To do that, we must make a copy of the palette and update
        ' the transparency index accordingly.
        Dim fi_Palette() As Long
        fi_Palette = FreeImage_GetPaletteExLong(fi_DIB)
        
        fi_Palette(palIndex) = tmpLayer.getOriginalTransparentColor()
        
    End If
        
    'Use that handle to save the image to PNG format
    If fi_DIB <> 0 Then
    
        'Embed a background color if available, and the user has requested it.
        If pngPreserveBKGD And srcPDImage.pngBackgroundColor <> -1 Then
            
            Dim rQuad As RGBQUAD
            rQuad.rgbRed = ExtractR(srcPDImage.pngBackgroundColor)
            rQuad.rgbGreen = ExtractG(srcPDImage.pngBackgroundColor)
            rQuad.rgbBlue = ExtractB(srcPDImage.pngBackgroundColor)
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
        fi_Check = FreeImage_SaveEx(fi_DIB, PNGPath, FIF_PNG, PNGFlags, outputColorDepth, , , , , True)
        
        If fi_Check = False Then
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            FreeLibrary hLib
            SavePNGImage = False
            Exit Function
        
        Else
            
            'If pngnq is being used to help with the 8bpp reduction, now is when we need to use it.
            If g_ImageFormats.pngnqEnabled And output8BPP Then
            
                'Build a full shell path for the pngnq operation
                Dim shellPath As String
                shellPath = g_PluginPath & "pngnq-s9.exe "
                
                'Force overwrite if a file with that name already exists
                shellPath = shellPath & "-f "
                
                'Display verbose status messages (consider removing this for production build)
                shellPath = shellPath & "-v "
                
                'Turn off the alpha importance heuristic (this leads to better results on semi-transparent images, and improves
                ' processing time for 24bpp images)
                shellPath = shellPath & "-A "
                
                'Now, add options that the user may have specified.
                
                'Alpha extenuation (only relevant for 32bpp images)
                If srcPDImage.mainLayer.getLayerColorDepth = 32 Then
                    If g_UserPreferences.GetPref_Boolean("Plugins", "Pngnq Alpha Extenuation", False) Then
                        shellPath = shellPath & "-t15 "
                    Else
                        shellPath = shellPath & "-t0 "
                    End If
                End If
        
                'YUV
                If g_UserPreferences.GetPref_Boolean("Plugins", "Pngnq YUV", True) Then
                    shellPath = shellPath & "-Cy "
                Else
                    shellPath = shellPath & "-Cr "
                End If
        
                'Color sample size
                shellPath = shellPath & "-s" & g_UserPreferences.GetPref_Long("Plugins", "Pngnq Color Sample", 3) & " "
        
                'Dithering
                If g_UserPreferences.GetPref_Long("Plugins", "Pngnq Dithering", 5) = 0 Then
                    shellPath = shellPath & "-Qn "
                Else
                    shellPath = shellPath & "-Q" & g_UserPreferences.GetPref_Long("Plugins", "Pngnq Dithering", 5) & " "
                End If
                
                'Append the name of the current image
                shellPath = shellPath & """" & PNGPath & """"
                
                'Use pngnq to create a new file
                Message "Using the pngnq-s9 plugin to write a high-quality 8bpp PNG file.  This may take a moment..."
                
                'Before launching the shell, launch a single DoEvents.  This gives us some leeway before Windows marks the program
                ' as unresponsive...
                DoEvents
                
                Dim shellCheck As Boolean
                shellCheck = ShellAndWait(shellPath, vbMinimizedNoFocus)
            
                'If the shell was successful and the image was created successfully, overwrite the original 32bpp save
                ' (from FreeImage) with the new 8bpp one (from pngnq-s9)
                If shellCheck Then
                
                    Message "Pngnq-s9 transformation complete.  Verifying output..."
                
                    'pngnq is going to create a new file with the name "filename-nq8.png".  We need to rename that file
                    ' to whatever name the user supplied
                    Dim srcFile As String
                    srcFile = PNGPath
                    StripOffExtension srcFile
                    srcFile = srcFile & "-nq8.png"
                    
                    'Make sure both FreeImage and pngnq were able to generate valid files, then rewrite the FreeImage one
                    ' with the pngnq one.
                    If FileExist(srcFile) And FileExist(PNGPath) Then
                        Kill PNGPath
                        FileCopy srcFile, PNGPath
                        Kill srcFile
                    Else
                        Message "Pngnq-s9 could not write file.  Saving 32bpp image instead..."
                    End If
                Else
                    Message "Pngnq-s9 could not write file.  Saving 32bpp image instead..."
                End If
            
            End If
            
            Message "%1 save complete.", sFileType
        End If
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SavePNGImage = False
        FreeLibrary hLib
        Exit Function
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    SavePNGImage = True
    Exit Function
    
SavePNGError:

    If hLib <> 0 Then FreeLibrary hLib

    SavePNGImage = False
    
End Function

'Save a PPM (Portable Pixmap) image
Public Function SavePPMImage(ByRef srcPDImage As pdImage, ByVal PPMPath As String, Optional ByVal ppmParams As String = "") As Boolean

    On Error GoTo SavePPMError

    'Parse all possible PPM parameters (at present there is only one possible parameter, which sets RAW vs ASCII encoding)
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(ppmParams) > 0 Then cParams.setParamString ppmParams
    Dim ppmFormat As Long
    ppmFormat = cParams.GetLong(1, 0)

    Dim sFileType As String
    sFileType = "PPM"

    'Make sure we found the plug-in when we loaded the program
    If g_ImageFormats.FreeImageEnabled = False Then
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SavePPMImage = False
        Exit Function
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
    
    Message "Preparing %1 image...", sFileType
    
    'Based on the user's preference, select binary or ASCII encoding for the PPM file
    Dim ppm_Encoding As FREE_IMAGE_SAVE_OPTIONS
    If ppmFormat = 0 Then ppm_Encoding = FISO_PNM_SAVE_RAW Else ppm_Encoding = FISO_PNM_SAVE_ASCII
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer srcPDImage.mainLayer
    If tmpLayer.getLayerColorDepth = 32 Then tmpLayer.convertTo24bpp
        
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
        
    'Use that handle to save the image to PPM format (ASCII)
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, PPMPath, FIF_PPM, ppm_Encoding, FICD_24BPP, , , , , True)
        If fi_Check = False Then
            Message "PPM save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
            FreeLibrary hLib
            SavePPMImage = False
            Exit Function
        Else
            Message "%1 save complete.", sFileType
        End If
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        FreeLibrary hLib
        SavePPMImage = False
        Exit Function
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
    SavePPMImage = True
    Exit Function
        
SavePPMError:

    If hLib <> 0 Then FreeLibrary hLib
    SavePPMImage = False
        
End Function

'Save to Targa (TGA) format.
Public Function SaveTGAImage(ByRef srcPDImage As pdImage, ByVal TGAPath As String, ByVal outputColorDepth As Long, Optional ByVal tgaParams As String = "") As Boolean
    
    On Error GoTo SaveTGAError
    
    'Parse all possible TGA parameters (at present there is only one possible parameter, which specifies RLE compression)
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(tgaParams) > 0 Then cParams.setParamString tgaParams
    Dim TGACompression As Boolean
    TGACompression = cParams.GetBool(1, False)
    
    Dim sFileType  As String
    sFileType = "TGA"
    
    'Make sure we found the plug-in when we loaded the program
    If g_ImageFormats.FreeImageEnabled = False Then
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveTGAImage = False
        Exit Function
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
    
    Message "Preparing %1 image...", sFileType
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer srcPDImage.mainLayer
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (srcPDImage.mainLayer.getLayerColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'Does this layer contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If tmpLayer.isAlphaBinary Then
            tmpLayer.applyAlphaCutoff
        Else
        
            'If we are in the midst of a batch conversion, we don't want to bother the user with alpha dialogs.
            ' Thus, use a default cut-off of 127 and continue on.
            If MacroStatus = MacroBATCH Then
                tmpLayer.applyAlphaCutoff
            
            'We're not in a batch conversion, so ask the user which cut-off they would like to use.
            Else
            
                Dim alphaCheck As VbMsgBoxResult
                alphaCheck = promptAlphaCutoff(tmpLayer)
                
                'If the alpha dialog is canceled, abandon the entire save
                If alphaCheck = vbCancel Then
                
                    tmpLayer.eraseLayer
                    Set tmpLayer = Nothing
                    SaveTGAImage = False
                    Exit Function
                
                'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                Else
                    tmpLayer.applyAlphaCutoff g_AlphaCutoff
                End If
            End If
            
        End If
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (srcPDImage.mainLayer.getLayerColorDepth = 32) And (outputColorDepth < 32) Then tmpLayer.compositeBackgroundColor 255, 255, 255
    
    End If
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'If the image contains alpha, we need to convert the FreeImage copy of the image to 8bpp
    If handleAlpha Then
        fi_DIB = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_NNQUANT, True)
        
        'We now need to find the palette index of a known transparent pixel
        Dim transpX As Long, transpY As Long
        tmpLayer.getTransparentLocation transpX, transpY
        
        Dim palIndex As Byte
        FreeImage_GetPixelIndex fi_DIB, transpX, transpY, palIndex
        
        'Request that FreeImage set that palette entry as the transparent index
        FreeImage_SetTransparentIndex fi_DIB, palIndex
        
        'Finally, because some software may not display the transparency correctly, we need to set that
        ' palette index to some normal color instead of bright magenta.  To do that, we must make a
        ' copy of the palette and update the transparency index accordingly.
        Dim fi_Palette() As Long
        fi_Palette = FreeImage_GetPaletteExLong(fi_DIB)
        
        fi_Palette(palIndex) = tmpLayer.getOriginalTransparentColor()
        
    End If
    
    'Finally, prepare a TGA save flag.  If the user has requested RLE encoding, pass that along to FreeImage.
    Dim TGAflags As Long
    TGAflags = TARGA_DEFAULT
            
    If TGACompression Then TGAflags = TARGA_SAVE_RLE
            
    
    'Use that handle to save the image to TGA format
    If fi_DIB <> 0 Then
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TGAPath, FIF_TARGA, TGAflags, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            FreeLibrary hLib
            SaveTGAImage = False
            Exit Function
        Else
            Message "%1 save complete.", sFileType
        End If
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        FreeLibrary hLib
        SaveTGAImage = False
        Exit Function
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
    SaveTGAImage = True
    Exit Function
    
SaveTGAError:

    If hLib <> 0 Then FreeLibrary hLib
    SaveTGAImage = False

End Function

'Save to JPEG using the FreeImage library.  FreeImage offers significantly more JPEG features than GDI+.
Public Function SaveJPEGImage(ByRef srcPDImage As pdImage, ByVal JPEGPath As String, ByVal jpegParams As String) As Boolean
    
    On Error GoTo SaveJPEGError
        
    'Parse all possible JPEG parameters
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(jpegParams) > 0 Then cParams.setParamString jpegParams
    Dim JPEGFlags As Long
    JPEGFlags = cParams.GetLong(1, 92)
    
    'If FreeImage is not available, fall back to GDI+.  If that is not available, fail the function.
    If Not g_ImageFormats.FreeImageEnabled Then
        If g_ImageFormats.GDIPlusEnabled Then
            SaveJPEGImage = GDIPlusSavePicture(srcPDImage, JPEGPath, ImageJPEG, 24, JPEGFlags)
        Else
            SaveJPEGImage = False
            Message "No %1 encoder found. Save aborted.", "JPEG"
        End If
        Exit Function
    End If
    
    Dim sFileType As String
    sFileType = "JPEG"
    
    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
    
    Message "Preparing %1 image...", sFileType
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer srcPDImage.mainLayer
    If tmpLayer.getLayerColorDepth = 32 Then tmpLayer.convertTo24bpp
        
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
        
    'If the image is grayscale, instruct FreeImage to internally mark the image as such
    Dim outputColorDepth As Long
    Message "Analyzing image color content..."
    If tmpLayer.isLayerGrayscale Then
        Message "No color found.  Saving 8bpp grayscale JPEG."
        outputColorDepth = 8
        fi_DIB = FreeImage_ConvertToGreyscale(fi_DIB)
    Else
        Message "Color found.  Saving 24bpp full-color JPEG."
        outputColorDepth = 24
    End If
        
    'Combine all received flags into one
    JPEGFlags = JPEGFlags Or cParams.GetLong(2, 0)
    
    'If a thumbnail has been requested, generate that now
    If cParams.GetLong(3, 0) <> 0 Then
    
        'Create the thumbnail using default settings (100x100px)
        Dim fThumbnail As Long
        fThumbnail = FreeImage_MakeThumbnail(fi_DIB, 100)
        
        'Embed the thumbnail into the main DIB
        FreeImage_SetThumbnail fi_DIB, fThumbnail
        
        'Erase the thumbnail
        FreeImage_Unload fThumbnail
        
    End If
        
    'Use that handle to save the image to JPEG format
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, JPEGPath, FIF_JPEG, JPEGFlags, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            FreeLibrary hLib
            SaveJPEGImage = False
            Exit Function
        Else
            Message "%1 save complete.", sFileType
        End If
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        FreeLibrary hLib
        SaveJPEGImage = False
        Exit Function
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
    SaveJPEGImage = True
    Exit Function
    
SaveJPEGError:

    If hLib <> 0 Then FreeLibrary hLib
    SaveJPEGImage = False
    
End Function

'Save a TIFF (Tagged Image File Format) image via FreeImage.  GDI+ can also do this.
Public Function SaveTIFImage(ByRef srcPDImage As pdImage, ByVal TIFPath As String, ByVal outputColorDepth As Long, Optional ByVal tiffParams As String = "") As Boolean
    
    On Error GoTo SaveTIFError
    
    'Parse all possible TIFF parameters
    ' (At present, two are possible: one for compression type, and another for CMYK encoding)
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(tiffParams) > 0 Then cParams.setParamString tiffParams
    Dim tiffEncoding As Long
    tiffEncoding = cParams.GetLong(1, 0)
    Dim tiffUseCMYK As Boolean
    tiffUseCMYK = cParams.GetBool(2, False)
    
    Dim sFileType As String
    sFileType = "TIFF"
    
    'Make sure we found the plug-in when we loaded the program
    If g_ImageFormats.FreeImageEnabled = False Then
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveTIFImage = False
        Exit Function
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
    
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
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer srcPDImage.mainLayer
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (srcPDImage.mainLayer.getLayerColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'Does this layer contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If tmpLayer.isAlphaBinary Then
            tmpLayer.applyAlphaCutoff
        Else
            
            'If we are in the midst of a batch conversion, we don't want to bother the user with alpha dialogs.
            ' Thus, use a default cut-off of 127 and continue on.
            If MacroStatus = MacroBATCH Then
                tmpLayer.applyAlphaCutoff
            
            'We're not in a batch conversion, so ask the user which cut-off they would like to use.
            Else
                Dim alphaCheck As VbMsgBoxResult
                alphaCheck = promptAlphaCutoff(tmpLayer)
                
                'If the alpha dialog is canceled, abandon the entire save
                If alphaCheck = vbCancel Then
                
                    tmpLayer.eraseLayer
                    Set tmpLayer = Nothing
                    SaveTIFImage = False
                    Exit Function
                
                'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                Else
                    tmpLayer.applyAlphaCutoff g_AlphaCutoff
                End If
            End If
            
        End If
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (srcPDImage.mainLayer.getLayerColorDepth = 32) And (outputColorDepth < 32) Then tmpLayer.compositeBackgroundColor 255, 255, 255
    
    End If
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'If the image is being reduced from some higher bit-depth to 1bpp, manually force a conversion with dithering
    If outputColorDepth = 1 Then fi_DIB = FreeImage_Dither(fi_DIB, FID_FS)
    
    'If the image contains alpha, we need to convert the FreeImage copy of the image to 8bpp
    If handleAlpha Then
        fi_DIB = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_NNQUANT, True)
        
        'We now need to find the palette index of a known transparent pixel
        Dim transpX As Long, transpY As Long
        tmpLayer.getTransparentLocation transpX, transpY
        
        Dim palIndex As Byte
        FreeImage_GetPixelIndex fi_DIB, transpX, transpY, palIndex
        
        'Request that FreeImage set that palette entry as the transparent index
        FreeImage_SetTransparentIndex fi_DIB, palIndex
        
        'Finally, because some software may not display the transparency correctly, we need to set that
        ' palette index to some normal color instead of bright magenta.  To do that, we must make a
        ' copy of the palette and update the transparency index accordingly.
        Dim fi_Palette() As Long
        fi_Palette = FreeImage_GetPaletteExLong(fi_DIB)
        
        fi_Palette(palIndex) = tmpLayer.getOriginalTransparentColor()
        
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
            tmpLayer.convertToCMYK32
            fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
        End If
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TIFPath, FIF_TIFF, TIFFFlags, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            FreeLibrary hLib
            SaveTIFImage = False
            Exit Function
        Else
            Message "%1 save complete.", sFileType
        End If
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        FreeLibrary hLib
        SaveTIFImage = False
        Exit Function
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
    SaveTIFImage = True
    Exit Function
    
SaveTIFError:

    If hLib <> 0 Then FreeLibrary hLib
    SaveTIFImage = False
        
End Function

'Save to JPEG-2000 format using the FreeImage library.
Public Function SaveJP2Image(ByRef srcPDImage As pdImage, ByVal jp2Path As String, ByVal outputColorDepth As Long, Optional ByVal jp2Params As String = "") As Boolean
    
    On Error GoTo SaveJP2Error
    
    'Parse all possible JPEG-2000 params
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(jp2Params) > 0 Then cParams.setParamString jp2Params
    Dim jp2Quality As Long
    If cParams.doesParamExist(1) Then jp2Quality = cParams.GetLong(1) Else jp2Quality = 16
    
    Dim sFileType As String
    sFileType = "JPEG-2000"
    
    'Make sure we found the plug-in when we loaded the program
    If g_ImageFormats.FreeImageEnabled = False Then
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveJP2Image = False
        Exit Function
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
    
    Message "Preparing %1 image...", sFileType
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer srcPDImage.mainLayer
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (srcPDImage.mainLayer.getLayerColorDepth = 32) Then tmpLayer.convertTo24bpp
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'Use that handle to save the image to JPEG format
    If fi_DIB <> 0 Then
                
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, jp2Path, FIF_JP2, jp2Quality, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            FreeLibrary hLib
            SaveJP2Image = False
            Exit Function
        Else
            Message "%1 save complete.", sFileType
        End If
    Else
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        FreeLibrary hLib
        SaveJP2Image = False
        Exit Function
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
    SaveJP2Image = True
    Exit Function
    
SaveJP2Error:

    If hLib <> 0 Then FreeLibrary hLib
    SaveJP2Image = False
    
End Function

