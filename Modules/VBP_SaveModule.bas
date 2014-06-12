Attribute VB_Name = "Saving"
'***************************************************************************
'File Saving Interface
'Copyright ©2001-2014 by Tanner Helland
'Created: 4/15/01
'Last updated: 19/May/14
'Last update: final work on custom Undo/Redo saving
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


'This routine will blindly save the composited layer contents (from the pdImage object specified by srcPDImage) to dstPath.
' It is up to the calling routine to make sure this is what is wanted. (Note: this routine will erase any existing image
' at dstPath, so BE VERY CAREFUL with what you send here!)
'
'INPUTS:
'   1) pdImage to be saved
'   2) Destination file path
'   3) Optional: imageID (if provided, the function can write information about the save to the relevant object in the pdImages array - this primarily exists for legacy reasons)
'   4) Optional: whether to display a form for the user to input additional save options (JPEG quality, etc)
'   5) Optional: a string of relevant save parameters.  If this is not provided, relevant parameters will be loaded from the preferences file.
Public Function PhotoDemon_SaveImage(ByRef srcPDImage As pdImage, ByVal dstPath As String, Optional ByVal imageID As Long = -1, Optional ByVal loadRelevantForm As Boolean = False, Optional ByVal saveParamString As String = "", Optional ByVal forceColorDepthMethod As Long = -1, Optional ByVal suspendMetadataActions As Boolean = False, Optional ByVal suspendMRUUpdating As Boolean = False) As Boolean
    
    'Only update the MRU list if 1) no form is shown (because the user may cancel it), 2) a form was shown and the user
    ' successfully navigated it, and 3) no errors occured during the export process.  By default, this is set to "do not update."
    Dim updateMRU As Boolean
    updateMRU = False
    
    'Start by determining the output format for this image (which was set either by a "Save As" common dialog box,
    ' or by copying the image's original format - or, if in the midst of a batch process, by the user via the batch wizard).
    Dim saveFormat As Long
    saveFormat = srcPDImage.currentFileFormat
    
    
    '****************************************************************************************************
    ' Determine exported color depth (for non-PDI formats)
    '****************************************************************************************************

    'The user is allowed to set a persistent preference for output color depth.  This setting affects a "color depth"
    ' parameter that will be sent to the various format-specific save file routines.  The available preferences are:
    ' 0) Mimic the file's original color depth (if available; this may not always be possible, e.g. saving a 32bpp PNG as JPEG)
    ' 1) Count the number of colors used, and save the file based on that (again, if possible)
    ' 2) Prompt the user for their desired export color depth
    '
    'Note that batch processing allows the user to overwrite their default preference with a specific preference for that
    ' batch process; if this occurs, the "forceColorDepthMethod" is utilized.
    
    Dim outputColorDepth As Long
    
    '100 is the magic number for saving PDI files (PhotoDemon's internal format).  PDI files do not need color depth checked,
    ' as the writer handles color depth independently for each layer.
    If saveFormat = 100 Then
        outputColorDepth = 32
        
    'The save format is not PDI.  Determine the ideal color depth.
    Else
    
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
                
                    'Retrieve a composited copy of the current image, which we will use to determine output color depth.
                    Dim tmpCompositeDIB As pdDIB
                    Set tmpCompositeDIB = New pdDIB
                    
                    srcPDImage.getCompositedImage tmpCompositeDIB, False
                    
                    'Validate the composited image's alpha channel; if it is pointless, we can request 24bpp output depth.
                    If Not tmpCompositeDIB.verifyAlphaChannel Then tmpCompositeDIB.convertTo24bpp
                    
                    'Count the number of colors in the image.  (The function will automatically cease if it hits 257 colors,
                    ' as anything above 256 colors is treated as 24bpp.)
                    Dim colorCountCheck As Long
                    Message "Counting image colors to determine optimal exported color depth..."
                    
                    If imageID <> -1 Then
                        colorCountCheck = getQuickColorCount(tmpCompositeDIB, imageID)
                    Else
                        colorCountCheck = getQuickColorCount(tmpCompositeDIB)
                    End If
                    
                    'If 256 or less colors were found in the image, mark it as 8bpp.  Otherwise, mark it as 24 or 32bpp.
                    outputColorDepth = getColorDepthFromColorCount(colorCountCheck, tmpCompositeDIB)
                    
                    'A special case arises when an image has <= 256 colors, but a non-binary alpha channel.  PNG allows for
                    ' this, but other formats do not.  Because even the PNG transformation is not lossless, set these types of
                    ' images to be exported as 32bpp.
                    If (outputColorDepth <= 8) And (tmpCompositeDIB.getDIBColorDepth = 32) Then
                        If (Not tmpCompositeDIB.isAlphaBinary) Then outputColorDepth = 32
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
    ' Based on the requested file type and color depth, call the appropriate save function
    '****************************************************************************************************
        
    Select Case saveFormat
        
        'JPEG
        Case FIF_JPEG
        
            'JPEG files may need to display a dialog box so the user can set compression quality
            If loadRelevantForm Then
                
                Dim gotSettings As VbMsgBoxResult
                gotSettings = promptJPEGSettings(srcPDImage)
                
                'If the dialog was canceled, note it.  Otherwise, remember that the user has seen the JPEG save screen at least once.
                If gotSettings = vbOK Then
                    srcPDImage.imgStorage.Item("hasSeenJPEGPrompt") = True
                    PhotoDemon_SaveImage = True
                Else
                    PhotoDemon_SaveImage = False
                    Message "Save canceled."
                    
                    Exit Function
                End If
                
                'If the user clicked OK, replace the function's save parameters with the ones set by the user
                cParams.setParamString Str(g_JPEGQuality)
                cParams.setParamString cParams.getParamString & "|" & Str(g_JPEGFlags)
                cParams.setParamString cParams.getParamString & "|" & Str(g_JPEGThumbnail)
                cParams.setParamString cParams.getParamString & "|" & Str(g_JPEGAutoQuality)
                cParams.setParamString cParams.getParamString & "|" & Str(g_JPEGAdvancedColorMatching)
                
            End If
            
            'Store these JPEG settings in the image object so we don't have to pester the user for it if they save again
            srcPDImage.saveParameters = cParams.getParamString
                            
            'I implement two separate save functions for JPEG images: FreeImage and GDI+.  GDI+ does not need to make a copy
            ' of the image before saving it - which makes it much faster - but FreeImage provides a number of additional
            ' parameters, like optimization, thumbnail embedding, and custom subsampling.  If no optional parameters are in use
            ' (or if FreeImage is unavailable), use GDI+.  Otherwise, use FreeImage.
            If g_ImageFormats.FreeImageEnabled And (cParams.doesParamExist(2) Or cParams.doesParamExist(3)) Then
                Screen.MousePointer = vbHourglass
                updateMRU = SaveJPEGImage(srcPDImage, dstPath, cParams.getParamString)
                Screen.MousePointer = vbDefault
            ElseIf g_ImageFormats.GDIPlusEnabled Then
                Screen.MousePointer = vbHourglass
                updateMRU = GDIPlusSavePicture(srcPDImage, dstPath, ImageJPEG, 24, cParams.GetLong(1, 92))
                Screen.MousePointer = vbDefault
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
                
                Exit Function
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
                cParams.setParamString Str(g_UserPreferences.GetPref_Long("File Formats", "PNG Compression", 9))
                cParams.setParamString cParams.getParamString() & "|" & Str(g_UserPreferences.GetPref_Boolean("File Formats", "PNG Interlacing", False))
                cParams.setParamString cParams.getParamString() & "|" & Str(g_UserPreferences.GetPref_Boolean("File Formats", "PNG Background Color", True))
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
            If Not cParams.doesParamExist(1) Then cParams.setParamString Str(g_UserPreferences.GetPref_Long("File Formats", "PPM Export Format", 0))
            updateMRU = SavePPMImage(srcPDImage, dstPath, cParams.getParamString)
                
        'TGA
        Case FIF_TARGA
            If Not cParams.doesParamExist(1) Then cParams.setParamString Str(g_UserPreferences.GetPref_Boolean("File Formats", "TGA RLE", False))
            updateMRU = SaveTGAImage(srcPDImage, dstPath, outputColorDepth, cParams.getParamString)
            
        'JPEG-2000
        Case FIF_JP2
        
            If loadRelevantForm Then
                
                Dim gotJP2Settings As VbMsgBoxResult
                gotJP2Settings = promptJP2Settings(srcPDImage)
                
                'If the dialog was canceled, note it.  Otherwise, remember that the user has seen the JPEG save screen at least once.
                If gotJP2Settings = vbOK Then
                    srcPDImage.imgStorage.Item("hasSeenJP2Prompt") = True
                    PhotoDemon_SaveImage = True
                Else
                    PhotoDemon_SaveImage = False
                    Message "Save canceled."
                    
                    Exit Function
                End If
                
                'If the user clicked OK, replace the functions save parameters with the ones set by the user
                cParams.setParamString Str(g_JP2Compression)
                
            End If
            
            'Store the JPEG-2000 quality in the image object so we don't have to pester the user for it if they save again
            srcPDImage.saveParameters = cParams.getParamString
        
            updateMRU = SaveJP2Image(srcPDImage, dstPath, outputColorDepth, cParams.getParamString)
            
        'TIFF
        Case FIF_TIFF
            
            'TIFFs use two parameters - compression type, and CMYK encoding (true/false)
            If Not cParams.doesParamExist(1) Then
                cParams.setParamString Str(g_UserPreferences.GetPref_Long("File Formats", "TIFF Compression", 0)) & "|" & Str(g_UserPreferences.GetPref_Boolean("File Formats", "TIFF CMYK", False))
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
        
        'WebP
        Case FIF_WEBP
        
            If loadRelevantForm Then
                
                Dim gotWebPSettings As VbMsgBoxResult
                gotWebPSettings = promptWebPSettings(srcPDImage)
                
                'If the dialog was canceled, note it.  Otherwise, remember that the user has seen the JPEG save screen at least once.
                If gotWebPSettings = vbOK Then
                    srcPDImage.imgStorage.Item("hasSeenWebPPrompt") = True
                    PhotoDemon_SaveImage = True
                Else
                    PhotoDemon_SaveImage = False
                    Message "Save canceled."
                    
                    Exit Function
                End If
                
                'If the user clicked OK, replace the functions save parameters with the ones set by the user
                cParams.setParamString Str(g_WebPCompression)
                
            End If
            
            'Store the JPEG-2000 quality in the image object so we don't have to pester the user for it if they save again
            srcPDImage.saveParameters = cParams.getParamString
        
            updateMRU = SaveWebPImage(srcPDImage, dstPath, outputColorDepth, cParams.getParamString)
        
        'JPEG XR
        Case FIF_JXR
        
            If loadRelevantForm Then
                
                Dim gotJXRSettings As VbMsgBoxResult
                gotJXRSettings = promptJXRSettings(srcPDImage)
                
                'If the dialog was canceled, note it.  Otherwise, remember that the user has seen the JPEG save screen at least once.
                If gotJXRSettings = vbOK Then
                    srcPDImage.imgStorage.Item("hasSeenJXRPrompt") = True
                    PhotoDemon_SaveImage = True
                Else
                    PhotoDemon_SaveImage = False
                    Message "Save canceled."
                    
                    Exit Function
                End If
                
                'If the user clicked OK, replace the functions save parameters with the ones set by the user
                cParams.setParamString Str(g_JXRCompression) & "|" & Str(g_JXRProgressive)
                
            End If
            
            'Store the JPEG-2000 quality in the image object so we don't have to pester the user for it if they save again
            srcPDImage.saveParameters = cParams.getParamString
        
            updateMRU = SaveJXRImage(srcPDImage, dstPath, outputColorDepth, cParams.getParamString)
            
        'Anything else must be a bitmap
        Case FIF_BMP
            
            'If the user has not provided explicit BMP parameters, load their default values from the preferences file
            If Not cParams.doesParamExist(1) Then cParams.setParamString Str(g_UserPreferences.GetPref_Boolean("File Formats", "Bitmap RLE", False))
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
    ' through an external plugin.  Exiftool requires an existant file to be used as a target, and an existant metadata file
    ' to be used as its source.
    
    'Note that updateMRU is used to track save file success, so it will only be TRUE if the image file was written successfully.
    ' If the file was not written successfully, abandon any attempts at metadata embedding.
    If updateMRU And g_ExifToolEnabled And (Not suspendMetadataActions) Then
        
        'Only attempt to export metadata if ExifTool was able to successfully cache and parse metadata prior to saving
        If srcPDImage.imgMetadata.hasXMLMetadata Then
            updateMRU = srcPDImage.imgMetadata.writeAllMetadata(dstPath, srcPDImage)
        Else
            Message "No metadata to export.  Continuing save..."
        End If
                
    End If
    
    'UpdateMRU should only be true if the save was successful
    If updateMRU And (Not suspendMRUUpdating) Then
    
        'Additionally, only add this MRU to the list (and generate an accompanying icon) if we are not in the midst
        ' of a batch conversion.
        If MacroStatus <> MacroBATCH Then
        
            'Add this file to the MRU list
            g_RecentFiles.MRU_AddNewFile dstPath, srcPDImage
        
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
            If Not srcPDImage.forInternalUseOnly Then syncInterfaceToCurrentImage
                        
            'Notify the thumbnail window that this image has been updated (so it can show/hide the save icon)
            If Not srcPDImage.forInternalUseOnly Then toolbar_ImageTabs.notifyUpdatedImage srcPDImage.imageID
            
        End If
    
    Else
        
        'If we aren't updating the MRU, something went wrong.  Display that the save was canceled and exit.
        ' (One exception to this is if the user requested us to not update the MRU; in this case, there is no error!)
        If Not suspendMRUUpdating Then
            Message "Save canceled."
            PhotoDemon_SaveImage = False
            Exit Function
        End If
        
    End If

    If Not suspendMRUUpdating Then Message "Save complete."

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
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.getCompositedImage tmpImageCopy, False
    
    'If the output color depth is 24 or 32bpp, or if both GDI+ and FreeImage are missing, use our own internal methods
    ' to save the BMP file.
    If ((outputColorDepth = 24) And (tmpImageCopy.getDIBColorDepth = 24)) Or ((outputColorDepth = 32) And (tmpImageCopy.getDIBColorDepth = 32)) Or ((Not g_ImageFormats.GDIPlusEnabled) And (Not g_ImageFormats.FreeImageEnabled)) Then
    
        Message "Saving %1 file...", sFileType
    
        'The DIB class is capable of doing this without any outside help.
        tmpImageCopy.writeToBitmapFile BMPPath
    
        Message "%1 save complete.", sFileType
        
    'If some other color depth is specified, use FreeImage or GDI+ to write the file
    Else
    
        If g_ImageFormats.FreeImageEnabled Then
            
            Message "Preparing %1 image...", sFileType
            
            'Copy the image into a temporary DIB
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB tmpImageCopy
            
            'If the output color depth is 24 but the current image is 32, composite the image against a white background
            If (outputColorDepth < 32) And (tmpDIB.getDIBColorDepth = 32) Then tmpDIB.convertTo24bpp
            
            'Convert our current DIB to a FreeImage-type DIB
            Dim fi_DIB As Long
            fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
                        
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
                
                If fi_Check Then
                    Message "%1 save complete.", sFileType
                Else
                    Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
                    SaveBMP = False
                    Exit Function
                End If
                
            Else
            
                Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
                SaveBMP = False
                Exit Function
                
            End If
            
        Else
            GDIPlusSavePicture srcPDImage, BMPPath, ImageBMP, outputColorDepth
        End If
    
    End If
    
    SaveBMP = True
    Exit Function
    
SaveBMPError:

    SaveBMP = False
    
End Function

'Save the current image to PhotoDemon's native PDI format
' TODO:
'  - Add support for storing a PNG copy of the fully composited image, preferably in the data chunk of the first node.
'  - Figure out a good way to store metadata; the problem is not so much storing the metadata itself, but storing any user edits.
'    I have postponed this until I get metadata editing working more fully.
'  - User-settable options for compression.  Some users may prefer extremely tight compression, at a trade-off of slower
'    image saves.  Similarly, compressing layers in PNG format instead of as a blind zLib stream would probably yield better
'    results (at a trade-off to performance).
'  - An option for compressing both the directory, and the completed data stream.  This would all take place in the pdPackage
'    class, and the directory would be decompressed automatically at load-time, so calling functions wouldn't need to worry
'    about catering to that class.  This could be helpful for further shrinking file size, especially for small images where
'    the header represents a larger portion of the file.
'  - Any number of other options might be helpful (e.g. password encryption, etc).  I should probably add a page about the PDI
'    format to the help documentation, where various ideas for future additions could be tracked.
Public Function SavePhotoDemonImage(ByRef srcPDImage As pdImage, ByVal PDIPath As String, Optional ByVal suppressMessages As Boolean = False, Optional ByVal compressHeaders As Boolean = True, Optional ByVal compressLayers As Boolean = True, Optional ByVal embedChecksums As Boolean = True, Optional ByVal writeHeaderOnlyFile As Boolean = False) As Boolean
    
    On Error GoTo SavePDIError
    
    Dim sFileType As String
    sFileType = "PDI"
    
    If Not suppressMessages Then Message "Saving %1 image...", sFileType
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of compressing individual layers,
    ' and storing everything to a running byte stream.
    Dim pdiWriter As pdPackager
    Set pdiWriter = New pdPackager
    If g_ZLibEnabled Then pdiWriter.init_ZLib g_PluginPath & "zlibwapi.dll"
    
    'When creating the actual package, we specify numOfLayers + 1 nodes.  The +1 is for the pdImage header itself, which
    ' gets its own node, separate from the individual layer nodes.
    pdiWriter.prepareNewPackage srcPDImage.getNumOfLayers + 1, PD_IMAGE_IDENTIFIER
    
    'We will use a temporary array for two main purposes: storing the byte equivalent of various XML strings returned by
    ' PhotoDemon objects, and for storing temporary copies of byte arrays of binary DIB data.
    Dim tmpData() As Byte
    
    'The first node we'll add is the pdImage header, in XML format.
    Dim nodeIndex As Long
    nodeIndex = pdiWriter.addNode("pdImage Header", -1, 0)
    
    Dim dataString As String
    srcPDImage.writeExternalData dataString, True
    
    pdiWriter.addNodeDataFromString nodeIndex, True, dataString, compressHeaders, , embedChecksums
    
    'The pdImage header only requires one of the two buffers in its node; the other can be happily left blank.
    
    'Next, we will add each pdLayer object to the stream.  This is done in two steps:
    ' 1) First, obtain the layer header in XML format and write it out
    ' 2) Second, obtain the current layer DIB (raw data only, no header!) and write it out
    Dim layerXMLHeader As String
    Dim layerDIBCopy() As Byte
    
    Dim i As Long
    For i = 0 To srcPDImage.getNumOfLayers - 1
    
        'Create a new node for this layer.  Note that the index is stored directly in the node name ("pdLayer (n)")
        ' while the layerID is stored as the nodeID.
        nodeIndex = pdiWriter.addNode("pdLayer " & i, srcPDImage.getLayerByIndex(i).getLayerID, 1)
        
        'Retrieve the layer header and add it to the header section of this node
        layerXMLHeader = srcPDImage.getLayerByIndex(i).getLayerHeaderAsXML(True)
        pdiWriter.addNodeDataFromString nodeIndex, True, layerXMLHeader, compressHeaders, , embedChecksums
        
        'If this is not a header-only file, retrieve the layer's DIB and add it to the data section of this node
        If Not writeHeaderOnlyFile Then
            srcPDImage.getLayerByIndex(i).layerDIB.copyImageBytesIntoStream layerDIBCopy
            pdiWriter.addNodeData nodeIndex, False, layerDIBCopy, compressLayers, , embedChecksums
        End If
    
    Next i
    
    'That's all there is to it!  Write the completed pdPackage out to file.
    SavePhotoDemonImage = pdiWriter.writePackageToFile(PDIPath)
    
    If Not suppressMessages Then Message "%1 save complete.", sFileType
    
    Exit Function
    
SavePDIError:

    SavePhotoDemonImage = False
    
End Function

'Save the requested layer to a variant of PhotoDemon's native PDI format.  Because this function is internal (it is used by the
' Undo/Redo engine only), it is not as fleshed-out as the actual SavePhotoDemonImage function.
Public Function SavePhotoDemonLayer(ByRef srcLayer As pdLayer, ByVal PDIPath As String, Optional ByVal suppressMessages As Boolean = False, Optional ByVal compressHeaders As Boolean = True, Optional ByVal compressLayers As Boolean = True, Optional ByVal embedChecksums As Boolean = True, Optional ByVal writeHeaderOnlyFile As Boolean = False) As Boolean
    
    On Error GoTo SavePDLayerError
    
    Dim sFileType As String
    sFileType = "PDI"
    
    If Not suppressMessages Then Message "Saving %1 layer...", sFileType
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of assembling the layer file.
    Dim pdiWriter As pdPackager
    Set pdiWriter = New pdPackager
    If g_ZLibEnabled Then pdiWriter.init_ZLib g_PluginPath & "zlibwapi.dll"
    
    'Unlike an actual PDI file, which stores a whole bunch of images, these temp layer files only have two pieces of data:
    ' the layer header, and the DIB bytestream.  Thus, we know there will only be 1 node required.
    pdiWriter.prepareNewPackage 1, PD_LAYER_IDENTIFIER
    
    'We will use a temporary array for two main purposes: storing the byte equivalent of various XML strings returned by
    ' PhotoDemon objects, and for storing temporary copies of byte arrays of binary DIB data.
    Dim tmpData() As Byte
    
    'The first (and only) node we'll add is the specific pdLayer header and DIB data.
    ' To help us reconstruct the node later, we also note the current layer's ID (stored as the node ID)
    '  and the current layer's index (stored as the node type).
    
    'Start by creating the node entry; if successful, this will return the index of the node, which we can use
    ' to supply the actual header and DIB data.
    Dim nodeIndex As Long
    nodeIndex = pdiWriter.addNode("pdLayer", srcLayer.getLayerID, pdImages(g_CurrentImage).getLayerIndexFromID(srcLayer.getLayerID))
    
    'Retrieve the layer header (in XML format), then write the XML stream to the pdPackage instance
    Dim dataString As String
    dataString = srcLayer.getLayerHeaderAsXML(True)
    
    pdiWriter.addNodeDataFromString nodeIndex, True, dataString, compressHeaders, , embedChecksums
    
    'If this is not a header-only request, retrieve the layer DIB (as a byte array), then copy the array
    ' into the pdPackage instance
    If Not writeHeaderOnlyFile Then
    
        Dim layerDIBCopy() As Byte
        srcLayer.layerDIB.copyImageBytesIntoStream layerDIBCopy
        pdiWriter.addNodeData nodeIndex, False, layerDIBCopy, compressLayers, , embedChecksums
        
    End If
    
    'That's all there is to it!  Write the completed pdPackage out to file.
    SavePhotoDemonLayer = pdiWriter.writePackageToFile(PDIPath)
    
    If Not suppressMessages Then Message "%1 save complete.", sFileType
    
    Exit Function
    
SavePDLayerError:

    SavePhotoDemonLayer = False
    
End Function

'Save a GIF (Graphics Interchange Format) image.  GDI+ can also do this.
Public Function SaveGIFImage(ByRef srcPDImage As pdImage, ByVal GIFPath As String, Optional ByVal forceAlphaConvert As Long = -1) As Boolean

    On Error GoTo SaveGIFError
    
    Dim sFileType As String
    sFileType = "GIF"

    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveGIFImage = False
        Exit Function
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.getCompositedImage tmpDIB, False
    
    'If the current image is 32bpp, we will need to apply some additional actions to the image to prepare the
    ' transparency.  Mark a bool value, because we will reference it in multiple places throughout the save function.
    Dim handleAlpha As Boolean
    If tmpDIB.getDIBColorDepth = 32 Then handleAlpha = True Else handleAlpha = False
    
    'If the current image contains transparency, we need to modify it in order to retain the alpha channel.
    If handleAlpha Then
    
        'Does this DIB contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If tmpDIB.isAlphaBinary Then
            tmpDIB.applyAlphaCutoff
        Else
            If forceAlphaConvert = -1 Then
                Dim alphaCheck As VbMsgBoxResult
                alphaCheck = promptAlphaCutoff(tmpDIB)
                
                'If the alpha dialog is canceled, abandon the entire save
                If alphaCheck = vbCancel Then
                
                    tmpDIB.eraseDIB
                    Set tmpDIB = Nothing
                    SaveGIFImage = False
                    Exit Function
                
                'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                Else
                    tmpDIB.applyAlphaCutoff g_AlphaCutoff, , g_AlphaCompositeColor
                End If
                
                'If the user decided to completely remove the image's alpha values, change handleAlpha to FALSE
                If g_AlphaCutoff = 0 Then handleAlpha = False
                
            Else
                tmpDIB.applyAlphaCutoff forceAlphaConvert
            End If
            
        End If
    
    End If
    
    Message "Writing %1 file...", sFileType
    
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
        
        fi_Palette(palIndex) = tmpDIB.getOriginalTransparentColor()
        
    End If
    
    'Use that handle to save the image to GIF format, with required 8bpp (256 color) conversion
    If fi_DIB <> 0 Then
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, GIFPath, FIF_GIF, , FICD_8BPP, , , , , True)
        
        If fi_Check Then
            Message "%1 save complete.", sFileType
        Else
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            SaveGIFImage = False
            Exit Function
        End If
        
    Else
    
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SaveGIFImage = False
        Exit Function
        
    End If
    
    SaveGIFImage = True
    Exit Function
    
SaveGIFError:

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
    If Not g_ImageFormats.FreeImageEnabled Then
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SavePNGImage = False
        Exit Function
    End If
    
    'Before doing anything else, make a special note of the outputColorDepth.  If it is 8bpp, we will use pngnq-s9 to help with the save.
    Dim output8BPP As Boolean
    If outputColorDepth = 8 Then output8BPP = True Else output8BPP = False
        
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.getCompositedImage tmpDIB, False
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'PhotoDemon now offers pngnq support via a plugin.  It can be used to render extremely high-quality 8bpp PNG files
        ' with "full" transparency.  If the pngnq-s9 executable is available, the export process is a bit different.
        
        'Before we can send stuff off to pngnq, however, we need to see if the image has more than 256 colors.  If it
        ' doesn't, we can save the file without pngnq's help.
        
        'Check to see if the current image had its colors counted before coming here.  If not, count it.
        Dim numColors As Long
        If g_LastImageScanned <> srcPDImage.imageID Then
            numColors = getQuickColorCount(tmpDIB, srcPDImage.imageID)
        Else
            numColors = g_LastColorCount
        End If
        
        'Pngnq can handle all types of transparency for us.  If pngnq cannot be found, we must rely on our own routines.
        If Not g_ImageFormats.pngnqEnabled Then
        
            'Does this DIB contain binary transparency?  If so, mark all transparent pixels with magic magenta.
            If tmpDIB.isAlphaBinary Then
                tmpDIB.applyAlphaCutoff
            Else
            
                'If we are in the midst of a batch conversion, we don't want to bother the user with alpha dialogs.
                ' Thus, use a default cut-off of 127 and continue on.
                If MacroStatus = MacroBATCH Then
                    tmpDIB.applyAlphaCutoff
                
                'We're not in a batch conversion, so ask the user which cut-off they would like to use.
                Else
            
                    Dim alphaCheck As VbMsgBoxResult
                    alphaCheck = promptAlphaCutoff(tmpDIB)
                    
                    'If the alpha dialog is canceled, abandon the entire save
                    If alphaCheck = vbCancel Then
                    
                        tmpDIB.eraseDIB
                        Set tmpDIB = Nothing
                        SavePNGImage = False
                        Exit Function
                    
                    'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                    Else
                        tmpDIB.applyAlphaCutoff g_AlphaCutoff, , g_AlphaCompositeColor
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
        If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth < 32) Then tmpDIB.compositeBackgroundColor 255, 255, 255
    
        'Also, if pngnq is enabled, we will use that for the transformation - so we need to reset the outgoing color depth to 24bpp
        If (tmpDIB.getDIBColorDepth = 24) And (outputColorDepth = 8) And g_ImageFormats.pngnqEnabled Then outputColorDepth = 24
    
    End If
    
    Message "Writing %1 file...", sFileType
    
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    'If the image is being reduced from some higher bit-depth to 1bpp, manually force a conversion with dithering
    If outputColorDepth = 1 Then fi_DIB = FreeImage_Dither(fi_DIB, FID_FS)
    
    'If the image contains alpha and pngnq is not available, we need to manually convert the FreeImage copy of the image to 8bpp.
    ' Then we need to apply alpha using the cut-off established earlier in this section.
    If handleAlpha And (Not g_ImageFormats.pngnqEnabled) Then
    
        fi_DIB = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_NNQUANT, True)
        
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
        
        fi_Palette(palIndex) = tmpDIB.getOriginalTransparentColor()
        
    End If
        
    'Use that handle to save the image to PNG format
    If fi_DIB <> 0 Then
    
        'Embed a background color if available, and the user has requested it.
        If pngPreserveBKGD And srcPDImage.imgStorage.Exists("pngBackgroundColor") Then
            
            Dim rQuad As RGBQUAD
            rQuad.Red = ExtractR(srcPDImage.imgStorage.Item("pngBackgroundColor"))
            rQuad.Green = ExtractG(srcPDImage.imgStorage.Item("pngBackgroundColor"))
            rQuad.Blue = ExtractB(srcPDImage.imgStorage.Item("pngBackgroundColor"))
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
        
        If Not fi_Check Then
            
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            SavePNGImage = False
            Exit Function
        
        'Save was successful.  'If pngnq is being used to help with the 8bpp reduction, now is when we use it.
        Else
            
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
                If tmpDIB.getDIBColorDepth = 32 Then
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
    If Len(ppmParams) > 0 Then cParams.setParamString ppmParams
    Dim ppmFormat As Long
    ppmFormat = cParams.GetLong(1, 0)

    Dim sFileType As String
    sFileType = "PPM"

    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
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
    srcPDImage.getCompositedImage tmpDIB, False
    
    'PPM only supports 24bpp
    If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.convertTo24bpp
        
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
        
    'Use that handle to save the image to PPM format (ASCII)
    If fi_DIB <> 0 Then
    
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, PPMPath, FIF_PPM, ppm_Encoding, FICD_24BPP, , , , , True)
        
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
    If Not g_ImageFormats.FreeImageEnabled Then
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveTGAImage = False
        Exit Function
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.getCompositedImage tmpDIB, False
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'Does this DIB contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If tmpDIB.isAlphaBinary Then
            tmpDIB.applyAlphaCutoff
        Else
        
            'If we are in the midst of a batch conversion, we don't want to bother the user with alpha dialogs.
            ' Thus, use a default cut-off of 127 and continue on.
            If MacroStatus = MacroBATCH Then
                tmpDIB.applyAlphaCutoff
            
            'We're not in a batch conversion, so ask the user which cut-off they would like to use.
            Else
            
                Dim alphaCheck As VbMsgBoxResult
                alphaCheck = promptAlphaCutoff(tmpDIB)
                
                'If the alpha dialog is canceled, abandon the entire save
                If alphaCheck = vbCancel Then
                
                    tmpDIB.eraseDIB
                    Set tmpDIB = Nothing
                    SaveTGAImage = False
                    Exit Function
                
                'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                Else
                    tmpDIB.applyAlphaCutoff g_AlphaCutoff, , g_AlphaCompositeColor
                End If
            End If
            
        End If
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth < 32) Then tmpDIB.compositeBackgroundColor 255, 255, 255
    
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
        
        fi_Palette(palIndex) = tmpDIB.getOriginalTransparentColor()
        
    End If
    
    'Finally, prepare a TGA save flag.  If the user has requested RLE encoding, pass that along to FreeImage.
    Dim TGAflags As Long
    TGAflags = TARGA_DEFAULT
            
    If TGACompression Then TGAflags = TARGA_SAVE_RLE
    
    'Use that handle to save the image to TGA format
    If fi_DIB <> 0 Then
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TGAPath, FIF_TARGA, TGAflags, outputColorDepth, , , , , True)
        
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

'Save to JPEG using the FreeImage library.  FreeImage offers significantly more JPEG features than GDI+.
Public Function SaveJPEGImage(ByRef srcPDImage As pdImage, ByVal JPEGPath As String, ByVal jpegParams As String) As Boolean
    
    On Error GoTo SaveJPEGError
    
    Dim sFileType As String
    sFileType = "JPEG"
    
    'Parse all possible JPEG parameters
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(jpegParams) > 0 Then cParams.setParamString jpegParams
    Dim jpegFlags As Long
    
    'Start by retrieving quality
    jpegFlags = cParams.GetLong(1, 92)
    
    'If FreeImage is enabled, check for an "automatically determine quality" request.
    If g_ImageFormats.FreeImageEnabled Then
    
        Dim autoQualityCheck As jpegAutoQualityMode
        autoQualityCheck = cParams.GetLong(4, 0)
        
        If autoQualityCheck <> doNotUseAutoQuality Then
        
            'The user has requested that we determine a proper quality value for them.  Do so now!
            Message "Testing JPEG quality values to determine best setting..."
            
            'Large images take a veeeery long time to search.  Force a max value of 1024x1024, and search the smaller image.
            ' (This should still result in a good value, but at a much smaller time investment.)
            Dim testDIB As pdDIB
            Set testDIB = New pdDIB
            srcPDImage.getCompositedImage testDIB, False
            If testDIB.getDIBColorDepth = 32 Then testDIB.convertTo24bpp
            
            If (testDIB.getDIBWidth > 1024) Or (testDIB.getDIBHeight > 1024) Then
            
                'Find new dimensions
                Dim newWidth As Long, newHeight As Long
                convertAspectRatio testDIB.getDIBWidth, testDIB.getDIBHeight, 1024, 1024, newWidth, newHeight
                
                'Create a temporary source image (resizing requires separate source and destination images)
                Dim tmpSourceDIB As pdDIB
                Set tmpSourceDIB = New pdDIB
                tmpSourceDIB.createFromExistingDIB testDIB
                
                'Resize the temp image and continue
                testDIB.createFromExistingDIB tmpSourceDIB, newWidth, newHeight
                Set tmpSourceDIB = Nothing
            
            End If
            
            jpegFlags = findQualityForDesiredJPEGPerception(testDIB, autoQualityCheck, cParams.GetBool(5, False))
            Message "Ideal quality of %1 found.  Continuing with save...", jpegFlags
            
            Set testDIB = Nothing
        
        Else
            Message "Preparing %1 image...", sFileType
        End If
    
    Else
        Message "Preparing %1 image...", sFileType
    End If
    
    'If FreeImage is not available, fall back to GDI+.  If that is not available, fail the function.
    If Not g_ImageFormats.FreeImageEnabled Then
    
        If g_ImageFormats.GDIPlusEnabled Then
            SaveJPEGImage = GDIPlusSavePicture(srcPDImage, JPEGPath, ImageJPEG, 24, jpegFlags)
        Else
            SaveJPEGImage = False
            Message "No %1 encoder found. Save aborted.", "JPEG"
        End If
        
        Exit Function
        
    End If
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.getCompositedImage tmpDIB, False
    
    'JPEGs can only save 24bpp images, so flatten the alpha as necessary
    If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.convertTo24bpp
        
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
        
    'If the image is grayscale, instruct FreeImage to internally mark the image as such
    Dim outputColorDepth As Long
    Message "Analyzing image color content..."
    
    If tmpDIB.isDIBGrayscale Then
    
        Message "No color found.  Saving 8bpp grayscale JPEG."
        outputColorDepth = 8
        fi_DIB = FreeImage_ConvertToGreyscale(fi_DIB)
        
    Else
    
        Message "Color found.  Saving 24bpp full-color JPEG."
        outputColorDepth = 24
        
    End If
        
    'Combine all received flags into one
    jpegFlags = jpegFlags Or cParams.GetLong(2, 0)
    
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
        fi_Check = FreeImage_SaveEx(fi_DIB, JPEGPath, FIF_JPEG, jpegFlags, outputColorDepth, , , , , True)
        
        If fi_Check Then
            Message "%1 save complete.", sFileType
        Else
            
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            SaveJPEGImage = False
            Exit Function
            
        End If
        
    Else
    
        Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
        SaveJPEGImage = False
        Exit Function
        
    End If
    
    SaveJPEGImage = True
    Exit Function
    
SaveJPEGError:

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
    If Not g_ImageFormats.FreeImageEnabled Then
        
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
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
    srcPDImage.getCompositedImage tmpDIB, False
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'Does this DIB contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If tmpDIB.isAlphaBinary Then
            tmpDIB.applyAlphaCutoff
        Else
            
            'If we are in the midst of a batch conversion, we don't want to bother the user with alpha dialogs.
            ' Thus, use a default cut-off of 127 and continue on.
            If MacroStatus = MacroBATCH Then
                tmpDIB.applyAlphaCutoff
            
            'We're not in a batch conversion, so ask the user which cut-off they would like to use.
            Else
                Dim alphaCheck As VbMsgBoxResult
                alphaCheck = promptAlphaCutoff(tmpDIB)
                
                'If the alpha dialog is canceled, abandon the entire save
                If alphaCheck = vbCancel Then
                
                    tmpDIB.eraseDIB
                    Set tmpDIB = Nothing
                    SaveTIFImage = False
                    Exit Function
                
                'If it wasn't canceled, use the value it provided to apply our alpha cut-off
                Else
                    tmpDIB.applyAlphaCutoff g_AlphaCutoff, , g_AlphaCompositeColor
                End If
            End If
            
        End If
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (tmpDIB.getDIBColorDepth = 32) And (outputColorDepth < 32) Then tmpDIB.compositeBackgroundColor 255, 255, 255
    
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
        
        fi_Palette(palIndex) = tmpDIB.getOriginalTransparentColor()
        
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
            tmpDIB.convertToCMYK32
            fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
            
        End If
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TIFPath, FIF_TIFF, TIFFFlags, outputColorDepth, , , , , True)
        
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
Public Function SaveJP2Image(ByRef srcPDImage As pdImage, ByVal jp2Path As String, ByVal outputColorDepth As Long, Optional ByVal jp2Params As String = "") As Boolean
    
    On Error GoTo SaveJP2Error
    
    'Parse all possible JPEG-2000 params
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(jp2Params) > 0 Then cParams.setParamString jp2Params
    Dim JP2Quality As Long
    If cParams.doesParamExist(1) Then JP2Quality = cParams.GetLong(1) Else JP2Quality = 16
    
    Dim sFileType As String
    sFileType = "JPEG-2000"
    
    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveJP2Image = False
        Exit Function
        
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.getCompositedImage tmpDIB, False
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (tmpDIB.getDIBColorDepth = 32) Then tmpDIB.convertTo24bpp
    
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    'Use that handle to save the image to JPEG format
    If fi_DIB <> 0 Then
                
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, jp2Path, FIF_JP2, JP2Quality, outputColorDepth, , , , , True)
        
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
Public Function SaveJXRImage(ByRef srcPDImage As pdImage, ByVal jxrPath As String, ByVal outputColorDepth As Long, Optional ByVal jxrParams As String = "") As Boolean
    
    On Error GoTo SaveJXRError
    
    'Parse all possible JXR params
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(jxrParams) > 0 Then cParams.setParamString jxrParams
    Dim jxrQuality As Long, jxrProgressive As Boolean
    
    'Quality is the first parameter
    If cParams.doesParamExist(1) Then jxrQuality = cParams.GetLong(1) Else jxrQuality = 0
    
    'Progressive encoding is the second parameter
    If cParams.doesParamExist(2) Then jxrProgressive = cParams.GetBool(2) Else jxrProgressive = False
    
    'FreeImage just accepts a single Long-type set of flags, so merge the progressive setting with the quality one
    If jxrProgressive Then jxrQuality = jxrQuality Or JXR_PROGRESSIVE
    
    Dim sFileType As String
    sFileType = "JPEG XR"
    
    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveJXRImage = False
        Exit Function
        
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.getCompositedImage tmpDIB, False
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (tmpDIB.getDIBColorDepth = 32) Then tmpDIB.convertTo24bpp
    
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    'Use that handle to save the image to JPEG XR format
    If fi_DIB <> 0 Then
                
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, jxrPath, FIF_JXR, jxrQuality, outputColorDepth, , , , , True)
        
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
Public Function SaveWebPImage(ByRef srcPDImage As pdImage, ByVal WebPPath As String, ByVal outputColorDepth As Long, Optional ByVal WebPParams As String = "") As Boolean
    
    On Error GoTo SaveWebPError
    
    'Parse all possible WebP params
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(WebPParams) > 0 Then cParams.setParamString WebPParams
    Dim WebPQuality As Long
    If cParams.doesParamExist(1) Then WebPQuality = cParams.GetLong(1) Else WebPQuality = 0
    
    Dim sFileType As String
    sFileType = "WebP"
    
    'Make sure we found the plug-in when we loaded the program
    If Not g_ImageFormats.FreeImageEnabled Then
        
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SaveWebPImage = False
        Exit Function
        
    End If
    
    Message "Preparing %1 image...", sFileType
    
    'Retrieve a composited copy of the image, at full size
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.getCompositedImage tmpDIB, False
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (tmpDIB.getDIBColorDepth = 32) Then tmpDIB.convertTo24bpp
    
    'Convert our current DIB to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    'Use that handle to save the image to WebP format
    If fi_DIB <> 0 Then
                
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, WebPPath, FIF_WEBP, WebPQuality, outputColorDepth, , , , , True)
        
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
Public Sub fillDIBWithJPEGVersion(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal JPEGQuality As Long, Optional ByVal jpegSubsample As Long = JPEG_SUBSAMPLING_422)

    'srcDIB may be 32bpp.  Convert it to 24bpp if necessary.
    If srcDIB.getDIBColorDepth = 32 Then srcDIB.convertTo24bpp

    'Pass the DIB to FreeImage, which will make a copy for itself.
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
    
    'Prepare matching flags for FreeImage's JPEG encoder
    Dim jpegFlags As Long
    jpegFlags = JPEGQuality Or jpegSubsample
        
    'Now comes the actual JPEG conversion, which is handled exclusively by FreeImage.  Basically, we ask it to save
    ' the image in JPEG format to a byte array; we then hand that byte array back to it and request a decompression.
    Dim jpegArray() As Byte
    Dim fi_Check As Long
    fi_Check = FreeImage_SaveToMemoryEx(FIF_JPEG, fi_DIB, jpegArray, jpegFlags, True)
    
    fi_DIB = FreeImage_LoadFromMemoryEx(jpegArray, FILO_JPEG_FAST)
    
    'Copy the newly decompressed JPEG into the destination pdDIB object.
    SetDIBitsToDevice dstDIB.getDIBDC, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight, 0, 0, 0, dstDIB.getDIBHeight, ByVal FreeImage_GetBits(fi_DIB), ByVal FreeImage_GetInfo(fi_DIB), 0&
    
    'Release the FreeImage copy of the DIB
    FreeImage_Unload fi_DIB
    Erase jpegArray

End Sub

'Given a source image and a desired JPEG perception quality, test various JPEG quality values until an ideal one is found
Public Function findQualityForDesiredJPEGPerception(ByRef srcImage As pdDIB, ByVal desiredPerception As jpegAutoQualityMode, Optional ByVal useHighQualityColorMatching As Boolean = False) As Long

    'If desiredPerception is 0 ("do not use auto check"), exit
    If desiredPerception = doNotUseAutoQuality Then
        findQualityForDesiredJPEGPerception = 0
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
        fillDIBWithJPEGVersion tmpJPEGImage, tmpJPEGImage, curJPEGQuality
        
        'Here is where high-quality and low-quality color-matching diverge.
        If useHighQualityColorMatching Then
        
            'Convert the JPEG-ified DIB to the L*a*b* color space
            convertEntireDIBToLabColor tmpJPEGImage, dstImageData
            
            'Retrieve a mean RMSD for the two images
            rmsdCheck = findMeanRMSDForTwoArrays(srcImageData, dstImageData, srcImage.getDIBWidth - 1, srcImage.getDIBHeight - 1)
            
        Else
        
            rmsdCheck = findMeanRMSDForTwoDIBs(srcImage, tmpJPEGImage)
        
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
        fillDIBWithJPEGVersion tmpJPEGImage, tmpJPEGImage, curJPEGQuality
        
        'Here is where high-quality and low-quality color-matching diverge.
        If useHighQualityColorMatching Then
        
            'Convert the JPEG-ified DIB to the L*a*b* color space
            convertEntireDIBToLabColor tmpJPEGImage, dstImageData
            
            'Retrieve a mean RMSD for the two images
            rmsdCheck = findMeanRMSDForTwoArrays(srcImageData, dstImageData, srcImage.getDIBWidth - 1, srcImage.getDIBHeight - 1)
            
        Else
        
            rmsdCheck = findMeanRMSDForTwoDIBs(srcImage, tmpJPEGImage)
        
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
    findQualityForDesiredJPEGPerception = curJPEGQuality

End Function

'This function takes two 24bpp DIBs and compares them, returning a single mean RMSD.
Public Function findMeanRMSDForTwoDIBs(ByRef srcDib1 As pdDIB, ByRef srcDib2 As pdDIB) As Double

    Dim totalRMSD As Double
    totalRMSD = 0

    Dim x As Long, y As Long, QuickX As Long
    
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    
    'Acquire pointers to both DIB arrays
    Dim tmpSA1 As SAFEARRAY2D, tmpSA2 As SAFEARRAY2D
    
    Dim srcArray1() As Byte, srcArray2() As Byte
    
    prepSafeArray tmpSA1, srcDib1
    prepSafeArray tmpSA2, srcDib2
    
    CopyMemory ByVal VarPtrArray(srcArray1()), VarPtr(tmpSA1), 4
    CopyMemory ByVal VarPtrArray(srcArray2()), VarPtr(tmpSA2), 4
    
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = srcDib1.getDIBWidth
    imgHeight = srcDib2.getDIBHeight
    
    For x = 0 To imgWidth - 1
        QuickX = x * 3
    For y = 0 To imgHeight - 1
    
        'Retrieve both sets of L*a*b* coordinates
        r1 = srcArray1(QuickX, y)
        g1 = srcArray1(QuickX + 1, y)
        b1 = srcArray1(QuickX + 2, y)
        
        r2 = srcArray2(QuickX, y)
        g2 = srcArray2(QuickX + 1, y)
        b2 = srcArray2(QuickX + 2, y)
        
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
    findMeanRMSDForTwoDIBs = totalRMSD / (imgWidth * imgHeight)

End Function


'This function assumes two 24bpp DIBs have been pre-converted to Single-type L*a*b* arrays.  Use the L*a*b* data to return
' a mean RMSD for the two images.
Public Function findMeanRMSDForTwoArrays(ByRef srcArray1() As Single, ByRef srcArray2() As Single, ByVal imgWidth As Long, ByVal imgHeight As Long) As Double

    Dim totalRMSD As Double
    totalRMSD = 0

    Dim x As Long, y As Long, QuickX As Long
    
    Dim LabL1 As Double, LabA1 As Double, LabB1 As Double
    Dim labL2 As Double, labA2 As Double, labB2 As Double
    
    For x = 0 To imgWidth - 1
        QuickX = x * 3
    For y = 0 To imgHeight - 1
    
        'Retrieve both sets of L*a*b* coordinates
        LabL1 = srcArray1(QuickX, y)
        LabA1 = srcArray1(QuickX + 1, y)
        LabB1 = srcArray1(QuickX + 2, y)
        
        labL2 = srcArray2(QuickX, y)
        labA2 = srcArray2(QuickX + 1, y)
        labB2 = srcArray2(QuickX + 2, y)
        
        'Calculate an RMSD
        totalRMSD = totalRMSD + distanceThreeDimensions(LabL1, LabA1, LabB1, labL2, labA2, labB2)
    
    Next y
    Next x
    
    'Divide the total RMSD by the number of pixels in the image, then exit
    findMeanRMSDForTwoArrays = totalRMSD / (imgWidth * imgHeight)

End Function

'Given a source and destination DIB reference, fill the destination with a post-JPEG-2000-compression of the original.  This
' is used to generate the live preview used in PhotoDemon's "export JPEG-2000" dialog.
Public Sub fillDIBWithJP2Version(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal JP2Quality As Long)
    
    'Pass the DIB to FreeImage, which will make a copy for itself.
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
        
    'Now comes the actual JPEG-2000 conversion, which is handled exclusively by FreeImage.  Basically, we ask it to save
    ' the image in JPEG-2000 format to a byte array; we then hand that byte array back to it and request a decompression.
    Dim jp2Array() As Byte
    Dim fi_Check As Long
    fi_Check = FreeImage_SaveToMemoryEx(FIF_JP2, fi_DIB, jp2Array, JP2Quality, True)
    
    fi_DIB = FreeImage_LoadFromMemoryEx(jp2Array, 0)
    
    'Copy the newly decompressed JPEG-2000 into the destination pdDIB object.
    SetDIBitsToDevice dstDIB.getDIBDC, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight, 0, 0, 0, dstDIB.getDIBHeight, ByVal FreeImage_GetBits(fi_DIB), ByVal FreeImage_GetInfo(fi_DIB), 0&
    
    'Release the FreeImage copy of the DIB.
    FreeImage_Unload fi_DIB
    Erase jp2Array

End Sub

'Given a source and destination DIB reference, fill the destination with a post-WebP-compression of the original.  This
' is used to generate the live preview used in PhotoDemon's "export WebP" dialog.
Public Sub fillDIBWithWebPVersion(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal WebPQuality As Long)
    
    'Pass the DIB to FreeImage, which will make a copy for itself.
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
        
    'Now comes the actual WebP conversion, which is handled exclusively by FreeImage.  Basically, we ask it to save
    ' the image in WebP format to a byte array; we then hand that byte array back to it and request a decompression.
    Dim webPArray() As Byte
    Dim fi_Check As Long
    fi_Check = FreeImage_SaveToMemoryEx(FIF_WEBP, fi_DIB, webPArray, WebPQuality, True)
    
    fi_DIB = FreeImage_LoadFromMemoryEx(webPArray, 0)
    
    'Copy the newly decompressed image into the destination pdDIB object.
    SetDIBitsToDevice dstDIB.getDIBDC, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight, 0, 0, 0, dstDIB.getDIBHeight, ByVal FreeImage_GetBits(fi_DIB), ByVal FreeImage_GetInfo(fi_DIB), 0&
    
    'Release the FreeImage copy of the DIB.
    FreeImage_Unload fi_DIB
    Erase webPArray

End Sub

'Given a source and destination DIB reference, fill the destination with a post-JXR-compression of the original.  This
' is used to generate the live preview used in PhotoDemon's "export JXR" dialog.
Public Sub fillDIBWithJXRVersion(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal jxrQuality As Long)
    
    'Pass the DIB to FreeImage, which will make a copy for itself.
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
        
    'Now comes the actual JPEG XR conversion, which is handled exclusively by FreeImage.  Basically, we ask it to save
    ' the image in JPEG XR format to a byte array; we then hand that byte array back to it and request a decompression.
    Dim jxrArray() As Byte
    Dim fi_Check As Long
    fi_Check = FreeImage_SaveToMemoryEx(FIF_JXR, fi_DIB, jxrArray, jxrQuality, True)
    
    fi_DIB = FreeImage_LoadFromMemoryEx(jxrArray, 0)
    
    'Copy the newly decompressed image into the destination pdDIB object.
    SetDIBitsToDevice dstDIB.getDIBDC, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight, 0, 0, 0, dstDIB.getDIBHeight, ByVal FreeImage_GetBits(fi_DIB), ByVal FreeImage_GetInfo(fi_DIB), 0&
    
    'Release the FreeImage copy of the DIB.
    FreeImage_Unload fi_DIB
    Erase jxrArray

End Sub

'Save a new Undo/Redo entry to file.  This function is only called by the createUndoData function in the pdUndo class.
' For the most part, this function simply wraps other save functions; however, certain odd types of Undo diff files (e.g. layer headers)
' may be directly processed and saved by this function.
'
'Note that this function interacts closely with the matching LoadUndo function in the Loading module.  Any novel Undo diff types added
' here must also be mirrored there.
Public Function saveUndoData(ByRef srcPDImage As pdImage, ByRef dstUndoFilename As String, ByVal processType As PD_UNDO_TYPE, Optional ByVal targetLayerID As Long = -1) As Boolean

    'What kind of Undo data we save is determined by the current processType.
    Select Case processType
    
        'EVERYTHING, meaning a full copy of the pdImage stack and any selection data
        Case UNDO_EVERYTHING
            Saving.SavePhotoDemonImage srcPDImage, dstUndoFilename, True, False, False, False
            srcPDImage.mainSelection.writeSelectionToFile dstUndoFilename & ".selection"
            
        'A full copy of the pdImage stack
        Case UNDO_IMAGE
            Saving.SavePhotoDemonImage srcPDImage, dstUndoFilename, True, False, False, False
        
        'A full copy of the pdImage stack, *without any layer DIB data*
        Case UNDO_IMAGEHEADER
            Saving.SavePhotoDemonImage srcPDImage, dstUndoFilename, True, False, False, False, True
        
        'Layer data only (full layer header + full layer DIB).
        Case UNDO_LAYER
            Saving.SavePhotoDemonLayer srcPDImage.getLayerByID(targetLayerID), dstUndoFilename & ".layer", True, True, False, False, False
        
        'Layer header data only (e.g. DO NOT WRITE OUT THE LAYER DIB)
        Case UNDO_LAYERHEADER
            Saving.SavePhotoDemonLayer srcPDImage.getLayerByID(targetLayerID), dstUndoFilename & ".layer", True, True, False, False, True
            
        'Selection data only
        Case UNDO_SELECTION
            srcPDImage.mainSelection.writeSelectionToFile dstUndoFilename & ".selection"
            
        'Anything else (for now, default to the full pdImage stack until all other undo types are covered!)
        Case Else
            Saving.SavePhotoDemonImage srcPDImage, dstUndoFilename, True, False, False, False
        
    End Select
    
End Function

'Quickly save a DIB to file in PNG format.  Things like PD's Recent File manager use this function to quickly write DIBs out to file.
Public Function QuickSaveDIBAsPNG(ByVal dstFilename As String, ByRef srcDIB As pdDIB) As Boolean

    If (srcDIB Is Nothing) Or (srcDIB.getDIBWidth = 0) Or (srcDIB.getDIBHeight = 0) Then
        QuickSaveDIBAsPNG = False
        Exit Function
    End If

    'If FreeImage is available, use it to save the PNG; otherwise, fall back to GDI+
    If g_ImageFormats.FreeImageEnabled Then
        
        'Convert the temporary DIB to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
    
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
            fi_Check = FreeImage_SaveEx(fi_DIB, dstFilename, FIF_PNG, FISO_PNG_Z_BEST_SPEED, fi_OutputColorDepth, , , , , True)
            If Not fi_Check Then Message "Thumbnail save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
            
        Else
            Message "Thumbnail save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
        End If
        
    'FreeImage is not available; try to use GDI+ to save a PNG thumbnail
    Else
        
        If Not GDIPlusQuickSavePNG(dstFilename, srcDIB) Then Message "Thumbnail save failed (unspecified GDI+ error)."
        
    End If

End Function
