Attribute VB_Name = "Plugin_FreeImage_Expanded_Interface"
'***************************************************************************
'FreeImage Interface (Advanced)
'Copyright ©2012-2014 by Tanner Helland
'Created: 3/September/12
'Last updated: 07/December/13
'Last update: if CMYK files have an embedded ICC profile, retrieve the raw CMYK bits and apply the sRGB conversion ourselves.
'
'This module represents a new - and significantly more comprehensive - approach to loading images via the
' FreeImage libary. It handles a variety of decisions on a per-format basis to ensure optimal load speed
' and quality.
'
'Please note that this module relies heavily on Carsten Klein's FreeImage wrapper for VB (included in this project
' as Outside_FreeImageV3; see that file for license details).  Thanks to Carsten for his work on integrating FreeImage
' into classic VB.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'When loading a multipage image, the user will be prompted to load each page as an individual image.  If the user agrees,
' this variable will be set to TRUE.  PreLoadImage will then use this variable to launch the import of the subsequent pages.
Public g_imageHasMultiplePages As Boolean
Public g_imagePageCount As Long
    
'Is FreeImage available as a plugin?  (NOTE: this is now determined separately from FreeImageEnabled.)
Public Function isFreeImageAvailable() As Boolean
    If FileExist(g_PluginPath & "freeimage.dll") Then isFreeImageAvailable = True Else isFreeImageAvailable = False
End Function
    
'Load an image via FreeImage.  It is assumed that the source file has already been vetted for things like "does it exist?"
Public Function LoadFreeImageV4(ByVal srcFilename As String, ByRef dstDIB As pdDIB, Optional ByVal pageToLoad As Long = 0, Optional ByVal showMessages As Boolean = True) As Boolean

    On Error GoTo FreeImageV4_AdvancedError
    
    '****************************************************************************
    ' Make sure FreeImage exists and is usable
    '****************************************************************************
    
    'Double-check that FreeImage.dll was located at start-up
    If Not g_ImageFormats.FreeImageEnabled Then
        LoadFreeImageV4 = False
        Exit Function
    End If
    
    '****************************************************************************
    ' Determine image format
    '****************************************************************************
    
    If showMessages Then Message "Analyzing filetype..."
    
    'While we could manually test our extension against the FreeImage database, it is capable of doing so itself.
    'First, check the file header to see if it matches a known head type
    Dim fileFIF As FREE_IMAGE_FORMAT
    fileFIF = FreeImage_GetFileType(srcFilename)
    
    'For certain filetypes (CUT, MNG, PCD, TARGA and WBMP, according to the FreeImage documentation), the lack of a reliable
    ' header may prevent GetFileType from working.  As a result, double-check the file using its file extension.
    If fileFIF = FIF_UNKNOWN Then fileFIF = FreeImage_GetFIFFromFilename(srcFilename)
    
    'By this point, if the file still doesn't show up in FreeImage's database, abandon the import attempt.
    If Not FreeImage_FIFSupportsReading(fileFIF) Then
    
        If showMessages Then Message "Filetype not supported by FreeImage.  Import abandoned."
        LoadFreeImageV4 = False
        Exit Function
        
    End If
    
    'Store this file format inside the DIB
    dstDIB.setOriginalFormat fileFIF
    
    
    '****************************************************************************
    ' Based on the detected format, prepare any necessary load flags
    '****************************************************************************
    
    If showMessages Then Message "Preparing import flags..."
    
    'Certain filetypes offer import options.  Check the FreeImage type to see if we want to enable any optional flags.
    Dim fi_ImportFlags As FREE_IMAGE_LOAD_OPTIONS
    fi_ImportFlags = 0
    
    'For JPEGs, specify a preference for accuracy and quality over load speed under normal circumstances,
    ' but when performing a batch conversion choose the reverse (speed over accuracy).  Also, if the showMessages parameter
    ' is false, we know that preview-quality is acceptable - so load the image as quickly as possible.
    If fileFIF = FIF_JPEG Then
        
        If (MacroStatus = MacroBATCH) Or (Not showMessages) Then
            fi_ImportFlags = FILO_JPEG_FAST Or FILO_JPEG_EXIFROTATE
        Else
            fi_ImportFlags = FILO_JPEG_ACCURATE Or FILO_JPEG_EXIFROTATE
        End If
        
    End If
    
    'Check for CMYK JPEGs, TIFFs, and PSD files.  If an image is CMYK and an ICC profile is present, ask FreeImage to load the
    ' raw CMYK data. If no ICC profile is present, FreeImage is free to perform the CMYK -> RGB translation for us.
    Dim isCMYK As Boolean
    isCMYK = False
    
    If (fileFIF = FIF_JPEG) Or (fileFIF = FIF_PSD) Or (fileFIF = FIF_TIFF) Then
    
        'To speed up the load process, only load the file header, and explicitly instruct FreeImage to leave CMYK images
        ' in CMYK format (otherwise we can't detect CMYK, as it will be auto-converted to RGB!).
        Dim additionalFlags As Long
        additionalFlags = FILO_LOAD_NOPIXELS
        
        Select Case fileFIF
        
            Case FIF_JPEG
                additionalFlags = additionalFlags Or FILO_JPEG_CMYK
            
            Case FIF_PSD
                additionalFlags = additionalFlags Or FILO_PSD_CMYK
            
            Case FIF_TIFF
                additionalFlags = additionalFlags Or TIFF_CMYK
        
        End Select
    
        Dim tmpFIHandle As Long
        tmpFIHandle = FreeImage_Load(fileFIF, srcFilename, fi_ImportFlags Or additionalFlags)
        
        'Check the file's color type
        If FreeImage_GetColorType(tmpFIHandle) = FIC_CMYK Then
        
            'File is CMYK.  Check for an ICC profile.
            If FreeImage_HasICCProfile(tmpFIHandle) Then
            
                'CMYK + ICC profile means we want FreeImage to load the data in CMYK format, and we'll perform the conversion to sRGB ourselves.
                isCMYK = True
                
                Select Case fileFIF
        
                    Case FIF_JPEG
                        fi_ImportFlags = fi_ImportFlags Or FILO_JPEG_CMYK
                    
                    Case FIF_PSD
                        fi_ImportFlags = fi_ImportFlags Or FILO_PSD_CMYK
                    
                    Case FIF_TIFF
                        fi_ImportFlags = fi_ImportFlags Or TIFF_CMYK
                
                End Select
                
            End If
        
        End If
        
        'Release our header-only copy of the image
        FreeImage_Unload tmpFIHandle
        
    End If
    
    
    'FreeImage is crazy slow at loading RAW-format files, so custom flags are specified to speed up the process
    If fileFIF = FIF_RAW Then
        
        'If this is not a primary image, RAW format files can load just their thumbnail
        If (Not showMessages) Then fi_ImportFlags = FILO_RAW_PREVIEW
        
    End If
        
    'For icons, we prefer a white background (default is black).
    ' NOTE: this check is now disabled, because it uses the AND mask incorrectly for mixed-format icons.  A better fix is
    ' provided below - see the section starting with "If fileFIF = FIF_ICO Then..."
    'If fileFIF = FIF_ICO Then fi_ImportFlags = FILO_ICO_MAKEALPHA
    
    '****************************************************************************
    ' If the user has requested a specific page from a multipage image, prepare a few extra items
    '****************************************************************************
    
    Dim fi_multi_hDIB As Long
    Dim needToCloseMulti As Boolean
    
    If pageToLoad > 0 Then needToCloseMulti = True Else needToCloseMulti = False
    
    '****************************************************************************
    ' Load the image into a FreeImage container
    '****************************************************************************
        
    'With all flags set and filetype correctly determined, import the image
    Dim fi_hDIB As Long
    
    If (pageToLoad <= 0) Then
        If showMessages Then Message "Importing image from file..."
        fi_hDIB = FreeImage_Load(fileFIF, srcFilename, fi_ImportFlags)
    Else
        If fileFIF = FIF_GIF Then
            If showMessages Then Message "Importing frame # %1 from animated GIF file...", pageToLoad + 1
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_GIF, srcFilename, , , , FILO_GIF_PLAYBACK)
        ElseIf fileFIF = FIF_ICO Then
            If showMessages Then Message "Importing icon # %1 from ICO file...", pageToLoad + 1
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_ICO, srcFilename, , , , 0)
        Else
            If showMessages Then Message "Importing page # %1 from multipage TIFF file...", pageToLoad + 1
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_TIFF, srcFilename, , , , 0)
        End If
        fi_hDIB = FreeImage_LockPage(fi_multi_hDIB, pageToLoad)
    End If
    
    'Icon files may use a simple mask for their alpha channel; in this case, re-load the icon with the FILO_ICO_MAKEALPHA flag
    If fileFIF = FIF_ICO Then
        
        'Check the bit-depth
        If FreeImage_GetBPP(fi_hDIB) < 32 Then
        
            'If this is the first frame of the icon, unload it and try again
            If (pageToLoad <= 0) Then
                If fi_hDIB <> 0 Then FreeImage_UnloadEx fi_hDIB
                fi_hDIB = FreeImage_Load(fileFIF, srcFilename, FILO_ICO_MAKEALPHA)
            
            'If this is not the first frame, the required load code is a bit different.
            Else
                
                'Unlock this page and close the multi-page bitmap
                FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                FreeImage_CloseMultiBitmap fi_multi_hDIB
                
                'Now re-open it with the proper flags
                fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_ICO, srcFilename, , , , FILO_ICO_MAKEALPHA)
                fi_hDIB = FreeImage_LockPage(fi_multi_hDIB, pageToLoad)
                                
            End If
            
        End If
        
    End If
    
    'If an empty handle is returned, abandon the import attempt.
    If fi_hDIB = 0 Then
    
        If showMessages Then Message "Import via FreeImage failed (blank handle)."
        LoadFreeImageV4 = False
        Exit Function
        
    End If
        
        
        
    '****************************************************************************
    ' Retrieve generic metadata, like X and Y resolution (if available)
    '****************************************************************************
    
    dstDIB.setDPI FreeImage_GetResolutionX(fi_hDIB), FreeImage_GetResolutionY(fi_hDIB), True
    
    
    
    '****************************************************************************
    ' Retrieve any attached ICC profiles, and copy their contents into this DIB's ICC manager
    '****************************************************************************
    
    If FreeImage_HasICCProfile(fi_hDIB) Then
    
        'This image has an attached profile.  Retrieve it and stick it inside the image.
        dstDIB.ICCProfile.loadICCFromFreeImage fi_hDIB
        
    End If
    
        
    '****************************************************************************
    ' Retrieve format-specific information, like PNG background color
    '****************************************************************************
    
    'Check to see if the image has a background color embedded
    If FreeImage_HasBackgroundColor(fi_hDIB) Then
                
        'FreeImage will pass the background color to an RGBquad type-variable
        Dim rQuad As RGBQUAD
        If FreeImage_GetBackgroundColor(fi_hDIB, rQuad) Then
        
            'Normally, we can reassemble the .r/g/b values in the object, but paletted images work a bit differently - the
            ' palette index is stored in .rgbReserved.  Check for that, and if it's non-zero, retrieve the palette value instead.
            If rQuad.rgbReserved <> 0 Then
                Dim fi_Palette() As Long
                fi_Palette = FreeImage_GetPaletteExLong(fi_hDIB)
                dstDIB.setBackgroundColor fi_Palette(rQuad.rgbReserved)
                
            'Otherwise it's easy - just reassemble the RGB values from the quad
            Else
                dstDIB.setBackgroundColor RGB(rQuad.rgbRed, rQuad.rgbGreen, rQuad.rgbBlue)
            End If
        
        End If
        
    End If
    
    
    '****************************************************************************
    ' Determine native color depth
    '****************************************************************************
    
    'Before we continue the import, we need to make sure the pixel data is in a format appropriate for PhotoDemon.
    
    If showMessages Then Message "Analyzing color depth..."
    
    'First thing we want to check is the color depth.  PhotoDemon is designed around 16 million color images.  This could
    ' change in the future, but for now, force high-bit-depth images to a more appropriate 24 or 32bpp.
    Dim fi_BPP As Long
    fi_BPP = FreeImage_GetBPP(fi_hDIB)
    
    'Because it could be helpful later on, also retrieve the image datatype.  This is an internal FreeImage value
    ' corresponding to various data encodings (floating-point, complex, integer, etc).  If we ever want to handle
    ' high-bit-depth images, that value will be crucial!
    Dim fi_DataType As FREE_IMAGE_TYPE
    fi_DataType = FreeImage_GetImageType(fi_hDIB)
    
    Debug.Print "Image bit-depth of " & fi_BPP & " and data type " & fi_DataType & " detected."
    
    'If a high bit-depth image is incoming, we need to use a temporary DIB to hold the image's alpha data (which will
    ' be erased by the tone-mapping algorithm we'll use).  This is that object
    Dim tmpAlphaRequired As Boolean, tmpAlphaCopySuccess As Boolean
    tmpAlphaRequired = False
    tmpAlphaCopySuccess = False
    
    Dim tmpAlphaDIB As pdDIB
    
    'A number of other variables may be required as we adjust the bit-depth of the image to match PhotoDemon's internal requirements.
    Dim new_hDIB As Long
    
    Dim fi_hasTransparency As Boolean
    Dim fi_transparentEntries As Long
    
    '****************************************************************************
    ' If the image is > 32bpp, downsample it to 24 or 32bpp
    '****************************************************************************
    
    'First, check source images without an alpha channel.  Convert these using the superior tone mapping method.
    If (fi_BPP = 48) Or (fi_BPP = 96) Then
    
        If showMessages Then Message "High bit-depth RGB image identified.  Checking for non-standard alpha encoding..."
        
        'While images with these bit-depths may not have an alpha channel, they can have binary transparency - check for that now.
        ' (Note: as of FreeImage 3.15.3 binary bit-depths are not detected correctly.  That said, they may someday be supported -
        ' so I've implemented two checks to cover both contingencies.
        fi_hasTransparency = FreeImage_IsTransparent(fi_hDIB)
        fi_transparentEntries = FreeImage_GetTransparencyCount(fi_hDIB)
    
        'As of 25 Nov '12, the user can choose to disable tone-mapping (which makes HDR loading much faster, but reduces image quality).
        ' Check that preference before tone-mapping the image.
        If g_UserPreferences.GetPref_Boolean("Loading", "HDR Tone Mapping", True) Then
            
            If showMessages Then If showMessages Then Message "Tone mapping HDR image to preserve tonal range..."
            
            new_hDIB = FreeImage_ToneMapping(fi_hDIB, FITMO_REINHARD05)
            
            If pageToLoad <= 0 Then
                If (fi_hDIB <> new_hDIB) Then FreeImage_UnloadEx fi_hDIB
            Else
                If (fi_hDIB <> new_hDIB) Then
                    needToCloseMulti = False
                    FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                    FreeImage_CloseMultiBitmap fi_multi_hDIB
                End If
            End If
            
            fi_hDIB = new_hDIB
            
            If showMessages Then Message "Tone mapping complete."
        
        Else
                
            If fi_hasTransparency Or (fi_transparentEntries <> 0) Then
            
                If showMessages Then Message "Alpha found, but further tone-mapping ignored at user's request."
                new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, False)
                
                If pageToLoad <= 0 Then
                    If (fi_hDIB <> new_hDIB) Then FreeImage_UnloadEx fi_hDIB
                Else
                    If (fi_hDIB <> new_hDIB) Then
                        needToCloseMulti = False
                        FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                        FreeImage_CloseMultiBitmap fi_multi_hDIB
                    End If
                End If
                
                fi_hDIB = new_hDIB
            
            Else
            
                If showMessages Then Message "No alpha found.  Further tone-mapping ignored at user's request."
                new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_24BPP, False)
                
                If pageToLoad <= 0 Then
                    If (fi_hDIB <> new_hDIB) Then FreeImage_UnloadEx fi_hDIB
                Else
                    If (fi_hDIB <> new_hDIB) Then
                        needToCloseMulti = False
                        FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                        FreeImage_CloseMultiBitmap fi_multi_hDIB
                    End If
                End If
            
                fi_hDIB = new_hDIB
            
            End If
        
        End If
        
    End If
    
    'BecaHDR Tone Mapping may not preserve alpha channels (the FreeImage documentation is unclear on this),
    ' we must do the same as above - manually make a copy of the image's alpha data, then reduce the image using tone mapping.
    ' Later in the process we will restore the alpha data to the image.
    If (fi_BPP = 64) Or (fi_BPP = 128) Then
    
        'Again, check for the user's preference on tone-mapping
        If g_UserPreferences.GetPref_Boolean("Loading", "HDR Tone Mapping", True) Then
        
            If showMessages Then Message "High bit-depth RGBA image identified.  Tone mapping HDR image to preserve tonal range..."
        
            'Now, convert the RGB data using the superior tone-mapping method.
            new_hDIB = FreeImage_ToneMapping(fi_hDIB, FITMO_REINHARD05)
            
            If pageToLoad <= 0 Then
                If (new_hDIB <> fi_hDIB) Then FreeImage_UnloadEx fi_hDIB
            Else
                If (new_hDIB <> fi_hDIB) Then
                    needToCloseMulti = False
                    FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                    FreeImage_CloseMultiBitmap fi_multi_hDIB
                End If
            End If
            
            fi_hDIB = new_hDIB
            
            If showMessages Then Message "Tone mapping complete."
            
        Else
        
            If showMessages Then Message "High bit-depth RGBA image identified.  Tone-mapping ignored at user's request."
            new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, False)
            
            If pageToLoad <= 0 Then
                If (fi_hDIB <> new_hDIB) Then FreeImage_UnloadEx fi_hDIB
            Else
                If (fi_hDIB <> new_hDIB) Then
                    needToCloseMulti = False
                    FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                    FreeImage_CloseMultiBitmap fi_multi_hDIB
                End If
            End If
            
            fi_hDIB = new_hDIB
        
        End If
        
    End If
    
    
    '****************************************************************************
    ' Now that we have filtered out > 32bpp images, store the current color depth of the original image.
    '****************************************************************************
    
    dstDIB.setOriginalColorDepth FreeImage_GetBPP(fi_hDIB)
    
    
    '****************************************************************************
    ' If the image is < 24bpp, upsample it to 24bpp or 32bpp
    '****************************************************************************
    
    'Similarly, check for low-bit-depth images
    If fi_BPP < 24 Then
        
        'Next, check to see if this is actually a high-bit-depth grayscale image masquerading as a low-bit-depth RGB image
        Dim fi_imageType As FREE_IMAGE_TYPE
        fi_imageType = FreeImage_GetImageType(fi_hDIB)
        
        'If it is such a grayscale image, it requires a unique conversion operation
        If fi_imageType = FIT_UINT16 Then
            
            If showMessages Then Message "Tone-mapping high-bit-depth grayscale image to 24bpp..."
            
            'First, convert it to a high-bit depth RGB image
            fi_hDIB = FreeImage_ConvertToRGB16(fi_hDIB)
            
            'Now use tone-mapping to reduce it back to 24bpp or 32bpp (contingent on the presence of transparency)
            fi_hasTransparency = FreeImage_IsTransparent(fi_hDIB)
            fi_transparentEntries = FreeImage_GetTransparencyCount(fi_hDIB)
        
            If fi_hasTransparency Or (fi_transparentEntries <> 0) Then
                new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, False)
                
                If pageToLoad <= 0 Then
                    If (new_hDIB <> fi_hDIB) Then FreeImage_UnloadEx fi_hDIB
                Else
                    If (new_hDIB <> fi_hDIB) Then
                        needToCloseMulti = False
                        FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                        FreeImage_CloseMultiBitmap fi_multi_hDIB
                    End If
                End If
            
                fi_hDIB = new_hDIB
            Else
                new_hDIB = FreeImage_ToneMapping(fi_hDIB, FITMO_REINHARD05)
                
                If pageToLoad <= 0 Then
                    If (new_hDIB <> fi_hDIB) Then FreeImage_UnloadEx fi_hDIB
                Else
                    If (new_hDIB <> fi_hDIB) Then
                        needToCloseMulti = False
                        FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                        FreeImage_CloseMultiBitmap fi_multi_hDIB
                    End If
                End If
            
                fi_hDIB = new_hDIB
            End If
            
        Else
        
            'Conversion to higher bit depths is contingent on the presence of an alpha channel
            fi_hasTransparency = FreeImage_IsTransparent(fi_hDIB)
        
            'Images with an alpha channel are converted to 32 bit.  Without, 24.
            If fi_hasTransparency = True Then
                new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, False)
                
                If pageToLoad <= 0 Then
                    If (new_hDIB <> fi_hDIB) Then FreeImage_UnloadEx fi_hDIB
                Else
                    If (new_hDIB <> fi_hDIB) Then
                        needToCloseMulti = False
                        FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                        FreeImage_CloseMultiBitmap fi_multi_hDIB
                    End If
                End If
            
                fi_hDIB = new_hDIB
            Else
                new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_24BPP, False)
                
                If pageToLoad <= 0 Then
                    If (new_hDIB <> fi_hDIB) Then FreeImage_UnloadEx fi_hDIB
                Else
                    If (new_hDIB <> fi_hDIB) Then
                        needToCloseMulti = False
                        FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                        FreeImage_CloseMultiBitmap fi_multi_hDIB
                    End If
                End If
            
                fi_hDIB = new_hDIB
            End If
            
        End If
        
    End If
    
    'By this point, we have loaded the image, and it is guaranteed to be at 24 or 32 bit color depth.  Verify it one final time.
    fi_BPP = FreeImage_GetBPP(fi_hDIB)
    
    
    '****************************************************************************
    ' Perform a special check for CMYK images.  They require additional handling.
    '****************************************************************************
    
    If isCMYK Then
    
        If showMessages Then Message "CMYK image detected.  Preparing transform into RGB space..."
        
        'Copy the CMYK data into a 32bpp DIB.  (Note that we could pass the FreeImage DIB copy directly into the function, but the resulting
        ' image would be top-down instead of bottom-up.  It's easier to just use our own PD-specific DIB object.)
        Dim tmpCMYKDIB As pdDIB
        Set tmpCMYKDIB = New pdDIB
        tmpCMYKDIB.createBlank FreeImage_GetWidth(fi_hDIB), FreeImage_GetHeight(fi_hDIB), 32
        SetDIBitsToDevice tmpCMYKDIB.getDIBDC, 0, 0, tmpCMYKDIB.getDIBWidth, tmpCMYKDIB.getDIBHeight, 0, 0, 0, tmpCMYKDIB.getDIBHeight, ByVal FreeImage_GetBits(fi_hDIB), ByVal FreeImage_GetInfo(fi_hDIB), 0&
        
        'Prepare a blank 24bpp DIB to receive the transformed sRGB data
        Dim tmpRGBDIB As pdDIB
        Set tmpRGBDIB = New pdDIB
        tmpRGBDIB.createBlank tmpCMYKDIB.getDIBWidth, tmpCMYKDIB.getDIBHeight, 24
        
        'Apply the transformation using the dedicated CMYK transform handler
        If applyCMYKTransform(dstDIB.ICCProfile.getICCDataPointer, dstDIB.ICCProfile.getICCDataSize, tmpCMYKDIB, tmpRGBDIB, dstDIB.ICCProfile.getSourceRenderIntent) Then
        
            Message "Copying newly transformed sRGB data..."
        
            'The transform was successful.  Copy the new sRGB data back into the FreeImage object, so the load process can continue.
            FreeImage_Unload fi_hDIB
            fi_hDIB = FreeImage_CreateFromDC(tmpRGBDIB.getDIBDC)
            fi_BPP = FreeImage_GetBPP(fi_hDIB)
            dstDIB.ICCProfile.markSuccessfulProfileApplication
            
        'Something went horribly wrong.  Re-load the image and use FreeImage to apply the CMYK -> RGB transform.
        Else
        
            Message "ICC-based CMYK transformation failed.  Falling back to default CMYK conversion..."
        
            FreeImage_Unload fi_hDIB
            fi_hDIB = FreeImage_Load(fileFIF, srcFilename, FILO_JPEG_ACCURATE Or FILO_JPEG_EXIFROTATE)
            fi_BPP = FreeImage_GetBPP(fi_hDIB)
        
        End If
        
        Set tmpCMYKDIB = Nothing
        Set tmpRGBDIB = Nothing
    
    End If
    
    
    '****************************************************************************
    ' PD's new rendering engine requires pre-multiplied alpha values.  Apply premultiplication now.
    '****************************************************************************
    
    If fi_BPP = 32 Then FreeImage_PreMultiplyWithAlpha fi_hDIB
    
    
    '****************************************************************************
    ' Create a blank pdDIB, which will receive a copy of the image in DIB format
    '****************************************************************************
    
    'We are now finally ready to load the image.
    
    If showMessages Then Message "Requesting memory for image transfer..."
    
    'Get width and height from the file, and create a new DIB to match
    Dim fi_Width As Long, fi_Height As Long
    fi_Width = FreeImage_GetWidth(fi_hDIB)
    fi_Height = FreeImage_GetHeight(fi_hDIB)
    
    Dim creationSuccess As Boolean
    
    'Update Dec '12: certain faulty TIFF files can confuse FreeImage and cause it to report wildly bizarre height and width
    ' values; check for this, and if it happens, abandon the load immediately.  (This is not ideal, because it leaks memory
    ' - but it prevents a hard program crash, so I consider it the lesser of two evils.)
    If (fi_Width > 1000000) Or (fi_Height > 1000000) Then
        LoadFreeImageV4 = False
        Exit Function
    Else
        creationSuccess = dstDIB.createBlank(fi_Width, fi_Height, fi_BPP)
    End If
    
    'Make sure the blank DIB creation worked
    If Not creationSuccess Then
        If showMessages Then Message "Import via FreeImage failed (couldn't create DIB)."
        
        If (pageToLoad <= 0) Or (Not needToCloseMulti) Then
            If fi_hDIB <> 0 Then FreeImage_UnloadEx fi_hDIB
        Else
            If (fi_hDIB <> 0) Then FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
            If (fi_multi_hDIB <> 0) Then FreeImage_CloseMultiBitmap fi_multi_hDIB
        End If
        
        LoadFreeImageV4 = False
        Exit Function
    End If
    
    '****************************************************************************
    ' Copy the data from the FreeImage object to the target pdDIB object
    '****************************************************************************
    
    If showMessages Then Message "Memory secured.  Finalizing image load..."
        
    'Copy the bits from the FreeImage DIB to our DIB
    SetDIBitsToDevice dstDIB.getDIBDC, 0, 0, fi_Width, fi_Height, 0, 0, 0, fi_Height, ByVal FreeImage_GetBits(fi_hDIB), ByVal FreeImage_GetInfo(fi_hDIB), 0&
    
    
    
    '****************************************************************************
    ' Release all FreeImage-specific structures and links
    '****************************************************************************
    
    'With the image bits now safely in our care, release the FreeImage DIB
    If (pageToLoad <= 0) Or (Not needToCloseMulti) Then
        FreeImage_UnloadEx fi_hDIB
    Else
        FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
        FreeImage_CloseMultiBitmap fi_multi_hDIB
    End If
    
    If showMessages Then Message "Image load successful.  FreeImage released."
    
    
    '****************************************************************************
    ' If necessary, restore any lost alpha data
    '****************************************************************************
    
    'We are almost done.  The last thing we need to do is restore the alpha values if this was a high-bit-depth image
    ' whose alpha data was lost during the tone-mapping phase.
    If tmpAlphaRequired Then
    
        If showMessages Then Message "Restoring alpha data..."
        
        dstDIB.copyAlphaFromExistingDIB tmpAlphaDIB
        dstDIB.fixPremultipliedAlpha True
        tmpAlphaDIB.eraseDIB
        Set tmpAlphaDIB = Nothing
        
        If showMessages Then Message "Alpha data restored successfully."
        
    End If
    
    '****************************************************************************
    ' Load complete
    '****************************************************************************
    
    'Mark this load as successful
    LoadFreeImageV4 = True
    
    Exit Function
    
    '****************************************************************************
    ' Error handling
    '****************************************************************************
    
FreeImageV4_AdvancedError:

    'Release the FreeImage DIB if available
    If fi_hDIB <> 0 Then FreeImage_UnloadEx fi_hDIB
    
    'Display a relevant error message
    If showMessages Then Message "Import via FreeImage failed (Err # %1)", Err.Number
    
    'Mark this load as unsuccessful
    LoadFreeImageV4 = False
    
End Function

'See if an image file is actually comprised of multiple files (e.g. animated GIFs, multipage TIFs).
' Input: file name to be checked
' Returns: 0 if only one image is found.  Page (or frame) count if multiple images are found.
Public Function isMultiImage(ByVal srcFilename As String) As Long

    On Error GoTo isMultiImage_Error
    
    'Double-check that FreeImage.dll was located at start-up
    If Not g_ImageFormats.FreeImageEnabled Then
        isMultiImage = 0
        Exit Function
    End If
        
    'Determine the file type.  (Currently, this feature only works on animated GIFs, multipage TIFFs, and icons.)
    Dim fileFIF As FREE_IMAGE_FORMAT
    fileFIF = FreeImage_GetFileType(srcFilename)
    If fileFIF = FIF_UNKNOWN Then fileFIF = FreeImage_GetFIFFromFilename(srcFilename)
    
    'If FreeImage can't determine the file type, or if the filetype is not GIF or TIF, return False
    If (Not FreeImage_FIFSupportsReading(fileFIF)) Or ((fileFIF <> FIF_GIF) And (fileFIF <> FIF_TIFF) And (fileFIF <> FIF_ICO)) Then
        isMultiImage = 0
        Exit Function
    End If
    
    'At this point, we are guaranteed that the image is a GIF, TIFF, or icon file.
    ' Open the file using the multipage function
    Dim fi_multi_hDIB As Long
    If fileFIF = FIF_GIF Then
        fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_GIF, srcFilename)
    ElseIf fileFIF = FIF_ICO Then
        fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_ICO, srcFilename)
    Else
        fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_TIFF, srcFilename)
    End If
    
    'Get the page count, then close the file
    Dim pageCheck As Long
    pageCheck = FreeImage_GetPageCount(fi_multi_hDIB)
    FreeImage_CloseMultiBitmap fi_multi_hDIB
    
    'Return the page count (which will be zero if only a single page or frame is present)
    isMultiImage = pageCheck
    
    Exit Function
    
isMultiImage_Error:

    isMultiImage = 0

End Function
