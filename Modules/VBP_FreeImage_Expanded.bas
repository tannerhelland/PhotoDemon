Attribute VB_Name = "Plugin_FreeImage_Expanded_Interface"
'***************************************************************************
'FreeImage Interface (Advanced)
'Copyright ©2012-2014 by Tanner Helland
'Created: 3/September/12
'Last updated: 08/July/14
'Last update: remove global page count variables in favor of more OOP-appropriate functions
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

Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean
    
'Is FreeImage available as a plugin?  (NOTE: this is now determined separately from FreeImageEnabled.)
Public Function isFreeImageAvailable() As Boolean
    If FileExist(g_PluginPath & "freeimage.dll") Then isFreeImageAvailable = True Else isFreeImageAvailable = False
End Function
    
'Load an image via FreeImage.  It is assumed that the source file has already been vetted for things like "does it exist?"
Public Function LoadFreeImageV4(ByVal srcFilename As String, ByRef dstDIB As pdDIB, Optional ByVal pageToLoad As Long = 0, Optional ByVal showMessages As Boolean = True, Optional ByRef targetImage As pdImage = Nothing) As Boolean

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
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Analyzing filetype..."
    #End If
    
    'While we could manually test our extension against the FreeImage database, it is capable of doing so itself.
    'First, check the file header to see if it matches a known head type
    Dim fileFIF As FREE_IMAGE_FORMAT
    fileFIF = FreeImage_GetFileType(srcFilename)
    
    'For certain filetypes (CUT, MNG, PCD, TARGA and WBMP, according to the FreeImage documentation), the lack of a reliable
    ' header may prevent GetFileType from working.  As a result, double-check the file using its file extension.
    If fileFIF = FIF_UNKNOWN Then fileFIF = FreeImage_GetFIFFromFilename(srcFilename)
    
    'By this point, if the file still doesn't show up in FreeImage's database, abandon the import attempt.
    If Not FreeImage_FIFSupportsReading(fileFIF) Then
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Filetype not supported by FreeImage.  Import abandoned."
        #End If
        
        LoadFreeImageV4 = False
        Exit Function
        
    End If
    
    'Store this file format inside the DIB
    dstDIB.setOriginalFormat fileFIF
    
    
    '****************************************************************************
    ' Based on the detected format, prepare any necessary load flags
    '****************************************************************************
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Preparing import flags..."
    #End If
    
    'Certain filetypes offer import options.  Check the FreeImage type to see if we want to enable any optional flags.
    Dim fi_ImportFlags As FREE_IMAGE_LOAD_OPTIONS
    fi_ImportFlags = 0
    
    'For JPEGs, specify a preference for accuracy and quality over load speed.
    If fileFIF = FIF_JPEG Then
        fi_ImportFlags = FILO_JPEG_ACCURATE
        
        'If the user has not suspended EXIF auto-rotation, request it from FreeImage
        If g_UserPreferences.GetPref_Boolean("Loading", "ExifAutoRotate", True) Then fi_ImportFlags = fi_ImportFlags Or FILO_JPEG_EXIFROTATE
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
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Importing image from file..."
        #End If
        
        fi_hDIB = FreeImage_Load(fileFIF, srcFilename, fi_ImportFlags)
        
    Else
        
        #If DEBUGMODE = 1 Then
            
            If fileFIF = FIF_GIF Then
                pdDebug.LogAction "Importing frame # " & pageToLoad + 1 & " from animated GIF file..."
            ElseIf fileFIF = FIF_ICO Then
                pdDebug.LogAction "Importing icon # " & pageToLoad + 1 & " from ICO file...", pageToLoad + 1
            Else
                pdDebug.LogAction "Importing page # " & pageToLoad + 1 & " from multipage TIFF file...", pageToLoad + 1
            End If
            
        #End If
        
        If fileFIF = FIF_GIF Then
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_GIF, srcFilename, , , , FILO_GIF_PLAYBACK)
        ElseIf fileFIF = FIF_ICO Then
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_ICO, srcFilename, , , , 0)
        Else
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
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Import via FreeImage failed (blank handle)."
        #End If
        
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
            If rQuad.Alpha <> 0 Then
                Dim fi_Palette() As Long
                fi_Palette = FreeImage_GetPaletteExLong(fi_hDIB)
                dstDIB.setBackgroundColor fi_Palette(rQuad.Alpha)
                
            'Otherwise it's easy - just reassemble the RGB values from the quad
            Else
                dstDIB.setBackgroundColor RGB(rQuad.Red, rQuad.Green, rQuad.Blue)
            End If
        
        End If
        
    End If
    
    
    '****************************************************************************
    ' Determine native color depth
    '****************************************************************************
    
    'Before we continue the import, we need to make sure the pixel data is in a format appropriate for PhotoDemon.
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Analyzing color depth..."
    #End If
    
    'First thing we want to check is the color depth.  PhotoDemon is designed around 16 million color images.  This could
    ' change in the future, but for now, force high-bit-depth images to a more appropriate 24 or 32bpp.
    Dim fi_BPP As Long
    fi_BPP = FreeImage_GetBPP(fi_hDIB)
    
    'Because it could be helpful later on, also retrieve the image datatype.  This is an internal FreeImage value
    ' corresponding to various data encodings (floating-point, complex, integer, etc).  If we ever want to handle
    ' high-bit-depth images, that value will be crucial!
    Dim fi_DataType As FREE_IMAGE_TYPE
    fi_DataType = FreeImage_GetImageType(fi_hDIB)
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Image bit-depth of " & fi_BPP & " and data type " & fi_DataType & " detected."
    #End If
    
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
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "High bit-depth RGB image identified.  Checking for non-standard alpha encoding..."
        #End If
        
        'While images with these bit-depths may not have an alpha channel, they can have binary transparency - check for that now.
        ' (Note: as of FreeImage 3.15.3 binary bit-depths are not detected correctly.  That said, they may someday be supported -
        ' so I've implemented two checks to cover both contingencies.
        fi_hasTransparency = FreeImage_IsTransparent(fi_hDIB)
        fi_transparentEntries = FreeImage_GetTransparencyCount(fi_hDIB)
    
        'As of 25 Nov '12, the user can choose to disable tone-mapping (which makes HDR loading much faster, but reduces image quality).
        ' Check that preference before tone-mapping the image.
        ' Also, as of 11 Sep '14, images with attached ICC profiles will preferentially use that over forcible tone-mapping.
        If g_UserPreferences.GetPref_Boolean("Loading", "HDR Tone Mapping", True) And (Not FreeImage_HasICCProfile(fi_hDIB)) Then
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Tone mapping HDR image to preserve tonal range..."
            #End If
            
            new_hDIB = FreeImage_ToneMapping(fi_hDIB, FITMO_REINHARD05)
            
            'Add a note to the target image that tone-mapping was forcibly applied to the incoming data
            If Not (targetImage Is Nothing) Then
                targetImage.imgStorage.Add "Tone-mapping", True
            End If
            
            'Immediately unload the original image copy (which is probably enormous, on account of its bit-depth)
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
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Tone mapping complete."
            #End If
        
        Else
                
            If fi_hasTransparency Or (fi_transparentEntries <> 0) Then
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Alpha found, but further tone-mapping ignored at user's request."
                #End If
                
                new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, False)
                
                'Immediately unload the original image copy (which is probably enormous, on account of its bit-depth)
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
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "No alpha found.  Further tone-mapping ignored at user's request."
                #End If
                
                '48bpp images can be converted automatically.  Unfortunately, in an absolutely massive oversight by the FreeImage team,
                ' 96bpp (RGBF) images cannot be auto-converted.  We must do it manually.
                Debug.Print fi_DataType, FIT_RGBF
                If (fi_DataType = FIT_RGBF) Then
                    new_hDIB = convertFreeImageRGBFTo24bppDIB(fi_hDIB)
                Else
                    new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_24BPP, False)
                End If
                
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
    
    'Because HDR Tone Mapping may not preserve alpha channels (the FreeImage documentation is unclear on this),
    ' we must do the same as above - manually make a copy of the image's alpha data, then reduce the image using tone mapping.
    ' Later in the process we will restore the alpha data to the image.
    If (fi_BPP = 64) Or (fi_BPP = 128) Then
    
        'Again, check for the user's preference on tone-mapping
        If g_UserPreferences.GetPref_Boolean("Loading", "HDR Tone Mapping", True) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "High bit-depth RGBA image identified.  Tone mapping HDR image to preserve tonal range..."
            #End If
        
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
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Tone mapping complete."
            #End If
            
        Else
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "High bit-depth RGBA image identified.  Tone-mapping ignored at user's request."
            #End If
            
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
            
            'Again, check for the user's preference on tone-mapping
            If g_UserPreferences.GetPref_Boolean("Loading", "HDR Tone Mapping", True) Then
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Tone-mapping high-bit-depth grayscale image to 24bpp..."
                #End If
                
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
            
            'User doesn't want tone-mapping, so perform a linear conversion to 24 or 32bpp
            Else
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "High bit-depth grayscale image identified.  Tone-mapping ignored at user's request."
                #End If
                
                'First, convert it to a high-bit depth RGB image
                fi_hDIB = FreeImage_ConvertToRGB16(fi_hDIB)
                
                'Look for transparency
                fi_hasTransparency = FreeImage_IsTransparent(fi_hDIB)
                fi_transparentEntries = FreeImage_GetTransparencyCount(fi_hDIB)
            
                'Convert to 24bpp or 32bpp as appropriate
                If fi_hasTransparency Or (fi_transparentEntries <> 0) Then
                    new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, False)
                Else
                    new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_24BPP, False)
                End If
                
                'Unload the original source
                If pageToLoad <= 0 Then
                    If (new_hDIB <> fi_hDIB) Then FreeImage_UnloadEx fi_hDIB
                Else
                    If (new_hDIB <> fi_hDIB) Then
                        needToCloseMulti = False
                        FreeImage_UnlockPage fi_multi_hDIB, fi_hDIB, False
                        FreeImage_CloseMultiBitmap fi_multi_hDIB
                    End If
                End If
                
                'Replace the main FreeImage DIB handle with the new copy
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
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "CMYK image detected.  Preparing transform into RGB space..."
        #End If
        
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
        If Color_Management.applyCMYKTransform(dstDIB.ICCProfile.getICCDataPointer, dstDIB.ICCProfile.getICCDataSize, tmpCMYKDIB, tmpRGBDIB, dstDIB.ICCProfile.getSourceRenderIntent) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Copying newly transformed sRGB data..."
            #End If
        
            'The transform was successful.  Copy the new sRGB data back into the FreeImage object, so the load process can continue.
            FreeImage_Unload fi_hDIB
            fi_hDIB = FreeImage_CreateFromDC(tmpRGBDIB.getDIBDC)
            fi_BPP = FreeImage_GetBPP(fi_hDIB)
            dstDIB.ICCProfile.markSuccessfulProfileApplication
            
        'Something went horribly wrong.  Re-load the image and use FreeImage to apply the CMYK -> RGB transform.
        Else
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "ICC-based CMYK transformation failed.  Falling back to default CMYK conversion..."
            #End If
        
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
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Requesting memory for image transfer..."
    #End If
    
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
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Import via FreeImage failed (couldn't create DIB)."
        #End If
        
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
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Memory secured.  Finalizing image load..."
    #End If
        
    'Copy the bits from the FreeImage DIB to our DIB
    SetDIBitsToDevice dstDIB.getDIBDC, 0, 0, fi_Width, fi_Height, 0, 0, 0, fi_Height, ByVal FreeImage_GetBits(fi_hDIB), ByVal FreeImage_GetInfo(fi_hDIB), 0&
    
    'Debug.Print fi_hDIB, fi_multi_hDIB
    
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
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Image load successful.  FreeImage released."
    #End If
    
    
    '****************************************************************************
    ' If necessary, restore any lost alpha data
    '****************************************************************************
    
    'We are almost done.  The last thing we need to do is restore the alpha values if this was a high-bit-depth image
    ' whose alpha data was lost during the tone-mapping phase.
    If tmpAlphaRequired Then
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Restoring alpha data..."
        #End If
        
        dstDIB.copyAlphaFromExistingDIB tmpAlphaDIB
        dstDIB.fixPremultipliedAlpha True
        tmpAlphaDIB.eraseDIB
        Set tmpAlphaDIB = Nothing
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Alpha data restored successfully."
        #End If
        
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

'Perform linear scaling of a 96bpp RGBF image to standard 24bpp.  Note that an intermediate pdDIB object is used for convenience, but the returned
' handle is to a FREEIMAGE OBJECT.
'
'Returns: a non-zero FreeImage 24bpp image handle if successful.  0 if unsuccessful.
'
'IMPORTANT NOTE: this function ONLY FREES THE INCOMING fi_Handle PARAMETER IF CONVERSION IS SUCCESSFUL.  If the function fails (returns 0),
' I assume the caller still wants the original handle so it can proceed accordingly.  Similarly, the 24bpp handle this function returns (if
' successful) must also be freed by the caller.  Ignore this, and the function will leak.
Private Function convertFreeImageRGBFTo24bppDIB(ByVal fi_Handle As Long, Optional ByVal toNormalize As PD_BOOL = PD_BOOL_AUTO, Optional ByVal ignoreNegative As Boolean = False) As Long
    
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
        mustNormalize = isNormalizeRequired(fi_Handle, minR, maxR, minG, maxG, minB, maxB)
    Else
        mustNormalize = (toNormalize = PD_BOOL_TRUE)
        If mustNormalize Then isNormalizeRequired fi_Handle, minR, maxR, minG, maxG, minB, maxB
    End If
    
    'I have no idea if normalization is supposed to include negative numbers or not; each high-bit-depth format has its own quirks, and none are
    ' clear on preferred defaults, so I'll leave this as manually settable for now.
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
    
    'This Single-type array will consistently be updated to point to the current line of pixels in the image (RGBF format, remember!)
    Dim srcImageData() As Single
    Dim srcSA As SAFEARRAY1D
    
    'Create a 24bpp DIB at the same size as the image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createBlank FreeImage_GetWidth(fi_Handle), FreeImage_GetHeight(fi_Handle), 24
    
    'Point a byte array at the temporary DIB
    Dim dstImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, tmpDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(tmpSA), 4
        
    'Iterate through each scanline in the source image, copying it to destination as we go.
    Dim iWidth As Long, iHeight As Long, iHeightInv As Long, iScanWidth As Long
    iWidth = FreeImage_GetWidth(fi_Handle) - 1
    iHeight = FreeImage_GetHeight(fi_Handle) - 1
    iScanWidth = FreeImage_GetWidth(fi_Handle) * 3
    
    'Due to the potential math involved in conversion (if gamma and other settings are being toggled), we need a lot of intermediate variables.
    ' Depending on the user's settings, some of these may go unused.
    Dim rSrcF As Double, gSrcF As Double, bSrcF As Double
    Dim rDstF As Double, gDstF As Double, bDstF As Double
    Dim rDstL As Long, gDstL As Long, bDstL As Long
    
    Dim x As Long, y As Long, QuickX As Long
    
    For y = 0 To iHeight
    
        'FreeImage DIBs are stored bottom-up; we invert them during processing
        iHeightInv = iHeight - y
    
        'Point a 1D VB array at this scanline
        With srcSA
            .cbElements = 4
            .cDims = 1
            .lBound = 0
            .cElements = iScanWidth
            .pvData = FreeImage_GetScanline(fi_Handle, y)
        End With
        CopyMemory ByVal VarPtrArray(srcImageData), VarPtr(srcSA), 4
        
        'Iterate through this line, converting values as we go
        For x = 0 To iWidth
            
            'Retrieve the source values.  This includes an implicit cast to Double, which I've done because some formats support IEEE constants
            ' like NaN or Infinity.  VB doesn't deal with these gracefully, and an implicit cast to Double seems to reduce unpredictable errors,
            ' possibly by giving any range-check code some breathing room.
            QuickX = x * 3
            rSrcF = CDbl(srcImageData(QuickX))
            gSrcF = CDbl(srcImageData(QuickX + 1))
            bSrcF = CDbl(srcImageData(QuickX + 2))
            
            'If normalization is required, apply it now
            If mustNormalize Then
                
                'If the caller has requested that we ignore negative values, clamp negative values to zero
                If ignoreNegative Then
                
                    If rSrcF < 0 Then rSrcF = 0
                    If gSrcF < 0 Then gSrcF = 0
                    If bSrcF < 0 Then bSrcF = 0
                
                'If negative values are considered valid, redistribute them on the range [0, Dist[Min, Max]]
                Else
                    rSrcF = rSrcF - minR
                    gSrcF = gSrcF - minG
                    bSrcF = bSrcF - minB
                End If
                
                'As a failsafe, make sure the image is not all black
                If rDist > 0 Then
                    rDstF = (rSrcF / rDist)
                    
                'If this channel is a single color, force it to black
                Else
                    rDstF = 0
                End If
                
                'Repeat for g and b channels
                If gDist > 0 Then
                    gDstF = (gSrcF / gDist)
                Else
                    gDstF = 0
                End If
                
                If bDist > 0 Then
                    bDstF = (bSrcF / bDist)
                Else
                    bDstF = 0
                End If
                
            'If an image does not need to be normalized, this step is much easier
            Else
                
                rDstF = rSrcF
                gDstF = gSrcF
                bDstF = bSrcF
                
            End If
                        
            'In the future, gamma correction, etc could be applied here.
            
            'Apply failsafe range checks now
            If rDstF < 0 Then
                rDstF = 0
            ElseIf rDstF > 1 Then
                rDstF = 1
            End If
                
            If gDstF < 0 Then
                gDstF = 0
            ElseIf gDstF > 1 Then
                gDstF = 1
            End If
                
            If bDstF < 0 Then
                bDstF = 0
            ElseIf bDstF > 1 Then
                bDstF = 1
            End If
            
            'Calculate corresponding integer values on the range [0, 255]
            rDstL = rDstF * 255
            gDstL = gDstF * 255
            bDstL = bDstF * 255
            
            'Technically, the RGB values should be guaranteed on [0, 255] at this point - but better safe than sorry when working with
            ' detailed floating-point conversions.
            If rDstL > 255 Then
                rDstL = 255
            ElseIf rDstL < 0 Then
                rDstL = 0
            End If
            
            If gDstL > 255 Then
                gDstL = 255
            ElseIf gDstL < 0 Then
                gDstL = 0
            End If
            
            If bDstL > 255 Then
                bDstL = 255
            ElseIf bDstL < 0 Then
                bDstL = 0
            End If
                        
            'Copy the final, safe values into the destination
            dstImageData(QuickX, iHeightInv) = bDstL
            dstImageData(QuickX + 1, iHeightInv) = gDstL
            dstImageData(QuickX + 2, iHeightInv) = rDstL
            
        Next x
        
        'Free our 1D array reference
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        
    Next y
    
    'Point dstImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    'Now that we are done with the original source FI DIB, free it
    FreeImage_Unload fi_Handle
    
    'Create a FreeImage object from our DIB, then release our DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    Set tmpDIB = Nothing
    
    'Success!
    convertFreeImageRGBFTo24bppDIB = fi_DIB

End Function

'Returns TRUE if an RGBF format FreeImage DIB contains values outside the [0, 1] range (meaning normalization is required).
' If normalization is required, the various min and max parameters will be filled for each channel.  It is up to the caller to determine how
' these values are used; this function is only diagnostic.
Private Function isNormalizeRequired(ByVal fi_Handle As Long, ByRef dstMinR As Double, ByRef dstMaxR As Double, ByRef dstMinG As Double, ByRef dstMaxG As Double, ByRef dstMinB As Double, ByRef dstMaxB As Double) As Boolean
    
    'Values within the [0, 1] range are considered normal.  Values outside this range are not normal, and normalization is thus required.
    ' Because an image does not have to include 0 or 1 values specifically, we return TRUE exclusively; e.g. if any value falls outside
    ' the [0, 1] range, normalization is required.
    Dim minR As Single, maxR As Single, minG As Single, maxG As Single, minB As Single, maxB As Single
    minR = 0: minG = 0: minB = 0
    maxR = 1: maxG = 1: maxB = 1
    
    'This Single-type array will consistently be updated to point to the current line of pixels in the image (RGBF format, remember!)
    Dim srcImageData() As Single
    Dim srcSA As SAFEARRAY1D
    
    'Iterate through each scanline in the source image, checking normalize parameters as we go.
    Dim iWidth As Long, iHeight As Long, iScanWidth As Long
    iWidth = FreeImage_GetWidth(fi_Handle) - 1
    iHeight = FreeImage_GetHeight(fi_Handle) - 1
    iScanWidth = FreeImage_GetWidth(fi_Handle) * 3
    
    Dim srcR As Single, srcG As Single, srcB As Single
    Dim x As Long, y As Long, QuickX As Long
    
    For y = 0 To iHeight
        
        'Point a 1D VB array at this scanline
        With srcSA
            .cbElements = 4
            .cDims = 1
            .lBound = 0
            .cElements = iScanWidth
            .pvData = FreeImage_GetScanline(fi_Handle, y)
        End With
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
        'Iterate through this line, checking values as we go
        For x = 0 To iWidth
            
            QuickX = x * 3
            
            srcR = srcImageData(QuickX)
            srcG = srcImageData(QuickX + 1)
            srcB = srcImageData(QuickX + 2)
            
            'Check max/min values independently for each channel
            If srcR < minR Then
                minR = srcR
            ElseIf srcR > maxR Then
                maxR = srcR
            End If
            
            If srcG < minG Then
                minG = srcG
            ElseIf srcG > maxG Then
                maxG = srcG
            End If
            
            If srcB < minB Then
                minB = srcB
            ElseIf srcB > maxB Then
                maxB = srcB
            End If
            
        Next x
        
        'Free our 1D array reference
        CopyMemory ByVal VarPtrArray(srcImageData()), 0&, 4
        
    Next y
    
    'Fill min/max RGB values regardless of normalization
    dstMinR = minR
    dstMaxR = maxR
    dstMinG = minG
    dstMaxG = maxG
    dstMinB = minB
    dstMaxB = maxB
    
    'If the max or min lie outside the image, notify the caller that normalization is required on this image
    If (minR < 0) Or (maxR > 1) Or (minG < 0) Or (maxG > 1) Or (minB < 0) Or (maxB > 1) Then
        isNormalizeRequired = True
    Else
        isNormalizeRequired = False
    End If
    
End Function

'Use FreeImage to resize a DIB.  (Technically, to copy a resized portion of a source image into a destination image.)
' The call is formatted similar to StretchBlt, as it used to replace StretchBlt when working with 32bpp data.
' This function is also declared identically to PD's GDI+ equivalent, specifically GDIPlusResizeDIB.  This was done
' so that the two functions can be used interchangeably.
Public Function FreeImageResizeDIB(ByRef dstDIB As pdDIB, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcDIB As pdDIB, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal interpolationType As FREE_IMAGE_FILTER, Optional ByVal destinationIsBlank As Boolean = False) As Boolean

    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Double
    'profileTime = Timer

    FreeImageResizeDIB = True

    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
                
        'Create a temporary DIB at the size of the source image
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.createBlank srcWidth, srcHeight, srcDIB.getDIBColorDepth, 0
        
        'Copy the relevant source portion of the image into the temporary DIB
        BitBlt tmpDIB.getDIBDC, 0, 0, srcWidth, srcHeight, srcDIB.getDIBDC, srcX, srcY, vbSrcCopy
        
        'Create a FreeImage copy of the temporary DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
        
        'Use that handle to request an image resize
        If fi_DIB <> 0 Then
            
            Dim returnDIB As Long
            returnDIB = FreeImage_RescaleByPixel(fi_DIB, dstWidth, dstHeight, True, interpolationType)
                        
            'Copy the bits from the FreeImage DIB to our DIB
            tmpDIB.createBlank dstWidth, dstHeight, 32, 0
            SetDIBitsToDevice tmpDIB.getDIBDC, 0, 0, dstWidth, dstHeight, 0, 0, 0, srcHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
            
            'If the destinationIsBlank flag is true, we can use BitBlt in place of AlphaBlend to copy the result
            ' onto the destination DIB; this shaves off a tiny bit of time.
            If destinationIsBlank Then
                BitBlt dstDIB.getDIBDC, dstX, dstY, dstWidth, dstHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
            Else
                AlphaBlend dstDIB.getDIBDC, dstX, dstY, dstWidth, dstHeight, tmpDIB.getDIBDC, 0, 0, dstWidth, dstHeight, 255 * &H10000 Or &H1000000
            End If
            
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            
        End If
                
    Else
        FreeImageResizeDIB = False
    End If
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format(CStr((Timer - profileTime) * 1000), "0000.00")
    
End Function

'Use FreeImage to resize a DIB, optimized against the use case where the full source image is being used.
' (Basically, something closer to BitBlt than StretchBlt, but without sourceX/Y parameters for an extra boost.)
Public Function FreeImageResizeDIBFast(ByRef dstDIB As pdDIB, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcDIB As pdDIB, ByVal interpolationType As FREE_IMAGE_FILTER, Optional ByVal destinationIsBlank As Boolean = False) As Boolean

    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Double
    'profileTime = Timer

    FreeImageResizeDIBFast = True

    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
        
        'Create a FreeImage copy of the source DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
        
        'Use that handle to request an image resize
        If fi_DIB <> 0 Then
            
            Dim returnDIB As Long
            returnDIB = FreeImage_RescaleByPixel(fi_DIB, dstWidth, dstHeight, True, interpolationType)
                        
            'Copy the bits from the FreeImage DIB to a temporary DIB
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createBlank dstWidth, dstHeight, 32, 0
            SetDIBitsToDevice tmpDIB.getDIBDC, 0, 0, dstWidth, dstHeight, 0, 0, 0, dstHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
            
            'If the destinationIsBlank flag is true, we can use BitBlt in place of AlphaBlend to copy the result
            ' onto the destination DIB; this shaves off a tiny bit of time.
            If destinationIsBlank Then
                BitBlt dstDIB.getDIBDC, dstX, dstY, dstWidth, dstHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
            Else
                AlphaBlend dstDIB.getDIBDC, dstX, dstY, dstWidth, dstHeight, tmpDIB.getDIBDC, 0, 0, dstWidth, dstHeight, 255 * &H10000 Or &H1000000
            End If
            
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            
        End If
                
    Else
        FreeImageResizeDIBFast = False
    End If
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format(CStr((Timer - profileTime) * 1000), "0000.00")
    
End Function

