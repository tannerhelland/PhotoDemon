Attribute VB_Name = "Plugin_FreeImage"
'***************************************************************************
'FreeImage Interface (Advanced)
'Copyright 2012-2016 by Tanner Helland
'Created: 3/September/12
'Last updated: 04/December/14
'Last update: overhaul all code related to high bit-depth images
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

Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean

'Additional variables for PD-specific tone-mapping functions
Private m_shoulderStrength As Double, m_linearStrength As Double, m_linearAngle As Double, m_linearWhitePoint As Single
Private m_toeStrength As Double, m_toeNumerator As Double, m_toeDenominator As Double, m_toeAngle As Double

'Cache for post-export image previews.  This array can be safely freed, as it will be properly initialized on-demand.
Private m_ExportPreviewBytes() As Byte
    
'Is FreeImage available as a plugin?  (NOTE: this is now determined separately from FreeImageEnabled.)
Public Function IsFreeImageAvailable() As Boolean
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(g_PluginPath & "freeimage.dll") Then IsFreeImageAvailable = True Else IsFreeImageAvailable = False
    
End Function

'Initialize FreeImage.  Do not call this until you have verified FreeImage's existence (typically via isFreeImageAvailable(), above)
Public Function InitializeFreeImage() As Boolean
    
    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    g_FreeImageHandle = LoadLibrary(g_PluginPath & "FreeImage.dll")
    InitializeFreeImage = CBool(g_FreeImageHandle <> 0)
    
End Function

    
'Load an image via FreeImage.  It is assumed that the source file has already been vetted for things like "does it exist?"
Public Function LoadFreeImageV4(ByVal srcFilename As String, ByRef dstDIB As pdDIB, Optional ByVal pageToLoad As Long = 0, Optional ByVal showMessages As Boolean = True, Optional ByRef targetImage As pdImage = Nothing) As PD_OPERATION_OUTCOME

    On Error GoTo FreeImageV4_AdvancedError
    
    '****************************************************************************
    ' Make sure FreeImage exists and is usable
    '****************************************************************************
    
    'Double-check that FreeImage.dll was located at start-up
    If Not g_ImageFormats.FreeImageEnabled Then
        LoadFreeImageV4 = PD_FAILURE_GENERIC
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
    fileFIF = FreeImage_GetFileTypeU(StrPtr(srcFilename))
    
    'For certain filetypes (CUT, MNG, PCD, TARGA and WBMP, according to the FreeImage documentation), the lack of a reliable
    ' header may prevent GetFileType from working.  As a result, double-check the file using its file extension.
    If fileFIF = FIF_UNKNOWN Then fileFIF = FreeImage_GetFIFFromFilenameU(StrPtr(srcFilename))
    
    'By this point, if the file still doesn't show up in FreeImage's database, abandon the import attempt.
    If Not FreeImage_FIFSupportsReading(fileFIF) Then
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Filetype not supported by FreeImage.  Import abandoned."
        #End If
        
        LoadFreeImageV4 = PD_FAILURE_GENERIC
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
    If fileFIF = PDIF_JPEG Then
        fi_ImportFlags = fi_ImportFlags Or FILO_JPEG_ACCURATE
        
        'If the user has not suspended EXIF auto-rotation, request it from FreeImage
        If g_UserPreferences.GetPref_Boolean("Loading", "ExifAutoRotate", True) Then fi_ImportFlags = fi_ImportFlags Or FILO_JPEG_EXIFROTATE
    End If
    
    'For PNG files, request that gamma is ignored (we will handle it ourselves, later in the load process)
    If fileFIF = PDIF_PNG Then
        fi_ImportFlags = fi_ImportFlags Or FILO_PNG_IGNOREGAMMA
    End If
    
    'Check for CMYK JPEGs, TIFFs, and PSD files.  If an image is CMYK and an ICC profile is present, ask FreeImage to load the
    ' raw CMYK data. If no ICC profile is present, FreeImage is free to perform the CMYK -> RGB translation for us.
    Dim isCMYK As Boolean
    isCMYK = False
    
    If (fileFIF = PDIF_JPEG) Or (fileFIF = PDIF_PSD) Or (fileFIF = PDIF_TIFF) Then
    
        'To speed up the load process, only load the file header, and explicitly instruct FreeImage to leave CMYK images
        ' in CMYK format (otherwise we can't detect CMYK, as it will be auto-converted to RGB!).
        Dim additionalFlags As Long
        additionalFlags = FILO_LOAD_NOPIXELS
        
        Select Case fileFIF
        
            Case PDIF_JPEG
                additionalFlags = additionalFlags Or FILO_JPEG_CMYK
            
            Case PDIF_PSD
                additionalFlags = additionalFlags Or FILO_PSD_CMYK
            
            Case PDIF_TIFF
                additionalFlags = additionalFlags Or TIFF_CMYK
        
        End Select
    
        Dim tmpFIHandle As Long
        tmpFIHandle = FreeImage_LoadUInt(fileFIF, StrPtr(srcFilename), fi_ImportFlags Or additionalFlags)
        
        'Check the file's color type
        If FreeImage_GetColorType(tmpFIHandle) = FIC_CMYK Then
        
            'File is CMYK.  Check for an ICC profile.
            If FreeImage_HasICCProfile(tmpFIHandle) Then
            
                'CMYK + ICC profile means we want FreeImage to load the data in CMYK format, and we'll perform the conversion to sRGB ourselves.
                isCMYK = True
                
                Select Case fileFIF
        
                    Case PDIF_JPEG
                        fi_ImportFlags = fi_ImportFlags Or FILO_JPEG_CMYK
                    
                    Case PDIF_PSD
                        fi_ImportFlags = fi_ImportFlags Or FILO_PSD_CMYK
                    
                    Case PDIF_TIFF
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
        If (Not showMessages) Then fi_ImportFlags = fi_ImportFlags Or FILO_RAW_PREVIEW
        
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
        
        fi_hDIB = FreeImage_LoadUInt(fileFIF, StrPtr(srcFilename), fi_ImportFlags)
        
    Else
        
        #If DEBUGMODE = 1 Then
            
            If fileFIF = PDIF_GIF Then
                pdDebug.LogAction "Importing frame # " & pageToLoad + 1 & " from animated GIF file..."
            ElseIf fileFIF = FIF_ICO Then
                pdDebug.LogAction "Importing icon # " & pageToLoad + 1 & " from ICO file...", pageToLoad + 1
            Else
                pdDebug.LogAction "Importing page # " & pageToLoad + 1 & " from multipage TIFF file...", pageToLoad + 1
            End If
            
        #End If
        
        If fileFIF = PDIF_GIF Then
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(PDIF_GIF, srcFilename, , , , FILO_GIF_PLAYBACK)
        ElseIf fileFIF = FIF_ICO Then
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(FIF_ICO, srcFilename, , , , 0)
        Else
            fi_multi_hDIB = FreeImage_OpenMultiBitmap(PDIF_TIFF, srcFilename, , , , 0)
        End If
        
        fi_hDIB = FreeImage_LockPage(fi_multi_hDIB, pageToLoad)
        
    End If
    
    'Store this original, untouched color depth now
    If fi_hDIB <> 0 Then dstDIB.setOriginalFreeImageColorDepth FreeImage_GetBPP(fi_hDIB)
    
    'Icon files may use a simple mask for their alpha channel; in this case, re-load the icon with the FILO_ICO_MAKEALPHA flag
    If fileFIF = FIF_ICO Then
        
        'Check the bit-depth
        If FreeImage_GetBPP(fi_hDIB) < 32 Then
        
            'If this is the first frame of the icon, unload it and try again
            If (pageToLoad <= 0) Then
                If fi_hDIB <> 0 Then FreeImage_UnloadEx fi_hDIB
                fi_hDIB = FreeImage_LoadUInt(fileFIF, StrPtr(srcFilename), FILO_ICO_MAKEALPHA)
            
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
        
        LoadFreeImageV4 = PD_FAILURE_GENERIC
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
        dstDIB.ICCProfile.LoadICCFromFreeImage fi_hDIB
        
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
            If rQuad.alpha <> 0 Then
                Dim fi_Palette() As Long
                fi_Palette = FreeImage_GetPaletteExLong(fi_hDIB)
                dstDIB.setBackgroundColor fi_Palette(rQuad.alpha)
                
            'Otherwise it's easy - just reassemble the RGB values from the quad
            Else
                dstDIB.setBackgroundColor RGB(rQuad.Red, rQuad.Green, rQuad.Blue)
            End If
        
        End If
     
    'No background color found; write -1 to notify of this.
    Else
        dstDIB.setBackgroundColor -1
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
    
    'A number of other variables may be required as we adjust the bit-depth of the image to match PhotoDemon's internal requirements.
    Dim new_hDIB As Long
    
    Dim fi_hasTransparency As Boolean
    Dim fi_transparentEntries As Long
    
    
    '****************************************************************************
    ' If the image is high bit-depth (e.g. > 8 bits per channel), downsample it to a standard 24 or 32bpp image.
    '****************************************************************************
    
    If fi_DataType <> FIT_Bitmap Then
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "HDR image identified.  Raising tone-map dialog..."
        #End If
    
        'Use the central tone-map handler to apply further tone-mapping
        Dim toneMappingOutcome As PD_OPERATION_OUTCOME
        toneMappingOutcome = RaiseToneMapDialog(fi_hDIB, new_hDIB)
        
        'A non-zero return signifies a successful tone-map operation.  Unload our old handle, and proceed with the new handle
        If (toneMappingOutcome = PD_SUCCESS) And (new_hDIB <> 0) Then
            
            'Add a note to the target image that tone-mapping was forcibly applied to the incoming data
            If Not (targetImage Is Nothing) Then
                targetImage.imgStorage.AddEntry "Tone-mapping", True
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
            
            'Replace the primary FI_DIB handle with the new one, then carry on with loading
            fi_hDIB = new_hDIB
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Tone mapping complete."
            #End If
        
        'The tone-mapper will return 0 if it failed.  If this happens, we cannot proceed with loading.
        Else
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Tone-mapping canceled due to user request or error.  Abandoning image import."
            #End If
            
            If fi_hDIB <> 0 Then FreeImage_Unload fi_hDIB
            
            If toneMappingOutcome <> PD_SUCCESS Then
                LoadFreeImageV4 = toneMappingOutcome
            Else
                LoadFreeImageV4 = PD_FAILURE_GENERIC
            End If
            
            Exit Function
        
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
            If fi_hasTransparency Then
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
        If Color_Management.ApplyCMYKTransform(dstDIB.ICCProfile.GetICCDataPointer, dstDIB.ICCProfile.GetICCDataSize, tmpCMYKDIB, tmpRGBDIB, dstDIB.ICCProfile.GetSourceRenderIntent) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Copying newly transformed sRGB data..."
            #End If
        
            'The transform was successful.  Copy the new sRGB data back into the FreeImage object, so the load process can continue.
            FreeImage_Unload fi_hDIB
            fi_hDIB = FreeImage_CreateFromDC(tmpRGBDIB.getDIBDC)
            fi_BPP = FreeImage_GetBPP(fi_hDIB)
            dstDIB.ICCProfile.MarkSuccessfulProfileApplication
            
        'Something went horribly wrong.  Re-load the image and use FreeImage to apply the CMYK -> RGB transform.
        Else
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "ICC-based CMYK transformation failed.  Falling back to default CMYK conversion..."
            #End If
        
            FreeImage_Unload fi_hDIB
            fi_hDIB = FreeImage_LoadUInt(fileFIF, StrPtr(srcFilename), FILO_JPEG_ACCURATE Or FILO_JPEG_EXIFROTATE)
            fi_BPP = FreeImage_GetBPP(fi_hDIB)
        
        End If
        
        Set tmpCMYKDIB = Nothing
        Set tmpRGBDIB = Nothing
    
    End If
    
    
    '****************************************************************************
    ' PD's new rendering engine requires pre-multiplied alpha values.  Apply premultiplication now - but ONLY if
    ' the image did not come from the clipboard.  (Clipboard images requires special treatment.)
    '****************************************************************************
    
    Dim specialClipboardHandlingRequired As Boolean
    Dim tmpClipboardInfo As PD_Clipboard_Info
    
    specialClipboardHandlingRequired = False
    
    If fi_BPP = 32 Then
        
        'If the clipboard is active, this image came from a Paste operation.  It may require extra alpha heuristics.
        If g_Clipboard.IsClipboardOpen Then
        
            'Retrieve a local copy of PD's clipboard info struct.  We're going to analyze it, to see if we need to
            ' run some alpha heuristics (because the clipboard is shit when it comes to handling alpha correctly.)
            tmpClipboardInfo = g_Clipboard.GetClipboardInfo
            
            'If the clipboard image was originally placed on the clipboard as a DDB, a whole variety of driver-specific
            ' issues may be present.
            If tmpClipboardInfo.pdci_OriginalFormat = CF_BITMAP Then
            
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
        LoadFreeImageV4 = PD_FAILURE_GENERIC
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
        
        LoadFreeImageV4 = PD_FAILURE_GENERIC
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
    ' Finalize alpha values in the target image
    '****************************************************************************
    
    'If this image came from the clipboard, and its alpha state is unknown, we're going to force all alpha values
    ' to 255 to avoid potential driver-specific issues with the PrtScrn key.
    If specialClipboardHandlingRequired Then
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Image came from the clipboard; finalizing alpha now..."
        #End If
        
        dstDIB.ForceNewAlpha 255
    
    End If
    
    'Regardless of original bit-depth, the final PhotoDemon image will always be 32-bits, with pre-multiplied alpha.
    dstDIB.setInitialAlphaPremultiplicationState True
    
    
    '****************************************************************************
    ' Load complete
    '****************************************************************************
    
    'Mark this load as successful
    LoadFreeImageV4 = PD_SUCCESS
    
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
    LoadFreeImageV4 = PD_FAILURE_GENERIC
    
End Function

'See if an image file is actually comprised of multiple files (e.g. animated GIFs, multipage TIFs).
' Input: file name to be checked
' Returns: 0 if only one image is found.  Page (or frame) count if multiple images are found.
Public Function IsMultiImage(ByVal srcFilename As String) As Long

    On Error GoTo isMultiImage_Error
    
    'Double-check that FreeImage.dll was located at start-up
    If Not g_ImageFormats.FreeImageEnabled Then
        IsMultiImage = 0
        Exit Function
    End If
        
    'Determine the file type.  (Currently, this feature only works on animated GIFs, multipage TIFFs, and icons.)
    Dim fileFIF As FREE_IMAGE_FORMAT
    fileFIF = FreeImage_GetFileTypeU(StrPtr(srcFilename))
    If fileFIF = FIF_UNKNOWN Then fileFIF = FreeImage_GetFIFFromFilenameU(StrPtr(srcFilename))
    
    'If FreeImage can't determine the file type, or if the filetype is not GIF or TIF, return False
    If (Not FreeImage_FIFSupportsReading(fileFIF)) Or ((fileFIF <> PDIF_GIF) And (fileFIF <> PDIF_TIFF) And (fileFIF <> FIF_ICO)) Then
        IsMultiImage = 0
        Exit Function
    End If
    
    'At this point, we are guaranteed that the image is a GIF, TIFF, or icon file.
    ' Open the file using the multipage function
    Dim fi_multi_hDIB As Long
    If fileFIF = PDIF_GIF Then
        fi_multi_hDIB = FreeImage_OpenMultiBitmap(PDIF_GIF, srcFilename)
    ElseIf fileFIF = FIF_ICO Then
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

'Given a FreeImage handle, return a 24 or 32bpp pdDIB object, as relevant.  Note that this function does not modify premultiplication
' status of 32bpp images.  The caller is responsible for applying that (as necessary).
'
'NOTE!  This function requires the FreeImage DIB to already be in 24 or 32bpp format.  It will fail if another bit-depth is used.
'ALSO NOTE!  This function does not set alpha premultiplication.  It's assumed that the caller knows that value in advance.
'ALSO NOTE!  This function does not free the incoming FreeImage handle, by design.
Public Function GetPDDibFromFreeImageHandle(ByVal srcFI_Handle As Long, ByRef dstDIB As pdDIB) As Boolean
    
    Dim fiHandleBackup As Long
    fiHandleBackup = srcFI_Handle
    
    'Double-check the FreeImage handle's bit depth
    Dim fiBPP As Long
    fiBPP = FreeImage_GetBPP(srcFI_Handle)
    
    If (fiBPP <> 24) And (fiBPP <> 32) Then
        
        'If the DIB is less than 24 bpp, upsample now
        If fiBPP < 24 Then
            
            'Conversion to higher bit depths is contingent on the presence of an alpha channel
            If FreeImage_IsTransparent(srcFI_Handle) Or (FreeImage_GetTransparentIndex(srcFI_Handle) <> -1) Then
                srcFI_Handle = FreeImage_ConvertColorDepth(srcFI_Handle, FICF_RGB_32BPP, False)
            Else
                srcFI_Handle = FreeImage_ConvertColorDepth(srcFI_Handle, FICF_RGB_24BPP, False)
            End If
            
            'Verify the new bit-depth
            fiBPP = FreeImage_GetBPP(srcFI_Handle)
            
            If (fiBPP <> 24) And (fiBPP <> 32) Then
                
                'If a new DIB was created, release it now.  (Note that the caller must still free the original handle.)
                If (srcFI_Handle <> 0) And (srcFI_Handle <> fiHandleBackup) Then FreeImage_Unload srcFI_Handle
                
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
    dstDIB.createBlank fi_Width, fi_Height, fiBPP, 0
    SetDIBitsToDevice dstDIB.getDIBDC, 0, 0, fi_Width, fi_Height, 0, 0, 0, fi_Height, ByVal FreeImage_GetBits(srcFI_Handle), ByVal FreeImage_GetInfo(srcFI_Handle), 0&
    
    'If we created a temporary DIB, free it now
    If srcFI_Handle <> fiHandleBackup Then
        FreeImage_Unload srcFI_Handle
        srcFI_Handle = fiHandleBackup
    End If
    
    GetPDDibFromFreeImageHandle = True
    
End Function

'Given a PD DIB, return a 24 or 32bpp FreeImage handle that simply WRAPS the DIB without copying it.  This is much faster (and less
' resource-intensive) than copying of the entire pixel array.  For situations where you only need non-destructive FreeImage behavior
' (like saving a DIB to file in some non-BMP format), please use this function.
'
'ALSO NOTE!  This function does not affect alpha premultiplication.  It's assumed that the caller sets that value in advance.
'ALSO NOTE!  The reverseScanlines parameter will be applied to the underlying pdDIB object - plan accordingly!
'ALSO NOTE!  This function does not free the outgoing FreeImage handle, by design.  Make sure to free it manually!
'ALSO NOTE!  The function returns zero for failure state; please check the return value before trying to use it!
Public Function GetFIHandleFromPDDib_NoCopy(ByRef srcDIB As pdDIB, Optional ByVal reverseScanlines As Boolean = False) As Long
    With srcDIB
        GetFIHandleFromPDDib_NoCopy = Outside_FreeImageV3.FreeImage_ConvertFromRawBitsEx(False, .getActualDIBBits, FIT_Bitmap, .getDIBWidth, .getDIBHeight, .getDIBArrayWidth, .getDIBColorDepth, , , , reverseScanlines)
    End With
End Function

'Paint a FreeImage DIB to an arbitrary clipping rect on some target pdDIB.  This does not free or otherwise modify the source FreeImage object
Public Function PaintFIDibToPDDib(ByRef dstDIB As pdDIB, ByVal fi_Handle As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long) As Boolean
    
    If (Not (dstDIB Is Nothing)) And (fi_Handle <> 0) Then
        Dim bmpInfo As BITMAPINFO
        Outside_FreeImageV3.FreeImage_GetInfoHeaderEx fi_Handle, VarPtr(bmpInfo.bmiHeader)
        If dstDIB.IsDIBTopDown Then bmpInfo.bmiHeader.biHeight = -1 * (bmpInfo.bmiHeader.biHeight)
        
        Dim iHeight As Long: iHeight = Abs(bmpInfo.bmiHeader.biHeight)
        PaintFIDibToPDDib = (SetDIBitsToDevice(dstDIB.getDIBDC, dstX, dstY, dstWidth, dstHeight, 0, 0, 0, iHeight, ByVal FreeImage_GetBits(fi_Handle), bmpInfo, 0&) <> 0)
    End If
    
End Function
'Prior to applying tone-mapping settings, query the user for their preferred behavior.  If the user doesn't want this dialog raised, this
' function will silently retrieve the proper settings from the preference file, and proceed with tone-mapping automatically.
'
'Returns: fills dst_fiHandle with a non-zero FreeImage 24 or 32bpp image handle if successful.  0 if unsuccessful.
'         The function itself will return a PD_OPERATION_OUTCOME value; this is important for determining if the user canceled the dialog.
'
'IMPORTANT NOTE!  If this function fails, further loading of the image must be halted.  PD cannot yet operate on anything larger than 32bpp,
' so if tone-mapping fails, we must abandon loading completely.  (A failure state can also be triggered by the user canceling the
' tone-mapping dialog.)
Private Function RaiseToneMapDialog(ByVal fi_Handle As Long, ByRef dst_fiHandle As Long) As PD_OPERATION_OUTCOME

    'Ask the user how they want to proceed.  Note that the dialog wrapper automatically handles the case of "do not prompt;
    ' use previous settings."  If that happens, it will retrieve the proper conversion settings for us, and return a dummy
    ' value of OK (as if the dialog were actually raised).
    Dim howToProceed As VbMsgBoxResult, toneMapSettings As String
    howToProceed = Dialog_Handler.PromptToneMapSettings(fi_Handle, toneMapSettings)
    
    'Check for a cancellation state; if encountered, abandon ship now.
    If (howToProceed <> vbOK) Then
        
        Debug.Print "Tone-map dialog appears to have been cancelled; result = " & howToProceed
        dst_fiHandle = 0
        RaiseToneMapDialog = PD_FAILURE_USER_CANCELED
        Exit Function
    
    'The ToneMapSettings string will now contain all the information we need to proceed with the tone-map.  Forward it to the
    ' central tone-mapping handler and use its success/fail state for this function as well.
    Else
        
        Debug.Print "Tone-map dialog appears to have been successful; result = " & howToProceed
        Message "Applying tone-mapping..."
        dst_fiHandle = ApplyToneMapping(fi_Handle, toneMapSettings)
        
        If dst_fiHandle = 0 Then
            Debug.Print "WARNING!  ApplyToneMapping() failed for reasons unknown."
            RaiseToneMapDialog = PD_FAILURE_GENERIC
        Else
            RaiseToneMapDialog = PD_SUCCESS
        End If
        
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
Public Function ApplyToneMapping(ByVal fi_Handle As Long, ByVal toneMapSettings As String) As Long
    
    'Retrieve the source image's bit-depth and data type.  These are crucial to successful tone-mapping operations.
    Dim fi_BPP As Long
    fi_BPP = FreeImage_GetBPP(fi_Handle)
    
    Dim fi_DataType As FREE_IMAGE_TYPE
    fi_DataType = FreeImage_GetImageType(fi_Handle)
    
    'Also, check for transparency in the source image.
    Dim hasTransparency As Boolean, transparentEntries As Long
    hasTransparency = FreeImage_IsTransparent(fi_Handle)
    transparentEntries = FreeImage_GetTransparencyCount(fi_Handle)
    If transparentEntries > 0 Then hasTransparency = True
    
    Dim newHandle As Long, rgbfHandle As Long
    
    'toneMapSettings contains all conversion instructions.  Parse it to determine which tone-map function to use.
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.SetParamString toneMapSettings
    
    'The first parameter contains the requested tone-mapping operation.
    Select Case cParams.GetLong(1)
    
        'Linear map
        Case PDTM_LINEAR
                
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
                
                If rgbfHandle = 0 Then
                    Debug.Print "WARNING!  FreeImage_ConvertToRGBA/F failed for reasons unknown."
                    ApplyToneMapping = 0
                    Exit Function
                Else
                    Debug.Print "FreeImage_ConvertToRGBA/F successful.  Proceeding with manual tone-mapping operation."
                End If
                
                newHandle = rgbfHandle
                
            End If
            
            'At this point, fi_Handle now represents a 32-bpc RGBF (or RGBAF) type FreeImage DIB.  Apply manual tone-mapping now.
            newHandle = ConvertFreeImageRGBFTo24bppDIB(newHandle, cParams.GetLong(3), cParams.GetBool(4), cParams.GetDouble(2))
            
            'Unload the intermediate RGBF handle as necessary
            If rgbfHandle <> 0 Then FreeImage_Unload rgbfHandle
            
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
                
                If rgbfHandle = 0 Then
                    Debug.Print "WARNING!  FreeImage_ConvertToRGBA/F failed for reasons unknown."
                    ApplyToneMapping = 0
                    Exit Function
                Else
                    Debug.Print "FreeImage_ConvertToRGBA/F successful.  Proceeding with manual tone-mapping operation."
                End If
                
                newHandle = rgbfHandle
                
            End If
            
            'At this point, fi_Handle now represents a 24bpp RGBF type FreeImage DIB.  Apply manual tone-mapping now.
            newHandle = ToneMapFilmic_RGBFTo24bppDIB(newHandle, cParams.GetDouble(2), cParams.GetDouble(3), , , , , , , cParams.GetDouble(4))
            
            'Unload the intermediate RGBF handle as necessary
            If rgbfHandle <> 0 Then FreeImage_Unload rgbfHandle
            
            ApplyToneMapping = newHandle
        
        'Adaptive logarithmic map
        Case PDTM_DRAGO
            ApplyToneMapping = FreeImage_TmoDrago03(fi_Handle, cParams.GetDouble(2), cParams.GetDouble(3))
            
        'Photoreceptor map
        Case PDTM_REINHARD
            ApplyToneMapping = FreeImage_TmoReinhard05Ex(fi_Handle, cParams.GetDouble(2), ByVal 0#, cParams.GetDouble(3), cParams.GetDouble(4))
        
    
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
        Debug.Print "Tone-mapping request invalid"
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
    
    'Create a 24bpp or 32bpp DIB at the same size as the image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    If fi_DataType = FIT_RGBF Then
        tmpDIB.createBlank FreeImage_GetWidth(fi_Handle), FreeImage_GetHeight(fi_Handle), 24
    Else
        tmpDIB.createBlank FreeImage_GetWidth(fi_Handle), FreeImage_GetHeight(fi_Handle), 32
    End If
    
    'Point a byte array at the temporary DIB
    Dim dstImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    PrepSafeArray tmpSA, tmpDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(tmpSA), 4
        
    'Iterate through each scanline in the source image, copying it to destination as we go.
    Dim iWidth As Long, iHeight As Long, iHeightInv As Long, iScanWidth As Long
    iWidth = FreeImage_GetWidth(fi_Handle) - 1
    iHeight = FreeImage_GetHeight(fi_Handle) - 1
    iScanWidth = FreeImage_GetPitch(fi_Handle)
    
    Dim qvDepth As Long
    If fi_DataType = FIT_RGBF Then qvDepth = 3 Else qvDepth = 4
    
    'Prep any other post-processing adjustments
    Dim gammaCorrection As Double
    gammaCorrection = 1 / newGamma
    
    'Due to the potential math involved in conversion (if gamma and other settings are being toggled), we need a lot of intermediate variables.
    ' Depending on the user's settings, some of these may go unused.
    Dim rSrcF As Double, gSrcF As Double, bSrcF As Double
    Dim rDstF As Double, gDstF As Double, bDstF As Double
    Dim rDstL As Long, gDstL As Long, bDstL As Long
    
    'Alpha is also a possibility, but unlike RGB values, we assume it is always normalized.  This allows us to skip any intermediate processing,
    ' and simply copy the value directly into the destination (after redistributing to the proper range, of course).
    Dim aDstF As Double, aDstL As Long
    
    Dim x As Long, y As Long, quickX As Long
    
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
            quickX = x * qvDepth
            rSrcF = CDbl(srcImageData(quickX))
            gSrcF = CDbl(srcImageData(quickX + 1))
            bSrcF = CDbl(srcImageData(quickX + 2))
            If qvDepth = 4 Then aDstF = CDbl(srcImageData(quickX + 3))
            
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
            
            'FYI, alpha is always unnormalized
                        
            'Apply gamma now (if any).  Unfortunately, lookup tables aren't an option because we're dealing with floating-point input,
            ' so this step is a little slow due to the exponent operator.
            If newGamma <> 1# Then
                If rDstF > 0 Then rDstF = rDstF ^ gammaCorrection
                If gDstF > 0 Then gDstF = gDstF ^ gammaCorrection
                If bDstF > 0 Then bDstF = bDstF ^ gammaCorrection
            End If
            
            'In the future, additional corrections could be applied here.
            
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
            
            'Handle alpha, if necessary
            If qvDepth = 4 Then
                If aDstF > 1 Then aDstF = 1
                If aDstF < 0 Then aDstF = 0
                aDstL = aDstF * 255
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
            dstImageData(quickX, iHeightInv) = bDstL
            dstImageData(quickX + 1, iHeightInv) = gDstL
            dstImageData(quickX + 2, iHeightInv) = rDstL
            If qvDepth = 4 Then dstImageData(quickX + 3, iHeightInv) = aDstL
            
        Next x
        
        'Free our 1D array reference
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        
    Next y
    
    'Point dstImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    'Create a FreeImage object from our pdDIB object, then release our pdDIB copy
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    Set tmpDIB = Nothing
    
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
Private Function ToneMapFilmic_RGBFTo24bppDIB(ByVal fi_Handle As Long, Optional ByVal newGamma As Single = 2.2, Optional ByVal exposureCompensation As Single = 2#, Optional ByVal shoulderStrength As Single = 0.22, Optional ByVal linearStrength As Single = 0.3, Optional ByVal linearAngle As Single = 0.1, Optional ByVal toeStrength As Single = 0.2, Optional ByVal toeNumerator As Single = 0.01, Optional ByVal toeDenominator As Single = 0.3, Optional ByVal linearWhitePoint As Single = 11.2) As Long
    
    'Before doing anything, check the incoming fi_Handle.  For performance reasons, this function only handles RGBF and RGBAF formats.
    ' Other formats are invalid.
    Dim fi_DataType As FREE_IMAGE_TYPE
    fi_DataType = FreeImage_GetImageType(fi_Handle)
    
    If (fi_DataType <> FIT_RGBF) And (fi_DataType <> FIT_RGBAF) Then
        Debug.Print "Tone-mapping request invalid"
        ToneMapFilmic_RGBFTo24bppDIB = 0
        Exit Function
    End If
    
    'This Single-type array will consistently be updated to point to the current line of pixels in the image (RGBF format, remember!)
    Dim srcImageData() As Single
    Dim srcSA As SAFEARRAY1D
    
    'Create a 24bpp or 32bpp DIB at the same size as the image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    If fi_DataType = FIT_RGBF Then
        tmpDIB.createBlank FreeImage_GetWidth(fi_Handle), FreeImage_GetHeight(fi_Handle), 24
    Else
        tmpDIB.createBlank FreeImage_GetWidth(fi_Handle), FreeImage_GetHeight(fi_Handle), 32
    End If
    
    'Point a byte array at the temporary DIB
    Dim dstImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    PrepSafeArray tmpSA, tmpDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(tmpSA), 4
        
    'Iterate through each scanline in the source image, copying it to destination as we go.
    Dim iWidth As Long, iHeight As Long, iHeightInv As Long, iScanWidth As Long
    iWidth = FreeImage_GetWidth(fi_Handle) - 1
    iHeight = FreeImage_GetHeight(fi_Handle) - 1
    iScanWidth = FreeImage_GetPitch(fi_Handle)
    
    Dim qvDepth As Long
    If fi_DataType = FIT_RGBF Then qvDepth = 3 Else qvDepth = 4
    
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
    
    'Prep any other post-processing adjustments
    Dim gammaCorrection As Double
    gammaCorrection = 1 / newGamma
    
    'Due to the potential math involved in conversion (if gamma and other settings are being toggled), we need a lot of intermediate variables.
    ' Depending on the user's settings, some of these may go unused.
    Dim rSrcF As Single, gSrcF As Single, bSrcF As Single
    Dim rDstF As Single, gDstF As Single, bDstF As Single
    Dim rDstL As Long, gDstL As Long, bDstL As Long
    
    'Alpha is also a possibility, but unlike RGB values, we assume it is always normalized.  This allows us to skip any intermediate processing,
    ' and simply copy the value directly into the destination (after redistributing to the proper range, of course).
    Dim aDstF As Double, aDstL As Long
    
    Dim x As Long, y As Long, quickX As Long
    
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
            
            'Retrieve the source values.
            quickX = x * qvDepth
            rSrcF = srcImageData(quickX)
            gSrcF = srcImageData(quickX + 1)
            bSrcF = srcImageData(quickX + 2)
            If qvDepth = 4 Then aDstF = CDbl(srcImageData(quickX + 3))
            
            'Apply filmic tone-mapping.  See http://fr.slideshare.net/ozlael/hable-john-uncharted2-hdr-lighting for details
            rDstF = fFilmicTonemap(exposureCompensation * rSrcF) / fWhitePoint
            gDstF = fFilmicTonemap(exposureCompensation * gSrcF) / fWhitePoint
            bDstF = fFilmicTonemap(exposureCompensation * bSrcF) / fWhitePoint
                                    
            'Apply gamma now (if any).  Unfortunately, lookup tables aren't an option because we're dealing with floating-point input,
            ' so this step is a little slow due to the exponent operator.
            If newGamma <> 1# Then
                If rDstF > 0 Then rDstF = rDstF ^ gammaCorrection
                If gDstF > 0 Then gDstF = gDstF ^ gammaCorrection
                If bDstF > 0 Then bDstF = bDstF ^ gammaCorrection
            End If
                        
            'Apply failsafe range checks
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
            
            'Handle alpha, if necessary
            If qvDepth = 4 Then
                If aDstF > 1 Then aDstF = 1
                If aDstF < 0 Then aDstF = 0
                aDstL = aDstF * 255
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
            dstImageData(quickX, iHeightInv) = bDstL
            dstImageData(quickX + 1, iHeightInv) = gDstL
            dstImageData(quickX + 2, iHeightInv) = rDstL
            If qvDepth = 4 Then dstImageData(quickX + 3, iHeightInv) = aDstL
            
        Next x
        
        'Free our 1D array reference
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        
    Next y
    
    'Point dstImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    'Create a FreeImage object from our pdDIB object, then release our pdDIB copy
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpDIB.getDIBDC)
    
    Set tmpDIB = Nothing
    
    'Success!
    ToneMapFilmic_RGBFTo24bppDIB = fi_DIB

End Function

'Filmic tone-map function
Private Function fFilmicTonemap(ByRef x As Single) As Single
    
    'In advance, calculate the filmic function for the white point
    Dim numFunction As Double, denFunction As Double
    
    numFunction = x * (m_shoulderStrength * x + m_linearStrength * m_linearAngle) + m_toeStrength * m_toeNumerator
    denFunction = x * (m_shoulderStrength * x + m_linearStrength) + m_toeStrength * m_toeDenominator
    
    'Failsafe check for DBZ errors
    If denFunction > 0 Then
        fFilmicTonemap = (numFunction / denFunction) - m_toeAngle
    Else
        fFilmicTonemap = 1
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
    Dim srcImageData() As Single
    Dim srcSA As SAFEARRAY1D
    
    'Iterate through each scanline in the source image, checking normalize parameters as we go.
    Dim iWidth As Long, iHeight As Long, iScanWidth As Long
    iWidth = FreeImage_GetWidth(fi_Handle) - 1
    iHeight = FreeImage_GetHeight(fi_Handle) - 1
    iScanWidth = FreeImage_GetPitch(fi_Handle)
    
    Dim qvDepth As Long
    If fi_DataType = FIT_RGBF Then qvDepth = 3 Else qvDepth = 4
    
    Dim srcR As Single, srcG As Single, srcB As Single
    Dim x As Long, y As Long, quickX As Long
    
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
            
            quickX = x * qvDepth
            
            srcR = srcImageData(quickX)
            srcG = srcImageData(quickX + 1)
            srcB = srcImageData(quickX + 2)
            
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
        IsNormalizeRequired = True
    Else
        IsNormalizeRequired = False
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
        fi_DIB = Plugin_FreeImage.GetFIHandleFromPDDib_NoCopy(tmpDIB)
        
        'Use that handle to request an image resize
        If fi_DIB <> 0 Then
            
            Dim returnDIB As Long
            returnDIB = FreeImage_RescaleByPixel(fi_DIB, dstWidth, dstHeight, True, interpolationType)
                        
            'Copy the bits from the FreeImage DIB to our DIB
            tmpDIB.createBlank dstWidth, dstHeight, 32, 0
            Plugin_FreeImage.PaintFIDibToPDDib tmpDIB, returnDIB, 0, 0, dstWidth, dstHeight
            
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
        fi_DIB = Plugin_FreeImage.GetFIHandleFromPDDib_NoCopy(srcDIB)
        
        'Use that handle to request an image resize
        If fi_DIB <> 0 Then
            
            Dim returnDIB As Long
            returnDIB = FreeImage_RescaleByPixel(fi_DIB, dstWidth, dstHeight, True, interpolationType)
            
            'If the destinationIsBlank flag is TRUE, we can copy the bits directly from the FreeImage bytes to the
            ' destination bytes, skipping the need for an intermediary DIB.
            If destinationIsBlank Then
                Plugin_FreeImage.PaintFIDibToPDDib dstDIB, returnDIB, dstX, dstY, dstWidth, dstHeight
            Else
                Dim tmpDIB As pdDIB
                Set tmpDIB = New pdDIB
                tmpDIB.createBlank dstWidth, dstHeight, 32, 0
                Plugin_FreeImage.PaintFIDibToPDDib tmpDIB, returnDIB, 0, 0, dstWidth, dstHeight
                AlphaBlend dstDIB.getDIBDC, dstX, dstY, dstWidth, dstHeight, tmpDIB.getDIBDC, 0, 0, dstWidth, dstHeight, 255 * &H10000 Or &H1000000
                Set tmpDIB = Nothing
            End If
            
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            
        End If
                
    Else
        FreeImageResizeDIBFast = False
    End If
    
    'If alpha is present, copy the alpha parameters between DIBs, as it will not have changed
    dstDIB.setInitialAlphaPremultiplicationState srcDIB.getAlphaPremultiplication
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format(CStr((Timer - profileTime) * 1000), "0000.00")
    
End Function

'Use FreeImage to rotate a DIB, optimized against the use case where the full source image is being used.
Public Function FreeImageRotateDIBFast(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByRef rotationAngle As Double, Optional ByVal enlargeCanvasToFit As Boolean = True, Optional ByVal applyPostAlphaPremultiplication As Boolean = True) As Boolean

    'Uncomment the two lines below, and the reporting line at the end of the sub, to send timing reports to the debug window.
    'Dim profileTime As Double
    'profileTime = Timer
    
    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
    
        'FreeImage uses positive values to indicate counter-clockwise rotation.  While mathematically correct, I find this
        ' unintuitive for casual users.  PD reverses the rotationAngle value so that POSITIVE values indicate CLOCKWISE rotation.
        rotationAngle = -rotationAngle
        
        'Rotation requires quite a few variables, including a number of handles for passing data back-and-forth with FreeImage.
        Dim fi_DIB As Long, returnDIB As Long
        Dim nWidth As Long, nHeight As Long
        
        'One of the FreeImage rotation variants requires an explicit center point; calculate one in advance.
        Dim cx As Double, cy As Double
        
        cx = srcDIB.getDIBWidth / 2
        cy = srcDIB.getDIBHeight / 2
            
        'Give FreeImage a handle to our temporary rotation image
        fi_DIB = Plugin_FreeImage.GetFIHandleFromPDDib_NoCopy(srcDIB)
        
        If fi_DIB <> 0 Then
            
            'There are two ways to rotate an image - enlarging the canvas to receive the fully rotated copy, or
            ' leaving the image the same size and truncating corners.  These require two different FreeImage functions.
            If enlargeCanvasToFit Then
                
                returnDIB = FreeImage_Rotate(fi_DIB, rotationAngle, 0)
                nWidth = FreeImage_GetWidth(returnDIB)
                nHeight = FreeImage_GetHeight(returnDIB)
                
            'Leave the canvas the same size
            Else
               
               returnDIB = FreeImage_RotateEx(fi_DIB, rotationAngle, 0, 0, cx, cy, True)
               nWidth = FreeImage_GetWidth(returnDIB)
               nHeight = FreeImage_GetHeight(returnDIB)
            
            End If
            
            'Unload the original FreeImage source
            FreeImage_UnloadEx fi_DIB
            
            If returnDIB <> 0 Then
            
                'Ask FreeImage to premultiply the image's alpha data, as necessary
                If applyPostAlphaPremultiplication Then FreeImage_PreMultiplyWithAlpha returnDIB
                
                'Create a blank DIB to receive the rotated image from FreeImage
                dstDIB.createBlank nWidth, nHeight, 32
                            
                'Copy the bits from the FreeImage DIB to our DIB
                Plugin_FreeImage.PaintFIDibToPDDib dstDIB, returnDIB, 0, 0, nWidth, nHeight
                
                'With the transfer complete, release any remaining FreeImage DIBs and exit
                FreeImage_UnloadEx returnDIB
                FreeImageRotateDIBFast = True
                
            Else
                FreeImageRotateDIBFast = False
            End If
            
        Else
            FreeImageRotateDIBFast = False
        End If
                
    Else
        FreeImageRotateDIBFast = False
    End If
    
    'If alpha is present, copy the alpha parameters between DIBs, as it will not have changed
    dstDIB.setInitialAlphaPremultiplicationState srcDIB.getAlphaPremultiplication
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format(CStr((Timer - profileTime) * 1000), "0000.00")
    
End Function

Public Function FreeImageErrorState() As Boolean
    FreeImageErrorState = CBool(Len(g_FreeImageErrorMessages(UBound(g_FreeImageErrorMessages))) <> 0)
End Function

Public Function GetFreeImageErrors(Optional ByVal eraseListUponReturn As Boolean = True) As String
    
    Dim listOfFreeImageErrors As String
    listOfFreeImageErrors = """"
    
    'Condense all recorded errors into a single string
    If UBound(g_FreeImageErrorMessages) > 0 Then
        Dim i As Long
        For i = 0 To UBound(g_FreeImageErrorMessages)
            listOfFreeImageErrors = listOfFreeImageErrors & g_FreeImageErrorMessages(i)
            If i < UBound(g_FreeImageErrorMessages) Then listOfFreeImageErrors = listOfFreeImageErrors & vbCrLf
        Next i
    Else
        listOfFreeImageErrors = listOfFreeImageErrors & g_FreeImageErrorMessages(0)
    End If
    
    listOfFreeImageErrors = listOfFreeImageErrors & """"
    GetFreeImageErrors = listOfFreeImageErrors
    
    If eraseListUponReturn Then ReDim g_FreeImageErrorMessage(0) As String
    
End Function

'Need a FreeImage object at a specific color depth?  Use this function.  Note that the source DIB is not touched or freed,
' and obviously you must manually free the returned FreeImage handle when you're done with it.
'
'Some combinations of values are not valid (e.g. alphaState and outputColorDepth must be mixed carefully); refer to the
' ImageExporter module for specific details on supported color modes.
'
'Finally, this function does not run heuristics on the incoming image.  For example, if you tell it to create a
' grayscale image, it *will* create a grayscale image, regardless of the input.  As such, you must run any
' intelligent heuristics *prior* to calling this function!
'
'Returns: a non-zero FI handle if successful; 0 if something goes horribly wrong
Public Function GetFIDib_SpecificColorMode(ByRef srcDIB As pdDIB, ByVal outputColorDepth As Long, Optional ByVal desiredAlphaState As PD_ALPHA_STATUS = PDAS_ComplicatedAlpha, Optional ByVal currentAlphaState As PD_ALPHA_STATUS = PDAS_ComplicatedAlpha, Optional ByVal alphaCutoff As Long = 127, Optional ByVal BackgroundColor As Long = vbWhite, Optional ByVal forceGrayscale As Boolean = False, Optional ByVal paletteCount As Long = 256, Optional ByVal RGB16bppUse565 As Boolean = True) As Long

    'The order of operations here is a bit tricky.  First, we need to deal with the problem of binary alpha values.
    ' Binary alpha values require us to leave the image in 32-bpp mode, but force it to use only "0" or "255" alpha
    ' values.  As part of this process, transparent pixels will have their color forcibly changed to magic magenta.
    ' This allows us to easily detect them post-quantization.  (Also, if the alpha cutoff is set to 0, we mark the
    ' image as *not* using alpha at all.)
    If (desiredAlphaState = PDAS_BinaryAlpha) Then
        
        If (currentAlphaState = PDAS_ComplicatedAlpha) Then
            srcDIB.ApplyAlphaCutoff alphaCutoff, , BackgroundColor
            If alphaCutoff = 0 Then desiredAlphaState = PDAS_NoAlpha
        
        'If the image already has binary alpha, don't waste time re-applying it.  Instead, make sure we're tracking
        ' the image's current transparent color, as well as the location of a known transparent pixel.  We'll need
        ' these if the caller requested additional quantization options or a wacky bit-depth.
        
        'TODO: see if we still need to use magic magenta here.  There's a chance that it's still relevant, as the
        ' source image may be using black as both a color, and a transparent marker.  The best solution may be
        ' to expose an optional parameter in MemorizeBinaryAlphaData, that still requests a magic magenta conversion.
        ElseIf (currentAlphaState = PDAS_BinaryAlpha) Then
            srcDIB.MemorizeBinaryAlphaData
        End If
        
    End If
    
    'Next, if the caller doesn't want us to use alpha at all, reduce to 24-bpp internally.
    If (desiredAlphaState = PDAS_NoAlpha) Or (outputColorDepth = 24) Or (outputColorDepth = 48) Or (outputColorDepth = 96) Then
        If (srcDIB.getDIBColorDepth = 32) Then srcDIB.convertTo24bpp BackgroundColor
    End If
    
    'Create a default FreeImage handle now
    Dim fi_DIB As Long, tmpFIHandle As Long
    fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
    
    '1-bpp is easy; handle it now
    If (outputColorDepth = 1) Then
        tmpFIHandle = FreeImage_Dither(fi_DIB, FID_FS)
        If (tmpFIHandle <> fi_DIB) Then
            FreeImage_Unload fi_DIB
            fi_DIB = tmpFIHandle
        End If
    
    'Non-1-bpp is harder
    Else
        
        'Handle grayscale variants first; they use their own dedicated conversion functions
        If forceGrayscale Then
            fi_DIB = GetGrayscaleFIDib(fi_DIB, outputColorDepth)
        
        'Non-grayscale variants are more complicated
        Else
        
            'Start with non-alpha color modes.  They are easier to handle.
            If (desiredAlphaState = PDAS_NoAlpha) Then
            
                'Walk down the list of valid outputs, starting at the low end
                If (outputColorDepth <= 8) Then
                    
                    'FreeImage supports a new "lossless" quantization method that is perfect for images that already
                    ' have 256 colors or less.  This method is basically just a hash table, and it lets us avoid
                    ' lossy quantization if at all possible.
                    If (paletteCount = 256) Then
                        tmpFIHandle = FreeImage_ColorQuantize(fi_DIB, FIQ_LFPQUANT)
                    Else
                        tmpFIHandle = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_LFPQUANT, False, paletteCount)
                    End If
                    
                    '0 means the image has > 256 colors, and must be quantized via lossy means
                    If (tmpFIHandle = 0) Then
                        
                        'If we're going straight to 4-bits, ignore the user's palette count in favor of a 16-bit one.
                        If (outputColorDepth = 4) Then
                            tmpFIHandle = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_WUQUANT, False, 16)
                        Else
                            If (paletteCount = 256) Then
                                tmpFIHandle = FreeImage_ColorQuantize(fi_DIB, FIQ_WUQUANT)
                            Else
                                tmpFIHandle = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_WUQUANT, False, paletteCount)
                            End If
                        End If
                        
                    End If
                    
                    If (tmpFIHandle <> fi_DIB) Then
                        FreeImage_Unload fi_DIB
                        fi_DIB = tmpFIHandle
                    End If
                    
                    'We now have an 8-bpp image.  Forcibly convert to 4-bpp if necessary.
                    If (outputColorDepth = 4) Then
                        tmpFIHandle = FreeImage_ConvertTo4Bits(fi_DIB)
                        If (tmpFIHandle <> fi_DIB) Then
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
                        If (outputColorDepth > 24) Then
                        
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
            
            'The image contains alpha, and the caller wants it preserved.  AAARRRGGGHHH
            Else
            
                'Skip 32-bpp, as it's already standardized
                If (outputColorDepth <> 32) Then
                
                    '< 8-bpp is the ugliest, so it comes first!
                    If (outputColorDepth < 32) Then
                    
                        'Technically, PNG supports the concept of a single "transparent color" for any bit-depth,
                        ' including high bit-depth images.  FreeImage, however, doesn't have that level of granularity.
                        ' As such, any transparency under 8-bpp is handled via a transparent palette index.
                        
                        'Start by getting the image into 8-bpp color mode.
                        
                        'FreeImage supports a new "lossless" quantization method that is perfect for images that already
                        ' have 256 colors or less.  This method is basically just a hash table, and it lets us avoid
                        ' lossy quantization if at all possible.
                        If (paletteCount = 256) Then
                            tmpFIHandle = FreeImage_ColorQuantize(fi_DIB, FIQ_LFPQUANT)
                        Else
                            tmpFIHandle = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_LFPQUANT, False, paletteCount)
                        End If
                    
                        '0 means the image has > 256 colors, and must be quantized via lossy means
                        If (tmpFIHandle = 0) Then
                            
                            'If we're going straight to 4-bits, ignore the user's palette count in favor of a 16-bit one.
                            If (outputColorDepth = 4) Then
                                tmpFIHandle = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_WUQUANT, False, 16)
                            Else
                                If (paletteCount = 256) Then
                                    tmpFIHandle = FreeImage_ColorQuantize(fi_DIB, FIQ_WUQUANT)
                                Else
                                    tmpFIHandle = FreeImage_ColorQuantizeEx(fi_DIB, FIQ_WUQUANT, False, paletteCount)
                                End If
                            End If
                            
                        End If
                    
                        If (tmpFIHandle <> fi_DIB) Then
                            FreeImage_Unload fi_DIB
                            fi_DIB = tmpFIHandle
                        End If
                    
                        'fi_DIB now contains an 8-bpp image.  We next need to find the palette index of a known transparent pixel
                        Dim transpX As Long, transpY As Long
                        srcDIB.getTransparentLocation transpX, transpY
                        
                        'Use that location to retrieve the matching transparent index
                        Dim palIndex As Byte
                        FreeImage_GetPixelIndex fi_DIB, transpX, transpY, palIndex
                        FreeImage_SetTransparentIndex fi_DIB, palIndex
            
                        'Finally, because some software may not display the transparency correctly, we need to set that
                        ' palette index color to its original value.  To do that, we must make a copy of the palette and
                        ' update the transparency index accordingly.
                        Dim fi_Palette() As Long
                        fi_Palette = FreeImage_GetPaletteExLong(fi_DIB)
                        fi_Palette(palIndex) = srcDIB.GetOriginalTransparentColor()
                        
                        'We now have an < 8-bpp image with a transparent index correctly marked.  Whew!
                    
                    'Output is > 32-bpp with transparency
                    Else
                        
                        'High bit-depth variants are covered last
                        If (outputColorDepth = 64) Then
                            tmpFIHandle = FreeImage_ConvertToRGBA16(fi_DIB)
                            If (tmpFIHandle <> fi_DIB) Then
                                FreeImage_Unload fi_DIB
                                fi_DIB = tmpFIHandle
                            End If
                        '128-bpp is the only other possibility
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
        
        'Output colordepth must be 16; any other values are invalid
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
Public Function GetExportPreview(ByRef srcFI_Handle As Long, ByRef dstDIB As pdDIB, ByVal dstFormat As PHOTODEMON_IMAGE_FORMAT, Optional ByVal fi_SaveFlags As Long = 0, Optional ByVal fi_LoadFlags As Long = 0)
    
    Dim fi_Check As Long, fi_Size As Long
    fi_Check = FreeImage_SaveToMemoryEx(dstFormat, srcFI_Handle, m_ExportPreviewBytes, fi_SaveFlags, False, fi_Size)
    If fi_Check <> 0 Then
        
        Dim fi_DIB As Long
        fi_DIB = FreeImage_LoadFromMemoryEx(Nothing, fi_LoadFlags, fi_Size, dstFormat, VarPtr(m_ExportPreviewBytes(0)))
        
        If (fi_DIB <> 0) Then
            FreeImage_FlipVertically fi_DIB
            If (FreeImage_GetBPP(fi_DIB) <> 24) And (FreeImage_GetBPP(fi_DIB) <> 32) Then
                If FreeImage_IsTransparent(fi_DIB) Then
                    fi_DIB = FreeImage_ConvertColorDepth(fi_DIB, FICF_RGB_32BPP, True)
                Else
                    fi_DIB = FreeImage_ConvertColorDepth(fi_DIB, FICF_RGB_24BPP, True)
                End If
            End If
            Plugin_FreeImage.PaintFIDibToPDDib dstDIB, fi_DIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight
            FreeImage_Unload fi_DIB
            GetExportPreview = True
        Else
            GetExportPreview = False
        End If
        
    Else
        GetExportPreview = False
    End If
    
End Function

'PD uses a persistent cache for generating post-export preview images.  This costs several MB of memory but greatly improves
' responsiveness of export dialogs.  When such a dialog is unloaded, you can call this function to forcibly reclaim the memory
' associated with that cache.
Public Sub ReleasePreviewCache()
    Erase m_ExportPreviewBytes
End Sub
