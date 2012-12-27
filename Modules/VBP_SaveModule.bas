Attribute VB_Name = "Saving"
'***************************************************************************
'File Saving Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 15/December/12
'Last update: Rewrote all save subs as functions.  They now report success or not to the calling routine.
'
'Module for handling all image saving.  It contains pretty much every routine that I find useful;
' the majority of the functions are simply interfaces to FreeImage, so if that is not enabled than
' only a subset of these matter.
'
'***************************************************************************

Option Explicit

'Save the current image to BMP format
Public Function SaveBMP(ByVal imageID As Long, ByVal BMPPath As String, ByVal outputColorDepth As Long) As Boolean
   
    On Error GoTo SaveBMPError
   
    'If the output color depth is 24 or 32bpp, or if both GDI+ and FreeImage are missing, use our own internal methods
    ' to save a BMP file
    If (outputColorDepth = 24) Or (outputColorDepth = 32) Or ((Not imageFormats.GDIPlusEnabled) And (Not imageFormats.FreeImageEnabled)) Then
    
        Message "Saving bitmap..."
    
        'The layer class is capable of doing this without any outside help.
        pdImages(imageID).mainLayer.writeToBitmapFile BMPPath
    
        Message "BMP save complete."
        
    'If some other color depth is specified, use FreeImage or GDI+ to write the file
    Else
    
        If imageFormats.FreeImageEnabled Then
            
            'Load FreeImage into memory
            Dim hLib As Long
            hLib = LoadLibrary(PluginPath & "FreeImage.dll")
            
            Message "Preparing BMP image..."
            
            'Copy the image into a temporary layer
            Dim tmpLayer As pdLayer
            Set tmpLayer = New pdLayer
            tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
            
            'If the output color depth is 24 but the current image is 32, composite the image against a white background
            If (outputColorDepth < 32) And (pdImages(imageID).mainLayer.getLayerColorDepth = 32) Then tmpLayer.convertTo24bpp
            
            'Convert our current layer to a FreeImage-type DIB
            Dim fi_DIB As Long
            fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
                        
            'If the image is being reduced from some higher bit-depth to 1bpp, manually force a conversion with dithering
            If outputColorDepth = 1 Then fi_DIB = FreeImage_Dither(fi_DIB, FID_FS)
            
            'Finally, prepare some BMP save flags.  If the user has requested RLE encoding, and this image is <= 8bpp,
            ' request RLE encoding from FreeImage.
            Dim BMPflags As Long
            BMPflags = BMP_DEFAULT
            
            If outputColorDepth = 8 And userPreferences.GetPreference_Boolean("General Preferences", "BitmapRLE", False) Then BMPflags = BMP_SAVE_RLE
            
            'Use that handle to save the image to BMP format, with required color conversion based on the outgoing color depth
            If fi_DIB <> 0 Then
                Dim fi_Check As Long
                fi_Check = FreeImage_SaveEx(fi_DIB, BMPPath, FIF_BMP, BMPflags, outputColorDepth, , , , , True)
                If fi_Check = False Then
                    Message "BMP save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
                    FreeLibrary hLib
                    SaveBMP = False
                    Exit Function
                Else
                    Message "BMP save complete."
                End If
            Else
                Message "BMP save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
                SaveBMP = False
                FreeLibrary hLib
                Exit Function
            End If
    
            'Release FreeImage from memory
            FreeLibrary hLib
            
        Else
            GDIPlusSavePicture imageID, BMPPath, ImageBMP, outputColorDepth
        End If
    
    End If
    
    SaveBMP = True
    Exit Function
    
SaveBMPError:

    If hLib <> 0 Then FreeLibrary hLib

    SaveBMP = False
    
End Function

'Save the current image to PhotoDemon's native PDI format
Public Function SavePhotoDemonImage(ByVal imageID As Long, ByVal PDIPath As String) As Boolean
    
    On Error GoTo SavePDIError
    
    Message "Saving PhotoDemon Image..."

    'First, have the layer write itself to file in BMP format
    pdImages(imageID).mainLayer.writeToBitmapFile PDIPath
    
    'Then compress the file using zLib
    CompressFile PDIPath
    
    Message "PDI Save complete."
    
    SavePhotoDemonImage = True
    Exit Function
    
SavePDIError:

    SavePhotoDemonImage = False
    
End Function

'Save a GIF (Graphics Interchange Format) image.  GDI+ can also do this.
Public Function SaveGIFImage(ByVal imageID As Long, ByVal GIFPath As String) As Boolean

    On Error GoTo SaveGIFError

    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        SaveGIFImage = False
        Exit Function
    End If
    
    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing GIF image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    
    'If the current image is 32bpp, we will need to apply some additional actions to the image to prepare the
    ' transparency.  Mark a bool value, because we will reference it in multiple places throughout the save function.
    Dim handleAlpha As Boolean
    If pdImages(imageID).mainLayer.getLayerColorDepth = 32 Then handleAlpha = True Else handleAlpha = False
    
    'If the current image contains transparency, we need to modify it in order to retain the alpha channel.
    If handleAlpha Then
    
        'Does this layer contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If tmpLayer.isAlphaBinary Then
            tmpLayer.applyAlphaCutoff
        Else
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
            Message "GIF save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
            FreeLibrary hLib
            SaveGIFImage = False
            Exit Function
        Else
            Message "GIF save complete."
        End If
    Else
        Message "GIF save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
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
Public Function SavePNGImage(ByVal imageID As Long, ByVal PNGPath As String, ByVal outputColorDepth As Long) As Boolean

    On Error GoTo SavePNGError

    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        SavePNGImage = False
        Exit Function
    End If
    
    'Before doing anything else, make a special note of the outputColorDepth.  If it is 8bpp, we will use pngnq-s9 to help with the save.
    Dim output8BPP As Boolean
    If outputColorDepth = 8 Then output8BPP = True Else output8BPP = False
    
    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing PNG image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (pdImages(imageID).mainLayer.getLayerColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'PhotoDemon now offers pngnq support via a plugin.  It can be used to render extremely high-quality 8bpp PNG files
        ' with "full" transparency.  If it is available, the export process is a bit different.
        
        'Before we can send stuff off to pngnq, however, we need to see if the image has more than 256 colors.  If it
        ' doesn't, we can save the file ourselves.
        
        'Check to see if the current image had its colors counted before coming here.  If not, count it.
        Dim numColors As Long
        If g_LastImageScanned <> imageID Then
            numColors = getQuickColorCount(pdImages(imageID), imageID)
        Else
            numColors = g_LastColorCount
        End If
        
        'Pngnq can handle all transparency for us, if it exists.  If it does not, we must rely on our own routines.
        If Not imageFormats.pngnqEnabled Then
        
            'Does this layer contain binary transparency?  If so, mark all transparent pixels with magic magenta.
            If tmpLayer.isAlphaBinary Then
                tmpLayer.applyAlphaCutoff
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
            
        'If pngnq is available, force the output to 32bpp.  PNGNQ will take care of the actual 8bpp reduction.
        Else
            outputColorDepth = 32
        End If
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (pdImages(imageID).mainLayer.getLayerColorDepth = 32) And (outputColorDepth < 32) Then tmpLayer.compositeBackgroundColor 255, 255, 255
    
        'Also, if pngnq is enabled, we will use that for the transformation - so we need to reset the outgoing color depth to 24bpp
        If (pdImages(imageID).mainLayer.getLayerColorDepth = 24) And (outputColorDepth = 8) And imageFormats.pngnqEnabled Then outputColorDepth = 24
    
    End If
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'If the image is being reduced from some higher bit-depth to 1bpp, manually force a conversion with dithering
    If outputColorDepth = 1 Then fi_DIB = FreeImage_Dither(fi_DIB, FID_FS)
    
    'If the image contains alpha and pngnq is not available, we need to manually convert the FreeImage copy of the image to 8bpp.
    ' Then we need to apply alpha using the cut-off established earlier in this section.
    If handleAlpha And (Not imageFormats.pngnqEnabled) Then
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
        If userPreferences.GetPreference_Boolean("General Preferences", "PNGBackgroundPreservation", True) And pdImages(imageID).pngBackgroundColor <> -1 Then
            
            Dim rQuad As RGBQUAD
            rQuad.rgbRed = ExtractR(pdImages(imageID).pngBackgroundColor)
            rQuad.rgbGreen = ExtractG(pdImages(imageID).pngBackgroundColor)
            rQuad.rgbBlue = ExtractB(pdImages(imageID).pngBackgroundColor)
            FreeImage_SetBackgroundColor fi_DIB, rQuad
        
        End If
    
        'Finally, prepare some PNG save flags based on the user's saved preferences
        Dim PNGFlags As Long
        
        'Compression level (1 to 9, but FreeImage also has a "no compression" option with a unique flag)
        PNGFlags = userPreferences.GetPreference_Long("General Preferences", "PNGCompression", 9)
        If PNGFlags = 0 Then PNGFlags = PNG_Z_NO_COMPRESSION
        
        'Interlacing
        If userPreferences.GetPreference_Boolean("General Preferences", "PNGInterlacing", False) Then PNGFlags = (PNGFlags Or PNG_INTERLACED)
    
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, PNGPath, FIF_PNG, PNGFlags, outputColorDepth, , , , , True)
        
        If fi_Check = False Then
            Message "PNG save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
            FreeLibrary hLib
            SavePNGImage = False
            Exit Function
        
        Else
            
            'If pngnq is being used to help with the 8bpp reduction, now is when we need to use it.
            If imageFormats.pngnqEnabled And output8BPP Then
            
                'Build a full shell path for the pngnq operation
                Dim shellPath As String
                shellPath = PluginPath & "pngnq-s9.exe "
                
                'Force overwrite if a file with that name already exists
                shellPath = shellPath & "-f "
                
                'Display verbose status messages (consider removing this for production build)
                shellPath = shellPath & "-v "
                
                'Turn off the alpha importance heuristic (this leads to better results on semi-transparent images, and improves
                ' processing time for 24bpp images)
                shellPath = shellPath & "-A "
                
                'Now, add options that the user may have specified.
                
                'Alpha extenuation (only relevant for 32bpp images)
                If pdImages(imageID).mainLayer.getLayerColorDepth = 32 Then
                    If userPreferences.GetPreference_Boolean("Plugin Preferences", "PngnqAlphaExtenuation", False) Then
                        shellPath = shellPath & "-t15 "
                    Else
                        shellPath = shellPath & "-t0 "
                    End If
                End If
        
                'YUV
                If userPreferences.GetPreference_Boolean("Plugin Preferences", "PngnqYUV", True) Then
                    shellPath = shellPath & "-Cy "
                Else
                    shellPath = shellPath & "-Cr "
                End If
        
                'Color sample size
                shellPath = shellPath & "-s" & userPreferences.GetPreference_Long("Plugin Preferences", "PngnqColorSample", 3) & " "
        
                'Dithering
                If userPreferences.GetPreference_Long("Plugin Preferences", "PngnqDithering", 5) = 0 Then
                    shellPath = shellPath & "-Qn "
                Else
                    shellPath = shellPath & "-Q" & userPreferences.GetPreference_Long("Plugin Preferences", "PngnqDithering", 5) & " "
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
                'Shell shellPath, vbMinimizedNoFocus
            
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
            
            Message "PNG save complete."
        End If
    Else
        Message "PNG save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
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
Public Function SavePPMImage(ByVal imageID As Long, ByVal PPMPath As String) As Boolean

    On Error GoTo SavePPMError

    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        SavePPMImage = False
        Exit Function
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing PPM image..."
    
    'Based on the user's preference, select binary or ASCII encoding for the PPM file
    Dim ppm_Encoding As FREE_IMAGE_SAVE_OPTIONS
    If userPreferences.GetPreference_Long("General Preferences", "PPMExportFormat", 0) = 0 Then ppm_Encoding = FISO_PNM_SAVE_RAW Else ppm_Encoding = FISO_PNM_SAVE_ASCII
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
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
            Message "PPM save complete."
        End If
    Else
        Message "PPM save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
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
Public Function SaveTGAImage(ByVal imageID As Long, ByVal TGAPath As String, ByVal outputColorDepth As Long) As Boolean
    
    On Error GoTo SaveTGAError
    
    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        SaveTGAImage = False
        Exit Function
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing TGA image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (pdImages(imageID).mainLayer.getLayerColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'Does this layer contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If tmpLayer.isAlphaBinary Then
            tmpLayer.applyAlphaCutoff
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
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (pdImages(imageID).mainLayer.getLayerColorDepth = 32) And (outputColorDepth < 32) Then tmpLayer.compositeBackgroundColor 255, 255, 255
    
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
            
    If userPreferences.GetPreference_Boolean("General Preferences", "TGARLE", False) Then TGAflags = TARGA_SAVE_RLE
            
    
    'Use that handle to save the image to TGA format
    If fi_DIB <> 0 Then
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TGAPath, FIF_TARGA, TGAflags, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "TGA save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
            FreeLibrary hLib
            SaveTGAImage = False
            Exit Function
        Else
            Message "TGA save complete."
        End If
    Else
        Message "TGA save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
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

'Save to JPEG using the FreeImage library.  This is more reliable than using GDI+.
Public Function SaveJPEGImage(ByVal imageID As Long, ByVal JPEGPath As String, ByVal jQuality As Long, Optional ByVal jOtherFlags As Long = 0, Optional ByVal jCreateThumbnail As Long = 0) As Boolean
    
    On Error GoTo SaveJPEGError
    
    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        SaveJPEGImage = False
        Exit Function
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing JPEG image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
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
    If jOtherFlags <> 0 Then jQuality = jQuality Or jOtherFlags
    
    'If a thumbnail has been requested, generate that now
    If jCreateThumbnail <> 0 Then
    
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
        fi_Check = FreeImage_SaveEx(fi_DIB, JPEGPath, FIF_JPEG, jQuality, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "JPEG save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
            FreeLibrary hLib
            SaveJPEGImage = False
            Exit Function
        Else
            Message "JPEG save complete."
        End If
    Else
        Message "JPEG save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
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
Public Function SaveTIFImage(ByVal imageID As Long, ByVal TIFPath As String, ByVal outputColorDepth As Long) As Boolean
    
    On Error GoTo SaveTIFError
    
    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        SaveTIFImage = False
        Exit Function
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing TIFF image..."
    
    'TIFFs have some unique considerations regarding compression techniques.  If a color-depth-specific compression
    ' technique has been requested, modify the output depth accordingly.
    Select Case userPreferences.GetPreference_Long("General Preferences", "TIFFCompression", 0)
        
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
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    
    'If the image is being saved to a lower bit-depth, we may have to adjust the alpha channel.  Check for that now.
    Dim handleAlpha As Boolean
    If (pdImages(imageID).mainLayer.getLayerColorDepth = 32) And (outputColorDepth = 8) Then handleAlpha = True Else handleAlpha = False
    
    'If this image is 32bpp but the output color depth is less than that, make necessary preparations
    If handleAlpha Then
        
        'Does this layer contain binary transparency?  If so, mark all transparent pixels with magic magenta.
        If tmpLayer.isAlphaBinary Then
            tmpLayer.applyAlphaCutoff
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
    
    Else
    
        'If we are not saving to 8bpp, check to see if we are saving to some other smaller bit-depth.
        ' If we are, composite the image against a white background.
        If (pdImages(imageID).mainLayer.getLayerColorDepth = 32) And (outputColorDepth < 32) Then tmpLayer.compositeBackgroundColor 255, 255, 255
    
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
        
        Select Case userPreferences.GetPreference_Long("General Preferences", "TIFFCompression", 0)
        
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
        
        If userPreferences.GetPreference_Boolean("General Preferences", "TIFFCMYK", False) Then TIFFFlags = (TIFFFlags Or TIFF_CMYK)
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TIFPath, FIF_TIFF, TIFFFlags, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "TIFF save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
            FreeLibrary hLib
            SaveTIFImage = False
            Exit Function
        Else
            Message "TIFF save complete."
        End If
    Else
        Message "TIFF save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
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

'Save to JPEG-2000 format using the FreeImage library.  This is currently deemed "experimental".
Public Function SaveJP2Image(ByVal imageID As Long, ByVal jp2Path As String, ByVal outputColorDepth As Long, Optional ByVal jp2Quality As Long = 16) As Boolean
    
    On Error GoTo SaveJP2Error
    
    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        SaveJP2Image = False
        Exit Function
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing JPEG-2000 image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (pdImages(imageID).mainLayer.getLayerColorDepth = 32) Then tmpLayer.convertTo24bpp
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'Use that handle to save the image to JPEG format
    If fi_DIB <> 0 Then
                
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, jp2Path, FIF_JP2, jp2Quality, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "JPEG-2000 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
            FreeLibrary hLib
            SaveJP2Image = False
            Exit Function
        Else
            Message "JPEG-2000 save complete."
        End If
    Else
        Message "JPEG-2000 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
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

