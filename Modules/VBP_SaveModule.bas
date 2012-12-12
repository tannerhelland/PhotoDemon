Attribute VB_Name = "Saving"
'***************************************************************************
'File Saving Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 11/December/12
'Last update: Rewrote all save functions to handle variable color depths properly
'
'Module for handling all image saving.  It contains pretty much every routine that I find useful;
' the majority of the functions are simply interfaces to FreeImage, so if that is not enabled than
' only a subset of these matter.
'
'***************************************************************************

Option Explicit

'Save the current image to BMP format
Public Sub SaveBMP(ByVal imageID As Long, ByVal BMPPath As String, ByVal outputColorDepth As Long)
   
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
            
            'Use that handle to save the image to GIF format, with required color conversion based on the outgoing color depth
            If fi_DIB <> 0 Then
                Dim fi_Check As Long
                fi_Check = FreeImage_SaveEx(fi_DIB, BMPPath, FIF_BMP, , outputColorDepth, , , , , True)
                If fi_Check = False Then
                    Message "BMP save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
                Else
                    Message "BMP save complete."
                End If
            Else
                Message "BMP save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
            End If
    
            'Release FreeImage from memory
            FreeLibrary hLib
            
        Else
            GDIPlusSavePicture imageID, BMPPath, ImageBMP, outputColorDepth
        End If
    
    End If
    
End Sub

'Save the current image to PhotoDemon's native PDI format
Public Sub SavePhotoDemonImage(ByVal imageID As Long, ByVal PDIPath As String)
    
    Message "Saving PhotoDemon Image..."

    'First, have the layer write itself to file in BMP format
    pdImages(imageID).mainLayer.writeToBitmapFile PDIPath
    
    'Then compress the file using zLib
    CompressFile PDIPath
    
    Message "PDI Save complete."
    
End Sub

'Save a GIF (Graphics Interchange Format) image.  GDI+ can also do this.
Public Sub SaveGIFImage(ByVal imageID As Long, ByVal GIFPath As String)

    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
    End If
    
    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing GIF image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    
    'If the current image is 32bpp, composite the image against a white background
    If pdImages(imageID).mainLayer.getLayerColorDepth = 32 Then tmpLayer.convertTo24bpp
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'Use that handle to save the image to GIF format, with required 8bpp (256 color) conversion
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, GIFPath, FIF_GIF, , FICD_8BPP, , , , , True)
        If fi_Check = False Then
            Message "GIF save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "GIF save complete."
        End If
    Else
        Message "GIF save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
End Sub

'Save a PNG (Portable Network Graphic) file.  GDI+ can also do this.
Public Sub SavePNGImage(ByVal imageID As Long, ByVal PNGPath As String, ByVal outputColorDepth As Long)

    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
    End If
    
    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing PNG image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (pdImages(imageID).mainLayer.getLayerColorDepth = 32) Then tmpLayer.convertTo24bpp
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'Use that handle to save the image to PNG format
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
                
        fi_Check = FreeImage_SaveEx(fi_DIB, PNGPath, FIF_PNG, FISO_PNG_Z_BEST_COMPRESSION, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "PNG save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "PNG save complete."
        End If
    Else
        Message "PNG save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib

End Sub

'Save a PPM (Portable Pixmap) image
Public Sub SavePPMImage(ByVal imageID As Long, ByVal PPMPath As String)

    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
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
        Else
            Message "PPM save complete."
        End If
    Else
        Message "PPM save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
        
End Sub

'Save to Targa (TGA) format.
Public Sub SaveTGAImage(ByVal imageID As Long, ByVal TGAPath As String, ByVal outputColorDepth As Long)
    
    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing TGA image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (pdImages(imageID).mainLayer.getLayerColorDepth = 32) Then tmpLayer.convertTo24bpp
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'Use that handle to save the image to TGA format
    If fi_DIB <> 0 Then
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TGAPath, FIF_TARGA, FILO_TARGA_DEFAULT, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "TGA save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "TGA save complete."
        End If
    Else
        Message "TGA save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib

End Sub

'Save to JPEG using the FreeImage library.  This is more reliable than using GDI+.
Public Sub SaveJPEGImage(ByVal imageID As Long, ByVal JPEGPath As String, ByVal jQuality As Long, Optional ByVal jOtherFlags As Long = 0, Optional ByVal jCreateThumbnail As Long = 0)
    
    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
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
        fi_Check = FreeImage_SaveEx(fi_DIB, JPEGPath, FIF_JPEG, jQuality, FICD_24BPP, , , , , True)
        If fi_Check = False Then
            Message "JPEG save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "JPEG save complete."
        End If
    Else
        Message "JPEG save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
End Sub

'Save a TIFF (Tagged Image File Format) image via FreeImage.  GDI+ can also do this.
Public Sub SaveTIFImage(ByVal imageID As Long, ByVal TIFPath As String, ByVal outputColorDepth As Long)
    
    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing TIFF image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    
    'If the output color depth is 24 but the current image is 32, composite the image against a white background
    If (outputColorDepth < 32) And (pdImages(imageID).mainLayer.getLayerColorDepth = 32) Then tmpLayer.convertTo24bpp
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'Use that handle to save the image to TIFF format
    If fi_DIB <> 0 Then
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TIFPath, FIF_TIFF, FISO_TIFF_DEFAULT, outputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "TIFF save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "TIFF save complete."
        End If
    Else
        Message "TIFF save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
        
End Sub

'Save to JPEG-2000 format using the FreeImage library.  This is currently deemed "experimental".
Public Sub SaveJP2Image(ByVal imageID As Long, ByVal jp2Path As String, ByVal outputColorDepth As Long, Optional ByVal jp2Quality As Long = 16)
    
    'Make sure we found the plug-in when we loaded the program
    If imageFormats.FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
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
        Else
            Message "JPEG-2000 save complete."
        End If
    Else
        Message "JPEG-2000 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
End Sub

