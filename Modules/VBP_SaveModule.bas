Attribute VB_Name = "Saving"
'***************************************************************************
'File Saving Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 04/September/12
'Last update: Rewrote all save code against the new layer class.  No more intermediary picture boxes!
'
'Module for handling all image saving.  It contains pretty much every routine that I find useful;
' the majority of the functions are simply interfaces to FreeImage, so if that is not enabled than
' only a subset of these matter.
'
'***************************************************************************

Option Explicit

'Save the current image to BMP format
Public Sub SaveBMP(ByVal imageID As Long, ByVal BMPPath As String)
    
    Message "Saving image..."
    
    'The layer class is capable of doing this without any outside help.
    pdImages(imageID).mainLayer.writeToBitmapFile BMPPath
    
    Message "Save complete."
    
End Sub

'Save the current image to PhotoDemon's native PDI format
Public Sub SavePhotoDemonImage(ByVal imageID As Long, ByVal PDIPath As String)
    
    Message "Saving image..."

    'First, have the layer write itself to file in BMP format
    pdImages(imageID).mainLayer.writeToBitmapFile PDIPath
    
    'Then compress the file using zLib
    CompressFile PDIPath
    
    Message "Save complete."
    
End Sub

'Save a GIF (Graphics Interchange Format) image.  GDI+ can also do this.
Public Sub SaveGIFImage(ByVal imageID As Long, ByVal GIFPath As String)

    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
    End If
    
    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(pdImages(imageID).mainLayer.getLayerDC)
    
    'Use that handle to save the image to GIF format, with required 8bpp (256 color) conversion
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, GIFPath, FIF_GIF, , FICD_8BPP, , , , , True)
        If fi_Check = False Then
            Message "Save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "Save complete."
        End If
    Else
        Message "Save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
End Sub

'Save a PNG (Portable Network Graphic) file.  GDI+ can also do this.
Public Sub SavePNGImage(ByVal imageID As Long, ByVal PNGPath As String, Optional ByVal PNGColorDepth As Long = &H18)

    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
    End If
    
    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(pdImages(imageID).mainLayer.getLayerDC)
    
    'Use that handle to save the image to PNG format
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        
        'In the future, the color depth of the output file should be user-controllable via a form.  Right now, however, just use
        ' the color depth of the current image
        Dim fi_OutputColorDepth As FREE_IMAGE_COLOR_DEPTH
        If pdImages(imageID).mainLayer.getLayerColorDepth = 24 Then
            fi_OutputColorDepth = FICD_24BPP
        Else
            fi_OutputColorDepth = FICD_32BPP
        End If
        
        fi_Check = FreeImage_SaveEx(fi_DIB, PNGPath, FIF_PNG, FISO_PNG_Z_BEST_COMPRESSION, fi_OutputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "Save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "Save complete."
        End If
    Else
        Message "Save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib

End Sub

'IMPORTANT NOTE: Only ASCII format PPM is currently enabled.  RAW IS NOT YET SUPPORTED!
Public Sub SavePPMImage(ByVal imageID As Long, ByVal PPMPath As String)

    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    If tmpLayer.getLayerColorDepth = 32 Then tmpLayer.compositeBackgroundColor 255, 255, 255
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'Use that handle to save the image to PPM format (ASCII)
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, PPMPath, FIF_PPM, , FICD_24BPP, , , , , True)
        If fi_Check = False Then
            Message "Save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "Save complete."
        End If
    Else
        Message "Save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
        
End Sub

'Save to Targa (TGA) format.
Public Sub SaveTGAImage(ByVal imageID As Long, ByVal TGAPath As String)
    
    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(pdImages(imageID).mainLayer.getLayerDC)
    
    'Use that handle to save the image to TGA format
    If fi_DIB <> 0 Then
    
        'In the future, the color depth of the output file should be user-controllable via a form.  Right now, however, just use
        ' the color depth of the current image
        Dim fi_OutputColorDepth As FREE_IMAGE_COLOR_DEPTH
        If pdImages(imageID).mainLayer.getLayerColorDepth = 24 Then
            fi_OutputColorDepth = FICD_24BPP
        Else
            fi_OutputColorDepth = FICD_32BPP
        End If
    
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TGAPath, FIF_TARGA, FILO_TARGA_DEFAULT, fi_OutputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "Save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "Save complete."
        End If
    Else
        Message "Save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib

End Sub

'Save to JPEG using the FreeImage library.  This is faster and more reliable than using GDI+.
Public Sub SaveJPEGImage(ByVal imageID As Long, ByVal JPEGPath As String, ByVal jQuality As Long)
    
    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    'Copy the image into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    If tmpLayer.getLayerColorDepth = 32 Then tmpLayer.compositeBackgroundColor 255, 255, 255
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(tmpLayer.getLayerDC)
    
    'Use that handle to save the image to JPEG format
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, JPEGPath, FIF_JPEG, JPEG_OPTIMIZE + jQuality, FICD_24BPP, , , , , True)
        If fi_Check = False Then
            Message "Save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "Save complete."
        End If
    Else
        Message "Save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
    
End Sub

'Save a TIFF (Tagged Image File Format) image via FreeImage.  GDI+ can also do this.
Public Sub SaveTIFImage(ByVal imageID As Long, ByVal TIFPath As String)
    
    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save could not be completed without FreeImage library access."
        Exit Sub
    End If

    'Load FreeImage into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(pdImages(imageID).mainLayer.getLayerDC)
    
    'Use that handle to save the image to TIFF format
    If fi_DIB <> 0 Then
    
        'In the future, the color depth of the output file should be user-controllable via a form.  Right now, however, just use
        ' the color depth of the current image
        Dim fi_OutputColorDepth As FREE_IMAGE_COLOR_DEPTH
        If pdImages(imageID).mainLayer.getLayerColorDepth = 24 Then
            fi_OutputColorDepth = FICD_24BPP
        Else
            fi_OutputColorDepth = FICD_32BPP
        End If
        
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TIFPath, FIF_TIFF, TIFF_NONE, fi_OutputColorDepth, , , , , True)
        If fi_Check = False Then
            Message "Save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report."
        Else
            Message "Save complete."
        End If
    Else
        Message "Save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report."
    End If
    
    'Release FreeImage from memory
    FreeLibrary hLib
        
End Sub

