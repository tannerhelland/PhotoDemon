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
Public Sub SaveBMP(ByVal ImageID As Long, ByVal BMPPath As String)
    
    Message "Saving image..."
    
    'The layer class is capable of doing this without any outside help.
    pdImages(ImageID).mainLayer.writeToBitmapFile BMPPath
    
    Message "Save complete."
    
End Sub

'Save the current image to PhotoDemon's native PDI format
Public Sub SavePhotoDemonImage(ByVal ImageID As Long, ByVal PDIPath As String)
    
    Message "Saving image..."

    'First, have the layer write itself to file in BMP format
    pdImages(ImageID).mainLayer.writeToBitmapFile PDIPath
    
    'Then compress the file using zLib
    CompressFile PDIPath
    
    Message "Save complete."
    
End Sub

Public Sub SaveGIFImage(ByVal ImageID As Long, ByVal GIFPath As String)

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
    fi_DIB = FreeImage_CreateFromDC(pdImages(ImageID).mainLayer.getLayerDC)
    
    'Use that handle to save the image to GIF format, with required 8bpp (256 color) conversion
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, GIFPath, FIF_GIF, , FICD_8BPP, , , , , False)
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

Public Sub SavePNGImage(ByVal ImageID As Long, ByVal PNGPath As String, Optional ByVal PNGColorDepth As Long = &H18)

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
    fi_DIB = FreeImage_CreateFromDC(pdImages(ImageID).mainLayer.getLayerDC)
    
    'Use that handle to save the image to PNG format
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, PNGPath, FIF_PNG, FISO_PNG_Z_BEST_COMPRESSION, PNGColorDepth, , , , , False)
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
Public Sub SavePPMImage(ByVal ImageID As Long, ByVal PPMPath As String)

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
    fi_DIB = FreeImage_CreateFromDC(pdImages(ImageID).mainLayer.getLayerDC)
    
    'Use that handle to save the image to PPM format (ASCII)
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, PPMPath, FIF_PPM, , FICD_24BPP, , , , , False)
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
Public Sub SaveTGAImage(ByVal ImageID As Long, ByVal TGAPath As String)
    
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
    fi_DIB = FreeImage_CreateFromDC(pdImages(ImageID).mainLayer.getLayerDC)
    
    'Use that handle to save the image to TGA format
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TGAPath, FIF_TARGA, FILO_TARGA_DEFAULT, FICD_24BPP, , , , , False)
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
Public Sub SaveJPEGImageUsingFreeImage(ByVal ImageID As Long, ByVal JPEGPath As String, ByVal jQuality As Long)
    
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
    fi_DIB = FreeImage_CreateFromDC(pdImages(ImageID).mainLayer.getLayerDC)
    
    'Use that handle to save the image to JPEG format
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, JPEGPath, FIF_JPEG, JPEG_OPTIMIZE + jQuality, FICD_24BPP, , , , , False)
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

Public Sub SaveJPEGImageUsingVB(ByVal ImageID As Long, ByVal JPEGPath As String, ByVal Quality As Long)
    
    'Use John's JPEG class
    Dim m_Jpeg As cJpeg
    Set m_Jpeg = New cJpeg
    
    m_Jpeg.Quality = Quality
    
    'The image can only be sampled AFTER the quality has been set
    GetImageData
    m_Jpeg.SampleHDC pdImages(ImageID).containingForm.BackBuffer.hDC, PicWidthL + 1, PicHeightL + 1
    
    'Delete file if it exists
    If FileExist(JPEGPath) Then
        Message "Deleting old file..."
        Kill JPEGPath
    End If
    
    'Save the JPG file
    Message "Saving JPEG image..."
    m_Jpeg.SaveFile JPEGPath

    'Save memory (not really necessary, but I do it out of habit)
    Set m_Jpeg = Nothing
    
    Message "Save complete."
    
End Sub

Public Sub SaveTIFImage(ByVal ImageID As Long, ByVal TIFPath As String)
    
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
    fi_DIB = FreeImage_CreateFromDC(pdImages(ImageID).mainLayer.getLayerDC)
    
    'Use that handle to save the image to TIFF format
    If fi_DIB <> 0 Then
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveEx(fi_DIB, TIFPath, FIF_TIFF, TIFF_NONE, FICD_24BPP, , , , , False)
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

