Attribute VB_Name = "Saving"
'***************************************************************************
'File Saving Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 03/June/12
'Last update: Rewrite all save code against FreeImage v3.15.3, which is faster, cleaner,
'             and more reliable than the ten-year-old version PD was originally designed
'             to work with.  The new version also enabled support for new file formats
'             (including GIF!).
'
'Module for handling all image saving.  It contains pretty much every routine that I find useful;
' the majority of the functions are simply interfaces to FreeImage, so if that is not enabled than
' only a subset of these matter.
'
'***************************************************************************

Option Explicit

Public Sub SaveBMP(ByVal ImageID As Long, ByVal BitmapFileName As String)
    Message "Saving image..."
    If FileExist(BitmapFileName) Then Kill BitmapFileName
    SavePicture pdImages(ImageID).containingForm.BackBuffer.Image, BitmapFileName
    Message "Save complete."
End Sub

Public Sub SavePhotoDemonImage(ByVal ImageID As Long, ByVal PDIPath As String)
    Message "Saving image..."

    SavePicture pdImages(ImageID).containingForm.BackBuffer.Image, PDIPath
    'Need to add error handling...
    CompressFile PDIPath
    
    Message "Save complete."
    
End Sub

Public Sub SaveGIFImage(ByVal ImageID As Long, ByVal GIFPath As String)

    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-ins (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy this file into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
        Exit Sub
    End If
    
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    SavePictureEx pdImages(ImageID).containingForm.BackBuffer.Picture, GIFPath, FIF_GIF, , FICD_8BPP
    
    Message "Save complete."
    
End Sub

Public Sub SavePNGImage(ByVal ImageID As Long, ByVal PNGPath As String, Optional ByVal PNGColorDepth As Long = &H18)

    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-ins (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy this file into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
        Exit Sub
    End If
    
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    SavePictureEx pdImages(ImageID).containingForm.BackBuffer.Picture, PNGPath, FIF_PNG, FISO_PNG_Z_BEST_COMPRESSION, PNGColorDepth
    
    Message "Save complete."
    
End Sub

'IMPORTANT NOTE: Only ASCII format PPM is currently enabled.  RAW IS NOT YET SUPPORTED!
Public Sub SavePPMImage(ByVal ImageID As Long, ByVal PPMPath As String)

    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-ins (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy this file into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
        Exit Sub
    End If
    
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    SavePictureEx pdImages(ImageID).containingForm.BackBuffer.Picture, PPMPath, FIF_PPM, , FICD_24BPP
    
    Message "Save complete."
    
End Sub

Public Sub SaveTGAImage(ByVal ImageID As Long, ByVal TGAPath As String)
    
    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-ins (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy this file into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
        Exit Sub
    End If
    
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    SavePictureEx pdImages(ImageID).containingForm.BackBuffer.Picture, TGAPath, FIF_TARGA, FILO_TARGA_DEFAULT, FICD_24BPP
    
    Message "Save complete."

End Sub

Public Sub SaveJPEGImageUsingFreeImage(ByVal ImageID As Long, ByVal JPEGPath As String, ByVal jQuality As Long)
    
    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-ins (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy this file into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
        Exit Sub
    End If
    
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    SavePictureEx pdImages(ImageID).containingForm.BackBuffer.Picture, JPEGPath, FIF_JPEG, JPEG_OPTIMIZE + jQuality, FICD_24BPP
    
    Message "Save complete."
    
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
        MsgBox "The FreeImage interface plug-ins (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy this file into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
        Exit Sub
    End If
    
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    SavePictureEx pdImages(ImageID).containingForm.BackBuffer.Picture, TIFPath, FIF_TIFF, TIFF_NONE, FICD_24BPP
    
    Message "Save complete."
    
End Sub

'PCX exporting is temporarily disabled.  It's such a rare use-case that I don't want to invest energy in it right now.
'Public Sub SavePCXImage(ByVal ImageID As Long, ByVal PCXPath As String, ByVal colorDepth As Long, ByVal useRLE As Long)
    'Dim TempPCX As New SavePCX
    'Set TempPCX = Nothing
   '
   ' Message "Saving image..."
   ' If FileExist(PCXPath) Then Kill PCXPath
   '
   ' TempPCX.SavePCXinFile PCXPath, pdImages(ImageID).containingForm.BackBuffer, colorDepth, useRLE
   '
   ' Set TempPCX = Nothing
   ' Message "Image saved."
    
'End Sub

