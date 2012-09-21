Attribute VB_Name = "FreeImage_Expanded_Interface"
'***************************************************************************
'FreeImage Interface (Advanced)
'Copyright ©2000-2012 by Tanner Helland
'Created: 3/September/12
'Last updated: 3/September/12
'Last update: initial build
'
'This module represents a new - and significantly more complex - approach to loading images via the FreeImage libary.
' The current FreeImage implementation (LoadFreeImageV3) relies on FreeImage to make its own decisions regarding format,
' color-depth, and color-space conversion.  The decisions FreeImage typically makes are not always consistent with the
' way this data is eventually handled in VB, so some explicit coding is necessary on a per-image-format basis.
'
'This module is based heavily on the work of Herman Liu, to whom I owe a large debt of gratitude.
'
'Additionally, this module continues to rely heavily on Carsten Klein's FreeImage wrapper for VB (included in this project
' as Outside_FreeImageV3; see that file for license details).  Thanks to Carsten for his work on integrating FreeImage
' into classic VB.
'
'***************************************************************************

Option Explicit

'DIB declarations
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
    
'Load an image via FreeImage.  It is assumed that the source file has already been vetted for things like "does it exist?"
Public Function LoadFreeImageV3_Advanced(ByVal srcFilename As String, ByRef dstLayer As pdLayer, ByRef dstImage As pdImage) As Boolean

    On Error GoTo FreeImageV3_AdvancedError
    
    'Double-check that FreeImage.dll was located at start-up
    If FreeImageEnabled = False Then
        LoadFreeImageV3_Advanced = False
        Exit Function
    End If
    
    'Load the FreeImage library from the plugin directory
    Dim hFreeImgLib As Long
    hFreeImgLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Analyzing filetype..."
    
    'While we could manually test our extension against the FreeImage database, it is capable of doing so itself.
    'First, check the file header to see if it matches a known head type
    Dim fileFIF As FREE_IMAGE_FORMAT
    fileFIF = FreeImage_GetFileType(srcFilename)
    
    'For certain filetypes (CUT, MNG, PCD, TARGA and WBMP, according to the FreeImage documentation), the lack of a reliable
    ' header may prevent GetFileType from working.  As a result, double-check the file using its file extension.
    If fileFIF = FIF_UNKNOWN Then fileFIF = FreeImage_GetFIFFromFilename(srcFilename)
    
    'By this point, if the file still doesn't show up in FreeImage's database, abandon the import attempt.
    If Not FreeImage_FIFSupportsReading(fileFIF) Then
        Message "Filetype not supported by FreeImage.  Import abandoned."
        LoadFreeImageV3_Advanced = False
        Exit Function
    End If
    
    'Store this file format inside the relevant pdImage object
    dstImage.OriginalFileFormat = fileFIF
    
    Message "Preparing import flags..."
    
    'Certain filetypes offer import options.  Check the FreeImage type to see if we want to enable any optional flags.
    Dim fi_ImportFlags As FREE_IMAGE_LOAD_OPTIONS
    fi_ImportFlags = 0
    
    'For JPEGs, specify a preference for accuracy and quality over load speed under normal circumstances,
    ' but when performing a batch conversion choose the reverse (speed over accuracy).
    If fileFIF = FIF_JPEG Then
        If MacroStatus = MacroBATCH Then
            fi_ImportFlags = FILO_JPEG_FAST
        Else
            fi_ImportFlags = FILO_JPEG_ACCURATE
        End If
    End If
    
    'For icons, we prefer a white background (default is black).
    ' NOTE: this is disabled, because it uses the AND mask incorrectly for mixed-format icons
    'If fileFIF = FIF_ICO Then fi_ImportFlags = FILO_ICO_MAKEALPHA
    
    Message "Importing image from file..."
    
    'With all flags set and filetype correctly determined, import the image
    Dim fi_hDIB As Long
    fi_hDIB = FreeImage_Load(fileFIF, srcFilename, fi_ImportFlags)
    
    'If an empty handle is returned, abandon the import attempt.
    If fi_hDIB = 0 Then
        Message "Import via FreeImage failed (blank handle)."
        LoadFreeImageV3_Advanced = False
        Exit Function
    End If
    
    'Before we continue the import, we need to make sure the pixel data is in a format appropriate for PhotoDemon.
    
    Message "Analyzing color depth..."
    
    'First thing we want to check is the color depth.  PhotoDemon is designed around 16 million color images.  This could
    ' change in the future, but for now, force high-bit-depth images to a more appropriate 24 or 32bpp.
    Dim fi_BPP As Long
    fi_BPP = FreeImage_GetBPP(fi_hDIB)
    
    Dim new_hDIB As Long
    
    Dim fi_hasTransparency As Boolean
    Dim fi_transparentEntries As Long
        
    'First, check source images without an alpha channel.  Convert these using the superior tone mapping method.
    If (fi_BPP = 48) Or (fi_BPP = 96) Then
        'While images with these bit-depths may not have an alpha channel, they can have binary transparency - check for that now.
        ' (Note: as of FreeImage 3.15 binary bit-depths are not detected correctly.  That said, they may someday be supported - so
        ' I've implemented two checks to cover both contingencies.
        fi_hasTransparency = FreeImage_IsTransparent(fi_hDIB)
        fi_transparentEntries = FreeImage_GetTransparencyCount(fi_hDIB)
        
        If fi_hasTransparency Or (fi_transparentEntries <> 0) Then
            new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, True)
            fi_hDIB = new_hDIB
        Else
            new_hDIB = FreeImage_ToneMapping(fi_hDIB, FITMO_REINHARD05)
            FreeImage_UnloadEx fi_hDIB
            fi_hDIB = new_hDIB
        End If
    End If
    
    'Because tone mapping may not preserve alpha channels (the FreeImage documentation is unclear on this),
    ' high-bit-depth images with alpha channels are converted using a traditional method, which is inferior from a
    ' quality standpoint, but at least guaranteed to preserve the alpha channel data.
    If (fi_BPP = 64) Or (fi_BPP = 128) Then
        new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, True)
        fi_hDIB = new_hDIB
    End If
        
    'Similarly, check for low-bit-depth images
    If fi_BPP < 24 Then
        
        'Next, check to see if this is actually a high-bit-depth grayscale image masquerading as a low-bit-depth RGB image
        Dim fi_imageType As FREE_IMAGE_TYPE
        fi_imageType = FreeImage_GetImageType(fi_hDIB)
        
        'If it is such a grayscale image, it requires a unique conversion operation
        If fi_imageType = FIT_UINT16 Then
            
            Message "Tone-mapping high-bit-depth grayscale image to 24bpp..."
            
            'First, convert it to a high-bit depth RGB image
            fi_hDIB = FreeImage_ConvertToRGB16(fi_hDIB)
            
            'Now use tone-mapping to reduce it back to 24bpp or 32bpp (contingent on the presence of transparency)
            fi_hasTransparency = FreeImage_IsTransparent(fi_hDIB)
            fi_transparentEntries = FreeImage_GetTransparencyCount(fi_hDIB)
        
            If fi_hasTransparency Or (fi_transparentEntries <> 0) Then
                new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, True)
                fi_hDIB = new_hDIB
            Else
                new_hDIB = FreeImage_ToneMapping(fi_hDIB, FITMO_REINHARD05)
                FreeImage_UnloadEx fi_hDIB
                fi_hDIB = new_hDIB
            End If
            
        Else
        
            'Conversion to higher bit depths is contingent on the presence of an alpha channel
            fi_hasTransparency = FreeImage_IsTransparent(fi_hDIB)
        
            'Images with an alpha channel are converted to 32 bit.  Without, 24.
            If fi_hasTransparency = True Then
                new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, True)
                fi_hDIB = new_hDIB
            Else
                new_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_24BPP, True)
                fi_hDIB = new_hDIB
            End If
            
        End If
        
    End If
    
    'By this point, we have loaded the image, and it is guaranteed to be at 24 or 32 bit color depth.  Verify it one final time.
    fi_BPP = FreeImage_GetBPP(fi_hDIB)
    
    'We are now finally ready to load the image.
    
    Message "Requesting memory for image transfer..."
    
    'Get width and height from the file, and create a new layer to match
    Dim fi_Width As Long, fi_Height As Long
    fi_Width = FreeImage_GetWidth(fi_hDIB)
    fi_Height = FreeImage_GetHeight(fi_hDIB)
    
    Dim creationSuccess As Boolean
    creationSuccess = dstLayer.createBlank(fi_Width, fi_Height, fi_BPP)
    
    'Make sure the blank DIB creation worked
    If creationSuccess = False Then
        Message "Import via FreeImage failed (couldn't create DIB)."
        LoadFreeImageV3_Advanced = False
        Exit Function
    End If
    
    'Copy the bits from the FreeImage DIB to our DIB
    'NOTE: investigate using AlphaBlend to copy the bits, with SourceConstantAlpha set to 255 (per http://msdn.microsoft.com/en-us/library/dd183393%28v=vs.85%29.aspx)
    ' This may be a way to preserve the alpha channel... assuming that this SetDIBitsToDevice is actually the problem, which it may not be.
    SetDIBitsToDevice dstLayer.getLayerDC, 0, 0, fi_Width, fi_Height, 0, 0, 0, fi_Height, ByVal FreeImage_GetBits(fi_hDIB), ByVal FreeImage_GetInfo(fi_hDIB), 0&
              
    'With the image bits now safely in our care, release the FreeImage DIB
    FreeImage_UnloadEx fi_hDIB
    
    'Release the FreeImage library
    FreeLibrary hFreeImgLib
    
    'If the loaded image has an alpha-channel, recomposite the image against a white background (by default it will be black)
    If dstLayer.getLayerColorDepth = 32 Then
        Message "Performing final alpha channel recomposite..."
        dstLayer.compositeBackgroundColor
    End If
    
    'Mark this load as successful
    LoadFreeImageV3_Advanced = True
    
    Exit Function
    
    
FreeImageV3_AdvancedError:

    'Release the FreeImage DIB if available
    If fi_hDIB <> 0 Then FreeImage_UnloadEx fi_hDIB
    
    'Release the FreeImage library
    If hFreeImgLib <> 0 Then FreeLibrary hFreeImgLib
    
    'Display a relevant error message
    Message "Import via FreeImage failed (Err#" & Err.Number & ")"
    
    'Mark this load as unsuccessful
    LoadFreeImageV3_Advanced = False
    
End Function
