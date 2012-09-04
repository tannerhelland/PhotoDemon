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
' as Outside_FreeImageV3; see that file for license details).
'
'***************************************************************************

Option Explicit

'DIB declarations
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
    
    
'Load an image via FreeImage.  It is assumed that the source file has already been vetted for things like "does it exist?"
Public Function LoadFreeImageV3_Advanced(ByVal srcFilename As String) As Boolean

    On Error GoTo FreeImageV3_AdvancedError
    
    'Double-check that FreeImage.dll was located at start-up
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        pdImages(CurrentImage).IsActive = False
        Unload FormMain.ActiveForm
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
    
    Message "Preparing import flags..."
    
    'Certain filetypes offer import options.  Check the FreeImage type to see if we want to enable any optional flags.
    Dim fi_ImportFlags As FREE_IMAGE_LOAD_OPTIONS
    fi_ImportFlags = 0
    
    'For JPEGs, specify a preference for accuracy and quality over load speed
    If fileFIF = FIF_JPEG Then fi_ImportFlags = FILO_JPEG_ACCURATE
    
    'For icons, we prefer a white background (default is black).
    If fileFIF = FIF_ICO Then fi_ImportFlags = FILO_ICO_MAKEALPHA
    
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
    
    'First, check source images without an alpha channel.  Convert these using the superior tone mapping method.
    If (fi_BPP = 48) Or (fi_BPP = 96) Then fi_hDIB = FreeImage_ToneMapping(fi_hDIB, FITMO_REINHARD05)
    
    'Because tone mapping may not preserve alpha channels (the FreeImage documentation is unclear on this),
    ' high-bit-depth images with alpha channels are converted using a traditional method, which is inferior from a
    ' quality standpoint, but at least guaranteed to preserve the alpha channel data.
    If (fi_BPP = 64) Or (fi_BPP = 128) Then fi_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, True)
    
    'Similarly, check for low-bit-depth images
    If fi_BPP < 24 Then
        
        'Conversion to higher bit depths is contingent on the presence of an alpha channel
        Dim fi_hasTransparency As Boolean
        fi_hasTransparency = FreeImage_IsTransparent(fi_hDIB)
        
        'Images with an alpha channel are converted to 32 bit.  Without, 24.
        If fi_hasTransparency = True Then
            fi_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_32BPP, True)
        Else
            fi_hDIB = FreeImage_ConvertColorDepth(fi_hDIB, FICF_RGB_24BPP, True)
        End If
        
    End If
    
    'By this point, we have loaded the image, and it is guaranteed to be at 24 or 32 bit color depth.
    'The last thing we need to do is specific to transparent images.  FreeImage will default to loading transparent images
    ' against a black background.  We prefer white.  FreeImage's composite function can be used to handle this.
    fi_BPP = FreeImage_GetBPP(fi_hDIB)
    If fi_BPP = 32 Then
        
        Message "Recompositing alpha channel..."
        
        Dim tmpWhite As RGBQUAD
        With tmpWhite
            .rgbBlue = 255
            .rgbGreen = 255
            .rgbRed = 255
            .rgbReserved = 0
        End With
        fi_hDIB = FreeImage_Composite(fi_hDIB, , tmpWhite)
    End If
    
    'We are now finally ready to load the image.
    
    Message "Requesting memory for image transfer..."
    
    'Get width and height from the file, and create a new layer to match
    Dim fi_Width As Long, fi_Height As Long
    fi_Width = FreeImage_GetWidth(fi_hDIB)
    fi_Height = FreeImage_GetHeight(fi_hDIB)
    
    Dim creationSuccess As Boolean
    creationSuccess = pdImages(CurrentImage).mainLayer.createBlank(fi_Width, fi_Height, fi_BPP)
    
    'Make sure the blank DIB creation worked
    If creationSuccess = False Then
        Message "Import via FreeImage failed (couldn't create DIB)."
        LoadFreeImageV3_Advanced = False
        Exit Function
    End If
    
    'Copy the bits from the FreeImage DIB to our DIB
    SetDIBitsToDevice pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, fi_Width, fi_Height, 0, 0, 0, fi_Height, ByVal FreeImage_GetBits(fi_hDIB), ByVal FreeImage_GetInfo(fi_hDIB), 0&
              
    'With the image bits now safely in our care, release the FreeImage DIB
    FreeImage_UnloadEx fi_hDIB
    
    'Release the FreeImage library
    FreeLibrary hFreeImgLib
    
    'Mark this load as successful
    LoadFreeImageV3_Advanced = True
    
    Exit Function
    
    
FreeImageV3_AdvancedError:

    'Reset the mouse pointer
    FormMain.MousePointer = vbDefault

    'We'll use this string to hold additional error data
    Dim AddInfo As String
    
    'This variable stores the message box type
    Dim mType As VbMsgBoxStyle
    
    'Tracks the user input from the message box
    Dim msgReturn As VbMsgBoxResult

    'FreeImage throws Error #5 if an invalid image is loaded
    If Err.Number = 5 Then
        AddInfo = "You have attempted to load an invalid picture.  This can happen if a file does not contain image data, or if it contains image data in an unsupported format." & vbCrLf & vbCrLf & "- If you downloaded this image from the Internet, the download may have terminated prematurely.  Please try downloading the image again." & vbCrLf & vbCrLf & "- If this image file came from a digital camera, scanner, or other image editing program, it's possible that " & PROGRAMNAME & " simply doesn't understand this particular file format.  Please save the image in a generic format (such as bitmap or JPEG), then reload it."
        Message "Invalid picture.  Image load cancelled."
        mType = vbCritical + vbOKOnly
    End If
    
    'Create the message box to return the error information
    msgReturn = MsgBox(PROGRAMNAME & " has experienced an error.  Details on the problem include:" & vbCrLf & vbCrLf & _
    "Error number " & Err.Number & vbCrLf & _
    "Description: " & Err.Description & vbCrLf & vbCrLf & _
    AddInfo & vbCrLf & vbCrLf & _
    "Sorry for the inconvenience," & vbCrLf & _
    "-Tanner Helland" & vbCrLf & PROGRAMNAME & " Developer" & vbCrLf & _
    "www.tannerhelland.com/contact", mType, PROGRAMNAME & " Error Handler: #" & Err.Number)
    
    'If an error was thrown, unload the active form (since it will just be empty and pictureless)
    pdImages(CurrentImage).IsActive = False
    Unload FormMain.ActiveForm

    'Mark this load as unsuccessful
    LoadFreeImageV3_Advanced = False
    
End Function
