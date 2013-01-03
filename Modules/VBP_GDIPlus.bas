Attribute VB_Name = "GDI_Plus"
'***************************************************************************
'GDI+ Interface
'Copyright ©2011-2013 by Tanner Helland
'Created: 1/September/12
'Last updated: 4/September/12
'Last update: full support for image saving (GIF, JPEG, PNG, TIFF) via GDI+.  These are all considered fallbacks; if FreeImage
'              is found, it will be given first priority, but at least this will allow PSC users to export four additional
'              file formats even if they don't download the FreeImage plugin.
'
'This interface provides a means for interacting with the unnecessarily complex (and overwrought) GDI+ module.  GDI+ is
' primarily used as a fallback for image loading and saving if the FreeImage DLL cannot be found.
'
'These routines are adapted from the work of a number of other talented VB programmers.  Since GDI+ is not well-documented
' for VB users, I have pieced this module together from the following pieces of code:
' Avery P's initial GDI+ deconstruction: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
' Carles P.V.'s iBMP implementation: http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
' Robert Rayment's PaintRR implementation: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1
' Many thanks to these individuals for their outstanding work on graphics in VB.
'
'***************************************************************************

Option Explicit

'GDI+ Enums
Public Enum GDIPlusImageFormat
    [ImageBMP] = 0
    [ImageGIF] = 1
    [ImageJPEG] = 2
    [ImagePNG] = 3
    [ImageTIFF] = 4
End Enum

Public Enum GDIPlusStatus
    [OK] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum

Private Enum EncoderParameterValueType
    [EncoderParameterValueTypeByte] = 1
    [EncoderParameterValueTypeASCII] = 2
    [EncoderParameterValueTypeShort] = 3
    [EncoderParameterValueTypeLong] = 4
    [EncoderParameterValueTypeRational] = 5
    [EncoderParameterValueTypeLongRange] = 6
    [EncoderParameterValueTypeUndefined] = 7
    [EncoderParameterValueTypeRationalRange] = 8
End Enum

Private Enum EncoderValue
    [EncoderValueColorTypeCMYK] = 0
    [EncoderValueColorTypeYCCK] = 1
    [EncoderValueCompressionLZW] = 2
    [EncoderValueCompressionCCITT3] = 3
    [EncoderValueCompressionCCITT4] = 4
    [EncoderValueCompressionRle] = 5
    [EncoderValueCompressionNone] = 6
    [EncoderValueScanMethodInterlaced] = 7
    [EncoderValueScanMethodNonInterlaced] = 8
    [EncoderValueVersionGif87] = 9
    [EncoderValueVersionGif89] = 10
    [EncoderValueRenderProgressive] = 11
    [EncoderValueRenderNonProgressive] = 12
    [EncoderValueTransformRotate90] = 13
    [EncoderValueTransformRotate180] = 14
    [EncoderValueTransformRotate270] = 15
    [EncoderValueTransformFlipHorizontal] = 16
    [EncoderValueTransformFlipVertical] = 17
    [EncoderValueMultiFrame] = 18
    [EncoderValueLastFrame] = 19
    [EncoderValueFlush] = 20
    [EncoderValueFrameDimensionTime] = 21
    [EncoderValueFrameDimensionResolution] = 22
    [EncoderValueFrameDimensionPage] = 23
    [EncoderValueColorTypeGray] = 24
    [EncoderValueColorTypeRGB] = 25
End Enum

Private Type clsid
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type

Private Type ImageCodecInfo
    ClassID           As clsid
    FormatID          As clsid
    CodecName         As Long
    DllName           As Long
    formatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type

Private Type EncoderParameter
    Guid           As clsid
    NumberOfValues As Long
    Type           As EncoderParameterValueType
    Value          As Long
End Type

Private Type EncoderParameters
    Count     As Long
    Parameter As EncoderParameter
End Type

Private Const EncoderCompression      As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Private Const EncoderColorDepth       As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Private Const EncoderColorSpace       As String = "{AE7A62A0-EE2C-49D8-9D07-1BA8A927596E}"
Private Const EncoderScanMethod       As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Private Const EncoderVersion          As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Private Const EncoderRenderMethod     As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Private Const EncoderQuality          As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const EncoderTransformation   As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Private Const EncoderLuminanceTable   As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Private Const EncoderChrominanceTable As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Private Const EncoderSaveAsCMYK       As String = "{A219BBC9-0A9D-4005-A3EE-3A421B8BB06C}"
Private Const EncoderSaveFlag         As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
Private Const CodecIImageBytes        As String = "{025D1823-6C7D-447B-BBDB-A3CBC3DFA2FC}"

'GDI+ recognizes a variety of pixel formats:
'Public Const PixelFormatIndexed = &H10000           ' Indexes into a palette
'Public Const PixelFormatGDI = &H20000               ' Is a GDI-supported format
'Public Const PixelFormatAlpha = &H40000             ' Has an alpha component
'Public Const PixelFormatPAlpha = &H80000            ' Pre-multiplied alpha
'Public Const PixelFormatExtended = &H100000         ' Extended color 16 bits/channel
'Public Const PixelFormatCanonical = &H200000
'
'Public Const PixelFormatUndefined = 0
'
'Public Const PixelFormat1bppIndexed = &H30101
'Public Const PixelFormat4bppIndexed = &H30402
'Public Const PixelFormat8bppIndexed = &H30803
'Public Const PixelFormat16bppGreyScale = &H101004
'Public Const PixelFormat16bppRGB555 = &H21005
'Public Const PixelFormat16bppRGB565 = &H21006
'Public Const PixelFormat16bppARGB1555 = &H61007

Private Const PixelFormat24bppRGB = &H21808
Private Const PixelFormat32bppARGB = &H26200A
'Private Const PixelFormat32bppRGB = &H22009
Private Const PixelFormat32bppPARGB = &HE200B
'Public Const PixelFormat48bppRGB = &H10300C
'Public Const PixelFormat64bppARGB = &H34400D
'Public Const PixelFormat64bppPARGB = &H1C400E
'Public Const PixelFormatMax = 15 '&HF

'GDI+ required types
Private Type GDIPlusStartupInput
    GDIPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

'OleCreatePictureIndirect types
Private Type PictDesc
    Size       As Long
    Type       As Long
    hBmpOrIcon As Long
    hPal       As Long
End Type

'BITMAP types
Private Type RGBQUAD
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Private Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    ImageSize As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    Colorused As Long
    ColorImportant As Long
End Type

Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Colors(0 To 255) As RGBQUAD
End Type

'Workaround for VB not exposing an IStream interface
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
    
'Start-up and shutdown
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef Token As Long, ByRef InputBuf As GDIPlusStartupInput, Optional ByVal OutputBuffer As Long = 0&) As GDIPlusStatus
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GDIPlusStatus

'Load image from file, process said file, etc.
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal lStride As Long, ByVal ePixelFormat As Long, ByRef Scan0 As Any, ByRef pBitmap As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal Bitmap As Long, hBmpReturn As Long, ByVal Background As Long) As GDIPlusStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GDIPlusStatus
Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BITMAPINFO, gdiBitmapData As Any, Bitmap As Long) As GDIPlusStatus
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As GDIPlusStatus
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As GDIPlusStatus
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As String, clsidEncoder As clsid, encoderParams As Any) As GDIPlusStatus
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As GDIPlusStatus

'OleCreatePictureIndirect is used to convert GDI+ images to VB's preferred StdPicture
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long

'CLSIDFromString is used to convert a mimetype into a CLSID required by the GDI+ image encoder
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pclsid As clsid) As Long

'Necessary for converting between ASCII and UNICODE strings
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long

'CopyMemory
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal cb As Long) As Long

'When GDI+ is initialized, it will assign us a token.  We use this to release GDI+ when the program terminates.
Private GDIPlusToken As Long

'Use GDI+ to load a picture into a StdPicture object - not ideal, as some information will be lost in the transition, but since
' this is only a fallback from FreeImage I'm not going out of my way to improve it.
Public Function GDIPlusLoadPicture(ByVal srcFilename As String, ByRef targetPicture As StdPicture) As Boolean

    'Used to hold the return values of various GDI+ calls
    Dim GDIPlusReturn As Long
      
    'Use GDI+ to load the image
    Dim hImage As Long
    GDIPlusReturn = GdipLoadImageFromFile(StrPtr(srcFilename), hImage)
    
    If (GDIPlusReturn <> [OK]) Then
        GdipDisposeImage hImage
        GDIPlusLoadPicture = False
        Exit Function
    End If
    
    'Copy the GDI+ image into a standard bitmap
    Dim hBitmap As Long
    GDIPlusReturn = GdipCreateHBITMAPFromBitmap(hImage, hBitmap, vbBlack)
    
    If (GDIPlusReturn <> [OK]) Then
        GdipDisposeImage hImage
        GDIPlusLoadPicture = False
        Exit Function
    End If
    
    'Now we can release the GDI+ copy of the image
    GDIPlusReturn = GdipDisposeImage(hImage)
    
    'Assuming the load/unload went okay, prepare to copy the bitmap object into an StdPicture
    If (GDIPlusReturn = [OK]) Then

        'Prepare the header required by OleCreatePictureIndirect
        Dim picHeader As PictDesc
        With picHeader
            .Size = Len(picHeader)
            .Type = vbPicTypeBitmap
            .hBmpOrIcon = hBitmap
            .hPal = 0
        End With
        
        'Populate the magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        Dim aGuid(0 To 3) As Long
        aGuid(0) = &H7BF80980
        aGuid(1) = &H101ABF32
        aGuid(2) = &HAA00BB8B
        aGuid(3) = &HAB0C3000
        
        'Using the bitmap indirectly created by GDI+, build an identical StdPicture object
        OleCreatePictureIndirect picHeader, aGuid(0), -1, targetPicture
    
        GDIPlusLoadPicture = True
    
    Else
        GDIPlusLoadPicture = False
    End If

End Function

'Save an image using GDI+.  Per the current save spec, ImageID must be specified.
' Additional save options are currently available for JPEGs (save quality, range [1,100]) and TIFFs (compression type).
Public Function GDIPlusSavePicture(ByVal imageID As Long, ByVal dstFilename As String, ByVal imgFormat As GDIPlusImageFormat, ByVal outputColorDepth As Long, Optional ByVal JPEGQuality As Long = 92) As Boolean

    On Error GoTo GDIPlusSaveError

    Message "Initializing GDI+..."

    'If the output format is 24bpp (e.g. JPEG) but the input image is 32bpp, composite it against white
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(imageID).mainLayer
    If tmpLayer.getLayerColorDepth <> 24 And imgFormat = [ImageJPEG] Then tmpLayer.compositeBackgroundColor 255, 255, 255

    'Begin by creating a generic bitmap header for the current layer
    Dim imgHeader As BITMAPINFO
    
    With imgHeader.Header
        .Size = Len(imgHeader.Header)
        .Planes = 1
        .BitCount = tmpLayer.getLayerColorDepth
        .Width = tmpLayer.getLayerWidth
        .Height = -tmpLayer.getLayerHeight
    End With

    'Use GDI+ to create a GDI+-compatible bitmap
    Dim GDIPlusReturn As Long
    Dim hImage As Long
    
    Message "Creating GDI+ compatible image copy..."
        
    'Different GDI+ calls are required for different color depths. GdipCreateBitmapFromGdiDib leads to a blank
    ' alpha channel for 32bpp images, so use GdipCreateBitmapFromScan0 in that case.
    If tmpLayer.getLayerColorDepth = 32 Then
        
        'Use GdipCreateBitmapFromScan0 to create a 32bpp DIB with alpha preserved
        GDIPlusReturn = GdipCreateBitmapFromScan0(tmpLayer.getLayerWidth, tmpLayer.getLayerHeight, tmpLayer.getLayerWidth * 4, PixelFormat32bppARGB, ByVal tmpLayer.getLayerDIBits, hImage)
    
    Else
        GDIPlusReturn = GdipCreateBitmapFromGdiDib(imgHeader, ByVal tmpLayer.getLayerDIBits, hImage)
    End If
    
    If (GDIPlusReturn <> [OK]) Then
        GdipDisposeImage hImage
        GDIPlusSavePicture = False
        Exit Function
    End If
    
    'Certain image formats require extra parameters, and because the values are passed ByRef, they can't be constants
    Dim GIF_EncoderVersion As Long
    GIF_EncoderVersion = [EncoderValueVersionGif89]
    
    Dim gdipColorDepth As Long
    gdipColorDepth = outputColorDepth
    
    Dim TIFF_Compression As Long
    TIFF_Compression = [EncoderValueCompressionLZW]
    
    'TIFF has some unique constraints on account of its many compression schemes.  Because it only supports a subset
    ' of compression types, we must adjust our code accordingly.
    If imgFormat = ImageTIFF Then
    
        Select Case userPreferences.GetPreference_Long("General Preferences", "TIFFCompression", 0)
        
            'Default settings (LZW for > 1bpp, CCITT Group 4 fax encoding for 1bpp)
            Case 0
                If gdipColorDepth = 1 Then TIFF_Compression = [EncoderValueCompressionCCITT4] Else TIFF_Compression = [EncoderValueCompressionLZW]
                
            'No compression
            Case 1
                TIFF_Compression = [EncoderValueCompressionNone]
                
            'Macintosh Packbits (RLE)
            Case 2
                TIFF_Compression = [EncoderValueCompressionRle]
            
            'Proper deflate (Adobe-style) - not supported by GDI+
            Case 3
                TIFF_Compression = [EncoderValueCompressionLZW]
            
            'Obsolete deflate (PKZIP or zLib-style) - not supported by GDI+
            Case 4
                TIFF_Compression = [EncoderValueCompressionLZW]
            
            'LZW
            Case 5
                TIFF_Compression = [EncoderValueCompressionLZW]
                
            'JPEG - not supported by GDI+
            Case 6
                TIFF_Compression = [EncoderValueCompressionLZW]
            
            'Fax Group 3
            Case 7
                gdipColorDepth = 1
                TIFF_Compression = [EncoderValueCompressionCCITT3]
            
            'Fax Group 4
            Case 8
                gdipColorDepth = 1
                TIFF_Compression = [EncoderValueCompressionCCITT4]
                
        End Select
    
    End If
    
    'Request an encoder from GDI+ based on the type passed to this routine
    Dim uEncCLSID As clsid
    Dim uEncParams As EncoderParameters
    Dim aEncParams() As Byte

    Message "Preparing GDI+ encoder for this filetype..."

    Select Case imgFormat
        
        'BMP export
        Case [ImageBMP]
            pvGetEncoderClsID "image/bmp", uEncCLSID
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = EncoderParameterValueTypeLong
                .Guid = pvDEFINE_GUID(EncoderColorDepth)
                .Value = VarPtr(gdipColorDepth)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
    
        'GIF export
        Case [ImageGIF]
            pvGetEncoderClsID "image/gif", uEncCLSID
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = EncoderParameterValueTypeLong
                .Guid = pvDEFINE_GUID(EncoderVersion)
                .Value = VarPtr(GIF_EncoderVersion)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
            
        'JPEG export (requires extra work to specify a quality for the encode)
        Case [ImageJPEG]
            pvGetEncoderClsID "image/jpeg", uEncCLSID
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = [EncoderParameterValueTypeLong]
                .Guid = pvDEFINE_GUID(EncoderQuality)
                .Value = VarPtr(JPEGQuality)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
        
        'PNG export
        Case [ImagePNG]
            pvGetEncoderClsID "image/png", uEncCLSID
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = EncoderParameterValueTypeLong
                .Guid = pvDEFINE_GUID(EncoderColorDepth)
                .Value = VarPtr(gdipColorDepth)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
        
        'TIFF export (requires extra work to specify compression and color depth for the encode)
        Case [ImageTIFF]
            pvGetEncoderClsID "image/tiff", uEncCLSID
            uEncParams.Count = 2
            ReDim aEncParams(1 To Len(uEncParams) + Len(uEncParams.Parameter) * 2)
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = [EncoderParameterValueTypeLong]
                .Guid = pvDEFINE_GUID(EncoderCompression)
                .Value = VarPtr(TIFF_Compression)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = [EncoderParameterValueTypeLong]
                .Guid = pvDEFINE_GUID(EncoderColorDepth)
                .Value = VarPtr(gdipColorDepth)
            End With
            
            CopyMemory aEncParams(Len(uEncParams) + 1), uEncParams.Parameter, Len(uEncParams.Parameter)
    
    End Select

    'With our encoder prepared, we can finally continue with the save
    
    'Check to see if a file already exists at this location
    If FileExist(dstFilename) Then Kill dstFilename
    
    Message "Saving the file..."
    
    'Perform the encode and save
    GDIPlusReturn = GdipSaveImageToFile(hImage, StrConv(dstFilename, vbUnicode), uEncCLSID, aEncParams(1))
    
    If (GDIPlusReturn <> [OK]) Then
        GdipDisposeImage hImage
        GDIPlusSavePicture = False
        Exit Function
    End If
    
    Message "Releasing all temporary image copies..."
    
    'Release the GDI+ copy of the image
    GDIPlusReturn = GdipDisposeImage(hImage)
    
    Message "Save complete."
    
    GDIPlusSavePicture = True
    Exit Function
    
GDIPlusSaveError:

    GDIPlusSavePicture = False
    
End Function

'At start-up, this function is called to determine whether or not we have GDI+ available on this machine.
Public Function isGDIPlusAvailable() As Boolean

    Dim gdiCheck As GDIPlusStartupInput
    gdiCheck.GDIPlusVersion = 1
    
    If (GdiplusStartup(GDIPlusToken, gdiCheck) <> [OK]) Then
        isGDIPlusAvailable = False
    Else
        isGDIPlusAvailable = True
    End If

End Function

'At shutdown, this function must be called to release our GDI+ instance
Public Function releaseGDIPlus()
    GdiplusShutdown GDIPlusToken
End Function

'Thanks to Carles P.V. for providing the following four functions, which are used as part of GDI+ image saving.
' You can download Carles's full project from http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
Private Function pvGetEncoderClsID(strMimeType As String, ClassID As clsid) As Long

  Dim Num      As Long
  Dim Size     As Long
  Dim lIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    
    pvGetEncoderClsID = -1 ' Failure flag
    
    '-- Get the encoder array size
    Call GdipGetImageEncodersSize(Num, Size)
    If (Size = 0) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageEncoders(Num, Size, Buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    
    '-- Loop through all the codecs
    For lIdx = 1 To Num
        '-- Must convert the pointer into a usable string
        If (StrComp(pvPtrToStrW(ICI(lIdx).MimeType), strMimeType, vbTextCompare) = 0) Then
            ClassID = ICI(lIdx).ClassID ' Save the Class ID
            pvGetEncoderClsID = lIdx      ' Return the index number for success
            Exit For
        End If
    Next lIdx
    '-- Free the memory
    Erase ICI
    Erase Buffer
End Function

Private Function pvDEFINE_GUID(ByVal sGuid As String) As clsid
'-- Courtesy of: Dana Seaman
'   Helper routine to convert a CLSID(aka GUID) string to a structure
'   Example ImageFormatBMP = {B96B3CAB-0728-11D3-9D7B-0000F81EF32E}
    Call CLSIDFromString(StrPtr(sGuid), pvDEFINE_GUID)
End Function

'"Convert" (technically, dereference) an ANSI or Unicode string to the BSTR used by VB
Private Function pvPtrToStrW(ByVal lpsz As Long) As String
    
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        pvPtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
End Function

'Same as above, but in reverse
Private Function pvPtrToStrA(ByVal lpsz As Long) As String
    
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenA(lpsz)

    If (lLen > 0) Then
        sOut = String$(lLen, vbNullChar)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen)
        pvPtrToStrA = sOut
    End If
End Function

'Implementation of an IStream-compatible interface.  Originally accessed from http://read.pudn.com/downloads151/sourcecode/graph/texture_mapping/657997/%E9%80%8F%E6%98%8E%E5%8A%A8%E6%80%81%E6%97%B6%E9%92%9F/Classes/cGDIPlus.cls__.htm
Private Function CreateStream(byteContent() As Byte, Optional byteOffset As Long = 0&) As IUnknown
     
    ' Purpose: Create an IStream-compatible IUnknown interface containing the
    ' passed byte aray. This IUnknown interface can be passed to GDI+ functions
    ' that expect an IStream interface -- neat hack
     
    On Error GoTo HandleError
    Dim o_lngByteCount  As Long
    Dim o_hMem As Long
    Dim o_lpMem  As Long
      
    If iparseIsArrayEmpty(VarPtrArray(byteContent)) = 0& Then ' create a growing stream as needed
         Call CreateStreamOnHGlobal(0, 1, CreateStream)
    Else                                        ' create a fixed stream
         o_lngByteCount = UBound(byteContent) - byteOffset + 1
         o_hMem = GlobalAlloc(&H2&, o_lngByteCount)
         If o_hMem <> 0 Then
             o_lpMem = GlobalLock(o_hMem)
             If o_lpMem <> 0 Then
                 CopyMemory ByVal o_lpMem, byteContent(byteOffset), o_lngByteCount
                 Call GlobalUnlock(o_hMem)
                 Call CreateStreamOnHGlobal(o_hMem, 1, CreateStream)
             End If
         End If
     End If
     
HandleError:
End Function

'Test if an array has been initialized
Private Function iparseIsArrayEmpty(FarPointer As Long) As Long
    CopyMemory iparseIsArrayEmpty, ByVal FarPointer, 4&
End Function

'Save an image to a PNG stream using GDI+.  Per the current save spec, ImageID must be specified.
Public Function GDIPlusSavePNGStream(ByVal imageID As Long, ByRef outStream() As Byte, ByRef IIStream As IUnknown) As Boolean

    'Message "Initializing GDI+..."

    'Begin by creating a generic bitmap header for the current layer
    Dim imgHeader As BITMAPINFO
    
    With imgHeader.Header
        .Size = Len(imgHeader.Header)
        .Planes = 1
        .BitCount = pdImages(imageID).mainLayer.getLayerColorDepth
        .Width = pdImages(imageID).mainLayer.getLayerWidth
        .Height = -pdImages(imageID).mainLayer.getLayerHeight
    End With

    'Use GDI+ to create a GDI+-compatible bitmap
    Dim GDIPlusReturn As Long
    Dim hImage As Long
    
    'Message "Creating GDI+ compatible image copy..."
    
    GDIPlusReturn = GdipCreateBitmapFromGdiDib(imgHeader, ByVal pdImages(imageID).mainLayer.getLayerDIBits, hImage)
    
    If GDIPlusReturn <> 0 Then Exit Function
    
    'PNG requires extra parameters, and because the values are passed ByRef, they can't be constants
    Dim PNG_ColorDepth As Long
    PNG_ColorDepth = pdImages(imageID).mainLayer.getLayerColorDepth
    
    'Request an encoder from GDI+ based on the type passed to this routine
    Dim uEncCLSID As clsid
    Dim uEncParams As EncoderParameters
    Dim aEncParams() As Byte

    'Message "Preparing GDI+ encoder for this filetype..."
    
    'PNG export
            pvGetEncoderClsID "image/png", uEncCLSID
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = EncoderParameterValueTypeLong
                .Guid = pvDEFINE_GUID(EncoderColorDepth)
                .Value = VarPtr(PNG_ColorDepth)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
        
    'With our encoder prepared, we can finally continue with the save
        
    'Message "Saving the file to memory stream..."
    
    'Perform the encode and save
    
    'First, create a null stream (IUnknown object)
    Erase outStream
    Set IIStream = CreateStream(outStream)
    
    GDIPlusReturn = GdipSaveImageToStream(hImage, IIStream, uEncCLSID, aEncParams(1))
    
    If GDIPlusReturn <> 0 Then Exit Function
    
    'Message "Releasing all temporary image copies..."
    
    'Release the GDI+ copy of the image
    GDIPlusReturn = GdipDisposeImage(hImage)
    
    If GDIPlusReturn <> 0 Then Exit Function
    
    'Message "Copy complete."
    
    GDIPlusSavePNGStream = True

End Function
