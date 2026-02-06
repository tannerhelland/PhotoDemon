Attribute VB_Name = "WIC"
'***************************************************************************
'WIC (Windows Imaging Component) Interface
'Copyright 2019-2026 by Tanner Helland
'Created: 26/December/19
'Last updated: 05/August/20
'Last update: further work on this module has been postponed until 8.0's release.  I've commented out
'             a bunch of unused structs and APIs until they are actually used.
'
'From MSDN (https://docs.microsoft.com/en-us/windows/win32/wic/-wic-lh):
' "The Windows Imaging Component (WIC) is an extensible platform that provides [a]
'  low-level API for digital images."
'
'At present, PhotoDemon makes limited use of WIC for importing esoteric file formats (e.g. HEIC/HEIF).
' Usage may expand in the future as needs require.
'
'Thank you to vbforums user Victor Bravo VI for first documenting lightweight WIC usage from VB6:
' http://www.vbforums.com/showthread.php?879695-Anyone-familiar-with-Windows-Imaging-Component-or-odd-IDL-pointer-types-in-general&p=5427029&viewfull=1#post5427029
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

''IStream declares; these are used indirectly when exporting files
'Private Enum WIN32_STGM
'
'    'Access
'    STGM_READ = &H0&
'    STGM_WRITE = &H1&
'    STGM_READWRITE = &H2&
'
'    'Sharing
'    STGM_SHARE_DENY_NONE = &H40&
'    STGM_SHARE_DENY_READ = &H30&
'    STGM_SHARE_DENY_WRITE = &H20&
'    STGM_SHARE_EXCLUSIVE = &H10&
'    STGM_PRIORITY = &H40000
'
'    'Creation
'    STGM_CREATE = &H1000&
'    STGM_CONVERT = &H20000
'    STGM_FAILIFTHERE = &H0&
'
'    'Transactioning
'    STGM_DIRECT = &H0&
'    STGM_TRANSACTED = &H10000
'
'    'Transactioning Performance
'    STGM_NOSCRATCH = &H100000
'    STGM_NOSNAPSHOT = &H200000
'
'    'Direct SWMR and Simple
'    STGM_SIMPLE = &H8000000
'    STGM_DIRECT_SWMR = &H400000
'
'    'Delete On Release
'    STGM_DELETEONRELEASE = &H4000000
'
'End Enum
'
'#If False Then
'    Private Const STGM_READ = &H0&, STGM_WRITE = &H1&, STGM_READWRITE = &H2&, STGM_SHARE_DENY_NONE = &H40&, STGM_SHARE_DENY_READ = &H30&, STGM_SHARE_DENY_WRITE = &H20&, STGM_SHARE_EXCLUSIVE = &H10&, STGM_PRIORITY = &H40000, STGM_CREATE = &H1000&, STGM_CONVERT = &H20000, STGM_FAILIFTHERE = &H0&, STGM_DIRECT = &H0&, STGM_TRANSACTED = &H10000, STGM_NOSCRATCH = &H100000, STGM_NOSNAPSHOT = &H200000, STGM_SIMPLE = &H8000000, STGM_DIRECT_SWMR = &H400000, STGM_DELETEONRELEASE = &H4000000
'#End If

'WIC constants are declared in wincodec.h
Private Const WINCODEC_SDK_VERSION1 As Long = &H236&
Private Const WINCODEC_SDK_VERSION2 As Long = &H237&

'COM HRESULT for "success"; this is used to check all WIC/IWIC returns for pass/fail.
' (Note that HRESULTs use bitfields, so they're not easily coerced into a VB6 enum - non-OK returns
' need to be dealt with manually.)
Private Const S_OK As Long = 0&

'WIC enums
'Private Enum WICColorContextType
'    WICColorContextUninitialized = &H0&
'    WICColorContextProfile = &H1&
'    WICColorContextExifColorSpace = &H2&
'End Enum
'
'#If False Then
'    Private Const WICColorContextUninitialized = &H0&, WICColorContextProfile = &H1&, WICColorContextExifColorSpace = &H2&
'#End If

Private Enum WICDecodeOptions
    WICDecodeMetadataCacheOnDemand = &H0&
    WICDecodeMetadataCacheOnLoad = &H1&
    WICMETADATACACHEOPTION_FORCE_DWORD = &H7FFFFFFF
End Enum

#If False Then
    Private Const WICDecodeMetadataCacheOnDemand = &H0&, WICDecodeMetadataCacheOnLoad = &H1&, WICMETADATACACHEOPTION_FORCE_DWORD = &H7FFFFFFF
#End If

Private Enum WICBitmapDitherType
    WICBitmapDitherTypeNone = 0&
    WICBitmapDitherTypeSolid = 0&
    WICBitmapDitherTypeOrdered4x4 = &H1&
    WICBitmapDitherTypeOrdered8x8 = &H2&
    WICBitmapDitherTypeOrdered16x16 = &H3&
    WICBitmapDitherTypeSpiral4x4 = &H4&
    WICBitmapDitherTypeSpiral8x8 = &H5&
    WICBitmapDitherTypeDualSpiral4x4 = &H6&
    WICBitmapDitherTypeDualSpiral8x8 = &H7&
    WICBitmapDitherTypeErrorDiffusion = &H8&
End Enum

#If False Then
    Private Const WICBitmapDitherTypeNone = 0&, WICBitmapDitherTypeSolid = 0&, WICBitmapDitherTypeOrdered4x4 = &H1&, WICBitmapDitherTypeOrdered8x8 = &H2&, WICBitmapDitherTypeOrdered16x16 = &H3&, WICBitmapDitherTypeSpiral4x4 = &H4&, WICBitmapDitherTypeSpiral8x8 = &H5&, WICBitmapDitherTypeDualSpiral4x4 = &H6&, WICBitmapDitherTypeDualSpiral8x8 = &H7&, WICBitmapDitherTypeErrorDiffusion = &H8&
#End If

'Private Enum WICBitmapEncoderCacheOption
'    WICBitmapEncoderCacheInMemory = 0&      'As of Jan 2020, "not supported"
'    WICBitmapEncoderCacheTempFile = &H1&    'As of Jan 2020, "not supported"
'    WICBitmapEncoderNoCache = &H2&
'End Enum
'
'#If False Then
'    Private Const WICBitmapEncoderCacheInMemory = 0&, WICBitmapEncoderCacheTempFile = &H1&, WICBitmapEncoderNoCache = &H2&
'#End If

Private Enum WICBitmapInterpolationMode
    WICBitmapInterpolationModeNearestNeighbor = 0&
    WICBitmapInterpolationModeLinear = &H1&
    WICBitmapInterpolationModeCubic = &H2&
    WICBitmapInterpolationModeFant = &H3&
    WICBitmapInterpolationModeHighQualityCubic = &H4&   'Win 10 only!
End Enum

#If False Then
    Private Const WICBitmapInterpolationModeNearestNeighbor = 0&, WICBitmapInterpolationModeLinear = &H1&, WICBitmapInterpolationModeCubic = &H2&, WICBitmapInterpolationModeFant = &H3&, WICBitmapInterpolationModeHighQualityCubic = &H4&
#End If

Private Enum WICBitmapPaletteType
    WICBitmapPaletteTypeCustom = 0&
    WICBitmapPaletteTypeMedianCut = &H1&
    WICBitmapPaletteTypeFixedBW = &H2&
    WICBitmapPaletteTypeFixedHalftone8 = &H3&
    WICBitmapPaletteTypeFixedHalftone27 = &H4&
    WICBitmapPaletteTypeFixedHalftone64 = &H5&
    WICBitmapPaletteTypeFixedHalftone125 = &H6&
    WICBitmapPaletteTypeFixedHalftone216 = &H7&
    WICBitmapPaletteTypeFixedWebPalette = WICBitmapPaletteTypeFixedHalftone216
    WICBitmapPaletteTypeFixedHalftone252 = &H8&
    WICBitmapPaletteTypeFixedHalftone256 = &H9&
    WICBitmapPaletteTypeFixedGray4 = &HA&
    WICBitmapPaletteTypeFixedGray16 = &HB&
    WICBitmapPaletteTypeFixedGray256 = &HC&
End Enum

#If False Then
    Private Const WICBitmapPaletteTypeCustom = 0&, WICBitmapPaletteTypeMedianCut = &H1&, WICBitmapPaletteTypeFixedBW = &H2&, WICBitmapPaletteTypeFixedHalftone8 = &H3&, WICBitmapPaletteTypeFixedHalftone27 = &H4&, WICBitmapPaletteTypeFixedHalftone64 = &H5&, WICBitmapPaletteTypeFixedHalftone125 = &H6&, WICBitmapPaletteTypeFixedHalftone216 = &H7&, WICBitmapPaletteTypeFixedWebPalette = WICBitmapPaletteTypeFixedHalftone216, WICBitmapPaletteTypeFixedHalftone252 = &H8&, WICBitmapPaletteTypeFixedHalftone256 = &H9&, WICBitmapPaletteTypeFixedGray4 = &HA&, WICBitmapPaletteTypeFixedGray16 = &HB&, WICBitmapPaletteTypeFixedGray256 = &HC&
#End If

'Stream helpers
'Private Declare Function SHCreateStreamOnFileEx Lib "shlwapi" (ByVal pszFile As Long, ByVal grfMode As WIN32_STGM, ByVal dwAttributes As Long, ByVal fCreate As Long, ByVal pstmTemplate As stdole.IUnknown, ByRef ppstm As stdole.IUnknown) As Long

'The WIC flat API is documented here: https://docs.microsoft.com/en-us/windows/win32/wic/wic-proxy-functions
'Private Declare Function IWICBitmapDecoder_GetColorContexts_Proxy Lib "windowscodecs" (ByVal ptrToBitmapDecoder As stdole.IUnknown, ByVal cCount As Long, ByVal ptrToIWICColorContext As Long, ByRef pcActualCount As Long) As Long
Private Declare Function IWICBitmapDecoder_GetFrameCount_Proxy Lib "windowscodecs" (ByVal ptrToObject As stdole.IUnknown, ByRef pCount As Long) As Long
Private Declare Function IWICBitmapDecoder_GetFrame_Proxy Lib "windowscodecs" (ByVal ptrToObject As stdole.IUnknown, ByVal frameIndex As Long, ByRef ppIBitmapFrame As stdole.IUnknown) As Long
'Private Declare Function IWICBitmapEncoder_Initialize_Proxy Lib "windowscodecs" (ByVal ptrToObject As stdole.IUnknown, ByRef dstStream As stdole.IUnknown, ByVal cacheOption As WICBitmapEncoderCacheOption) As Long
Private Declare Function IWICBitmapScaler_Initialize_Proxy Lib "windowscodecs" (ByVal ptrToObject As stdole.IUnknown, ByVal pISource As stdole.IUnknown, ByVal uiWidth As Long, ByVal uiHeight As Long, ByVal scalerMode As WICBitmapInterpolationMode) As Long
Private Declare Function IWICBitmapSource_CopyPixels_Proxy Lib "windowscodecs" (ByVal ptrToObject As stdole.IUnknown, ByVal ptrToSrcRectL As Long, ByVal cbStride As Long, ByVal cbBufferSize As Long, ByVal pbBuffer As Long) As Long
Private Declare Function IWICBitmapSource_GetSize_Proxy Lib "windowscodecs" (ByVal ptrToObject As stdole.IUnknown, ByRef puiWidth As Long, ByRef puiHeight As Long) As Long
Private Declare Function IWICFormatConverter_Initialize_Proxy Lib "windowscodecs" (ByVal ptrToObject As stdole.IUnknown, ByVal pISource As stdole.IUnknown, ByRef dstFormat As Guid, ByVal ditherType As WICBitmapDitherType, ByVal pIPalette As stdole.IUnknown, ByVal alphaThresholdPercent As Double, ByVal paletteTranslate As WICBitmapPaletteType) As Long
Private Declare Function IWICImagingFactory_CreateBitmapScaler_Proxy Lib "windowscodecs" (ByVal pFactory As stdole.IUnknown, ByRef ppIBitmapScaler As stdole.IUnknown) As Long
Private Declare Function IWICImagingFactory_CreateDecoderFromFilename_Proxy Lib "windowscodecs" (ByVal pFactory As stdole.IUnknown, ByVal wzFilename As Long, ByRef pguidVendor As Guid, ByVal dwDesiredAccess As Long, ByVal metadataOptions As WICDecodeOptions, ByRef ppIDecoder As stdole.IUnknown) As Long
'Private Declare Function IWICImagingFactory_CreateEncoder_Proxy Lib "windowscodecs" (ByVal ptrToObject As stdole.IUnknown, ByRef guidContainerFormat As Guid, ByRef guidVendor As Guid, ByRef dstIWICBitmapEncoder As stdole.IUnknown) As Long
Private Declare Function IWICImagingFactory_CreateFormatConverter_Proxy Lib "windowscodecs" (ByVal pFactory As stdole.IUnknown, ByRef ppIFormatConverter As stdole.IUnknown) As Long
Private Declare Function WICCreateImagingFactory_Proxy Lib "windowscodecs" (ByVal sdkVersion As Long, ByRef ppIImagingFactory As stdole.IUnknown) As Long
'Private Declare Function WICCreateColorContext_Proxy Lib "windowscodecs" (ByVal ptrToImagingFactory As stdole.IUnknown, ByRef dstWICColorContext As stdole.IUnknown) As Long

'Test lib availability
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long

'If WIC is available, we cache status for the rest of the session.
' (Note that the lib is self-freeing, so we don't need to manually un-initialize it at shutdown.)
Private m_IsLibAvailable As Boolean

'A WIC factory only needs to be created once; it can be reused on subsequent calls
Private m_WICImagingFactory As stdole.IUnknown

'At present, PD only uses WIC to load HEIC/HEIF images.  HEIF support is complicated because depending on
' your Win 10 version, you may need to perform 1-2 extra downloads from the MS Store as described in this link:
' https://www.windowscentral.com/how-open-heic-and-hevc-files-windows-10s-photos-app
'
'I have not yet devised a good way to explain this to the user... but perhaps a pop-up could be used in
' the future if the user attempts a HEIF load but WIC fails with "file type unknown".
Public Function IsWICAvailable() As Boolean
    
    If (Not m_IsLibAvailable) Then
        
        Dim hMod As Long, libName As String
        libName = "windowscodecs.dll"
        hMod = LoadLibrary(StrPtr(libName))
        m_IsLibAvailable = (hMod <> 0)
        
        If m_IsLibAvailable Then
            Dim testAddress As Long
            testAddress = GetProcAddress(hMod, "WICCreateImagingFactory_Proxy")
            m_IsLibAvailable = (testAddress <> 0)
            FreeLibrary hMod
        End If
        
    End If
    
    IsWICAvailable = m_IsLibAvailable
    
End Function

'Load an arbitrary path to an arbitrary pdDIB object.
' Based heavily on work originally done by Victor Bravo VI (http://www.vbforums.com/showthread.php?879695-Anyone-familiar-with-Windows-Imaging-Component-or-odd-IDL-pointer-types-in-general&p=5427029&viewfull=1#post5427029)
Public Function LoadFileToDIB(ByRef dstDIB As pdDIB, ByRef srcFile As String) As Boolean
    
    LoadFileToDIB = False
    
    'Attempt to initialize an imaging factory
    If (Not StartWICImagingFactory()) Then Exit Function
    
    'Next, we need a decoder; this can be auto-generated using CreateDecoderFromFilename
    Const GENERIC_READ As Long = &H80000000
    Dim iWICBitmapDecoder As stdole.IUnknown
    If (IWICImagingFactory_CreateDecoderFromFilename_Proxy(m_WICImagingFactory, StrPtr(srcFile), GUID_NULL, GENERIC_READ, WICDecodeMetadataCacheOnDemand, iWICBitmapDecoder) <> S_OK) Then Exit Function
    
    'One of the big improvements made in WIC (vs GDI+) is support for multipage/frame image formats.
    ' Generally speaking, we just want to retrieve the first page, but this function is already set up
    ' to deal with multipage formats (if desired in the future).
    
    'Start by retrieving a page/frame count for this file
    Dim frameCount As Long
    If (IWICBitmapDecoder_GetFrameCount_Proxy(iWICBitmapDecoder, frameCount) <> S_OK) Then Exit Function
    
    'For now, use the decoder to retrieve just the first frame
    Dim curFrame As Long: curFrame = 0
    Dim iWICBitmapFrameDecode As stdole.IUnknown
    If (IWICBitmapDecoder_GetFrame_Proxy(iWICBitmapDecoder, curFrame, iWICBitmapFrameDecode) <> S_OK) Then Exit Function
    
    'Relevant frame data can now be queried.
    
    'Get width/height of the page
    Dim frameWidth As Long, frameHeight As Long
    If (IWICBitmapSource_GetSize_Proxy(iWICBitmapFrameDecode, frameWidth, frameHeight) <> S_OK) Then Exit Function
    
    'ICC profiles are not supported by WIC stubs; instead, we use ExifTool to retrieve the color profile (if any).
    ' In the future, we could perhaps use a TLB to retrieve color context(s) here.
    
    'PD operates solely in premultipled RGBA color space(s); prep a relevant converter
    Dim iWICFormatConverter As stdole.IUnknown
    If (IWICImagingFactory_CreateFormatConverter_Proxy(m_WICImagingFactory, iWICFormatConverter) <> S_OK) Then Exit Function
    If (IWICFormatConverter_Initialize_Proxy(iWICFormatConverter, iWICBitmapFrameDecode, GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, Nothing, 0#, WICBitmapPaletteTypeCustom) <> S_OK) Then Exit Function
    
    'Prep the destination DIB
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateBlank frameWidth, frameHeight, 32, 0, 0
    
    'WIC pixel copiers require a BitmapScaler object (even when you aren't resizing the image);
    ' prep a dummy one now, using an interpolation setting safe for all Windows versions.
    ' (High-quality bicubic is Win10 only.)
    Dim iWICBitmapScaler As stdole.IUnknown
    If (IWICImagingFactory_CreateBitmapScaler_Proxy(m_WICImagingFactory, iWICBitmapScaler) <> S_OK) Then Exit Function
    If (IWICBitmapScaler_Initialize_Proxy(iWICBitmapScaler, iWICFormatConverter, frameWidth, frameHeight, WICBitmapInterpolationModeFant) <> S_OK) Then Exit Function
    
    'Copy the pixels into place
    LoadFileToDIB = (IWICBitmapSource_CopyPixels_Proxy(iWICBitmapScaler, 0&, dstDIB.GetDIBStride, dstDIB.GetDIBStride * dstDIB.GetDIBHeight, dstDIB.GetDIBPointer) = S_OK)
    
    'In PD, RGBA pixels are always premultiplied; set the corresponding flag before exiting
    If LoadFileToDIB Then dstDIB.SetInitialAlphaPremultiplicationState True
    
End Function

'Under construction - do not use yet!
'Public Function SaveImageToFile_HEIF(ByRef srcImage As pdImage, ByRef dstFilename As String, Optional ByVal imgFileType As PD_IMAGE_FORMAT = PDIF_HEIF) As Boolean
'
'    'Attempt to initialize an imaging factory
'    If (Not StartWICImagingFactory()) Then Exit Function
'
'    'Next, we need an encoder; these are obviously format-specific
'    Dim iWICBitmapEncoder As stdole.IUnknown
'    If (IWICImagingFactory_CreateEncoder_Proxy(m_WICImagingFactory, GUID_ContainerFormatHeif, GUID_NULL, iWICBitmapEncoder) <> S_OK) Then Exit Function
'
'    'Next, we need to initialize the encoder against a stream.
'
'    'Start by creating the stream
'    Dim tmp_StreamTemplate As stdole.IUnknown, dstStream As stdole.IUnknown
'    If (SHCreateStreamOnFileEx(StrPtr(dstFilename), STGM_READWRITE Or STGM_SHARE_DENY_NONE Or STGM_CREATE, &H80&, &H1&, tmp_StreamTemplate, dstStream) <> S_OK) Then Exit Function
'
'    'Next, initialize the previously created encoder
'    If (IWICBitmapEncoder_Initialize_Proxy(iWICBitmapEncoder, dstStream, WICBitmapEncoderNoCache) <> S_OK) Then Exit Function
'
'
'End Function

'Convenience GUID functions.  Original declares are in wincodec.h
Private Function GUID_WICPixelFormat32bppPBGRA() As Guid
    DEFINE_GUID GUID_WICPixelFormat32bppPBGRA, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H10
End Function

Private Function GUID_WICHeifEncoder() As Guid
    DEFINE_GUID GUID_WICHeifEncoder, &HDBECEC1, &H9EB3, &H4860, &H9C, &H6F, &HDD, &HBE, &H86, &H63, &H45, &H75
End Function

Private Function GUID_ContainerFormatHeif() As Guid
    DEFINE_GUID GUID_ContainerFormatHeif, &HE1E62521, &H6787, &H405B, &HA3, &H39, &H50, &H7, &H15, &HB5, &H76, &H3F
End Function

'VB6 hacks for standard DEFINE_GUID macros; thank you to Victor Bravo VI for the original versions (http://www.vbforums.com/showthread.php?879695-Anyone-familiar-with-Windows-Imaging-Component-or-odd-IDL-pointer-types-in-general&highlight=Windows+Imaging+Component)
' MSDN: https://docs.microsoft.com/en-us/windows-hardware/drivers/kernel/defining-and-exporting-new-guids
Private Function GUID_NULL() As Guid

End Function

Private Sub DEFINE_GUID(ByRef u As Guid, ByVal d1 As Long, ByVal d2 As Integer, ByVal d3 As Integer, ByVal d4_0 As Byte, ByVal d4_1 As Byte, ByVal d4_2 As Byte, ByVal d4_3 As Byte, ByVal d4_4 As Byte, ByVal d4_5 As Byte, ByVal d4_6 As Byte, ByVal d4_7 As Byte)
    u.Data1 = d1
    u.Data2 = d2: u.Data3 = d3
    u.Data4(0) = d4_0: u.Data4(1) = d4_1: u.Data4(2) = d4_2: u.Data4(3) = d4_3: u.Data4(4) = d4_4: u.Data4(5) = d4_5: u.Data4(6) = d4_6: u.Data4(7) = d4_7
End Sub

'Returns TRUE if a WIC imaging factory is available; FALSE otherwise
Private Function StartWICImagingFactory() As Boolean
    
    On Error GoTo FactoryProblem
    
    'Attempt to create a factory object (starting with newest version first)
    If (m_WICImagingFactory Is Nothing) Then
        If (WICCreateImagingFactory_Proxy(WINCODEC_SDK_VERSION2, m_WICImagingFactory) <> S_OK) Then
            If (WICCreateImagingFactory_Proxy(WINCODEC_SDK_VERSION1, m_WICImagingFactory) <> S_OK) Then
                StartWICImagingFactory = False
                Exit Function
            End If
        End If
    End If
    
FactoryProblem:
    StartWICImagingFactory = Not (m_WICImagingFactory Is Nothing)
    
End Function
