Attribute VB_Name = "ImageFormats"
'***************************************************************************
'PhotoDemon Image Format Manager
'Copyright 2012-2026 by Tanner Helland
'Created: 18/November/12
'Last updated: 29/January/26
'Last update: HDR import now handled by native VB6 implementation (not FreeImage)
'
'This module determines run-time read/write support for various image formats.
'
'Based on available plugins, this class generates a list of file formats that PhotoDemon is capable
' of importing and exporting.  Import/export lists are separately maintained, and the presence of a
' format in the Import category does not guarantee a similar presence in the Export category.
'
'Some esoteric formats rely on external libraries for loading and/or saving.  Similarly, some formats
' support multiple import engines (e.g. PNGs are preferentially loaded by an internal PNG decoder,
' but we could theoretically hand them off to other libraries).  From this module alone, it won't be
' clear which engine or third-party library (if any) is used to load a given format - for that,
' consult the relevant debug log ([PD path]/Data/Debug) after loading an image file.
'
'Note also that as of 2020, many formats use native PhotoDemon-specific encoder/decoder classes.
' These formats are *always* available regardless of 3rd-party library status, but some formats
' may have add-on features that require third-party libraries - for example, PSD files are
' typically RLE (PackBits) encoded, which PhotoDemon can decode natively, but HDR PSD images can
' optionally use DEFLATE compression which requires PD to tap into a 3rd-party library (libdeflate).
' These complexities are all managed silently, but if you want to use image support code in another
' project, you'll need to dive deeper than just this module.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Is the FreeImage DLL available?  We store this value here because we fall back to FreeImage
' for images that don't load via traditional means (which provides last-ditch coverage for
' esoteric and/or legacy formats).
Private m_FreeImageEnabled As Boolean

'Number of available input, output formats
Private m_numOfInputFormats As Long, m_numOfOutputFormats As Long

'Array of available input, output extensions.
Private inputExtensions() As String, outputExtensions() As String

'Array of "friendly" descriptions for input, output formats
Private inputDescriptions() As String, outputDescriptions() As String

'Array of corresponding image format constants
Private inputPDIFs() As PD_IMAGE_FORMAT, outputPDIFs() As PD_IMAGE_FORMAT

'Array of common-dialog-formatted input/output filetypes.  (Common dialogs require different pipe-based formatting
' than normal lists, as you must add human-readable text descriptions to the list.)
Private m_CommonDialogInputs As String, m_CommonDialogOutputs As String

'Common dialog also require a specialized "default extension" string for output files
Private m_cdOutputDefaultExtensions As String

'When analyzing our current plugin options, we have to make a lot of on-the-fly decisions about format availability.
' This value is shared among multiple functions while making such decisions.  Do not treat it as a meaningful value.
Private m_curFormatIndex As Long

'Return the index of given input PDIF
Public Function GetIndexOfInputPDIF(ByVal srcFIF As PD_IMAGE_FORMAT) As Long
    
    Dim i As Long
    For i = 0 To GetNumOfInputFormats
        If inputPDIFs(i) = srcFIF Then
            GetIndexOfInputPDIF = i
            Exit Function
        End If
    Next i
    
    'If we reach this line of code, no match was found.  Return -1.
    GetIndexOfInputPDIF = -1
    
End Function

'Return the PDIF ("PD image format" constant) at a given index
Public Function GetInputPDIF(ByVal dIndex As Long) As Long
    If (dIndex >= 0) And (dIndex <= m_numOfInputFormats) Then
        GetInputPDIF = inputPDIFs(dIndex)
    Else
        GetInputPDIF = PDIF_UNKNOWN
    End If
End Function

'Return the friendly input format description at a given index
Public Function GetInputFormatDescription(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= m_numOfInputFormats) Then
        GetInputFormatDescription = inputDescriptions(dIndex)
    Else
        GetInputFormatDescription = vbNullString
    End If
End Function

'Return the input format extension at a given index
Public Function GetInputFormatExtensions(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= m_numOfInputFormats) Then
        GetInputFormatExtensions = inputExtensions(dIndex)
    Else
        GetInputFormatExtensions = vbNullString
    End If
End Function

'Return a list of all input formats supported by the current session.  By default, the list is delimited with commas,
' and each extension is listed as "*.abc".
Public Function GetListOfInputFormats(Optional ByVal listDelimiter As String = ";", Optional ByVal includeStarDot As Boolean = True) As String
    
    'The first entry in the extensions collection is used for the "All supported formats" common dialog option;
    ' as such, it already contains a full list of valid extensions.
    GetListOfInputFormats = inputExtensions(0)
    
    If Strings.StringsNotEqual(listDelimiter, ";", False) Then GetListOfInputFormats = Replace$(GetListOfInputFormats, ";", listDelimiter)
    If (Not includeStarDot) Then GetListOfInputFormats = Replace$(GetListOfInputFormats, "*.", vbNullString)
    
End Function

'Return the number of available input format types
Public Function GetNumOfInputFormats() As Long
    GetNumOfInputFormats = m_numOfInputFormats
End Function

'Return a list of input filetypes formatted for use with a common dialog box
Public Function GetCommonDialogInputFormats() As String
    GetCommonDialogInputFormats = m_CommonDialogInputs
End Function

'Return the index of given output FIF.  Returns -1 if PD cannot export a given format.
Public Function GetIndexOfOutputPDIF(ByVal srcFIF As PD_IMAGE_FORMAT) As Long
    
    GetIndexOfOutputPDIF = -1
    
    Dim i As Long
    For i = 0 To GetNumOfOutputFormats
        If (outputPDIFs(i) = srcFIF) Then
            GetIndexOfOutputPDIF = i
            Exit For
        End If
    Next i
    
End Function

'Return the FIF (image format constant) at a given index
Public Function GetOutputPDIF(ByVal dIndex As Long) As PD_IMAGE_FORMAT
    If (dIndex >= 0) And (dIndex <= m_numOfOutputFormats) Then
        GetOutputPDIF = outputPDIFs(dIndex)
    Else
        GetOutputPDIF = PDIF_UNKNOWN
    End If
End Function

'Return the friendly output format description at a given index
Public Function GetOutputFormatDescription(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= m_numOfOutputFormats) Then
        GetOutputFormatDescription = outputDescriptions(dIndex)
    Else
        GetOutputFormatDescription = vbNullString
    End If
End Function

'Return the output format extension at a given index
Public Function GetOutputFormatExtension(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= m_numOfOutputFormats) Then
        GetOutputFormatExtension = outputExtensions(dIndex)
    Else
        GetOutputFormatExtension = vbNullString
    End If
End Function

'Return the number of available output format types
Public Function GetNumOfOutputFormats() As Long
    GetNumOfOutputFormats = m_numOfOutputFormats
End Function

'Return a list of output filetypes formatted for use with a common dialog box
Public Function GetCommonDialogOutputFormats() As String
    GetCommonDialogOutputFormats = m_CommonDialogOutputs
End Function

'Return a list of output default extensions formatted for use with a common dialog box
Public Function GetCommonDialogDefaultExtensions() As String
    GetCommonDialogDefaultExtensions = m_cdOutputDefaultExtensions
End Function

'Generate a list of available import formats
Public Sub GenerateInputFormats()

    'Prepare a list of possible INPUT formats based on the 3rd-party libraries available at run-time.
    
    '(These format lists are automatically trimmed after library status is evaluated.
    ' The arbitrary upper limit only needs to be revisited if I greatly expand format support in the future!)
    Const MAX_NUM_INPUT_FORMATS As Long = 75
    ReDim inputExtensions(0 To MAX_NUM_INPUT_FORMATS - 1) As String
    ReDim inputDescriptions(0 To MAX_NUM_INPUT_FORMATS - 1) As String
    ReDim inputPDIFs(0 To MAX_NUM_INPUT_FORMATS - 1) As PD_IMAGE_FORMAT
    
    'Formats should be added in alphabetical order, as this class has no "sort" functionality.

    'Always start with an "All Compatible Images" option
    inputDescriptions(0) = g_Language.TranslateMessage("All Compatible Images")
    
    'Unique to that first "All Compatible Images" entry is a matching list of compatible extensions.
    ' We don't need to supply it; this class automatically generates a matching list as we go.
    
    'Set the location tracker to "0".  Beyond this point, it will be automatically updated.
    m_curFormatIndex = 0
    
    'Next, add individual formats.  Some formats are condidtional on third-party libraries;
    ' others use internal or system-provided encoders that will always be available.
    
    'AV1-based codecs require an external plugin, but PD can download it automatically (with permission)
    ' if the current system is compatible, even if the plugin isn't installed at startup.
    If OS.OSSupports64bitExe() Then
        AddInputFormat "AVIF - AV1 Image File", "*.heif;*.heifs;*.heic;*.heics;*.avci;*.avcs;*.avif;*.avifs", PDIF_AVIF
    End If
    
    AddInputFormat "BMP - Windows or OS/2 Bitmap", "*.bmp", PDIF_BMP
    AddInputFormat "CBZ - Comic Book Archive", "*.cbz", PDIF_CBZ
    
    'FreeImage will only be used on 32-bit systems; its DDS support is atrocious but better than nothing
    If Plugin_DDS.IsDirectXTexAvailable() Or m_FreeImageEnabled Then
        AddInputFormat "DDS - DirectDraw Surface", "*.dds", PDIF_DDS
    End If
        
    If m_FreeImageEnabled Then AddInputFormat "DNG - Adobe Digital Negative", "*.dng", PDIF_RAW
    
    'EMFs are handled via GDI+ for improved rendering and feature compatibility
    AddInputFormat "EMF - Enhanced Metafile", "*.emf", PDIF_EMF
    
    If m_FreeImageEnabled Then
        AddInputFormat "EXR - Industrial Light and Magic", "*.exr", PDIF_EXR
        AddInputFormat "G3 - Digital Fax Format", "*.g3", PDIF_FAXG3
    End If
    
    AddInputFormat "GIF - Graphics Interchange Format", "*.gif;*.agif", PDIF_GIF
    AddInputFormat "HDR - Radiance High Dynamic Range", "*.hdr;*.rgbe;*.xyze;*.pic", PDIF_HDR
    
    'HEIF support requires libheif and several additional support libraries
    If Plugin_Heif.IsLibheifEnabled() Then
        AddInputFormat "HEIC/HEIF - High Efficiency Image Format", "*.heif;*.heifs;*.heic;*.heics;*.avci;*.avcs;*.hif", PDIF_HEIF
    End If
    
    AddInputFormat "HGT - Shuttle Radar Topography Mission (SRTM)", "*.hgt", PDIF_HGT
    AddInputFormat "ICO - Windows Icon", "*.ico", PDIF_ICO
    
    If m_FreeImageEnabled Then AddInputFormat "IFF - Amiga Interchange Format", "*.iff", PDIF_IFF
    
    'Lossless JPEG requires the 3rd-party CharLS library
    If PluginManager.IsPluginCurrentlyEnabled(CCP_CharLS) Then AddInputFormat "JLS - JPEG-LS", "*.jls", PDIF_JLS
    
    If m_FreeImageEnabled Then AddInputFormat "JNG - JPEG Network Graphics", "*.jng", PDIF_JNG
    
    'OpenJPEG is used for JPEG-2000 handling
    If PluginManager.IsPluginCurrentlyEnabled(CCP_OpenJPEG) Then
        AddInputFormat "JP2/J2K - JPEG 2000 File or Codestream", "*.jp2;*.j2k;*.jpt;*.j2c;*.jpc;*.jpx;*.jpf;*.jph", PDIF_JP2
    End If
    
    AddInputFormat "JPG/JPEG - Joint Photographic Experts Group", "*.jpg;*.jpeg;*.jpe;*.jif;*.jfif", PDIF_JPEG
    
    'JPEG XL images require an external plugin, but PD can download it automatically (with permission)
    ' if the current system is compatible, even if the plugin isn't installed at startup.
    If OS.IsWin7OrLater() Then AddInputFormat "JXL - JPEG XL", "*.jxl", PDIF_JXL
    
    If m_FreeImageEnabled Then
        AddInputFormat "JXR/HDP - JPEG XR (HD Photo)", "*.jxr;*.hdp;*.wdp", PDIF_JXR
        AddInputFormat "KOA/KOALA - Commodore 64", "*.koa;*.koala", PDIF_KOALA
        AddInputFormat "LBM - Deluxe Paint", "*.lbm", PDIF_LBM
    End If
    
    AddInputFormat "MBM - Symbian Bitmap", "*.mbm;*.mbw;*.mcl;*.aif;*.abw;*.acl", PDIF_MBM
    AddInputFormat "ORA - OpenRaster", "*.ora", PDIF_ORA
    
    If m_FreeImageEnabled Then
        AddInputFormat "PBM - Portable Bitmap", "*.pbm", PDIF_PBM
        AddInputFormat "PCD - Kodak PhotoCD", "*.pcd", PDIF_PCD
    End If
    
    'In 2025, I wrote a custom DCX/PCX parser
    AddInputFormat "PCX/DCX - Zsoft Paintbrush", "*.pcx;*.pcc;*.dcx", PDIF_PCX
    
    'In v10.0, support was added for PDFs (via pdfium)
    If Plugin_PDF.IsPDFiumAvailable() Then AddInputFormat "PDF - Portable Document Format", "*.pdf", PDIF_PDF
    
    'PDI (PhotoDemon's native file format) is always available!
    AddInputFormat "PDI - PhotoDemon Image", "*.pdi", PDIF_PDI
        
    If m_FreeImageEnabled Then
        AddInputFormat "PFM - Portable Floatmap", "*.pfm", PDIF_PFM
        AddInputFormat "PGM - Portable Graymap", "*.pgm", PDIF_PGM
        AddInputFormat "PIC/PICT - Macintosh Picture", "*.pict;*.pct;*.pic", PDIF_PICT
    End If
    
    AddInputFormat "PNG/APNG - Portable Network Graphic", "*.png;*.apng", PDIF_PNG
        
    If m_FreeImageEnabled Then
        AddInputFormat "PNM - Portable Anymap", "*.pnm", PDIF_PPM
        AddInputFormat "PPM - Portable Pixmap", "*.ppm", PDIF_PPM
    End If
    
    'In v8.0, PhotoDemon received a custom PSD parser
    AddInputFormat "PSD - Adobe Photoshop", "*.psd;*.psb", PDIF_PSD
    
    'In v9.0, PhotoDemon received a custom PSP parser
    AddInputFormat "PSP - PaintShop Pro", "*.psp;*.pspimage;*.tub;*.psptube;*.pfr;*.pspframe;*.msk;*.pspmask;*.pspbrush", PDIF_PSP
    
    'In v9.0, PhotoDemon received a custom QOI parser
    AddInputFormat "QOI - Quite OK Image", "*.qoi", PDIF_QOI
    
    If m_FreeImageEnabled Then
        AddInputFormat "RAS - Sun Raster File", "*.ras", PDIF_RAS
        AddInputFormat "RAW, etc - Raw image data", "*.3fr;*.arw;*.bay;*.bmq;*.cap;*.cine;*.cr2;*.crw;*.cs1;*.dc2;*.dcr;*.dng;*.drf;*.dsc;*.erf;*.fff;*.ia;*.iiq;*.k25;*.kc2;*.kdc;*.mdc;*.mef;*.mos;*.mrw;*.nef;*.nrw;*.orf;*.pef;*.ptx;*.pxn;*.qtk;*.raf;*.raw;*.rdc;*.rw2;*.rwz;*.sr2;*.srf;*.sti", PDIF_RAW
        AddInputFormat "SGI/RGB/BW - Silicon Graphics Image", "*.sgi;*.rgb;*.rgba;*.bw;*.int;*.inta", PDIF_SGI
    End If
    
    'resvg is required for SVG images
    If Plugin_resvg.IsResvgEnabled() Then AddInputFormat "SVG - Scalable Vector Graphics", "*.svg;*.svgz", PDIF_SVG
    
    If m_FreeImageEnabled Then AddInputFormat "TGA - Truevision (TARGA)", "*.tga", PDIF_TARGA
    
    'FreeImage or GDI+ works for loading TIFFs
    AddInputFormat "TIF/TIFF - Tagged Image File Format", "*.tif;*.tiff", PDIF_TIFF
        
    'In v10, I wrote a custom WBMP parser
    AddInputFormat "WBMP - Wireless Bitmap", "*.wbmp;*.wbm;*.wap", PDIF_WBMP
        
    'libwebp is required for WebP files
    If Plugin_WebP.IsWebPEnabled() Then AddInputFormat "WEBP - Google WebP", "*.webp", PDIF_WEBP
    
    'I don't know if anyone still uses WMFs, but GDI+ provides support "for free"
    AddInputFormat "WMF - Windows Metafile", "*.wmf", PDIF_WMF
    
    'In v10, I wrote a custom XBM parser (after being notified that FreeImage hard-crashes on all XBM files)
    AddInputFormat "XBM - X Bitmap", "*.xbm", PDIF_XBM
    
    'In v9.0, I wrote a custom XCF parser for PD
    AddInputFormat "XCF - GIMP (GNU Image Manipulation Program)", "*.xcf;*.xcfgz;*.xcf.gz", PDIF_XCF
    
    If m_FreeImageEnabled Then AddInputFormat "XPM - X Pixmap", "*.xpm", PDIF_XPM
    
    'Finish out the list with an obligatory "All files" option
    AddInputFormat g_Language.TranslateMessage("All files"), "*.*", -1
    
    'Resize our description and extension arrays to match their final size
    m_numOfInputFormats = m_curFormatIndex
    ReDim Preserve inputDescriptions(0 To m_numOfInputFormats) As String
    ReDim Preserve inputExtensions(0 To m_numOfInputFormats) As String
    ReDim Preserve inputPDIFs(0 To m_numOfInputFormats) As PD_IMAGE_FORMAT
    
    'Now that all input files have been added, we can compile a common-dialog-friendly version of this index
    
    'Loop through each entry in the arrays, and append them to the common-dialog-formatted string
    Dim x As Long
    For x = 0 To m_numOfInputFormats
    
        'Index 0 is a special case; everything else is handled in the same manner.
        If (x <> 0) Then
            m_CommonDialogInputs = m_CommonDialogInputs & "|" & inputDescriptions(x) & "|" & inputExtensions(x)
        Else
            m_CommonDialogInputs = inputDescriptions(x) & "|" & inputExtensions(x)
        End If
    
    Next x
    
    'Input format generation complete!
    
End Sub

'Add support for another input format.  A descriptive string and extension list are required.
Private Sub AddInputFormat(ByVal formatDescription As String, ByVal extensionList As String, ByVal correspondingPDIF As PD_IMAGE_FORMAT)
    
    'Increment the counter
    m_curFormatIndex = m_curFormatIndex + 1
    
    'Add the descriptive text to our collection
    inputDescriptions(m_curFormatIndex) = formatDescription
    
    'Add any relevant extension(s) to our collection
    inputExtensions(m_curFormatIndex) = extensionList
    
    'Add the FIF constant
    inputPDIFs(m_curFormatIndex) = correspondingPDIF
    
    'If applicable, add these extensions to the "All Compatible Images" list
    If (extensionList <> "*.*") Then
        If (m_curFormatIndex <> 1) Then
            inputExtensions(0) = inputExtensions(0) & ";" & extensionList
        Else
            inputExtensions(0) = inputExtensions(0) & extensionList
        End If
    End If
            
End Sub

'Generate a list of available export formats
Public Sub GenerateOutputFormats()

    Const MAX_NUM_OUTPUT_FORMATS As Long = 50
    ReDim outputExtensions(0 To MAX_NUM_OUTPUT_FORMATS - 1) As String
    ReDim outputDescriptions(0 To MAX_NUM_OUTPUT_FORMATS - 1) As String
    ReDim outputPDIFs(0 To MAX_NUM_OUTPUT_FORMATS - 1) As PD_IMAGE_FORMAT
    
    'Formats must be added in alphabetical order, as this class has no "sort" functionality.
    
    'If a given export format requires a matching 3rd-party or non-standard system library,
    ' you *must* ensure availability of said library BEFORE adding support for that format.
    
    'Start by effectively setting the location tracker to "0".
    ' (Beyond this point, it will be automatically updated.)
    m_curFormatIndex = -1
    
    'AV1-based codecs require an external plugin, but PD can download it automatically (with permission)
    ' if the current system is compatible, even if the plugin isn't installed at startup.
    If OS.OSSupports64bitExe() Then AddOutputFormat "AVIF - AV1 Image File", "avif", PDIF_AVIF
    
    AddOutputFormat "BMP - Windows Bitmap", "bmp", PDIF_BMP
    AddOutputFormat "DDS - DirectDraw Surface", "dds", PDIF_DDS
    AddOutputFormat "GIF - Graphics Interchange Format", "gif", PDIF_GIF
    AddOutputFormat "HDR - Radiance High Dynamic Range", "hdr", PDIF_HDR
    
    'HEIF support requires libheif and several additional support libraries
    If Plugin_Heif.IsLibheifEnabled() Then AddOutputFormat "HEIC/HEIF - High Efficiency Image File Format", "heic", PDIF_HEIF
    
    AddOutputFormat "ICO - Windows Icon", "ico", PDIF_ICO
    If PluginManager.IsPluginCurrentlyEnabled(CCP_OpenJPEG) Then AddOutputFormat "JP2 - JPEG 2000", "jp2", PDIF_JP2
    AddOutputFormat "JPG - Joint Photographic Experts Group", "jpg", PDIF_JPEG
    
    'JPEG XL images require an external plugin, but PD can download it automatically (with permission)
    ' if the current system is compatible, even if the plugin isn't installed at startup.
    If OS.IsWin7OrLater() Then AddOutputFormat "JXL - JPEG XL", "jxl", PDIF_JXL
    
    If m_FreeImageEnabled Then AddOutputFormat "JXR - JPEG XR (HD Photo)", "jxr", PDIF_JXR
    AddOutputFormat "ORA - OpenRaster", "ora", PDIF_ORA
    AddOutputFormat "PCX - Zsoft Paintbrush", "pcx", PDIF_PCX
    AddOutputFormat "PDI - PhotoDemon Image", "pdi", PDIF_PDI
    AddOutputFormat "PNG - Portable Network Graphic", "png", PDIF_PNG
    If m_FreeImageEnabled Then AddOutputFormat "PNM - Portable Anymap (Netpbm)", "pnm", PDIF_PNM
    AddOutputFormat "PSD - Adobe Photoshop", "psd", PDIF_PSD
    AddOutputFormat "PSP - PaintShop Pro", "psp", PDIF_PSP
    AddOutputFormat "QOI - Quite OK Image", "qoi", PDIF_QOI
    If m_FreeImageEnabled Then AddOutputFormat "TGA - Truevision (TARGA)", "tga", PDIF_TARGA
    AddOutputFormat "TIFF - Tagged Image File Format", "tif", PDIF_TIFF
    AddOutputFormat "WBMP - Wireless Bitmap", "wbm", PDIF_WBMP
    If Plugin_WebP.IsWebPEnabled() Then AddOutputFormat "WEBP - Google WebP", "webp", PDIF_WEBP
    
    'Resize our description and extension arrays to match their final size
    m_numOfOutputFormats = m_curFormatIndex
    ReDim Preserve outputDescriptions(0 To m_numOfOutputFormats) As String
    ReDim Preserve outputExtensions(0 To m_numOfOutputFormats) As String
    ReDim Preserve outputPDIFs(0 To m_numOfOutputFormats) As PD_IMAGE_FORMAT
    
    'Now that all output files have been added, we can compile a common-dialog-friendly version of this index
    
    'Loop through each entry in the arrays, and append them to the common-dialog-formatted string
    Dim x As Long
    For x = 0 To m_numOfOutputFormats
    
        'Index 0 is a special case; everything else is handled in the same manner.
        If (x <> 0) Then
            m_CommonDialogOutputs = m_CommonDialogOutputs & "|" & outputDescriptions(x) & "|*." & outputExtensions(x)
            m_cdOutputDefaultExtensions = m_cdOutputDefaultExtensions & "|." & outputExtensions(x)
        Else
            m_CommonDialogOutputs = outputDescriptions(x) & "|*." & outputExtensions(x)
            m_cdOutputDefaultExtensions = "." & outputExtensions(x)
        End If
    
    Next x
    
    'Output format generation complete!
        
End Sub

'Add support for another output format.  A descriptive string and extension list are required.
Private Sub AddOutputFormat(ByVal formatDescription As String, ByVal extensionList As String, ByVal correspondingPDIF As PD_IMAGE_FORMAT)
    
    m_curFormatIndex = m_curFormatIndex + 1
    
    'Add the descriptive text to our collection
    outputDescriptions(m_curFormatIndex) = formatDescription
    
    'Add the primary extension for this format type
    outputExtensions(m_curFormatIndex) = extensionList
    
    'Add the corresponding FIF
    outputPDIFs(m_curFormatIndex) = correspondingPDIF
    
End Sub

'Given a PDIF (PhotoDemon image format constant), return the default extension.
Public Function GetExtensionFromPDIF(ByVal srcPDIF As PD_IMAGE_FORMAT) As String

    Select Case srcPDIF
        
        Case PDIF_AVIF
            GetExtensionFromPDIF = "avif"
        Case PDIF_BMP
            GetExtensionFromPDIF = "bmp"
        Case PDIF_CBZ
            GetExtensionFromPDIF = "cbz"
        Case PDIF_CUT
            GetExtensionFromPDIF = "cut"
        Case PDIF_DDS
            GetExtensionFromPDIF = "dds"
        Case PDIF_EMF
            GetExtensionFromPDIF = "emf"
        Case PDIF_EXR
            GetExtensionFromPDIF = "exr"
        Case PDIF_FAXG3
            GetExtensionFromPDIF = "g3"
        Case PDIF_GIF
            GetExtensionFromPDIF = "gif"
        Case PDIF_HDR
            GetExtensionFromPDIF = "hdr"
        Case PDIF_HEIF
            GetExtensionFromPDIF = "heic"
        Case PDIF_HGT
            GetExtensionFromPDIF = "hgt"
        Case PDIF_ICO
            GetExtensionFromPDIF = "ico"
        Case PDIF_IFF
            GetExtensionFromPDIF = "iff"
        Case PDIF_J2K
            GetExtensionFromPDIF = "j2k"
        Case PDIF_JLS
            GetExtensionFromPDIF = "jls"
        Case PDIF_JNG
            GetExtensionFromPDIF = "jng"
        Case PDIF_JP2
            GetExtensionFromPDIF = "jp2"
        Case PDIF_JPEG
            GetExtensionFromPDIF = "jpg"
        Case PDIF_JXL
            GetExtensionFromPDIF = "jxl"
        Case PDIF_JXR
            GetExtensionFromPDIF = "jxr"
        Case PDIF_KOALA
            GetExtensionFromPDIF = "koa"
        Case PDIF_LBM
            GetExtensionFromPDIF = "lbm"
        Case PDIF_MBM
            GetExtensionFromPDIF = "mbm"
        Case PDIF_MNG
            GetExtensionFromPDIF = "mng"
        Case PDIF_ORA
            GetExtensionFromPDIF = "ora"
        'Case PDIF_PBM                      'NOTE: for simplicity, all PPM extensions are condensed to PNM
        '    GetExtensionFromPDIF = "pbm"
        'Case PDIF_PBMRAW
        '    GetExtensionFromPDIF = "pbm"
        Case PDIF_PCD
            GetExtensionFromPDIF = "pcd"
        Case PDIF_PCX
            GetExtensionFromPDIF = "pcx"
        Case PDIF_PDF
            GetExtensionFromPDIF = "pdf"
        Case PDIF_PDI
            GetExtensionFromPDIF = "pdi"
        'Case PDIF_PFM
        '    GetExtensionFromPDIF = "pfm"
        'Case PDIF_PGM
        '    GetExtensionFromPDIF = "pgm"
        'Case PDIF_PGMRAW
        '    GetExtensionFromPDIF = "pgm"
        Case PDIF_PICT
            GetExtensionFromPDIF = "pct"
        Case PDIF_PNG
            GetExtensionFromPDIF = "png"
        Case PDIF_PBM, PDIF_PBMRAW, PDIF_PFM, PDIF_PGM, PDIF_PGMRAW, PDIF_PNM, PDIF_PPM, PDIF_PPMRAW
            GetExtensionFromPDIF = "pnm"
        'Case PDIF_PPM
        '    GetExtensionFromPDIF = "ppm"
        'Case PDIF_PPMRAW
        '    GetExtensionFromPDIF = "ppm"
        Case PDIF_PSD
            GetExtensionFromPDIF = "psd"
        Case PDIF_PSP
            GetExtensionFromPDIF = "psp"
        Case PDIF_QOI
            GetExtensionFromPDIF = "qoi"
        Case PDIF_RAS
            GetExtensionFromPDIF = "ras"
        'RAW is an interesting case; because PD can write HDR images, which support nearly all features
        ' of all major RAW formats, we use HDR as the default extension for RAW-type images.
        Case PDIF_RAW
            GetExtensionFromPDIF = "hdr"
        Case PDIF_SGI
            GetExtensionFromPDIF = "sgi"
        Case PDIF_TARGA
            GetExtensionFromPDIF = "tga"
        Case PDIF_TIFF
            GetExtensionFromPDIF = "tif"
        Case PDIF_WBMP
            GetExtensionFromPDIF = "wbm"
        Case PDIF_WEBP
            GetExtensionFromPDIF = "webp"
        Case PDIF_WMF
            GetExtensionFromPDIF = "wmf"
        Case PDIF_XBM
            GetExtensionFromPDIF = "xbm"
        Case PDIF_XCF
            GetExtensionFromPDIF = "xcf"
        Case PDIF_XPM
            GetExtensionFromPDIF = "xpm"
        
        Case Else
            GetExtensionFromPDIF = vbNullString
    
    End Select

End Function

'Given a PDIF (PhotoDemon image format constant), return TRUE if the extension passed is safely associated
' with that format.  (If it isn't, PD will auto-suggest a correct extension instead of using whatever extension
' the file originally had.)
'
'Returns: TRUE if the passed extension is OK; FALSE otherwise.
'         If FALSE is returned, please substitute a correct extension (obtainable from GetExtensionFromPDIF, above).
Public Function IsExtensionOkayForPDIF(ByVal srcPDIF As PD_IMAGE_FORMAT, ByRef srcExtension As String) As Boolean

    Select Case srcPDIF
        
        Case PDIF_AVIF
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "avif", "avci", "avcs", "avifs")
        Case PDIF_BMP
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "bmp")
        Case PDIF_CBZ
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "cbz")
        Case PDIF_CUT
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "cut")
        Case PDIF_DDS
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "dds")
        Case PDIF_EMF
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "emf")
        Case PDIF_EXR
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "exr")
        Case PDIF_FAXG3
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "g3")
        Case PDIF_GIF
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "gif", "agif")
        Case PDIF_HDR
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "hdr", "rgbe", "xyze", "pic")
        Case PDIF_HEIF
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "heif", "heic", "heifs", "heics", "hif")
        Case PDIF_HGT
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "hgt")
        Case PDIF_ICO
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "ico", "cur")
        Case PDIF_IFF
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "iff", "lbm")
        Case PDIF_JP2
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "jp2", "j2k", "jpt", "j2c", "jpc", "jpx", "jpf", "jph", "jpc")
        Case PDIF_JLS
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "jls")
        Case PDIF_JNG
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "jng")
        Case PDIF_JPEG
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "jpg", "jpeg", "jpe", "jfif", "jif")
        Case PDIF_JXL
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "jxl")
        Case PDIF_JXR
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "jxr", "wdp", "hdp")
        Case PDIF_KOALA
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "koa", "koala")
        Case PDIF_LBM
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "lbm", "iff")
        Case PDIF_MBM
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "mbm", "mbw", "mcl", "aif", "abw", "acl")
        Case PDIF_MNG
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "mng")
        Case PDIF_ORA
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "ora")
        'Case PDIF_PBM                      'NOTE: for simplicity, all PPM extensions are condensed to PNM
        '    GetExtensionFromPDIF = "pbm"
        'Case PDIF_PBMRAW
        '    GetExtensionFromPDIF = "pbm"
        Case PDIF_PCD
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "pcd")
        Case PDIF_PCX
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "pcx", "pcc", "dcx")
        Case PDIF_PDF
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "pdf")
        Case PDIF_PDI
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "pdi")
        'Case PDIF_PFM
        '    GetExtensionFromPDIF = "pfm"
        'Case PDIF_PGM
        '    GetExtensionFromPDIF = "pgm"
        'Case PDIF_PGMRAW
        '    GetExtensionFromPDIF = "pgm"
        Case PDIF_PICT
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "pic", "pict", "pct")
        Case PDIF_PNG
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "png", "apng")
        Case PDIF_PBM, PDIF_PBMRAW, PDIF_PFM, PDIF_PGM, PDIF_PGMRAW, PDIF_PNM, PDIF_PPM, PDIF_PPMRAW
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "ppm", "pbm", "pfm", "pgm", "pnm")
        'Case PDIF_PPM
        '    GetExtensionFromPDIF = "ppm"
        'Case PDIF_PPMRAW
        '    GetExtensionFromPDIF = "ppm"
        Case PDIF_PSD
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "psd", "psb")
        Case PDIF_PSP
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "psp", "pspimage", "tub", "psptube", "pfr", "pspframe", "msk", "pspmask", "pspbrush")
        Case PDIF_QOI
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "qoi")
        Case PDIF_RAS
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "ras")
        'RAW is an interesting case; because PD can write HDR images, which support nearly all features
        ' of all major RAW formats, we use HDR as the default extension for RAW-type images.
        Case PDIF_RAW
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "hdr", "3fr", "arw", "bay", "bmq", "cap", "cine", "cr2", "crw", "cs1", "dc2", "dcr", "dng", "drf", "dsc", "erf", "fff", "ia", "iiq", "k25", "kc2", "kdc", "mdc", "mef", "mos", "mrw", "nef", "nrw", "orf", "pef", "ptx", "pxn", "qtk", "raf", "raw", "rdc", "rw2", "rwz", "sr2", "srf", "sti")
        Case PDIF_SGI
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "sgi", "rgb", "rgba", "bw", "int", "inta")
        Case PDIF_SVG
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "svg", "svgz")
        Case PDIF_TARGA
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "tga", "targa")
        Case PDIF_TIFF
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "tif", "tiff")
        Case PDIF_WBMP
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "wbm", "wbmp", "wap")
        Case PDIF_WEBP
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "webp")
        Case PDIF_WMF
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "wmf")
        Case PDIF_XBM
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "xbm")
        Case PDIF_XCF
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "xcf", "xcfgz", "gz")
        Case PDIF_XPM
            IsExtensionOkayForPDIF = Strings.StringsEqualAny(srcExtension, True, "xpm")
        
        Case Else
            IsExtensionOkayForPDIF = False
    
    End Select

End Function

'Given a file extension, attempt to match it to *any* known file format.
' Returns PDIF_UNKNOWN if no matching format+extension combo is found.
Public Function IsExtensionOkayForAnyPDIF(ByRef srcExtension As String) As PD_IMAGE_FORMAT
        
    IsExtensionOkayForAnyPDIF = PDIF_UNKNOWN
    
    Dim i As PD_IMAGE_FORMAT
    For i = 0 To PDIF_MAX_ID
        If IsExtensionOkayForPDIF(i, srcExtension) Then
            IsExtensionOkayForAnyPDIF = i
            Exit Function
        End If
    Next i
    
End Function

'Given a file extension, return the corresponding best-guess PDIF (PhotoDemon image format constant).
'
'By default, this function will *NOT* return PDIF_UNKNOWN for extensions without a corresponding format.
' Instead, it will return PDIF_PDI (PD's default image format).  This is by design so that images with unknown
' or missing extensions correctly default to PD's standard file format at save-time.
'
'If you *want* this function to return PDIF_UNKNOWN when an unclear extension is encountered, set the
' "allowUnknownAsReturn" paramter to TRUE.
Public Function GetPDIFFromExtension(ByVal srcExtension As String, Optional ByVal allowUnknownAsReturn As Boolean = False) As PD_IMAGE_FORMAT
    
    'Shortcut check for non-existent extensions
    If (LenB(srcExtension) = 0) Then
        If allowUnknownAsReturn Then
            GetPDIFFromExtension = PDIF_UNKNOWN
        Else
            GetPDIFFromExtension = PDIF_PDI
        End If
        Exit Function
    End If
    
    srcExtension = LCase$(srcExtension)
    
    'Note that not all possible image format extensions are covered, by design.  Only extensions known to PD
    ' (and available as supported export formats) are checked and returned.
    Select Case srcExtension
        
        Case "avif"
            GetPDIFFromExtension = PDIF_AVIF
        Case "bmp"
            GetPDIFFromExtension = PDIF_BMP
        Case "cbz"
            GetPDIFFromExtension = PDIF_CBZ
        Case "cut"
            GetPDIFFromExtension = PDIF_CUT
        Case "dds"
            GetPDIFFromExtension = PDIF_DDS
        Case "emf"
            GetPDIFFromExtension = PDIF_EMF
        Case "exr"
            GetPDIFFromExtension = PDIF_EXR
        Case "g3"
            GetPDIFFromExtension = PDIF_FAXG3
        Case "gif", "agif"
            GetPDIFFromExtension = PDIF_GIF
        Case "hdr", "rgbe", "xyze", "pic"
            GetPDIFFromExtension = PDIF_HDR
        Case "heif", "heic", "hif"
            GetPDIFFromExtension = PDIF_HEIF
        Case "hgt"
            GetPDIFFromExtension = PDIF_HGT
        Case "ico"
            GetPDIFFromExtension = PDIF_ICO
        Case "iff"
            GetPDIFFromExtension = PDIF_IFF
            GetPDIFFromExtension = PDIF_JP2
        Case "jls"
            GetPDIFFromExtension = PDIF_JLS
        Case "jng"
            GetPDIFFromExtension = PDIF_JNG
        Case "jp2", "j2k", "j2c"
            GetPDIFFromExtension = PDIF_JP2
        Case "jpg", "jpe", "jpeg"
            GetPDIFFromExtension = PDIF_JPEG
        Case "jxl"
            GetPDIFFromExtension = PDIF_JXL
        Case "jxr"
            GetPDIFFromExtension = PDIF_JXR
        Case "koa"
            GetPDIFFromExtension = PDIF_KOALA
        Case "lbm"
            GetPDIFFromExtension = PDIF_LBM
        Case "mbm", "aif"
            GetPDIFFromExtension = PDIF_MBM
        Case "mng"
            GetPDIFFromExtension = PDIF_MNG
        Case "ora"
            GetPDIFFromExtension = PDIF_ORA
        Case "pcd"
            GetPDIFFromExtension = PDIF_PCD
        Case "pcx"
            GetPDIFFromExtension = PDIF_PCX
        Case "pdf"
            GetPDIFFromExtension = PDIF_PDF
        Case "pdi"
            GetPDIFFromExtension = PDIF_PDI
        Case "pct"
            GetPDIFFromExtension = PDIF_PICT
        Case "png", "apng"
            GetPDIFFromExtension = PDIF_PNG
        Case "pbm", "pfm", "pgm", "pnm", "ppm"
            GetPDIFFromExtension = PDIF_PNM
        Case "psd", "psb"
            GetPDIFFromExtension = PDIF_PSD
        Case "psp", "pspimage"
            GetPDIFFromExtension = PDIF_PSP
        Case "qoi"
            GetPDIFFromExtension = PDIF_QOI
        Case "ras"
            GetPDIFFromExtension = PDIF_RAS
        Case "sgi"
            GetPDIFFromExtension = PDIF_SGI
        Case "tga"
            GetPDIFFromExtension = PDIF_TARGA
        Case "tif", "tiff"
            GetPDIFFromExtension = PDIF_TIFF
        Case "wbm", "wbmp"
            GetPDIFFromExtension = PDIF_WBMP
        Case "webp"
            GetPDIFFromExtension = PDIF_WEBP
        Case "wmf"
            GetPDIFFromExtension = PDIF_WMF
        Case "xbm"
            GetPDIFFromExtension = PDIF_XBM
        Case "xcf"
            GetPDIFFromExtension = PDIF_XCF
        Case "xpm"
            GetPDIFFromExtension = PDIF_XPM
        
        Case Else
            If allowUnknownAsReturn Then
                GetPDIFFromExtension = PDIF_UNKNOWN
            Else
                GetPDIFFromExtension = PDIF_PDI
            End If
            
    End Select

End Function

'Given an output PDIF, return the ideal metadata format for that image format.
Public Function GetIdealMetadataFormatFromPDIF(ByVal outputPDIF As PD_IMAGE_FORMAT) As PD_METADATA_FORMAT
    Select Case outputPDIF
        Case PDIF_BMP
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        Case PDIF_GIF
            GetIdealMetadataFormatFromPDIF = PDMF_XMP
        Case PDIF_HDR
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        Case PDIF_HEIF
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        Case PDIF_ICO
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        Case PDIF_JP2
            GetIdealMetadataFormatFromPDIF = PDMF_XMP
        Case PDIF_JPEG
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        Case PDIF_JXL
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        Case PDIF_JXR
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        Case PDIF_MBM
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        Case PDIF_PCX
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        Case PDIF_PDF
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        Case PDIF_PDI
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        Case PDIF_PNG
            GetIdealMetadataFormatFromPDIF = PDMF_XMP
        Case PDIF_PNM
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        Case PDIF_PSD
            GetIdealMetadataFormatFromPDIF = PDMF_XMP
        Case PDIF_PSP
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        Case PDIF_QOI
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        Case PDIF_TARGA
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        Case PDIF_TIFF
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        Case PDIF_WEBP
            GetIdealMetadataFormatFromPDIF = PDMF_XMP
        Case PDIF_XCF
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        Case Else
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
    End Select
End Function

'Given an output PDIF, return a BOOLEAN specifying whether the export format supports animation.
Public Function IsAnimationSupported(ByVal outputPDIF As PD_IMAGE_FORMAT) As Boolean
    Select Case outputPDIF
        Case PDIF_GIF, PDIF_PNG, PDIF_WEBP, PDIF_JXL
            IsAnimationSupported = True
        Case Else
            IsAnimationSupported = False
    End Select
End Function

'Given an output PDIF, return a BOOLEAN specifying whether Exif metadata is allowed for that image format.
' (Technically, ExifTool can write non-standard Exif chunks for formats like PNG and JPEG-2000,
' but PD prefers not to do this as the resulting data can't be read by anything but ExifTool itself.)

'If an Exif tag can't be converted to a corresponding XMP tag, it should simply be removed from the new file.)
Public Function IsExifAllowedForPDIF(ByVal outputPDIF As PD_IMAGE_FORMAT) As Boolean
    Select Case outputPDIF
        Case PDIF_HEIF
            IsExifAllowedForPDIF = True
        Case PDIF_JPEG
            IsExifAllowedForPDIF = True
        Case PDIF_JXL
            IsExifAllowedForPDIF = True
        Case PDIF_JXR
            IsExifAllowedForPDIF = True
        Case PDIF_PDI
            IsExifAllowedForPDIF = True
        Case PDIF_PSD
            IsExifAllowedForPDIF = True
        Case PDIF_PSP
            IsExifAllowedForPDIF = True
        Case PDIF_TIFF
            IsExifAllowedForPDIF = True
        Case Else
            IsExifAllowedForPDIF = False
    End Select
End Function

'Given an output PDIF, return a BOOLEAN specifying whether PD has implemented an export dialog for that image format.
Public Function IsExportDialogSupported(ByVal outputPDIF As PD_IMAGE_FORMAT) As Boolean
    Select Case outputPDIF
        Case PDIF_AVIF
            IsExportDialogSupported = True
        Case PDIF_BMP
            IsExportDialogSupported = True
        Case PDIF_DDS
            IsExportDialogSupported = True
        Case PDIF_GIF
            IsExportDialogSupported = True
        Case PDIF_HEIF
            IsExportDialogSupported = True
        Case PDIF_ICO
            IsExportDialogSupported = True
        Case PDIF_JP2
            IsExportDialogSupported = True
        Case PDIF_JPEG
            IsExportDialogSupported = True
        Case PDIF_JXL
            IsExportDialogSupported = True
        Case PDIF_JXR
            IsExportDialogSupported = True
        Case PDIF_PNG
            IsExportDialogSupported = True
        Case PDIF_PBM, PDIF_PGM, PDIF_PNM, PDIF_PPM
            IsExportDialogSupported = True
        Case PDIF_PSD
            IsExportDialogSupported = True
        Case PDIF_PSP
            IsExportDialogSupported = True
        Case PDIF_TIFF
            IsExportDialogSupported = True
        Case PDIF_WEBP
            IsExportDialogSupported = True
        Case Else
            IsExportDialogSupported = False
    End Select
End Function

Public Function IsExifToolRelevant(ByVal srcFormat As PD_IMAGE_FORMAT) As Boolean
    Select Case srcFormat
        Case PDIF_CBZ
            IsExifToolRelevant = False
        Case PDIF_ORA
            IsExifToolRelevant = False
        Case PDIF_PDF
            IsExifToolRelevant = False
        Case PDIF_PDI
            IsExifToolRelevant = False
        Case PDIF_QOI
            IsExifToolRelevant = False
        Case Else
            IsExifToolRelevant = True
    End Select
End Function

'Helper wrappers to declare FreeImage availability as a last-resort fallback
Public Function IsFreeImageEnabled() As Boolean
    IsFreeImageEnabled = m_FreeImageEnabled
End Function

Public Sub SetFreeImageEnabled(ByVal newState As Boolean)
    m_FreeImageEnabled = newState
End Sub

'When the active language changes, we need to calculate new translations for text like "All Compatible Images"
Public Sub NotifyLanguageChanged()
    GenerateInputFormats
    GenerateOutputFormats
End Sub
