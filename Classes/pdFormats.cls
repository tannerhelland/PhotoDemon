VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Format Manager
'Copyright 2012-2018 by Tanner Helland
'Created: 18/November/12
'Last updated: 02/May/16
'Last update: expand PNM variant support
'
'This class is the new handler for PhotoDemon's vast image format support library.  Previously, individual functions
' (such as batch convert) had to manage specialized versions of this data, which made it extremely cumbersome to add
' or adjust PhotoDemon's support for individual formats.  To remedy this problem, this class was created.
'
'Based on the available plugins, this class generates a list of file formats that PhotoDemon is capable of writing
' and reading.  These format lists are separately maintained, and the presence of a format in e.g. the Import category
' does not guarantee a similar presence in the Export category.
'
'Most non-standard formats rely on FreeImage for loading and/or saving, and they can additionally be tested against
' individual FreeImage version numbers if necessary.  (Thus far, this has *not* proven necessary.)
'
'GDI+ is optionally used as a fallback if FreeImage is not present, but it should be noted that GDI+ interoperability
' is not as well-tested as FreeImage.  In some cases, however, it IS faster to use GDI+ due to the need to make a
' specialized copy of DIB data for FreeImage's internal handling - so in the case of batch processing, GDI+ may be
' used preferentially over FreeImage for reading and writing certain filetypes.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Is the FreeImage DLL available?
Public FreeImageEnabled As Boolean

'Is GDI+ available?
Public GDIPlusEnabled As Boolean

'Is pngQuant available?
Public pngQuantEnabled As Boolean

'Number of available input, output formats
Private numOfInputFormats As Long, numOfOutputFormats As Long

'Array of available input, output extensions.
Private inputExtensions() As String
Private outputExtensions() As String

'Array of "friendly" descriptions for input, output formats
Private inputDescriptions() As String
Private outputDescriptions() As String

'Array of corresponding image format constants
Private inputPDIFs() As PD_IMAGE_FORMAT
Private outputPDIFs() As PD_IMAGE_FORMAT

'Array of common-dialog-formatted input/output filetypes.  (Common dialogs require different pipe-based formatting
' than normal lists, as you must add human-readable text descriptions to the list.)
Private m_CommonDialogInputs As String, m_CommonDialogOutputs As String

'Common dialog also require a specialized "default extension" string for output files
Private m_cdOutputDefaultExtensions As String

'When analyzing our current plugin options, we have to make a lot of on-the-fly decisions about format availability.
' This value is shared among multiple functions while making such decisions.  Do not treat it as a meaningful value.
Private m_curFormatIndex As Long

'Return the index of given input PDIF
Friend Function GetIndexOfInputPDIF(ByVal srcFIF As PD_IMAGE_FORMAT) As Long
    
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
Friend Function GetInputPDIF(ByVal dIndex As Long) As Long
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        GetInputPDIF = inputPDIFs(dIndex)
    Else
        GetInputPDIF = FIF_UNKNOWN
    End If
End Function

'Return the friendly input format description at a given index
Friend Function GetInputFormatDescription(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        GetInputFormatDescription = inputDescriptions(dIndex)
    Else
        GetInputFormatDescription = vbNullString
    End If
End Function

'Return the input format extension at a given index
Friend Function GetInputFormatExtensions(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        GetInputFormatExtensions = inputExtensions(dIndex)
    Else
        GetInputFormatExtensions = vbNullString
    End If
End Function

'Return a list of all input formats supported by the current session.  By default, the list is delimited with commas,
' and each extension is listed as "*.abc".
Friend Function GetListOfInputFormats(Optional ByVal listDelimiter As String = ";", Optional ByVal includeStarDot As Boolean = True) As String
    
    'The first entry in the extensions collection is used for the "All supported formats" common dialog option;
    ' as such, it already contains a full list of valid extensions.
    GetListOfInputFormats = inputExtensions(0)
    
    If Strings.StringsNotEqual(listDelimiter, ";", False) Then GetListOfInputFormats = Replace$(GetListOfInputFormats, ";", listDelimiter)
    If (Not includeStarDot) Then GetListOfInputFormats = Replace$(GetListOfInputFormats, "*.", vbNullString)
    
End Function

'Return the number of available input format types
Friend Function GetNumOfInputFormats() As Long
    GetNumOfInputFormats = numOfInputFormats
End Function

'Return a list of input filetypes formatted for use with a common dialog box
Friend Function GetCommonDialogInputFormats() As String
    GetCommonDialogInputFormats = m_CommonDialogInputs
End Function

'Return the index of given output FIF
Friend Function GetIndexOfOutputPDIF(ByVal srcFIF As PD_IMAGE_FORMAT) As Long
    
    Dim i As Long
    For i = 0 To GetNumOfOutputFormats
        If outputPDIFs(i) = srcFIF Then
            GetIndexOfOutputPDIF = i
            Exit Function
        End If
    Next i
    
    'If we reach this line of code, no match was found.  Return -1.
    GetIndexOfOutputPDIF = -1
    
End Function

'Return the FIF (image format constant) at a given index
Friend Function GetOutputPDIF(ByVal dIndex As Long) As PD_IMAGE_FORMAT
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        GetOutputPDIF = outputPDIFs(dIndex)
    Else
        GetOutputPDIF = FIF_UNKNOWN
    End If
End Function

'Return the friendly output format description at a given index
Friend Function GetOutputFormatDescription(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfOutputFormats) Then
        GetOutputFormatDescription = outputDescriptions(dIndex)
    Else
        GetOutputFormatDescription = vbNullString
    End If
End Function

'Return the output format extension at a given index
Friend Function GetOutputFormatExtension(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfOutputFormats) Then
        GetOutputFormatExtension = outputExtensions(dIndex)
    Else
        GetOutputFormatExtension = vbNullString
    End If
End Function

'Return the number of available output format types
Friend Function GetNumOfOutputFormats() As Long
    GetNumOfOutputFormats = numOfOutputFormats
End Function

'Return a list of output filetypes formatted for use with a common dialog box
Friend Function GetCommonDialogOutputFormats() As String
    GetCommonDialogOutputFormats = m_CommonDialogOutputs
End Function

'Return a list of output default extensions formatted for use with a common dialog box
Friend Function GetCommonDialogDefaultExtensions() As String
    GetCommonDialogDefaultExtensions = m_cdOutputDefaultExtensions
End Function

'Generate a list of available import formats
Friend Sub GenerateInputFormats()

    'Prepare a list of possible INPUT formats based on the plugins available to us
    ReDim inputExtensions(0 To 50) As String
    ReDim inputDescriptions(0 To 50) As String
    ReDim inputPDIFs(0 To 50) As PD_IMAGE_FORMAT

    'Formats should be added in alphabetical order, as this class has no "sort" functionality.

    'Always start with an "All Compatible Images" option
    inputDescriptions(0) = g_Language.TranslateMessage("All Compatible Images")
    
    'Unique to this first one is the full list of compatible extensions.  Instead of generating a full list here,
    ' it will be automatically generated as we go.
    
    'Set the location tracker to "0".  Beyond this point, it will be automatically updated.
    m_curFormatIndex = 0
    
    'Bitmap files require no plugins; they are always supported.
    AddInputFormat "BMP - Windows or OS/2 Bitmap", "*.bmp", PDIF_BMP
    
    If FreeImageEnabled Then
        AddInputFormat "DDS - DirectDraw Surface", "*.dds", PDIF_DDS
        AddInputFormat "DNG - Adobe Digital Negative", "*.dng", PDIF_RAW
    End If
    
    If GDIPlusEnabled Then AddInputFormat "EMF - Enhanced Metafile", "*.emf", PDIF_EMF
    
    If FreeImageEnabled Then
        AddInputFormat "EXR - Industrial Light and Magic", "*.exr", PDIF_EXR
        AddInputFormat "G3 - Digital Fax Format", "*.g3", PDIF_FAXG3
    End If
    
    AddInputFormat "GIF - Compuserve", "*.gif", PDIF_GIF
    
    If FreeImageEnabled Then AddInputFormat "HDR - High Dynamic Range", "*.hdr", PDIF_HDR
    
    AddInputFormat "ICO - Windows Icon", "*.ico", PDIF_ICO
    
    If FreeImageEnabled Then
        AddInputFormat "IFF - Amiga Interchange Format", "*.iff", PDIF_IFF
        AddInputFormat "JNG - JPEG Network Graphics", "*.jng", PDIF_JNG
        AddInputFormat "JP2/J2K - JPEG 2000 File or Codestream", "*.jp2;*.j2k;*.jpc;*.jpx;*.jpf", PDIF_JP2
    End If
    
    AddInputFormat "JPG/JPEG - Joint Photographic Experts Group", "*.jpg;*.jpeg;*.jpe;*.jif;*.jfif", PDIF_JPEG
    
    If FreeImageEnabled Then
        AddInputFormat "JXR/HDP - JPEG XR (HD Photo)", "*.jxr;*.hdp;*.wdp", PDIF_JXR
        AddInputFormat "KOA/KOALA - Commodore 64", "*.koa;*.koala", PDIF_KOALA
        AddInputFormat "LBM - Deluxe Paint", "*.lbm", PDIF_LBM
        AddInputFormat "MNG - Multiple Network Graphics", "*.mng", PDIF_MNG
        AddInputFormat "PBM - Portable Bitmap", "*.pbm", PDIF_PBM
        AddInputFormat "PCD - Kodak PhotoCD", "*.pcd", PDIF_PCD
        AddInputFormat "PCX - Zsoft Paintbrush", "*.pcx", PDIF_PCX
    End If
    
    'PDI (PhotoDemon's native file format) is always available!
    AddInputFormat "PDI - PhotoDemon Image", "*.pdi", PDIF_PDI
        
    If FreeImageEnabled Then
        AddInputFormat "PFM - Portable Floatmap", "*.pfm", PDIF_PFM
        AddInputFormat "PGM - Portable Graymap", "*.pgm", PDIF_PGM
        AddInputFormat "PIC/PICT - Macintosh Picture", "*.pict;*.pct;*.pic", PDIF_PICT
    End If
    
    'FreeImage or GDI+ works for loading PNGs
    If FreeImageEnabled Or GDIPlusEnabled Then AddInputFormat "PNG - Portable Network Graphic", "*.png", PDIF_PNG
        
    If FreeImageEnabled Then
        AddInputFormat "PNM - Portable Anymap", "*.pnm", PDIF_PPM
        AddInputFormat "PPM - Portable Pixmap", "*.ppm", PDIF_PPM
        AddInputFormat "PSD - Adobe Photoshop", "*.psd;*.psb", PDIF_PSD
        AddInputFormat "RAS - Sun Raster File", "*.ras", PDIF_RAS
        AddInputFormat "RAW, etc - Raw image data", "*.3fr;*.arw;*.bay;*.bmq;*.cap;*.cine;*.cr2;*.crw;*.cs1;*.dc2;*.dcr;*.dng;*.drf;*.dsc;*.erf;*.fff;*.ia;*.iiq;*.k25;*.kc2;*.kdc;*.mdc;*.mef;*.mos;*.mrw;*.nef;*.nrw;*.orf;*.pef;*.ptx;*.pxn;*.qtk;*.raf;*.raw;*.rdc;*.rw2;*.rwz;*.sr2;*.srf;*.sti", PDIF_RAW
        AddInputFormat "SGI/RGB/BW - Silicon Graphics Image", "*.sgi;*.rgb;*.rgba;*.bw;*.int;*.inta", PDIF_SGI
        AddInputFormat "TGA - Truevision (TARGA)", "*.tga", PDIF_TARGA
    End If
    
    'FreeImage or GDI+ works for loading TIFFs
    If FreeImageEnabled Or GDIPlusEnabled Then AddInputFormat "TIF/TIFF - Tagged Image File Format", "*.tif;*.tiff", PDIF_TIFF
        
    If FreeImageEnabled Then
        AddInputFormat "WBMP - Wireless Bitmap", "*.wbmp;*.wbm", PDIF_WBMP
        AddInputFormat "WEBP - Google WebP", "*.webp", PDIF_WEBP
    End If
    
    If GDIPlusEnabled Then AddInputFormat "WMF - Windows Metafile", "*.wmf", PDIF_WMF
    
    'Finish out the list with an obligatory "All files" option
    AddInputFormat g_Language.TranslateMessage("All files"), "*.*", -1
    
    'Resize our description and extension arrays to match their final size
    numOfInputFormats = m_curFormatIndex
    ReDim Preserve inputDescriptions(0 To numOfInputFormats) As String
    ReDim Preserve inputExtensions(0 To numOfInputFormats) As String
    ReDim Preserve inputPDIFs(0 To numOfInputFormats) As PD_IMAGE_FORMAT
    
    'Now that all input files have been added, we can compile a common-dialog-friendly version of this index
    
    'Loop through each entry in the arrays, and append them to the common-dialog-formatted string
    Dim x As Long
    For x = 0 To numOfInputFormats
    
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
Friend Sub GenerateOutputFormats()

    ReDim outputExtensions(0 To 50) As String
    ReDim outputDescriptions(0 To 50) As String
    ReDim outputPDIFs(0 To 50) As PD_IMAGE_FORMAT

    'Formats should be added in alphabetical order, as this class has no "sort" functionality.
    
    'Start by effectively setting the location tracker to "0".  Beyond this point, it will be automatically updated.
    m_curFormatIndex = -1

    AddOutputFormat "BMP - Windows Bitmap", "bmp", PDIF_BMP
    
    If FreeImageEnabled Or GDIPlusEnabled Then AddOutputFormat "GIF - Graphics Interchange Format", "gif", PDIF_GIF
    
    If FreeImageEnabled Then
        AddOutputFormat "HDR - High Dynamic Range", "hdr", PDIF_HDR
        AddOutputFormat "JP2 - JPEG 2000", "jp2", PDIF_JP2
    End If
        
    If FreeImageEnabled Or GDIPlusEnabled Then AddOutputFormat "JPG - Joint Photographic Experts Group", "jpg", PDIF_JPEG
        
    If FreeImageEnabled Then AddOutputFormat "JXR - JPEG XR (HD Photo)", "jxr", PDIF_JXR
        
    AddOutputFormat "PDI - PhotoDemon Image", "pdi", PDIF_PDI
    
    If FreeImageEnabled Or GDIPlusEnabled Then AddOutputFormat "PNG - Portable Network Graphic", "png", PDIF_PNG
    
    If FreeImageEnabled Then
        AddOutputFormat "PNM - Portable Anymap (Netpbm)", "pnm", PDIF_PNM
        AddOutputFormat "PSD - Photoshop Document", "psd", PDIF_PSD
        AddOutputFormat "TGA - Truevision (TARGA)", "tga", PDIF_TARGA
    End If
    
    If FreeImageEnabled Or GDIPlusEnabled Then AddOutputFormat "TIFF - Tagged Image File Format", "tif", PDIF_TIFF
    
    If FreeImageEnabled Then AddOutputFormat "WEBP - Google WebP", "webp", PDIF_WEBP
    
    'Resize our description and extension arrays to match their final size
    numOfOutputFormats = m_curFormatIndex
    ReDim Preserve outputDescriptions(0 To numOfOutputFormats) As String
    ReDim Preserve outputExtensions(0 To numOfOutputFormats) As String
    ReDim Preserve outputPDIFs(0 To numOfOutputFormats) As PD_IMAGE_FORMAT
    
    'Now that all output files have been added, we can compile a common-dialog-friendly version of this index
    
    'Loop through each entry in the arrays, and append them to the common-dialog-formatted string
    Dim x As Long
    For x = 0 To numOfOutputFormats
    
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
    
    'Increment the counter
    m_curFormatIndex = m_curFormatIndex + 1
    
    'Add the descriptive text to our collection
    outputDescriptions(m_curFormatIndex) = formatDescription
    
    'Add the primary extension for this format type
    outputExtensions(m_curFormatIndex) = extensionList
    
    'Add the corresponding FIF
    outputPDIFs(m_curFormatIndex) = correspondingPDIF
            
End Sub

'Given a PDIF (PhotoDemon image format constant), return the default extension.
Friend Function GetExtensionFromPDIF(ByVal srcPDIF As PD_IMAGE_FORMAT) As String

    Select Case srcPDIF
    
        Case PDIF_BMP
            GetExtensionFromPDIF = "bmp"
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
        Case PDIF_ICO
            GetExtensionFromPDIF = "ico"
        Case PDIF_IFF
            GetExtensionFromPDIF = "iff"
        Case PDIF_J2K
            GetExtensionFromPDIF = "j2k"
        Case PDIF_JNG
            GetExtensionFromPDIF = "jng"
        Case PDIF_JP2
            GetExtensionFromPDIF = "jp2"
        Case PDIF_JPEG
            GetExtensionFromPDIF = "jpg"
        Case PDIF_JXR
            GetExtensionFromPDIF = "jxr"
        Case PDIF_KOALA
            GetExtensionFromPDIF = "koa"
        Case PDIF_LBM
            GetExtensionFromPDIF = "lbm"
        Case PDIF_MNG
            GetExtensionFromPDIF = "mng"
        'Case PDIF_PBM                      'NOTE: for simplicity, all PPM extensions are condensed to PNM
        '    GetExtensionFromPDIF = "pbm"
        'Case PDIF_PBMRAW
        '    GetExtensionFromPDIF = "pbm"
        Case PDIF_PCD
            GetExtensionFromPDIF = "pcd"
        Case PDIF_PCX
            GetExtensionFromPDIF = "pcx"
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
        Case PDIF_RAS
            GetExtensionFromPDIF = "ras"
        'RAW is an interesting case; because PD can write HDR images, which support nearly all features of all major RAW formats,
        ' we use HDR as the default extension for RAW-type images.
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
        Case PDIF_XPM
            GetExtensionFromPDIF = "xpm"
        
        Case Else
            GetExtensionFromPDIF = vbNullString
    
    End Select

End Function

'This can be used to see if an output format supports multiple color depths.
Friend Function DoesPDIFSupportMultipleColorDepths(ByVal outputPDIF As PD_IMAGE_FORMAT) As Boolean

    Select Case outputPDIF
    
        Case PDIF_GIF, PDIF_HDR, PDIF_JPEG
            DoesPDIFSupportMultipleColorDepths = False
            
        Case Else
            DoesPDIFSupportMultipleColorDepths = True
    
    End Select

End Function

'Given a file format and color depth, are the two compatible?  (NOTE: this function takes into account the availability of FreeImage and/or GDI+)
Friend Function IsColorDepthSupported(ByVal outputPDIF As Long, ByVal desiredColorDepth As Long) As Boolean
    
    'Check the special case of PDI (internal PhotoDemon images)
    If (outputPDIF = PDIF_PDI) Then
        IsColorDepthSupported = True
        Exit Function
    End If

    'By default, report that a given color depth is NOT supported
    IsColorDepthSupported = False
    
    'First, address formats handled only by FreeImage
    If FreeImageEnabled Then
        
        Select Case outputPDIF
        
            'BMP
            Case PDIF_BMP
            
                Select Case desiredColorDepth
        
                    Case 1
                        IsColorDepthSupported = True
                    Case 4
                        IsColorDepthSupported = True
                    Case 8
                        IsColorDepthSupported = True
                    Case 16
                        IsColorDepthSupported = True
                    Case 24
                        IsColorDepthSupported = True
                    Case 32
                        IsColorDepthSupported = True
                    Case Else
                        IsColorDepthSupported = False
                        
                End Select
        
            'GIF
            Case PDIF_GIF
            
                If desiredColorDepth = 8 Then IsColorDepthSupported = True Else IsColorDepthSupported = False
                
            'HDR
            Case PDIF_HDR
            
                If desiredColorDepth = 24 Then IsColorDepthSupported = True Else IsColorDepthSupported = False
                
            'JP2 (JPEG 2000)
            Case PDIF_JP2
            
                Select Case desiredColorDepth
                
                    Case 24
                        IsColorDepthSupported = True
                    Case 32
                        IsColorDepthSupported = True
                    Case Else
                        IsColorDepthSupported = False
                        
                End Select
                
            'JPEG
            Case PDIF_JPEG
            
                If desiredColorDepth = 24 Then IsColorDepthSupported = True Else IsColorDepthSupported = False
            
            'JXR (JPEG XR)
            Case PDIF_JXR
            
                Select Case desiredColorDepth
                
                    Case 24
                        IsColorDepthSupported = True
                    Case 32
                        IsColorDepthSupported = True
                    Case Else
                        IsColorDepthSupported = False
                        
                End Select
                
            'PNG
            Case PDIF_PNG
        
                Select Case desiredColorDepth
        
                    Case 1
                        IsColorDepthSupported = True
                    Case 4
                        IsColorDepthSupported = True
                    Case 8
                        IsColorDepthSupported = True
                    Case 24
                        IsColorDepthSupported = True
                    Case 32
                        IsColorDepthSupported = True
                    Case Else
                        IsColorDepthSupported = False
                        
                End Select
                
            'PPM (Portable Pixmap)
            Case PDIF_PPM
            
                If desiredColorDepth = 24 Then IsColorDepthSupported = True Else IsColorDepthSupported = False
            
            'PSD (Photoshop document)
            Case PDIF_PSD
                    
                Select Case desiredColorDepth
        
                    Case 1
                        IsColorDepthSupported = True
                    Case 4
                        IsColorDepthSupported = True
                    Case 8
                        IsColorDepthSupported = True
                    Case 24
                        IsColorDepthSupported = True
                    Case 32
                        IsColorDepthSupported = True
                    
                    'Higher bit-depths may be supported, but I'm not enabling them until high bit-depth output is
                    ' better supported throughout all of PD.
                    Case Else
                        IsColorDepthSupported = False
                        
                End Select
            
            'TGA (Targa)
            Case PDIF_TARGA
            
                Select Case desiredColorDepth
                
                    Case 8
                        IsColorDepthSupported = True
                    Case 24
                        IsColorDepthSupported = True
                    Case 32
                        IsColorDepthSupported = True
                    Case Else
                        IsColorDepthSupported = False
                
                End Select
                
            'TIFF
            Case PDIF_TIFF
            
                Select Case desiredColorDepth
        
                    Case 1
                        IsColorDepthSupported = True
                    Case 4
                        IsColorDepthSupported = True
                    Case 8
                        IsColorDepthSupported = True
                    Case 24
                        IsColorDepthSupported = True
                    Case 32
                        IsColorDepthSupported = True
                    Case Else
                        IsColorDepthSupported = False
                        
                End Select
                
            'WebP
            Case PDIF_WEBP
            
                Select Case desiredColorDepth
                
                    Case 24
                        IsColorDepthSupported = True
                    Case 32
                        IsColorDepthSupported = True
                    Case Else
                        IsColorDepthSupported = False
                        
                End Select
        
        End Select
        
        'Because FreeImage covers every available file type, we can now exit the function with whatever value has been set
        Exit Function
        
    End If
    
    'If we have reached this line of code, FreeImage is not available.
    
    'So check to see what color depths GDI+ can offer
    If GDIPlusEnabled Then
    
        Select Case outputPDIF
        
            'GIF
            Case PDIF_GIF
            
                If desiredColorDepth = 8 Then IsColorDepthSupported = True Else IsColorDepthSupported = False
                
            'JPEG
            Case PDIF_JPEG
            
                If desiredColorDepth = 24 Then IsColorDepthSupported = True Else IsColorDepthSupported = False
                
            'PNG
            Case PDIF_PNG
        
                Select Case desiredColorDepth
        
                    Case 1
                        IsColorDepthSupported = True
                    Case 4
                        IsColorDepthSupported = True
                    Case 8
                        IsColorDepthSupported = True
                    Case 24
                        IsColorDepthSupported = True
                    Case 32
                        IsColorDepthSupported = True
                    Case Else
                        IsColorDepthSupported = False
                        
                End Select
                
            'TIFF
            Case PDIF_TIFF
            
                Select Case desiredColorDepth
        
                    Case 1
                        IsColorDepthSupported = True
                    Case 4
                        IsColorDepthSupported = True
                    Case 8
                        IsColorDepthSupported = True
                    Case 24
                        IsColorDepthSupported = True
                    Case 32
                        IsColorDepthSupported = True
                    Case Else
                        IsColorDepthSupported = False
                        
                End Select
        
        End Select
        
        'Because GDI+ covers every available file type (if FreeImage can't be found), we can now exit the function with whatever value has been set
        Exit Function
    
    End If
    
    'Finally, if FreeImage can't be found, and GDI+ can't be found, the user better like outputting BMP files, because that's all they get.
    If outputPDIF = PDIF_BMP Then
    
        If (desiredColorDepth = 24) Or (desiredColorDepth = 32) Then IsColorDepthSupported = True Else IsColorDepthSupported = False
        
    End If

End Function

'Given a file format and desired color depth, return the next-best color depth that can be used (assuming the desired one is not available)
' (NOTE: this function takes into account the availability of FreeImage and/or GDI+)
Friend Function GetClosestColorDepth(ByVal outputPDIF As Long, ByVal desiredColorDepth As Long) As Long
    
    'Check the special case of PDI (internal PhotoDemon images)
    If (outputPDIF = PDIF_PDI) Then
        GetClosestColorDepth = desiredColorDepth
        Exit Function
    End If

    'By default, report that 24bpp is the preferred alternative
    GetClosestColorDepth = 24

    'Certain file formats only support one output color depth, so they are easily handled (e.g. GIF)
    
    'Some file formats support many color depths (PNG, for example, can handle 1/4/8/24/32)
    
    'Some file formats offer different color depths based on what encoders are available (BMP, for example, has access to 1/4/8 bpp
    ' if either of FreeImage or GDI+ is enabled)
    
    'Based on all these factors, return the nearest color depth to the one the user would like to use (but can't, presumably)
    Select Case outputPDIF
    
        'BMP support changes based on the available encoder
        Case PDIF_BMP
        
            'FreeImage and GDI+ offer identical bpp support
            If FreeImageEnabled Or GDIPlusEnabled Then
            
                If desiredColorDepth <= 1 Then
                    GetClosestColorDepth = 1
                ElseIf desiredColorDepth <= 4 Then
                    GetClosestColorDepth = 4
                ElseIf desiredColorDepth <= 8 Then
                    GetClosestColorDepth = 8
                ElseIf desiredColorDepth <= 24 Then
                    GetClosestColorDepth = 24
                Else
                    GetClosestColorDepth = 32
                End If
            
            'If both FreeImage and GDI+ are missing, we can only write 24 and 32bpp bitmaps
            Else
                If desiredColorDepth <= 24 Then
                    GetClosestColorDepth = 24
                Else
                    GetClosestColorDepth = 32
                End If
            End If
        
        'GIF only supports 8bpp
        Case PDIF_GIF
            GetClosestColorDepth = 8
            
        'HDR only supports 24bpp
        Case PDIF_HDR
            GetClosestColorDepth = 24
            
        'JP2 (JPEG 2000) supports 24/32bpp
        Case PDIF_JP2
            If desiredColorDepth <= 24 Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
        'JPEG only supports 24bpp (8bpp grayscale is not currently handled)
        Case PDIF_JPEG
            GetClosestColorDepth = 24
        
        'JXR (JPEG XR) supports 24/32bpp
        Case PDIF_JXR
            If desiredColorDepth <= 24 Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
        'PNG supports all available color depths
        Case PDIF_PNG
        
            If desiredColorDepth <= 1 Then
                GetClosestColorDepth = 1
            ElseIf desiredColorDepth <= 4 Then
                GetClosestColorDepth = 4
            ElseIf desiredColorDepth <= 8 Then
                GetClosestColorDepth = 8
            ElseIf desiredColorDepth <= 24 Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
        'PPM only supports 24bpp
        Case PDIF_PPM
            GetClosestColorDepth = 24
        
        'PSD supports all available color depths
        Case PDIF_PSD
        
            If desiredColorDepth <= 1 Then
                GetClosestColorDepth = 1
            ElseIf desiredColorDepth <= 4 Then
                GetClosestColorDepth = 4
            ElseIf desiredColorDepth <= 8 Then
                GetClosestColorDepth = 8
            ElseIf desiredColorDepth <= 24 Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
        'TGA supports 8/24/32
        Case PDIF_TARGA
            If desiredColorDepth <= 8 Then
                GetClosestColorDepth = 8
            ElseIf desiredColorDepth <= 24 Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
        'TIFF supports all available color depths
        Case PDIF_TIFF
            
            If desiredColorDepth <= 1 Then
                GetClosestColorDepth = 1
            ElseIf desiredColorDepth <= 4 Then
                GetClosestColorDepth = 4
            ElseIf desiredColorDepth <= 8 Then
                GetClosestColorDepth = 8
            ElseIf desiredColorDepth <= 24 Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
            
        'WebP supports 24/32bpp
        Case PDIF_WEBP
            If desiredColorDepth <= 24 Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
    End Select
    
End Function

'Given an output PDIF, return the ideal metadata format for that image format.
Friend Function GetIdealMetadataFormatFromPDIF(ByVal outputPDIF As PD_IMAGE_FORMAT) As PD_METADATA_FORMAT

    Select Case outputPDIF
    
        Case PDIF_BMP
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        
        Case PDIF_GIF
            GetIdealMetadataFormatFromPDIF = PDMF_XMP
        
        Case PDIF_HDR
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        
        Case PDIF_JP2
            GetIdealMetadataFormatFromPDIF = PDMF_XMP
        
        Case PDIF_JPEG
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        
        Case PDIF_JXR
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        
        Case PDIF_PDI
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        
        Case PDIF_PNG
            GetIdealMetadataFormatFromPDIF = PDMF_XMP
        
        Case PDIF_PNM
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
            
        Case PDIF_PSD
            GetIdealMetadataFormatFromPDIF = PDMF_XMP
        
        Case PDIF_TARGA
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        
        Case PDIF_TIFF
            GetIdealMetadataFormatFromPDIF = PDMF_EXIF
        
        Case PDIF_WEBP
            GetIdealMetadataFormatFromPDIF = PDMF_XMP
        
        Case Else
            GetIdealMetadataFormatFromPDIF = PDMF_NONE
        
    End Select
    
End Function

'Given an output PDIF, return a BOOLEAN specifying whether Exif metadata is allowed for that image format.
' (Technically, ExifTool can write non-standard Exif chunks for formats like PNG and JPEG-2000, but PD prefers not to do this.
'  If an Exif tag can't be converted to a corresponding XMP tag, it should simply be removed from the new file.)
Friend Function IsExifAllowedForPDIF(ByVal outputPDIF As PD_IMAGE_FORMAT) As Boolean

    Select Case outputPDIF
    
        Case PDIF_BMP
            IsExifAllowedForPDIF = False
        
        Case PDIF_GIF
            IsExifAllowedForPDIF = False
        
        Case PDIF_HDR
            IsExifAllowedForPDIF = False
        
        Case PDIF_JP2
            IsExifAllowedForPDIF = False
        
        Case PDIF_JPEG
            IsExifAllowedForPDIF = True
        
        Case PDIF_JXR
            IsExifAllowedForPDIF = True
        
        Case PDIF_PDI
            IsExifAllowedForPDIF = True
        
        Case PDIF_PNG
            IsExifAllowedForPDIF = False
        
        Case PDIF_PNM
            IsExifAllowedForPDIF = False
        
        Case PDIF_PSD
            IsExifAllowedForPDIF = True
        
        Case PDIF_TARGA
            IsExifAllowedForPDIF = False
        
        Case PDIF_TIFF
            IsExifAllowedForPDIF = True
        
        Case PDIF_WEBP
            IsExifAllowedForPDIF = False
        
        Case Else
            IsExifAllowedForPDIF = False
        
    End Select
    
End Function

'Given an output PDIF, return a BOOLEAN specifying whether PD has implemented an export dialog for that image format.
Friend Function IsExportDialogSupported(ByVal outputPDIF As PD_IMAGE_FORMAT) As Boolean

    Select Case outputPDIF
    
        Case PDIF_BMP
            IsExportDialogSupported = True
        
        Case PDIF_GIF
            IsExportDialogSupported = True
        
        Case PDIF_HDR
            IsExportDialogSupported = False
        
        Case PDIF_JP2
            IsExportDialogSupported = True
        
        Case PDIF_JPEG
            IsExportDialogSupported = True
        
        Case PDIF_JXR
            IsExportDialogSupported = True
        
        Case PDIF_PDI
            IsExportDialogSupported = False
        
        Case PDIF_PNG
            IsExportDialogSupported = True
        
        Case PDIF_PBM, PDIF_PGM, PDIF_PNM, PDIF_PPM
            IsExportDialogSupported = True
        
        Case PDIF_PSD
            IsExportDialogSupported = False
        
        Case PDIF_TARGA
            IsExportDialogSupported = False
        
        Case PDIF_TIFF
            IsExportDialogSupported = True
        
        Case PDIF_WEBP
            IsExportDialogSupported = True
        
        Case Else
            IsExportDialogSupported = False
        
    End Select
    
End Function
