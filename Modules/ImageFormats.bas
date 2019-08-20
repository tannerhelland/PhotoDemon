Attribute VB_Name = "ImageFormats"
'***************************************************************************
'PhotoDemon Image Format Manager
'Copyright 2012-2019 by Tanner Helland
'Created: 18/November/12
'Last updated: 13/January/19
'Last update: add support for OpenRaster images
'
'This module determines run-time read/write support for various image formats.
'
'Based on available plugins, this class generates a list of file formats that PhotoDemon is capable of writing
' and reading.  These format lists are separately maintained, and the presence of a format in the Import category
' does not guarantee a similar presence in the Export category.
'
'Many esoteric formats rely on FreeImage.dll for loading and/or saving.  In some cases, GDI+ is used preferentially
' over FreeImage (e.g. loading JPEGs; FreeImage has better coverage of non-standard JPEG encodings, but GDI+ is
' significantly faster).  From this module alone, it won't be clear which plugin, if any, is used to load a given
' file - for that, you'd need to consult the relevant debug log after loading an image file.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Is the FreeImage DLL available?
Private m_FreeImageEnabled As Boolean

'Is pngQuant available?
Private m_pngQuantEnabled As Boolean

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
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        GetInputPDIF = inputPDIFs(dIndex)
    Else
        GetInputPDIF = FIF_UNKNOWN
    End If
End Function

'Return the friendly input format description at a given index
Public Function GetInputFormatDescription(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        GetInputFormatDescription = inputDescriptions(dIndex)
    Else
        GetInputFormatDescription = vbNullString
    End If
End Function

'Return the input format extension at a given index
Public Function GetInputFormatExtensions(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
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
    GetNumOfInputFormats = numOfInputFormats
End Function

'Return a list of input filetypes formatted for use with a common dialog box
Public Function GetCommonDialogInputFormats() As String
    GetCommonDialogInputFormats = m_CommonDialogInputs
End Function

'Return the index of given output FIF
Public Function GetIndexOfOutputPDIF(ByVal srcFIF As PD_IMAGE_FORMAT) As Long
    
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
Public Function GetOutputPDIF(ByVal dIndex As Long) As PD_IMAGE_FORMAT
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        GetOutputPDIF = outputPDIFs(dIndex)
    Else
        GetOutputPDIF = FIF_UNKNOWN
    End If
End Function

'Return the friendly output format description at a given index
Public Function GetOutputFormatDescription(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfOutputFormats) Then
        GetOutputFormatDescription = outputDescriptions(dIndex)
    Else
        GetOutputFormatDescription = vbNullString
    End If
End Function

'Return the output format extension at a given index
Public Function GetOutputFormatExtension(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfOutputFormats) Then
        GetOutputFormatExtension = outputExtensions(dIndex)
    Else
        GetOutputFormatExtension = vbNullString
    End If
End Function

'Return the number of available output format types
Public Function GetNumOfOutputFormats() As Long
    GetNumOfOutputFormats = numOfOutputFormats
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

    'Prepare a list of possible INPUT formats based on the plugins available to us.
    ' (These format lists are automatically trimmed after plugin status has been assessed;
    '  the arbitrary upper limit of "50" would only need to be revisited if we greatly
    '  expand format support in the future.)
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
    
    If m_FreeImageEnabled Then
        AddInputFormat "DDS - DirectDraw Surface", "*.dds", PDIF_DDS
        AddInputFormat "DNG - Adobe Digital Negative", "*.dng", PDIF_RAW
    End If
    
    'EMFs will be loaded via GDI+ for improved rendering and feature compatibility
    AddInputFormat "EMF - Enhanced Metafile", "*.emf", PDIF_EMF
    
    If m_FreeImageEnabled Then
        AddInputFormat "EXR - Industrial Light and Magic", "*.exr", PDIF_EXR
        AddInputFormat "G3 - Digital Fax Format", "*.g3", PDIF_FAXG3
    End If
    
    AddInputFormat "GIF - Graphics Interchange Format", "*.gif", PDIF_GIF
    
    If m_FreeImageEnabled Then AddInputFormat "HDR - High Dynamic Range", "*.hdr", PDIF_HDR
    
    AddInputFormat "ICO - Windows Icon", "*.ico", PDIF_ICO
    
    If m_FreeImageEnabled Then
        AddInputFormat "IFF - Amiga Interchange Format", "*.iff", PDIF_IFF
        AddInputFormat "JNG - JPEG Network Graphics", "*.jng", PDIF_JNG
        AddInputFormat "JP2/J2K - JPEG 2000 File or Codestream", "*.jp2;*.j2k;*.jpc;*.jpx;*.jpf", PDIF_JP2
    End If
    
    AddInputFormat "JPG/JPEG - Joint Photographic Experts Group", "*.jpg;*.jpeg;*.jpe;*.jif;*.jfif", PDIF_JPEG
    
    If m_FreeImageEnabled Then
        AddInputFormat "JXR/HDP - JPEG XR (HD Photo)", "*.jxr;*.hdp;*.wdp", PDIF_JXR
        AddInputFormat "KOA/KOALA - Commodore 64", "*.koa;*.koala", PDIF_KOALA
        AddInputFormat "LBM - Deluxe Paint", "*.lbm", PDIF_LBM
        AddInputFormat "MNG - Multiple Network Graphics", "*.mng", PDIF_MNG
    End If
    
    AddInputFormat "ORA - OpenRaster", "*.ora", PDIF_ORA
    
    If m_FreeImageEnabled Then
        AddInputFormat "PBM - Portable Bitmap", "*.pbm", PDIF_PBM
        AddInputFormat "PCD - Kodak PhotoCD", "*.pcd", PDIF_PCD
        AddInputFormat "PCX - Zsoft Paintbrush", "*.pcx", PDIF_PCX
    End If
    
    'PDI (PhotoDemon's native file format) is always available!
    AddInputFormat "PDI - PhotoDemon Image", "*.pdi", PDIF_PDI
        
    If m_FreeImageEnabled Then
        AddInputFormat "PFM - Portable Floatmap", "*.pfm", PDIF_PFM
        AddInputFormat "PGM - Portable Graymap", "*.pgm", PDIF_PGM
        AddInputFormat "PIC/PICT - Macintosh Picture", "*.pict;*.pct;*.pic", PDIF_PICT
    End If
    
    'We actually have three PNG decoders: an internal one (preferred), FreeImage, and GDI+
    AddInputFormat "PNG/APNG - Portable Network Graphic", "*.png;*.apng", PDIF_PNG
        
    If m_FreeImageEnabled Then
        AddInputFormat "PNM - Portable Anymap", "*.pnm", PDIF_PPM
        AddInputFormat "PPM - Portable Pixmap", "*.ppm", PDIF_PPM
        AddInputFormat "PSD - Adobe Photoshop", "*.psd;*.psb", PDIF_PSD
        AddInputFormat "RAS - Sun Raster File", "*.ras", PDIF_RAS
        AddInputFormat "RAW, etc - Raw image data", "*.3fr;*.arw;*.bay;*.bmq;*.cap;*.cine;*.cr2;*.crw;*.cs1;*.dc2;*.dcr;*.dng;*.drf;*.dsc;*.erf;*.fff;*.ia;*.iiq;*.k25;*.kc2;*.kdc;*.mdc;*.mef;*.mos;*.mrw;*.nef;*.nrw;*.orf;*.pef;*.ptx;*.pxn;*.qtk;*.raf;*.raw;*.rdc;*.rw2;*.rwz;*.sr2;*.srf;*.sti", PDIF_RAW
        AddInputFormat "SGI/RGB/BW - Silicon Graphics Image", "*.sgi;*.rgb;*.rgba;*.bw;*.int;*.inta", PDIF_SGI
        AddInputFormat "TGA - Truevision (TARGA)", "*.tga", PDIF_TARGA
    End If
    
    'FreeImage or GDI+ works for loading TIFFs
    AddInputFormat "TIF/TIFF - Tagged Image File Format", "*.tif;*.tiff", PDIF_TIFF
        
    If m_FreeImageEnabled Then
        AddInputFormat "WBMP - Wireless Bitmap", "*.wbmp;*.wbm", PDIF_WBMP
        AddInputFormat "WEBP - Google WebP", "*.webp", PDIF_WEBP
    End If
    
    'I don't know if anyone still uses WMFs, but GDI+ provides support "for free"
    AddInputFormat "WMF - Windows Metafile", "*.wmf", PDIF_WMF
    
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
Public Sub GenerateOutputFormats()

    ReDim outputExtensions(0 To 50) As String
    ReDim outputDescriptions(0 To 50) As String
    ReDim outputPDIFs(0 To 50) As PD_IMAGE_FORMAT

    'Formats should be added in alphabetical order, as this class has no "sort" functionality.
    
    'Start by effectively setting the location tracker to "0".  Beyond this point, it will be automatically updated.
    m_curFormatIndex = -1

    AddOutputFormat "BMP - Windows Bitmap", "bmp", PDIF_BMP
    
    'FreeImage or GDI+ can write GIFs for us
    AddOutputFormat "GIF - Graphics Interchange Format", "gif", PDIF_GIF
    
    If m_FreeImageEnabled Then
        AddOutputFormat "HDR - High Dynamic Range", "hdr", PDIF_HDR
        AddOutputFormat "JP2 - JPEG 2000", "jp2", PDIF_JP2
    End If
        
    'FreeImage or GDI+ can write JPEGs for us
    AddOutputFormat "JPG - Joint Photographic Experts Group", "jpg", PDIF_JPEG
        
    If m_FreeImageEnabled Then AddOutputFormat "JXR - JPEG XR (HD Photo)", "jxr", PDIF_JXR
    
    AddOutputFormat "ORA - OpenRaster", "ora", PDIF_ORA
    AddOutputFormat "PDI - PhotoDemon Image", "pdi", PDIF_PDI
    
    'FreeImage or GDI+ can write PNGs for us
    AddOutputFormat "PNG - Portable Network Graphic", "png", PDIF_PNG
    
    If m_FreeImageEnabled Then
        AddOutputFormat "PNM - Portable Anymap (Netpbm)", "pnm", PDIF_PNM
        AddOutputFormat "PSD - Photoshop Document", "psd", PDIF_PSD
        AddOutputFormat "TGA - Truevision (TARGA)", "tga", PDIF_TARGA
    End If
    
    'FreeImage or GDI+ can write TIFFs for us
    AddOutputFormat "TIFF - Tagged Image File Format", "tif", PDIF_TIFF
    
    If m_FreeImageEnabled Then AddOutputFormat "WEBP - Google WebP", "webp", PDIF_WEBP
    
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
Public Function GetExtensionFromPDIF(ByVal srcPDIF As PD_IMAGE_FORMAT) As String

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
Public Function DoesPDIFSupportMultipleColorDepths(ByVal outputPDIF As PD_IMAGE_FORMAT) As Boolean

    Select Case outputPDIF
    
        Case PDIF_GIF, PDIF_HDR, PDIF_JPEG, PDIF_ORA
            DoesPDIFSupportMultipleColorDepths = False
            
        Case Else
            DoesPDIFSupportMultipleColorDepths = True
    
    End Select

End Function

'Given a file format and color depth, are the two compatible?  (NOTE: this function takes into account the availability of FreeImage and/or GDI+)
Public Function IsColorDepthSupported(ByVal outputPDIF As Long, ByVal desiredColorDepth As Long) As Boolean
    
    'Internal engines are covered first; in the absence of these, we'll rely on feature sets in
    ' either FreeImage (if available) or GDI+
    
    'Check the special case of PDI (internal PhotoDemon images)
    If (outputPDIF = PDIF_PDI) Then
        IsColorDepthSupported = True
        Exit Function
    End If
    
    'OpenRaster support uses an internal engine
    If (outputPDIF = PDIF_ORA) Then
        IsColorDepthSupported = (desiredColorDepth = 32)
        Exit Function
    End If
    
    'All subsequent checks rely on FreeImage or GDI+

    'By default, report that a given color depth is NOT supported
    IsColorDepthSupported = False
    
    'First, address formats handled only by FreeImage
    If m_FreeImageEnabled Then
        
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
        
    'If FreeImage isn't available, fall back to GDI+ features
    Else
    
        Select Case outputPDIF
        
            'GIF
            Case PDIF_GIF
                IsColorDepthSupported = (desiredColorDepth = 8)
                
            'JPEG
            Case PDIF_JPEG
                IsColorDepthSupported = (desiredColorDepth = 24)
                
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
     
    End If

End Function

'Given a file format and desired color depth, return the next-best color depth that can be used (assuming the desired one is not available)
' (NOTE: this function takes into account the availability of FreeImage and/or GDI+)
Public Function GetClosestColorDepth(ByVal outputPDIF As PD_IMAGE_FORMAT, ByVal desiredColorDepth As Long) As Long
    
    'Internal export engines are handled first, as they tend to have the most comprehensive
    ' (and reliable) color depth coverage
    If (outputPDIF = PDIF_PDI) Then
        GetClosestColorDepth = desiredColorDepth
        Exit Function
    End If
    
    If (outputPDIF = PDIF_ORA) Then
        GetClosestColorDepth = 32
        Exit Function
    End If
    
    'Subsequent formats are covered by either FreeImage or GDI+

    'By default, report that 24bpp is the preferred alternative
    GetClosestColorDepth = 24
    
    'Certain file formats only support one output color depth, so they are easily handled (e.g. GIF)
    
    'Some file formats support many color depths (PNG, for example, can handle 1/4/8/24/32)
    
    'This function attempts to return the color depth nearest to the one the user has requested
    Select Case outputPDIF
    
        'BMP support changes based on the available encoder
        Case PDIF_BMP
        
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
            
        'GIF only supports 8bpp
        Case PDIF_GIF
            GetClosestColorDepth = 8
            
        'HDR only supports 24bpp
        Case PDIF_HDR
            GetClosestColorDepth = 24
            
        'JP2 (JPEG 2000) supports 24/32bpp
        Case PDIF_JP2
            If (desiredColorDepth <= 24) Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
        'JPEG only supports 24bpp (8bpp grayscale is currently handled automatically by the encoder)
        Case PDIF_JPEG
            GetClosestColorDepth = 24
        
        'JXR (JPEG XR) supports 24/32bpp
        Case PDIF_JXR
            If (desiredColorDepth <= 24) Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
        Case PDIF_ORA
            GetClosestColorDepth = 32
            
        'PNG supports all available color depths
        Case PDIF_PNG
        
            If (desiredColorDepth <= 1) Then
                GetClosestColorDepth = 1
            ElseIf (desiredColorDepth <= 4) Then
                GetClosestColorDepth = 4
            ElseIf (desiredColorDepth <= 8) Then
                GetClosestColorDepth = 8
            ElseIf (desiredColorDepth <= 24) Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
        'PPM only supports 24bpp
        Case PDIF_PPM
            GetClosestColorDepth = 24
        
        'PSD supports all available color depths
        Case PDIF_PSD
        
            If (desiredColorDepth <= 1) Then
                GetClosestColorDepth = 1
            ElseIf (desiredColorDepth <= 4) Then
                GetClosestColorDepth = 4
            ElseIf (desiredColorDepth <= 8) Then
                GetClosestColorDepth = 8
            ElseIf (desiredColorDepth <= 24) Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
        'TGA supports 8/24/32
        Case PDIF_TARGA
            If (desiredColorDepth <= 8) Then
                GetClosestColorDepth = 8
            ElseIf (desiredColorDepth <= 24) Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
        
        'TIFF supports all available color depths
        Case PDIF_TIFF
            If (desiredColorDepth <= 1) Then
                GetClosestColorDepth = 1
            ElseIf (desiredColorDepth <= 4) Then
                GetClosestColorDepth = 4
            ElseIf (desiredColorDepth <= 8) Then
                GetClosestColorDepth = 8
            ElseIf (desiredColorDepth <= 24) Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
            End If
            
        'WebP supports 24/32bpp
        Case PDIF_WEBP
            If (desiredColorDepth <= 24) Then
                GetClosestColorDepth = 24
            Else
                GetClosestColorDepth = 32
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
Public Function IsExifAllowedForPDIF(ByVal outputPDIF As PD_IMAGE_FORMAT) As Boolean

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
        
        Case PDIF_ORA
            IsExifAllowedForPDIF = False
            
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
Public Function IsExportDialogSupported(ByVal outputPDIF As PD_IMAGE_FORMAT) As Boolean

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
        
        Case PDIF_ORA
            IsExportDialogSupported = False
        
        Case PDIF_PDI
            IsExportDialogSupported = False
        
        Case PDIF_PNG
            IsExportDialogSupported = True
        
        Case PDIF_PBM, PDIF_PGM, PDIF_PNM, PDIF_PPM
            IsExportDialogSupported = True
        
        Case PDIF_PSD
            IsExportDialogSupported = True
        
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

Public Function IsExifToolRelevant(ByVal srcFormat As PD_IMAGE_FORMAT) As Boolean

    Select Case srcFormat
    
        Case PDIF_PDI
            IsExifToolRelevant = False
            
        Case PDIF_ORA
            IsExifToolRelevant = False
            
        Case Else
            IsExifToolRelevant = True
    
    End Select

End Function

Public Function IsFreeImageEnabled() As Boolean
    IsFreeImageEnabled = m_FreeImageEnabled
End Function

Public Sub SetFreeImageEnabled(ByVal newState As Boolean)
    m_FreeImageEnabled = newState
End Sub

Public Function IsPngQuantEnabled() As Boolean
    IsPngQuantEnabled = m_pngQuantEnabled
End Function

Public Sub SetPngQuantEnabled(ByVal newState As Boolean)
    m_pngQuantEnabled = newState
End Sub

'When the active language changes, we need to calculate new translations for text like "All Compatible Images"
Public Sub NotifyLanguageChanged()
    GenerateInputFormats
    GenerateOutputFormats
End Sub
