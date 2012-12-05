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
'Copyright ©2011-2012 by Tanner Helland
'Created: 18/November/12
'Last updated: 26/November/12
'Last update: added Industrial Light and Magic EXR format support
'
'This class is the new handler for PhotoDemon's vast image format support library.  Previously, individual functions
' (such as batch convert) had to manage specialized versions of this data, which made it extremely cumbersome to add
' or adjust PhotoDemon's support for individual formats.  To remedy this problem, this class was created.
'
'Based on the available plugins, this class generates a list of file formats that PhotoDemon is capable of writing
' and reading.  These format lists are separately maintained, and the presence of a format in one column does not
' have any connection to the other.
'
'Most non-standard formats rely on FreeImage for loading and/or saving, and they can additionally be tested against
' individual FreeImage version numbers if necessary.  As of FreeImage 3.15.4 (released October 2012) no formats are
' disabled due to known FreeImage bugs.
'
'GDIPlus is optionally used as a fallback if FreeImage is not present, but it should be noted that GDI+ interoperability
' is not as well-tested as FreeImage.  In some cases, however, it IS faster to use GDI+ due to the need to make a
' specialized copy of DIB data for FreeImage's internal handling - so in the case of batch conversions, GDI+ is actually
' used preferentially over FreeImage for reading and writing certain filetypes.
'
'***************************************************************************

Option Explicit

'Is FreeImage.dll enabled?
Public FreeImageEnabled As Boolean

'Is GDI+ available?
Public GDIPlusEnabled As Boolean

'Number of available input, output formats
Private numOfInputFormats As Long, numOfOutputFormats As Long

'Array of available input, output extensions.
Private inputExtensions() As String
Private outputExtensions() As String

'Array of "friendly" descriptions for input, output formats
Private inputDescriptions() As String
Private outputDescriptions() As String

'Array of corresponding image format constants (using the FreeImage convention)
Private inputFIFs() As Long
Private outputFIFs() As Long

'Array of common-dialog-formatted input/output filetypes
Private commonDialogInputs As String, commonDialogOutputs As String

'Common dialog also require a specialized "default extension" string for output files
Private cdOutputDefaultExtensions As String

'This value is used to track the number of currently used input/output formats
Private curFormatIndex As Long

'Return the index of given input FIF
Public Function getIndexOfInputFIF(ByVal srcFIF As Long) As Long
    
    Dim i As Long
    For i = 0 To getNumOfInputFormats
        If inputFIFs(i) = srcFIF Then
            getIndexOfInputFIF = i
            Exit Function
        End If
    Next i
    
    'If we reach this line of code, no match was found.  Return -1.
    getIndexOfInputFIF = -1
    
End Function

'Return the FIF (image format constant) at a given index
Public Function getInputFIF(ByVal dIndex As Long) As Long
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        getInputFIF = inputFIFs(dIndex)
    Else
        getInputFIF = FIF_UNKNOWN
    End If
End Function

'Return the friendly input format description at a given index
Public Function getInputFormatDescription(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        getInputFormatDescription = inputDescriptions(dIndex)
    Else
        getInputFormatDescription = ""
    End If
End Function

'Return the input format extension at a given index
Public Function getInputFormatExtensions(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        getInputFormatExtensions = inputExtensions(dIndex)
    Else
        getInputFormatExtensions = ""
    End If
End Function

'Return the number of available input format types
Public Function getNumOfInputFormats() As Long
    getNumOfInputFormats = numOfInputFormats
End Function

'Return a list of input filetypes formatted for use with a common dialog box
Public Function getCommonDialogInputFormats() As String
    getCommonDialogInputFormats = commonDialogInputs
End Function

'Return the index of given output FIF
Public Function getIndexOfOutputFIF(ByVal srcFIF As Long) As Long
    
    Dim i As Long
    For i = 0 To getNumOfOutputFormats
        If outputFIFs(i) = srcFIF Then
            getIndexOfOutputFIF = i
            Exit Function
        End If
    Next i
    
    'If we reach this line of code, no match was found.  Return -1.
    getIndexOfOutputFIF = -1
    
End Function

'Return the FIF (image format constant) at a given index
Public Function getOutputFIF(ByVal dIndex As Long) As Long
    If (dIndex >= 0) And (dIndex <= numOfInputFormats) Then
        getOutputFIF = outputFIFs(dIndex)
    Else
        getOutputFIF = FIF_UNKNOWN
    End If
End Function

'Return the friendly output format description at a given index
Public Function getOutputFormatDescription(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfOutputFormats) Then
        getOutputFormatDescription = outputDescriptions(dIndex)
    Else
        getOutputFormatDescription = ""
    End If
End Function

'Return the output format extension at a given index
Public Function getOutputFormatExtension(ByVal dIndex As Long) As String
    If (dIndex >= 0) And (dIndex <= numOfOutputFormats) Then
        getOutputFormatExtension = outputExtensions(dIndex)
    Else
        getOutputFormatExtension = ""
    End If
End Function

'Return the number of available output format types
Public Function getNumOfOutputFormats() As Long
    getNumOfOutputFormats = numOfOutputFormats
End Function

'Return a list of output filetypes formatted for use with a common dialog box
Public Function getCommonDialogOutputFormats() As String
    getCommonDialogOutputFormats = commonDialogOutputs
End Function

'Return a list of output default extensions formatted for use with a common dialog box
Public Function getCommonDialogDefaultExtensions() As String
    getCommonDialogDefaultExtensions = cdOutputDefaultExtensions
End Function

'Generate a list of available import formats
Public Sub generateInputFormats()

    'Prepare a list of possible INPUT formats based on the plugins available to us
    ReDim inputExtensions(0 To 50) As String
    ReDim inputDescriptions(0 To 50) As String
    ReDim inputFIFs(0 To 50) As Long

    'Formats should be added in alphabetical order, as this class has no "sort" functionality.

    'Always start with an "All Compatible Images" option
    inputDescriptions(0) = "All Compatible Images"
    
    'Unique to this first one is the full list of compatible extensions.  Instead of generating a full list here,
    ' it will be automatically generated as we go.
    
    'Set the location tracker to "0".  Beyond this point, it will be automatically updated.
    curFormatIndex = 0
    
    'Bitmap files require no plugins; they are always supported.
    addInputFormat "BMP - Windows or OS/2 Bitmap", "*.bmp", FIF_BMP
    
    If FreeImageEnabled Then
        addInputFormat "DDS - DirectDraw Surface", "*.dds", FIF_DDS
        addInputFormat "EXR - Industrial Light and Magic", "*.exr", FIF_EXR
        addInputFormat "G3 - Digital Fax Format", "*.g3", FIF_FAXG3
    End If
    
    addInputFormat "GIF - Compuserve", "*.gif", FIF_GIF
    
    If FreeImageEnabled Then addInputFormat "HDR - High Dynamic Range", "*.hdr", FIF_HDR
    
    addInputFormat "ICO - Windows Icon", "*.ico", FIF_ICO
    
    If FreeImageEnabled Then
        addInputFormat "IFF - Amiga Interchange Format", "*.iff", FIF_IFF
        addInputFormat "JNG - JPEG Network Graphics", "*.jng", FIF_JNG
        addInputFormat "JP2/J2K - JPEG 2000 File or Codestream", "*.jp2;*.j2k", FIF_JP2
        ' At present, the batch convert box gets overwhelmed by too many file formats - so I've shrunk the list to cut a few rare JPG2000 extensions.
        ' addInputFormat "JP2/J2K - JPEG 2000 File or Codestream", "*.jp2;*.j2k;*.jpc;*.jpx;*.jpf", FIF_JP2
    End If
    
    addInputFormat "JPG/JPEG - Joint Photographic Experts Group", "*.jpg;*.jpeg;*.jif;*.jfif", FIF_JPEG
    
    If FreeImageEnabled Then
        addInputFormat "KOA/KOALA - Commodore 64", "*.koa;*.koala", FIF_KOALA
        addInputFormat "LBM - Deluxe Paint", "*.lbm", FIF_LBM
        addInputFormat "MNG - Multiple Network Graphics", "*.mng", FIF_MNG
        addInputFormat "PBM - Portable Bitmap", "*.pbm", FIF_PBM
        addInputFormat "PCD - Kodak PhotoCD", "*.pcd", FIF_PCD
        addInputFormat "PCX - Zsoft Paintbrush", "*.pcx", FIF_PCX
    End If
    
    'Only allow PDI (PhotoDemon's native file format) loading if the zLib dll has been properly detected
    If zLibEnabled Then addInputFormat "PDI - PhotoDemon Image", "*.pdi", 100
        
    If FreeImageEnabled Then
        addInputFormat "PGM - Portable Graymap", "*.pgm", FIF_PGM
        addInputFormat "PIC/PICT - Macintosh Picture", "*.pict;*.pct;*.pic", FIF_PICT
    End If
    
    'FreeImage or GDI+ works for loading PNGs
    If FreeImageEnabled Or GDIPlusEnabled Then addInputFormat "PNG - Portable Network Graphic", "*.png", FIF_PNG
        
    If FreeImageEnabled Then
        addInputFormat "PPM - Portable Pixelmap", "*.ppm", FIF_PPM
        addInputFormat "PSD - Adobe Photoshop", "*.psd", FIF_PSD
        addInputFormat "RAS - Sun Raster File", "*.ras", FIF_RAS
        addInputFormat "SGI/RGB/BW - Silicon Graphics Image", "*.sgi;*.rgb;*.rgba;*.bw;*.int;*.inta", FIF_SGI
        addInputFormat "TGA - Truevision Targa", "*.tga", FIF_TARGA
    End If
    
    'FreeImage or GDI+ works for loading TIFFs
    If FreeImageEnabled Or GDIPlusEnabled Then addInputFormat "TIF/TIFF - Tagged Image File Format", "*.tif;*.tiff", FIF_TIFF
        
    If FreeImageEnabled Then addInputFormat "WBMP - Wireless Bitmap", "*.wbmp;*.wbm", FIF_WBMP
        
    'Finish out the list with an obligatory "All files" option
    addInputFormat "All files", "*.*", -1
    
    'Resize our description and extension arrays to match their final size
    numOfInputFormats = curFormatIndex
    ReDim Preserve inputDescriptions(0 To numOfInputFormats) As String
    ReDim Preserve inputExtensions(0 To numOfInputFormats) As String
    ReDim Preserve inputFIFs(0 To numOfInputFormats) As Long
    
    'Now that all input files have been added, we can compile a common-dialog-friendly version of this index
    
    'Loop through each entry in the arrays, and append them to the common-dialog-formatted string
    Dim x As Long
    For x = 0 To numOfInputFormats
    
        'Index 0 is a special case; everything else is handled in the same manner.
        If x <> 0 Then
            commonDialogInputs = commonDialogInputs & "|" & inputDescriptions(x) & "|" & inputExtensions(x)
        Else
            commonDialogInputs = inputDescriptions(x) & "|" & inputExtensions(x)
        End If
    
    Next x
    
    'Input format generation complete!
    
End Sub

'Add support for another input format.  A descriptive string and extension list are required.
Private Sub addInputFormat(ByVal formatDescription As String, ByVal extensionList As String, ByVal correspondingFIF As Long)
    
    'Increment the counter
    curFormatIndex = curFormatIndex + 1
    
    'Add the descriptive text to our collection
    inputDescriptions(curFormatIndex) = formatDescription
    
    'Add any relevant extension(s) to our collection
    inputExtensions(curFormatIndex) = extensionList
    
    'Add the FIF constant
    inputFIFs(curFormatIndex) = correspondingFIF
    
    'If applicable, add these extensions to the "All Compatible Images" list
    If extensionList <> "*.*" Then
        If curFormatIndex <> 1 Then
            inputExtensions(0) = inputExtensions(0) & ";" & extensionList
        Else
            inputExtensions(0) = inputExtensions(0) & extensionList
        End If
    End If
            
End Sub

'Generate a list of available export formats
Public Sub generateOutputFormats()

    ReDim outputExtensions(0 To 50) As String
    ReDim outputDescriptions(0 To 50) As String
    ReDim outputFIFs(0 To 50) As Long

    'Formats should be added in alphabetical order, as this class has no "sort" functionality.
    
    'Start by effectively setting the location tracker to "0".  Beyond this point, it will be automatically updated.
    curFormatIndex = -1

    addOutputFormat "BMP - Windows Bitmap", "bmp", FIF_BMP
    
    If FreeImageEnabled Or GDIPlusEnabled Then addOutputFormat "GIF - Graphics Interchange Format", "gif", FIF_GIF
        
    If FreeImageEnabled Then addOutputFormat "JP2 - JPEG 2000", "jp2", FIF_JP2
        
    If FreeImageEnabled Or GDIPlusEnabled Then addOutputFormat "JPG - Joint Photographic Experts Group", "jpg", FIF_JPEG
        
    If zLibEnabled Then addOutputFormat "PDI - PhotoDemon Image", "pdi", 100
    
    If FreeImageEnabled Or GDIPlusEnabled Then addOutputFormat "PNG - Portable Network Graphic", "png", FIF_PNG
        
    If FreeImageEnabled Then
        addOutputFormat "PPM - Portable Pixel Map", "ppm", FIF_PPM
        addOutputFormat "TGA - Truevision Targa", "tga", FIF_TARGA
    End If
    
    If FreeImageEnabled Or GDIPlusEnabled Then addOutputFormat "TIFF - Tagged Image File Format", "tif", FIF_TIFF
        
    'Resize our description and extension arrays to match their final size
    numOfOutputFormats = curFormatIndex
    ReDim Preserve outputDescriptions(0 To numOfOutputFormats) As String
    ReDim Preserve outputExtensions(0 To numOfOutputFormats) As String
    ReDim Preserve outputFIFs(0 To numOfOutputFormats) As Long
    
    'Now that all output files have been added, we can compile a common-dialog-friendly version of this index
    
    'Loop through each entry in the arrays, and append them to the common-dialog-formatted string
    Dim x As Long
    For x = 0 To numOfOutputFormats
    
        'Index 0 is a special case; everything else is handled in the same manner.
        If x <> 0 Then
            commonDialogOutputs = commonDialogOutputs & "|" & outputDescriptions(x) & "|*." & outputExtensions(x)
            cdOutputDefaultExtensions = cdOutputDefaultExtensions & "|." & outputExtensions(x)
        Else
            commonDialogOutputs = outputDescriptions(x) & "|*." & outputExtensions(x)
            cdOutputDefaultExtensions = "." & outputExtensions(x)
        End If
    
    Next x
    
    'Output format generation complete!
        
End Sub

'Add support for another output format.  A descriptive string and extension list are required.
Private Sub addOutputFormat(ByVal formatDescription As String, ByVal extensionList As String, ByVal correspondingFIF As Long)
    
    'Increment the counter
    curFormatIndex = curFormatIndex + 1
    
    'Add the descriptive text to our collection
    outputDescriptions(curFormatIndex) = formatDescription
    
    'Add the primary extension for this format type
    outputExtensions(curFormatIndex) = extensionList
    
    'Add the corresponding FIF
    outputFIFs(curFormatIndex) = correspondingFIF
            
End Sub

'Given an FIF (image format constant), return the default extension.
Public Function getExtensionFromFIF(ByVal srcFIF As Long) As String

    Select Case srcFIF
    
        Case FIF_BMP
            getExtensionFromFIF = "bmp"
        Case FIF_ICO
            getExtensionFromFIF = "ico"
        Case FIF_JPEG
            getExtensionFromFIF = "jpg"
        Case FIF_JNG
            getExtensionFromFIF = "jng"
        Case FIF_KOALA
            getExtensionFromFIF = "koa"
        Case FIF_LBM
            getExtensionFromFIF = "lbm"
        Case FIF_IFF
            getExtensionFromFIF = "iff"
        Case FIF_MNG
            getExtensionFromFIF = "mng"
        Case FIF_PBM
            getExtensionFromFIF = "pbm"
        Case FIF_PBMRAW
            getExtensionFromFIF = "pbm"
        Case FIF_PCD
            getExtensionFromFIF = "pcd"
        Case FIF_PCX
            getExtensionFromFIF = "pcx"
        Case FIF_PGM
            getExtensionFromFIF = "pgm"
        Case FIF_PGMRAW
            getExtensionFromFIF = "pgm"
        Case FIF_PNG
            getExtensionFromFIF = "png"
        Case FIF_PPM
            getExtensionFromFIF = "ppm"
        Case FIF_PPMRAW
            getExtensionFromFIF = "ppm"
        Case FIF_RAS
            getExtensionFromFIF = "ras"
        Case FIF_TARGA
            getExtensionFromFIF = "tga"
        Case FIF_TIFF
            getExtensionFromFIF = "tif"
        Case FIF_WBMP
            getExtensionFromFIF = "wbm"
        Case FIF_PSD
            getExtensionFromFIF = "psd"
        Case FIF_CUT
            getExtensionFromFIF = "cut"
        Case FIF_XBM
            getExtensionFromFIF = "xbm"
        Case FIF_XPM
            getExtensionFromFIF = "xpm"
        Case FIF_DDS
            getExtensionFromFIF = "dds"
        Case FIF_GIF
            getExtensionFromFIF = "gif"
        Case FIF_HDR
            getExtensionFromFIF = "hdr"
        Case FIF_FAXG3
            getExtensionFromFIF = "g3"
        Case FIF_SGI
            getExtensionFromFIF = "sgi"
        Case FIF_EXR
            getExtensionFromFIF = "exr"
        Case FIF_J2K
            getExtensionFromFIF = "j2k"
        Case FIF_JP2
            getExtensionFromFIF = "jp2"
        Case FIF_PFM
            getExtensionFromFIF = "pfm"
        Case FIF_PICT
            getExtensionFromFIF = "pct"
        Case FIF_RAW = 34
            getExtensionFromFIF = "raw"
        Case Else
            getExtensionFromFIF = ""
    
    End Select

End Function
