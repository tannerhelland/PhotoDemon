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
' used preferentially over FreeImage for reading and writing JPEGs.
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

'Array of common-dialog-formatted input/output filetypes
Private commonDialogInputs As String, commonDialogOutputs As String

'Common dialog also require a specialized "default extension" string for output files
Private cdOutputDefaultExtensions As String

'This value is used to track the number of currently used input/output formats
Private curFormatIndex As Long

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
    ReDim inputExtensions(0 To 100) As String
    ReDim inputDescriptions(0 To 100) As String

    'Formats should be added in alphabetical order, as this class has no "sort" functionality.

    'Always start with an "All Compatible Images" option
    inputDescriptions(0) = "All Compatible Images"
    
    'Unique to this first one is the full list of compatible extensions.  Instead of generating a full list here,
    ' it will be automatically generated as we go.
    
    'Set the location tracker to "0".  Beyond this point, it will be automatically updated.
    curFormatIndex = 0
    
    'Bitmap files require no plugins; they are always supported.
    addInputFormat "BMP - Windows or OS/2 Bitmap", "*.bmp"
    
    If FreeImageEnabled Then
        addInputFormat "DDS - DirectDraw Surface", "*.dds"
        addInputFormat "EXR - Industrial Light and Magic", "*.exr"
        addInputFormat "G3 - Digital Fax Format", "*.g3"
    End If
    
    addInputFormat "GIF - Compuserve", "*.gif"
    
    If FreeImageEnabled Then addInputFormat "HDR - High Dynamic Range", "*.hdr"
    
    addInputFormat "ICO - Windows Icon", "*.ico"
    
    If FreeImageEnabled Then
        addInputFormat "IFF - Amiga Interchange Format", "*.iff"
        addInputFormat "JNG - JPEG Network Graphics", "*.jng"
        addInputFormat "JP2/J2K - JPEG 2000 File or Codestream", "*.jp2;*.j2k;*.jpc;*.jpx;*.jpf"
    End If
    
    addInputFormat "JPG/JPEG - Joint Photographic Experts Group", "*.jpg;*.jpeg;*.jif;*.jfif"
    
    If FreeImageEnabled Then
        addInputFormat "KOA/KOALA - Commodore 64", "*.koa;*.koala"
        addInputFormat "LBM - Deluxe Paint", "*.lbm"
        addInputFormat "MNG - Multiple Network Graphics", "*.mng"
        addInputFormat "PBM - Portable Bitmap", "*.pbm"
        addInputFormat "PCD - Kodak PhotoCD", "*.pcd"
        addInputFormat "PCX - Zsoft Paintbrush", "*.pcx"
    End If
    
    'Only allow PDI (PhotoDemon's native file format) loading if the zLib dll has been properly detected
    If zLibEnabled Then addInputFormat "PDI - PhotoDemon Image", "*.pdi"
        
    If FreeImageEnabled Then
        addInputFormat "PGM - Portable Graymap", "*.pgm"
        addInputFormat "PIC/PICT - Macintosh Picture", "*.pict;*.pct;*.pic"
    End If
    
    'FreeImage or GDI+ works for loading PNGs
    If FreeImageEnabled Or GDIPlusEnabled Then addInputFormat "PNG - Portable Network Graphic", "*.png"
        
    If FreeImageEnabled Then
        addInputFormat "PPM - Portable Pixelmap", "*.ppm"
        addInputFormat "PSD - Adobe Photoshop", "*.psd"
        addInputFormat "RAS - Sun Raster File", "*.ras"
        addInputFormat "SGI/RGB/BW - Silicon Graphics Image", "*.sgi;*.rgb;*.rgba;*.bw;*.int;*.inta"
        addInputFormat "TGA - Truevision Targa", "*.tga"
    End If
    
    'FreeImage or GDI+ works for loading TIFFs
    If FreeImageEnabled Or GDIPlusEnabled Then addInputFormat "TIF/TIFF - Tagged Image File Format", "*.tif;*.tiff"
        
    If FreeImageEnabled Then addInputFormat "WBMP - Wireless Bitmap", "*.wbmp;*.wbm"
        
    'Finish out the list with an obligatory "All files" option
    addInputFormat "All files", "*.*"
    
    'Resize our description and extension arrays to match their final size
    numOfInputFormats = curFormatIndex
    ReDim Preserve inputDescriptions(0 To numOfInputFormats) As String
    ReDim Preserve inputExtensions(0 To numOfInputFormats) As String
    
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
Private Sub addInputFormat(ByVal formatDescription As String, ByVal extensionList As String)
    
    'Increment the counter
    curFormatIndex = curFormatIndex + 1
    
    'Add the descriptive text to our array
    inputDescriptions(curFormatIndex) = formatDescription
    
    'Add the extension list to our VB-specific array
    inputExtensions(curFormatIndex) = extensionList
    
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

    ReDim outputExtensions(0 To 100) As String
    ReDim outputDescriptions(0 To 100) As String

    'Formats should be added in alphabetical order, as this class has no "sort" functionality.
    
    'Start by effectively setting the location tracker to "0".  Beyond this point, it will be automatically updated.
    curFormatIndex = -1

    addOutputFormat "BMP - Windows Bitmap", "bmp"
    
    If FreeImageEnabled Or GDIPlusEnabled Then addOutputFormat "GIF - Graphics Interchange Format", "gif"
        
    If FreeImageEnabled Then addOutputFormat "JP2 - JPEG 2000", "jp2"
        
    If FreeImageEnabled Or GDIPlusEnabled Then addOutputFormat "JPG - Joint Photographic Experts Group", "jpg"
        
    If zLibEnabled Then addOutputFormat "PDI - PhotoDemon Image", "pdi"
    
    If FreeImageEnabled Or GDIPlusEnabled Then addOutputFormat "PNG - Portable Network Graphic", "png"
        
    If FreeImageEnabled Then
        addOutputFormat "PPM - Portable Pixel Map", "ppm"
        addOutputFormat "TGA - Truevision Targa", "tga"
    End If
    
    If FreeImageEnabled Or GDIPlusEnabled Then addOutputFormat "TIFF - Tagged Image File Format", "tif"
        
    'Resize our description and extension arrays to match their final size
    numOfOutputFormats = curFormatIndex
    ReDim Preserve outputDescriptions(0 To numOfOutputFormats) As String
    ReDim Preserve outputExtensions(0 To numOfOutputFormats) As String
    
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
Private Sub addOutputFormat(ByVal formatDescription As String, ByVal extensionList As String)
    
    'Increment the counter
    curFormatIndex = curFormatIndex + 1
    
    'Add the descriptive text to our array
    outputDescriptions(curFormatIndex) = formatDescription
    
    'Add the extension list to our VB-specific array
    outputExtensions(curFormatIndex) = extensionList
            
End Sub
