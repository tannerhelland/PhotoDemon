VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon ICO (Icon) Container and Import/Export engine
'Copyright 2020-2020 by Tanner Helland
'Created: 07/May/20
'Last updated: 11/May/20
'Last update: wrap up initial work on ICO import; export is up next!
'
'This class (and its associated pdICO- child classes) handle import and export of
' old-school ICO files.  The class was custom-built for PhotoDemon, with an emphasis on
' performance, robustness, and automatic coverage of core ICO features - e.g. users
' shouldn't need to know any details about the ICO format to produce useable ICO files.
'
'Unless otherwise noted, all code in this class is my original work.  I've based my work
' off info found in the following articles (links good as of May 2020):
' https://devblogs.microsoft.com/oldnewthing/20101018-00/?p=12513
' https://docs.microsoft.com/en-us/previous-versions/ms997538(v=msdn.10)
' https://docs.microsoft.com/en-us/windows/win32/uxguide/vis-icons
' https://en.wikipedia.org/wiki/ICO_(file_format)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps all kinds of
' diagnostic information to the debug log.
Private Const ICO_DEBUG_VERBOSE As Boolean = True

'ICO parsing is complicated, and a lot of things can go wrong.  Instead of returning
' binary "success/fail" values, we return specific flags; "warnings" may be recoverable
' and you can still attempt to load the file.  "Failure" returns are unrecoverable and
' processing *must* be abandoned.  (As a convenience, you can treat the "warning" and
' "failure" values as flags; specific warning/failure states in each category will share
' the same high flag bit.)
'
'As I get deeper into this class, I may expand this enum to include more detailed states.
Public Enum PD_ICOResult
    ico_Success = 0
    ico_Warning = 256
    ico_Failure = 65536
    ico_FileNotICO = 16777217
End Enum

#If False Then
    Private Const ico_Success = 0, ico_Warning = 256, ico_Failure = 65536, ico_FileNotICO = 16777217
#End If

'Icons and cursors are both supported, although cursor hotspots are not currently preserved
Private Enum PD_IconSubtype
    ico_Unknown = 0
    ico_Icon = 1
    ico_Cursor = 2
End Enum

#If False Then
    Private Const ico_Unknown = 0, ico_Icon = 1, ico_Cursor = 2
#End If

Private m_IconSubtype As PD_IconSubtype

'Icon bitmaps use standard 40-byte BITMAPINFOHEADER structs as headers
Private Type BITMAPINFOHEADER
    size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    ImageSize As Long
    xPelsPerMeter As Long
    yPelsPerMeter As Long
    ColorUsed As Long
    ColorImportant As Long
End Type

'Individual icons are loaded to this struct
Private Type PD_Icon
    'Offset#     Size (in bytes)     Purpose
    '0   1   Specifies image width in pixels. Can be any number between 0 and 255. Value 0 means image width is 256 pixels.
    ico_Width As Byte
    '1   1   Specifies image height in pixels. Can be any number between 0 and 255. Value 0 means image height is 256 pixels.
    ico_Height As Byte
    '2   1   Specifies number of colors in the color palette. Should be 0 if the image does not use a color palette.
    ico_PalSize As Byte
    '3   1   Reserved. Should be 0.[Notes 2]
    ico_Reserved As Byte
    '4   2   In ICO format: Specifies color planes. Should be 0 or 1.[Notes 3]
    '        In CUR format: Specifies the horizontal coordinates of the hotspot in number of pixels from the left.
    ico_ColorPlanes As Integer
    '6   2   In ICO format: Specifies bits per pixel. [Notes 4]
    '        In CUR format: Specifies the vertical coordinates of the hotspot in number of pixels from the top.
    ico_BPP As Integer
    '8   4   Specifies the size of the image's data in bytes
    ico_SizeInBytes As Long
    '12  4   Specifies the offset of BMP or PNG data from the beginning of the ICO/CUR file
    ico_OffsetInFile As Long
    'Data in file, unprocessed
    ico_RawData() As Byte
    'Data is in PNG format
    ico_IsPNG As Boolean
    'Set to TRUE if the frame was constructed OK (as far as we know, anyway)
    ico_OK As Boolean
    'Finished pdDIB object
    ico_DIB As pdDIB
End Type

Private m_NumIcons As Long
Private m_Icons() As PD_Icon

'Byte-by-byte access is provided, as always, by a pdStream instance
Private m_Stream As pdStream

'Validate a source filename as ICO format.  Validation *does* touch the file - we validate the icon
' count to make sure there is at least 1 valid icon in the file.
Friend Function IsFileICO(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True, Optional ByVal onSuccessLeaveStreamOpen As Boolean = False) As Boolean
    
    Dim potentiallyICO As Boolean
    potentiallyICO = Files.FileExists(srcFilename)
    If potentiallyICO Then potentiallyICO = (Files.FileLenW(srcFilename) > 6)
    
    'Check extension up front, if requested.  Note that icons and cursors share many attributes,
    ' but we do *not* treat cursors as valid icons, at present.
    If potentiallyICO And requireValidFileExtension Then
        potentiallyICO = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "ico", True)
        'If (Not potentiallyICO) Then potentiallyICO = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "cur", True)
    End If
    
    'Proceed with deeper validation as necessary
    If potentiallyICO Then
        
        'Attempt to load the file
        Set m_Stream = New pdStream
        If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename) Then
        
            'Next is a series of WORDs.  These must be as follows for a valid file:
            ' 2-bytes: 0 (reserved)
            ' 2-bytes: 1 (for icon), 2 (for cursor), all other values are invalid
            ' 2-bytes: number of images in the file (must be > 0)
            potentiallyICO = (m_Stream.ReadInt() = 0)
            If potentiallyICO Then
                m_IconSubtype = m_Stream.ReadInt()
                If (m_IconSubtype <> ico_Icon) And (m_IconSubtype <> ico_Cursor) Then m_IconSubtype = ico_Unknown
                potentiallyICO = (m_IconSubtype <> ico_Unknown)
            End If
            If potentiallyICO Then
                m_NumIcons = m_Stream.ReadIntUnsigned()
                potentiallyICO = (m_NumIcons > 0)
                If ICO_DEBUG_VERBOSE Then PDDebug.LogAction "Valid icon file found; " & CStr(m_NumIcons) & " embedded icon(s) reported"
            End If
        
        End If
        
    End If
    
    IsFileICO = potentiallyICO
    If (Not IsFileICO) Or (Not onSuccessLeaveStreamOpen) Then Set m_Stream = Nothing
    
End Function

'Simplified wrapper to load an icon and produce a pdImage object where each layer represents
' an icon frame from the file.
Friend Function LoadICO(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal checkExtension As Boolean = True) As PD_ICOResult

    'Reset some internal parameters to ensure subsequent reads are accurate.  (This is critical if multiple PSDs
    ' are read back-to-back.)
    Me.Reset
    
    'Validate the file
    If Me.IsFileICO(srcFile, checkExtension, True) Then
    
        'The file is validated.  Start loading icon headers; there must be m_NumIcons entries
        LoadICO = ico_Success
        ReDim m_Icons(0 To m_NumIcons - 1) As PD_Icon
        
        Dim i As Long
        For i = 0 To m_NumIcons - 1
            
            'Headers are fixed in size; interpretation does vary slightly for icons vs cursors
            With m_Icons(i)
                
                'Width/height can be 0; this means the frame is >= 256 pixels in that dimension.
                ' (Note that the precise size is listed in the BITMAPINFOHEADER for this frame,
                ' which uses normal 4-byte integers for dimensions.)
                .ico_Width = m_Stream.ReadByte()
                .ico_Height = m_Stream.ReadByte()
                .ico_PalSize = m_Stream.ReadByte()
                .ico_Reserved = m_Stream.ReadByte()
                .ico_ColorPlanes = m_Stream.ReadIntUnsigned()
                .ico_BPP = m_Stream.ReadIntUnsigned()
                .ico_SizeInBytes = m_Stream.ReadLong()
                .ico_OffsetInFile = m_Stream.ReadLong()
            End With
            
        Next i
        
        'In a properly constructed icon file, the stream pointer will now point at the
        ' icon data for the first icon in the collection.  Icon files may not be constructed
        ' properly, so we're going to manually position the file pointer as we load each icon.
        
        'We're also going to perform some minor validation on header data and raise warnings
        ' if we encounter anything strange.  (Note that the 50mb size limit below is arbitrary,
        ' and intended as a failsafe against malicious input only.)
        Const MAX_ICON_SIZE_IN_BYTES As Long = 50000000
        For i = 0 To m_NumIcons - 1
            
            With m_Icons(i)
                
                'Validate the source icon size
                If (.ico_SizeInBytes < MAX_ICON_SIZE_IN_BYTES) Then
                    
                    'Validate that the offset + size lies within file size bounds
                    If ((.ico_OffsetInFile + .ico_SizeInBytes) <= m_Stream.GetStreamSize()) Then
                        
                        'This frame looks okay.  For now, just cache its raw bytes;
                        ' we'll validate and process them in a subsequent step
                        ReDim .ico_RawData(0 To .ico_SizeInBytes - 1) As Byte
                        m_Stream.SetPosition .ico_OffsetInFile, FILE_BEGIN
                        m_Stream.ReadBytesToBarePointer VarPtr(.ico_RawData(0)), .ico_SizeInBytes
                        
                    Else
                        InternalError "LoadICO", "icon #" & CStr(i + 1) & " lies outside file bounds: " & .ico_SizeInBytes
                        LoadICO = ico_Warning
                    End If
                    
                Else
                    InternalError "LoadICO", "icon #" & CStr(i + 1) & " is too big: " & .ico_SizeInBytes
                    LoadICO = ico_Warning
                End If
                
            End With
            
        Next i
        
        'All icon data has now been loaded.  Close the source file.
        m_Stream.StopStream
        
        'Next, we must attempt to produce usable pdLayer objects from the data
        ' we pulled from file.  This involves tasks like generating alpha channels
        ' from underlying icon masks.
        If (LoadICO < ico_Failure) Then LoadICO = LoadICO_GenerateFrames(dstImage, LoadICO)
        
    Else
        InternalError "LoadICO", "source file isn't in ICO format"
        LoadICO = ico_Failure
    End If

End Function

Friend Sub Reset()
    m_NumIcons = 0
    ReDim m_Icons(0) As PD_Icon
End Sub

'Save an icon file.  Importantly, the passed param string must contain a bunch of information about
' how to generate frames within the icon file.
Friend Function SaveICO_ToFile(ByRef dstFile As String, ByRef srcPDImage As pdImage, ByRef fullParamString As String) As Boolean



End Function

'Given an index into the m_Icons() array, produce a finished DIB for said index's raw icon data
Private Function CreateDIBForIndex(ByRef srcHeader As BITMAPINFOHEADER, ByRef srcStream As pdStream, ByVal srcIndex As Long) As Boolean
    
    Dim i As Long
    
    'If the underlying image requires a palette, it immediately follows the header
    Dim palSize As Long, srcPalette() As RGBQuad
    If (srcHeader.BitCount <= 8) Then palSize = 2 ^ srcHeader.BitCount Else palSize = 0
    
    If (palSize > 0) Then
        ReDim srcPalette(0 To palSize - 1) As RGBQuad
        For i = 0 To palSize - 1
            srcStream.ReadBytesToBarePointer VarPtr(srcPalette(i)), 4&
            srcPalette(i).Alpha = 255
        Next i
    End If
    
    'With the header and palette successfully retrieved, we can now proceed with retrieving
    ' the image's pixel data.  Note that the source DIB's size may be 2x the actual DIB's size,
    ' owing to the presence of a mask.  (The mask immediately follows the actual pixel data,
    ' if it exists.)
    Dim pxWidth As Long, xFinal As Long, pxBitCount As Long
    pxWidth = srcHeader.Width
    xFinal = pxWidth - 1
    pxBitCount = srcHeader.BitCount
    
    Dim pxScanline() As Byte, scanlineSize As Long
    If (pxBitCount = 1) Then
        scanlineSize = (pxWidth + 7) \ 8
    ElseIf (pxBitCount = 2) Then
        scanlineSize = (pxWidth + 3) \ 4
    ElseIf (pxBitCount = 4) Then
        scanlineSize = (pxWidth + 1) \ 2
    ElseIf (pxBitCount = 8) Then
        scanlineSize = pxWidth
    ElseIf (pxBitCount = 24) Then
        scanlineSize = (pxWidth * 3 + 3) And &HFFFFFFFC
    ElseIf (pxBitCount = 32) Then
        scanlineSize = pxWidth * 4
    Else
        InternalError "CreateDIBForIndex", "bad bitcount: " & pxBitCount
        CreateDIBForIndex = False
        Exit Function
    End If
    
    'Validate calculate scanline size
    If (scanlineSize <= 0) Or (Not srcStream.AreBytesAvailable(scanlineSize)) Then
        InternalError "CreateDIBForIndex", "bad scanlinesize: " & scanlineSize
        CreateDIBForIndex = False
        Exit Function
    Else
    
        'Regardless of bit-depth and width, scanline size must always be a multiple of 4,
        ' like all Windows bitmaps.
        scanlineSize = (scanlineSize + 3) And &HFFFFFFFC
    
    End If
    
    'The icon's size appears to be valid.  Initialize the destination DIB and a temporary
    ' array for holding raw scanline data (before it's proceed to 32-bpp).
    ReDim pxScanline(0 To scanlineSize - 1) As Byte
    
    '32-bpp layers don't have an embedded mask, so we need to determine how many scanlines
    ' to process for color data.  (Any icon with an embedded mask will report its height
    ' as 2x its actual height - 1x height for color data, 1x height for mask data.)
    Dim numScanlines As Long, maskExists As Boolean
    maskExists = False
    
    If (srcHeader.Height = srcHeader.Width) Then
        numScanlines = srcHeader.Width
    ElseIf (srcHeader.Height = srcHeader.Width * 2) Then
        numScanlines = srcHeader.Width
        maskExists = True
    Else
        numScanlines = srcHeader.Height \ 2
    End If
    
    Set m_Icons(srcIndex).ico_DIB = New pdDIB
    m_Icons(srcIndex).ico_DIB.CreateBlank pxWidth, numScanlines, 32, 0, 255
    
    'Some bit-depths are easier to handle with lookup tables.  (In effect, we pre-convert
    ' each scanline to 8-bpp.)
    Dim preConvert() As Byte, bitFlags() As Byte
    If (pxBitCount < 8) Then
    
        ReDim preConvert(0 To pxWidth - 1) As Byte
        
        If (pxBitCount = 1) Then
            ReDim bitFlags(0 To 7) As Byte
            bitFlags(0) = 2 ^ 7
            bitFlags(1) = 2 ^ 6
            bitFlags(2) = 2 ^ 5
            bitFlags(3) = 2 ^ 4
            bitFlags(4) = 2 ^ 3
            bitFlags(5) = 2 ^ 2
            bitFlags(6) = 2 ^ 1
            bitFlags(7) = 1
        ElseIf (pxBitCount = 2) Then
            ReDim bitFlags(0 To 3) As Byte
            bitFlags(0) = 2 ^ 6
            bitFlags(1) = 2 ^ 4
            bitFlags(2) = 2 ^ 2
            bitFlags(3) = 1
        End If
    
    End If
    
    'Process each scanline in turn
    Dim x As Long, y As Long, alphaFound As Boolean
    alphaFound = False
    
    Dim tmpSA1D As SafeArray1D, dstPixels() As RGBQuad
    Dim srcByte As Byte, numPixelsProcessed As Long
    
    For y = 0 To numScanlines - 1
    
        'Retrieve the raw source scanline values
        srcStream.ReadBytesToBarePointer VarPtr(pxScanline(0)), scanlineSize
        
        'For low bit-depth images, immediately upsample to 8-bpp
        If (pxBitCount < 8) Then
            
            numPixelsProcessed = 0
            If (pxBitCount = 1) Then
                
                For x = 0 To scanlineSize - 1
                    
                    srcByte = pxScanline(x)
                    
                    'Ignore empty bytes at the end of each scanline
                    For i = 0 To 7
                        If (numPixelsProcessed <= xFinal) Then
                            If (bitFlags(i) = (srcByte And bitFlags(i))) Then preConvert(numPixelsProcessed) = 1 Else preConvert(numPixelsProcessed) = 0
                            numPixelsProcessed = numPixelsProcessed + 1
                        End If
                    Next i
                    
                Next x
            
            ElseIf (pxBitCount = 2) Then
            
                For x = 0 To scanlineSize - 1
                    srcByte = pxScanline(x)
                    For i = 0 To 3
                        If (numPixelsProcessed <= xFinal) Then
                            preConvert(numPixelsProcessed) = (srcByte \ bitFlags(i)) And &H3
                            numPixelsProcessed = numPixelsProcessed + 1
                        End If
                    Next i
                Next x
            
            ElseIf (pxBitCount = 4) Then
            
                For x = 0 To scanlineSize - 1
                    
                    srcByte = pxScanline(x)
                    preConvert(numPixelsProcessed) = (srcByte \ 16) And &HF
                    numPixelsProcessed = numPixelsProcessed + 1
                    
                    If (numPixelsProcessed <= xFinal) Then
                        preConvert(numPixelsProcessed) = srcByte And &HF
                        numPixelsProcessed = numPixelsProcessed + 1
                    End If
                    
                Next x
            
            End If
        
        End If
        
        'Point a destination array at the target DIB
        m_Icons(srcIndex).ico_DIB.WrapRGBQuadArrayAroundScanline dstPixels, tmpSA1D, numScanlines - (y + 1)
        
        'Process each pixel in turn
        For x = 0 To xFinal
        
            Select Case pxBitCount
            
                Case 1, 2, 4
                    dstPixels(x) = srcPalette(preConvert(x))
                    
                Case 8
                    dstPixels(x) = srcPalette(pxScanline(x))
                
                Case 24
                    dstPixels(x).Blue = pxScanline(x * 3)
                    dstPixels(x).Green = pxScanline(x * 3 + 1)
                    dstPixels(x).Red = pxScanline(x * 3 + 2)
                    dstPixels(x).Alpha = 255
                    
                Case 32
                    GetMem4_Ptr VarPtr(pxScanline(x * 4)), VarPtr(dstPixels(x))
                    If (dstPixels(x).Alpha > 0) Then alphaFound = True
            
            End Select
        
        Next x
    
    Next y
    
    'If this frame is not already 32-bpp (or if it is 32-bpp with a blank alpha channel,
    ' e.g. 0RGB format - which is a valid icon color format), we need to generate an alpha
    ' channel using the mask that follows the image.  Note that this step is deliberately
    ' ignored for PNG frames; they provide their own alpha channel, if one exists.
    If maskExists And (Not alphaFound) Then
    
        'The mask is guaranteed to be a 1-bpp channel at the same dimensions as the underlying
        ' image.  Because it is a literal AND mask, 0 = opaque, 1 = transparent.  Note that
        ' a value of 1 presumes black in the image; PD premultiplies alpha so this is a given
        ' in the final image, but in actual icon rendering, it could potentially produce
        ' bizarre behavior depending on the screen contents beneath the icon render.  PD does
        ' not attempt to cover the (esoteric) use-case of AND + non-zero XOR at present.
        
        'Anyway, populate a 1-bpp LUT
        ReDim bitFlags(0 To 7) As Byte
        bitFlags(0) = 2 ^ 7
        bitFlags(1) = 2 ^ 6
        bitFlags(2) = 2 ^ 5
        bitFlags(3) = 2 ^ 4
        bitFlags(4) = 2 ^ 3
        bitFlags(5) = 2 ^ 2
        bitFlags(6) = 2 ^ 1
        bitFlags(7) = 1
        
        'Calculate scanline size (remembering DWORD alignment)
        scanlineSize = (pxWidth + 7) \ 8
        scanlineSize = (scanlineSize + 3) And &HFFFFFFFC
        ReDim pxScanline(0 To scanlineSize - 1) As Byte
        
        'Start iterating through scanlines in the mask
        For y = 0 To numScanlines - 1
            
            'Retrieve the raw mask scanline values
            srcStream.ReadBytesToBarePointer VarPtr(pxScanline(0)), scanlineSize
            
            'Load the contents of each flag into the target image's alpha channel
            m_Icons(srcIndex).ico_DIB.WrapRGBQuadArrayAroundScanline dstPixels, tmpSA1D, numScanlines - (y + 1)
            
            numPixelsProcessed = 0
            For x = 0 To scanlineSize - 1
                srcByte = pxScanline(x)
                For i = 0 To 7
                    If (numPixelsProcessed <= xFinal) Then
                        If (bitFlags(i) = (srcByte And bitFlags(i))) Then dstPixels(numPixelsProcessed).Alpha = 0
                        numPixelsProcessed = numPixelsProcessed + 1
                    End If
                Next i
            Next x
            
        Next y
        
    End If
    
    'Release our unsafe DIB array wrapper
    m_Icons(srcIndex).ico_DIB.UnwrapRGBQuadArrayFromDIB dstPixels
    
    'Premultiply our finished alpha channel
    m_Icons(srcIndex).ico_DIB.SetAlphaPremultiplication True
    
    CreateDIBForIndex = True
    
End Function

'Given an index into the m_Icons() array, return a potential layer name for that icon
Private Function GetNameOfLayer(ByVal layerIndex As Long) As String
    If (layerIndex >= 0) And (layerIndex < m_NumIcons) Then
        With m_Icons(layerIndex)
            GetNameOfLayer = CStr(.ico_DIB.GetDIBWidth) & "x" & CStr(.ico_DIB.GetDIBHeight) & " ("
            If .ico_IsPNG Then
                GetNameOfLayer = GetNameOfLayer & "PNG)"
            Else
                GetNameOfLayer = GetNameOfLayer & .ico_BPP & "-bpp)"
            End If
        End With
    Else
        InternalError "GetNameOfLayer", "bad layerIndex"
    End If
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdICO." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdICO." & funcName & "() reported an error: " & errDescription
    End If
End Sub

'After an icon file has been successfully parsed, call this function to produce usable pdLayer objects
' from the raw frame data.
Private Function LoadICO_GenerateFrames(ByRef dstImage As pdImage, ByVal origLoadState As PD_ICOResult) As PD_ICOResult
    
    'Mirror the original load state
    LoadICO_GenerateFrames = origLoadState
    
    'As before, we're going to work through each individual frame in turn.  The goal of this stage
    ' is to produce a usable 32-bpp DIB for the underlying icon data.  If an icon consists of a
    ' 1-bpp mask, we will produce a full alpha channel from it.
    Dim i As Long
    For i = 0 To m_NumIcons - 1
    
        'The raw bitstream for each icon has been dumped into each entry's .ico_RawData() array.
        ' The bitstream can consist of 1 of 2 possible entries:
        ' 1) A traditional Win32 BITMAPINFOHEADER, followed by a color table, pixels, and mask...
        ' 2) ...or, a full PNG file, embedded as-is
        
        'For option (2), a pdPNG class instance will do all the hard work of loading the icon
        ' for us.  Note that all other parameters (e.g. BITMAPINFOHEADER) must not exist.
        
        'Because of this, we always start by attempting to validate the data as PNG format.
        ' If that fails, we'll try to load the data as a traditional Win32 icon.
        Dim cPNG As pdPNG, tmpImage As pdImage
        
        With m_Icons(i)
            
            'Start by attempting to validate the data as PNG
            Set cPNG = New pdPNG
            If (cPNG.LoadPNG_Simple(vbNullString, tmpImage, .ico_DIB, False, VarPtr(.ico_RawData(0)), .ico_SizeInBytes) = png_Success) Then
                
                'This frame is a PNG!  We've done everything we need to do; no other work is required.
                Set cPNG = Nothing
                Erase .ico_RawData
                .ico_OK = True
                .ico_IsPNG = True
                
            Else
                
                'The embedded frame is *not* a PNG.  It must be a traditional Win32 icon.
                .ico_IsPNG = False
                
                'We now need to parse the raw bytestream as if it were a BMP file.
                ' (In the future, perhaps we could dump this off to a dedicated BMP reader.)
                Dim okToContinue As Boolean
                
                Dim cStream As pdStream: Set cStream = New pdStream
                okToContinue = cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, .ico_SizeInBytes, VarPtr(.ico_RawData(0)))
                If (Not okToContinue) Then
                    LoadICO_GenerateFrames = ico_Failure
                    InternalError "LoadICO_GenerateFrames", "couldn't start stream on raw frame bytes"
                    Exit Function
                End If
                
                'The first entry in the source bytes will be a traditional BITMAPINFOHEADER.
                ' Retrieve it first.
                Dim tmpBIHeader As BITMAPINFOHEADER
                okToContinue = (cStream.ReadBytesToBarePointer(VarPtr(tmpBIHeader), LenB(tmpBIHeader)) = LenB(tmpBIHeader))
                If (Not okToContinue) Then
                    LoadICO_GenerateFrames = ico_Failure
                    InternalError "LoadICO_GenerateFrames", "ran out of bytes for BIH"
                    Exit Function
                End If
                
                'Perform basic validation on the header
                
                'Per the icon file spec on MSDN (https://docs.microsoft.com/en-us/previous-versions/ms997538(v=msdn.10)?redirectedfrom=MSDN)
                ' "The icHeader member has the form of a DIB BITMAPINFOHEADER. Only the following members
                '  are used: biSize, biWidth, biHeight, biPlanes, biBitCount, biSizeImage. All other
                '  members must be 0. The biHeight member specifies the combined height of the XOR and
                '  AND masks. The members of icHeader define the contents and sizes of the other elements
                '  of the ICONIMAGE structure in the same way that the BITMAPINFOHEADER structure defines
                '  a CF_DIB format DIB."
                Dim headerValid As Boolean
                headerValid = (tmpBIHeader.size = 40)
                If (Not headerValid) Then InternalError "LoadICO_GenerateFrames", "bad header size: " & tmpBIHeader.size
                headerValid = (tmpBIHeader.Width > 0)
                If (Not headerValid) Then InternalError "LoadICO_GenerateFrames", "bad header width: " & tmpBIHeader.Width
                headerValid = (tmpBIHeader.Height > 0)
                If (Not headerValid) Then InternalError "LoadICO_GenerateFrames", "bad header height: " & tmpBIHeader.Height
                headerValid = (tmpBIHeader.Planes = 1)
                If (Not headerValid) Then InternalError "LoadICO_GenerateFrames", "bad header planes: " & tmpBIHeader.Planes
                headerValid = (tmpBIHeader.BitCount > 0)
                If (Not headerValid) Then InternalError "LoadICO_GenerateFrames", "bad header bitcount: " & tmpBIHeader.BitCount
                
                'biSizeImage doesn't actually matter; per the MSDN spec (https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader)
                ' "biSizeImage: Specifies the size, in bytes, of the image. This can be set to
                '  0 for uncompressed RGB bitmaps."
                
                'Because of this, we don't attempt validation on that member.
                
                'If all checks passed, this is probably a valid icon!  Generate a unique
                ' frame for it.
                If headerValid Then m_Icons(i).ico_OK = CreateDIBForIndex(tmpBIHeader, cStream, i)
                
            End If
            
        End With
    
    Next i
    
    'With all frames parsed, we now need to construct a new layer for each frame in the
    ' destination pdImage object.
    
    'Start by finding the largest frame in the file; we'll use this for our master image dimensions
    Dim maxWidth As Long, maxHeight As Long
    For i = 0 To m_NumIcons - 1
        If m_Icons(i).ico_OK Then
            If (Not m_Icons(i).ico_DIB Is Nothing) Then
                maxWidth = PDMath.Max2Int(maxWidth, m_Icons(i).ico_DIB.GetDIBWidth())
                maxHeight = PDMath.Max2Int(maxHeight, m_Icons(i).ico_DIB.GetDIBHeight())
            End If
        End If
    Next i
    
    'Ensure both width and height are non-zero
    If (maxWidth > 0) And (maxHeight > 0) Then
        
        'We have enough data to produce a usable image.  Start by initializing basic pdImage attributes.
        dstImage.SetOriginalFileFormat PDIF_ICO
        dstImage.Width = maxWidth
        dstImage.Height = maxHeight
        dstImage.SetDPI 96#, 96#
        
        'Next, we want to figure out which layer to activate + make visible.  This should be the...
        ' 1) largest image in the file...
        ' 2) ...that also has the highest bit-depth...
        ' 3) ...with preference given to PNG frames
        Dim activeLayerIndex As Long, highestBitDepth As Long
        For i = 0 To m_NumIcons - 1
            If (m_Icons(i).ico_DIB.GetDIBWidth = maxWidth) And (m_Icons(i).ico_DIB.GetDIBHeight = maxHeight) Then
            
                'This layer matches the largest layer size we have so far.  If it *also* has the
                ' highest bit-depth, flag it as the new active index.
                If (m_Icons(i).ico_BPP > highestBitDepth) Or (m_Icons(i).ico_IsPNG) Then
                    highestBitDepth = m_Icons(i).ico_BPP
                    If m_Icons(i).ico_IsPNG Then highestBitDepth = 48   'Give PNGs an arbitrarily large preference
                    activeLayerIndex = i
                End If
            
            End If
        Next i
        
        'Next, we want to produce a pdLayer object for each valid frame
        Dim tmpLayer As pdLayer, newLayerID As Long
        
        For i = 0 To m_NumIcons - 1
            
            'Skip frames that didn't validate during loading
            If m_Icons(i).ico_OK And (Not m_Icons(i).ico_DIB Is Nothing) Then
                
                'Ensure alpha is premultiplied
                If (Not m_Icons(i).ico_DIB.GetAlphaPremultiplication()) Then m_Icons(i).ico_DIB.SetAlphaPremultiplication True
                
                'Prep a new layer object and initialize it with the image bits we've retrieved
                newLayerID = dstImage.CreateBlankLayer()
                Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                tmpLayer.InitializeNewLayer PDL_Image, GetNameOfLayer(i), m_Icons(i).ico_DIB
                
                'If this layer's dimensions match the largest layer, make this layer visible.
                ' (All other layers will be hidden, by default.)
                tmpLayer.SetLayerVisibility (i = activeLayerIndex)
                If tmpLayer.GetLayerVisibility Then dstImage.SetActiveLayerByID newLayerID
                
                'Notify the layer of new changes, so it knows to regenerate internal caches on next access
                tmpLayer.NotifyOfDestructiveChanges
                
            End If
        
        Next i
        
        'Notify the image of destructive changes, so it can rebuild internal caches
        dstImage.NotifyImageChanged UNDO_Everything
        dstImage.SetActiveLayerByIndex activeLayerIndex
        
    Else
        LoadICO_GenerateFrames = ico_Failure
        InternalError "LoadICO_GenerateFrames", "no frames with non-zero width/height"
        Exit Function
    End If

End Function

Private Sub Class_Initialize()
    Set m_Stream = New pdStream
    Me.Reset
End Sub

Private Sub Class_Terminate()
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen() Then m_Stream.StopStream True
    End If
End Sub
