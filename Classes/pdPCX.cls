VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PCX (ZSoft "PiCture eXchange") Image Decoder
'Copyright 2025-2025 by Tanner Helland
'Created: 02/January/25
'Last updated: 02/January/25
'Last update: initial build
'
'The PCX file format is almost as old as I am!  PhotoDemon originally handled PCX files via the 3rd-party
' FreeImage library, but FreeImage doesn't support some valid PCX files and has unreliable behavior on
' others (including modern exported ones from well-known software like GIMP).
'
'So in 2025, I wrote my own PCX decoder.  It definitely falls on the "quick-and-dirty" side of things,
' but it's chock full of error-handling (and safety asserts) and I've tested it rigorously on a variety
' of image dimensions (tiny to enormous), including both legacy files and exports from other modern software.
'
'The end result is a fast and lightweight parser that covers many PCX cases that FreeImage does not.
'
'Encoding support is not planned at present, but this may be revisited (as always) if users complain.
'
'A full copy of the PCX spec is available here (and is what I used in writing this class):
' https://www.fileformat.info/format/pcx/spec/a10e75307b3a4cc49c3bbe6db4c41fa2/view.htm
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps additional diagnostic information
' to PD's primary debug log.
Private Const PCX_DEBUG_VERBOSE As Boolean = True

'The full header is copied into this struct at validation time.
Private Type PD_PCXHeader
    pcx_ValidationByte As Byte
    pcx_VersionNumber As Byte
    pcx_Compression As Byte         '0 = none (rare), 1 = RLE (expected)
    pcx_BitsPerPixel As Byte
    pcx_XMin As Long                'Encoded in file as unsigned short
    pcx_YMin As Long                'Encoded in file as unsigned short
    pcx_XMax As Long                'Encoded in file as unsigned short
    pcx_YMax As Long                'Encoded in file as unsigned short
    pcx_Xdpi As Long                'Encoded in file as unsigned short
    pcx_Ydpi As Long                'Encoded in file as unsigned short
    pcx_EGAPalette(0 To 47) As Byte 'EGA palette for 16-color images
    pcx_Reserved1 As Byte           'Should be 0
    pcx_NumChannels As Byte         'Should validate to 1, 3, or 4
    pcx_BytesPerScanline As Long    'Encoded in file as unsigned short, and MUST BE EVEN
    pcx_PaletteMode As Integer      '1 = monochrome OR color?, 2 = grayscale
    pcx_XScreenResolution As Long   'Encoded in file as unsigned short
    pcx_YScreenResolution As Long   'Encoded in file as unsigned short
    pcx_Reserved2(0 To 53) As Byte  'Reserved for future use, should be 0s
End Type

'Filled by any call to IsFilePCX; contents are unreliable if IsFilePCX returns FALSE
Private m_Header As PD_PCXHeader

'Will be TRUE if IsFilePCX validated the header successfully
Private m_HeaderValid As Boolean

'Current palette, if any; will contain the contents of the EGA palette by default,
' but can be overridden by a trailing palette if the file defines one.
Private m_Palette(0 To 255) As RGBQuad

'Byte-by-byte access is provided, as always, by a pdStream instance.
Private m_Stream As pdStream

'The last filename loaded.  We use this to skip validation during loading, if the caller already called
' IsFilePCX on the current filename they want loaded.
Private m_OpenFilename As String

'PCX files use a simple but well-defined header (https://en.wikipedia.org/wiki/PCX#PCX_file_format).
' This makes validation straightforward.
Friend Function IsFilePCX(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True, Optional ByVal leaveStreamOpenIfValid As Boolean = True) As Boolean
        
    Const FUNC_NAME As String = "IsFilePCX"
        
    IsFilePCX = False
    On Error GoTo BadPCXFile
    
    m_OpenFilename = vbNullString
    m_HeaderValid = False
    ResetHeader
    
    Dim potentialMatch As Boolean
    potentialMatch = Files.FileExists(srcFilename)
    
    'Check extension too, as requested.
    If (potentialMatch And requireValidFileExtension) Then
        potentialMatch = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "pcx", True)
    End If
    
    'PCX files have a fixed-size 128-byte header; smaller files can't be PCX.
    If potentialMatch Then potentialMatch = (Files.FileLenW(srcFilename) > 128)
    
    'If any of the failsafe checks failed, exit immediately
    If (Not potentialMatch) Then Exit Function
    
    'Populate a PCX header, then look for valid values in relevant header fields.
    If (m_Stream Is Nothing) Then Set m_Stream = New pdStream
    If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename, optimizeAccess:=OptimizeSequentialAccess) Then
        
        m_OpenFilename = srcFilename
        
        'Refer to the header definition, above, for the order and size of header fields.
        With m_Header
            .pcx_ValidationByte = m_Stream.ReadByte()
            .pcx_VersionNumber = m_Stream.ReadByte()
            .pcx_Compression = m_Stream.ReadByte()
            .pcx_BitsPerPixel = m_Stream.ReadByte()
            .pcx_XMin = m_Stream.ReadIntUnsigned()
            .pcx_YMin = m_Stream.ReadIntUnsigned()
            .pcx_XMax = m_Stream.ReadIntUnsigned()
            .pcx_YMax = m_Stream.ReadIntUnsigned()
            .pcx_Xdpi = m_Stream.ReadIntUnsigned()
            .pcx_Ydpi = m_Stream.ReadIntUnsigned()
            m_Stream.ReadBytesToBarePointer VarPtr(.pcx_EGAPalette(0)), 48
            .pcx_Reserved1 = m_Stream.ReadByte()
            .pcx_NumChannels = m_Stream.ReadByte()
            .pcx_BytesPerScanline = m_Stream.ReadIntUnsigned()
            .pcx_PaletteMode = m_Stream.ReadInt()
            .pcx_XScreenResolution = m_Stream.ReadIntUnsigned()
            .pcx_YScreenResolution = m_Stream.ReadIntUnsigned()
            m_Stream.ReadBytesToBarePointer VarPtr(.pcx_Reserved2(0)), 54
        End With
        
        'The file offset should now be 128.  Validate relevant header bits before exiting.
        With m_Header
            
            'Validate marker byte
            Const ZSOFT_VALIDATION_BYTE As Byte = &HA
            IsFilePCX = (.pcx_ValidationByte = ZSOFT_VALIDATION_BYTE)
            If (IsFilePCX And PCX_DEBUG_VERBOSE) Then PDDebug.LogAction "File appears to be PCX format"
            
            'Validate compression
            If IsFilePCX Then
                IsFilePCX = (.pcx_Compression < 2)
                If (Not IsFilePCX) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Bad compression flag: " & .pcx_Compression
            End If
            
            'Validate bpp
            If IsFilePCX Then
                IsFilePCX = (.pcx_BitsPerPixel = 1) Or (.pcx_BitsPerPixel = 2) Or (.pcx_BitsPerPixel = 4) Or (.pcx_BitsPerPixel = 8) Or (.pcx_BitsPerPixel = 24) Or (.pcx_BitsPerPixel = 32)
                If (Not IsFilePCX) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Bad bpp: " & .pcx_BitsPerPixel
            End If
            
            'Validate x/y dimensions
            If IsFilePCX Then
                IsFilePCX = (.pcx_XMax >= .pcx_XMin)
                If (Not IsFilePCX) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Bad XMin/Max: " & .pcx_XMin & ", " & .pcx_XMax
            End If
            If IsFilePCX Then
                IsFilePCX = (.pcx_YMax >= .pcx_YMin)
                If (Not IsFilePCX) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Bad YMin/Max: " & .pcx_YMin & ", " & .pcx_YMax
            End If
            
            'Validate channel count
            If IsFilePCX Then
                IsFilePCX = (.pcx_NumChannels = 1) Or (.pcx_NumChannels = 3) Or (.pcx_NumChannels = 4)
                If (Not IsFilePCX) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Bad channel count: " & .pcx_NumChannels
            End If
            
        End With
        
        'IsFilePCX now reflects a (pretty confident!) guess at the file being PCX format
        
    '/Couldn't start stream; file may be locked or inaccessible
    Else
        GoTo BadPCXFile
    End If
    
    'Close the file stream before exiting, as requested
    If (Not leaveStreamOpenIfValid) Then
        If (Not m_Stream Is Nothing) Then m_Stream.StopStream True
    End If
    
    'Remember validation state so we can skip it if the caller proceeds with loading
    m_HeaderValid = IsFilePCX
    
    If PCX_DEBUG_VERBOSE And IsFilePCX Then PDDebug.LogAction "PCX format looks valid!"
    
    Exit Function
    
'On any parse error, this function jumps to this branch and simply closes the underlying file, then exits
BadPCXFile:
    
    Set m_Stream = Nothing
    InternalError FUNC_NAME, "critical parse failure"
    IsFilePCX = False
    
End Function

'Validate and load a candidate PCX file
Friend Function LoadPCX_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "LoadPCX_FromFile"
    LoadPCX_FromFile = False
    
    On Error GoTo BadPCXFile
    
    'Validate the file as necessary
    Dim needToValidate As Boolean
    needToValidate = True
    If m_HeaderValid And (m_OpenFilename = srcFile) Then
        If (Not m_Stream Is Nothing) Then
            needToValidate = (Not m_Stream.IsOpen())
        End If
    End If
    
    If needToValidate Then needToValidate = Not Me.IsFilePCX(srcFile, True, True)
    
    'If the underlying stream is open and the header validated successfully, proceed!
    If m_HeaderValid Then
        
        'Just in case, forcibly reset the file pointer to offset 128 (where pixel data begins)
        m_Stream.SetPosition 128, FILE_BEGIN
        
        'Construct a default palette object using the contents of the file header.
        ' (We may overwrite this with other data later, as necessary.)
        Dim i As Long, palOffset As Long
        palOffset = 0
        
        For i = 0 To 15
            m_Palette(i).Red = m_Header.pcx_EGAPalette(palOffset)
            m_Palette(i).Green = m_Header.pcx_EGAPalette(palOffset + 1)
            m_Palette(i).Blue = m_Header.pcx_EGAPalette(palOffset + 2)
            m_Palette(i).Alpha = 255
            palOffset = palOffset + 3
        Next i
        
        'If the header specifies EGA mode, overwrite the embedded palette (which may be all-zeroes)
        ' with a stock EGA palette instead.
        If (m_Header.pcx_VersionNumber < 2) Or (m_Header.pcx_VersionNumber = 3) Then Palettes.GetStockPalette_EGA m_Palette, False
        
        'Calculate image dimensions, in pixels
        Dim imgWidth As Long, imgHeight As Long, imgStrideInBytes As Long, numChannels As Long
        imgWidth = (m_Header.pcx_XMax - m_Header.pcx_XMin) + 1
        imgHeight = (m_Header.pcx_YMax - m_Header.pcx_YMin) + 1
        imgStrideInBytes = m_Header.pcx_BytesPerScanline * m_Header.pcx_NumChannels
        
        If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Preparing to decode PCX image (" & imgWidth & "x" & imgHeight & ") @ " & m_Header.pcx_BitsPerPixel & "x" & m_Header.pcx_NumChannels
        
        Dim actualBPP As Long
        actualBPP = m_Header.pcx_BitsPerPixel * m_Header.pcx_NumChannels
        
        'Note the current file position and store it; we'll forcibly calculate correct offsets
        ' for each line as we go.
        Dim initFileOffset As Long, maxStreamSize As Long
        initFileOffset = m_Stream.GetPosition()
        maxStreamSize = m_Stream.GetStreamSize()
        
        Dim xFinal As Long
        xFinal = imgWidth - 1
        
        'Before decoding, we need to grab the trailing palette (if any).
        If (m_Header.pcx_VersionNumber = 5) Then
        
            'Here are the direct instructions from the spec on retrieving a 256-color palette:
            '   "To access a 256 color palette:
            '   "First, check the version number in the header; if it contains a 5 there is a palette."
            '   "Second, read to the end of the file and count back 769 bytes.  The value you find
            '    should be a 12 decimal, showing the presence of a 256 color palette."
            m_Stream.SetPosition m_Stream.GetStreamSize - 769, FILE_BEGIN
            If (m_Stream.ReadByte() = 12) Then
                If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "EOF palette found; retrieving before parsing..."
                For i = 0 To 255
                    m_Palette(i).Red = m_Stream.ReadByte()
                    m_Palette(i).Green = m_Stream.ReadByte()
                    m_Palette(i).Blue = m_Stream.ReadByte()
                    m_Palette(i).Alpha = 255
                Next i
            End If
            
            'Before continuing, reset the file pointer to its original offset.
            m_Stream.SetPosition initFileOffset, FILE_BEGIN
            
        End If
        
        'Prepare the destination DIB (we'll fill its contents as we go).
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank imgWidth, imgHeight, 32, 255, 255
        
        Dim tmpSA1D As SafeArray1D, dstPixels() As Byte, dstPixelsRGBQuad() As RGBQuad
        
        'The RLE compression used in PCX files operates on byte-level data (*not* pixel data).
        ' This simplifies decoding somewhat, since decompression is identical for all color-depths.
        Dim uncompressedBytes() As Byte
        ReDim uncompressedBytes(0 To imgStrideInBytes - 1) As Byte
        
        Dim compressedBytes() As Byte
        ReDim compressedBytes(0 To imgStrideInBytes * 2 - 1) As Byte
        
        'Start iterating scanlines.  We'll build the image one line at a time as we go.
        Dim x As Long, y As Long, curChannel As Long, maxChannels As Long
        maxChannels = m_Header.pcx_NumChannels - 1
        
        For y = 0 To imgHeight - 1
            
            'Start by decoding the current line in the source data.  This consists of decompressing
            ' the RLE data into our (uncompressed bytes) array, then parsing that into specific pixel info.
            
            'For speed purposes, pull a worst-case number of bytes from the source data into a local array.
            ' (This is much faster than complex function calls on every RLE byte.)
            Dim worstCaseBytes As Long
            worstCaseBytes = imgStrideInBytes * 2
            If (worstCaseBytes > maxStreamSize - initFileOffset) Then worstCaseBytes = maxStreamSize - initFileOffset
            m_Stream.ReadBytes compressedBytes, imgStrideInBytes * 2, False
            
            'Reset the offset into the uncompressed data holder
            Dim xOffset As Long, srcBytesRead As Long
            xOffset = 0
            srcBytesRead = 0
            
            Do While (xOffset < imgStrideInBytes)
                
                Dim curByte As Byte
                curByte = compressedBytes(srcBytesRead)
                srcBytesRead = srcBytesRead + 1
                
                'If the first two bits in this byte are 1, (e.g. 0b11000000 mask) then
                ' the remaining bytes indicate a run count for the byte that follows.
                ' The run count must be non-zero or the file is invalid.
                Const PCX_RLE_MARKER As Byte = &HC0, PCX_RLE_MASK As Byte = &H3F
                If ((curByte And PCX_RLE_MARKER) = PCX_RLE_MARKER) Then
                    
                    'Calculate run length for the next byte
                    Dim runLength As Long
                    runLength = curByte And PCX_RLE_MASK
                    
                    'Retrieve the next byte, and copy it [n] times into the destination.
                    curByte = compressedBytes(srcBytesRead)
                    srcBytesRead = srcBytesRead + 1
                    
                    'Failsafe check for overruns
                    If (xOffset + runLength > imgStrideInBytes) Then runLength = imgStrideInBytes - xOffset
                    
                    'Failsafe check for null runs (invalid)
                    If (runLength = 0) Then
                        InternalError FUNC_NAME, "bad RLE encoding"
                        GoTo BadPCXFile
                    End If
                    
                    'Write the run, advance the offset accordingly, and continue!
                    For x = xOffset To xOffset + runLength - 1
                        uncompressedBytes(x) = curByte
                    Next x
                    
                    xOffset = xOffset + runLength
                    
                'If the two high bytes are not set, copy over this byte as-is
                Else
                    uncompressedBytes(xOffset) = curByte
                    xOffset = xOffset + 1
                End If
                
            Loop
            
            'Reset the stream pointer to its correct location for the next scanline
            initFileOffset = initFileOffset + srcBytesRead
            m_Stream.SetPosition initFileOffset, FILE_BEGIN
            
            'uncompressedBytes() now contains the full data required for this scanline,
            ' *without* RLE compression but with no further processing.  Note also that there
            ' may be dummy bytes between channels based on the difference between imgWidth
            ' (which is pixel-precise) and imgBytesPerScanline (which *must* be even regardless
            ' of image dimensions).
            
            'We will start copying bytes from position [0] in the source array
            Dim srcOffset As Long
            srcOffset = 0
            
            'Point a VB array at the target line in the image
            If (actualBPP >= 24) Then
                dstDIB.WrapArrayAroundScanline dstPixels, tmpSA1D, y
            Else
                dstDIB.WrapRGBQuadArrayAroundScanline dstPixelsRGBQuad, tmpSA1D, y
            End If
            
            'We now want to iterate channels, and copy data into the final pixel buffer
            For curChannel = 0 To maxChannels
                
                'Separate handling by bit-depth
                If (actualBPP >= 24) Then
                    
                    'R/B channels need to be swizzled as we go
                    Dim channelOffset As Long
                    If (curChannel = 0) Or (curChannel = 2) Then
                        channelOffset = 2 - curChannel
                    Else
                        channelOffset = curChannel
                    End If
                    
                    'For 24-bit data, just copy it as-is into the destination buffer, but remember the channel offsets!
                    xOffset = curChannel * m_Header.pcx_BytesPerScanline
                    x = 0
                    Do While (x < imgWidth)
                        dstPixels(x * 4 + channelOffset) = uncompressedBytes(xOffset + x)
                        x = x + 1
                    Loop
                    
                ElseIf (actualBPP = 8) Then
                    
                    'Copy data directly from the master palette as we go
                    x = 0
                    Do While (x < imgWidth)
                        dstPixelsRGBQuad(x) = m_Palette(uncompressedBytes(x))
                        x = x + 1
                    Loop
                    
                End If
                    
            Next curChannel
            
            'Unwrap the temporary destination array
            If (actualBPP >= 24) Then
                dstDIB.UnwrapArrayFromDIB dstPixels
            Else
                dstDIB.UnwrapRGBQuadArrayFromDIB dstPixelsRGBQuad
            End If
            
        Next y
        
'        'How we retrieve pixels depends on color-depth, obviously.  We're going to do this
'        ' in two passes to simplify the process of handling multiple color-depths and RLE decoding.
'        Dim pxScanline() As Byte, scanlineSize As Long
'        If (m_Header.pcx_BitsPerPixel = 1) Then
'            scanlineSize = (imgWidth + 7) \ 8
'        ElseIf (pxBitCount = 2) Then
'            scanlineSize = (imgWidth + 3) \ 4
'        ElseIf (pxBitCount = 4) Then
'            scanlineSize = (imgWidth + 1) \ 2
'        ElseIf (pxBitCount = 8) Then
'            scanlineSize = imgWidth
'        ElseIf (pxBitCount = 12) Then
'            scanlineSize = imgWidth * 2  '16-bit alignment; 4-bits go unused
'        ElseIf (pxBitCount = 16) Then
'            scanlineSize = imgWidth * 2
'        ElseIf (pxBitCount = 24) Then
'            scanlineSize = imgWidth * 3
'        ElseIf (pxBitCount = 32) Then
'            scanlineSize = imgWidth * 4
'        Else
'            InternalError FUNC_NAME, "bad bitcount: " & pxBitCount
'            GoTo NextFrame
'        End If
'
'        If (scanlineSize <= 0) Then
'            InternalError FUNC_NAME, "bad scanline size: " & scanlineSize
'            GoTo NextFrame
'        End If
        
        LoadPCX_FromFile = True
        
    '/File is not PCX; silently ignore it
    End If
    
    Exit Function
    
BadPCXFile:
    InternalError FUNC_NAME, "abandoned load due to critical error"
    LoadPCX_FromFile = False
    
End Function

'Clear m_Header
Private Sub ResetHeader()
    VBHacks.FillMemory VarPtr(m_Header), LenB(m_Header), 0
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdPCX." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdPCX." & funcName & "() reported an error: " & errDescription
    End If
End Sub

'The underlying stream would auto-free naturally, but I like being tidy
Private Sub Class_Terminate()
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen Then m_Stream.StopStream True
    End If
End Sub
