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
'Last updated: 04/January/25
'Last update: add a small bit of helper code for recovering badly encoded PCX files
'
'The PCX file format is nearly as old as I am!  PhotoDemon originally handled PCX files via the 3rd-party
' FreeImage library, but FreeImage doesn't support some valid PCX files and has unreliable behavior on
' others (including modern exported ones from well-known software like GIMP).
'
'So in 2025, I wrote my own PCX decoder.  It definitely falls on the "quick-and-dirty" side, but it's full of
' detailed error-handling (and safety asserts) and I've tested it rigorously on a variety of image dimensions
' (tiny to enormous), including both legacy files and exports from other modern software.
'
'The end result is a fast and lightweight parser that covers many PCX cases that FreeImage does not,
' and that provides *better* coverage than Photoshop (and coverage of all features that Photoshop does, too).
' The new decoder is also capable of recovering some "broken" PCX files, which is kinda nice.
'
'Encoding support is not planned at present, but this may be revisited (as always) if users complain.
'
'A full copy of the PCX spec is available here (and is what I used in writing this class):
' https://www.fileformat.info/format/pcx/spec/a10e75307b3a4cc49c3bbe6db4c41fa2/view.htm
'
'A number of other articles were used in developing this class.  Thank you to the authors:
' https://moddingwiki.shikadi.net/wiki/PCX_Format
' https://www.fileformat.info/format/pcx/egff.htm
' https://en.wikipedia.org/wiki/PCX#PCX_file_format
' http://fileformats.archiveteam.org/wiki/PCX
' https://www.fysnet.net/pcxfile.htm
' http://cd.textfiles.com/clipart1996/
' https://stackoverflow.com/questions/48172059/is-is-possible-to-determine-from-the-header-of-a-1-bit-pcx-if-the-palette-shoul
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps additional diagnostic information
' to PD's primary debug log.
Private Const PCX_DEBUG_VERBOSE As Boolean = False

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
Private m_Palette() As RGBQuad

'Byte-by-byte access is provided, as always, by a pdStream instance.
Private m_Stream As pdStream

'The last filename loaded.  We use this to skip validation during loading, if the caller already called
' IsFilePCX on the current filename they want loaded.
Private m_OpenFilename As String

'These properties are filled after loading, and primarily exist as helpers so PD can store basic data
' if the user chooses "use original esttings" at export time.
Private m_EquivalentColorDepth As Long
Private m_HasAlpha As Boolean
Private m_IsGrayscale As Boolean

'PCX uses simple RLE encoding.  Each scanline is encoded individually.  The spec is explicit that RLE runs
' should *NOT* cross scanlines, but as with many older image formats, there are images in the wild that
' ignore this.  Instead of just throwing those files out, let's try and rescue them.
Private m_BadRLECount() As Long, m_BadRLEValue() As Long

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
                If (Not IsFilePCX) Then IsFilePCX = (.pcx_BitsPerPixel = 1) And (.pcx_NumChannels = 2)
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
    
    'This exists only as a failsafe; potential encoding errors are explicitly handled via other means
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
    
    'If the underlying stream is open and the header validated successfully, attempt a full load
    If m_HeaderValid Then
        
        'Forcibly reset the file pointer to offset 128 (where the header ends and pixel data begins).
        m_Stream.SetPosition 128, FILE_BEGIN
        
        'Construct a default palette object using the contents of the file header.
        ' (We may overwrite this with other data later, as necessary - later PCX files can store
        '  a trailing palette at the end of the file.)
        Dim i As Long, palOffset As Long
        palOffset = 0
        ReDim m_Palette(0 To 255) As RGBQuad
        
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
        
        'Untested, but as a failsafe: if a PCX specifies a grayscale palette, explicitly populate that now
        If (m_Header.pcx_PaletteMode = 2) Then
            If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Palette marked as grayscale; replacing colors now..."
            Palettes.GetPalette_GrayscaleEx m_Palette, 16, True
            m_IsGrayscale = True
        Else
            m_IsGrayscale = False
        End If
        
        'PCX files use max/min dimensions; convert these to 0-based
        Dim imgWidth As Long, imgHeight As Long, imgStrideInBytes As Long
        imgWidth = (m_Header.pcx_XMax - m_Header.pcx_XMin) + 1
        imgHeight = (m_Header.pcx_YMax - m_Header.pcx_YMin) + 1
        imgStrideInBytes = m_Header.pcx_BytesPerScanline * m_Header.pcx_NumChannels
        
        If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Preparing to decode PCX image (" & imgWidth & "x" & imgHeight & ") @ " & m_Header.pcx_BitsPerPixel & "x" & m_Header.pcx_NumChannels
        
        'PCX uses a strange hybrid of "planes" and "bit-depths" that is unlike other, more modern formats.
        ' We need to multiply these together to arrive at an "actual" bit-depth (although note that it
        ' doesn't always map cleanly - for example, PCX supports a "1-bit, 3-plane" encoding that's
        ' equivalent to "3-bits-per-pixel" - an oddity not supported by other formats).
        Dim actualBPP As Long
        actualBPP = m_Header.pcx_BitsPerPixel * m_Header.pcx_NumChannels
        m_HasAlpha = False
        
        Select Case actualBPP
            Case 1
                m_EquivalentColorDepth = 1
            Case 2
                m_EquivalentColorDepth = 2
            Case 3, 4
                m_EquivalentColorDepth = 4
            Case 8
                m_EquivalentColorDepth = 8
            Case 24
                m_EquivalentColorDepth = 24
            
            'Technically, PCX does *not* support alpha, but it does support an arbitrary number of
            ' color planes, so we'll load 8-bit, 4-channel PCX without complaint (and treat the fourth
            ' channel as alpha).
            Case 32
                m_EquivalentColorDepth = 32
                m_HasAlpha = True
        End Select
        
        'Note the current file position and store it; we'll forcibly calculate correct offsets
        ' for each line as we go.
        Dim initFileOffset As Long, maxStreamSize As Long
        initFileOffset = m_Stream.GetPosition()
        maxStreamSize = m_Stream.GetStreamSize()
        
        Dim xFinal As Long
        xFinal = imgWidth - 1
        
        'Before decoding, we may need to grab a trailing palette (if any).
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
        
        'Palettes in PCX files are complex, and may require further processing.
        ' (This link was very helpful: https://moddingwiki.shikadi.net/wiki/PCX_Format)
        If (actualBPP < 24) Then
            
            If PCX_DEBUG_VERBOSE Then
                PDDebug.LogAction "File version: " & m_Header.pcx_VersionNumber
                PDDebug.LogAction "Palette mode: " & m_Header.pcx_PaletteMode
                PDDebug.LogAction "Plane count: " & m_Header.pcx_NumChannels
                PDDebug.LogAction "BPP: " & m_Header.pcx_BitsPerPixel
                PDDebug.LogAction "Performing palette maintenance..."
            End If
            
            'Photoshop always loads monochrome PCX files as pure black-and-white, regardless of
            ' the current palette.  I'm okay mimicking that behavior too.
            If (actualBPP = 1) Then
                m_Palette(0) = Colors.GetRGBQuadFromHex("#000000")
                m_Palette(1) = Colors.GetRGBQuadFromHex("#FFFFFF")
            End If
            
        End If
        
        'Prepare a destination DIB (we'll fill its contents as we go).
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank imgWidth, imgHeight, 32, 0, 255
        
        Dim tmpSA1D As SafeArray1D, dstPixels() As Byte, dstPixelsRGBQuad() As RGBQuad
        
        'The RLE compression used in PCX files operates on byte-level data (*not* pixel data).
        ' This simplifies decoding somewhat, since decompression is identical for all color-depths.
        Dim uncompressedBytes() As Byte
        ReDim uncompressedBytes(0 To imgStrideInBytes - 1) As Byte
        
        Dim compressedBytes() As Byte
        ReDim compressedBytes(0 To imgStrideInBytes * 2 - 1) As Byte
        
        Dim shuffledBytes() As Byte
        ReDim shuffledBytes(0 To imgWidth - 1) As Byte
        
        'VB6's lack of bit shift functions means some bit-depths are easier to handle via lookup tables.
        Dim bitFlags() As Byte
        If (actualBPP < 8) Then
            If (m_Header.pcx_BitsPerPixel = 1) Then
                ReDim bitFlags(0 To 7) As Byte
                bitFlags(0) = 2 ^ 7
                bitFlags(1) = 2 ^ 6
                bitFlags(2) = 2 ^ 5
                bitFlags(3) = 2 ^ 4
                bitFlags(4) = 2 ^ 3
                bitFlags(5) = 2 ^ 2
                bitFlags(6) = 2 ^ 1
                bitFlags(7) = 1
            ElseIf (m_Header.pcx_BitsPerPixel = 2) Then
                ReDim bitFlags(0 To 3) As Byte
                bitFlags(0) = 2 ^ 6
                bitFlags(1) = 2 ^ 4
                bitFlags(2) = 2 ^ 2
                bitFlags(3) = 1
            End If
        End If
        
        'The PCX spec explicitly disallows RLE runs from crossing scanlines.  (This makes a lot of
        ' sense for the era in which PCX was designed - you only wanted a one-line buffer for pixel
        ' values as you loaded them, to minimize memory requirements.)  Unfortunately, like any
        ' old image format there are images "in the wild" that don't follow this rule.  We'll try
        ' and rescue them by noting bad runs and overflowing them onto the next line.  Note,
        ' however, that this solution is *not* foolproof - for security reasons, we'll never
        ' overflow multiple lines at once, which means that tiny images with RLE runs that span
        ' more than two lines will *not* load correctly (which I'm fine with).
        '
        'Note also that these arrays are explicitly allocated with an extra trailing line, so we
        ' don't need to check boundaries when writing "next line" values to the array.
        ReDim m_BadRLECount(0 To imgHeight) As Long
        ReDim m_BadRLEValue(0 To imgHeight) As Long
        
        'Start iterating scanlines.  We'll build the image one line at a time as we go.
        Dim x As Long, y As Long, j As Long, curChannel As Long, maxChannels As Long
        maxChannels = m_Header.pcx_NumChannels - 1
        
        For y = 0 To imgHeight - 1
            
            'Start by decompressing the current scanline.  This step will migrate data from its simple
            ' RLE format to an "uncompressed bytes" array.
            
            'For speed purposes, pull a worst-case number of compressed bytes into memory.
            ' (Operating on this will be much faster than touching the file on every RLE byte.)
            Dim worstCaseBytes As Long
            worstCaseBytes = imgStrideInBytes * 2
            
            'If our worst-case size extends past the end of the file, default any remaining bytes with
            ' white (in case this is actually a malformed file) then pull the data in from file.
            If (worstCaseBytes > maxStreamSize - initFileOffset) Then
                VBHacks.FillMemory VarPtr(compressedBytes(0)), imgStrideInBytes * 2, 255
                worstCaseBytes = maxStreamSize - initFileOffset
            End If
            m_Stream.ReadBytes compressedBytes, imgStrideInBytes * 2, False
            
            'Reset the offset into the uncompressed data holder
            Dim xOffset As Long, srcBytesRead As Long
            xOffset = 0
            srcBytesRead = 0
            
            'Before iterating x values, pre-fill the start of the line with any "leftover" RLE
            ' runs from the previous line.  (This will only occur on broken PCX files, but I'd
            ' like to try and recover them if I can.)
            If (m_BadRLECount(y) > 0) Then
                
                Dim amtToFill As Long
                amtToFill = PDMath.Min2Int(m_BadRLECount(y), imgStrideInBytes - 1)
                
                For x = xOffset To xOffset + amtToFill - 1
                    uncompressedBytes(x) = m_BadRLEValue(y)
                Next x
                
                xOffset = amtToFill
                
            End If
            
            'Keep decoding until a full scanline has been processed.  (Per the spec, RLE runs cannot
            ' extend across scanlines; we enforce this strictly.)
            Do While (xOffset < imgStrideInBytes)
                
                Dim curByte As Byte
                curByte = compressedBytes(srcBytesRead)
                srcBytesRead = srcBytesRead + 1
                
                'If the first two bits in this byte are 1 (e.g. 0b11000000 mask),
                ' the remaining bits indicate a run count for the byte that follows.
                Const PCX_RLE_MARKER As Byte = &HC0, PCX_RLE_MASK As Byte = &H3F
                If ((curByte And PCX_RLE_MARKER) = PCX_RLE_MARKER) Then
                    
                    'Mask off the actual run length
                    Dim runLength As Long
                    runLength = curByte And PCX_RLE_MASK
                    
                    'Retrieve the next byte, and copy it [n] times into the destination.
                    curByte = compressedBytes(srcBytesRead)
                    srcBytesRead = srcBytesRead + 1
                    
                    'Failsafe check for RLE runs that would cross into the next scanline.
                    If (xOffset + runLength > imgStrideInBytes) Then
                        
                        'Store the bad run amount and value for the *next* line; we'll try to
                        ' rescue the encoding there.
                        m_BadRLECount(y + 1) = (xOffset + runLength) - imgStrideInBytes
                        m_BadRLEValue(y + 1) = curByte
                        
                        'This RLE runs breaks spec and extends to the next line.  Decoding from this
                        ' point forward may not work, but you could theoretically attempt to fix this
                        ' by prepending the run to the next line, then going from there.
                        runLength = imgStrideInBytes - xOffset
                        If PCX_DEBUG_VERBOSE Then InternalError FUNC_NAME, "RLE overrun on line " & y
                        
                    End If
                    
                    'Failsafe check for null run lengths
                    If (runLength <= 0) Then
                        
                        'Technically, a run of 0 shouldn't be valid - but some files "in the wild"
                        ' contain these and Photoshop loads them without complaint.
                        ' Let's do the same here (by simply treating this is a nop).
                        
                    Else
                        
                        'Write the run, advance the destination offset accordingly, and continue!
                        For x = xOffset To xOffset + runLength - 1
                            uncompressedBytes(x) = curByte
                        Next x
                        
                    End If
                        
                    xOffset = xOffset + runLength
                    
                'If the two high bytes are not set, copy over this byte as-is (it's not a run).
                Else
                    uncompressedBytes(xOffset) = curByte
                    xOffset = xOffset + 1
                End If
                
            Loop
            
            'uncompressedBytes() now contains the full data required for this scanline,
            ' *without* RLE compression but with no further processing.  Note also that there
            ' may be trailing dummy bytes based on the difference between imgWidth
            ' (which is pixel-precise) and imgBytesPerScanline (which *must* be even regardless
            ' of image dimensions, per the spec).
            Dim srcOffset As Long
            srcOffset = 0
            
            '24- and 32-bpp data (which we'll load, even though it's technically "not allowed"),
            ' require no additional processing, so we'll access direct bytes in the image instead
            ' of palette quads.
            Dim useBareBytes As Boolean
            If (actualBPP >= 24) Then useBareBytes = True
            
            'Point a VB array at the target scanline in the final layer
            If useBareBytes Then
                dstDIB.WrapArrayAroundScanline dstPixels, tmpSA1D, y
            Else
                dstDIB.WrapRGBQuadArrayAroundScanline dstPixelsRGBQuad, tmpSA1D, y
            End If
            
            'From here, pixel placement is separated by bit-depth
            If (actualBPP >= 24) Then
                
                '24-bpp mode is easy: iterate channels, and simply copy data into the final pixel buffer
                ' as we go!
                For curChannel = 0 To maxChannels
                    
                    'R/B channels need to be swizzled
                    Dim channelOffset As Long
                    If (curChannel = 0) Or (curChannel = 2) Then
                        channelOffset = 2 - curChannel
                    Else
                        channelOffset = curChannel
                    End If
                    
                    xOffset = curChannel * m_Header.pcx_BytesPerScanline
                    x = 0
                    Do While (x < imgWidth)
                        dstPixels(x * 4 + channelOffset) = uncompressedBytes(xOffset + x)
                        x = x + 1
                    Loop
                    
                Next curChannel
            
            '8-bpp data can be copied directly from that palette (which was constructed as BGRA)
            ElseIf (actualBPP = 8) Then
                
                x = 0
                Do While (x < imgWidth)
                    dstPixelsRGBQuad(x) = m_Palette(uncompressedBytes(x))
                    x = x + 1
                Loop
            
            'Lower color depths are more tedious to handle
            ElseIf (actualBPP >= 1) And (actualBPP <= 4) Then
                
                'Okay, I lied - monochrome mode is easy!
                If (actualBPP = 1) Then
                    
                    x = 0: xOffset = 0
                    Do While (x < imgWidth)
                        
                        'Pull the full byte from the post-RLE stream
                        curByte = uncompressedBytes(xOffset)
                        
                        'Iterate bits (but stop once the scanline is complete)
                        For j = 0 To 7
                            
                            If (x < imgWidth) Then
                            
                                'Use our pre-built LUT for simplicity
                                If (bitFlags(j) = (curByte And bitFlags(j))) Then i = 1 Else i = 0
                                
                                'Match palette indices and move to the next pixel.  Note that this
                                ' requires the palette to be correctly set *in advance of this step*.
                                dstPixelsRGBQuad(x) = m_Palette(i)
                                x = x + 1
                                
                            Else
                                Exit For
                            End If
                            
                        Next j
                        
                        xOffset = xOffset + 1
                        
                    Loop
                
                '2-, 3- (!!!), and 4-byte images can actually be encoded two ways:
                ' 1) as 2x1, and 4x1-bit planes, respectively, or...
                ' 2) as 1x2, 1x3, and 1x4-bit planes, respectively.
                '
                'These two modes require separate handling, with the first case in particular
                ' requiring us to iterate the full scanline of uncompressed bytes and shuffle
                ' them into useable palette indices *before* matching them.
                ElseIf (actualBPP >= 2) Or (actualBPP <= 4) Then
                    
                    'Multi-plane mode requires use of another buffer before palette matching
                    If (m_Header.pcx_NumChannels > 1) Then
                        
                        'Black out the shuffled storage before proceeding
                        curChannel = 0
                        VBHacks.FillMemory VarPtr(shuffledBytes(0)), UBound(shuffledBytes) + 1, 0
                        
                        Dim powFlag As Long
                        powFlag = 1
                        
                        'Count pixels and iterate as we go (so we can ignore dead space at the end
                        ' of individual scanlines).
                        x = 0: xOffset = 0
                        Do While (x < imgWidth) And (curChannel <= maxChannels)
                            
                            'Pull the full byte from the post-RLE stream and iterate bits
                            curByte = uncompressedBytes(xOffset)
                            For j = 0 To 7
                                
                                If (x < imgWidth) And (curChannel <= maxChannels) Then
                                
                                    'Use our pre-built LUT for simplicity
                                    If (bitFlags(j) = (curByte And bitFlags(j))) Then i = 1 Else i = 0
                                    
                                    'Use this flag to update the corresponding "shuffled" entry
                                    If (i = 1) Then
                                        shuffledBytes(x) = shuffledBytes(x) Or powFlag
                                    End If
                                    
                                    'Update the power flag as we rotate through channels
                                    x = x + 1
                                    If (x >= imgWidth) Then
                                        x = 0
                                        curChannel = curChannel + 1
                                        powFlag = 2 ^ curChannel
                                        Exit For
                                    End If
                                    
                                Else
                                    Exit For
                                End If
                                
                            Next j
                            
                            xOffset = xOffset + 1
                            
                        Loop
                        
                        'With all bytes shuffled, we can now do standard palette lookup to assign
                        ' pixel values.
                        x = 0
                        Do While (x < imgWidth)
                            dstPixelsRGBQuad(x) = m_Palette(shuffledBytes(x))
                            x = x + 1
                        Loop
                    
                    '"Normal" low-bit palette indices work like any other image format
                    Else
                        
                        If (actualBPP = 2) Then
                            
                            x = 0: xOffset = 0
                            Do While (x < imgWidth)
                                
                                'Pull the full byte from the post-RLE stream
                                curByte = uncompressedBytes(xOffset)
                                
                                'Iterate bit-pairs
                                For j = 0 To 3
                                    
                                    If (x < imgWidth) Then
                                    
                                        'Use our pre-built LUT for simplicity
                                        i = (curByte \ bitFlags(j)) And &H3
                                        
                                        'Match palette indices and move to the next pixel.  Note that this
                                        ' requires the palette to be correctly set *in advance of this step*.
                                        dstPixelsRGBQuad(x) = m_Palette(i)
                                        x = x + 1
                                        
                                    Else
                                        Exit For
                                    End If
                                    
                                Next j
                                
                                xOffset = xOffset + 1
                                
                            Loop
                        
                        ElseIf (actualBPP = 4) Then
                            
                            x = 0: xOffset = 0
                            Do While (x < imgWidth)
                                
                                curByte = uncompressedBytes(xOffset)
                                dstPixelsRGBQuad(x) = m_Palette((curByte \ 16) And &HF)
                                x = x + 1
                                
                                If (x <= imgWidth) Then
                                    dstPixelsRGBQuad(x) = m_Palette(curByte And &HF)
                                    x = x + 1
                                End If
                                
                                xOffset = xOffset + 1
                                
                            Loop
                            
                        End If
                        
                    End If
                    
                '/end (actualBPP >= 2) Or (actualBPP <= 4) case
                End If
                
            Else
                '/bad bit-depth number; this branch is a failsafe only
            End If
            
            'Unwrap the safearray from the target scanline in the finished image
            If useBareBytes Then
                dstDIB.UnwrapArrayFromDIB dstPixels
            Else
                dstDIB.UnwrapRGBQuadArrayFromDIB dstPixelsRGBQuad
            End If
            
            'Reset the stream pointer to its correct location for the next scanline
            initFileOffset = initFileOffset + srcBytesRead
            
            If (Not m_Stream.SetPosition(initFileOffset, FILE_BEGIN)) Then
                InternalError FUNC_NAME, "WARNING: EOF reached prematurely - pixels beyond this point not guaranteed!"
                Exit For
            End If
            
        Next y
        
        LoadPCX_FromFile = True
        
    '/File is not PCX; silently ignore it
    End If
    
    Exit Function
    
BadPCXFile:
    InternalError FUNC_NAME, "abandoned load due to critical error"
    LoadPCX_FromFile = False
    
End Function

'The next four functions are only valid *after* a call to LoadPCX.
Friend Function EquivalentColorDepth() As Long
    EquivalentColorDepth = m_EquivalentColorDepth
End Function

Friend Function GetDPI(ByRef dstXDPI As Single, ByRef dstYDPI As Single) As Boolean
    GetDPI = (m_Header.pcx_XScreenResolution <> 0) And (m_Header.pcx_YScreenResolution <> 0)
    dstXDPI = m_Header.pcx_XScreenResolution
    dstYDPI = m_Header.pcx_YScreenResolution
End Function

Friend Function HasAlpha() As Boolean
    HasAlpha = m_HasAlpha
End Function

Friend Function HasGrayscale() As Boolean
    HasGrayscale = m_IsGrayscale
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
