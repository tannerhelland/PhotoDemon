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
'Last updated: 07/January/25
'Last update: PCX images can now be exported (at all bit-depths!)
'
'The PCX file format is nearly as old as I am!  PhotoDemon originally handled PCX files via the 3rd-party
' FreeImage library, but FreeImage doesn't support some valid PCX files and has unreliable behavior on
' others (including modern exported ones from popular software like GIMP).
'
'So in 2025, I wrote my own PCX decoder and encoder.  They are fast, lightweight, comprehensive, secure,
' and both have been tested rigorously on a variety of image dimensions (tiny to enormous), including both
' "real-world" legacy PCX files and PCX exports from other modern software.
'
'The end result is a fast and lightweight PCX engine that covers many edge-cases FreeImage does not,
' and that provides *better* coverage than Photoshop (and coverage of all features that Photoshop does, too).
' The new decoder is also capable of recovering some "broken" PCX files, which is kinda neat.
'
'All color-depths and PCX features are fully supported, to my knowledge.
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

'PCX files use a single-byte (!!!) marker to start
Private Const ZSOFT_VALIDATION_BYTE As Byte = &HA

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
    pcx_ChannelCount As Byte        'Technically planes, not channels, so can be 1, 2, 3 or 4
    pcx_BytesPerScanline As Long    'Encoded in file as unsigned short, and MUST BE EVEN
    pcx_PaletteMode As Integer      '1 = monochrome OR color, 2 = grayscale; basically unused by modern readers
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
            If (.pcx_Compression <> 1) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "WARNING: unexpected compression value found: " & .pcx_Compression
            .pcx_BitsPerPixel = m_Stream.ReadByte()
            .pcx_XMin = m_Stream.ReadIntUnsigned()
            .pcx_YMin = m_Stream.ReadIntUnsigned()
            .pcx_XMax = m_Stream.ReadIntUnsigned()
            .pcx_YMax = m_Stream.ReadIntUnsigned()
            .pcx_Xdpi = m_Stream.ReadIntUnsigned()
            .pcx_Ydpi = m_Stream.ReadIntUnsigned()
            m_Stream.ReadBytesToBarePointer VarPtr(.pcx_EGAPalette(0)), 48
            .pcx_Reserved1 = m_Stream.ReadByte()
            .pcx_ChannelCount = m_Stream.ReadByte()
            .pcx_BytesPerScanline = m_Stream.ReadIntUnsigned()
            .pcx_PaletteMode = m_Stream.ReadInt()
            .pcx_XScreenResolution = m_Stream.ReadIntUnsigned()
            .pcx_YScreenResolution = m_Stream.ReadIntUnsigned()
            m_Stream.ReadBytesToBarePointer VarPtr(.pcx_Reserved2(0)), 54
        End With
        
        'The file offset should now be 128.  Validate relevant header bits before exiting.
        With m_Header
            
            'Validate marker byte
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
                IsFilePCX = (.pcx_ChannelCount = 1) Or (.pcx_ChannelCount = 3) Or (.pcx_ChannelCount = 4)
                If (Not IsFilePCX) Then IsFilePCX = (.pcx_BitsPerPixel = 1) And (.pcx_ChannelCount = 2)
                If (Not IsFilePCX) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Bad channel count: " & .pcx_ChannelCount
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
        ' with a stock system palette instead.  (Technically an EGA palette should be used, but Adobe
        ' uses a slightly modified version of an EGA palette and it seems to produce better-looking
        ' PCX images than I've been able to do - so I'm just stealing their palette and running with it!)
        If (m_Header.pcx_VersionNumber < 2) Or (m_Header.pcx_VersionNumber = 3) Then Palettes.GetStockPalette pdsp_PSLegacy, m_Palette, False
        
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
        imgStrideInBytes = m_Header.pcx_BytesPerScanline * m_Header.pcx_ChannelCount
        
        If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Preparing to decode PCX image (" & imgWidth & "x" & imgHeight & ") @ " & m_Header.pcx_BitsPerPixel & "x" & m_Header.pcx_ChannelCount
        
        'PCX uses a strange hybrid of "planes" and "bit-depths" that is unlike other, more modern formats.
        ' We need to multiply these together to arrive at an "actual" bit-depth (although note that it
        ' doesn't always map cleanly - for example, PCX supports a "1-bit, 3-plane" encoding that's
        ' equivalent to "3-bits-per-pixel" - an oddity not supported by other formats).
        Dim actualBPP As Long
        actualBPP = m_Header.pcx_BitsPerPixel * m_Header.pcx_ChannelCount
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
                PDDebug.LogAction "Plane count: " & m_Header.pcx_ChannelCount
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
        maxChannels = m_Header.pcx_ChannelCount - 1
        
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
                    If (m_Header.pcx_ChannelCount > 1) Then
                        
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
                                    If (bitFlags(j) = (curByte And bitFlags(j))) Then
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

'Export data to a compatible PCX file
Friend Function SavePCX_ToFile(ByRef srcImage As pdImage, ByRef dstFile As String) As Boolean

    Const FUNC_NAME As String = "SavePCX_ToFile"
    SavePCX_ToFile = False
    
    'This exists only as a failsafe; potential encoding errors are explicitly handled via other means
    On Error GoTo SaveFailed
    
    'Open a stream on the target file
    Set m_Stream = New pdStream
    If (Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFile, 1024, 0, OptimizeSequentialAccess)) Then
        InternalError FUNC_NAME, "couldn't start file stream: " & dstFile
        Exit Function
    End If
    
    'PCX doesn't support multiple frames.  (We may add DCX support someday - that variant *does* support frames.)
    ' As such, grab a composited copy of the current image.
    Dim compositeDIB As pdDIB
    srcImage.GetCompositedImage compositeDIB, True
    
    'For now, composite against white.  (32-bit export works, but is largely unsupported elsewhere.)
    compositeDIB.CompositeBackgroundColor 255, 255, 255
    
    'Color-depth is currently locked at 24/32-bit export (other depths remain TODO).
    ' Note that 32-bit PCX images aren't really a "thing", given that the PCX spec was last
    ' (officially) updated in 2025.  However, some apps *will* load 32-bit PCX images.
    '
    'I'm happy to disable this pending user feedback, but for now, PD will try and preserve alpha.
    ' (Loading a PCX, editing it, then saving it back out to file is likely the most common use-case,
    '  and that's not gonna add transparency anyway - so this may all be a moot point.)
    Dim outputColorDepth As Long, numColorsInPalette As Long, pxIndexed() As Byte
    numColorsInPalette = 257
    
    If DIBs.IsDIBTransparent(compositeDIB) Then
        outputColorDepth = 32
    Else
        outputColorDepth = 24
        
        'Look for lower color depths and grab the relevant palette accordingly
        numColorsInPalette = Palettes.GetDIBColorCount_FastAbort(compositeDIB, m_Palette)
        If (numColorsInPalette <= 2) Then
            outputColorDepth = 1
        ElseIf (numColorsInPalette <= 4) Then
            outputColorDepth = 2
        ElseIf (numColorsInPalette <= 16) Then
            outputColorDepth = 4
        ElseIf (numColorsInPalette <= 256) Then
            outputColorDepth = 8
        End If
        
        'If this image can be represented with a palette, palettize it now
        If (outputColorDepth <= 8) Then
            numColorsInPalette = DIBs.GetDIBAs8bpp_RGBA_SrcPalette(compositeDIB, m_Palette, pxIndexed)
        End If
        
    End If
    
    'We now have an accurate output color depth target, and for images <= 256 colors,
    ' we also have a pre-filled palette in m_Palette().
    
    'We could populate a header struct, but VB padding issues make this annoying, so instead we're just
    ' gonna dump header bits out to file as we go.
    With m_Stream
        .WriteByte ZSOFT_VALIDATION_BYTE
        
        '32-bit RGBA data technically isn't valid, but some software will load it.  I'm still undecided
        ' on whether to allow it at all.  (It requires no work on my part to allow, but I don't want users
        ' writing files that are unusable elsewhere.  Choices choices.)
        If (outputColorDepth = 32) Then
            .WriteByte 5
            
        '24-bit RGB data should always be version 3 per https://www.fileformat.info/format/pcx/egff.htm
        ElseIf (outputColorDepth = 24) Then
            .WriteByte 3
        
        '8-bit RGB data needs v5, so we can append a larger palette to the tail-end of the image
        ElseIf (outputColorDepth = 8) Then
            .WriteByte 5
        
        'Images with <= 16 colors can use the preallocated palette in the header
        ' (instead of appending their own).
        ElseIf (outputColorDepth = 4) Or (outputColorDepth = 2) Or (outputColorDepth = 1) Then
            .WriteByte 2
        Else
            InternalError FUNC_NAME, "bad output depth"
            GoTo SaveFailed
        End If
        
        'The only supported compression value is 1 (RLE-compressed)
        .WriteByte 1
        
        'Next is bits-per-pixel.  Note that PCX files can define "color-depth" in terms of both
        ' bits and planes (e.g. bits-per-plane or planes-per-bit).  We only use the former.
        If (outputColorDepth >= 8) Then
            .WriteByte 8
        Else
            .WriteByte outputColorDepth
        End If
        
        'Instead of dimensions, PCX files use x/y min/max values.  These don't mean anything in
        ' modern software and it's always recommended to zero-out the min offsets (we do the same).
        ' Note, however, that the formula for final size is (max - min) + 1.
        
        'x/y minimums
        .WriteIntU 0
        .WriteIntU 0
        
        'x/y maximums
        .WriteIntU compositeDIB.GetDIBWidth - 1
        .WriteIntU compositeDIB.GetDIBHeight - 1
        
        'x/y DPI
        .WriteIntU srcImage.GetDPI()
        .WriteIntU srcImage.GetDPI()
        
        'Next comes a 48-byte (16x3) palette.  We'll write a standard system palette for all bit-depths,
        ' unless the image actually has a custom <= 16-color palette.
        Dim tmpPalette() As RGBQuad
        Palettes.GetStockPalette pdsp_PSLegacy, tmpPalette, True
        
        Dim i As Long
        If (outputColorDepth <= 4) Then
            
            'Write as many useable colors as the palette contains, then fill remaining entries with black
            For i = 0 To numColorsInPalette - 1
                .WriteByte m_Palette(i).Red
                .WriteByte m_Palette(i).Green
                .WriteByte m_Palette(i).Blue
            Next i
            If (numColorsInPalette < 16) Then
                For i = numColorsInPalette To 15
                    .WriteByte 0
                    .WriteByte 0
                    .WriteByte 0
                Next i
            End If
        
        'If the image doesn't have its own palette (or requires a full 256-color one),
        ' write a dummy EGA palette.
        Else
            For i = 0 To 15
                .WriteByte tmpPalette(i).Red
                .WriteByte tmpPalette(i).Green
                .WriteByte tmpPalette(i).Blue
            Next i
        End If
        
        'Next is a reserved byte, followed by the number of channels
        .WriteByte 0
        
        Dim channelCount As Long
        If (outputColorDepth >= 8) Then
            channelCount = outputColorDepth \ 8
        Else
            channelCount = 1
        End If
        .WriteByte channelCount
        
        'Next is bytes per scanline.  This value *MUST* be even, per the spec which explicitly states:
        ' "Number of bytes to allocate for a scanline plane.  MUST be an EVEN number.  Do NOT calculate from Xmax-Xmin."
        Dim bpScanline As Long
        If (outputColorDepth >= 8) Then
            bpScanline = compositeDIB.GetDIBWidth
        ElseIf (outputColorDepth = 4) Then
            bpScanline = (compositeDIB.GetDIBWidth + 1) \ 2
        ElseIf (outputColorDepth = 2) Then
            bpScanline = (compositeDIB.GetDIBWidth + 3) \ 4
        ElseIf (outputColorDepth = 1) Then
            bpScanline = (compositeDIB.GetDIBWidth + 7) \ 8
        End If
        
        If ((bpScanline And 1) = 1) Then bpScanline = bpScanline + 1
        .WriteIntU bpScanline
        
        'Next comes a palette "mode" which defines either color/bw or grayscale.  The spec explicitly says this
        ' value is ignored in modern versions (back in the day, grayscale was unachievable on CGA/EGA but could
        ' be nicely approximated on VGA so displays could use this to know whether to grab only the first byte
        ' of any color values in the palette as necessary).  PD does not explicitly set this to avoid unpredictable
        ' behavior in legacy software - instead, we always explicitly write the palette we want.
        .WriteIntU 1
        
        'Next is a useless x/y resolution of the screen.  Almost always 0 in "real-world" PCXs.
        .WriteIntU 0
        .WriteIntU 0
        
        'Next is 54 empty reserved bytes to pad out the 128-byte header
        .WritePadding 54
        
    End With
    
    'With that, the header is complete!  Time to write pixel values.
    Dim tmpSA1D As SafeArray1D, srcPixels() As Byte
    
    'Image stride (in bytes) represents how many bytes we *must* encode into a single RLE stream before
    ' breaking for the next line.  (This amount is non-negotiable, even though it potentiall adds empty
    ' bytes at the end of the RLE stream since PCXs enforce a strict rule of "scanlines must be even".)
    Dim imgStrideInBytes As Long
    imgStrideInBytes = bpScanline * channelCount
    
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = compositeDIB.GetDIBWidth
    imgHeight = compositeDIB.GetDIBHeight
    
    'The RLE compression used in PCX files operates on byte-level data (*not* pixel data).
    ' This simplifies encoding somewhat, since compression is identical for all color-depths.
    ' All we have to do first is a small "pre-processing" step to organize pixel data by plane,
    ' and then we can run the RLE compression scheme uniformly against the data without regard
    ' to its source format.
    Dim uncompressedBytes() As Byte
    ReDim uncompressedBytes(0 To imgStrideInBytes - 1) As Byte
    
    'RLE compression has a worst-case encoding of 2x original scanline size
    Dim compressedBytes() As Byte
    ReDim compressedBytes(0 To imgStrideInBytes * 2 - 1) As Byte
    
    Dim x As Long, y As Long, idxLast As Long, dstOffset As Long, offsetInByte As Long
    For y = 0 To imgHeight - 1
        
        'Point a scanline at this line in the source image
        compositeDIB.WrapArrayAroundScanline srcPixels, tmpSA1D, y
        
        'Further handling separates by color-depth
        
        '24/32-bit color
        If (outputColorDepth >= 24) Then
            
            'Copy over bytes in planar order (RRR GGG BBB instead of RGB RGB RGB)
            For x = 0 To imgWidth - 1
                uncompressedBytes(x) = srcPixels(x * 4 + 2)
                uncompressedBytes(imgWidth + x) = srcPixels(x * 4 + 1)
                uncompressedBytes(imgWidth * 2 + x) = srcPixels(x * 4)
                If (outputColorDepth = 32) Then uncompressedBytes(imgWidth * 3 + x) = srcPixels(x * 4 + 3)
            Next x
            
            'Note the index of the last item added to the output array.  We can potentially add a
            ' matching dummy byte after this (if the array has empty space) to save a little file space
            ' during RLE compression.
            If (outputColorDepth = 32) Then
                idxLast = imgWidth * 4 - 1
            Else
                idxLast = imgWidth * 3 - 1
            End If
            
        'Write only palette entries
        ElseIf (outputColorDepth = 8) Then
        
            For x = 0 To imgWidth - 1
                uncompressedBytes(x) = pxIndexed(x, y)
            Next x
            idxLast = imgWidth - 1
            
        ElseIf (outputColorDepth = 4) Then
        
            For x = 0 To imgWidth - 1
                dstOffset = x \ 2
                If (x And 1) Then
                    uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or pxIndexed(x, y)
                Else
                    uncompressedBytes(dstOffset) = pxIndexed(x, y) * 16
                End If
            Next x
            idxLast = dstOffset
            
        ElseIf (outputColorDepth = 2) Then
            
            offsetInByte = 0
            For x = 0 To imgWidth - 1
                dstOffset = x \ 4
                Select Case offsetInByte
                    Case 0
                        uncompressedBytes(dstOffset) = pxIndexed(x, y) * 64
                    Case 1
                        uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or (pxIndexed(x, y) * 16)
                    Case 2
                        uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or (pxIndexed(x, y) * 4)
                    Case 3
                        uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or pxIndexed(x, y)
                End Select
                offsetInByte = offsetInByte + 1
                If (offsetInByte > 3) Then offsetInByte = 0
            Next x
            idxLast = dstOffset
            
        ElseIf (outputColorDepth = 1) Then
            
            offsetInByte = 0
            For x = 0 To imgWidth - 1
                dstOffset = x \ 8
                Select Case offsetInByte
                    Case 0
                        uncompressedBytes(dstOffset) = pxIndexed(x, y) * 128
                    Case 1
                        uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or (pxIndexed(x, y) * 64)
                    Case 2
                        uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or (pxIndexed(x, y) * 32)
                    Case 3
                        uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or (pxIndexed(x, y) * 16)
                    Case 4
                        uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or (pxIndexed(x, y) * 8)
                    Case 5
                        uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or (pxIndexed(x, y) * 4)
                    Case 6
                        uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or (pxIndexed(x, y) * 2)
                    Case 7
                        uncompressedBytes(dstOffset) = uncompressedBytes(dstOffset) Or pxIndexed(x, y)
                End Select
                offsetInByte = offsetInByte + 1
                If (offsetInByte > 7) Then offsetInByte = 0
            Next x
            idxLast = dstOffset
            
        End If
        
        'Unwrap our temporary array from the source image
        compositeDIB.UnwrapArrayFromDIB srcPixels
        
        'The uncompressedBytes() array now holds the image's pixel data in PCX order.
        
        'To improve compression a wee bit, ensure the null trailing bytes (if any) can simply be added
        ' to a trailing RLE packet matching the last meaningful byte.
        If (idxLast < imgStrideInBytes - 1) Then
            For i = idxLast + 1 To imgStrideInBytes - 1
                uncompressedBytes(i) = uncompressedBytes(idxLast)
            Next i
        End If
        
        'Next, we need to perform a (pretty simple) RLE compression on the data.  Note that this
        ' *can* increase data size because PCX's RLE scheme is primitive.  (In a worst-case scenario
        ' it can potentially double image size - see the above declare for compressedBytes() for our
        ' contingencies against this.)
        
        'Remember: we must encode a full "imgStrideInBytes" number of pixels, even if the trailing bytes
        ' lie outside the image's actual boundaries.
        Dim xOffset As Long
        xOffset = 0
        x = 0
        
        Do While (x < imgStrideInBytes)
            
            'See if the current byte must be encoded as a singleton (e.g. not part of a run)
            Dim encodeAsSingleton As Boolean
            encodeAsSingleton = (x = imgStrideInBytes - 1)  'last byte in the scanline
            If (Not encodeAsSingleton) Then encodeAsSingleton = (uncompressedBytes(x) <> uncompressedBytes(x + 1))
            
            'Singleton encoding is a little weird, because certain byte values actually need to be encoded
            ' as *two* bytes.
            If encodeAsSingleton Then
                
                'PCX uses the first two bits (0b11000000 or 0xC0) to indicate an RLE run.
                ' This prevents singleton encoding of anything 192 or higher.
                If (uncompressedBytes(x) >= 192) Then
                    
                    'Write this value as an RLE packet with length 1
                    compressedBytes(xOffset) = 193    '0b11000001 or 0xC1
                    compressedBytes(xOffset + 1) = uncompressedBytes(x)
                    xOffset = xOffset + 2
                    
                'Values smaller than 192 can just be dumped as-is
                Else
                    compressedBytes(xOffset) = uncompressedBytes(x)
                    xOffset = xOffset + 1
                End If
                
                'Increment the source array pointer
                x = x + 1
            
            'This item can be encoded as an RLE run.
            Else
                
                'See how many times we can repeat this pixel.  Note that we must stop repeating under
                ' several conditions:
                ' 1) End of scanline
                ' 2) Byte that no longer matches the run value
                ' 3) Max run size (63) reached.
                '
                'Remember; we only have 6 bits to write the RLE length, which is why 63 (0b00111111)
                ' is the maximum run length.
                Const MAX_RUN_LENGTH As Long = 63
                
                Dim idxTest As Long
                idxTest = x + 1
                
                Dim keepScanning As Boolean
                keepScanning = True
                
                Do
                    idxTest = idxTest + 1
                    keepScanning = (idxTest - x) < MAX_RUN_LENGTH
                    If keepScanning Then keepScanning = (idxTest < imgStrideInBytes)
                    If keepScanning Then keepScanning = (uncompressedBytes(x) = uncompressedBytes(idxTest))
                Loop While keepScanning
                
                'For one reason or another, the RLE run ended, and idxTest points at at the uncompressed byte
                ' *just past* where encoding *must* stop.
                
                'Knowing that, write the RLE data out to file.
                compressedBytes(xOffset) = 192 + (idxTest - x)
                compressedBytes(xOffset + 1) = uncompressedBytes(x)
                xOffset = xOffset + 2
                x = idxTest
                
            End If
            
        Loop
        
        'The compressedBytes() array now contains the RLE stream for this scanline.  Dump it
        ' (up to the relevant offset) out to file.
        m_Stream.WriteBytesFromPointer VarPtr(compressedBytes(0)), xOffset
        
    Next y
    
    'If the image is 8-bit (and *just* 8-bit!) write a trailing palette.  It must always be 256-colors
    ' even if the image only uses < 256.
    If (outputColorDepth = 8) Then
        m_Stream.WriteByte 12   'Magic number; see the spec at https://www.fileformat.info/format/pcx/spec/a10e75307b3a4cc49c3bbe6db4c41fa2/view.htm
        For i = 0 To numColorsInPalette - 1
            m_Stream.WriteByte m_Palette(i).Red
            m_Stream.WriteByte m_Palette(i).Green
            m_Stream.WriteByte m_Palette(i).Blue
        Next i
        If (numColorsInPalette < 256) Then
            For i = numColorsInPalette To 255
                m_Stream.WriteByte 0
                m_Stream.WriteByte 0
                m_Stream.WriteByte 0
            Next i
        End If
    End If
    
    'No trailing bytes are required!  Close the stream and exit.
    m_Stream.StopStream True
    SavePCX_ToFile = True
    
    Exit Function
    
SaveFailed:
    InternalError FUNC_NAME, "abandoned save due to critical error"
    SavePCX_ToFile = False
    
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen Then m_Stream.StopStream True
    End If
    
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
