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
'Copyright 2025-2026 by Tanner Helland
'Created: 02/January/25
'Last updated: 24/February/25
'Last update: tweak importer to fix recovery of 4-bpp, 4-plane data, but as I only have one test image
'             with these settings I'm *still* not sure this handling is correct.
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
'All color-depths and PCX features are fully supported, to my knowledge.  DCX images (multi-frame PCX)
' can also be loaded, with PD placing each frame into its own layer.
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

'The PCX spec is relatively unambiguous, but due to its age (pre-widespread Internet availability),
' I'm guessing many implementors didn't actually read the spec and just reverse-engineered its behavior
' from files in the wild.  This leads to a lot of weird spec deviations in actual images, some of which
' are recoverable (with extra handling on our part).  To enforce stricter per-spec behavior, set this to
' FALSE.  In production builds, I recommend leaving as TRUE, to improve compatibility with more images.
Private Const PCX_DEBUG_RESCUE_BROKEN As Boolean = True

'PCX files use a single-byte (!!!) marker to start
Private Const ZSOFT_VALIDATION_BYTE As Byte = &HA

'At validation time, the full header is copied into this struct.  (Fields are manually copied in one-at-a-time
' to avoid alignment issues.)
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

'Filled by any call to IsFilePCX; contents are unreliable when IsFilePCX returns FALSE
Private m_Header As PD_PCXHeader

'Will be TRUE if IsFilePCX validated the header successfully
Private m_HeaderValid As Boolean

'Current palette, if any; will contain the contents of the header's EGA palette by default,
' but can be overridden by a trailing palette when the file defines one.
Private m_Palette() As RGBQuad

'Byte-by-byte access is provided, as always, by a pdStream instance.
Private m_Stream As pdStream

'The last filename loaded.  We use this to skip validation during loading, if the caller already called
' IsFilePCX() on the filename they want loaded.
Private m_OpenFilename As String

'These properties are filled after loading, and primarily exist as helpers so PD can store basic data
' if the user chooses "use original settings" at export time.
Private m_EquivalentColorDepth As Long
Private m_HasAlpha As Boolean
Private m_IsGrayscale As Boolean

'PCX uses simple RLE encoding.  Each scanline is encoded individually.  The spec is explicit that RLE runs
' should *NOT* cross scanlines, but as with many older image formats, there are images in the wild that
' ignore this.  Instead of just throwing those files out, we try and rescue them.
Private m_BadRLECount() As Long, m_BadRLEValue() As Long

'This class handles both PCX and DCX images.  DCX images are a set of PCX images with a small header
' that points to all embedded PCX offsets.  If the loaded file is a DCX, this flag will be set to TRUE
' at validation time.
Private m_DCXMode As Boolean

'If a file is DCX, we pre-load all PCX offsets and attempt to load each frame as a separate layer inside
' the destination pdImage object.
Private m_dcxCount As Long, m_dcxOffsets() As Long

'PCX files use a simple but well-defined header (https://en.wikipedia.org/wiki/PCX#PCX_file_format).
' This makes validation straightforward.
Friend Function IsFilePCX(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True, Optional ByVal leaveStreamOpenIfValid As Boolean = True) As Boolean
        
    Const FUNC_NAME As String = "IsFilePCX"
        
    IsFilePCX = False
    m_DCXMode = False
    On Error GoTo BadPCXFile
    
    m_OpenFilename = vbNullString
    m_HeaderValid = False
    ResetHeader
    
    Dim potentialMatch As Boolean
    potentialMatch = Files.FileExists(srcFilename)
    
    'Check extension too, as requested.
    If (potentialMatch And requireValidFileExtension) Then
        potentialMatch = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "pcx", True)
        potentialMatch = potentialMatch Or Strings.StringsEqual(Files.FileGetExtension(srcFilename), "pcc", True)
        potentialMatch = potentialMatch Or Strings.StringsEqual(Files.FileGetExtension(srcFilename), "dcx", True)
    End If
    
    'PCX files have a fixed-size 128-byte header; smaller files can't be PCX.
    If potentialMatch Then potentialMatch = (Files.FileLenW(srcFilename) > 128)
    
    'If any of the failsafe checks failed, exit immediately
    If (Not potentialMatch) Then Exit Function
    
    'Populate a PCX header, then look for valid values in relevant header fields.
    If (m_Stream Is Nothing) Then Set m_Stream = New pdStream
    If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename, optimizeAccess:=OptimizeSequentialAccess) Then
        
        m_OpenFilename = srcFilename
        IsFilePCX = LoadPCXHeader()
        
        'If validation failed, attempt DCX validation instead.
        If (Not IsFilePCX) Then
            m_Stream.SetPosition 0
            IsFilePCX = LoadDCXHeader()
        End If
        
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

'DCX files are multi-frame PCX images.  We can load each frame as a layer in PD.
Private Function LoadDCXHeader() As Boolean

    Const FUNC_NAME = "LoadDCXHeader"
    LoadDCXHeader = False
    
    'Failsafe checks
    If (m_Stream Is Nothing) Then Exit Function
    If (Not m_Stream.IsOpen()) Then Exit Function
    
    Dim initStreamPos As Long
    initStreamPos = m_Stream.GetPosition()
    
    Const DCX_MAGIC_NUMBER As Long = &H3ADE68B1
    m_Stream.SetPosition 0
    If (m_Stream.ReadLong = DCX_MAGIC_NUMBER) Then
        
        If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "DCX header found.  Scanning frame list..."
        
        'This is very likely a DCX file.  Set the class-wide DCX flag, then pre-load the
        ' initial table of PCX offsets and store it at module-level.
        
        '(Note that DCX files impose a hard max-limit of 1024 frames.  Files will almost never
        ' include that many frames, but for compatibility reasons they may still include
        ' [1024-n] blank entries to ensure a fixed-size header.)
        Const MAX_OFFSET As Long = 4 * 1024
        
        'Pull the first entry manually and ensure it's not zero.
        Dim posNextFrame As Long
        posNextFrame = m_Stream.ReadLong()
        If (posNextFrame = 0) Or (posNextFrame > m_Stream.GetStreamSize() - 128) Then
            InternalError FUNC_NAME, "bad first frame offset: " & posNextFrame
            m_Stream.SetPosition initStreamPos
            Exit Function
        Else
            If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "First PCX frame found at offset " & posNextFrame
        End If
        
        'Prep storage for all frames (including an extra position, "just in case", to avoid
        ' [n+1] allocation problems)
        ReDim m_dcxOffsets(0 To 1025) As Long
        
        Do
            
            'Store the retrieved frame, then pull the next one
            m_dcxOffsets(m_dcxCount) = posNextFrame
            m_dcxCount = m_dcxCount + 1
            posNextFrame = m_Stream.ReadLong()
            
        Loop While (posNextFrame <> 0) And (m_Stream.GetPosition < MAX_OFFSET)
        
        If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Found " & m_dcxCount & " PCX frames."
        LoadDCXHeader = (m_dcxCount > 0)
        
        'Set the frame pointer to the first frame, then exit
        If LoadDCXHeader Then
            m_Stream.SetPosition m_dcxOffsets(0), FILE_BEGIN
            ReDim Preserve m_dcxOffsets(0 To m_dcxCount - 1) As Long
            m_DCXMode = True
        Else
            m_Stream.SetPosition initStreamPos
            Erase m_dcxOffsets
        End If
        
    Else
        If PCX_DEBUG_VERBOSE Then InternalError FUNC_NAME, "not a DCX image"
    End If
    
End Function

'Before calling this function, m_Stream must point at the location where a PCX header is expected.
' If header validation passes, this function will return TRUE (otherwise, FALSE).
' When TRUE is returned, stream position is guaranteed to be at end-of-header.
' On a FALSE return, stream position will be reset to whatever it was when this function was called.
Private Function LoadPCXHeader() As Boolean
    
    Const FUNC_NAME = "LoadPCXHeader"
    LoadPCXHeader = False
    
    'Failsafe checks
    If (m_Stream Is Nothing) Then Exit Function
    If (Not m_Stream.IsOpen()) Then Exit Function
    
    Dim initStreamPos As Long
    initStreamPos = m_Stream.GetPosition()
    
    'Refer to the header definition (at top) for comments on the order and size of header fields.
    m_Header.pcx_ValidationByte = m_Stream.ReadByte()
    
    'Don't continue if the first byte fails
    If (m_Header.pcx_ValidationByte <> ZSOFT_VALIDATION_BYTE) Then
        If PCX_DEBUG_VERBOSE Then InternalError FUNC_NAME, "not a PCX image (will check DCX next)"
        m_Stream.SetPosition initStreamPos
        Exit Function
    Else
        LoadPCXHeader = True
        If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Frame appears to be PCX format; loading rest of header..."
    End If
    
    'If we're still here, the image might be a PCX file.  Load expected fields and we'll perform
    ' bulk validation at the end.
    With m_Header
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
        .pcx_ChannelCount = m_Stream.ReadByte()
        .pcx_BytesPerScanline = m_Stream.ReadIntUnsigned()
        .pcx_PaletteMode = m_Stream.ReadInt()
        .pcx_XScreenResolution = m_Stream.ReadIntUnsigned()
        .pcx_YScreenResolution = m_Stream.ReadIntUnsigned()
        m_Stream.ReadBytesToBarePointer VarPtr(.pcx_Reserved2(0)), 54
    End With
    
    'The file offset should now be 128.  Validate relevant header bits before exiting.
    With m_Header
        
        'Validate compression
        If LoadPCXHeader Then
            
            'PCX files should only ever use a value of "1", which means "RLE compressed".  Some files
            ' in the wild have weird values here but decode OK, so I no longer reject files based on
            ' a bad compression value.
            If (.pcx_Compression <> 1) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "WARNING: unexpected compression value found: " & .pcx_Compression
            
        End If
        
        'Validate bpp
        If LoadPCXHeader Then
            LoadPCXHeader = (.pcx_BitsPerPixel = 1) Or (.pcx_BitsPerPixel = 2) Or (.pcx_BitsPerPixel = 4) Or (.pcx_BitsPerPixel = 8) Or (.pcx_BitsPerPixel = 24) Or (.pcx_BitsPerPixel = 32)
            If (Not LoadPCXHeader) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Bad bpp: " & .pcx_BitsPerPixel
        End If
        
        'Validate x/y dimensions
        If LoadPCXHeader Then
            LoadPCXHeader = (.pcx_XMax >= .pcx_XMin)
            If (Not LoadPCXHeader) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Bad XMin/Max: " & .pcx_XMin & ", " & .pcx_XMax
        End If
        If LoadPCXHeader Then
            LoadPCXHeader = (.pcx_YMax >= .pcx_YMin)
            If (Not LoadPCXHeader) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Bad YMin/Max: " & .pcx_YMin & ", " & .pcx_YMax
        End If
        
        'Validate channel count
        If LoadPCXHeader Then
            LoadPCXHeader = (.pcx_ChannelCount = 1) Or (.pcx_ChannelCount = 3) Or (.pcx_ChannelCount = 4)
            If (Not LoadPCXHeader) Then LoadPCXHeader = (.pcx_BitsPerPixel = 1) And (.pcx_ChannelCount = 2)
            If (Not LoadPCXHeader) And PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Bad channel count: " & .pcx_ChannelCount
        End If
        
    End With
    
    'LoadPCXHeader now reflects a (pretty confident!) guess at the file being PCX format.
    
    'If the the file *isn't* PCX, reset the stream header before exiting.
    If (Not LoadPCXHeader) Then m_Stream.SetPosition initStreamPos
    Exit Function
    
'On any parse error, this function jumps to this branch and simply closes the underlying file, then exits
BadPCXFile:
    
    InternalError FUNC_NAME, "critical header failure"
    m_Stream.SetPosition initStreamPos, FILE_BEGIN
    LoadPCXHeader = False
    
End Function

'If the source file was loaded as a multi-frame DCX, this will return TRUE.
' (Note: return value is not useful if IsFilePCX() initially returned FALSE.)
Friend Function IsFileActuallyDCX() As Boolean
    IsFileActuallyDCX = m_DCXMode
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
    
    If needToValidate Then
        If (Not Me.IsFilePCX(srcFile, True, True)) Then Exit Function
    End If
    
    'If the underlying stream is open and the header validated successfully, attempt a full load
    If m_HeaderValid And (Not m_Stream Is Nothing) Then
        
        'Failsafe only; should not be possible
        If (Not m_Stream.IsOpen()) Then GoTo BadPCXFile
        
        '(Failsafe only; the pdImage object should always be initialized by the caller)
        If (dstImage Is Nothing) Then Set dstImage = New pdImage
            
        'We now have two possibilities:
        ' 1) the file is a normal PCX file and contains just one PCX frame (encompassing the whole file)
        ' 2) the file is a DCX file and contains multiple PCX frames
        '
        'The only difference between the two is how many times we iterate the PCX frame load function
        If (Not m_DCXMode) Then
            m_dcxCount = 1
            ReDim m_dcxOffsets(0) As Long
            m_dcxOffsets(0) = 0
        End If
        
        'Try loading each PCX frame in turn.  As frames are generated, we'll add it to the running pdImage object.
        Dim i As Long, numPagesAdded As Long
        numPagesAdded = 0
        
        For i = 0 To m_dcxCount - 1
            
            Set dstDIB = New pdDIB
            
            Dim posInit As Long, posEOF As Long
            posInit = m_dcxOffsets(i)
            If (i < m_dcxCount - 1) Then posEOF = m_dcxOffsets(i + 1) Else posEOF = m_Stream.GetStreamSize()
            PDDebug.LogAction "Attempting to load PCX at offset " & posInit
            If LoadPCXFrame(posInit, posEOF, dstDIB, dstImage) Then
                
                numPagesAdded = numPagesAdded + 1
                        
                'PCX files do not support color management and should be premultiplied
                ' (although non-255 alpha is extremely rare, owing to the format's age)
                dstDIB.SetColorManagementState cms_ProfileConverted
                dstDIB.SetAlphaPremultiplication True
                
                'Prep a new layer object and initialize it with the image bits we've retrieved
                Dim newLayerID As Long
                newLayerID = dstImage.CreateBlankLayer()
                
                Dim tmpLayer As pdLayer
                Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                
                'Layer name will be frame index in DCX files; filename in PCX files
                Dim nameOfLayer As String
                If (m_dcxCount = 1) Then
                    nameOfLayer = Layers.GenerateInitialLayerName(m_OpenFilename, srcImage:=dstImage)
                Else
                    nameOfLayer = g_Language.TranslateMessage("Page %1", numPagesAdded)
                End If
                tmpLayer.InitializeNewLayer PDL_Image, nameOfLayer, dstDIB, True
                
                'Fill any remaining layer properties with default values
                tmpLayer.SetLayerBlendMode BM_Normal
                tmpLayer.SetLayerOpacity 100!
                
                'In multi-frame mode, allow frames to be positioned away from (0, 0)
                If m_DCXMode Then
                    tmpLayer.SetLayerOffsetX m_Header.pcx_XMin
                    tmpLayer.SetLayerOffsetY m_Header.pcx_YMin
                Else
                    tmpLayer.SetLayerOffsetX 0
                    tmpLayer.SetLayerOffsetY 0
                End If
                
                'In multi-frame files, only the first frame is visible by default
                tmpLayer.SetLayerVisibility (numPagesAdded = 1)
                
                'Notify the layer and image of new changes, so it knows to regenerate internal caches
                ' on next access
                tmpLayer.NotifyOfDestructiveChanges
                dstImage.NotifyImageChanged UNDO_Image
                
            Else
                If PCX_DEBUG_VERBOSE Then InternalError FUNC_NAME, "Bad frame #" & i
            End If
            
        Next i
        
        LoadPCX_FromFile = (dstImage.GetNumOfLayers > 0)
        
        'Ensure the parent image size is updated against the size of the first layer
        If LoadPCX_FromFile Then dstImage.UpdateSize
        
    '/File is not PCX; silently ignore it
    End If
    
    Exit Function
    
BadPCXFile:
    InternalError FUNC_NAME, "abandoned load due to critical error"
    LoadPCX_FromFile = False
    
End Function

'Load an individual frame from a PCX file.  The caller is responsible for correctly sending the initial stream position
' and the expected EOF *for this frame*.  They also need to specify a target pdDIB object; that DIB will be filled
' by this function, if loading is successful.
'
'Returns: TRUE if load successful; FALSE otherwise
Private Function LoadPCXFrame(ByVal posInit As Long, ByVal posEOF As Long, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As Boolean
    
    Const FUNC_NAME As String = "LoadPCXFrame"
    LoadPCXFrame = False
    
    'This exists only as a failsafe; potential encoding errors are explicitly handled via other means
    On Error GoTo BadPCXFile
    
    'Additional failsafe checks
    If (m_Stream Is Nothing) Then Exit Function
    If (Not m_Stream.IsOpen()) Then Exit Function
    
    'We are now guaranteed an open stream which some previous function validated as holding PCX contents.
    
    'Starting from the current offset, load this frame's header
    m_Stream.SetPosition posInit, FILE_BEGIN
    If (Not LoadPCXHeader()) Then
        InternalError FUNC_NAME, "stream misaligned; couldn't load header"
        Exit Function
    End If
    
    'If the underlying stream is open and the header validated successfully, attempt a full load
    If m_HeaderValid Then
        
        'Forcibly reset the file pointer to offset 128 (where the header ends and pixel data begins).
        m_Stream.SetPosition posInit + 128, FILE_BEGIN
        
        'Construct a default palette object using the contents of the file header.
        ' (We may overwrite this with other data later, as necessary - v5 PCX files can store
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
        ' PCX images than a bare EGA one - so I've just stolen their palette here!)
        If (m_Header.pcx_VersionNumber < 2) Or (m_Header.pcx_VersionNumber = 3) Then Palettes.GetStockPalette pdsp_PSLegacy, m_Palette, False
        
        'I don't have a way to test this (no relevant images found "in the wild" yet),
        ' but if a PCX specifies a grayscale palette, should we explicitly populate one...?
        If (m_Header.pcx_PaletteMode = 2) Then
            If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Palette marked as grayscale; replacing colors now..."
            Palettes.GetPalette_GrayscaleEx m_Palette, 16, True
            m_IsGrayscale = True
        Else
            m_IsGrayscale = False
        End If
        
        'PCX files use max/min dimensions; convert these to normal 1-based width/height
        Dim imgWidth As Long, imgHeight As Long, imgStrideInBytes As Long
        imgWidth = (m_Header.pcx_XMax - m_Header.pcx_XMin) + 1
        imgHeight = (m_Header.pcx_YMax - m_Header.pcx_YMin) + 1
        imgStrideInBytes = m_Header.pcx_BytesPerScanline * m_Header.pcx_ChannelCount
        
        'Multi-plane images at low bit-depths (a quirk of PCX encoding) make it cumbersome to track
        ' where one plane ends and another begins due to the way PCX alignment works.  Cache the
        ' bytes-per-scanline value in a local, non-UDT value so we can access it faster inside the
        ' decoding loop.
        Dim bytesPerSingleScanline As Long
        bytesPerSingleScanline = m_Header.pcx_BytesPerScanline
        If PCX_DEBUG_VERBOSE Then
            If ((bytesPerSingleScanline And 1) = 1) Then PDDebug.LogAction "WARNING: bytes per scanline value is bad; will attempt load anyway."
            If ((m_Header.pcx_XMin <> 0) Or (m_Header.pcx_YMin <> 0)) Then PDDebug.LogAction "Note: unexpected x/y min values: " & m_Header.pcx_XMin & ", " & m_Header.pcx_YMin
            PDDebug.LogAction "Preparing to decode PCX image (" & imgWidth & "x" & imgHeight & ") @ " & m_Header.pcx_BitsPerPixel & "x" & m_Header.pcx_ChannelCount & " (" & imgStrideInBytes & ")"
        End If
        
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
            
            'Failsafe only
            Case Else
                If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "Unknown equivalent bit-depth: " & actualBPP
            
        End Select
        
        'Note the current file position and store it; we'll forcibly calculate correct offsets
        ' for each line as we go.
        Dim initFileOffset As Long, maxStreamSize As Long
        initFileOffset = m_Stream.GetPosition()
        maxStreamSize = posEOF  'In PCX files, this will be EOF; in DCX files, this will be the start of the next frame
        
        Dim xFinal As Long
        xFinal = imgWidth - 1
        
        'Before decoding, we may need to grab a trailing palette (if any).
        If (m_Header.pcx_VersionNumber = 5) Then
        
            'Here are the direct instructions from the spec on retrieving a 256-color palette:
            '   "To access a 256 color palette:
            '   "First, check the version number in the header; if it contains a 5 there is a palette."
            '   "Second, read to the end of the file and count back 769 bytes.  The value you find
            '    should be a 12 decimal, showing the presence of a 256 color palette."
            If (posEOF > 769) Then
                m_Stream.SetPosition posEOF - 769, FILE_BEGIN
                If (m_Stream.ReadByte() = 12) Then
                    If PCX_DEBUG_VERBOSE Then PDDebug.LogAction "EOF palette found; retrieving before parsing..."
                    For i = 0 To 255
                        m_Palette(i).Red = m_Stream.ReadByte()
                        m_Palette(i).Green = m_Stream.ReadByte()
                        m_Palette(i).Blue = m_Stream.ReadByte()
                        m_Palette(i).Alpha = 255
                    Next i
                End If
            End If
            
            'Before continuing, reset the file pointer to its original offset.
            m_Stream.SetPosition initFileOffset, FILE_BEGIN
            
        End If
        
        'PCX files have relatively complex palette behavior, owing to the format's age.
        ' This link was very helpful in understanding some of the nuances:
        ' https://moddingwiki.shikadi.net/wiki/PCX_Format
        If (actualBPP < 24) Then
            
            If PCX_DEBUG_VERBOSE Then
                PDDebug.LogAction "File version: " & m_Header.pcx_VersionNumber
                PDDebug.LogAction "Palette mode: " & m_Header.pcx_PaletteMode
                PDDebug.LogAction "Plane count: " & m_Header.pcx_ChannelCount
                PDDebug.LogAction "BPP: " & m_Header.pcx_BitsPerPixel
                PDDebug.LogAction "Performing palette maintenance..."
            End If
            
            'Photoshop always loads monochrome PCX files as pure black-and-white, regardless of
            ' the current palette.  GIMP only enforces this if the first two palette colors match
            ' (which appears to be a heuristic to catch wrongly encoded files...?)
            '
            'There's no right/wrong answer here, so I'm inclined to follow GIMP's lead (of not ignoring
            ' in-file information unless absolutely necessary).  Newer versions of GIMP have also added
            ' a pop-up prompt where you can choose between the PCX's embedded palette and a standard
            ' monochrome black-and-white palette.  This is such an incredibly esoteric use-case that I'm
            ' not sure I care to mimic it, but it could be added in the future as an option.
            '
            '(Note also that if the file is explicitly marked as "no embedded palette" (e.g. v3),
            ' we should ignore the embedded palette; this does potentially risk breaking some images with
            ' incorrectly marked headers, but if we use the embedded palette all the time we'll break
            ' images that are encoded correctly.  There's no right answer; "try to favor preserving
            ' correctly encoded images wherever possible" is the current strategy.)
            If (actualBPP = 1) Then
                
                'Files explicitly marked as "no embedded palette" need to manually set palette
                ' entries 0 and 1 to black and white, respectively
                If (m_Header.pcx_VersionNumber = 1) Or (m_Header.pcx_VersionNumber = 3) Then
                    m_Palette(0) = Colors.GetRGBQuadFromHex("#000000")
                    m_Palette(1) = Colors.GetRGBQuadFromHex("#FFFFFF")
                
                'Otherwise, rely on heuristics.  (Presently, if the RGB values in the first two
                ' palette entries have a difference of just 0 or 1, use black and white; I use this
                ' heuristic because some "in the wild" images inexplicably set the color values of
                ' the embedded palette to their *numerical index*, either 0/1/2/3... etc or
                ' reverse-ordered 15/14/13/12... etc.  This is presumably a glitch or deliberate
                ' flag of some sort, but it produces unusable images.)
                Else
                    If (Abs(CLng(m_Palette(0).Red) - CLng(m_Palette(1).Red)) < 2) And _
                    (Abs(CLng(m_Palette(0).Green) - CLng(m_Palette(1).Green)) < 2) And _
                    (Abs(CLng(m_Palette(0).Red) - CLng(m_Palette(1).Red)) < 2) Then
                        m_Palette(0) = Colors.GetRGBQuadFromHex("#000000")
                        m_Palette(1) = Colors.GetRGBQuadFromHex("#FFFFFF")
                    End If
                End If
                
            End If
            
            'If this image is a PCX (not a DCX), save this palette to the target image object
            ' as the "original" image palette.
            If (Not dstImage Is Nothing) And (Not m_DCXMode) Then
                Dim numPossiblePalColors As Long
                numPossiblePalColors = 2 ^ actualBPP
                If (numPossiblePalColors > 256) Then numPossiblePalColors = 256
                dstImage.SetOriginalPalette m_Palette, numPossiblePalColors
            End If
            
        End If
        
        'Initialize the destination DIB (we'll fill its contents as we go).
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank imgWidth, imgHeight, 32, 0, 255
        
        Dim tmpSA1D As SafeArray1D, dstPixels() As Byte, dstPixelsRGBQuad() As RGBQuad
        
        'The RLE compression used in PCX files operates on byte-level data (*not* pixel data).
        ' This simplifies decoding somewhat, since decompression is identical for all color-depths.
        Dim uncompressedBytes() As Byte
        ReDim uncompressedBytes(0 To imgStrideInBytes * 2) As Byte  'Extra padding included for overrun safety; only imgStrideInBytes is strictly necessary
        
        'Some bit-depth/color-plane combinations require us to decode RLE data, then shuffle bit-flags
        ' to arrive at actual palette indices.
        Dim compressedBytes() As Byte
        ReDim compressedBytes(0 To imgStrideInBytes * 2) As Byte
        
        Dim shuffledBytes() As Byte
        ReDim shuffledBytes(0 To imgWidth - 1) As Byte
        
        'Some software is capable of writing PCX files without compression (ImageMagick?).
        ' This is against-spec, but we can handle it without too much extra code.
        Dim noCompression As Boolean
        noCompression = (m_Header.pcx_Compression = 0)
        
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
        
        'I only have one test image with these settings, so I'm forced to imagine what the correct solution is!
        ' This produces a useable image so we're running with it for now...
        ElseIf (actualBPP = 16) Then
            ReDim bitFlags(0 To 7) As Byte
            bitFlags(0) = 2 ^ 7
            bitFlags(1) = 2 ^ 6
            bitFlags(2) = 2 ^ 5
            bitFlags(3) = 2 ^ 4
            bitFlags(4) = 2 ^ 3
            bitFlags(5) = 2 ^ 2
            bitFlags(6) = 2 ^ 1
            bitFlags(7) = 1
        End If
        
        'The PCX spec explicitly disallows RLE runs from crossing scanlines.  (This makes a lot of
        ' sense for the era in which PCX was designed - you only wanted a one-line buffer for pixel
        ' values as you loaded them, to minimize memory requirements.)  Unfortunately, like any
        ' old image format there are plenty of real-world images that don't follow this rule.
        '
        'We try to rescue faulty RLE runs by noting overflows and (safely) overflowing relevant bits
        ' onto the next line.  Note, however, that this solution is *not* foolproof - for security
        ' reasons, I never overflow multiple lines at once, which means that tiny images with RLE runs
        ' that span more than two lines will *not* load correctly (which I'm fine with).
        '
        'Note also that these arrays are explicitly allocated with an extra trailing line, so we
        ' don't need to check boundaries when writing "next line" values to the array.
        ReDim m_BadRLECount(0 To imgHeight) As Long
        ReDim m_BadRLEValue(0 To imgHeight) As Long
        
        'Start iterating scanlines.  We'll construct the final image one line at a time as we go.
        Dim x As Long, y As Long, j As Long, curChannel As Long, maxChannels As Long
        maxChannels = m_Header.pcx_ChannelCount - 1
        
        For y = 0 To imgHeight - 1
            
            'Start by decompressing the current scanline.  This step will migrate data from its simple
            ' RLE format to a temporary "uncompressed bytes" buffer.
            
            'For speed purposes, we always pull a worst-case number of compressed bytes into memory.
            ' (Operating on this is much faster than touching the file on each and every RLE byte.)
            Dim worstCaseBytes As Long
            worstCaseBytes = imgStrideInBytes * 2
            
            'If our worst-case size extends past the end of the file, default any remaining bytes with
            ' white (in case this is actually a malformed file) then pull the data in from file.
            If (worstCaseBytes > maxStreamSize - initFileOffset) Then
                VBHacks.FillMemory VarPtr(compressedBytes(0)), imgStrideInBytes * 2, 255
                worstCaseBytes = maxStreamSize - initFileOffset
            End If
            m_Stream.ReadBytes compressedBytes, imgStrideInBytes * 2, False
            
            'Reset all offsets
            Dim xOffset As Long, srcBytesRead As Long
            xOffset = 0
            srcBytesRead = 0
            
            'PCX files should always be compressed, but some images "in the wild" may have been
            ' encoded without compression.  Try to rescue them if found by simply grabbing the
            ' source bits as-is.
            If noCompression Then
                VBHacks.CopyMemoryStrict VarPtr(uncompressedBytes(0)), VarPtr(compressedBytes(0)), imgStrideInBytes
                srcBytesRead = imgStrideInBytes
                
            '99.99% of PCX files use simple RLE compression
            Else
                
                'Before iterating x values, pre-fill the start of the line with any "leftover" RLE
                ' runs from the previous line.  (This only occurs on broken PCX files, but doing this
                ' often allows us to successfully recover them.)
                If (m_BadRLECount(y) > 0) Then
                    
                    Dim amtToFill As Long
                    amtToFill = PDMath.Min2Int(m_BadRLECount(y), imgStrideInBytes - 1)
                    
                    For x = xOffset To xOffset + amtToFill - 1
                        uncompressedBytes(x) = m_BadRLEValue(y)
                    Next x
                    
                    xOffset = amtToFill
                    
                End If
                
                'Keep decoding until a full scanline has been processed.
                ' (Per the spec, RLE runs cannot extend across scanlines; we enforce this strictly.)
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
                            
                            If PCX_DEBUG_RESCUE_BROKEN Then
                                
                                'Store the bad run amount and value for the *next* line; we'll try to
                                ' rescue the encoding there.
                                m_BadRLECount(y + 1) = (xOffset + runLength) - imgStrideInBytes
                                m_BadRLEValue(y + 1) = curByte
                                runLength = imgStrideInBytes - xOffset
                                
                            End If
                            
                            If PCX_DEBUG_VERBOSE Then InternalError FUNC_NAME, "RLE overrun on line " & y
                            
                        End If
                        
                        'Failsafe check for null run lengths
                        If (runLength <= 0) Then
                            
                            'Technically, a run of 0 shouldn't be valid - but some real-world files
                            ' contain these and Photoshop loads them without complaint.
                            ' Let's do the same here (by simply treating this as a nop).
                            
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
                
                'Useful RLE info for broken files:
                'Debug.Print imgWidth, imgStrideInBytes, xOffset, srcBytesRead, m_Header.pcx_BytesPerScanline, m_Header.pcx_ChannelCount
                
            End If
            
            'uncompressedBytes() now contains the full data required for this scanline,
            ' *without* RLE compression but with no further processing.  Note also that there
            ' may be trailing dummy bytes based on the difference between imgWidth
            ' (which is pixel-precise) and imgBytesPerScanline (which *must* be even regardless
            ' of image dimensions, per the spec).
            Dim srcOffset As Long
            srcOffset = 0
            
            '24- and 32-bpp data require no additional processing, so we'll access direct bytes
            ' in the destination image instead of referring to palette quads.
            Dim useBareBytes As Boolean
            If (actualBPP >= 16) Then useBareBytes = True
            
            'Point a VB array at the target scanline in the final layer
            If useBareBytes Then
                dstDIB.WrapArrayAroundScanline dstPixels, tmpSA1D, y
            Else
                dstDIB.WrapRGBQuadArrayAroundScanline dstPixelsRGBQuad, tmpSA1D, y
            End If
            
            Dim numBytesThisChannel As Long
            Dim powFlag As Long
            
            'From here, pixel placement is separated by bit-depth
            If (actualBPP >= 24) Then
                
                '24/32-bpp mode is easy: iterate channels, and copy data as-is into the final pixel buffer
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
            
            'I've only found one image with this (strange!) combination, so this branch is
            ' not well-tested.  This will at least load the data as it exists in the file;
            ' whether or not that's what you'd actually expect to see is a different question...
            ElseIf (actualBPP = 16) Then
                
                powFlag = 1
                numBytesThisChannel = 0
                
                'Count pixels and iterate as we go (so we can ignore dead space at the end
                ' of individual scanlines).
                x = 0: xOffset = 0: curChannel = 0
                Do While (x < imgWidth) And (curChannel <= maxChannels)
                    
                    'Pull the full byte from the post-RLE stream and iterate bits
                    curByte = uncompressedBytes(xOffset)
                    For j = 0 To 7
                        
                        If (x < imgWidth) And (curChannel <= maxChannels) Then
                        
                            'Use our pre-built LUT for simplicity, and set the corresponding
                            ' color channel to maximum intensity (255).
                            If (bitFlags(j) = (curByte And bitFlags(j))) Then
                                dstPixels(x * 4 + curChannel) = 255
                            End If
                            
                            'Update the power flag as we rotate through channels
                            x = x + 1
                            If (x >= imgWidth) Then
                                x = 0
                                curChannel = curChannel + 1
                                powFlag = 2 ^ curChannel
                                
                                'Ensure padding is maintained if image width doesn't map cleanly
                                ' to scanline widths at this bit-depth.
                                If (numBytesThisChannel < bytesPerSingleScanline - 1) Then xOffset = xOffset + 1
                                
                                'Reset how many bytes we've read *for this channel* (to -1,
                                ' because it will be incremented again outside the for loop)
                                numBytesThisChannel = -1
                                
                                Exit For
                            End If
                            
                        Else
                            Exit For
                        End If
                        
                    Next j
                    
                    xOffset = xOffset + 1
                    numBytesThisChannel = numBytesThisChannel + 1
                    
                Loop
                
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
                        
                        powFlag = 1
                        numBytesThisChannel = 0
                        
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
                                        
                                        'Ensure padding is maintained if image width doesn't map cleanly
                                        ' to scanline widths at this bit-depth.
                                        If (numBytesThisChannel < bytesPerSingleScanline - 1) Then xOffset = xOffset + 1
                                        
                                        'Reset how many bytes we've read *for this channel* (to -1,
                                        ' because it will be incremented again outside the for loop)
                                        numBytesThisChannel = -1
                                        
                                        Exit For
                                    End If
                                    
                                Else
                                    Exit For
                                End If
                                
                            Next j
                            
                            xOffset = xOffset + 1
                            numBytesThisChannel = numBytesThisChannel + 1
                            
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
        
        LoadPCXFrame = True
        
    '/File is not PCX; silently ignore it
    End If
    
    Exit Function
    
BadPCXFile:
    InternalError FUNC_NAME, "abandoned load due to critical error"
    LoadPCXFrame = False
    
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
    
    'PCX doesn't support multiple frames.  (We could add DCX export support someday - that variant *does* support frames.)
    ' As such, grab a composited copy of the current image.
    ' (NOTE: if exporting to 32-bpp is enabled in the future, you will want to set premultiplied alpha to FALSE here.)
    Dim compositeDIB As pdDIB
    srcImage.GetCompositedImage compositeDIB, True
    
    'Export color-depth is currently limited to <= 24-bit.
    ' Note that 32-bit PCX images aren't really a "thing", given that the PCX spec was last
    ' (officially) updated in 1995.  However, some apps *will* load 32-bit PCX images.
    '
    'I'm happy to enable 32-bit output pending user feedback, but for now, PD forcibly removes alpha.
    ' (Loading a PCX, editing it, then saving it back out to file is likely the most common use-case,
    '  and that's not gonna add transparency anyway - so this may all be a moot point.)
    compositeDIB.CompositeBackgroundColor 255, 255, 255
    
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
            Palettes.SortPaletteByPopularity_RGB compositeDIB, m_Palette
            numColorsInPalette = DIBs.GetDIBAs8bpp_RGBA_SrcPalette(compositeDIB, m_Palette, pxIndexed)
            
            'NOTE: if this file was originally a PCX file, and you wanted to retain the image's original palette,
            ' you could do it like this:
            'Dim tmpPDPalette As pdPalette
            'srcImage.GetOriginalPalette tmpPDPalette
            'tmpPDPalette.CopyPaletteToArray m_Palette
            'numColorsInPalette = DIBs.GetDIBAs8bpp_RGBA_SrcPalette(compositeDIB, m_Palette, pxIndexed)
            
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
    ' breaking for the next line.  (This amount is non-negotiable, even though it potentially adds empty
    ' bytes at the end of the RLE stream since PCXs enforce a strict rule of "scanlines must be even".)
    Dim imgStrideInBytes As Long
    imgStrideInBytes = bpScanline * channelCount
    
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = compositeDIB.GetDIBWidth
    imgHeight = compositeDIB.GetDIBHeight
    
    'Remember: image width needs to be padded to an even number.  (Some PCXs "in the wild"
    ' may even use 4-byte alignment; they probably relied on a system-level macro for alignment.)
    Dim imgWidthChannelPadded As Long
    imgWidthChannelPadded = imgWidth
    If ((imgWidthChannelPadded And 1) = 1) Then imgWidthChannelPadded = imgWidthChannelPadded + 1
    
    'The RLE compression used in PCX files operates on byte-level data (*not* pixel data).
    ' This simplifies encoding somewhat, since compression is identical for all color-depths.
    ' All we have to do first is a small "pre-processing" step to organize pixel data by plane,
    ' and then we can run the RLE compression scheme uniformly against the data without regard
    ' to its source format.
    Dim uncompressedBytes() As Byte
    ReDim uncompressedBytes(0 To imgStrideInBytes) As Byte
    
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
                uncompressedBytes(imgWidthChannelPadded + x) = srcPixels(x * 4 + 1)
                uncompressedBytes(imgWidthChannelPadded * 2 + x) = srcPixels(x * 4)
                If (outputColorDepth = 32) Then uncompressedBytes(imgWidthChannelPadded * 3 + x) = srcPixels(x * 4 + 3)
            Next x
            
            'Note the index of the last item added to the output array.  We can potentially add a
            ' matching dummy byte after this (if the array has empty space) to save a little file space
            ' during RLE compression.
            If (outputColorDepth = 32) Then
                idxLast = imgWidthChannelPadded * 4 - 1
            Else
                idxLast = imgWidthChannelPadded * 3 - 1
            End If
            
            'Pad each color channel as relevant, so that padding bytes don't inadvertently add an extra
            ' RLE packet (because they differ from the previous pixel).
            If (imgWidth < imgWidthChannelPadded) Then
                uncompressedBytes(imgWidthChannelPadded - 1) = uncompressedBytes(imgWidthChannelPadded - 2)
                uncompressedBytes(imgWidthChannelPadded * 2 - 1) = uncompressedBytes(imgWidthChannelPadded * 2 - 2)
                uncompressedBytes(imgWidthChannelPadded * 3 - 1) = uncompressedBytes(imgWidthChannelPadded * 3 - 2)
                If (outputColorDepth = 32) Then uncompressedBytes(imgWidthChannelPadded * 4 - 1) = uncompressedBytes(imgWidthChannelPadded * 4 - 2)
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
        ' to a trailing RLE packet matching the last meaningful byte.  (Note that 24/32-bpp have already
        ' been handled; they require special work because each channel has its own padding.)
        If (idxLast < imgStrideInBytes - 1) And (outputColorDepth <= 8) Then
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
                
                'See how many times we can repeat this byte.  Note that we must stop repeating under
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
