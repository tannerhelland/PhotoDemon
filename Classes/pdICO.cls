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
'Copyright 2020-2026 by Tanner Helland
'Created: 07/May/20
'Last updated: 28/October/22
'Last update: improve alpha thresholding behavior for better color quality when creating tiny icon sizes
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
Private Const ICO_DEBUG_VERBOSE As Boolean = False

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
    Size As Long
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
    
    'The following settings are *only* used at export time
    
    'Export width/height; these need to be *precise*, even if larger than 255
    ico_ExportWidth As Long
    ico_ExportHeight As Long
    
    'Source layer index to use when generating mask + ico bits
    ico_ExportLayerIndex As Long
    
End Type

Private m_NumIcons As Long
Private m_Icons() As PD_Icon

'If the user wants us to use the merged image as source, we'll generate it just once, prior to export.
' (Note that this *will* be empty if the user requests "match each layer to the ideal icon frame" mode.)
Private m_mergedImage As pdDIB

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
    
    SaveICO_ToFile = False
    
    'All export settings are found in the param string
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString fullParamString
    
    'Does the user want to use the merged image as their source, or do they want us to match each layer
    ' to its most appropriate output frame?
    Dim useMergedImage As Boolean
    useMergedImage = cParams.GetBool("use-merged-image", False, True)
    
    '(regardless of user settings, use the merged image if this is a single-layer image)
    If (Not srcPDImage Is Nothing) Then
        useMergedImage = useMergedImage Or (srcPDImage.GetNumOfLayers <= 1)
    Else
        useMergedImage = False
    End If
    
    'Generate a merged image, as appropriate
    If useMergedImage Then srcPDImage.GetCompositedImage m_mergedImage, True
    
    'How many frames are we writing?
    Dim outFrames() As PD_Icon, numFrames As Long
    numFrames = cParams.GetLong("icon-count", 0, True)
    If (numFrames <= 0) Then
        InternalError "SaveICO_ToFile", "zero icon count"
        SaveICO_ToFile = False
        Exit Function
    End If
    
    'numFrames is now guaranteed > 0
    ReDim outFrames(0 To numFrames - 1) As PD_Icon
    
    'Build a list of required frame sizes and color-depths.
    Dim i As Long, j As Long, numSettings As Long, icoSettings() As Long
    For i = 1 To numFrames
    
        'A properly formatted settings string must supply three values: width, height, and color-depth (PNG = 64)
        numSettings = Strings.SplitIntegers(cParams.GetString("ico-" & i, vbNullString, True), icoSettings, False)
        If (numSettings = 3) Then
            With outFrames(i - 1)
                .ico_ExportWidth = icoSettings(0)
                .ico_ExportHeight = icoSettings(1)
                .ico_BPP = icoSettings(2)
                If (.ico_BPP > 32) Then
                    .ico_IsPNG = True
                    .ico_BPP = 32
                End If
            End With
        Else
            'nop; param string is bad!
        End If
    Next i
    
    'Next, we need to match the list of requested frames against layer indices in the source image.
    ' We want to match (as best we can) idealized layer sizes and color-counts against frame requirements.
    ' We also give preferences to layers whose names exactly match requested export settings; that means
    ' the original file was likely an ICO, and we're just re-saving it after edits.
    '
    '(The user can override this behavior either explicitly - by requesting that we use the merged image -
    ' or implicitly, if the source image only has one layer.)
    If (Not useMergedImage) Then
        
        'To avoid the need for determining layer color-depth multiple times (because PD stores everything
        ' as 32-bpp RGBA internally, and we need to know more detailed color counts), we cache layer data
        ' as we calculate it.  (Note that in the case of perfect matches between layer names and export
        ' requirements, we defer to that - this gives the user a way to "tell us" which layer to use for
        ' which export.)
        For i = 0 To numFrames - 1
            
            Dim tWidth As Long, tHeight As Long, tColorDepth As Long, tPNG As Boolean
            Dim lName As String, matchFound As Boolean, lyrIsPNG As Boolean
            tWidth = outFrames(i).ico_ExportWidth
            tHeight = outFrames(i).ico_ExportHeight
            tColorDepth = outFrames(i).ico_BPP
            tPNG = outFrames(i).ico_IsPNG
            If tPNG Then tColorDepth = 64
            
            matchFound = False
            
            'Look for any layers whose layer name matches the requested export settings
            For j = 0 To srcPDImage.GetNumOfLayers() - 1
            
                lName = srcPDImage.GetLayerByIndex(j).GetLayerName()
                numSettings = Strings.SplitIntegers(lName, icoSettings, False)
            
                'Note that we only look for 2 settings (width and height).  PNG-encoded frames
                ' won't return a numeric color-depth; we'll check this special case in a moment.
                If (numSettings >= 2) Then
                    
                    'Compare size numbers; if they match exactly, great!
                    If (icoSettings(0) = tWidth) And (icoSettings(1) = tHeight) Then
                    
                        'Sizes match.  Compare color-depths next.
                        If (numSettings > 2) Then
                            matchFound = (icoSettings(2) = tColorDepth)
                        Else
                            lyrIsPNG = (Strings.StrStrI(StrPtr(lName), StrPtr("PNG")) > 0)
                            matchFound = (lyrIsPNG And tPNG)
                        End If
                        
                    End If
                    
                    'If we found an exact match, mark it and continue with the next frame.
                    If matchFound Then
                        If ICO_DEBUG_VERBOSE Then PDDebug.LogAction "Matched layer name for frame " & i & " and layer " & j
                        outFrames(i).ico_ExportLayerIndex = j
                        GoTo NextFrame  'VB6 doesn't provide a Continue statement; sorry for the goto!
                    End If
                    
                End If
            
            Next j
            
            'If we're still here, an exact layer name match was *not* found for this frame.
            ' That means we need to perform heuristics on source layers.
            Dim closestSize As Long, closestIndex As Long
            closestIndex = -1
            closestSize = &HFFFFFFF
            
            Dim testSize As Long, compareSize As Long
            
            'The goal here is to find the layer that is closest in size to the requested icon size.
            For j = 0 To srcPDImage.GetNumOfLayers() - 1
                testSize = PDMath.Max2Int(srcPDImage.GetLayerByIndex(j).GetLayerWidth(True), srcPDImage.GetLayerByIndex(j).GetLayerHeight(True))
                compareSize = Abs(testSize - PDMath.Max2Int(outFrames(i).ico_ExportWidth, outFrames(i).ico_ExportHeight))
                If (compareSize < closestSize) Then
                    closestIndex = j
                    closestSize = compareSize
                End If
            Next j
            
            If ICO_DEBUG_VERBOSE Then PDDebug.LogAction "Approximating frame " & i & " with source layer " & closestIndex
            
            'TODO: factor layer color-depth into the decision tree??
            
            'Set the nearest-size index as the target index
            outFrames(i).ico_ExportLayerIndex = closestIndex
NextFrame:
        Next i
        
    End If
    
    'Every frame is now associated with a layer in the original image.  We have what we need to write
    ' our icon file!
    Set m_Stream = New pdStream
    If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFile) Then
    
        'Start with a standard header.
        ' 2-bytes: 0 (reserved)
        ' 2-bytes: 1 (for icon), 2 (for cursor), all other values are invalid
        ' 2-bytes: number of images in the file (must be > 0)
        m_Stream.WriteInt 0
        m_Stream.WriteInt 1
        m_Stream.WriteInt numFrames
        
        'Next, we need to write headers for each frame.  This requires us to know the size of each
        ' frame's pixel + mask (or PNG data), so we need to calculate all frames in advance.
        Dim tmpDIB As pdDIB
        For i = 0 To numFrames - 1
            CreateFrameForExport outFrames(i), srcPDImage.GetLayerByIndex(outFrames(i).ico_ExportLayerIndex), srcPDImage, tmpDIB, useMergedImage
        Next i
        
        '6-byte header + (numFrames) * 16-byte frame header = initial offset into bitmap data
        Dim curOffset As Long
        curOffset = 6& + numFrames * 16
        
        'With all frame data assembled, we can now dump all headers to file!
        For i = 0 To numFrames - 1
        
            'Headers are fixed in size; interpretation does vary slightly for icons vs cursors
            With outFrames(i)
                
                'Width/height can be 0; this means the frame is >= 256 pixels in that dimension.
                ' (Note that the precise size is listed in the BITMAPINFOHEADER for this frame,
                ' which uses normal 4-byte integers for dimensions.)
                If (.ico_ExportWidth > 255) Then m_Stream.WriteByte 0 Else m_Stream.WriteByte .ico_ExportWidth
                If (.ico_ExportHeight > 255) Then m_Stream.WriteByte 0 Else m_Stream.WriteByte .ico_ExportHeight
                
                'Color count in palette (byte)
                If (.ico_BPP = 1) Then
                    m_Stream.WriteByte 2
                ElseIf (.ico_BPP = 2) Then
                    m_Stream.WriteByte 4
                ElseIf (.ico_BPP = 4) Then
                    m_Stream.WriteByte 16
                Else
                    m_Stream.WriteByte 0
                End If
                
                'Reserved byte; must be zero, with the following caveat from Wikipedia (https://en.wikipedia.org/wiki/ICO_(file_format)):
                ' "Although Microsoft's technical documentation states that this value must be zero,
                ' the icon encoder built into .NET (System.Drawing.Icon.Save) sets this value to 255.
                ' It appears that the operating system ignores this value altogether."
                m_Stream.WriteByte 0
                
                'Color planes; must be 1 (2-bytes)
                m_Stream.WriteIntU 1
                
                'Bits-per-pixel
                If .ico_IsPNG Then
                    m_Stream.WriteIntU 32
                Else
                    m_Stream.WriteIntU .ico_BPP
                End If
                
                'Size of the image data, in bytes (4-bytes)
                m_Stream.WriteLong .ico_SizeInBytes
                
                'Starting offset of the data, in bytes
                m_Stream.WriteLong curOffset
                
                'Advance the offset automatically
                curOffset = curOffset + .ico_SizeInBytes
                
            End With
            
        Next i
        
        'With all headers written, we can now write image data out to file.  Note that it was pre-filled
        ' by the CreateFrameForExport function, above.
        For i = 0 To numFrames - 1
            m_Stream.WriteByteArray outFrames(i).ico_RawData, outFrames(i).ico_SizeInBytes
        Next i
        
        'We're done!  Close the stream and exit
        m_Stream.StopStream
        SaveICO_ToFile = True
    
    End If
    
End Function

'Given a PD_Icon struct, populate the ico_RawData() member with a valid icon stream, based on the params
' already filled in relevant PD_Icon members.  To reduce memory churn, a persistent pdDIB object is also required;
' PD will use this to generate intermediate frame copies.
'
'NOTE: at present, this function only produces square thumbnails (by design).  Odd sizes may be accomodated
' in the future.
Private Function CreateFrameForExport(ByRef srcIcoData As PD_Icon, ByRef srcLayer As pdLayer, ByRef srcPDImage As pdImage, ByRef tmpDIB As pdDIB, ByVal useMergedImage As Boolean) As Boolean
    
    'Before doing anything else, we need to generate a reference 32-bpp image of the source layer
    If (tmpDIB Is Nothing) Then Set tmpDIB = New pdDIB
    
    'For now, PD forces icon width and height to be identical.  This could be changed in the future,
    ' but for now, I prefer to write icons with a minimal chance of incompatibility.
    Dim icoSizeEqual As Long
    icoSizeEqual = PDMath.Max2Int(srcIcoData.ico_ExportWidth, srcIcoData.ico_ExportHeight)
    
    'Next, a quick note - PD's layer objects have an internal thumbnail renderer that produces
    ' nice layer thumbnails for PD's UI (like the layer toolbox), a la:
    ' srcLayer.RequestThumbnail tmpDIB, icoSizeEqual, True
    ' Unfortunately, that internal function uses a quality parameter controllable by the user
    ' (via Tools > Settings), which deliberately keeps the source layerat its original
    ' orientation, instead of modifying it by any/all non-destructive transforms.  As such,
    ' it cannot be used here.  We must manually produce a highest-possible-quality version
    ' of the image at the requested size.  The complexity of this task depends on how severely
    ' the source layer has been modified.
    Dim referenceDIB As pdDIB
    If useMergedImage Then
        
        'Make a local copy of the merged image
        Set referenceDIB = New pdDIB
        referenceDIB.CreateFromExistingDIB m_mergedImage
        
    Else
        
        If srcLayer.AffineTransformsActive Then
        
            'We need to produce an affine-transformed version of this layer as our reference DIB.
            ' (Note that layer offsets don't matter; we will use the entire source layer as-is.)
            Dim irrelevantX As Long, irrelevantY As Long
            srcLayer.GetAffineTransformedDIB referenceDIB, irrelevantX, irrelevantY, srcPDImage.Width, srcPDImage.Height
        
        'The source layer does *not* have any active affine transforms.  This makes it much faster
        ' (and simpler!) to produce an icon-sized DIB, because all our coordinate math involves
        ' rectangles.
        Else
            Set referenceDIB = srcLayer.GetLayerDIB
        End If
        
    End If
    
    'Perform a quick shortcut check.  If the reference DIB is identical in size to the requested
    ' icon frame size, simply copy over the reference as-is (instead of attempting to resize it)
    If (icoSizeEqual = referenceDIB.GetDIBWidth) And (icoSizeEqual = referenceDIB.GetDIBHeight) Then
        tmpDIB.CreateFromExistingDIB referenceDIB
        
    'If the sizes don't match, we need to perform a manual resize
    Else
        
        'Create the icon
        tmpDIB.CreateBlank icoSizeEqual, icoSizeEqual, 32, 0, 0
        tmpDIB.SetInitialAlphaPremultiplicationState True
        
        'Figure out where to position the DIB when we draw it
        Dim newWidth As Long, newHeight As Long
        PDMath.ConvertAspectRatio referenceDIB.GetDIBWidth, referenceDIB.GetDIBHeight, icoSizeEqual, icoSizeEqual, newWidth, newHeight, True
        
        'Render the source layer into the destination thumbnail at maximum quality
        Dim dstX As Long, dstY As Long
        dstX = (icoSizeEqual - newWidth) \ 2
        dstY = (icoSizeEqual - newHeight) \ 2
        
        'In the past, we used GDI+ for this step, which only requires the following line of code:
        'GDI_Plus.GDIPlus_StretchBlt tmpDIB, dstX, dstY, newWidth, newHeight, referenceDIB, 0, 0, referenceDIB.GetDIBWidth, referenceDIB.GetDIBHeight, interpolationType:=GP_IM_HighQualityBicubic, dstCopyIsOkay:=True
        
        '...But we can get much higher quality using our internal resampling engine.
        ' (The downside of this is that I haven't added arbitrary x/y offset support to PD's internal resampler,
        '  so any offsets need to be handled in a separate step.)
        Dim tmpHolderDIB As pdDIB
        Set tmpHolderDIB = New pdDIB
        tmpHolderDIB.CreateBlank newWidth, newHeight, 32, 0, 0
        tmpHolderDIB.SetInitialAlphaPremultiplicationState True
        
        'Perform the resample
        If (Not referenceDIB.GetAlphaPremultiplication()) Then referenceDIB.SetAlphaPremultiplication True
        Resampling.ResampleImageI tmpHolderDIB, referenceDIB, newWidth, newHeight, rf_Lanczos, False
        
        'Blt the final result into place
        GDI.BitBltWrapper tmpDIB.GetDIBDC, dstX, dstY, newWidth, newHeight, tmpHolderDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Free any temporary DIBs
        Set tmpHolderDIB = Nothing
        
    End If
    
    'The reference DIB is no longer required.  It may be large (especially if we had to calculate
    ' affine transforms to produce it) so free it immediately
    Set referenceDIB = Nothing
    
    'Previously, we would now un-premultiply alpha - but some preprocessing steps are actually
    ' faster with premultiplied alpha.  As such, we don't want to un-premultiply alpha just yet!
    
    'PD uses top-down DIBs internally.  Reverse scanlines before writing out to file.
    ' (But *only* for non-PNGs!  PD's PNG engine handles this automatically.)
    If (Not srcIcoData.ico_IsPNG) Then DIBs.ReverseScanlines tmpDIB
    
    'tmpDIB now contains a 32-bpp copy of the source layer.
    
    'If this frame is in PNG format, we want to skip any further processing and instead,
    ' immediately produce a PNG stream for the frame.
    
    'NOTE: the closest thing to an "official" spec from Microsoft on PNG-embedded frames is Raymond Chen's
    ' series of articles on icons (https://devblogs.microsoft.com/oldnewthing/20101018-00/?p=12513)
    ' In part 4, he explicitly states: "The format of a PNG-compressed image consists simply of a PNG image,
    ' starting with the PNG file signature. The image must be in 32bpp ARGB format."
    '
    'PD follows this recommendation and *always* produces a 32-bpp PNG file, regardless of possible other
    ' color-depths.
    If srcIcoData.ico_IsPNG Then
        
        'Un-premultiply alpha before handing the image off to the PNG engine
        If tmpDIB.GetAlphaPremultiplication() Then tmpDIB.SetAlphaPremultiplication False
        
        Dim cPNG As pdPNG
        Set cPNG = New pdPNG
        srcIcoData.ico_SizeInBytes = cPNG.SavePNG_ToMemory(srcIcoData.ico_RawData, tmpDIB, Nothing, png_TruecolorAlpha, 8, 12, vbNullString, png_FilterOptimal, True)
        CreateFrameForExport = (srcIcoData.ico_SizeInBytes > 0)
        If ICO_DEBUG_VERBOSE Then PDDebug.LogAction "Embedded PNG frame size: " & srcIcoData.ico_SizeInBytes
        
    'The embedded frame is *not* a PNG.  We need to generate old-fashioned icon data ourselves.
    Else
        
        'Start by creating a pdStream object.  First we'll write a standard bitmap header,
        ' then the actual frame pixel data (in whatever color format the user specified).
        ' Last follows the mandatory alpha mask.
        Dim tmpStream As pdStream
        Set tmpStream = New pdStream
        tmpStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
        
        'Start with a standard BITMAPINFOHEADER struct
        tmpStream.WriteLong 40&                     'Struct size
        tmpStream.WriteLong tmpDIB.GetDIBWidth      'Width
        tmpStream.WriteLong tmpDIB.GetDIBHeight * 2 'Height * 2 (due to mask)
        tmpStream.WriteIntU 1&                      'Planes (always 1)
        tmpStream.WriteIntU srcIcoData.ico_BPP      'BPP
        tmpStream.WriteLong 0&                      'Compression (always 0 - None)
        tmpStream.WriteLong 0&                      'Image size  (spec allows 0 when compression is 0)
        tmpStream.WriteLong 0&                      'Horizontal resolution (not used for icon frames)
        tmpStream.WriteLong 0&                      'Vertical resolution (not used for icon frames)
        
        'Palette size; per MSDN, "If biClrUsed is zero, the array contains the maximum number of colors
        ' for the given bitdepth; that is, 2^biBitCount colors."
        '
        'After some testing, I found that values other than 0 are not always handled correctly
        ' by other ICO readers.  As such, we always write a max palette size for the current frame
        ' depth, even if the frame uses fewer colors.
        tmpStream.WriteLong 0&
        tmpStream.WriteLong 0&  'Important color count (set to 0 for "all colors are important")
        
        'Next, if a palette exists, we need to write it - see the color-depth switches
        ' below for details.
        Dim x As Long, y As Long, tmpSA As SafeArray1D, pxData() As Byte, dstOffsetX As Long
        Dim pxScanline() As Byte, scanlineSize As Long, scanlinePadding As Long
        
        Dim pxWidth As Long, pxHeight As Long
        pxWidth = tmpDIB.GetDIBWidth
        pxHeight = tmpDIB.GetDIBHeight
        
        Dim numColors As Long, imgPalette() As RGBQuad, maskBytes() As Byte
        
        'Regardless of color-depth, we will always write a trailing 1-bpp monochrome mask for
        ' this frame (yes, even for 32-bpp frames).  Prep a lookup table in advance.
        Dim oneBppLUT(0 To 7) As Byte
        oneBppLUT(0) = 128
        oneBppLUT(1) = 64
        oneBppLUT(2) = 32
        oneBppLUT(3) = 16
        oneBppLUT(4) = 8
        oneBppLUT(5) = 4
        oneBppLUT(6) = 2
        oneBppLUT(7) = 1
        
        'Do any preliminary work to the DIB (based on color-depth requirements).
        Select Case srcIcoData.ico_BPP
        
            'No work required; write the pixel data as-is!
            Case 32
                If tmpDIB.GetAlphaPremultiplication() Then tmpDIB.SetAlphaPremultiplication False
                tmpStream.WriteBytesFromPointer tmpDIB.GetDIBPointer, 4 * tmpDIB.GetDIBWidth * tmpDIB.GetDIBHeight
            
            'Write RGB data, but *not* alpha data
            Case 24
                
                'Before writing RGB data, threshold alpha (e.g. set it to fully opaque or fully transparent).
                ' We do this because we need to composite semi-transparent pixels against a matte color
                ' (white by default) *before* writing them out to file.
                
                'Note also that alpha premultiplication of the pixel data doesn't matter after alpha
                ' thresholding, because the only alpha values left in the file will be 0 and 255 which
                ' produce identical RGB data regardless of alpha premultiplication.
                DIBs.ThresholdAlphaChannel tmpDIB, 127, PDDM_FloydSteinberg, 50!, RGB(255, 255, 255), True
                scanlineSize = (pxWidth * 3 + 3) And &HFFFFFFFC
                ReDim pxScanline(0 To scanlineSize - 1) As Byte
                
                For y = 0 To pxHeight - 1
                    tmpDIB.WrapArrayAroundScanline pxData, tmpSA, y
                    
                    'This FillMemory technically isn't necessary (as all data will be overwritten anyway),
                    ' but I leave it as a reminder that we can't blindly copy pixel data into the
                    ' icon stream - we need to convert it to a properly padded scanline first, with a
                    ' width that's a multiple of 4.
                    FillMemory VarPtr(pxScanline(0)), scanlineSize, 0
                    For x = 0 To pxWidth - 1
                        pxScanline(x * 3) = pxData(x * 4)
                        pxScanline(x * 3 + 1) = pxData(x * 4 + 1)
                        pxScanline(x * 3 + 2) = pxData(x * 4 + 2)
                    Next x
                    tmpStream.WriteBytesFromPointer VarPtr(pxScanline(0)), scanlineSize
                    
                Next y
                
                tmpDIB.UnwrapArrayFromDIB pxData
                
            'Next come indexed (palette) pixel formats.
            
            ' For compatibility reasons, note that PD always writes a full palette
            ' (e.g. 256-colors for an 8-bpp image), even if fewer colors are actually used.
            ' Many ICO readers assume a full palette is present, and this improves compatibility.
            Case 8
            
                'Before writing RGB data, threshold alpha (e.g. set it to fully opaque or fully transparent).
                ' Note that we attempt to dither alpha if the icon is larger than 32x32; this can improve
                ' visual quality somewhat, but below that size, it just devolves into noise.
                If (tmpDIB.GetDIBWidth > 32) Then
                    DIBs.ThresholdAlphaChannel tmpDIB, 127, PDDM_FloydSteinberg, 50!, RGB(255, 255, 255), True
                Else
                    DIBs.ThresholdAlphaChannel tmpDIB, 127, PDDM_None, 0!, RGB(255, 255, 255), True
                End If
                
                'The returned image will always force transparent pixels to black, so premultiplication
                ' is irrelevant.  Set the corresponding flag, however, to prevent the palettizer from
                ' attempting to unpremultiply first.
                tmpDIB.SetInitialAlphaPremultiplicationState False
                
                'Next, retrieve separate palettized and mask channels, as well as an optimized palette
                numColors = DIBs.GetDIBAs8bpp_RGBMask_Forcibly(tmpDIB, imgPalette, pxData, maskBytes, 256)
                
                'Regardless of color count, write the full 256-color table out to file
                tmpStream.WriteBytesFromPointer VarPtr(imgPalette(0)), numColors * 4
                If (numColors < 256) Then tmpStream.WritePadding (256 - numColors) * 4
                
                'Ensure scanline width is a multiple of 4
                scanlineSize = (pxWidth + 3) And &HFFFFFFFC
                scanlinePadding = scanlineSize - pxWidth
                
                'Write each scanline out to file, with extra padding as necessary
                For y = 0 To pxHeight - 1
                    tmpStream.WriteBytesFromPointer VarPtr(pxData(0, y)), pxWidth
                    If (scanlinePadding > 0) Then tmpStream.WritePadding scanlinePadding
                Next y
            
            Case 4
            
                'Before writing RGB data, threshold alpha (e.g. set it to fully opaque or fully transparent).
                ' Note that we attempt to dither alpha if the icon is larger than 32x32; this can improve
                ' visual quality somewhat, but below that size, it just devolves into noise.
                If (tmpDIB.GetDIBWidth > 32) Then
                    DIBs.ThresholdAlphaChannel tmpDIB, 127, PDDM_FloydSteinberg, 50!, RGB(255, 255, 255), True
                Else
                    DIBs.ThresholdAlphaChannel tmpDIB, 127, PDDM_None, 0!, RGB(255, 255, 255), True
                End If
                
                'Next, retrieve separate palettized and mask channels (as well as an optimized palette)
                tmpDIB.SetInitialAlphaPremultiplicationState False
                numColors = DIBs.GetDIBAs8bpp_RGBMask_Forcibly(tmpDIB, imgPalette, pxData, maskBytes, 16)
                
                'Regardless of color count, write a full 16-color table out to file
                tmpStream.WriteBytesFromPointer VarPtr(imgPalette(0)), numColors * 4
                If (numColors < 16) Then tmpStream.WritePadding (16 - numColors) * 4
                
                'Ensure scanline width is a multiple of 4
                scanlineSize = (pxWidth + 1) \ 2
                scanlineSize = (scanlineSize + 3) And &HFFFFFFFC
                ReDim pxScanline(0 To scanlineSize - 1) As Byte
                
                'Downsample source pixels to 4-bpp, then write each scanline out to file
                For y = 0 To pxHeight - 1
                    FillMemory VarPtr(pxScanline(0)), scanlineSize, 0
                    For x = 0 To pxWidth - 1
                        dstOffsetX = x \ 2
                        If (x And 1) Then
                            pxScanline(dstOffsetX) = pxScanline(dstOffsetX) Or pxData(x, y)
                        Else
                            pxScanline(dstOffsetX) = pxData(x, y) * 16
                        End If
                    Next x
                    tmpStream.WriteBytesFromPointer VarPtr(pxScanline(0)), scanlineSize
                Next y
                
            Case 1
        
                'Before writing RGB data, threshold alpha (e.g. set it to fully opaque or fully transparent).
                ' Note that we attempt to dither alpha if the icon is larger than 32x32; this can improve
                ' visual quality somewhat, but below that size, it just devolves into noise.
                If (tmpDIB.GetDIBWidth > 32) Then
                    DIBs.ThresholdAlphaChannel tmpDIB, 127, PDDM_FloydSteinberg, 50!, RGB(255, 255, 255), True
                Else
                    DIBs.ThresholdAlphaChannel tmpDIB, 127, PDDM_None, 0!, RGB(255, 255, 255), True
                End If
                
                'Next, retrieve separate palettized and mask channels (as well as an optimized palette)
                tmpDIB.SetInitialAlphaPremultiplicationState False
                numColors = DIBs.GetDIBAs8bpp_RGBMask_Forcibly(tmpDIB, imgPalette, pxData, maskBytes, 2)
                
                'Write a corresponding color table out to file
                tmpStream.WriteBytesFromPointer VarPtr(imgPalette(0)), numColors * 4
                If (numColors < 2) Then tmpStream.WritePadding (2 - numColors) * 4
                
                'Ensure scanline width is a multiple of 4
                scanlineSize = (pxWidth + 7) \ 8
                scanlineSize = (scanlineSize + 3) And &HFFFFFFFC
                ReDim pxScanline(0 To scanlineSize - 1) As Byte
                
                'Downsample source pixels to 4-bpp, then write each scanline out to file
                For y = 0 To pxHeight - 1
                    FillMemory VarPtr(pxScanline(0)), scanlineSize, 0
                    For x = 0 To pxWidth - 1
                        dstOffsetX = x \ 8
                        If (pxData(x, y) > 0) Then pxScanline(dstOffsetX) = pxScanline(dstOffsetX) Or oneBppLUT(x And &H7&)
                    Next x
                    tmpStream.WriteBytesFromPointer VarPtr(pxScanline(0)), scanlineSize
                Next y
                
        End Select
        
        'With pixel data written successfully, pull transparency data from the original DIB, threshold it
        ' (only for 32-bit - other bit-depths will have already done this!) - then write a 1-bpp mask.
        If (srcIcoData.ico_BPP = 32) Then DIBs.ThresholdAlphaChannel tmpDIB, 127, PDDM_None, 0!, RGB(255, 255, 255), True
        
        'Masks are written a scanline at a time - and importantly, each scanline *must* be padded a multiple of 4!
        scanlineSize = (pxWidth + 7) \ 8
        scanlineSize = (scanlineSize + 3) And &HFFFFFFFC
        ReDim pxScanline(0 To scanlineSize - 1) As Byte
        
        For y = 0 To pxHeight - 1
            
            tmpDIB.WrapArrayAroundScanline pxData, tmpSA, y
            FillMemory VarPtr(pxScanline(0)), scanlineSize, 0
            
            For x = 0 To pxWidth - 1
                dstOffsetX = x \ 8
                If (pxData(x * 4 + 3) < 128) Then pxScanline(dstOffsetX) = pxScanline(dstOffsetX) Or oneBppLUT(x And &H7&)
            Next x
            
            'Write the finished mask out to file
            tmpStream.WriteBytesFromPointer VarPtr(pxScanline(0)), scanlineSize
            
        Next y
        
        tmpDIB.UnwrapArrayFromDIB pxData
        
        'Note the final stream size, copy the stream data into a persistent array, then exit!
        Dim copySize As Long
        copySize = tmpStream.GetStreamSize()
        srcIcoData.ico_SizeInBytes = copySize
        
        ReDim srcIcoData.ico_RawData(0 To copySize - 1) As Byte
        CopyMemoryStrict VarPtr(srcIcoData.ico_RawData(0)), tmpStream.Peek_PointerOnly(0, copySize), copySize
        
        tmpStream.StopStream True
        
    End If
    
End Function

'Given an index into the m_Icons() array, produce a finished DIB for said index's raw icon data
Private Function CreateDIBForIndex(ByRef srcHeader As BITMAPINFOHEADER, ByRef srcStream As pdStream, ByVal srcIndex As Long) As Boolean
    
    Dim i As Long
    
    'If the underlying image requires a palette, it immediately follows the header.
    ' TODO: maybe consider retrieving color table sizes from the BIH itself?  I've never seen
    ' an ICO that supplied a non-normal palette size, but it is theoretically possible.
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
                headerValid = (tmpBIHeader.Size = 40)
                If (Not headerValid) Then InternalError "LoadICO_GenerateFrames", "bad header size: " & tmpBIHeader.Size
                headerValid = (tmpBIHeader.Width > 0)
                If (Not headerValid) Then InternalError "LoadICO_GenerateFrames", "bad header width: " & tmpBIHeader.Width
                headerValid = (tmpBIHeader.Height > 0)
                If (Not headerValid) Then InternalError "LoadICO_GenerateFrames", "bad header height: " & tmpBIHeader.Height
                headerValid = (tmpBIHeader.Planes = 1)
                If (Not headerValid) Then InternalError "LoadICO_GenerateFrames", "bad header planes: " & tmpBIHeader.Planes
                headerValid = (tmpBIHeader.BitCount > 0)
                
                'Sometimes, the top-level icon header will supply a null bit-depth (which is bad
                ' and against spec, because it makes selecting the "best" icon problematic in
                ' e.g. Explorer, but such are Windows icons!).  If this occurs, overwrite that
                ' null bit-depth with this one, retrieved from the pixel stream.
                If headerValid Then
                    If (m_Icons(i).ico_BPP = 0) Then m_Icons(i).ico_BPP = tmpBIHeader.BitCount
                Else
                    InternalError "LoadICO_GenerateFrames", "bad header bitcount: " & tmpBIHeader.BitCount
                End If
                
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
    
    'Start by finding the largest frame in the file; we'll use this for the parent image dimensions
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
            If m_Icons(i).ico_OK And (Not m_Icons(i).ico_DIB Is Nothing) Then
                If (m_Icons(i).ico_DIB.GetDIBWidth = maxWidth) And (m_Icons(i).ico_DIB.GetDIBHeight = maxHeight) Then
                
                    'This layer matches the largest layer size we have so far.  If it *also* has the
                    ' highest bit-depth, flag it as the new active index.
                    If (m_Icons(i).ico_BPP > highestBitDepth) Or (m_Icons(i).ico_IsPNG) Then
                        highestBitDepth = m_Icons(i).ico_BPP
                        If m_Icons(i).ico_IsPNG Then highestBitDepth = 48   'Give PNGs an arbitrarily large preference
                        activeLayerIndex = i
                    End If
                
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
