VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdXBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon X Bitmap (XBM) Decoder
'Copyright 2024-2026 by Tanner Helland
'Created: 15/March/24
'Last updated: 18/March/24
'Last update: wrap up initial build
'
'The XBM file format is *ancient*.  PhotoDemon originally just handed XBM support off to the 3rd-party
' FreeImage library, but FreeImage crashes on pretty much all XBM files (including modern exported ones
' from well-known software like GIMP).
'
'So in 2024, I wrote my own XBM decoder.  It definitely falls on the "quick-and-dirty" side of things,
' but it's chock full of error-handling (and safety asserts) and I've tested it rigorously on a variety
' of image dimensions (tiny to enormous), including all possible scanline padding and GIMP encoding cases.
' I've also tested some old XBM files from various online archives without problems, and I've made sure
' to handle some weird edge-cases like files that start with a comment (which GIMP explicitly supports
' when writing, though it does warn the user that this may "break some software").
'
'The end result is a fast and lightweight parser that covers many XBM cases that FreeImage does not,
' and without any more risk of hard-crashes just from attempting to load an XBM image!
'
'Encoding support is not planned at present, but this may be revisited (as always) if users complain.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps additional diagnostic information
' to PD's primary debug log.
Private Const XBM_DEBUG_VERBOSE As Boolean = False

'Embedded name of the image (this text precedes width/height/bitmap declares)
Private m_embeddedName As String

'Image width/height, in pixels, after a file has been validated.
' (Calling "IsFileXBM" is enough to populate these.)
Private m_Width As Long, m_Height As Long

'Byte-by-byte access is provided, as always, by a pdStream instance.
Private m_Stream As pdStream

'The last filename loaded.  We use this to skip validation during loading, if the caller already called
' IsFileXBM on the current filename they want loaded.
Private m_OpenFilename As String

'XBM files are just raw C code, so they don't validate normally.  Instead, I treat the file as a string
' and attempt to find key consts (declaring reasonable image dimensions).
Friend Function IsFileXBM(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True) As Boolean
        
    Const FUNC_NAME As String = "IsFileXBM"
        
    IsFileXBM = False
    On Error GoTo BadXBMFile
    
    m_OpenFilename = vbNullString
    m_Width = 0: m_Height = 0: m_embeddedName = vbNullString
    
    Dim potentialMatch As Boolean
    potentialMatch = Files.FileExists(srcFilename)
    
    'Check extension too, as requested.
    If (potentialMatch And requireValidFileExtension) Then
        potentialMatch = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "xbm", True)
    End If
    
    'If any of the failsafe checks failed, exit immediately
    If (Not potentialMatch) Then Exit Function
    
    'XBM files are just C code (as you'd use to embed an icon manually).
    ' Here is an example from Wikipedia (https://en.wikipedia.org/wiki/X_BitMap):
    '#define test_width 16
    '#define test_height 7
    'static unsigned char test_bits[] = {
    '0x13, 0x00, 0x15, 0x00, 0x93, 0xcd, 0x55, 0xa5, 0x93, 0xc5, 0x00, 0x80,
    '0x00, 0x60 };
    
    'Here is another example, an actual export from GIMP but truncated for brevity.
    ' Note the preceding comment (which should be treated as valid on an importer):
    ' /* Created with GIMP */
    ' #define cloudy__GIMP_comment__width 64
    ' #define cloudy__GIMP_comment__height 64
    ' static unsigned char cloudy__GIMP_comment__bits[] = {
    '  0x2a, 0x2a, 0x2a, 0x2a, 0x2a, 0x2a, 0x2a, 0x2a, 0x55, 0x55, 0x55, 0x55, ...
    
    'What we're looking for here are the width/height declare lines.
    
    'Grab the first 1024 chars of the file (if that many even exist) and look for the necessary declares.
    If (m_Stream Is Nothing) Then Set m_Stream = New pdStream
    If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename, optimizeAccess:=OptimizeSequentialAccess) Then
        
        m_OpenFilename = srcFilename
        
        'Load the first few kb of file contents into a string
        Const ARBITRARY_TEST_SIZE As Long = 2048
        Dim aBitOfFile As String
        aBitOfFile = m_Stream.ReadString_UnknownEncoding(ARBITRARY_TEST_SIZE, True)
        
        'We now want to look for two #define statements: one for width, and another for height.
        ' (We are not strict about which comes first, or the presence of comments strewn throughout the code.)
        Const DEF_STATEMENT As String = "#define", SPACE_CHAR As String = " "
        
        Dim posText As Long, containingLine As String, tokenizedLine() As String, tokenCount As Long
        posText = InStr(1, aBitOfFile, DEF_STATEMENT, vbTextCompare)
        
        'Width/height declares must use the same variable name to be considered valid
        Dim posTextInner As Long, widthName As String, heightName As String
        
        'Continue searching for #define statements until both width *and* height are found
        Do While (posText > 0)
        
            'Grab the whole line of text containing the define statement.
            containingLine = Trim$(Strings.GetLineContainingPosition(aBitOfFile, posText))
            
            'Search for the word "width" or "height" in the text
            posTextInner = InStr(1, containingLine, "_width", vbBinaryCompare)
            If (posTextInner > 0) Then
                
                'This line contains a width indicator.  We want two pieces of information from it:
                ' the preceding variable name, and the number following it.
                
                'Start by tokenizing the line into individual components
                tokenizedLine = Split(containingLine, SPACE_CHAR)
                tokenCount = UBound(tokenizedLine) + 1
                
                'Must have at least three tokens: "#define", "name_width", "<width-as-integer>"
                If (tokenCount < 3) Then GoTo BadXBMFile
                
                'Grab the actual width (must be the last entry on this line - comments are invalid)
                If IsNumeric(tokenizedLine(tokenCount - 1)) Then
                    m_Width = CLng(tokenizedLine(tokenCount - 1))
                Else
                    GoTo BadXBMFile
                End If
                
                'Grab the name prefix used for width (must match height and pixel bits as well);
                ' length is calculated as <token-length> - <length-of-"_width">
                widthName = Left$(tokenizedLine(tokenCount - 2), Len(tokenizedLine(tokenCount - 2)) - 6)
                
                'If we already have height, we can exit the loop
                If (m_Height <> 0) Then Exit Do
                
            Else
                
                posTextInner = InStr(1, containingLine, "_height", vbBinaryCompare)
                If (posTextInner > 0) Then
                    
                    'This line contains a height indicator.  We want two pieces of information from it:
                    ' the preceding variable name, and the number following it.
                    
                    '(Because steps here are identical to width, above, comments are omitted.)
                    tokenizedLine = Split(containingLine, SPACE_CHAR)
                    tokenCount = UBound(tokenizedLine) + 1
                    If (tokenCount < 3) Then GoTo BadXBMFile
                    
                    If IsNumeric(tokenizedLine(tokenCount - 1)) Then
                        m_Height = CLng(tokenizedLine(tokenCount - 1))
                    Else
                        GoTo BadXBMFile
                    End If
                    
                    heightName = Left$(tokenizedLine(tokenCount - 2), Len(tokenizedLine(tokenCount - 2)) - 7)
                    If (m_Width <> 0) Then Exit Do
                    
                '/This line is not a height declare; keep searching for one (don't need to branch).
                End If
                
            End If
            
            'Keep looking for more #define statements, starting from the trailing position of the last-found one
            posText = InStr(posText + 7, aBitOfFile, DEF_STATEMENT, vbTextCompare)
            
        Loop
        
        'Check for success/failure by comparing width/height variable names and discovered width/height,
        ' and enforce a signed short upper limit on embedded sizes (intensely overkill, as XBM was only
        ' really used for icons!)
        If Strings.StringsNotEqual(widthName, heightName) Or (m_Width = 0) Or (m_Height = 0) Or (m_Width > 32767) Or (m_Height > 32767) Then
            InternalError FUNC_NAME, "mismatched or invalid width/height"
            GoTo BadXBMFile
        End If
        
    '/Couldn't start stream; file may be locked or inaccessible
    Else
        GoTo BadXBMFile
    End If
    
    'Ensure we have valid, non-zero width/height and matching width/height variable names
    If (widthName = heightName) And (m_Width > 0) And (m_Height > 0) Then
        m_embeddedName = widthName
        IsFileXBM = True
    End If
    
    'Close the file stream before exiting
    If (Not m_Stream Is Nothing) Then m_Stream.StopStream True
    
    Exit Function
    
'On any parse error, this function jumps to this branch and simply closes the underlying file, then exits
BadXBMFile:
    
    Set m_Stream = Nothing
    InternalError FUNC_NAME, "critical parse failure"
    IsFileXBM = False
    
End Function

'Validate and load a candidate XBM file
Friend Function LoadXBM_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "LoadXBM_FromFile"
    LoadXBM_FromFile = False
    
    On Error GoTo BadXBMFile
    
    'Validate the file
    If Me.IsFileXBM(srcFile, False) Then
        
        'If validation passed, the width and height (if any) will be stored in m_width and m_height.
        If (m_Width <= 0) Or (m_Height <= 0) Then
            InternalError FUNC_NAME, "bad dimensions: " & m_Width & "x" & m_Height
            Exit Function
        End If
        
        'We also need to know the embedded name of the image file
        If (LenB(m_embeddedName) = 0) Then
            InternalError FUNC_NAME, "bad embedded name"
            Exit Function
        End If
        
        If XBM_DEBUG_VERBOSE Then
            PDDebug.LogAction "XBM dimensions: " & m_Width & "x" & m_Height
            PDDebug.LogAction "Starting to parse pixel data..."
        End If
        
        'Open a stream on the target file
        If (m_Stream Is Nothing) Then Set m_Stream = New pdStream
        If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
            
            'Pull the entire source file into memory
            Dim fileContents As String
            fileContents = m_Stream.ReadString_UTF8(m_Stream.GetStreamSize, False)
            
            'Close the stream (it's no longer needed)
            m_Stream.StopStream True
            
            'Because we know the embedded image name, we can search directly for the pixel bits header
            Dim posBits As Long
            posBits = InStr(1, fileContents, m_embeddedName & "_bits", vbBinaryCompare)
            If (posBits <= 0) Then
                InternalError FUNC_NAME, "no bits found"
                Exit Function
            End If
            
            'Still here?  That means we found the bits entry.
            
            'After the bits entry, we want to find the starting bracket for the array that follows.
            posBits = InStr(posBits, fileContents, "{", vbBinaryCompare)
            If (posBits <= 0) Then
                InternalError FUNC_NAME, "bad bits array"
                Exit Function
            End If
            
            'Find the trailing bracket that marks the end of the pixel data
            posBits = posBits + 1
            
            Dim posEnd As Long
            posEnd = InStr(posBits, fileContents, "}", vbBinaryCompare)
            If (posBits <= 0) Then
                InternalError FUNC_NAME, "bad bits array"
                Exit Function
            End If
            
            'We now have positional delimiters that mark the start and end of the pixel array.
            ' Alas, in VB6 there is no easy way to convert the contents of said array into a stream,
            ' so we're gonna need to iterate it manually and convert as-we-go.
            
            '(Could we do a simple Split and iterate that?  Yes.  Would that be easier?  Also yes.
            ' But there's no obvious upper limit on XBM dimensions, and after testing some large images
            ' exported from GIMP, it's clear that the pixel array can easily become large enough to
            ' "break" Split (by exhausting BSTR space).  So rather than risk that, let's just iterate
            ' the string manually and convert values as we go.)
            
            'Advance the initial char pointer to the first valid C-style hex numerical char
            Do While (posBits < posEnd) And (Not IsValidCHexChar(fileContents, posBits))
                posBits = posBits + 1
            Loop
            
            'Failsafe check for bad array
            If (posBits >= posEnd) Then
                InternalError FUNC_NAME, "bad bits array"
                Exit Function
            End If
            
            'As we calculate them, pixel values will get pushed into a new stream.
            Dim dstStream As pdStream
            Set dstStream = New pdStream
            dstStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
            
            'We also need to know the size of each scanline of bits, which can vary based on whether
            ' bytes or shorts are used to encode data in the file.
            Dim sizeOfEachHex As Long: sizeOfEachHex = 0
            
            'Start iterating the string, pulling out pixel values as we go
            Do While (posBits < posEnd)
                
                'Find the next non-number, non-hex character
                Dim endOfToken As Long
                endOfToken = posBits + 1
                Do While (endOfToken < posEnd) And IsValidCHexChar(fileContents, endOfToken)
                    endOfToken = endOfToken + 1
                Loop
                
                'Ensure the entry is at least 4 chars long (e.g 0x00) and double-check the initial 0x just to be safe
                Const C_HEX_PREFIX As String = "0x", VB_HEX_PREFIX As String = "&H"
                
                If ((endOfToken - posBits) >= 4) And (LCase$(Mid$(fileContents, posBits, 2)) = C_HEX_PREFIX) Then
                    
                    'This looks like a valid hex entry!  Pull out the stuff VB can understand and push it
                    ' onto our running pixel stream.
                    Dim justTheNumberBits As String
                    justTheNumberBits = VB_HEX_PREFIX & Mid$(fileContents, posBits + 2, endOfToken - posBits - 2)
                    
                    'XBM bits appear to come in two sizes: bytes and shorts.  It's possible that longer hex
                    ' combinations could be used, but I've never seen that "in the wild".
                    Select Case Len(justTheNumberBits) - 2
                        Case 2
                            dstStream.WriteByte CByte(justTheNumberBits)
                        Case 4
                            dstStream.WriteInt CInt(justTheNumberBits)
                        Case 8
                            dstStream.WriteLong CLng(justTheNumberBits)
                        Case Else
                            InternalError FUNC_NAME, "invalid hex size: " & Len(justTheNumberBits) & ", " & justTheNumberBits
                            GoTo BadXBMFile
                    End Select
                    
                    'Note the size of each hex entry (can be chars or shorts; PD also supports ints,
                    ' but I've never seen such a file "in the wild")
                    If (sizeOfEachHex = 0) Then sizeOfEachHex = (Len(justTheNumberBits) - 2) \ 2
                    
                    'Advance the token index beyond the end of this entry
                    posBits = endOfToken + 1
                    Do While (posBits < posEnd) And (Not IsValidCHexChar(fileContents, posBits))
                        posBits = posBits + 1
                    Loop
                    
                Else
                    InternalError FUNC_NAME, "unexpected parse"
                    GoTo BadXBMFile
                End If
                
            Loop
            
            'dstStream now contains the full contents of the XBM file, stored as a normal byte array.
            ' We need to iterate that array, and pull out individual bits and (set b/w accordingly).
            
            'Perform a failsafe check - we must have enough pixel data to supply the full image.
            If ((m_Width * m_Height) > (dstStream.GetStreamSize() * 8)) Then
                InternalError FUNC_NAME, "file ends prematurely"
                GoTo BadXBMFile
            End If
            
            'Prep image buffer; we'll dump intensity values straight into it.
            Set dstDIB = New pdDIB
            If dstDIB.CreateBlank(m_Width, m_Height, 32, vbWhite, 255) Then
                
                'There's not a performance-friendly way to mask flags in VB, so let's just use a byte array for clarity
                Dim bitFlags(0 To 7) As Byte
                bitFlags(0) = 1
                bitFlags(1) = 2
                bitFlags(2) = 4
                bitFlags(3) = 8
                bitFlags(4) = 16
                bitFlags(5) = 32
                bitFlags(6) = 64
                bitFlags(7) = 128
                
                'To avoid having to pull pixels out one-by-one (which is slow), grab the whole source data chunk
                ' as a local byte array.
                Dim srcPixels() As Byte, idxSrcByte As Long
                dstStream.SetPosition 0, FILE_BEGIN
                dstStream.ReadBytes srcPixels, -1, True
                dstStream.StopStream True
                
                'We'll need to track an index into the source data; it'll be updated "as we go"
                idxSrcByte = 0
                
                'Wrap an array around the destination DIB.  (This is unsafe, and must be manually
                ' freed before this function xits.)
                Dim imgPixels() As Long, imgSA As SafeArray1D
                dstDIB.WrapLongArrayAroundDIB_1D imgPixels, imgSA
                
                'We know how many pixels we should be addressing (based on the underlying width/height)
                Dim totalNumPixels As Long, numPixelsScanline As Long, curByte As Byte
                totalNumPixels = m_Width * m_Height
                
                'Dummy bits at the end of each scanline can be ignored.  The formula works like this:
                ' maxPixelsScanline = Int((m_Width + 7) \ 8) * 8
                ' ...where "8" is the number of bits *per hex entry in the file* (e.g. 8 = 0x00, 16 = 0x0000, etc)
                
                'Convert embedded hex size from bytes to bits
                sizeOfEachHex = sizeOfEachHex * 8
                
                Dim maxPixelsScanline As Long
                maxPixelsScanline = Int((m_Width + (sizeOfEachHex - 1)) \ sizeOfEachHex) * sizeOfEachHex
                
                'Start iterating bytes and converting bits to monochrome colors!
                Dim x As Long
                Do While (x < totalNumPixels)
                    
                    'We don't need to validate these reads, as we already asserted source size in a previous step
                    curByte = srcPixels(idxSrcByte)
                    
                    'Parse each bit in turn
                    Dim i As Long
                    For i = 0 To 7
                        
                        'Ignore empty bytes at the end of the image
                        If (x < totalNumPixels) Then
                            
                            'Ignore empty bytes at the end of each scanline
                            If (numPixelsScanline < m_Width) Then
                                
                                'Draw black bits into the image
                                If (bitFlags(i) = (curByte And bitFlags(i))) Then
                                    imgPixels(x) = &HFF000000
                                '/No else required, because the base image is already opaque white
                                End If
                                
                                'Increment pixel pointer
                                x = x + 1
                                
                            End If
                            
                        End If
                        
                        'Keep track of how many pixels we've copied *on this scanline*
                        numPixelsScanline = numPixelsScanline + 1
                        
                    Next i
                    
                    'If we've reached the end of a scanline, reset the scanline pixel counter
                    If (numPixelsScanline >= maxPixelsScanline) Then numPixelsScanline = 0
                    
                    'Advance to the next byte
                    idxSrcByte = idxSrcByte + 1
                    
                Loop
                
                'Release our unsafe array wrapper
                dstDIB.UnwrapLongArrayFromDIB imgPixels
                
                'File was loaded successfully!
                LoadXBM_FromFile = True
                
                'The returned data is always premultiplied
                If LoadXBM_FromFile Then dstDIB.SetInitialAlphaPremultiplicationState True
                
            Else
                InternalError FUNC_NAME, "out of memory"
                Set m_Stream = Nothing
                Exit Function
            End If
            
        Else
            InternalError FUNC_NAME, "bad stream"
            Exit Function
        End If
    
    '/File is not XBM; silently ignore it
    End If
    
    Exit Function
    
BadXBMFile:
    InternalError FUNC_NAME, "abandoned load due to critical error"
    LoadXBM_FromFile = False
    
End Function

'Test a character and return TRUE if the char can appear in a C-style hex number (e.g. 0x1234abcd)
Private Function IsValidCHexChar(ByRef srcString As String, ByVal idxChar As Long) As Boolean
    
    Dim lAsc As Long
    lAsc = AscW(LCase$(Mid$(srcString, idxChar, 1)))
    
    '0-9
    If ((lAsc >= 48) And (lAsc <= 57)) Then
        IsValidCHexChar = True
    
    'a-f
    ElseIf ((lAsc >= 97) And (lAsc <= 102)) Then
        IsValidCHexChar = True
    
    'x (as in 0x0123)
    Else
        IsValidCHexChar = (lAsc = 120)
    End If
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdXBM." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdXBM." & funcName & "() reported an error: " & errDescription
    End If
End Sub

'The underlying stream would auto-free naturally, but I like being tidy
Private Sub Class_Terminate()
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen Then m_Stream.StopStream True
    End If
End Sub
