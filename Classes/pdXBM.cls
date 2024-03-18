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
'Copyright 2024-2024 by Tanner Helland
'Created: 15/March/24
'Last updated: 15/March/24
'Last update: initial build
'
'The XBM file format is *ancient*.  PhotoDemon originally just handed XBM support off to the 3rd-party
' FreeImage library, but FreeImage crashes on all XBM files from well-known software like GIMP.
' So in 2024, I wrote my own little XBM parser.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps additional diagnostic information
' to PD's primary debug log.
Private Const XBM_DEBUG_VERBOSE As Boolean = False

'Image width/height, in pixels, after a file has been validated.
' (Calling "IsFileXBM" is enough to populate these.)
Private m_Width As Long, m_Height As Long

'Byte-by-byte access is provided, as always, by a pdStream instance.
Private m_Stream As pdStream

'The last filename loaded.  The contents of the filename will be stored in m_FileAsString, below.
Private m_OpenFilename As String

'Once the file is validated, its raw text is dumped into this string.
Private m_FileAsString As String

'XBM files are just raw C code, so they don't validate normally.  Instead, I treat the file as a string
' and attempt to find key consts (declaring reasonable image dimensions).
Friend Function IsFileXBM(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True) As Boolean
        
    Const FUNC_NAME As String = "IsFileXBM"
        
    IsFileXBM = False
    On Error GoTo BadXBMFile
    
    m_OpenFilename = vbNullString
    m_Width = 0: m_Height = 0
    
    Dim potentialMatch As Boolean
    potentialMatch = Files.FileExists(srcFilename)
    
    'Check extension too, as requested.
    If (potentialMatch And requireValidFileExtension) Then
        potentialMatch = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "xbm", True)
    End If
    
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
            
            'Debug.Print "|" & containingLine & "|"
            
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
                    
                'This line is not a height declare; keep searching for one...
                Else
                
                End If
                
            End If
            
            'Keep looking for more #define statements, starting from the trailing position of the last-found one
            posText = InStr(posText + 7, aBitOfFile, DEF_STATEMENT, vbTextCompare)
            
        Loop
        
        Debug.Print widthName, heightName, m_Width, m_Height
        
        'Check for success/failure by comparing width/height variable names and discovered width/height
        If Strings.StringsNotEqual(widthName, heightName) Or (m_Width = 0) Or (m_Height = 0) Then
            InternalError FUNC_NAME, "mismatched or invalid width/height"
            GoTo BadXBMFile
        End If
        
        'Debug.Print m_Width, m_Height
        
    '/Couldn't start stream; file may be locked or inaccessible
    Else
        GoTo BadXBMFile
    End If
    
    'TODO: validate existence of size consts
    potentialMatch = False
    IsFileXBM = potentialMatch
    
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
    
    'Validate the file
    If Me.IsFileXBM(srcFile, False) Then
        
        'If validation passed, we already have the (probably correct?) width and height stored in m_width and m_height.
        If (m_Width <= 0) Or (m_Height <= 0) Then
            InternalError FUNC_NAME, "bad dimensions: " & m_Width & "x" & m_Height
            Exit Function
        End If
        
        If XBM_DEBUG_VERBOSE Then
            PDDebug.LogAction "HGT dimensions: " & m_Width & "x" & m_Height
            PDDebug.LogAction "Starting parse..."
        End If
        
        'Open a stream on the target file
        Dim cStream As pdStream
        Set cStream = New pdStream
        
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
            
            'Prep image buffer; we'll dump intensity values straight into it.
            Set dstDIB = New pdDIB
            If dstDIB.CreateBlank(m_Width, m_Height, 32, vbWhite, 255) Then
                
                LoadXBM_FromFile = True
                
                'The returned data is always premultiplied
                If LoadXBM_FromFile Then dstDIB.SetInitialAlphaPremultiplicationState True
                
                'Regardless of outcome, free the underlying stream
                Set cStream = Nothing
                
            Else
                InternalError FUNC_NAME, "out of memory"
                Set m_Stream = Nothing
                Exit Function
            End If
            
        Else
            InternalError FUNC_NAME, "bad stream"
            Exit Function
        End If
        
    Else
        'File is not XBM; silently ignore it
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
