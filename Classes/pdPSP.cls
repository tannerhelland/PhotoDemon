VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PSP (PaintShop Pro) Container and Parser
'Copyright 2020-2020 by Tanner Helland
'Created: 30/December/20
'Last updated: 30/December/20
'Last update: start work on initial build
'
'This class (and its associated pdPSP- child classes) handle JASC/Corel Paint Shop Pro image parsing.
' All code has been custom-built for PhotoDemon, with a special emphasis on parsing performance.
'
'As of December 2020, this class is under heavy construction.  Many features are not yet supported.
'
'As always, import support is being tackled first.  Export support may be added later.
'
'As with all 3rd-party PSP engines, Paint Shop Pro has many features that don't have direct analogs
' in PhotoDemon.  Such features are still extracted by this class, but they will not "appear" in the
' final loaded image.  My ongoing goal is to expand support in this class as various PSP features are
' implemented in PD itself.
'
'Finally, there is always a possibility that a particular file format may prove too cumbersome to
' fully support. I make no guarantee that this class will ever reach full integration into PhotoDemon.
'
'Unless otherwise noted, all code in this class is my original work.  I've based my work off the
' "official" PSP spec at this URL (link good as of December 2020):
' ftp://ftp.corel.com/pub/documentation/PSP/
'
'Older PSP specs were also useful.  You may be able to find them here (link good as of December 2020);
' look for files with names like "psp8spec.pdf":
' http://www.telegraphics.com.au/svn/pspformat/trunk
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PSP files contain a *lot* of information.  To aid debugging, you can activate "verbose" output; this will
' dump all kinds of diagnostic information to the debug log.  (Note that other PSP-adjacent classes have
' their own version of this constant.)
Private Const PSP_DEBUG_VERBOSE As Boolean = False

'PSP loading is complicated, and a lot of things can go wrong.  Instead of returning binary "success/fail"
' values, we return specific flags; "warnings" may be recoverable and you can still attempt to load the file.
' "Failure" returns are unrecoverable and processing *must* be abandoned.  (As a convenience, you can treat
' the "warning" and "failure" values as flags; specific warning/failure states in each category will share
' the same high flag bit.)
'
'As I get deeper into this class, I may expand this enum to include more detailed states.
Public Enum PD_PSPResult
    psp_Success = &H0
    psp_Warning = &H10
    psp_Failure = &H100
    psp_FileNotPSP = &H1000
End Enum

#If False Then
    Private Const psp_Success = &H0, psp_Warning = &H10, psp_Failure = &H100, psp_FileNotPSP = &H1000
#End If

'Like many other formats, PSP files consist of "blocks".  Blocks have hard-coded 2-byte IDs.
Public Enum PSPBlockID
    PSP_IMAGE_BLOCK = 0             '// General Image Attributes Block (main)
    PSP_CREATOR_BLOCK               '// Creator Data Block (main)
    PSP_COLOR_BLOCK                 '// Color Palette Block (main and sub)
    PSP_LAYER_START_BLOCK           '// Layer Bank Block (main)
    PSP_LAYER_BLOCK                 '// Layer Block (sub)
    PSP_CHANNEL_BLOCK               '// Channel Block (sub)
    PSP_SELECTION_BLOCK             '// Selection Block (main)
    PSP_ALPHA_BANK_BLOCK            '// Alpha Bank Block (main)
    PSP_ALPHA_CHANNEL_BLOCK         '// Alpha Channel Block (sub)
    PSP_COMPOSITE_IMAGE_BLOCK       '// Composite Image Block (sub)
    PSP_EXTENDED_DATA_BLOCK         '// Extended Data Block (main)
    PSP_TUBE_BLOCK                  '// Picture Tube Data Block (main)
    PSP_ADJUSTMENT_EXTENSION_BLOCK  '// Adjustment Layer Block (sub)
    PSP_VECTOR_EXTENSION_BLOCK      '// Vector Layer Block (sub)
    PSP_SHAPE_BLOCK                 '// Vector Shape Block (sub)
    PSP_PAINTSTYLE_BLOCK            '// Paint Style Block (sub)
    PSP_COMPOSITE_IMAGE_BANK_BLOCK  '// Composite Image Bank (main)
    PSP_COMPOSITE_ATTRIBUTES_BLOCK  '// Composite Image Attr. (sub)
    PSP_JPEG_BLOCK                  '// JPEG Image Block (sub)
    PSP_LINESTYLE_BLOCK             '// Line Style Block (sub)
    PSP_TABLE_BANK_BLOCK            '// Table Bank Block (main)
    PSP_TABLE_BLOCK                 '// Table Block (sub)
    PSP_PAPER_BLOCK                 '// Vector Table Paper Block (sub)
    PSP_PATTERN_BLOCK               '// Vector Table Pattern Block (sub)
    PSP_GRADIENT_BLOCK              '// Vector Table Gradient Block (not used)
    PSP_GROUP_EXTENSION_BLOCK       '// Group Layer Block (sub)
    PSP_MASK_EXTENSION_BLOCK        '// Mask Layer Block (sub)
    PSP_BRUSH_BLOCK                 '// Brush Data Block (main)
    PSP_ART_MEDIA_BLOCK             '// Art Media Layer Block (main)
    PSP_ART_MEDIA_MAP_BLOCK         '// Art Media Layer map data Block (main)
    PSP_ART_MEDIA_TILE_BLOCK        '// Art Media Layer map tile Block (main)
    PSP_ART_MEDIA_TEXTURE_BLOCK     '// AM Layer map texture Block (main)
    PSP_COLORPROFILE_BLOCK          '// ICC Color profile block
End Enum

#If False Then
    Private Const PSP_IMAGE_BLOCK = 0, PSP_CREATOR_BLOCK = 1, PSP_COLOR_BLOCK = 2, PSP_LAYER_START_BLOCK = 3, PSP_LAYER_BLOCK = 4, PSP_CHANNEL_BLOCK = 5, PSP_SELECTION_BLOCK = 6, PSP_ALPHA_BANK_BLOCK = 7, PSP_ALPHA_CHANNEL_BLOCK = 8, PSP_COMPOSITE_IMAGE_BLOCK = 9
    Private Const PSP_EXTENDED_DATA_BLOCK = 10, PSP_TUBE_BLOCK = 11, PSP_ADJUSTMENT_EXTENSION_BLOCK = 12, PSP_VECTOR_EXTENSION_BLOCK = 13, PSP_SHAPE_BLOCK = 14, PSP_PAINTSTYLE_BLOCK = 15, PSP_COMPOSITE_IMAGE_BANK_BLOCK = 16, PSP_COMPOSITE_ATTRIBUTES_BLOCK = 17, PSP_JPEG_BLOCK = 18, PSP_LINESTYLE_BLOCK = 19
    Private Const PSP_TABLE_BANK_BLOCK = 20, PSP_TABLE_BLOCK = 21, PSP_PAPER_BLOCK = 22, PSP_PATTERN_BLOCK = 23, PSP_GRADIENT_BLOCK = 24, PSP_GROUP_EXTENSION_BLOCK = 25, PSP_MASK_EXTENSION_BLOCK = 26, PSP_BRUSH_BLOCK = 27, PSP_ART_MEDIA_BLOCK = 28, PSP_ART_MEDIA_MAP_BLOCK = 29
    Private Const PSP_ART_MEDIA_TILE_BLOCK = 30, PSP_ART_MEDIA_TEXTURE_BLOCK = 31, PSP_COLORPROFILE_BLOCK = 32
#End If

'PSP images are ultimately comprised of a collection of blocks.  These are analagous to
' something like "chunks" in a PNG file (although PSP "blocks" can contain sub-blocks and
' chunks within themselves, which is confusing to say the least).  As an early parse step,
' we quickly scan all blocks and cache their IDs and offsets.  This is helpful for both
' validation purposes, and general parsing purposes, because it lets us operate at a
' more helpful "block" level.
Private m_numOfBlocks As Long
Private m_Blocks() As pdPSPBlock

'Byte-by-byte access is provided, as always, by a pdStream instance
Private m_Stream As pdStream

'At present, we require the caller to pass an identical source file path to every load function.
' (This is a cheap and easy way to ensure no funny business.)  If the PSP is loaded directly from memory,
' we flag this with a special name.
Private m_SourceFilename As String
Private Const PSP_LOADED_FROM_MEMORY As String = "LoadFromPtr*"
Private m_SourcePtr As Long, m_SourcePtrLen As Long

'If warnings are encountered during processing, I push their messages onto a string stack.  (I may
' decide to report these to the user... but haven't decided yet.  Either way, it's very helpful
' while debugging; see associated Warnings_XYZ functions for details.)
Private m_Warnings As pdStringStack

'Perform basic validation on a potential PSP file.  For strict file extension matching (e.g. only
' .psp or .pspimage extension allowed), set checkExtension to TRUE.
Friend Function IsFilePSP(ByRef srcFile As String) As Boolean
    IsFilePSP = (Step1_ValidateHeader(srcFile) < psp_Failure)
End Function

'Simplified wrapper to load a PSD automatically.
Friend Function LoadPSP(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal loadFromPtr As Long = 0, Optional ByVal loadFromPtrLen As Long = 0) As PD_PSPResult
    
    'Reset some internal parameters to ensure subsequent reads are accurate.
    ' (This is critical if multiple PSPs are read back-to-back.)
    Me.Reset
    
    'We support PSD loading from both file and memory; if the passed loadFromPtr value is non-zero,
    ' treat it as a pointer and we'll wrap our stream around it instead.
    If (loadFromPtr <> 0) And (loadFromPtrLen <> 0) Then
        srcFile = PSP_LOADED_FROM_MEMORY
        m_SourcePtr = loadFromPtr
        m_SourcePtrLen = loadFromPtrLen
    Else
        m_SourcePtr = 0
        m_SourcePtrLen = 0
    End If

    'Try to validate the source file
    Dim keepLoading As PD_PSPResult
    keepLoading = Step1_ValidateHeader(srcFile)
    
    'If this identifies as a PSP file, build and validate the block directory
    If (keepLoading < psp_Failure) Then keepLoading = Step2_BuildBlockDirectory(srcFile)
    
    'If the block directory looks good, parse the general image header.
    
    
    'For now, PSP imports are deliberately marked as failures.  We have a lot more work to do
    ' before we can actually load a useable image!
    keepLoading = psp_Failure
    
    LoadPSP = keepLoading
    
    'Even if the PSP loaded, one (or more) surprises may have affected our ability to render
    ' the file correctly.  Regardless of this class's debug verbosity setting, warnings always
    ' get dumped out to the debug log.  This is especially useful when new versions of PaintShop Pro
    ' release, as they often bring new parsing surprises/nightmares.
    If (Me.Warnings_GetCount() > 0) Then
        PDDebug.LogAction "PhotoDemon's PSP parser generated one or more warnings.  Here is a full list:"
        Me.Warnings_DumpToDebugger
    End If
    
End Function

Friend Sub Reset()
    Set m_Warnings = New pdStringStack
End Sub

Private Function Step1_ValidateHeader(ByRef srcFile As String) As PD_PSPResult
    
    On Error GoTo InternalVBError
    
    'If the passed path is zero, assume the caller is loading the PSP from memory.
    If (LenB(srcFile) = 0) Then m_SourceFilename = PSP_LOADED_FROM_MEMORY Else m_SourceFilename = srcFile
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'Open a stream over the source data.
    Set m_Stream = New pdStream
    
    If (m_SourceFilename = PSP_LOADED_FROM_MEMORY) Then
        If Not m_Stream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_SourcePtrLen, m_SourcePtr) Then
            m_Warnings.AddString "Couldn't start in-memory stream against passed pointer: " & m_SourcePtr
            okToProceed = psp_Failure
        End If
    Else
        If Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, m_SourceFilename, , , OptimizeSequentialAccess) Then
            m_Warnings.AddString "Can't read file; it may be locked or in an inaccessible location."
            okToProceed = psp_Failure
        End If
    End If
    
    'The stream is open.  Validate both the PSP's ASCII identifier and its version.
    ' (We will attempt to parse on a version failure, but I have not tested old PSP file versions
    ' as aggressively so no promises are made on correctness!)
    Dim pspValidated As Boolean
    If (okToProceed < psd_Failure) Then
        
        'PSP files use a lovely 32-byte identifier that completely eliminates the possibility
        ' of false-positives.  Yay!
        Dim pspID() As Byte, checkBytes() As Byte
        
        '(Note: this hex string translates to “Paint Shop Pro Image File\n\x1a”, padded with zeroes
        ' to 32 bytes).
        If Strings.BytesFromHex(pspID, "5061696E742053686F702050726F20496D6167652046696C650A1A00000000") Then
            If m_Stream.AreBytesAvailable(32) Then
                If (m_Stream.ReadBytes(checkBytes, 32, True) = 32) Then
                    pspValidated = VBHacks.MemCmp(VarPtr(pspID(0)), VarPtr(checkBytes(0)), 32)
                End If
            End If
        End If
        
        If pspValidated Then
            okToProceed = psp_Success
        Else
            okToProceed = psp_FileNotPSP
        End If
        
        'Validate version number only if the original checksum passed
        If (okToProceed < psp_Failure) Then
        
            Dim pspVersionMajor As Integer, pspVersionMinor As Integer
            pspVersionMajor = m_Stream.ReadInt()
            pspVersionMinor = m_Stream.ReadInt()
            
            'Per v8 of the spec, PSP versions are as follows:
            'Current version of the spec (v8) is 6.0
            ' 3.0 = Paint Shop Pro 5
            ' 4.0 = Paint Shop Pro 6
            ' 5.0 = Paint Shop Pro 7
            ' 6.0 = Paint Shop Pro 8
            ' 7.0 = Paint Shop Pro 9
            
            'PhotoDemon will attempt to import *any* version, but versions outside the above
            ' list are not guaranteed to load correctly.
            If (pspVersionMajor < 3) Or (pspVersionMajor > 7) Then m_Warnings.AddString "Unexpected PSP major version encountered: " & pspVersionMinor
            If (pspVersionMinor <> 0) Then m_Warnings.AddString "Unexpected PSP minor version encountered: " & pspVersionMinor
            
        Else
            m_Warnings.AddString "PSP header failed basic validation.  (This is not a PSP file.)"
            okToProceed = psp_FileNotPSP
        End If
        
    End If
    
    'Validation complete.  If the file validated successfully, this function guarantees that m_Stream
    ' points at the first byte PAST the valid PSP file header.
    Step1_ValidateHeader = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError "Step1_ValidateHeader", "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 1, #" & Err.Number & ": " & Err.Description
    Step1_ValidateHeader = psp_Failure

End Function

'Do *NOT* call import steps out-of-order!
Private Function Step2_BuildBlockDirectory(ByRef srcFile As String) As PD_PSPResult
    
    On Error GoTo InternalVBError
    
    Step2_BuildBlockDirectory = psp_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "Step2_BuildBlockDirectory", "filename has changed since original validation!"
        Step2_BuildBlockDirectory = psp_Failure
        Exit Function
    End If
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'PSP files consist of a series of "blocks".  Total block count is unknown; the only way
    ' to know how many blocks are in the file is to "walk" the file manually.
    
    'In this step, we build an initial "block" collection.  This lets us perform some
    ' additional validation (like ensuring that critical blocks are in valid order).
    
    'Start by resetting our current block collection, if any
    Const DEFAULT_BLOCK_COUNT As Long = 8
    m_numOfBlocks = 0
    ReDim m_Blocks(0 To DEFAULT_BLOCK_COUNT - 1) As pdPSPBlock
    
    Dim curBlockID As PSPBlockID, curBlockLength As Long
    
    'After validation, we will be pointing at the first block in our collection.
    
    'All blocks use a uniform 32-bit "flag" ID, followed by a block ID and length
    Const PSP_BLOCK_START As Long = &H4B427E
    
    Do While (m_Stream.ReadLong() = PSP_BLOCK_START)
        
        'Ensure there's room in our block collection
        If (m_numOfBlocks > UBound(m_Blocks)) Then ReDim Preserve m_Blocks(0 To m_numOfBlocks * 2 - 1) As pdPSPBlock
        
        'Retrieve 2-byte ID
        curBlockID = m_Stream.ReadIntUnsigned()
        curBlockLength = m_Stream.ReadLong()
        
        'Add this block to the collection
        Set m_Blocks(m_numOfBlocks) = New pdPSPBlock
        m_Blocks(m_numOfBlocks).InitializeBlock m_Stream.GetPosition(), curBlockID, curBlockLength
        m_numOfBlocks = m_numOfBlocks + 1
        
        'Move the pointer ahead by the specified block length; we should be pointing at either a
        ' new block, or the end of the file!
        m_Stream.SetPosition curBlockLength, FILE_CURRENT
        If (m_Stream.GetPosition() >= m_Stream.GetStreamSize()) Then Exit Do
        
    Loop
    
    'With all blocks flagged, we can now do a bit of minimal validation on blocks and block order.
    ' (The PSP spec requires certain blocks appear in certain places and/or orders.)
    
    'From the spec:
    ' With regard to the order of the ten main blocks, there are two hard and fast rules.
    '  - First, the General Image Attributes block must come immediately after the File Header.
    '  - Second, the Composite Image Bank Block, if present, must come before the Layer Bank Block.
    '  - The order of other blocks is not mandated.
    
    'Start by ensuring the first block is the General Image Attributes block
    If (m_numOfBlocks > 0) Then
        If (Not m_Blocks(0) Is Nothing) Then
            If (m_Blocks(0).GetBlockID <> PSP_IMAGE_BLOCK) Then
                okToProceed = psp_Failure
                InternalError "Step2_BuildBlockDirectory", "first block is not General Image Attributes"
            End If
        Else
            okToProceed = psp_Failure
            InternalError "Step2_BuildBlockDirectory", "first block is corrupt"
        End If
    Else
        okToProceed = psp_Failure
        InternalError "Step2_BuildBlockDirectory", "no valid blocks found"
    End If
    
    'If the first block passed, look for the composite image bank block and the layer bank block.
    ' We need to ensure the former appears before the latter.
    Dim posCompositeBank As Long, posLayerBank As Long
    If (okToProceed < psp_Failure) Then
        
        'Set the composite and layer bank indices to invalid values
        posLayerBank = -2
        posCompositeBank = -1
        
        Dim i As Long
        For i = 0 To m_numOfBlocks - 1
            If (Not m_Blocks(i) Is Nothing) Then
                If (m_Blocks(i).GetBlockID = PSP_COMPOSITE_IMAGE_BANK_BLOCK) Then posCompositeBank = i
                If (m_Blocks(i).GetBlockID = PSP_LAYER_START_BLOCK) Then posLayerBank = i
            End If
        Next i
        
        If (posLayerBank <= posCompositeBank) Then
            okToProceed = psp_Failure
            InternalError "Step2_BuildBlockDirectory", "critical blocks out of order"
        End If
        
    End If
    
    'Finally, PSP images could theoretically be "empty" (e.g. containing just a header, but no
    ' layer or pixel data), but we don't want to mess with this possibility in PD.  Throw out
    ' such files.
    If (okToProceed < psp_Failure) Then
        If (m_numOfBlocks < 2) Then
            okToProceed = psp_Failure
            InternalError "Step2_BuildBlockDirectory", "too few blocks"
        End If
    End If
    
    'Block directory validation and assembly is complete!
    Step2_BuildBlockDirectory = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError "Step2_BuildBlockDirectory", "internal VB error #" & Err.Number & ": " & Err.Description
    If (Not m_Stream Is Nothing) Then If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 2, #" & Err.Number & ": " & Err.Description
    Step2_BuildBlockDirectory = psp_Failure

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

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdPSP." & funcName & "() reported an error on file """ & m_SourceFilename & """: " & errDescription
    Else
        Debug.Print "pdPSP." & funcName & "() reported an error on file """ & m_SourceFilename & """: " & errDescription
    End If
End Sub

'Want data on warnings?  Use these helper functions.
Friend Function Warnings_GetCount() As Long
    Warnings_GetCount = m_Warnings.GetNumOfStrings()
End Function

Friend Sub Warnings_CopyList(ByRef dstStack As pdStringStack)
    Set dstStack = m_Warnings
End Sub

Friend Sub Warnings_DumpToDebugger()
    If (m_Warnings.GetNumOfStrings() > 0) Then
        Dim i As Long
        For i = 0 To m_Warnings.GetNumOfStrings() - 1
            PDDebug.LogAction "(" & CStr(i + 1) & ") WARNING: pdPSP reported: " & m_Warnings.GetString(i)
        Next i
    End If
End Sub

