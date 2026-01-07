Attribute VB_Name = "ExifTool"
'***************************************************************************
'ExifTool Plugin Interface
'Copyright 2013-2026 by Tanner Helland
'Created: 24/May/13
'Last updated: 29/September/23
'Last update: manually write pHYs block when exporting PNGs
'
'Module for handling all ExifTool interfacing.  This module is pointless without the accompanying ExifTool plugin,
' which can be found in the App/PhotoDemon/Plugins subdirectory as "exiftool.exe".  The ExifTool plugin is
' available by default in all versions of PhotoDemon after (and including) 6.0.
'
'ExifTool is a comprehensive image metadata handler written by Phil Harvey.  No DLL or VB-compatible library
' is available, so PhotoDemon relies on the stock Windows ExifTool executable file for all interfacing.
' You can read more about ExifTool at its homepage:
'
'https://exiftool.org/
'
'As of PhotoDemon 6.1.499, all ExifTool interaction is piped across stdin/out.  This includes sending requests
' to ExifTool, retrieving ExifTool results, and checking ExifTool returns for success/failure.  All of PhotoDemon's
' metadata code has been rewritten to take advantage of this new asynchronous implementation.
'
'Prior to that release, I used a sample VB module as a valuable reference (c/o Michael Wandel):
'http://owl.phy.queensu.ca/~phil/exiftool/modExiftool_101.zip
'
'That reference code has long since been replaced with the current custom-built async implementation,
' but I thought it worthwhile to mention it in case you want a (much simpler!) look at how you might interact
' with ExifTool from VB6.
'
'This project is periodically tested against the newest build of ExifTool.  In most cases, you can drop in the
' latest ExifTool Windows EXE release without problems, but occasionally you may need to delete the Data/PluginData
' folder and allow ExifTool to rebuild it from scratch.  (This folder is where PD's copy of ExifTool extracts its
' portable Perl runtime, and version changes can muck things up unpredictably.)
'
'Additional documentation regarding the use of ExifTool can be found in the official ExifTool package,
' downloadable from https://exiftool.org/
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'DEBUGGING ONLY!  Do not enable this constant in production builds, as it does obnoxious things like
' overwrite the clipboard with streaming metadata information.
Public Const EXIFTOOL_DEBUGGING_ENABLED As Long = 0&

'ExifTool needs to unpack a portable copy of the perl runtime; hence the awkward SetEnvironmentVariableW usage
' (this allows unpacking to a specific folder - /Data/PluginData in our case - instead of the system temp folder)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function SetEnvironmentVariableW Lib "kernel32" (ByVal ptrToEnvName As Long, ByVal ptrToEnvValue As Long) As Long

'A great deal of extra code is required for finding ExifTool instances left by previous unsafe shutdowns, and silently terminating them.
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, newState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function LookupPrivilegeValueW Lib "advapi32" (ByVal ptrSystemName As Long, ByVal ptrName As Long, ByRef lpLuid As LUID) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
 
Private Type LUID
    lowPart As Long
    highPart As Long
End Type
 
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type
 
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
 
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH_LEN
End Type

'This type is what PhotoDemon uses internally for storing and displaying metadata.  Its complexity is a testament to the
' nightmare that is metadata management.
Public Type PDMetadataItem
    
    TagGroupAndName As String       'Something like "IFD0:ResolutionUnit".  This is the tag name used by ExifTool
    TagGroup As String              'The first half of FullGroupAndName
    TagGroupFriendly As String      'The PhotoDemon-specific categoriziation of this tag.  Users do not generally need to know subgroups; instead, we use naming conventions similar to other photo editors.
    tagName As String               'The second half of FullGroupAndName
    TagNameFriendly As String       'The human-friendly tag name (supports spaces and special chars)
    TagTable As String              'The primary categorization of this tag, e.g. "Exif::Main", "JPEG::Adobe", "ICC_Profile::Main"
    TagID As String                 'The low-level, format-specific ID of a given tag.  For many tags, this is just a string matching the tag's Exiftool name.  For some tag types, however, (e.g. EXIF), this will be a numeric ID corresponding to the actual spec-defined "id" of a given tag.
    TagValueFriendly As String      'The human-readable version of a tag value, e.g. "YCbCr4:4:4 (1 1)" instead of "1 1".
    TagValue As String              'The low-level version of a tag value.  For many tags, this is the same as TagValueFriendly.
    HasIndex As Boolean             'Indicates the presence of an "et:index" identifier in the RDF description.  This is only supplied under rare circumstances, e.g. if the same tag appears in multiple groups.
    IsTagList As Boolean            'Indicates the presence of a list-type tag, common with XMP chunks coming from Photoshop.  The friendly tag name contains a semicolor-delimited list of tag values.
    IsTagBinary As Boolean          'Indicates the presence of a base64-encoded binary tag.
    WasBinaryExtracted As Boolean   'Normally, we skip the extraction of binary data as it can be enormous and time-consuming to process.  However, if the user requests binary extraction via preference, we can go ahead and retrieve the data for them.  This tag indicates that we performed such an extraction, and the result is stored inside TagBase64Value.
    InternalUseOnly As Boolean      'Some tags (like ExifTool version) have no relevance to the end-user.  We still want to track these, but we tag them so that they are not exposed to the user.
    TagIndexInternal As Long        'Only meaningful if HasIndex (above) is TRUE.
    TagBase64Value As String        'Only meaningful if IsTagBinary (above) is TRUE.
    
    'Used to flag tags that need to be removed by the image export engine.  Users can set this flag
    ' through the Image > Metadata > Edit Metadata dialog.
    TagMarkedForRemoval As Boolean
    
    'When saving image files, users can also select the convenient "remove tags with privacy concerns" option.
    ' When that flag is set, PD will mark this internal removal flag.  (This is kept distinct from the
    ' normal marked-for-removal flag, because we don't want to save this value persistently to PDI files,
    ' or display it to the user in the normal metadata edit dialog.
    TagMarkedForRemoval_Export As Boolean
    
    'Used to flag tags that the user has touched from the metadata editing dialog
    UserModifiedThisSession As Boolean
    UserModifiedAllSessions As Boolean
    UserValueNew As String
    UserIDNew As String
    
    'IMPORTANT NOTE!  All values past this line are *not* filled in automatically.  They must be manually filled by parsing
    ' the ExifTool database file for the tag's matching attributes.  This is typically handled by the Metadata editing window.
    
    'These five values will always be loaded by a database pass
    DB_TagHitDatabase As Boolean    'TRUE if this tag object has already been filled with its database information.
                                    ' (This is updated per-session, so closing the image and reloading it will reset this value.)
    DB_IsWritable As Boolean        'TRUE if ExifTool is capable of writing/updating this tag
    DB_TypeCount As Long            'TRUE if this tag has a type like "byte x 4" instead of "int_32"
    DB_DataType As String           'The string representation of this tag's datatype; this is primarily used for debugging
    DB_DataTypeStrict As PD_Metadata_Datatype   'Please use this version of the tag's datatype, not the string
    
    'These values are filled on an as-needed basis; they are only specified if a tag requires it.
    ' (On most tags, these will be FALSE.)
    DBF_IsAvoid As Boolean
    DBF_IsBag As Boolean
    DBF_IsBinary As Boolean
    DBF_IsFlattened As Boolean
    DBF_IsList As Boolean
    DBF_IsMandatory As Boolean
    DBF_IsPermanent As Boolean
    DBF_IsProtected As Boolean
    DBF_IsSequence As Boolean
    DBF_IsUnknown As Boolean
    DBF_IsUnsafe As Boolean
    
    'Database description should always match the "friendly name", above, but we retrieve a database copy "just in case"
    DB_Description As String
    
    'If a tag provides its own hard-coded list of possible values, this will be set to TRUE, and the stacks will be populated
    ' with (DB_HardcodedListSize - 1) values
    DB_HardcodedList As Boolean
    DB_HardcodedListSize As Long
    DB_StackIDs As pdStringStack
    DB_StackValues As pdStringStack
    
    'Raw copy of the database XML packet associated with this tag, "just in case".  Do not use this for anything but parsing.
    TagDebugData As String
    
End Type

Public Enum PD_Metadata_Datatype
    MD_int8s       '- Signed 8-bit integer                    (EXIF 'SBYTE')
    MD_int8u       '- Unsigned 8-bit integer                  (EXIF 'BYTE')
    MD_int16s      '- Signed 16-bit integer                   (EXIF 'SSHORT')
    MD_int16u      '- Unsigned 16-bit integer                 (EXIF 'SHORT')
    MD_int32s      '- Signed 32-bit integer                   (EXIF 'SLONG')
    MD_int32u      '- Unsigned 32-bit integer                 (EXIF 'LONG')
    MD_int64s      '- Signed 64-bit integer                   (BigTIFF 'SLONG8')
    MD_int64u      '- Unsigned 64-bit integer                 (BigTIFF 'LONG8')
    MD_rational32s '- Rational consisting of 2 int16s values
    MD_rational32u '- Rational consisting of 2 int16u values
    MD_rational64s '- Rational consisting of 2 int32s values  (EXIF 'SRATIONAL')
    MD_rational64u '- Rational consisting of 2 int32u values  (EXIF 'RATIONAL')
    MD_fixed16s    '- Signed 16-bit fixed point value
    MD_fixed16u    '- Unsigned 16-bit fixed point value
    MD_fixed32s    '- Signed 32-bit fixed point value
    MD_fixed32u    '- Unsigned 32-bit fixed point value
    MD_float       '- 32-bit IEEE floating point value        (EXIF 'FLOAT')
    MD_double      '- 64-bit IEEE floating point value        (EXIF 'DOUBLE')
    MD_extended    '- 80-bit extended floating float
    MD_ifd         '- Unsigned 32-bit integer sub-IFD pointer (EXIF 'IFD')
    MD_ifd64       '- Unsigned 64-bit integer sub-IFD pointer (BigTIFF 'IFD8')
    MD_string      '- Series of 8-bit ASCII characters        (EXIF 'ASCII').
                   '  Note that PD condenses other string types down to MD_string, while silently handling conversions.
    MD_undef       '- Undefined-format binary data            (EXIF 'UNDEFINED')
    MD_binary      '- Binary data (same as 'undef')
    MD_integerstring    '- XMP type (e.g. encoded as a string, but it must adhere to certain formatting criteria)
    MD_floatstring      '- XMP type (e.g. encoded as a string, but it must adhere to certain formatting criteria)
    MD_rationalstring   '- XMP type (e.g. encoded as a string, but it must adhere to certain formatting criteria)
    MD_datestring       '- XMP type (e.g. encoded as a string, but it must adhere to certain formatting criteria)
    MD_booleanstring    '- XMP type (e.g. encoded as a string, but it must adhere to certain formatting criteria)
    MD_digits      '- IPTC type, basically a list of ASCII digits, restricted by count
    
End Enum

#If False Then
Private Const MD_int8s = 0, MD_int8u = 0, MD_int16s = 0, MD_int16u = 0, MD_int16uRev = 0, MD_int32s = 0, MD_int32u = 0
Private Const MD_int64s = 0, MD_int64u = 0, MD_rational32s = 0, MD_rational32u = 0, MD_rational64s = 0, MD_rational64u = 0
Private Const MD_fixed16s = 0, MD_fixed16u = 0, MD_fixed32s = 0, MD_fixed32u = 0, MD_float = 0, MD_double = 0
Private Const MD_extended = 0, MD_ifd = 0, MD_ifd64 = 0, MD_string = 0, MD_undef = 0, MD_binary = 0
Private Const MD_integerstring = 0, MD_floatstring = 0, MD_rationalstring = 0, MD_datestring = 0, MD_booleanstring = 0, MD_digits = 0
#End If

'pdFSO is used for Unicode file interop
Private m_FSO As pdFSO

'Once ExifTool has been run at least once, this will be set to TRUE.  If TRUE, this means that the pdPipeAsync
' class declared below is active and connected to ExifTool, and can be used to send/receive input and output.
Private m_IsExifToolRunning As Boolean
Private m_Async As pdPipeAsync

'Because ExifTool parses metadata asynchronously, we will gather its output as it comes.  This string will hold whatever
' XML data ExifTool has returned so far.
Private m_currentMetadataText As String

'While capture mode is active (e.g. while retrieving metadata information we care about) this will be set to TRUE.
Private m_captureModeActive As Boolean

'While verification mode is active (e.g. we only care about ExifTool succeeding or failing), this will be set to TRUE.
Private m_VerificationModeActive As Boolean
Private m_VerificationString As String

'Once in its lifetime, ExifTool will be asked to write its full tag database to an XML file (~6.2 mb worth of data).
' This will normally happen the first time the Metadata Editor is initiated.  While that special mode is running,
' this variable will be set to TRUE. (Note that the matching database string will be available under a variety of
' circumstances, as it is filled whenever the user initiates a metadata editing session.)
Private m_DatabaseModeActive As Boolean
Private m_DatabaseHandle As Long, m_DatabaseString As String, m_DatabasePath As String

'As of 6.6 alpha, technical metadata reports can now be generated for a given file.  While this mode is active, we do not
' want to immediately delete the report; use this boolean to check for that particular state.
Private m_technicalReportModeActive As Boolean
Private m_technicalReportSrcImage As String

'As of 7.0, ExifTool can be used to extract esoteric ICC profiles that FreeImage might miss.
Private m_ICCExtractionModeActive As Boolean
Private m_ICCExtractionSrcImage As String

'Prior to writing out a file's metadata, we must cache the information we want written in a temp file.  (ExifTool requires
' a source file when writing metadata out to file; the alternative is to manually request the writing of each tag in turn,
' but if we do this, we lose many built-in utilities like automatically removing duplicate tags, and reassigning invalid
' tags to preferred categories.)  Because writing out metadata is asynchronous, we have to wait for ExifTool to finish
' before deleting the temp file, so we keep a copy of the file's path here.  The StopVerificationMode (which is
' automatically triggered by the NewMetadataReceived function as necessary) will remove the file at this location.
Private m_tmpMetadataFilePath As String

'If multiple images are loaded simultaneously, we have to do some tricky handling to parse out their individual bits.  As such, we store
' the ID of the last image metadata request we received; only when this ID is returned successfully do we consider metadata processing
' "complete", and return TRUE for isMetadataFinished.
Private m_LastRequestID As Long

'ExifTool is built with Perl, and Perl requires a temp folder where it can dump a bunch of Perl-related resources.  By default,
' this is the current user's temp folder, but we prefer to cache this folder locally, inside PD's /Data subfolder.
' (This path is set when ExifTool is started.)  Note that ExifTool's database is also stored here.
Private m_ExifToolDataFolder As String

'Parsing the ExifTool database is a complicated and unpleasant process; limited local caching helps alleviate some of the pain
Private Type ET_GROUP
    GroupName As String
    GroupStart As Long
    GroupEnd As Long
End Type

Private Const INIT_GROUP_CACHE_SIZE As Long = 8
Private m_GroupCache() As ET_GROUP
Private m_NumGroupsInCache As Long
Private m_ModalWaitWindowActive As Boolean

'This module has to do a ton of XML parsing.  Rather than recreate a parser on every call, we just reuse a stock one
Private m_ParseXML As pdSerialize

Public Function IsDatabaseModeActive() As Boolean
    IsDatabaseModeActive = m_DatabaseModeActive
End Function

Public Function IsVerificationModeActive() As Boolean
    IsVerificationModeActive = m_VerificationModeActive
End Function

'm_Async will asynchronously trigger this function whenever it receives new metadata from ExifTool.
Public Sub NewMetadataReceived()
    
    If m_captureModeActive Then
        m_currentMetadataText = m_currentMetadataText & m_Async.GetDataAsString()
    
        'Sometimes it is helpful to retrieve metadata from a failed-to-import image.  If you want to
        ' examine ExifTool's returned XML output yourself (e.g. while debugging), you can dump the
        ' full XML contents to the clipboard by enabling the EXIFTOOL_DEBUGGING_ENABLED constant
        ' inside the Exiftool module.  (Note that a forcible ANSI conversion will be applied,
        ' so extended chars may not survive.)
        If (EXIFTOOL_DEBUGGING_ENABLED <> 0) Then
            If ExifTool.IsMetadataFinished() Then
                Clipboard.Clear
                Clipboard.SetText m_currentMetadataText
            End If
        End If
        
    'During verification mode, we must also check to see if verification has finished
    ElseIf m_VerificationModeActive Then
        m_VerificationString = m_VerificationString & m_Async.GetDataAsString()
        If IsMetadataFinished() Then StopVerificationMode
        
    'During database mode, check for a finish state, then write the retrieved database out to file!
    ElseIf m_DatabaseModeActive Then
        If IsMetadataFinished() Then
            WriteMetadataDatabaseToFile
            If m_ModalWaitWindowActive Then g_UnloadWaitWindow = True
        End If
    End If
    
End Sub

Private Sub WriteMetadataDatabaseToFile()
    
    'Remember that we haven't actually removed any data from the async class - it has cached everything for us.
    ' As such, we just want it to dump everything it has, as-is, to the metadata file (which is still open).
    ' Note that we also forcibly remove the "{ready}" flag, plus a linebreak, from the end of the buffer.
    Dim writeSize As Long
    writeSize = m_Async.GetSizeOfInputBuffer - (7 + 2)  'Len("{ready}") + Len(vbCrLf)
    If (writeSize > 0) Then m_FSO.FileWriteData m_DatabaseHandle, m_Async.PeekPointer(0), writeSize
    m_FSO.FileCloseHandle m_DatabaseHandle
    
    If (writeSize > 0) Then
        PDDebug.LogAction "ExifTool Metadata database created successfully (" & Files.FileLenW(m_DatabasePath) & " bytes)"
    Else
        PDDebug.LogAction "WARNING!  ExifTool.WriteMetadataDatabaseToFile failed to write the metadata database to file."
    End If
        
    'Reset the database mode tracker, so the database doesn't accidentally get rebuilt again!
    m_DatabaseModeActive = False
    m_Async.ResetInputBuffer
    
End Sub

'When we only care about ExifTool succeeding or failing, use this function to enter "verification mode", which simply checks
' to see if ExifTool has finished its previous request.
Private Sub StartVerificationMode()
    m_VerificationModeActive = True
    m_VerificationString = vbNullString
End Sub

'When verification mode ends (as triggered by an automatic check in newMetadataReceived), we must also remove the temporary
' file we created that held the data being exported via ExifTool.
Private Sub StopVerificationMode()
    
    m_VerificationModeActive = False
    
    'Standard metadata embed requests are easy - we just need to delete our temporary file
    If (Not m_technicalReportModeActive) And (Not m_ICCExtractionModeActive) Then
        
        m_VerificationString = vbNullString
        
        'Verification mode is a bit different.  We need to erase our temporary metadata file if it exists; then we can exit.
        If Files.FileDeleteIfExists(m_tmpMetadataFilePath) Then
            PDDebug.LogAction "Metadata embedding finished."
        Else
            PDDebug.LogAction "WARNING: metadata embedding finished, but temp file remains: " & m_tmpMetadataFilePath
        End If
        
    Else
        
        If m_technicalReportModeActive Then
        
            'Replace the {ready} text supplied by ExifTool itself, which will be at the end of the metadata report
            If (LenB(m_VerificationString) <> 0) Then m_VerificationString = Replace$(m_VerificationString, "{ready}", vbNullString)
            
            'Write the completed technical report out to a temp file
            Dim tmpFilename As String
            tmpFilename = UserPrefs.GetTempPath & "MetadataReport_" & Files.FileGetName(m_technicalReportSrcImage, True) & ".html"
            Files.FileSaveAsText m_VerificationString, tmpFilename
            
            'Shell the default HTML viewer for the user
            m_VerificationString = vbNullString
            Web.OpenURL tmpFilename
            
            m_technicalReportSrcImage = vbNullString
            m_technicalReportModeActive = False
            
        End If
        
        If m_ICCExtractionModeActive Then m_ICCExtractionModeActive = False
        
    End If
    
End Sub

'After using ExifTool to extract an ICC profile to a standalone file, you can retrieve the destination filename
' via this function.
Public Function GetExtractedICCProfilePath() As String
    GetExtractedICCProfilePath = m_ICCExtractionSrcImage
End Function

'Returns TRUE if ExifTool is currently parsing metadata asynchronously
Public Function IsMetadataPipeActive() As Boolean
    IsMetadataPipeActive = m_captureModeActive
End Function

'When ExifTool has completed work on metadata, it will send "{ready}" to stdout.  We check for the presence of "{ready}" in
' the metadata string to see if ExifTool is done.
Public Function IsMetadataFinished() As Boolean
    
    'If ExifTool is not available, or if it failed to start, simply return TRUE which will allow any waiting code
    ' to continue.
    If (Not m_IsExifToolRunning) Then
        IsMetadataFinished = True
        Exit Function
    End If
    
    'If no capture or verification modes are active, end any lingering metadata processing.
    If (Not m_captureModeActive) And (Not m_VerificationModeActive) And (Not m_DatabaseModeActive) Then
        IsMetadataFinished = True
        Exit Function
    End If
    
    'Different verification modes require different checks for completion.
    If m_captureModeActive Then
        m_captureModeActive = (InStr(1, m_currentMetadataText, "{ready" & m_LastRequestID & "}", vbBinaryCompare) = 0)
        IsMetadataFinished = (Not m_captureModeActive)
        
    ElseIf m_VerificationModeActive Then
        m_VerificationModeActive = (InStr(1, m_VerificationString, "{ready}", vbBinaryCompare) = 0)
        IsMetadataFinished = (Not m_VerificationModeActive)
    
    'In database mode, we only want to peek at the last few bytes received from ExifTool.  If they equal
    ' {ready}, it means ExifTool has finished sending us data.
    ElseIf m_DatabaseModeActive Then
        m_DatabaseModeActive = (InStr(1, m_Async.PeekLastNBytes(), "{ready}", vbBinaryCompare) = 0)
        IsMetadataFinished = (Not m_DatabaseModeActive)
    
    End If
    
End Function

'When metadata is ready (as determined by a call to isMetadataFinished), it can be retrieved via this function
Public Function RetrieveMetadataString() As String
    
    If (LenB(m_currentMetadataText) <> 0) Then
    
        'Because we request metadata in XML format, ExifTool escapes disallowed XML characters.  Convert those back
        ' to standard characters before returning the retrieved metadata.
        If (InStr(1, m_currentMetadataText, "&#39;") > 0) Then m_currentMetadataText = Replace(m_currentMetadataText, "&#39;", "'")
        If (InStr(1, m_currentMetadataText, "&quot;") > 0) Then m_currentMetadataText = Replace(m_currentMetadataText, "&quot;", """")
        If (InStr(1, m_currentMetadataText, "&gt;") > 0) Then m_currentMetadataText = Replace(m_currentMetadataText, "&gt;", ">")
        If (InStr(1, m_currentMetadataText, "&lt;") > 0) Then m_currentMetadataText = Replace(m_currentMetadataText, "&lt;", "<")
        If (InStr(1, m_currentMetadataText, "&amp;") > 0) Then m_currentMetadataText = Replace(m_currentMetadataText, "&amp;", "&")
        
    End If
        
    'Return the processed string, then erase our copy of it
    RetrieveMetadataString = m_currentMetadataText
    m_currentMetadataText = vbNullString
    
End Function

Public Function RetrieveUntouchedMetadataString() As String
    RetrieveUntouchedMetadataString = m_currentMetadataText
End Function

'If asynchronous metadata loading is too slow, it gets handed off to a timer on FormMain.  (Why not site the
' timer here?  Because modules in VB6 can't receive events without ugly workarounds, ugh.)  That timer will
' ping us when loading finishes.
Public Sub FinishAsyncMetadataLoading()

    'I don't like resorting to hackneyed error-handling, but ExifTool can be unpredictable,
    ' especially when loading a bajillion images simultaneously.  Rather than bring down the
    ' whole program, I'd prefer to simply ignore metadata for problematic image(s).
    On Error GoTo MetadataBroken

    'Retrieve the completed metadata string
    Dim mdString As String, tmpString As String
    mdString = ExifTool.RetrieveMetadataString()
    
    Dim curImageID As Long
    
    'Now comes some messy string parsing.  If the user has loaded multiple images at once,
    ' the metadata string returned by ExifTool will contain ALL METADATA for ALL IMAGES
    ' in one giant string.  We need to parse out each image's metadata, supply it to the
    ' correct pdImage object, then repeat until all images have received their metadata.
    
    'Start by finding the first occurrence of ExifTool's unique "{ready}" message,
    ' which signifies its success in completing a single coherent "-execute" request.
    Dim startPosition As Long, terminalPosition As Long
    startPosition = 1
    terminalPosition = InStr(1, mdString, "{ready", vbBinaryCompare)
    
    Do While (terminalPosition <> 0)
    
        'terminalPosition now contains the position of ExifTool's "{ready123}" tag, where 123 is
        ' the ID of the image whose metadata is contained *prior* to that point.
        '
        'Start by figuring out what that ID number actually is.
        Dim lenFailsafe As Long
        
        If (terminalPosition + 6 < Len(mdString)) Then
            lenFailsafe = InStr(terminalPosition + 6, mdString, "}", vbBinaryCompare) - (terminalPosition + 6)
        Else
            lenFailsafe = 0
        End If
        
        If (lenFailsafe <> 0) Then
            
            'Attempt to retrieve the relevant image ID for this section of metadata
            If ((terminalPosition + 6 + lenFailsafe) < Len(mdString)) Then
            
                tmpString = Mid$(mdString, terminalPosition + 6, lenFailsafe)
                
                If IsNumeric(tmpString) Then
                    curImageID = CLng(tmpString)
                Else
                    PDDebug.LogAction "Metadata ID invalid: " & tmpString
                    Exit Sub
                End If
                
                'Now we know where the metadata for this image *ends*, but we still need to determine
                ' where it *starts*.  All metadata XML entries start with a standard XML header.
                ' Search backwards from the {ready123} message until such a header is found.
                startPosition = InStrRev(mdString, "<?xml", terminalPosition, vbBinaryCompare)
                
                'Using the start and final markers, extract the relevant metadata and forward it to the
                ' matching pdImage object
                If (startPosition > 0) And ((terminalPosition - startPosition) > 0) Then
                    
                    'Make sure we calculated our curImageID value correctly
                    If PDImages.IsImageActive(curImageID) Then
                        
                        'Create the imgMetadata object as necessary, and load the selected metadata into it!
                        If (PDImages.GetImageByID(curImageID).ImgMetadata Is Nothing) Then Set PDImages.GetImageByID(curImageID).ImgMetadata = New pdMetadata
                        PDImages.GetImageByID(curImageID).ImgMetadata.LoadAllMetadata Mid$(mdString, startPosition, terminalPosition - startPosition), curImageID
                        
                        'Now comes kind of a weird requirement.  Because metadata is loaded asynchronously, it may
                        ' arrive AFTER the image import engine has already written our first Undo entry out to file
                        ' (this happens at image load-time, so we have a backup if the original file disappears).
                        '
                        'If this occurs, request a rewrite from the Undo engine to ensure that metadata gets added
                        ' to the base entry of the Undo/Redo stack.
                        If PDImages.GetImageByID(curImageID).UndoManager.HasFirstUndoWriteOccurred Then
                            PDDebug.LogAction "Adding late-arrival metadata to original undo entry..."
                            PDImages.GetImageByID(curImageID).UndoManager.ForceLastUndoDataToIncludeEverything
                        End If
                        
                    End If
                    
                    'Find the next chunk of image metadata, if any
                    terminalPosition = InStr(terminalPosition + 6, mdString, "{ready", vbBinaryCompare)
                    
                Else
                    PDDebug.LogAction "metadata parse failure"
                    terminalPosition = 0
                End If
                                    
            Else
                PDDebug.LogAction "metadata parse failure"
                terminalPosition = 0
            End If
            
        Else
            PDDebug.LogAction "metadata parse failure"
            terminalPosition = 0
        End If
    
    Loop
    
    Exit Sub
    
MetadataBroken:
    PDDebug.LogAction "WARNING: couldn't retrieve metadata."
    
End Sub

'Retrieve the currently installed ExifTool version.  (If ExifTool cannot be found, this will return FALSE.)
Public Function GetExifToolVersion() As String
    
    If PluginManager.IsPluginCurrentlyInstalled(CCP_ExifTool) Then
        
        Const EXIFTOOL_EXE_NAME As String = "exiftool.exe"
        Dim exiftoolPath As String
        exiftoolPath = PluginManager.GetPluginPath & EXIFTOOL_EXE_NAME
        
        Dim cShell As pdPipeSync
        Set cShell = New pdPipeSync
        If cShell.RunAndCaptureOutput(exiftoolPath, EXIFTOOL_EXE_NAME & " -ver", False) Then
            
            'The output string will generally be a simple version number, e.g. "XX.YY", and it will be
            ' terminated by a vbCrLf character.  Remove vbCrLf now.
            Dim outputString As String
            outputString = Trim$(cShell.GetStdOutDataAsString())
            If (InStr(1, outputString, vbCrLf, vbBinaryCompare) <> 0) Then outputString = Replace(outputString, vbCrLf, vbNullString)
            
            'Development versions of ExifTool (e.g. any version number that is not a multiple of 10) may include
            ' a warning about the current "official" version of the library.  This warning is placed at the end
            ' of the version number, using formatting like: "10.01 [Warning: Library version is 10.00]".
            
            'Look for such trailing tags and remove them if present.
            If (InStr(1, outputString, "[", vbBinaryCompare) <> 0) Then outputString = Left$(outputString, InStr(1, outputString, "[", vbBinaryCompare) - 1)
            GetExifToolVersion = Trim$(outputString)
            
        Else
            GetExifToolVersion = vbNullString
        End If
        
    End If
    
End Function

'Start an ExifTool instance (if one isn't already active), and have it process an image file.  Because we now run ExifTool
' asynchronously, this should be done early in the image load process.
Public Function StartMetadataProcessing(ByVal srcFile As String, ByRef dstImage As pdImage) As Boolean

    'If ExifTool is not running, start it.  If it cannot be started, exit.
    If (Not m_IsExifToolRunning) Then
        If (Not StartExifTool()) Then
            Message "ExifTool could not be started.  Metadata unavailable for this session."
            StartMetadataProcessing = False
            Exit Function
        End If
    End If
    
    'Notify the program that stdout capture has begun
    m_captureModeActive = True
    
    'Erase any previous metadata caches.
    ' NOTE! Upon implementing PD's new asynchronous metadata retrieval mechanism, we don't want to
    '        erase the original metadata string, as its construction may lag behind the rest of the
    '        image load process.  When a full metadata string is retrieved, the RetrieveMetadataString()
    '        function now handles clearing for us.
    'm_currentMetadataText = vbNullString
    
    'Start building a string of ExifTool parameters.  We will send these parameters to stdIn, but ExifTool expects
    ' them in ARGFILE format, e.g. each parameter on its own line.  pdString is particularly well-suited to
    ' constructing arguments like this.
    Dim cmdParams As pdString
    Set cmdParams = New pdString
    
    'Ignore minor errors and warnings
    cmdParams.AppendLine "-m"
    
    'To support Unicode filenames, explicitly request UTF-8-compatible parsing.
    cmdParams.AppendLine "-charset"
    cmdParams.AppendLine "filename=UTF8"
    
    'Output long-format data
    cmdParams.AppendLine "-l"
        
    'Request a custom separator for list-type values
    cmdParams.AppendLine "-sep"
    cmdParams.AppendLine ";;;"
        
    'If a translation is active, request descriptions in the current language
    If g_Language.TranslationActive Then
        cmdParams.AppendLine "-lang"
        cmdParams.AppendLine g_Language.GetCurrentLanguage((g_Language.GetCurrentLanguage(False) = "zh"))
    End If
    
    'If the user wants us to estimate JPEG quality, do so now
    If UserPrefs.GetPref_Boolean("Loading", "Metadata Estimate JPEG", True) Then
        cmdParams.AppendLine "-api"
        cmdParams.AppendLine "RequestTags=JPEGQualityEstimate,JPEGDigest"
    End If
    
    'If the user wants us to extract binary data, do so now
    If UserPrefs.GetPref_Boolean("Loading", "Metadata Extract Binary", False) Then
        cmdParams.AppendLine "-b"
    End If
    
    'If the user wants us to extract unknown tags, do so now.
    If UserPrefs.GetPref_Boolean("Loading", "Metadata Extract Unknown", False) Then
        cmdParams.AppendLine "-u"
    End If
    
    'If the user wants us to expose duplicate tags, do so now.  (Default behavior is to suppress duplicates.)
    If UserPrefs.GetPref_Boolean("Loading", "Metadata Hide Duplicates", True) Then
        cmdParams.AppendLine "--a"
    End If
    
    'Forcibly retrieve a binary copy of the entire ICC profile, if available
    cmdParams.AppendLine "-api"
    cmdParams.AppendLine "RequestTags=ICC_Profile"
    
    'Historically, we needed to explicitly set a charset; this shouldn't be necessary with current versions
    ' (as UTF-8 is automatically supported), but if desired, specific metadata types can be coerced into
    ' a character set like so:
    'cmdParams.AppendLine "-charset"
    'cmdParams.AppendLine "UTF8"
    
    'IPTC metadata is an exception to the rule above.  IPTC supports charset markers, but in my experience, these are
    ' rarely used.  ExifTool defaults to the current code page for conversion if we don't specify otherwise, so UTF-8
    ' is preferable for this one metadata type.
    cmdParams.AppendLine "-charset"
    cmdParams.AppendLine "iptc=UTF8"
            
    'If we requested binary data in a previous step, this will cause preview and thumbnail images to be processed.
    ' If this behavior is unwanted, we can deny them specifically - but at present, we parse them alongside other
    ' binary data, if the user has gone to the trouble of requesting binary extraction.
    'cmdParams.AppendLine "-x"
    'cmdParams.AppendLine "PreviewImage"
    'cmdParams.AppendLine "-x"
    'cmdParams.AppendLine "ThumbnailImage"
    'cmdParams.AppendLine "-x"
    'cmdParams.AppendLine "PhotoshopThumbnail"
    
    'Request XML output.  This results in far more complex extraction, but it's the only way to retrieve descriptions
    ' and names simultaneously.
    cmdParams.AppendLine "-X"
    
    'Include tag table information (e.g. additional technical details on each tag).  Note that this setting affects
    ' the XML parsing code, so you cannot comment it out without making matching changes inside pdMetadata.
    cmdParams.AppendLine "-t"
    
    'Add the source image's path
    cmdParams.AppendLine srcFile
    
    'Finally, add the special command "-execute" which tells ExifTool to start working.  We also append the current
    ' image's canonical ID value.  When ExifTool finishes, it will return this value; this is crucial for assigning
    ' incoming metadata to the correct image, if the user loads a bunch of images at once (and metadata streams in
    ' asynchronously).
    cmdParams.AppendLine "-execute" & dstImage.imageID
    
    'Note this request ID as being the last one we received; only when this ID is returned by ExifTool will we
    ' actually consider the metadata "complete".
    m_LastRequestID = dstImage.imageID
    
    'Ask the async user control to start processing this image's metadata.  It handle interop from here.
    If (Not m_Async Is Nothing) Then m_Async.SendData cmdParams.ToString()
    
    StartMetadataProcessing = True
    
End Function

'ExifTool has a lot of great facilities for analyzing image metadata.  Technical users in particular might want to take advantage
' of ExifTool's "htmldump" facility, which provides a detailed hex report of all metadata in a file.  This function can be used
' to generate such a report, but note that it only works for images that exist on disk (obviously).
Public Function CreateTechnicalMetadataReport(ByRef srcImage As pdImage) As Boolean
    
    'Start by checking for an existing copy of the XML database.  If it already exists, no need to recreate it.
    If Files.FileExists(srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk")) Then
    
        'A string builder improves performance and cuts down on string table thrashing
        Dim cmdParams As pdString
        Set cmdParams = New pdString
        
        'Add the htmldump command
        cmdParams.AppendLine "-htmldump"
        
        'Add -u, which will also report unknown tags
        cmdParams.AppendLine "-u"
                
        'To support Unicode filenames, explicitly request UTF-8-compatible parsing.
        cmdParams.AppendLine "-charset"
        cmdParams.AppendLine "filename=UTF8"
        
        'Add the source image to the list
        m_technicalReportSrcImage = srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk")
        cmdParams.AppendLine srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk")
        
        'Finally, add the special command "-execute" which tells ExifTool to start operations
        cmdParams.AppendLine "-execute"
        
        'Activate verification mode.  This will asynchronously wait for the metadata to be written out to file, and when it
        ' has finished, it will erase our temp file.
        m_technicalReportModeActive = True
        StartVerificationMode
        
        'Ask the user control to start processing this image's metadata.  It will handle things from here.
        If (Not m_Async Is Nothing) Then m_Async.SendData cmdParams.ToString()
        
        CreateTechnicalMetadataReport = True
    
    Else
        CreateTechnicalMetadataReport = False
    End If

End Function

'Extract an image's ICC profile to a standalone file.  If no destination filename is used, a temporary file will be generated.
' Use the ExifTool.GetExtractedICCProfilePath() function to retrieve said filename.  If you pass your own filename,
' make *absolutely certain* it ends in .icc or icm, or ExifTool may not extract the profile correctly.
Public Function ExtractICCMetadataToFile(ByRef srcImage As pdImage, Optional ByVal dstFilename As String = vbNullString) As Boolean
    
    'For this to work, the target file must exist on disk.  (ExifTool requires a disk copy to extract the
    ' ICC profile out to file.)
    If Files.FileExists(srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk")) Then
    
        Dim cmdParams As pdString
        Set cmdParams = New pdString
        
        'Extract the icc profile
        cmdParams.AppendLine "-icc_profile"
        
        'To do this correctly, we must also request the processing of binary-type tags
        cmdParams.AppendLine "-b"
        
        'We want to *write* a new file, but instead of using "-w" (which only takes an extension argument),
        ' use "-o" which lets us specify the full output path.
        Dim tmpFilename As String
        If (LenB(dstFilename) = 0) Then
            tmpFilename = Files.RequestTempFile()
            tmpFilename = tmpFilename & ".icc"
        Else
            tmpFilename = dstFilename
        End If
        
        cmdParams.AppendLine "-o"
        cmdParams.AppendLine tmpFilename
        
        'Cache the filename at module level, so we can retrieve it when we're done
        m_ICCExtractionSrcImage = tmpFilename
        
        'Finally, add the original filename
        cmdParams.AppendLine srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk")
        
        'Finally, add the special command "-execute" which tells ExifTool to start operations
        cmdParams.AppendLine "-execute"
        
        'Activate verification mode.  This will asynchronously wait for the metadata to be written out to file, and when it
        ' has finished, it will erase our temp file.
        m_ICCExtractionModeActive = True
        StartVerificationMode
        
        'Ask the user control to start processing this image's metadata.  It will handle things from here.
        If (Not m_Async Is Nothing) Then m_Async.SendData cmdParams.ToString()
        
        ExtractICCMetadataToFile = True
    
    Else
        ExtractICCMetadataToFile = False
    End If

End Function

Public Function DoesTagDatabaseExist() As Boolean
    DoesTagDatabaseExist = Files.FileExists(m_DatabasePath)
End Function

Public Function ShowMetadataDialog(ByRef srcImage As pdImage, Optional ByRef parentForm As Form = Nothing) As Boolean

    'Perform a failsafe check to make sure the metadata object exists.  (If ExifTool is missing, it may
    ' not be present!)
    If (Not srcImage.ImgMetadata Is Nothing) Then
        
        'In the future, we'll allow the user to add their own metadata to the current image.  At present,
        ' however, there's not much point in displaying a dialog if the image doesn't have metadata.
        If srcImage.ImgMetadata.HasMetadata Then
            
            'Make sure the metadata database exists.  If it doesn't, create it.
            If (Not ExifTool.DoesTagDatabaseExist) Or ExifTool.IsDatabaseModeActive Then
                 
                If (Not ExifTool.DoesTagDatabaseExist) Then ExifTool.WriteTagDatabase
                
                Dim waitTitle As String, waitDescription As String
                waitTitle = g_Language.TranslateMessage("Please wait while the tag database is created...")
                waitDescription = g_Language.TranslateMessage("The tag database handles technical details of the 20,000+ metadata tags supported by PhotoDemon.  Creating the database takes 10 to 15 seconds, and it only needs to be created once, when the metadata editor is used for the first time.")
                
                'When raising the metadata dialog from a save dialog, we cannot display the "please wait" window modelessly
                ' (as the save dialog will be modal).  As such, we must use different methods for unloading it.
                If Not (parentForm Is Nothing) Then
                    m_ModalWaitWindowActive = True
                    Interface.DisplayWaitScreen waitTitle, parentForm, waitDescription, True
                    m_ModalWaitWindowActive = False
                    
                'Raising the metadata dialog from the main window can be done modelessly
                Else
                    Interface.DisplayWaitScreen waitTitle, FormMain, waitDescription, False
                End If
                
                Do
                    DoEvents
                Loop While ExifTool.IsDatabaseModeActive
                
                Interface.HideWaitScreen
                
            End If
            
            'With the database successfully constructed, we now need to load it into memory
            If (LenB(m_DatabaseString) = 0) Then Files.FileLoadAsString m_DatabasePath, m_DatabaseString
            
            'Metadata caching is performed on a per-image basis, so we need to reset the cache on each invocation
            ExifTool.StartNewDatabaseCache
            
            ShowPDDialog vbModal, FormMetadata
        
        'TODO 8.2: still raise the form, and allow the user to add their own metadata to the image
        Else
            Message "No metadata available"
            PDMsgBox "This image does not contain metadata.", vbInformation Or vbOKOnly, "No metadata available"
        End If
        
    End If
            
End Function

'If the user wants to edit an image's metadata, we need to know which tags are writeable and which are not.  Also, it's helpful to
' know things like each tag's datatype (to verify output before it's passed along to ExifTool).  If ExifTool is successfully initialized
' at program startup, this function will be called, and its job is to populate ExifTool's tag database.
Public Function WriteTagDatabase() As Boolean

    'Start by checking for an existing copy of the XML database.  If it already exists, no need to recreate it.
    ' (TODO: check the database version number, as new tags may be added between releases...)
    If Files.FileExists(m_DatabasePath) Then
        WriteTagDatabase = True
    Else
    
        'Database wasn't found.  Generate a new copy now.
        
        'Start metadata database retrieval mode
        m_DatabaseModeActive = True
        m_DatabaseString = vbNullString
        
        'Open a persistent handle to the database itself.  We'll stream data from ExifTool directly into
        ' this file.
        Set m_FSO = New pdFSO
        If (Not m_FSO.FileCreateHandle(m_DatabasePath, m_DatabaseHandle, True, True, OptimizeSequentialAccess)) Then
            PDDebug.LogAction "WARNING!  Failed to create ExifTool database.  Metadata editing is disabled for this session."
            m_DatabaseModeActive = False
            WriteTagDatabase = False
            Exit Function
        End If
        
        'To simplify text file heuristics, forcibly write a UTF-8 BOM to the start of the file.
        Dim bomMarker(0 To 2) As Byte
        bomMarker(0) = &HEF: bomMarker(1) = &HBB: bomMarker(2) = &HBF
        m_FSO.FileWriteData m_DatabaseHandle, VarPtr(bomMarker(0)), 3
                
        'Request a database rewrite from ExifTool
        Dim cmdParams As pdString
        Set cmdParams = New pdString
        
        cmdParams.AppendLine "-listx"
        cmdParams.AppendLine "-lang"
        cmdParams.AppendLine "en"
        cmdParams.AppendLine "-f"
        cmdParams.AppendLine "-execute"
        
        'Send the data over to ExifTool.  The database will stream here asynchronously.
        If (Not m_Async Is Nothing) Then m_Async.SendData cmdParams.ToString()
        
        WriteTagDatabase = True
    
    End If

End Function

'Given a path to a valid metadata file, and a second path to a valid image file, use ExifTool to write the contents of
' the metadata file into the image file.
Public Function WriteMetadata(ByRef srcMetadataFile As String, ByRef dstImageFile As String, ByRef srcPDImage As pdImage, Optional ByVal forciblyAnonymize As Boolean = False, Optional ByRef originalMetadataParams As String = vbNullString) As Boolean
    
    'If ExifTool is not running, start it.  If it cannot be started, exit.
    If (Not m_IsExifToolRunning) Then
        If (Not StartExifTool()) Then
            Message "ExifTool could not be started.  Metadata unavailable for this session."
            WriteMetadata = False
            Exit Function
        End If
    End If
    
    'See if the output file format supports metadata.  If it doesn't, exit now.
    ' (Note that we return TRUE despite not writing any metadata - this lets the caller know that there were no errors.)
    Dim outputMetadataFormat As PD_METADATA_FORMAT
    outputMetadataFormat = ImageFormats.GetIdealMetadataFormatFromPDIF(srcPDImage.GetCurrentFileFormat)
    
    If (outputMetadataFormat = PDMF_NONE) Then
        PDDebug.LogAction "This file format does not support metadata.  Metadata processing skipped."
        If (LenB(srcMetadataFile) <> 0) Then Files.FileDeleteIfExists srcMetadataFile
        WriteMetadata = True
        Exit Function
    End If
    
    'If we are exporting a multipage image, different considerations are in place.  (For example, exporting a multipage
    ' TIFF requires us to not mess with the IFD### blocks - in JPEGs, these represent thumbnail images, but in a TIFF
    ' image they might be entire useful pages!)
    Dim saveIsMultipage As Boolean, saveIsMultipageTIFF As Boolean
    saveIsMultipage = srcPDImage.ImgStorage.GetEntry_Boolean("MultipageExportActive", False)
    If saveIsMultipage Then saveIsMultipageTIFF = (srcPDImage.GetCurrentFileFormat = PDIF_TIFF) Else saveIsMultipageTIFF = False
    
    'If an additional metadata parameter string was supplied, create a parser for it.  This may contain specialized
    ' processing instructions.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString originalMetadataParams
    
    'The preferred metadata format affects many of the requests sent to ExifTool.  Tag write requests are typically prefixed by
    ' the preferred tag group.  This string represents that group.
    Dim tagGroupPrefix As String
    Select Case outputMetadataFormat
    
        Case PDMF_EXIF
            tagGroupPrefix = "exif:"
        
        Case PDMF_XMP
            tagGroupPrefix = "xmp:"
            
        Case PDMF_IPTC
            tagGroupPrefix = "iptc:"
        
        Case Else
            tagGroupPrefix = vbNullString
            
    End Select
    
    'Start building a string of ExifTool parameters.  We will send these parameters to stdIn, but ExifTool expects them in
    ' ARGFILE format, e.g. each parameter on its own line.
    Dim cmdParams As pdString
    Set cmdParams = New pdString
    
    'Ignore minor errors and warnings
    cmdParams.AppendLine "-m"
    
    'Overwrite the original destination file (but only if the metadata was embedded successfully!)
    cmdParams.AppendLine "-overwrite_original"
    
    'To support Unicode filenames, explicitly request UTF-8-compatible parsing.
    cmdParams.AppendLine "-charset"
    cmdParams.AppendLine "filename=UTF8"
    
    'If a temporary file was given, mark it now.  It serves as our initial source of all metadata operations.
    Dim srcFileAvailable As Boolean
    srcFileAvailable = (LenB(srcMetadataFile) <> 0)
    If srcFileAvailable Then
    
        'Copy all tags.  It is important to do this first, because ExifTool applies operations in a left-to-right order -
        ' so we must start by copying all tags, then applying manual updates as necessary.
        cmdParams.AppendLine "-tagsfromfile"
        cmdParams.AppendLine srcMetadataFile
        cmdParams.AppendLine dstImageFile
    
    End If
    
    'Do *not* transfer over any thumbnail information (otherwise, this risks overwriting PD's existing thumbnail, if any!)
    ' Note that we ignore this when writing TIFFs, as they may be multipage, and there will be tons of IFD### blocks.
    If (Not saveIsMultipageTIFF) Then
        cmdParams.AppendLine "--IFD1:all"
    End If
    
    'Allow HTML entities (we need these for things like newlines)
    cmdParams.AppendLine "-E"
    
    'Define the same custom separator we used when initially reading the metadata
    cmdParams.AppendLine "-sep"
    cmdParams.AppendLine ";;;"
    
    'Next, we need to manually request the update of any tags that the user has manually modified via the metadata editor.
    Dim i As Long, tmpMetadata As PDMetadataItem, tmpEscapedValue As String
    If srcPDImage.ImgMetadata.HasMetadata Then
    
        For i = 0 To srcPDImage.ImgMetadata.GetMetadataCount - 1
            tmpMetadata = srcPDImage.ImgMetadata.GetMetadataEntry(i)
            
            If tmpMetadata.UserModifiedAllSessions Then
                cmdParams.Append "-" & tmpMetadata.TagGroupAndName & "="
                
                'Some tag-types require special escaping behavior (e.g. multiline comments, list-type values)
                If DoesTagValueRequireEscaping(tmpMetadata, tmpEscapedValue) Then
                    cmdParams.Append tmpEscapedValue
                Else
                    cmdParams.Append tmpMetadata.UserValueNew
                End If
                
                cmdParams.AppendLineBreak
            
            'For reasons I don't currently understand, unedited list-type tags may not be handled correctly by ExifTool.
            ' As such, specifically request their writing.
            Else
                If tmpMetadata.IsTagList Or tmpMetadata.DBF_IsBag Or tmpMetadata.DBF_IsList Or tmpMetadata.DBF_IsSequence Then
                    cmdParams.AppendLine "-" & tmpMetadata.TagGroupAndName & "=" & tmpMetadata.TagValueFriendly
                End If
            End If
                
            'Also, if the user has manually requested removal of a tag, mirror that request to ExifTool.
            '(NOTE: as of 7.0, this step is skipped.  To ensure removal, we forcibly strip problematic metadata entries
            ' from the source XML string (see pdMetadata.RetrieveModifiedXMLString() for details).  This spares us from
            ' needing to rely on ExifTool for the behavior, and similarly, if we request removal "just in case",
            ' ExifTool is likely to spew a whole bunch of warnings, which we don't want - but I've left this code here
            ' as an example, in case I need to someday reinstate it.
            'If tmpMetadata.TagMarkedForRemoval or tmpMetadata.TagMarkedForRemoval_Export Then
            '    cmdParams.AppendLine "-" & tmpMetadata.TagGroupAndName & "="
            'End If
            
        Next i
        
    End If
    
    'Regardless of the type of metadata copy we're performing, we need to alter or remove some tags because their
    ' original values are no longer relevant.
    If srcFileAvailable Then
    
        If (Not saveIsMultipageTIFF) Then
            cmdParams.AppendLine "-IFD2:ImageWidth="
            cmdParams.AppendLine "-IFD2:ImageHeight="
        End If
        
        cmdParams.AppendLine "--Padding"
    
        'Remove YCbCr subsampling data from the tags, as we may be using a different system than the previous save,
        ' and this information is not useful anyway - the JPEG header contains a copy of the subsampling data for
        ' the decoder, and that's sufficient!
        cmdParams.AppendLine "-YCbCrSubSampling="
        cmdParams.AppendLine "-IFD0:YCbCrSubSampling="
        
        'Remove YCbCrPositioning tags as well.  If no previous values are found, ExifTool will automatically
        ' repopulate these with a value matching the JPEG header.
        cmdParams.AppendLine "-YCbCrPositioning="
        
        'Photoshop embeds a bunch of problematic Photoshop-specific data, whose values may no longer be relevant
        cmdParams.AppendLine "-XMP-photoshop:ColorMode="
        cmdParams.AppendLine "-XMP-photoshop:ICCProfileName="
        
        'Note that some EXIF tags do not translate well to XMP.  ExifTool will copy these over anyway, but we want
        ' to manually remove them.
        cmdParams.AppendLine "-ExifIFD:BitsPerSample="
        cmdParams.AppendLine "-ExifIFD:ColorSpace="
        cmdParams.AppendLine "-ExifIFD:ComponentsConfiguration="
        cmdParams.AppendLine "-ExifIFD:CompressedBitsPerPixel="
        cmdParams.AppendLine "-ExifIFD:Compression="
        cmdParams.AppendLine "-xmp:BitsPerSample="
        cmdParams.AppendLine "-xmp:ColorSpace="
        cmdParams.AppendLine "-xmp:ComponentsConfiguration="
        cmdParams.AppendLine "-xmp:CompressedBitsPerPixel="
        cmdParams.AppendLine "-xmp:Compression="
        
    End If
    
    'If PD is embedding its own thumbnail, we will have saved a temporary thumbnail image file in a previous step.
    ' ExifTool needs to be notified of this file.
    Dim needToEmbedThumbnail As Boolean: needToEmbedThumbnail = False
    If cParams.GetBool("MetadataEmbedThumbnail", False) Then
        
        Dim tmpString As String
        tmpString = cParams.GetString("MetadataTempFilename")
        
        If (LenB(tmpString) <> 0) Then
            cmdParams.AppendLine "-ThumbnailImage<=" & tmpString
            needToEmbedThumbnail = True
        End If
        
    End If
    
    'Other software may have added tags related to an embedded thumbnail.  If PD is *not* embedding its own thumbnail,
    ' we want to forcibly remove any existing thumbnail information.
    If (Not needToEmbedThumbnail) And (Not saveIsMultipageTIFF) And srcFileAvailable Then
        cmdParams.AppendLine "-IFD1:Compression="
        cmdParams.AppendLine "-IFD1:all="
    End If
    
    'Now, we want to add a number of tags whose values should always be written, as they can be crucial to
    ' understanding the contents of the image.
    cmdParams.AppendLine "-" & tagGroupPrefix & "Orientation=Horizontal"
    cmdParams.AppendLine "-" & tagGroupPrefix & "XResolution=" & srcPDImage.GetDPI()
    cmdParams.AppendLine "-" & tagGroupPrefix & "YResolution=" & srcPDImage.GetDPI()
    cmdParams.AppendLine "-" & tagGroupPrefix & "ResolutionUnit=inches"
    
    'Various specs are unclear on the meaning of sRGB checks, and browser developers also have varying views
    ' on what an sRGB chunk means (see https://code.google.com/p/chromium/issues/detail?id=354883).
    ' Until these ambiguities are resolved, I've decided to skip sRGB flags for all image formats.
    'cmdParams.AppendLine "-" & tagGroupPrefix & "ColorSpace=sRGB"
    
    'Size tags are written to different areas based on the type of metadata being written.  JPEGs in particular
    ' have special rules; see the spec for details: http://www.cipa.jp/std/documents/e/DC-008-2012_E.pdf
    If (srcPDImage.GetCurrentFileFormat = PDIF_JPEG) Then
        cmdParams.AppendLine "-" & tagGroupPrefix & "ImageWidth="
        cmdParams.AppendLine "-" & tagGroupPrefix & "ImageHeight="
    Else
        cmdParams.AppendLine "-" & tagGroupPrefix & "ImageWidth=" & srcPDImage.Width
        cmdParams.AppendLine "-" & tagGroupPrefix & "ImageHeight=" & srcPDImage.Height
    End If
    
    If (outputMetadataFormat = PDMF_EXIF) Then
        cmdParams.AppendLine "-ExifIFD:ExifImageWidth=" & srcPDImage.Width
        cmdParams.AppendLine "-ExifIFD:ExifImageHeight=" & srcPDImage.Height
    ElseIf outputMetadataFormat = PDMF_XMP Then
        cmdParams.AppendLine "-xmp-exif:ExifImageWidth=" & srcPDImage.Width
        cmdParams.AppendLine "-xmp-exif:ExifImageHeight=" & srcPDImage.Height
    End If
    
    'JPEGs have the unique issue of needing their resolution values also updated in the JFIF header,
    ' so we make an additional request here for JPEGs specifically.
    If (srcPDImage.GetCurrentFileFormat = PDIF_JPEG) Then
        cmdParams.AppendLine "-JFIF:XResolution=" & srcPDImage.GetDPI()
        cmdParams.AppendLine "-JFIF:YResolution=" & srcPDImage.GetDPI()
        cmdParams.AppendLine "-JFIF:ResolutionUnit=inches"
    
    'Similarly, not all PNG editors fully support XMP, so let's write a pHYs tag too.
    ElseIf (srcPDImage.GetCurrentFileFormat = PDIF_PNG) Then
        cmdParams.AppendLine "-PNG:PixelsPerUnitX=" & Int((srcPDImage.GetDPI() / 2.54) * 100# + 0.5)
        cmdParams.AppendLine "-PNG:PixelsPerUnitY=" & Int((srcPDImage.GetDPI() / 2.54) * 100# + 0.5)
        cmdParams.AppendLine "-PNG:PixelUnits=meters"
    End If
    
    'If we are exporting a multipage TIFF object, add some per-page information now
    If saveIsMultipageTIFF And (Not forciblyAnonymize) Then AddMultipageData srcPDImage, cmdParams
    
    'If we were asked to remove GPS data, do so now
    If forciblyAnonymize And srcFileAvailable Then
        cmdParams.AppendLine "-gps:all="
      
    'If we were *not* asked to remove GPS data, and the source file contains GPS coordinates,
    ' we need to write them manually to prevent a loss of precision.
    Else
    
        If srcPDImage.ImgMetadata.HasGPSMetadata() Then
            Dim latString As String, lonString As String
            srcPDImage.ImgMetadata.FillLatitudeLongitude latString, lonString
            cmdParams.AppendLine "-GPS:GPSLatitude=" & latString
            cmdParams.AppendLine "-GPS:GPSLongitude=" & lonString
        End If
        
    End If
    
    'The incoming parameter "forciblyAnonymize" indicates the user wants privacy tags removed.
    ' If the user has NOT requested anonymization, list PD as the processing software.  (Note that this behavior can also be
    ' disabled from the Preferences dialog.)
    If (Not forciblyAnonymize) Then
        If UserPrefs.GetPref_Boolean("Saving", "MetadataListPD", True) Then cmdParams.AppendLine "-Software=" & GetPhotoDemonNameAndVersion()
    End If
    
    'ExifTool will always note itself as the XMP toolkit unless we specifically tell it not to; when "privacy mode" is active,
    ' do not list any toolkit at all.
    If forciblyAnonymize Then cmdParams.AppendLine "-XMPToolkit="
    
    'Finally, remove any other groups with potential privacy implications.
    ' (Normally, we want to remove individual tags instead of doing this, but ExifTool may attempt
    ' to be "helpful" and auto-convert data from other categories to these ones.  I've had recurring
    ' issues with this and iPhone photos, for example - so this adds another layer of failsafes.)
    If forciblyAnonymize Then
        cmdParams.AppendLine "-XMP-mwg-rs:all="
        cmdParams.AppendLine "-XMP-apdi:all="
        cmdParams.AppendLine "-XMP-depthData:all="
        cmdParams.AppendLine "-XMP-depthBlurEffect:all="
        cmdParams.AppendLine "-XMP-portraitLightingEffect:all="
        cmdParams.AppendLine "-XMP-portraitEffectsMatte:all="
        cmdParams.AppendLine "-XMP-semanticSegmentationMatte:all="
    End If
    
    'On some files, we prefer to use XMP over Exif.  This command instructs ExifTool to convert Exif tags to XMP tags where possible.
    If (outputMetadataFormat = PDMF_XMP) And srcFileAvailable Then
        
        'Convert all tags to XMP
        cmdParams.AppendLine "-xmp:all<all"
        
    End If
    
    'If the output format does not support Exif whatsoever, we can ask ExifTool to forcibly remove any remaining
    ' Exif tags. (This includes any tags it was unable to convert to XMP or IPTC format.)
    If (Not ImageFormats.IsExifAllowedForPDIF(srcPDImage.GetCurrentFileFormat)) And srcFileAvailable Then
        cmdParams.AppendLine "-exif:all="
    End If
    
    'If a temporary file was *not* given, supply our filename last
    If (Not srcFileAvailable) Then
        cmdParams.AppendLine dstImageFile
    End If
    
    'Finally, add the special command "-execute" which tells ExifTool to start operations
    cmdParams.AppendLine "-execute"
    
    'Activate verification mode.  This will asynchronously wait for the metadata to be written out to file, and when it
    ' has finished, it will erase our temp file.
    m_tmpMetadataFilePath = srcMetadataFile
    StartVerificationMode
    
    'Ask the user control to start processing this image's metadata.  It will handle things from here.
    If (Not m_Async Is Nothing) Then m_Async.SendData cmdParams.ToString()
    
    WriteMetadata = True
    
End Function

'Current, only TIFFs support specialty multipage metadata
Private Sub AddMultipageData(ByRef srcPDImage As pdImage, ByRef cmdParams As pdString)
    
    'Failsafe check; these should never be triggered
    If (srcPDImage Is Nothing) Or (cmdParams Is Nothing) Then Exit Sub
    
    Dim i As Long
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        If srcPDImage.GetLayerByIndex(i).GetLayerVisibility Then
            cmdParams.AppendLine "-IFD" & CStr(i) & ":PageName=" & srcPDImage.GetLayerByIndex(i).GetLayerName
        End If
    Next i
    
End Sub

Private Function DoesTagValueRequireEscaping(ByRef srcMetadata As PDMetadataItem, ByRef dstEscapedTag As String) As Boolean
    
    If (InStr(1, srcMetadata.UserValueNew, vbCrLf, vbBinaryCompare) <> 0) Then
        
        DoesTagValueRequireEscaping = True
        
        'List-type tags are escaped differently!
        If srcMetadata.DBF_IsBag Or srcMetadata.DBF_IsList Or srcMetadata.DBF_IsSequence Then
            dstEscapedTag = Replace$(srcMetadata.UserValueNew, vbCrLf, ";;;", , , vbBinaryCompare)
        Else
            dstEscapedTag = Replace$(srcMetadata.UserValueNew, vbCrLf, "&#xd;&#xa;", , , vbBinaryCompare)
        End If
        
    End If
    
End Function

'Start ExifTool.  We now use m_Async (a pdPipeAsync instance) to pass data to/from ExifTool.  This greatly
' reduces the overhead involved in repeatedly starting new ExifTool instances.  It also means that we can
' asynchronously start ExifTool early in the image load process, rather than waiting for the image to finish
' loading via FreeImage or GDI+.
Public Function StartExifTool() As Boolean
    
    'Start by creating a dedicated temp folder for ExifTool's required run-time extraction of Perl
    m_ExifToolDataFolder = UserPrefs.GetDataPath() & "PluginData\"
    Dim cFSO As pdFSO
    Set cFSO = New pdFSO
    
    If Files.PathExists(m_ExifToolDataFolder) Then
        
        'Ensure that the version of ExifTool hasn't changed since the last extraction.
        ' (If it has, errors can occur - this happened to PD after updating from the ExifTool 12.44 release
        ' to a later build.)
        '
        'To do this, we can quickly grab the main ExifTool Perl file and compare it against our current expected
        ' ExifTool version.  If the two do not match, kill the existing ExifTool folder and allow the .exe to
        ' re-create it anew.
        Const VERSION_REF_FILE As String = "inc\script\exiftool"
        
        Dim fullPathToRefFile As String
        fullPathToRefFile = m_ExifToolDataFolder & VERSION_REF_FILE
        
        If Files.FileExists(fullPathToRefFile) Then
            
            Const SAFE_BUFFER_SIZE As Long = 1024
            
            Dim cStream As pdStream
            Set cStream = New pdStream
            If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, fullPathToRefFile, SAFE_BUFFER_SIZE, optimizeAccess:=OptimizeSequentialAccess) Then
                
                'Grab just the first 1024 bytes.  Version should appear around char ~450, so this is a
                ' sufficiently safe sample size.
                Dim fileContents As String
                fileContents = cStream.ReadString_UTF8(SAFE_BUFFER_SIZE, False)
                
                'We no longer need access to the file
                cStream.StopStream True
                Set cStream = Nothing
                
                Const VERSION_STRING_PREFIX As String = "my $version = '"
                
                Dim posVersion As Long
                posVersion = InStr(1, fileContents, VERSION_STRING_PREFIX, vbBinaryCompare)
                If (posVersion > 0) Then
                    
                    'Increment past the initial single-quote
                    posVersion = posVersion + Len(VERSION_STRING_PREFIX)
                    
                    'Find the version string trailer single-quote
                    Dim posVersionEnd As Long
                    posVersionEnd = InStr(posVersion, fileContents, "'", vbBinaryCompare)
                    If (posVersionEnd > posVersion + 1) Then
                        
                        Dim efVersion As String
                        efVersion = Mid$(fileContents, posVersion, posVersionEnd - posVersion)
                        
                        'Compare the embedded version to the expected version for this PhotoDemon build.
                        Dim versionMatch As Boolean
                        versionMatch = Strings.StringsEqual(efVersion, PluginManager.ExpectedPluginVersion(CCP_ExifTool), True)
                        
                        If (Not versionMatch) Then
                            
                            PDDebug.LogAction "Warning: ExifTool data folder is out of date.  Preparing to rebuild..."
                            
                            'We now need to remove all files and subfolders from the /PluginData folder so that
                            ' ExifTool can rebuild it according to its current version.
                            Dim filesToDelete As pdStringStack
                            Set filesToDelete = New pdStringStack
                            If Files.RetrieveAllFiles(m_ExifToolDataFolder, filesToDelete, True, False) Then
                                
                                Dim numFilesDeleted As Long
                                numFilesDeleted = 0
                                
                                'In a perfect world, we could just remove files one-by-one and that would be that!
                                ' Unfortunately, the Perl packer ExifTool uses generates a lot of files with weird
                                ' permissions, so if we try to just blind-delete everything we will generate a ton
                                ' of permission errors.
                                '
                                'I have instead tried to blacklist file locations where I've encountered permission
                                ' errors while testing this feature, which should allows us to delete files that have
                                ' caused upgrade errors in past versions *without* engaging with any portable Perl
                                ' files that we don't have delete permissions on.  (Theoretically we could also check
                                ' permissions on each file before attempting to delete, but mass-checking arbitrary
                                ' file permissions risks upsetting some virus scanners, so I simply avoid interacting
                                ' with unnecessary files at all.)
                                Dim i As Long
                                For i = 0 To filesToDelete.GetNumOfStrings - 1
                                    
                                    Dim targetFile As String
                                    targetFile = filesToDelete.GetString(i)
                                    
                                    Dim okToDelete As Boolean
                                    okToDelete = Files.FileExists(targetFile)   'Failsafe only
                                    
                                    'This blacklist is hard-coded per manual file permission checks in the past.
                                    If okToDelete Then
                                        okToDelete = okToDelete And (InStr(1, targetFile, "\inc\lib\auto\", vbTextCompare) <= 0)
                                        If okToDelete Then okToDelete = okToDelete And (InStr(1, targetFile, "\inc\lib\unicore\", vbTextCompare) <= 0)
                                        If okToDelete Then okToDelete = okToDelete And (InStr(1, targetFile, "\inc\lib\Win32API\File\", vbTextCompare) <= 0)
                                        If okToDelete Then okToDelete = okToDelete And (InStr(1, targetFile, "\inc\lib\bytes_heavy.pl", vbTextCompare) <= 0)
                                        If okToDelete Then okToDelete = okToDelete And (InStr(1, targetFile, "\inc\lib\Config_git.pl", vbTextCompare) <= 0)
                                        If okToDelete Then okToDelete = okToDelete And (InStr(1, targetFile, "\inc\lib\Config_heavy.pl", vbTextCompare) <= 0)
                                    End If
                                    
                                    'Remove valid files (ExifTool will auto-regenerate them as necessary)
                                    If okToDelete Then
                                        If Files.FileDeleteIfExists(targetFile) Then
                                            numFilesDeleted = numFilesDeleted + 1
                                        Else
                                            PDDebug.LogAction "Failed to delete: " & filesToDelete.GetString(i)
                                        End If
                                    End If
                                    
                                Next i
                                
                                PDDebug.LogAction "ExifTool upgrade prep resulted in " & numFilesDeleted & " of " & filesToDelete.GetNumOfStrings & " files removed."
                                
                                'With all files gone, ExifTool can now rebuild itself accordingly
                                PDDebug.LogAction "ExifTool will now automatically rebuild its file cache..."
                                
                            End If
                            
                        '/if versions match, do nothing!  We are good to go.
                        Else
                            PDDebug.LogAction "(ExifTool version OK!)"
                        End If
                        
                    '/some kind of formatting problem; file structure has possibly changed since I last investigated?
                    Else
                        PDDebug.LogAction "WARNING: EXIFTOOL VERSION FORMAT UNEXPECTED"
                    End If
                    
                '/Version not found - file format has possibly changed?
                Else
                    PDDebug.LogAction "WARNING: NO EXIFTOOL VERSION IN PLUGINDATA"
                End If
                
            '/If we can't open the file, do not attempt to analyze (ExifTool may have it open already,
            ' in which case it's too late to do anything about it)
            End If
            
        '/no else required; if the target file doesn't exist, ExifTool will create it automatically
        End If
        
    'Create the folder anew
    Else
        Files.PathCreate m_ExifToolDataFolder
    End If
    
    'Set any other, related ExifTool paths now
    m_DatabasePath = m_ExifToolDataFolder & "exifToolDatabase.xml"
    
    'Next, set a local environment variable for Perl's temp folder, matching our temp folder above.  (If we do this prior
    ' to shelling ExifTool as a child process, ExifTool will automatically pick up the environment variable.)
    Dim envName As String, envValue As String
    envName = "PAR_GLOBAL_TEMP"
    envValue = m_ExifToolDataFolder
    SetEnvironmentVariableW StrPtr(envName), StrPtr(envValue)
    
    'Grab the ExifTool path, which we will shell and pipe in a moment
    Dim appLocation As String
    appLocation = PluginManager.GetPluginPath & "exiftool.exe"
    
    'Next, build a string of command-line parameters.  These will modify ExifTool's behavior to make it compatible with our code.
    Dim cmdParams As String
    
    'Tell ExifTool to stay open (e.g. do not exit after completing its operation), and to accept input from stdIn.
    ' (Note that exiftool.exe must be included as param [0], per C convention)
    cmdParams = cmdParams & "exiftool.exe -stay_open true -@ -"
    
    'Attempt to open ExifTool
    Set m_Async = New pdPipeAsync
    If m_Async.Run(appLocation, cmdParams) Then
        PDDebug.LogAction "ExifTool initiated successfully.  Ready to process metadata."
        m_IsExifToolRunning = True
        StartExifTool = True
    Else
        PDDebug.LogAction "WARNING! ExifTool Input/Output pipes could not be created."
        m_IsExifToolRunning = False
        StartExifTool = False
    End If

End Function

'Make sure to terminate ExifTool politely when the program closes.
Public Sub TerminateExifTool()

    'If for some reason, we still have an open handle to the ExifTool database, close it now.
    ' (This should never happen, barring a catastrophic crash or something similarly bad.)
    If (m_DatabaseHandle <> 0) Then
        If (Not m_FSO Is Nothing) Then m_FSO.FileCloseHandle m_DatabaseHandle
    End If

    If m_IsExifToolRunning Then
        
        If (Not m_Async Is Nothing) Then
            
            'Prepare a termination order
            Dim cmdParams As pdString
            Set cmdParams = New pdString
            
            cmdParams.AppendLine "-stay_open"
            cmdParams.AppendLine "False"
            cmdParams.AppendLine "-execute"
            
            'Ask ExifTool to terminate
            m_Async.SendData cmdParams.ToString()
        
            'Wait a little bit for ExifTool to receive the order and shut down on its own
            VBHacks.SleepAPI 500
            
            'ExifTool should be gone by now, but if it isn't, forcibly terminate it
            m_Async.TerminateChildProcess
            
        End If
        
        'As a failsafe, mark the plugin as no longer available
        PluginManager.SetPluginEnablement CCP_ExifTool, False
        
    End If

End Sub

'If an unclean shutdown is detected, use this function to try and terminate any ExifTool instances left over by the previous session.
' Many thanks to http://www.vbforums.com/showthread.php?321553-VB6-Killing-Processes&p=1898861#post1898861 for guidance on this task.
Public Sub KillStrandedExifToolInstances()
    
    'Prepare to purge all running ExifTool instances
    Const TH32CS_SNAPPROCESS As Long = 2&
    Const PROCESS_ALL_ACCESS As Long = 0&
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long, hSnapShot As Long, myProcess As Long
    Dim szExename As String
    Dim i As Long
    
    On Error GoTo CouldntKillExiftoolInstances
    
    'Prepare a generic process reference
    uProcess.dwSize = Len(uProcess)
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapShot, uProcess)
    
    Dim numProcessesTerminated As Long
    
    Dim procTarget As String
    procTarget = "exiftool.exe"
    
    'Iterate through all running processes, looking for ExifTool instances
    Do While (rProcessFound <> 0)
    
        'Retrieve the EXE name of this process
        i = InStr(1, uProcess.szExeFile, ChrW$(0), vbBinaryCompare)
        If (i > 1) Then
            
            szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
            
            'If the process name is "exiftool.exe", terminate it
            If Strings.StringsEqual(Right$(szExename, Len(procTarget)), procTarget, True) Then
                
                'Retrieve a handle to the ExifTool instance
                myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
                
                'Attempt to kill it
                If KillProcess(uProcess.th32ProcessID, 0) Then
                    PDDebug.LogAction "(Old ExifTool instance " & uProcess.th32ProcessID & " terminated successfully.)"
                    numProcessesTerminated = numProcessesTerminated + 1
                Else
                    PDDebug.LogAction "(Old ExifTool instance " & uProcess.th32ProcessID & " was not terminated; sorry!)"
                End If
                 
            End If
            
        End If
        
        'Find the next process, then continue
        rProcessFound = ProcessNext(hSnapShot, uProcess)
    
    Loop
    
    'Release our generic process snapshot, then exit
    If (hSnapShot <> 0) Then CloseHandle hSnapShot
    If (numProcessesTerminated > 0) Then PDDebug.LogAction CStr(numProcessesTerminated) & " old ExifTool instances were auto-terminated successfully."
    
    Exit Sub
    
CouldntKillExiftoolInstances:
    PDDebug.LogAction "Old ExifTool instances could not be auto-terminated due to errors.  Sorry!"
    
End Sub
 
'Terminate a process (referenced by its handle), and return success/failure
Function KillProcess(ByVal hProcessID As Long, Optional ByVal exitCode As Long) As Boolean

    Dim hToken As Long
    Dim hProcess As Long
    Dim tp As TOKEN_PRIVILEGES
     
    'Any number of things can cause the termination process to fail, unfortunately.  Check several known issues in advance.
    If (OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0) Then GoTo CleanUp
    
    Dim privName As String
    privName = "SeDebugPrivilege"
    If (LookupPrivilegeValueW(0, StrPtr(privName), tp.LuidUDT) = 0) Then GoTo CleanUp
    
    tp.PrivilegeCount = 1
    tp.Attributes = SE_PRIVILEGE_ENABLED
     
    If (AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0) Then GoTo CleanUp
     
    'Try to access the ExifTool process
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
    
    'Access granted!  Terminate the process
    If (hProcess <> 0) Then
        KillProcess = (TerminateProcess(hProcess, exitCode) <> 0)
        CloseHandle hProcess
    End If
    
    'Restore original privileges
    tp.Attributes = 0
    AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
     
CleanUp:
    
    'Free our privilege handle
    If (hToken <> 0) Then CloseHandle hToken
    
End Function

Public Sub StartNewDatabaseCache()
    ReDim m_GroupCache(0 To INIT_GROUP_CACHE_SIZE - 1) As ET_GROUP
    m_NumGroupsInCache = 0
End Sub

'Once the metadata database has been loaded, you can query it for additional tag information.  (This function will fail if
' 1) the database has not been loaded, or 2) the tag cannot be found.  (2) typically only happens if you haven't properly
' populated the dstMetadata object with data from an image file.)
Public Function FillTagFromDatabase(ByRef dstMetadata As PDMetadataItem) As Boolean
    
    'If this tag has already been processed during this session, ignore it.
    If dstMetadata.DB_TagHitDatabase Then
        FillTagFromDatabase = True
        Exit Function
    End If
    
    'Pulling a tag from the database is fairly complex.  We start by tracking down the relevant table in question.
    ' (Many different tag groups contain tags with the same name, so we must retrieve only the relevant entry.)
    
    'FYI - table segments (EXIF, IPTC, XMP, etc) look like this:
    ' <table name='Exif::Main' g0='EXIF' g1='IFD0' g2='Image'>
    '  <desc lang='en'>Exif</desc>
    ' ...tag values...
    ' </table>
    
    'We cache table positions as we locate them, so a separate function is used
    Dim tableStart As Long, tableEnd As Long
    If GetTagGroup(dstMetadata.TagTable, tableStart, tableEnd) Then
        
        'We now have a region of the string to search for this tag.  Tag database lines all follow a predictable pattern:
        '<tag id='1' name='InteropIndex' type='string' writable='true' flags='Unsafe' g1='InteropIFD'>
        ' <desc lang='en'>Interoperability Index</desc>
        ' ...potentially more info...
        '</tag>
        
        'Of that initial block, only the ID, NAME, TYPE, and WRITABLE attributes are guaranteed to exist.
        ' Extra FLAGS and GROUP IDs (e.g. "g1") are optional, as are some other obscure values.
        ' At present, we want to track down the start and end lines of the tag in question, so we can parse
        ' out any additional attributes.
        Dim srcTagSearch As String, tagStart As Long, tagEnd As Long
        srcTagSearch = "<tag id='" & dstMetadata.TagID & "' name='" & dstMetadata.tagName & "'"
        tagStart = InStr(tableStart, m_DatabaseString, srcTagSearch, vbBinaryCompare)
        If (tagStart <> 0) Then
            tagEnd = InStr(tagStart, m_DatabaseString, "</tag>", vbBinaryCompare)
            If (tagEnd <> 0) Then
            
                'Make sure the tag boundaries lie within the table boundaries
                If (tagStart > tableStart) And (tagEnd < tableEnd) Then
                
                    'We now know exactly where to parse out this tag's values.  Copy them into a dedicated string,
                    ' then pass that string off to a separate parse routine.  (The +6 is used to trap the closing
                    ' "</tag>" as well.)
                    Dim tagChunk As String
                    tagChunk = Mid$(m_DatabaseString, tagStart, (tagEnd - tagStart) + 6)
                    
                    ''During debugging, it may be helpful to cache the pure data returned from ExifTool; uncomment to allow this
                    'dstMetadata.TagDebugData = tagChunk
                    
                    If ParseTagDatabaseEntry(dstMetadata, tagChunk) Then
                        'To prevent future passes from hitting the database again, set the relevant shortcut flag
                        dstMetadata.DB_TagHitDatabase = True
                    Else
                        Debug.Print "WARNING!  ExifTool.ParseTagDatabaseEntry() failed on a tag: " & dstMetadata.TagTable & ">>" & dstMetadata.tagName
                    End If
                
                Else
                    Debug.Print "WARNING!  ExifTool.FillTagFromDatabase() found a tag, but it lies outside the required table boundaries: " & dstMetadata.TagTable & ">>" & dstMetadata.tagName
                End If
                
            Else
                Debug.Print "WARNING!  ExifTool.FillTagFromDatabase() failed to find a closing tag: " & dstMetadata.TagTable & ">>" & dstMetadata.tagName
            End If
        Else
            Debug.Print "WARNING!  ExifTool.FillTagFromDatabase() failed to find a matching table: " & dstMetadata.TagTable & ">>" & dstMetadata.tagName
        End If
        
    Else
        Debug.Print "WARNING!  ExifTool.GetTagGroup() failed to find a matching table: " & dstMetadata.TagTable
    End If

End Function

Private Function GetTagGroup(ByVal srcTableName As String, ByRef tableStart As Long, ByRef tableEnd As Long) As Boolean

    GetTagGroup = False
    
    'If caching is active, check there first
    If (m_NumGroupsInCache > 0) Then
        
        Dim i As Long
        For i = 0 To m_NumGroupsInCache - 1
            If Strings.StringsEqual(srcTableName, m_GroupCache(i).GroupName, False) Then
                tableStart = m_GroupCache(i).GroupStart
                tableEnd = m_GroupCache(i).GroupEnd
                GetTagGroup = True
                Exit For
            End If
        Next i
        
        If GetTagGroup Then Exit Function
    
    End If
    
    'If caching is inactive, or we couldn't find this group in the cache, we have to search the database manually
    Dim srcTableSearch As String
    srcTableSearch = "<table name='" & srcTableName & "'"
    tableStart = InStr(1, m_DatabaseString, srcTableSearch, vbBinaryCompare)
    
    'If the table wasn't located, there is literally nothing we can do!
    If (tableStart <> 0) Then
        
        'Find where this table ends.  The target tag must exist in this region, or it will be considered invalid.
        tableEnd = InStr(tableStart, m_DatabaseString, "</table>", vbBinaryCompare)
        If (tableEnd <> 0) Then
            
            'Add this group to our running cache
            With m_GroupCache(m_NumGroupsInCache)
                .GroupName = srcTableName
                .GroupStart = tableStart
                .GroupEnd = tableEnd
            End With
            
            m_NumGroupsInCache = m_NumGroupsInCache + 1
            If (m_NumGroupsInCache > UBound(m_GroupCache)) Then ReDim Preserve m_GroupCache(0 To m_NumGroupsInCache * 2 - 1) As ET_GROUP
            
            GetTagGroup = True
            
        Else
            Debug.Print "WARNING!  ExifTool.FillTagFromDatabase() failed to find a closing table tag: " & srcTableName
        End If
        
    Else
        Debug.Print "WARNING!  ExifTool.FillTagFromDatabase() failed to find a matching table: " & srcTableName
    End If
    
End Function

Private Function ParseTagDatabaseEntry(ByRef dstMetadata As PDMetadataItem, ByRef srcXML As String) As Boolean

    'This function assumes that the XML packet it receives is well-formed,
    ' and properly parsed out of the appropriate table in the primary ExifTool database.
    ' If these criteria are not met, all bets are off.
    
    'The first thing we want to do is break the XML into lines.  ExifTool spits out
    ' well-formed XML where each entry is placed on its own line, and this simplifies parsing.
    Dim xmlLines() As String
    xmlLines = Split(srcXML, vbCrLf, , vbBinaryCompare)
    
    If VBHacks.IsArrayInitialized(xmlLines) Then
        
        'To understand the next phase of the parsing process, let's look at the layout of two typical metadata tags:
        
        ' <tag id='1' name='InteropIndex' type='string' writable='true' flags='Unsafe' g1='InteropIFD'>
        '  <desc lang='en'>Interoperability Index</desc>
        '  <values>
        '   <key id='R03'>
        '    <val lang='en'>R03 - DCF option file (Adobe RGB)</val>
        '   </key>
        '   <key id='R98'>
        '    <val lang='en'>R98 - DCF basic file (sRGB)</val>
        '   </key>
        '   <key id='THM'>
        '    <val lang='en'>THM - DCF thumbnail file</val>
        '   </key>
        '  </values>
        ' </tag>
        
        ' <tag id='2' name='InteropVersion' type='undef' writable='true' flags='Mandatory,Unsafe' g1='InteropIFD'>
        '  <desc lang='en'>Interoperability Version</desc>
        ' </tag>
        
        'The first example represents tags where the range of possible values is discrete, and each entry is mapped to
        ' a specific, predetermined value.
        
        'The second example is a more freeform tag, where the user can theoretically place anything the way (within the
        ' constraints of the given type, obviously).
        
        'Regardless of which category a tag falls into, we can always parse the initial tag entry and the following
        ' description line using identical code.  (Also, note that we used the "id" and "name" values to locate this tag
        ' line in the first place, so those entries do not need to be parsed.  Instead, we want to get the "type" and
        ' "writable" values (which should always be present), and if they are available, any "count" or "flags" values.
        dstMetadata.DB_IsWritable = Strings.StringsEqual(GetXMLAttribute(xmlLines(0), "writable"), "true", False)
        dstMetadata.DB_DataType = GetXMLAttribute(xmlLines(0), "type")
        dstMetadata.DB_DataTypeStrict = GetStrictMDDatatype(dstMetadata.DB_DataType)
        
        'Some tags have a type like "byte x 4", which is common for things like RGBA definitions.  For purposes of
        ' presenting this value to the user, we must treat the value differently from a Long.  Tags like this can be
        ' identified by the presence of a "count" tag, which will always be >= 2.  ("0" and "1" would be redundant,
        ' as a single value is the assumed default.)
        Dim tmpString As String
        tmpString = GetXMLAttribute(xmlLines(0), "count")
        If (LenB(tmpString) <> 0) Then dstMetadata.DB_TypeCount = CLng(tmpString)
        
        'Flag retrieval is a bit convoluted, as flags are presented as a comma-delimited list.
        tmpString = GetXMLAttribute(xmlLines(0), "flags")
        If (LenB(tmpString) <> 0) Then
            dstMetadata.DBF_IsAvoid = (InStr(1, tmpString, "Avoid", vbBinaryCompare) <> 0)
            dstMetadata.DBF_IsBag = (InStr(1, tmpString, "Bag", vbBinaryCompare) <> 0)
            dstMetadata.DBF_IsBinary = (InStr(1, tmpString, "Binary", vbBinaryCompare) <> 0)
            dstMetadata.DBF_IsFlattened = (InStr(1, tmpString, "Flattened", vbBinaryCompare) <> 0)
            dstMetadata.DBF_IsList = (InStr(1, tmpString, "List", vbBinaryCompare) <> 0)
            dstMetadata.DBF_IsMandatory = (InStr(1, tmpString, "Mandatory", vbBinaryCompare) <> 0)
            dstMetadata.DBF_IsPermanent = (InStr(1, tmpString, "Permanent", vbBinaryCompare) <> 0)
            dstMetadata.DBF_IsProtected = (InStr(1, tmpString, "Protected", vbBinaryCompare) <> 0)
            dstMetadata.DBF_IsSequence = (InStr(1, tmpString, "Sequence", vbBinaryCompare) <> 0)
            dstMetadata.DBF_IsUnknown = (InStr(1, tmpString, "Unknown", vbBinaryCompare) <> 0)
            dstMetadata.DBF_IsUnsafe = (InStr(1, tmpString, "Unsafe", vbBinaryCompare) <> 0)
        End If
        
        'The second line is always a description
        If (UBound(xmlLines) >= 1) Then dstMetadata.DB_Description = ExifTool.PARSE_UnescapeXML(GetXMLValue_SingleLine(xmlLines(1)))
        
        'The third line will be one of two things:
        ' 1) A closing tag (literally, "</tag>")
        ' 2) A <values> tag, which indicates that one or more hard-coded key/value pairs follow
        If (UBound(xmlLines) >= 2) Then
            If (InStr(1, xmlLines(2), "</tag>", vbBinaryCompare) = 0) Then
                If (InStr(1, xmlLines(2), "<values>", vbBinaryCompare) <> 0) Then
                    
                    dstMetadata.DB_HardcodedList = True
                    Dim numOfKeys As Long: numOfKeys = 0
                    Dim curLine As Long: curLine = 3
                    
                    'Values follow a nice, predictable pattern:
                    '<key id='-1'>
                    ' <val lang='en'>n/a</val>
                    '</key>
                    '<key id='1'>
                    ' <val lang='en'>Canon EF 50mm f/1.8</val>
                    '</key>
                    '<key id='2'>
                    ' <val lang='en'>Canon EF 28mm f/2.8</val>
                    '</key>
                    '...many more lines...
                    
                    'Unfortunately, there's no way to know how many values are present, short of manually parsing until
                    ' we hit something that *isn't* a value.
                    
                    'To simplify this step, each metadata entry supplies two pdStringStacks.  One is used for ID values,
                    ' the other for values.  (Someday, we may support translations here, because you can coax them out of
                    ' ExifTool, but that's another project for another day.)
                    Set dstMetadata.DB_StackIDs = New pdStringStack
                    Set dstMetadata.DB_StackValues = New pdStringStack
                    
                    Do
                        dstMetadata.DB_StackIDs.AddString GetXMLAttribute(xmlLines(curLine), "id")
                        dstMetadata.DB_StackValues.AddString ExifTool.PARSE_UnescapeXML(GetXMLValue_SingleLine(xmlLines(curLine + 1)))
                        
                        numOfKeys = numOfKeys + 1
                        curLine = 3 + numOfKeys * 3
                        If curLine > UBound(xmlLines) Then Exit Do
                    Loop While (InStr(1, xmlLines(curLine), "<key id=", vbBinaryCompare) <> 0)
                    
                    dstMetadata.DB_HardcodedListSize = numOfKeys
                    
                Else
                    Debug.Print "WARNING!  Strange 3rd line found inside ExifTool.ParseTagDatabaseEntry: " & xmlLines(2)
                End If
            
            End If
        End If
        
        'Before exiting, we apply some manual changes to certain tags, to align with the way PD processes metadata.
        
        'First, we want to mark some tags as "Protected", even though they are technically not (e.g. JPEG Orientation, which we
        ' overwrite automatically in order to produce valid files).  ExifTool lets the user mess with these tags, but we don't.
        If Strings.StringsEqual(dstMetadata.TagNameFriendly, "Orientation", False) Then dstMetadata.DBF_IsProtected = True
        
        ParseTagDatabaseEntry = True
    
    Else
        ParseTagDatabaseEntry = False
    End If

End Function

Private Function GetXMLAttribute(ByRef srcXML As String, ByRef atrbName As String) As String
    
    GetXMLAttribute = vbNullString
    
    Dim strSearch As String
    strSearch = " " & atrbName & "='"
    
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, srcXML, strSearch, vbBinaryCompare)
    If (startPos <> 0) Then
        startPos = startPos + Len(strSearch)
        endPos = InStr(startPos + 1, srcXML, "'", vbBinaryCompare)
        If (endPos <> 0) Then GetXMLAttribute = Mid$(srcXML, startPos, endPos - startPos)
    End If
    
End Function

Private Function GetXMLValue_SingleLine(ByRef srcLine As String) As String
    Dim sPos As Long, ePos As Long
    sPos = InStr(1, srcLine, ">", vbBinaryCompare)
    If (sPos > 0) Then
        ePos = InStrRev(srcLine, "<", , vbBinaryCompare) - 1
        If (ePos > 0) And (ePos > sPos) Then
            GetXMLValue_SingleLine = Mid$(srcLine, sPos + 1, ePos - sPos)
        End If
    End If
End Function

Private Function GetStrictMDDatatype(ByRef textRepresentation As String) As PD_Metadata_Datatype
    
    If Strings.StringsEqual(textRepresentation, "int8s") Then
        GetStrictMDDatatype = MD_int8s
    ElseIf Strings.StringsEqual(textRepresentation, "int8u") Then
        GetStrictMDDatatype = MD_int8u
    ElseIf Strings.StringsEqual(textRepresentation, "int16s") Then
        GetStrictMDDatatype = MD_int16s
    ElseIf Strings.StringsEqual(textRepresentation, "int16u") Or Strings.StringsEqual(textRepresentation, "int16uRev") Then
        GetStrictMDDatatype = MD_int16u
    ElseIf Strings.StringsEqual(textRepresentation, "int32s") Then
        GetStrictMDDatatype = MD_int32s
    ElseIf Strings.StringsEqual(textRepresentation, "int32u") Then
        GetStrictMDDatatype = MD_int32u
    ElseIf Strings.StringsEqual(textRepresentation, "int64s") Then
        GetStrictMDDatatype = MD_int64s
    ElseIf Strings.StringsEqual(textRepresentation, "int64u") Then
        GetStrictMDDatatype = MD_int64u
    ElseIf Strings.StringsEqual(textRepresentation, "rational32s") Then
        GetStrictMDDatatype = MD_rational32s
    ElseIf Strings.StringsEqual(textRepresentation, "rational32u") Then
        GetStrictMDDatatype = MD_rational32u
    ElseIf Strings.StringsEqual(textRepresentation, "rational64s") Then
        GetStrictMDDatatype = MD_rational64s
    ElseIf Strings.StringsEqual(textRepresentation, "rational64u") Then
        GetStrictMDDatatype = MD_rational64u
    ElseIf Strings.StringsEqual(textRepresentation, "fixed16s") Then
        GetStrictMDDatatype = MD_fixed16s
    ElseIf Strings.StringsEqual(textRepresentation, "fixed16u") Then
        GetStrictMDDatatype = MD_fixed16u
    ElseIf Strings.StringsEqual(textRepresentation, "fixed32s") Then
        GetStrictMDDatatype = MD_fixed32s
    ElseIf Strings.StringsEqual(textRepresentation, "fixed32u") Then
        GetStrictMDDatatype = MD_fixed32u
    ElseIf Strings.StringsEqual(textRepresentation, "float") Then
        GetStrictMDDatatype = MD_float
    ElseIf Strings.StringsEqual(textRepresentation, "double") Then
        GetStrictMDDatatype = MD_double
    ElseIf Strings.StringsEqual(textRepresentation, "extended") Then
        GetStrictMDDatatype = MD_extended
    ElseIf Strings.StringsEqual(textRepresentation, "ifd") Then
        GetStrictMDDatatype = MD_ifd
    ElseIf Strings.StringsEqual(textRepresentation, "ifd64") Then
        GetStrictMDDatatype = MD_ifd64
    ElseIf Strings.StringsEqual(textRepresentation, "string") Then
        GetStrictMDDatatype = MD_string
    ElseIf Strings.StringsEqual(textRepresentation, "undef") Or Strings.StringsEqual(textRepresentation, "?") Then
        GetStrictMDDatatype = MD_undef
    ElseIf Strings.StringsEqual(textRepresentation, "binary") Then
        GetStrictMDDatatype = MD_binary
    
    'This group of data types are XMP-specific.  They are always stored as strings, but said strings may need to
    ' observe particular formatting to work.
    ElseIf Strings.StringsEqual(textRepresentation, "integer") Then
        GetStrictMDDatatype = MD_integerstring
    ElseIf Strings.StringsEqual(textRepresentation, "real") Then
        GetStrictMDDatatype = MD_floatstring
    ElseIf Strings.StringsEqual(textRepresentation, "rational") Then
        GetStrictMDDatatype = MD_rationalstring
    ElseIf Strings.StringsEqual(textRepresentation, "date") Then
        GetStrictMDDatatype = MD_datestring
    ElseIf Strings.StringsEqual(textRepresentation, "boolean") Then
        GetStrictMDDatatype = MD_booleanstring
    ElseIf Strings.StringsEqual(textRepresentation, "lang-alt") Then
        GetStrictMDDatatype = MD_string
    'All XMP information is stored as character strings. The Writable column specifies the information format:
    ' integer is a string of digits (possibly beginning with a '+' or '-'),
    ' real is a floating point number
    ' rational is two integer strings separated by a '/' character
    ' date is a date/time string in the format "YYYY:MM:DD HH:MM:SS[+/-HH:MM]"
    ' boolean is either "True" or "False", and lang-alt is a list of string alternatives in different languages.
    'Individual languages for lang-alt tags are accessed by suffixing the tag name with a '-', followed by an RFC 3066 language code (ie. "XMP:Title-fr", or "Rights-en-US"). A lang-alt tag with no language code accesses the "x-default" language, but causes other languages to be deleted when writing. The "x-default" language code may be specified when writing a new value to write only the default language, but note that all languages are still deleted if "x-default" tag is deleted. When reading, "x-default" is not specified.
    
    'Data types past this point do not appear in the official ExifTool documentation, but they have been observed in
    ' the database.  This list may not be all-inclusive.
    ElseIf Strings.StringsEqual(textRepresentation, "digits") Then
        GetStrictMDDatatype = MD_integerstring
    ElseIf (InStr(1, textRepresentation, "string", vbBinaryCompare) <> 0) Then
        GetStrictMDDatatype = MD_string
    ElseIf (InStr(1, textRepresentation, "str", vbBinaryCompare) <> 0) Then
        GetStrictMDDatatype = MD_string
    
    Else
        Debug.Print "WARNING!  ExifTool.GetStrictMDDataType could not resolve this type: " & textRepresentation
    End If
    
End Function

'ExifTool produces properly escaped chars for the five predefined XML entities (<, >, &, ', ").
' Numeric character references are *not* unescaped, but support could be added in the future.
Public Function PARSE_UnescapeXML(ByRef srcString As String) As String
    PARSE_UnescapeXML = srcString
    If InStr(1, PARSE_UnescapeXML, "&lt;", vbBinaryCompare) Then PARSE_UnescapeXML = Replace$(PARSE_UnescapeXML, "&lt;", "<")
    If InStr(1, PARSE_UnescapeXML, "&gt;", vbBinaryCompare) Then PARSE_UnescapeXML = Replace$(PARSE_UnescapeXML, "&gt;", ">")
    If InStr(1, PARSE_UnescapeXML, "&amp;", vbBinaryCompare) Then PARSE_UnescapeXML = Replace$(PARSE_UnescapeXML, "&amp;", "&")
    If InStr(1, PARSE_UnescapeXML, "&apos;", vbBinaryCompare) Then PARSE_UnescapeXML = Replace$(PARSE_UnescapeXML, "&apos;", "'")
    If InStr(1, PARSE_UnescapeXML, "&quot;", vbBinaryCompare) Then PARSE_UnescapeXML = Replace$(PARSE_UnescapeXML, "&quot;", """")
End Function

'Someday, it would be nice for the caller to have some control over this list.  For example, they may want to strip GPS,
' but preserve Copyright data.  That's not possible with a hard-coded implementation.
'
'For now, however, you can call this function to see if PD will remove the tag if the Anonymization option is checked.
Public Function DoesTagHavePrivacyConcerns(ByRef srcTag As PDMetadataItem) As Boolean
    
    Dim potentialConcern As Boolean
    potentialConcern = False
    
    'Non-writable categories can be skipped in advance, as their tags cannot physically be written by ExifTool.
    Dim sCategoryName As String
    sCategoryName = UCase$(srcTag.TagGroupFriendly)
    
    Dim groupSkippable As Boolean: groupSkippable = False
    If Strings.StringsEqual(sCategoryName, "SYSTEM", True) Then groupSkippable = True
    If Strings.StringsEqual(sCategoryName, "FILE", True) Then groupSkippable = True
    If Strings.StringsEqual(sCategoryName, "ICC PROFILE", True) Then groupSkippable = True
    
    'Technically, we should be able to get away with not checking inferred tags (called "Composite" by ExifTool), per this link:
    ' https://exiftool.org/TagNames/Composite.html
    '
    'But because this is a sensitive topic, we err on the side of caution and check tags in the inferred group.
    'If StrComp(sCategoryName, "INFERRED", vbBinaryCompare) = 0 Then groupSkippable = True
    
    'Only proceed with further checks if this category is a potentially writable one
    If (Not groupSkippable) Then
        
        'First, we check technical tag names for known problematic text
        Dim sMetadataName As String
        sMetadataName = LCase$(Trim$(srcTag.tagName))
        
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "firmware", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "about", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "artist", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "author", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "baby", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "by-line", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "city", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "comment", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "copyright", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "country", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "creator", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "date", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "description", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "digest", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "documentid", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "gps", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "history", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "host", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "info", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "instanceid", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "lens", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "location", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "make", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "model", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "name", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "program", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "region", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "rights", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "serial", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "software", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "subject", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "time", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "title", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "tool", vbBinaryCompare) > 0)
        potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "version", vbBinaryCompare) > 0)
        
        'Next, check actual tag values for known problematic text
        If (Not potentialConcern) Then
            Dim sMetadataValue As String
            sMetadataValue = LCase$(Trim$(srcTag.TagValue))
            potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "xmp.iid", vbBinaryCompare) > 0)
            potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "xmp.did", vbBinaryCompare) > 0)
            potentialConcern = potentialConcern Or (InStr(1, sMetadataName, "uuid", vbBinaryCompare) > 0)
        End If
        
    End If
    
    DoesTagHavePrivacyConcerns = potentialConcern
                    
End Function

'Some tags should never be written.  ExifTool is smart about handling such tags, but it will throw warnings about them,
' which clutters up our debug tracking.  Use this function to forcibly suppress some known metadata entries that should
' never be manually written.
Public Function ShouldTagNeverBeWritten(ByRef srcTag As PDMetadataItem) As Boolean

    Dim potentialConcern As Boolean: potentialConcern = False
    potentialConcern = Strings.StringsEqual(srcTag.tagName, "ExifToolVersion", True) Or potentialConcern
    potentialConcern = Strings.StringsEqual(srcTag.tagName, "JFIFVersion", True) Or potentialConcern
    potentialConcern = Strings.StringsEqual(srcTag.TagGroup, "File", True) Or potentialConcern
    potentialConcern = Strings.StringsEqual(srcTag.TagGroup, "System", True) Or potentialConcern
    potentialConcern = Strings.StringsEqual(srcTag.TagGroup, "XMP-depthData", True) Or potentialConcern
    ShouldTagNeverBeWritten = potentialConcern
                    
End Function

'The pdMetadata class now handles this function internally.  This copy exists solely for legacy
' PDI files that used the old format.
Public Sub RecoverSerializedTag_Legacy(ByRef srcString As String, ByRef dstMetadata As PDMetadataItem)
    
    If (LenB(srcString) <> 0) Then
        
        If (m_ParseXML Is Nothing) Then Set m_ParseXML = New pdSerialize
        m_ParseXML.SetParamString srcString
        
        'Basically, this is just a long-ass process of retrieving all tag properties from their specific XML tags.
        With dstMetadata
            .TagGroupAndName = m_ParseXML.GetString("PDMD_TagGroupAndName", , True)
            .TagGroup = m_ParseXML.GetString("PDMD_TagGroup", , True)
            .TagGroupFriendly = m_ParseXML.GetString("PDMD_TagGroupFriendly", , True)
            .tagName = m_ParseXML.GetString("PDMD_TagName", , True)
            .TagNameFriendly = m_ParseXML.GetString("PDMD_TagNameFriendly", , True)
            .TagTable = m_ParseXML.GetString("PDMD_TagTable", , True)
            .TagID = m_ParseXML.GetString("PDMD_TagID", , True)
            .TagValueFriendly = m_ParseXML.GetString("PDMD_TagValueFriendly", , True)
            .TagValue = m_ParseXML.GetString("PDMD_TagValue", , True)
            .HasIndex = m_ParseXML.GetBool("PDMD_HasIndex")
            .IsTagList = m_ParseXML.GetBool("PDMD_IsTagList")
            .IsTagBinary = m_ParseXML.GetBool("PDMD_IsTagBinary")
            .TagBase64Value = m_ParseXML.GetString("PDMD_TagBase64", , True)
            .WasBinaryExtracted = m_ParseXML.GetBool("PDMD_WasBinaryExtracted")
            .InternalUseOnly = m_ParseXML.GetBool("PDMD_InternalUseOnly")
            .TagIndexInternal = m_ParseXML.GetLong("PDMD_TagIndexInternal")
            .TagBase64Value = m_ParseXML.GetString("PDMD_TagBase64Value", , True)
            .TagMarkedForRemoval = m_ParseXML.GetBool("PDMD_TagMarkedForRemoval")
            .UserModifiedThisSession = m_ParseXML.GetBool("PDMD_UserModifiedThisSession")
            .UserModifiedAllSessions = m_ParseXML.GetBool("PDMD_UserModifiedAllSessions")
            .UserValueNew = m_ParseXML.GetString("PDMD_UserValueNew", , True)
            .UserIDNew = m_ParseXML.GetString("PDMD_UserIDNew", , True)
            .DB_TagHitDatabase = m_ParseXML.GetBool("PDMD_DBTagHitDatabase")
            .DB_IsWritable = m_ParseXML.GetBool("PDMD_DBISWritable")
            .DB_TypeCount = m_ParseXML.GetLong("PDMD_DBTypeCount")
            .DB_DataType = m_ParseXML.GetString("PDMD_DBDataType", , True)
            .DB_DataTypeStrict = m_ParseXML.GetLong("PDMD_DBDataTypeStrict")
            .DBF_IsAvoid = m_ParseXML.GetBool("PDMD_DBFIsAvoid")
            .DBF_IsBag = m_ParseXML.GetBool("PDMD_DBFIsBag")
            .DBF_IsBinary = m_ParseXML.GetBool("PDMD_DBFIsBinary")
            .DBF_IsFlattened = m_ParseXML.GetBool("PDMD_DBFIsFlattened")
            .DBF_IsList = m_ParseXML.GetBool("PDMD_DBFIsList")
            .DBF_IsMandatory = m_ParseXML.GetBool("PDMD_DBFIsMandatory")
            .DBF_IsPermanent = m_ParseXML.GetBool("PDMD_DBFIsPermanent")
            .DBF_IsProtected = m_ParseXML.GetBool("PDMD_DBFIsProtected")
            .DBF_IsSequence = m_ParseXML.GetBool("PDMD_DBFIsSequence")
            .DBF_IsUnknown = m_ParseXML.GetBool("PDMD_DBFIsUnknown")
            .DBF_IsUnsafe = m_ParseXML.GetBool("PDMD_DBFIsUnsafe")
            .DB_Description = m_ParseXML.GetString("PDMD_DBDescription", , True)
            .DB_HardcodedList = m_ParseXML.GetBool("PDMD_DBHardCodedList")
            .DB_HardcodedListSize = m_ParseXML.GetLong("PDMD_DBHardCodedListSize")
            If .DB_HardcodedList And (.DB_HardcodedListSize > 0) Then
                Set .DB_StackIDs = New pdStringStack
                Set .DB_StackValues = New pdStringStack
                .DB_StackIDs.RecreateStackFromSerializedString m_ParseXML.GetString("PDMD_StackIDs")
                .DB_StackValues.RecreateStackFromSerializedString m_ParseXML.GetString("PDMD_StackValues")
            End If
            
            'During debugging, it may be helpful to cache the pure data returned from ExifTool; uncomment to allow this
            '.TagDebugData = m_ParseXML.GetString("PDMD_TagDebugData", , True)
            
        End With
        
    End If
    
End Sub

Public Sub RemoveAllMetadata(ByRef srcImage As pdImage)
    If (Not srcImage Is Nothing) Then srcImage.ImgMetadata.Reset
End Sub
