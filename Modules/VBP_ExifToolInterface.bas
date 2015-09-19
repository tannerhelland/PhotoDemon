Attribute VB_Name = "Plugin_ExifTool_Interface"
'***************************************************************************
'ExifTool Plugin Interface
'Copyright 2013-2015 by Tanner Helland
'Created: 24/May/13
'Last updated: 22/October/14
'Last update: many technical improvements to metadata writing.  Formats that support only XMP or Exif will now have
'              as many tags as humanly possible converted to the relevant format.  Unconverted tags will be ignored.
'              When writing new tags, the preferred metadata format for a given image format will be preferentially
'              used (e.g. XMP for PNG files, Exif for JPEGs, etc).
'
'Module for handling all ExifTool interfacing.  This module is pointless without the accompanying ExifTool plugin,
' which can be found in the App/PhotoDemon/Plugins subdirectory as "exiftool.exe".  The ExifTool plugin is
' available by default in all versions of PhotoDemon after and including 6.0.
'
'ExifTool is a comprehensive image metadata handler written by Phil Harvey.  No DLL or VB-compatible library is
' available, so PhotoDemon relies on the stock Windows ExifTool executable file for all interfacing.  You can read
' more about ExifTool at its homepage:
'
'http://www.sno.phy.queensu.ca/~phil/exiftool/
'
'As of version 6.1 build 499, all ExifTool interaction is piped across stdin/out.  This includes sending requests to
' ExifTool, and checking results for success/failure.  All of PhotoDemon's metadata code has been rewritten to take
' advantage of the asynchronous abilities this provides.
'
'In the first draft of this code, I used a sample VB module as a reference courtesy of Michael Wandel:
'
'http://owl.phy.queensu.ca/~phil/exiftool/modExiftool_101.zip
'
'...as well as a modified piping function, derived from code originally written by Joacim Andersson:
'
'http://www.vbforums.com/showthread.php?364219-Classic-VB-How-do-I-shell-a-command-line-program-and-capture-the-output
'
'Those code modules are no longer relevant to the current implementation, but I thought it worthwhile to mention them
' in case others find them useful.
'
'This project was originally designed against v9.37 of ExifTool (14 Sep '13).  While I do test newer versions, it's
' impossible to test all metadata possibilities, so problems may arise if used with other versions of the software.
' Additional documentation regarding the use of ExifTool can be found in the official ExifTool package, downloadable
' from http://www.sno.phy.queensu.ca/~phil/exiftool/
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'A number of API functions are required to pipe stdout
Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100
Private Const SW_NORMAL = 1
Private Const SW_HIDE = 0
Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2

'Potential error codes (not used at present, but could be added in the future)
'Private Const ERROR_BROKEN_PIPE = 109

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long

'A great deal of extra code is required for finding ExifTool instances left by previous unsafe shutdowns, and silently terminating them.
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
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
 
Type PROCESSENTRY32
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


'This type is what PhotoDemon uses internally for storing and displaying metadata.  Its complexity is a testament to the nightmare
' that is metadata management.
Public Type mdItem
    FullGroupAndName As String
    Group As String
    SubGroup As String
    Name As String
    Description As String
    Value As String
    ActualValue As String
    Base64Value As String
    UserModified As Boolean
    isValueBinary As Boolean
    isValueList As Boolean
    isActualValueBinary As Boolean
    isActualValueList As Boolean
    isValueMultiLine As Boolean
    isValueBase64 As Boolean
    markedForRemoval As Boolean
End Type

'Once ExifTool has been run at least once, this will be set to TRUE.  If TRUE, this means that the shellPipeMain user control
' on FormMain is active and connected to ExifTool, and can be used to send/receive input and output.
Private isExifToolRunning As Boolean

'Because ExifTool parses metadata asynchronously, we will gather its output as it comes.  This string will hold whatever
' XML data ExifTool has returned so far.
Private curMetadataString As String

'While capture mode is active (e.g. while retrieving metadata information we care about) this will be set to TRUE.
Private captureModeActive As Boolean

'While verification mode is active (e.g. we only care about ExifTool succeeding or failing), this will be set to TRUE.
Private verificationModeActive As Boolean
Private verificationString As String

'Once in its lifetime, ExifTool will be asked to write its full tag database to an XML file (~1.5 mb worth of data).  This
' will normally happen the first time PD is run.  While that special mode is running, this variable will be set to TRUE.
Private databaseModeActive As Boolean
Private databaseString As String

'As of 6.6 alpha, technical metadata reports can now be generated for a given file.  While this mode is active, we do not
' want to immediately delete the report; use this boolean to check for that particular state.
Private technicalReportModeActive As Boolean
Private technicalReportSrcImage As String

'Prior to writing out a file's metadata, we must cache the information we want written in a temp file.  (ExifTool requires
' a source file when writing metadata out to file; the alternative is to manually request the writing of each tag in turn,
' but if we do this, we lose many built-in utilities like automatically removing duplicate tags, and reassigning invalid
' tags to preferred categories.)  Because writing out metadata is asynchronous, we have to wait for ExifTool to finish
' before deleting the temp file, so we keep a copy of the file's path here.  The stopVerificationMode (which is
' automatically triggered by the newMetadataReceived function as necessary) will remove the file at this location.
Private tmpMetadataFilePath As String

'If multiple images are loaded simultaneously, we have to do some tricky handling to parse out their individual bits.  As such, we store
' the ID of the last image metadata request we received; only when this ID is returned successfully do we consider metadata processing
' "complete", and return TRUE for isMetadataFinished.
Private m_LastRequestID As Long

Public Function isDatabaseModeActive() As Boolean
    isDatabaseModeActive = databaseModeActive
End Function

'The FormMain.ShellPipeMain user control will asynchronously trigger this function whenever it receives new metadata
' from ExifTool.
Public Sub newMetadataReceived(ByVal newMetadata As String)
    
    If captureModeActive Then
        curMetadataString = curMetadataString & newMetadata
    
    'During verification mode, we must also check to see if verification has finished
    ElseIf verificationModeActive Then
        verificationString = verificationString & newMetadata
        If isMetadataFinished() Then stopVerificationMode
        
    'During database mode, check for a finish state, then write the retrieved database out to file!
    ElseIf databaseModeActive Then
        databaseString = databaseString & newMetadata
        If isMetadataFinished() Then writeMetadataDatabaseToFile
    End If
    
End Sub

Private Sub writeMetadataDatabaseToFile()

    Dim mdDatabasePath As String
    mdDatabasePath = g_PluginPath & "ExifToolDatabase.xml"
        
    'Replace the {ready} text supplied by ExifTool itself, which will be at the end of the metadata database
    If Len(databaseString) <> 0 Then databaseString = Replace$(databaseString, "{ready}", "")
    
    'Write our XML string out to file
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    cFile.SaveStringToTextFile databaseString, mdDatabasePath
        
    'Reset the database mode trackers, so the database doesn't accidentally get rebuilt again
    databaseString = ""
    databaseModeActive = False
    
End Sub

'When we only care about ExifTool succeeding or failing, use this function to enter "verification mode", which simply checks
' to see if ExifTool has finished its previous request.
Private Sub startVerificationMode()
    verificationModeActive = True
    verificationString = ""
End Sub

'When verification mode ends (as triggered by an automatic check in newMetadataReceived), we must also remove the temporary
' file we created that held the data being exported via ExifTool.
Private Sub stopVerificationMode()
    
    verificationModeActive = False
    
    'All file interactions are handled through pdFSO, PhotoDemon's Unicode-compatible file system object
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If Not technicalReportModeActive Then
    
        verificationString = ""
        
        'Verification mode is a bit different.  We need to erase our temporary metadata file if it exists; then we can exit.
        cFile.KillFile tmpMetadataFilePath
        
    Else
    
        'Replace the {ready} text supplied by ExifTool itself, which will be at the end of the metadata report
        If Len(verificationString) <> 0 Then verificationString = Replace$(verificationString, "{ready}", "")
        
        'Write the completed technical report out to a temp file
        Dim tmpFilename As String
        tmpFilename = g_UserPreferences.GetTempPath & "MetadataReport_" & getFilenameWithoutExtension(technicalReportSrcImage) & ".html"
        
        cFile.SaveStringToTextFile verificationString, tmpFilename  ', True, False
        
        'Shell the default HTML viewer for the user
        verificationString = ""
        OpenURL tmpFilename
        
        technicalReportSrcImage = ""
        technicalReportModeActive = False
        
    End If
    
End Sub

'When ExifTool has completed work on metadata, it will send "{ready}" to stdout.  We check for the presence of "{ready}" in
' the metadata string to see if ExifTool is done.
Public Function isMetadataFinished() As Boolean
    
    'If ExifTool is not available, or if it failed to start, simple return TRUE which will allow any waiting code
    ' to continue.
    If Not isExifToolRunning Then
        isMetadataFinished = True
        Exit Function
    End If
    
    'I don't know much about asynchronous string handling in VB, but just to be safe, make a copy of the current
    ' metadata string (to avoid collisions?).
    Dim tmpMetadata As String
    
    If captureModeActive Then
        tmpMetadata = curMetadataString
    ElseIf verificationModeActive Then
        tmpMetadata = verificationString
    ElseIf databaseModeActive Then
        tmpMetadata = databaseString
    Else
        isMetadataFinished = True
        Exit Function
    End If
    
    'If there is no temporary metadata string, exit now
    If Len(tmpMetadata) = 0 Then Exit Function
    
    'Different verification modes require different checks for completion.
    If captureModeActive Then
        
        If InStr(1, tmpMetadata, "{ready" & m_LastRequestID & "}", vbBinaryCompare) > 0 Then
            
            'Terminate the relevant mode
            captureModeActive = False
            isMetadataFinished = True
            
        Else
            isMetadataFinished = False
        End If
        
    Else
    
        If InStr(1, tmpMetadata, "{ready}", vbBinaryCompare) > 0 Then
            
            'Terminate the relevant mode
            If verificationModeActive Then verificationModeActive = False
            If databaseModeActive Then databaseModeActive = False
            
            isMetadataFinished = True
            
        Else
            isMetadataFinished = False
        End If
    
    End If
    
End Function

'When metadata is ready (as determined by a call to isMetadataFinished), it can be retrieved via this function
Public Function retrieveMetadataString() As String
    
    If Len(curMetadataString) <> 0 Then
    
        'Because we request metadata in XML format, ExifTool escapes disallowed XML characters.  Convert those back
        ' to standard characters before returning the retrieved metadata.
        If InStr(1, curMetadataString, "&amp;") > 0 Then curMetadataString = Replace(curMetadataString, "&amp;", "&")
        If InStr(1, curMetadataString, "&#39;") > 0 Then curMetadataString = Replace(curMetadataString, "&#39;", "'")
        If InStr(1, curMetadataString, "&quot;") > 0 Then curMetadataString = Replace(curMetadataString, "&quot;", """")
        If InStr(1, curMetadataString, "&gt;") > 0 Then curMetadataString = Replace(curMetadataString, "&gt;", ">")
        If InStr(1, curMetadataString, "&lt;") > 0 Then curMetadataString = Replace(curMetadataString, "&lt;", "<")
        
    End If
        
    'Return the processed string, then erase our copy of it
    retrieveMetadataString = curMetadataString
    curMetadataString = ""
    
End Function

Public Function retrieveUntouchedMetadataString() As String
    retrieveUntouchedMetadataString = curMetadataString
End Function

'Is ExifTool available as a plugin?
Public Function isExifToolAvailable() As Boolean
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(g_PluginPath & "exiftool.exe") Then isExifToolAvailable = True Else isExifToolAvailable = False
    
End Function

'Retrieve the ExifTool version.
Public Function getExifToolVersion() As String

    If Not isExifToolAvailable Then
        getExifToolVersion = ""
        Exit Function
    Else
        
        Dim exifPath As String
        exifPath = g_PluginPath & "exiftool.exe"
        
        Dim outputString As String
        If ShellExecuteCapture(exifPath, "exiftool.exe -ver", outputString) Then
        
            'The output string will be a simple version number, e.g. "XX.YY", and it will be terminated by a vbCrLf character.
            ' Remove vbCrLf now.
            outputString = Trim$(outputString)
            If InStr(outputString, vbCrLf) <> 0 Then outputString = Replace(outputString, vbCrLf, "")
            getExifToolVersion = outputString
            
        Else
            getExifToolVersion = ""
        End If
        
    End If
    
End Function

'Start an ExifTool instance (if one isn't already active), and have it process an image file.  Because we now run ExifTool
' asynchronously, this should be done early in the image load process.
Public Function startMetadataProcessing(ByVal srcFile As String, ByVal srcFormat As Long, ByVal targetImageID As Long) As Boolean

    'If ExifTool is not running, start it.  If it cannot be started, exit.
    If Not isExifToolRunning Then
        If Not startExifTool() Then
            Message "ExifTool could not be started.  Metadata unavailable for this session."
            startMetadataProcessing = False
            Exit Function
        End If
    End If
    
    'Notify the program that stdout capture has begun
    captureModeActive = True
    
    'Erase any previous metadata caches.
    ' NOTE! Upon implementing PD's new asynchronous metadata retrieval mechanism, we don't want to erase the master metadata string,
    '        as its construction may lag behind the rest of the image load process.  When a full metadata string is retrieved,
    '        the retrieveMetadataString() function will handle clearing for us.
    'curMetadataString = ""
    
    'Many ExifTool options are delimited by quotation marks (").  Because VB has the worst character escaping scheme ever conceived, I use
    ' a variable to hold the ASCII equivalent of a quotation mark.  This makes things slightly more readable.
    Dim Quotes As String
    Quotes = Chr(34)
    
    'Start building a string of ExifTool parameters.  We will send these parameters to stdIn, but ExifTool expects them in
    ' ARGFILE format, e.g. each parameter on its own line.
    Dim cmdParams As String
    cmdParams = ""
    
    'Ignore minor errors and warnings
    cmdParams = cmdParams & "-m" & vbCrLf
    
    'Output long-format data
    cmdParams = cmdParams & "-l" & vbCrLf
        
    'Request a custom separator for list-type values
    cmdParams = cmdParams & "-sep" & vbCrLf & ";" & vbCrLf
        
    'If a translation is active, request descriptions in the current language
    If g_Language.translationActive Then
        cmdParams = cmdParams & "-lang" & vbCrLf & g_Language.getCurrentLanguage() & vbCrLf
    End If
    
    'TEST! On JPEGs, request a digest as well
    'If srcFormat = FIF_JPEG Then cmdParams = cmdParams & "-jpegdigest" & vbCrLf
    
    'Request that binary data be processed.  We have no use for this data within PD, but when it comes time to write
    ' our metadata back out to file, we may want to have a copy of it.
    'cmdParams = cmdParams & "-b" & vbCrLf
    
    'Historically, we needed to explicitly set a charset; this shouldn't be necessary with current versions (as UTF-8 is
    ' automatically supported), but if desired, specific metadata types can be coerced into requested character sets.
    'cmdParams = cmdParams & "-charset" & vbCrLf & "UTF8" & vbCrLf
    
    'To support Unicode filenames, explicitly request UTF-8-compatible parsing.
    cmdParams = cmdParams & "-charset" & vbCrLf & "filename=UTF8" & vbCrLf
    
    'Actually, we now forcibly request IPTC data as UTF-8.  IPTC supports charset markers, but in my experience, these are
    ' rarely used.  ExifTool will default to the current code page for conversion if we don't specify otherwise, so UTF-8
    ' is preferable here.
    cmdParams = cmdParams & "-charset" & vbCrLf & "iptc=UTF8" & vbCrLf
            
    'Requesting binary data also means preview and thumbnail images will be processed.  We DEFINITELY don't want these,
    ' so deny them specifically.
    'cmdParams = cmdParams & "-x" & vbCrLf & "PreviewImage" & vbCrLf
    'cmdParams = cmdParams & "-x" & vbCrLf & "ThumbnailImage" & vbCrLf
    'cmdParams = cmdParams & "-x" & vbCrLf & "PhotoshopThumbnail" & vbCrLf
    
    'Output XML data (a lot more complex, but the only way to retrieve descriptions and names simultaneously)
    cmdParams = cmdParams & "-X" & vbCrLf
    
    'Add the image path
    cmdParams = cmdParams & srcFile & vbCrLf
    
    'Finally, add the special command "-execute" which tells ExifTool to start operations
    cmdParams = cmdParams & "-execute" & targetImageID & vbCrLf
    
    'Note this request ID as being the last one we received; only when this ID is returned by ExifTool will we actually consider our
    ' work complete.
    m_LastRequestID = targetImageID
    
    'DEBUG ONLY! Display the param list we have assembled.
    'Debug.Print cmdParams
    
    'Ask the user control to start processing this image's metadata.  It will handle things from here.
    FormMain.shellPipeMain.SendData cmdParams
    
End Function

'ExifTool has a lot of great facilities for analyzing image metadata.  Technical users in particular might want to take advantage
' of ExifTool's "htmldump" facility, which provides a detailed hex report of all metadata in a file.  This function can be used
' to generate such a report, but note that it only works for images that exist on disk (obviously).
Public Function createTechnicalMetadataReport(ByRef srcImage As pdImage) As Boolean

    'Start by checking for an existing copy of the XML database.  If it already exists, no need to recreate it.
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(srcImage.locationOnDisk) Then
    
        Dim cmdParams As String
        cmdParams = ""
        
        'Add the htmldump command
        cmdParams = cmdParams & "-htmldump" & vbCrLf
        
        'Add -u, which will also report unknown tags
        cmdParams = cmdParams & "-u" & vbCrLf
                
        'To support Unicode filenames, explicitly request UTF-8-compatible parsing.
        cmdParams = cmdParams & "-charset" & vbCrLf & "filename=UTF8" & vbCrLf
                
        'Add the source image to the list
        technicalReportSrcImage = srcImage.locationOnDisk
        cmdParams = cmdParams & srcImage.locationOnDisk & vbCrLf
        
        'Finally, add the special command "-execute" which tells ExifTool to start operations
        cmdParams = cmdParams & "-execute" & vbCrLf
        
        'Activate verification mode.  This will asynchronously wait for the metadata to be written out to file, and when it
        ' has finished, it will erase our temp file.
        technicalReportModeActive = True
        startVerificationMode
        
        'Ask the user control to start processing this image's metadata.  It will handle things from here.
        FormMain.shellPipeMain.SendData cmdParams
        
        createTechnicalMetadataReport = True
    
    Else
    
        createTechnicalMetadataReport = False
    
    End If

End Function

'If the user wants to edit an image's metadata, we need to know which tags are writeable and which are not.  Also, it's helpful to
' know things like each tag's datatype (to verify output before it's passed along to ExifTool).  If ExifTool is successfully initialized
' at program startup, this function will be called, and its job is to populate ExifTool's tag database.
Public Function writeTagDatabase() As Boolean

    'Start by checking for an existing copy of the XML database.  If it already exists, no need to recreate it.
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(g_PluginPath & "exifToolDatabase.xml") Then
    
        'Database already exists - no need to recreate it!
        writeTagDatabase = True
    
    Else
    
        'Database wasn't found.  Generate a new copy now.
        
        'Start metadata database retrieval mode
        databaseModeActive = True
        databaseString = ""
        
        'Request a database rewrite from ExifTool
        Dim cmdParams As String
        cmdParams = ""
        
        cmdParams = cmdParams & "-listx" & vbCrLf
        cmdParams = cmdParams & "-s" & vbCrLf
        cmdParams = cmdParams & "-execute" & vbCrLf
        
        FormMain.shellPipeMain.SendData cmdParams
        
        writeTagDatabase = True
    
    End If

End Function

'Given a path to a valid metadata file, and a second path to a valid image file, use ExifTool to write the contents of
' the metadata file into the image file.
Public Function writeMetadata(ByVal srcMetadataFile As String, ByVal dstImageFile As String, ByRef srcPDImage As pdImage, Optional ByVal removeGPS As Boolean = False) As Boolean
    
    'If ExifTool is not running, start it.  If it cannot be started, exit.
    If Not isExifToolRunning Then
        If Not startExifTool() Then
            Message "ExifTool could not be started.  Metadata unavailable for this session."
            writeMetadata = False
            Exit Function
        End If
    End If
    
    'See if the output file format supports metadata.  If it doesn't, exit now.
    ' (Note that we return TRUE despite not writing any metadata - this lets the caller know that there were no errors.)
    Dim outputMetadataFormat As PD_METADATA_FORMAT
    outputMetadataFormat = g_ImageFormats.getIdealMetadataFormatFromFIF(srcPDImage.currentFileFormat)
    
    If outputMetadataFormat = PDMF_NONE Then
        Message "This file format does not support metadata.  Metadata processing skipped."
        writeMetadata = True
        Exit Function
    End If
    
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
            tagGroupPrefix = ""
            
    End Select
    
    'Many ExifTool options are delimited by quotation marks (").  Because VB has the worst character escaping scheme ever conceived, I use
    ' a variable to hold the ASCII equivalent of a quotation mark.  This makes things slightly more readable.
    Dim Quotes As String
    Quotes = Chr(34)
    
    'Start building a string of ExifTool parameters.  We will send these parameters to stdIn, but ExifTool expects them in
    ' ARGFILE format, e.g. each parameter on its own line.
    Dim cmdParams As String
    cmdParams = ""
    
    'Ignore minor errors and warnings
    cmdParams = cmdParams & "-m" & vbCrLf
        
    'Overwrite the original destination file, but only if the metadata was embedded succesfully
    cmdParams = cmdParams & "-overwrite_original" & vbCrLf
    
    'To support Unicode filenames, explicitly request UTF-8-compatible parsing.
    cmdParams = cmdParams & "-charset" & vbCrLf & "filename=UTF8" & vbCrLf
    
    'Copy all tags.  It is important to do this first, because ExifTool applies operations in a left-to-right order - so we must
    ' start by copying all tags, then applying manual updates as necessary.
    cmdParams = cmdParams & "-tagsfromfile" & vbCrLf & srcMetadataFile & vbCrLf
    cmdParams = cmdParams & dstImageFile & vbCrLf
    
    'On some files, we prefer to use XMP over Exif.  This command instructs ExifTool to convert Exif tags to XMP tags where possible.
    If outputMetadataFormat = PDMF_XMP Then
        cmdParams = cmdParams & "-xmp:all<all" & vbCrLf
    End If
    
    'Regardless of the type of metadata copy we're performing, we need to alter or remove some tags because their
    ' original values are no longer relevant.
    cmdParams = cmdParams & "--IFD2:ImageWidth" & vbCrLf & "--IFD2:ImageHeight" & vbCrLf
    cmdParams = cmdParams & "--Padding" & vbCrLf
            
    'Remove YCbCr subsampling data from the tags, as we may be using a different system than the previous save, and this information
    ' is not useful anyway - the JPEG header contains a copy of the subsampling data for the decoder, and that's sufficient!
    cmdParams = cmdParams & "--YCbCrSubSampling" & vbCrLf
    cmdParams = cmdParams & "--IFD0:YCbCrSubSampling" & vbCrLf
    
    'Remove YCbCrPositioning tags as well.  If no previous values are found, ExifTool will automatically repopulate these with
    ' the right value according to the JPEG header.
    cmdParams = cmdParams & "--YCbCrPositioning" & vbCrLf
    
    'Other software may have added Exif tags for an embedded thumbnail.  PD obeys the JFIF spec and writes the thumbnail into the
    ' JFIF header, so we don't want those extra Exif tags included.
    cmdParams = cmdParams & "-IFD1:all=" & vbCrLf
    
    'Now, we want to add a number of tags whose values should always be written, as they can be crucial to understanding the
    ' contents of the image.
    cmdParams = cmdParams & "-" & tagGroupPrefix & "Orientation=Horizontal" & vbCrLf
    cmdParams = cmdParams & "-" & tagGroupPrefix & "XResolution=" & srcPDImage.getDPI() & vbCrLf
    cmdParams = cmdParams & "-" & tagGroupPrefix & "YResolution=" & srcPDImage.getDPI() & vbCrLf
    cmdParams = cmdParams & "-" & tagGroupPrefix & "ResolutionUnit=inches" & vbCrLf
    
    'Various specs are unclear on the meaning of sRGB checks, and browser developers also have varying views on what an sRGB chunk means
    ' (see https://code.google.com/p/chromium/issues/detail?id=354883)
    ' Until such point as I can resolve these ambiguities, sRGB flags are now skipped for all formats.
    'cmdParams = cmdParams & "-" & tagGroupPrefix & "ColorSpace=sRGB" & vbCrLf
    
    'Size tags are written to different areas based on the type of metadata being written.  JPEGs require special rules; see the spec
    ' for details: http://www.cipa.jp/std/documents/e/DC-008-2012_E.pdf
    If srcPDImage.currentFileFormat = FIF_JPEG Then
        cmdParams = cmdParams & "--" & tagGroupPrefix & "ImageWidth" & vbCrLf
        cmdParams = cmdParams & "--" & tagGroupPrefix & "ImageHeight" & vbCrLf
    Else
        cmdParams = cmdParams & "-" & tagGroupPrefix & "ImageWidth=" & srcPDImage.Width & vbCrLf
        cmdParams = cmdParams & "-" & tagGroupPrefix & "ImageHeight=" & srcPDImage.Height & vbCrLf
    End If
    
    If outputMetadataFormat = PDMF_EXIF Then
        cmdParams = cmdParams & "-ExifIFD:ExifImageWidth=" & srcPDImage.Width & vbCrLf
        cmdParams = cmdParams & "-ExifIFD:ExifImageHeight=" & srcPDImage.Height & vbCrLf
    ElseIf outputMetadataFormat = PDMF_XMP Then
        cmdParams = cmdParams & "-xmp-exif:ExifImageWidth=" & srcPDImage.Width & vbCrLf
        cmdParams = cmdParams & "-xmp-exif:ExifImageHeight=" & srcPDImage.Height & vbCrLf
    End If
    
    
    
    'JPEGs have the unique issue of needing their resolution values also updated in the JFIF header, so we make
    ' an additional request here for JPEGs specifically.
    If srcPDImage.currentFileFormat = FIF_JPEG Then
        cmdParams = cmdParams & "-JFIF:XResolution=" & srcPDImage.getDPI() & vbCrLf
        cmdParams = cmdParams & "-JFIF:YResolution=" & srcPDImage.getDPI() & vbCrLf
        cmdParams = cmdParams & "-JFIF:ResolutionUnit=inches" & vbCrLf
    End If
    
    'If we were asked to remove GPS data, do so now
    If removeGPS Then cmdParams = cmdParams & "-gps:all=" & vbCrLf
    
    'GPS removal indicates the user wants privacy tags removed; if the user has NOT requested removal of these, list PD as
    ' the processing software.
    If Not removeGPS Then cmdParams = cmdParams & "-Software=" & getPhotoDemonNameAndVersion() & vbCrLf
    
    'ExifTool will always note itself as the XMP toolkit unless we specifically tell it not to; when "privacy mode" is active,
    ' do not list any toolkit at all.
    If removeGPS Then cmdParams = cmdParams & "-XMPToolkit=" & vbCrLf
        
    'If the output format does not support Exif whatsoever, we can ask ExifTool to forcibly remove any remaining Exif tags.
    ' (This includes any tags it was unable to convert to XMP or IPTC format.)
    If Not g_ImageFormats.isExifAllowedForFIF(srcPDImage.currentFileFormat) Then
        cmdParams = cmdParams & "-exif:all=" & vbCrLf
    End If
        
    'Finally, add the special command "-execute" which tells ExifTool to start operations
    cmdParams = cmdParams & "-execute" & vbCrLf
    
    'Activate verification mode.  This will asynchronously wait for the metadata to be written out to file, and when it
    ' has finished, it will erase our temp file.
    tmpMetadataFilePath = srcMetadataFile
    startVerificationMode
    
    'Ask the user control to start processing this image's metadata.  It will handle things from here.
    FormMain.shellPipeMain.SendData cmdParams
    
    writeMetadata = True
    
End Function

'Start ExifTool.  We now use FormMain.shellPipeMain (a user control of type ShellPipe) to pass data to/from ExifTool.  This greatly
' reduces the overhead involved in repeatedly starting new ExifTool instances.  It also means that we can asynchronously start
' ExifTool early in the load process, rather than waiting for an image to be loaded.
Public Function startExifTool() As Boolean
    
    'Many ExifTool options are delimited by quotation marks (").  Because VB has the worst character escaping scheme ever conceived, I use
    ' a variable to hold the ASCII equivalent of a quotation mark.  This makes things slightly more readable.
    Dim Quotes As String
    Quotes = Chr(34)
    
    'Grab the ExifTool path, which we will shell and pipe in a moment
    Dim appLocation As String
    appLocation = g_PluginPath & "exiftool.exe"
    
    'Next, build a string of command-line parameters.  These will modify ExifTool's behavior to make it compatible with our code.
    Dim cmdParams As String
    
    'Tell ExifTool to stay open (e.g. do not exit after completing its operation), and to accept input from stdIn.
    ' (Note that exiftool.exe must be included as param [0], per C convention)
    cmdParams = cmdParams & "exiftool.exe -charset filename=UTF8 -stay_open true -@ -"
    
    'Attempt to open ExifTool
    Dim returnVal As SP_RESULTS
    returnVal = FormMain.shellPipeMain.Run(appLocation, cmdParams)
    returnVal = SP_CREATEPIPEFAILED
    returnVal = SP_CREATEPROCFAILED
    returnVal = SP_SUCCESS
    Select Case returnVal
    
        Case SP_SUCCESS
            Message "ExifTool initiated successfully.  Ready to process metadata."
            isExifToolRunning = True
            startExifTool = True
            
        Case SP_CREATEPIPEFAILED
            Message "WARNING! ExifTool Input/Output pipes could not be created."
            isExifToolRunning = False
            startExifTool = False
            
        Case SP_CREATEPROCFAILED
            Message "WARNING! ExifTool.exe could not be started."
            isExifToolRunning = False
            startExifTool = False
    
    End Select

End Function

'Make sure to terminate ExifTool politely when the program closes.
Public Sub terminateExifTool()

    If isExifToolRunning Then

        'Prepare a termination order for ExifTool
        Dim cmdParams As String
        cmdParams = ""
        
        cmdParams = cmdParams & "-stay_open" & vbCrLf
        cmdParams = cmdParams & "False" & vbCrLf
        cmdParams = cmdParams & "-execute" & vbCrLf
        
        'Submit the order
        FormMain.shellPipeMain.SendData cmdParams
        
        'Wait a little bit for ExifTool to receive the order and shut down on its own
        Sleep 500
        
        'Close our own pipe handles and exit
        FormMain.shellPipeMain.FinishChild
        
        'As a failsafe, mark the plugin as no longer available
        g_ExifToolEnabled = False
        
    End If

End Sub

'Capture output from the requested command-line executable and return it as a string.  At present, this is only used to check the
' ExifTool version number, which is only done on-demand if the Plugin Manager is loaded.
' ALSO NOTE: This function is a heavily modified version of code originally written by Joacim Andersson. A download link to his
' original version is available at the top of this module.
Public Function ShellExecuteCapture(ByVal sApplicationPath As String, sCommandLineParams As String, ByRef dstString As String, Optional bShowWindow As Boolean = False) As Boolean

    Dim hPipeRead As Long, hPipeWrite As Long
    Dim hCurProcess As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim si As STARTUPINFO
    Dim PI As PROCESS_INFORMATION
    Dim baOutput() As Byte
    Dim sNewOutput As String
    Dim lBytesRead As Long
    
    'This pipe buffer size is effectively arbitrary, but I haven't had any problems with 1024
    Const BUFSIZE = 1024

    dstString = ""
    
    ReDim baOutput(BUFSIZE - 1) As Byte

    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1
    End With

    If CreatePipe(hPipeRead, hPipeWrite, sa, BUFSIZE) = 0 Then
        ShellExecuteCapture = False
        Message "Failed to start plugin service (couldn't create pipe)."
        Exit Function
    End If

    hCurProcess = GetCurrentProcess()

    'Replace the inheritable read handle with a non-inheritable one. (MSDN suggestion)
    DuplicateHandle hCurProcess, hPipeRead, hCurProcess, hPipeRead, 0&, 0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE

    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .hStdOutput = hPipeWrite
        
        'NOTE: calling functions typically request that the shelled window be shown in the IDE but not in the compiled .exe
        If bShowWindow Then .wShowWindow = SW_NORMAL Else .wShowWindow = SW_HIDE
        
    End With
    
    If CreateProcess(sApplicationPath, sCommandLineParams, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, vbNullString, si, PI) Then

        'Close the thread handle, as we have no use for it
        CloseHandle PI.hThread

        'Also close the pipe's write handle. This is important, because ReadFile will not return until all write handles
        ' are closed or the buffer is full.
        CloseHandle hPipeWrite
        hPipeWrite = 0
        
        Do
            
            If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then Exit Do
            
            sNewOutput = StrConv(baOutput, vbUnicode)
            dstString = dstString & Left$(sNewOutput, lBytesRead)
            
        Loop

        CloseHandle PI.hProcess
        CloseHandle hCurProcess
    Else
        ShellExecuteCapture = False
        Message "Failed to start plugin service (couldn't create process: %1).", Err.LastDllError
        Exit Function
    End If
    
    CloseHandle hPipeRead
    If hPipeWrite Then CloseHandle hPipeWrite
    If hCurProcess Then CloseHandle hCurProcess
    
    ShellExecuteCapture = True
    
End Function

'If an unclean shutdown is detected, use this function to try and terminate any ExifTool instances left over by the previous session.
' Many thanks to http://www.vbforums.com/showthread.php?321553-VB6-Killing-Processes&p=1898861#post1898861 for guidance on this task.
Public Sub killStrandedExifToolInstances()
    
    'Prepare to purge all running ExifTool instances
    Const TH32CS_SNAPPROCESS As Long = 2&
    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long, hSnapShot As Long, myProcess As Long
    Dim szExename As String
    Dim i As Long
    
    On Local Error GoTo CouldntKillExiftoolInstances
    
    'Prepare a generic process reference
    uProcess.dwSize = Len(uProcess)
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapShot, uProcess)
    
    'Iterate through all running processes, looking for ExifTool instances
    Do While rProcessFound
    
        'Retrieve the EXE name of this process
        i = InStr(1, uProcess.szExeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
        
        'If the process name is "exiftool.exe", terminate it
        If Right$(szExename, Len("exiftool.exe")) = "exiftool.exe" Then
            
            'Retrieve a handle to the ExifTool instance
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            
            'Attempt to kill it
            If KillProcess(uProcess.th32ProcessID, 0) Then
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "(Old ExifTool instance " & uProcess.th32ProcessID & " terminated successfully.)"
                #End If
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "(Old ExifTool instance " & uProcess.th32ProcessID & " was not terminated; sorry!)"
                #End If
            End If
             
        End If
        
        'Find the next process, then continue
        rProcessFound = ProcessNext(hSnapShot, uProcess)
    
    Loop
    
    'Release our generic process snapshot, then exit
    CloseHandle hSnapShot
    #If DEBUGMODE = 1 Then
        Debug.Print "All old ExifTool instances were auto-terminated successfully."
    #End If
    
    Exit Sub
    
CouldntKillExiftoolInstances:
    
    #If DEBUGMODE = 1 Then
        Debug.Print "Old ExifTool instances could not be auto-terminated.  Sorry!"
    #End If
    
End Sub
 
'Terminate a process (referenced by its handle), and return success/failure
Function KillProcess(ByVal hProcessID As Long, Optional ByVal exitCode As Long) As Boolean

    Dim hToken As Long
    Dim hProcess As Long
    Dim tp As TOKEN_PRIVILEGES
     
    'Any number of things can cause the termination process to fail, unfortunately.  Check several known issues in advance.
    If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then GoTo CleanUp
    If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then GoTo CleanUp
    
    tp.PrivilegeCount = 1
    tp.Attributes = SE_PRIVILEGE_ENABLED
     
    If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0 Then GoTo CleanUp
     
    'Try to access the ExifTool process
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
    
    'Access granted!  Terminate the process
    If hProcess Then
     
        KillProcess = (TerminateProcess(hProcess, exitCode) <> 0)
        CloseHandle hProcess
        
    End If
    
    'Restore original privileges
    tp.Attributes = 0
    AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
     
CleanUp:
    
    'Free our privilege handle
    If hToken Then CloseHandle hToken
    
End Function
