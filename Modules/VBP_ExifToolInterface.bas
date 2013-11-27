Attribute VB_Name = "Plugin_ExifTool_Interface"
'***************************************************************************
'ExifTool Plugin Interface
'Copyright ©2012-2013 by Tanner Helland
'Created: 24/May/13
'Last updated: 26/November/13
'Last update: huge metadata rewrite.  Loading/saving is now asynchronous, meaning it does not interfere with normal
'              program operation!  Also the option to "preserve all metadata, regardless of relevance" has been removed.
'              The reasons for this are many, but basically there is *no physical way* to preserve metadata exactly, so
'              it is misleading to claim to do so.  PD is better off not providing options like that, so I have reworked
'              the metadata handler to only operate on relevant data.  Irrelevant or invalid data is now forcibly removed.
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
'Those code modules are no longer relevant to the current implementation, but I thought it worthwhile to mention them.
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

'This type is what PhotoDemon uses internally for storing and displaying metadata
Public Type mdItem
    FullGroupAndName As String
    Group As String
    SubGroup As String
    Name As String
    Description As String
    Value As String
    ActualValue As String
    isValueBinary As Boolean
    isValueList As Boolean
    isActualValueBinary As Boolean
    isActualValueList As Boolean
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

'Prior to writing out a file's metadata, we must cache the information we want written in a temp file.  (ExifTool requires
' a source file when writing metadata out to file; the alternative is to manually request the writing of each tag in turn,
' but if we do this, we lose many built-in utilities like automatically removing duplicate tags, and reassigning invalid
' tags to preferred categories.)  Because writing out metadata is asynchronous, we have to wait for ExifTool to finish
' before deleting the temp file, so we keep a copy of the file's path here.  The stopVerificationMode (which is
' automatically triggered by the newMetadataReceived function as necessary) will remove the file at this location.
Private tmpMetadataFilePath As String

'The FormMain.ShellPipeMain user control will asynchronously trigger this function whenever it receives new metadata
' from ExifTool.
Public Sub newMetadataReceived(ByVal newMetadata As String)
    
    If captureModeActive Then
        curMetadataString = curMetadataString & newMetadata
    
    'During verification mode, we must also check to see if verification has finished
    ElseIf verificationModeActive Then
        verificationString = verificationString & newMetadata
        If isMetadataFinished() Then stopVerificationMode
    End If
    
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
    verificationString = ""
    
    'Verification mode is a bit different.  We need to erase our temporary metadata file if it exists; then we can exit.
    If FileExist(tmpMetadataFilePath) Then Kill tmpMetadataFilePath
    
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
    End If
    
    If InStr(1, tmpMetadata, "{ready}", vbBinaryCompare) > 0 Then
        
        'Terminate the relevant mode
        If captureModeActive Then captureModeActive = False
        If verificationModeActive Then verificationModeActive = False
        
        isMetadataFinished = True
        
    Else
        isMetadataFinished = False
    End If
    
End Function

'When metadata is ready (as determined by a call to isMetadataFinished), it can be retrieved via this function
Public Function retrieveMetadataString() As String
        
    'Because we request metadata in XML format, ExifTool escapes disallowed XML characters.  Convert those back
    ' to standard characters before returning the retrieved metadata.
    If InStr(1, curMetadataString, "&amp;") > 0 Then curMetadataString = Replace(curMetadataString, "&amp;", "&")
    If InStr(1, curMetadataString, "&#39;") > 0 Then curMetadataString = Replace(curMetadataString, "&#39;", "'")
    If InStr(1, curMetadataString, "&quot;") > 0 Then curMetadataString = Replace(curMetadataString, "&quot;", """")
    If InStr(1, curMetadataString, "&gt;") > 0 Then curMetadataString = Replace(curMetadataString, "&gt;", ">")
    If InStr(1, curMetadataString, "&lt;") > 0 Then curMetadataString = Replace(curMetadataString, "&lt;", "<")
    
    'Replace the {ready} text supplied by ExifTool itself
    curMetadataString = Replace$(curMetadataString, "{ready}", "")
    
    'Return the processed string, then erase our copy of it
    retrieveMetadataString = curMetadataString
    curMetadataString = ""
    
End Function

'Is ExifTool available as a plugin?
Public Function isExifToolAvailable() As Boolean
    If FileExist(g_PluginPath & "exiftool.exe") Then isExifToolAvailable = True Else isExifToolAvailable = False
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
        
            'The output string will be a simple version number, e.g. "X.YY", and it will be terminated by a vbCrLf character.
            ' Remove vbCrLf now.
            outputString = Trim$(outputString)
            outputString = Replace(outputString, vbCrLf, "")
            getExifToolVersion = outputString
            
        Else
            getExifToolVersion = ""
        End If
        
    End If
    
End Function

'Start an ExifTool instance (if one isn't already active), and have it process an image file.  Because we now run ExifTool
' asynchronously, this should be done early in the image load process.
Public Function startMetadataProcessing(ByVal srcFile As String, ByVal srcFormat As Long) As Boolean

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
    
    'Erase any previous metadata caches
    curMetadataString = ""
    
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
    
    'Request that binary data be processed.  We have no use for this data within PD, but when it comes time to write
    ' our metadata back out to file, we need to have a copy of it.
    cmdParams = cmdParams & "-b" & vbCrLf
    
    'Requesting binary data also means preview and thumbnail images will be processed.  We DEFINITELY don't want these,
    ' so deny them specifically.
    cmdParams = cmdParams & "-x" & vbCrLf & "PreviewImage" & vbCrLf
    cmdParams = cmdParams & "-x" & vbCrLf & "ThumbnailImage" & vbCrLf
    cmdParams = cmdParams & "-x" & vbCrLf & "PhotoshopThumbnail" & vbCrLf
    
    'Output XML data (a lot more complex, but the only way to retrieve descriptions and names simultaneously)
    cmdParams = cmdParams & "-X" & vbCrLf
    
    'Add the image path
    cmdParams = cmdParams & srcFile & vbCrLf
    
    'Finally, add the special command "-execute" which tells ExifTool to start operations
    cmdParams = cmdParams & "-execute" & vbCrLf
    
    'DEBUG ONLY! Display the param list we have assembled.
    'Debug.Print cmdParams
    
    'Ask the user control to start processing this image's metadata.  It will handle things from here.
    FormMain.shellPipeMain.SendData cmdParams
    
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
    
    'Many ExifTool options are delimited by quotation marks (").  Because VB has the worst character escaping scheme ever conceived, I use
    ' a variable to hold the ASCII equivalent of a quotation mark.  This makes things slightly more readable.
    Dim Quotes As String
    Quotes = Chr(34)
    
    'Grab the ExifTool path, which we will shell and pipe in a moment
    Dim appLocation As String
    appLocation = g_PluginPath & "exiftool.exe"
    
    'Start building a string of ExifTool parameters.  We will send these parameters to stdIn, but ExifTool expects them in
    ' ARGFILE format, e.g. each parameter on its own line.
    Dim cmdParams As String
    cmdParams = ""
    
    'Ignore minor errors and warnings
    cmdParams = cmdParams & "-m" & vbCrLf
        
    'Overwrite the original destination file, but only if the metadata was embedded succesfully
    cmdParams = cmdParams & "-overwrite_original" & vbCrLf
    
    'Copy all tags
    cmdParams = cmdParams & "-tagsfromfile" & vbCrLf & srcMetadataFile & vbCrLf & dstImageFile & vbCrLf
    
    'Regardless of the type of metadata copy we're performing, we need to alter or remove some tags because their
    ' original values are no longer relevant.
    cmdParams = cmdParams & "--Orientation" & vbCrLf
    cmdParams = cmdParams & "--IFD2:ImageWidth" & vbCrLf & "--IFD2:ImageHeight" & vbCrLf
    cmdParams = cmdParams & "-ImageWidth=" & srcPDImage.Width & vbCrLf & "-ExifIFD:ExifImageWidth=" & srcPDImage.Width & vbCrLf
    cmdParams = cmdParams & "-ImageHeight=" & srcPDImage.Height & vbCrLf & "-ExifIFD:ExifImageHeight=" & srcPDImage.Height & vbCrLf
    cmdParams = cmdParams & " -ColorSpace=sRGB" & vbCrLf
    cmdParams = cmdParams & "--Padding" & vbCrLf
    
    'If we were asked to remove GPS data, do so now
    If removeGPS Then cmdParams = cmdParams & "-gps:all=" & vbCrLf
    
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
    cmdParams = cmdParams & "exiftool.exe -stay_open true -@ -"
    
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
        
        'Close our own pipe handles and exit
        FormMain.shellPipeMain.FinishChild
        
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
