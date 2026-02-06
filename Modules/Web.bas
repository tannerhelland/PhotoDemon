Attribute VB_Name = "Web"
'***************************************************************************
'Internet helper functions
'Copyright 2001-2026 by Tanner Helland
'Created: 6/12/01
'Last updated: 12/July/17
'Last update: reorganize the Files module to place web-related stuff here.
'
'PhotoDemon doesn't provide much Internet interop, but when it does, the required functions can be found here.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function ShellExecuteW Lib "shell32" (ByVal hWnd As Long, ByVal ptrToOperationString As Long, ByVal ptrToFileString As Long, ByVal ptrToParameters As Long, ByVal ptrToDirectory As Long, ByVal nShowCmd As Long) As Long

Private Declare Function HttpQueryInfoW Lib "wininet" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByVal ptrSBuffer As Long, ByRef lBufferLength As Long, ByRef lIndex As Long) As Long
Private Declare Function InternetCanonicalizeUrlW Lib "wininet" (ByVal lpszUrl As Long, ByVal lpszBuffer As Long, ByRef lpdwBufferLength As Long, ByVal dwFlags As InternetCanonicalizeUrlFlags) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Long
Private Declare Function InternetOpenW Lib "wininet" (ByVal lpszAgent As Long, ByVal dwAccessType As Long, ByVal lpszProxyName As Long, ByVal lpszProxyBypass As Long, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrlW Lib "wininet" (ByVal hInternetSession As Long, ByVal lpszUrl As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal ptrToBuffer As Long, ByVal dwNumberOfBytesToRead As Long, ByRef lNumberOfBytesRead As Long) As Long

Private Const HTTP_QUERY_CONTENT_LENGTH As Long = 5
Private Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000
Private Const SW_SHOWNORMAL As Long = 1

'Downloading a webpage to a standalone string requires a few wininet declares
Private Enum InternetCanonicalizeUrlFlags
    ICU_NO_ENCODE = &H20000000
    ICU_DECODE = &H10000000
    ICU_NO_META = &H8000000
    ICU_ENCODE_SPACES_ONLY = &H4000000
    ICU_BROWSER_MODE = &H2000000
    ICU_ENCODE_PERCENT = &H1000&
End Enum

#If False Then
    Private Const ICU_NO_ENCODE = &H20000000, ICU_DECODE = &H10000000, ICU_NO_META = &H8000000, ICU_ENCODE_SPACES_ONLY = &H4000000, ICU_BROWSER_MODE = &H2000000, ICU_ENCODE_PERCENT = &H1000&
#End If

'Download code from some random dialogs (e.g. "download image") needs to be moved here! TODO

'Open a string as a hyperlink in the user's default browser
Public Sub OpenURL(ByRef targetURL As String)
    Dim targetAction As String: targetAction = "Open"
    ShellExecuteW FormMain.hWnd, StrPtr(targetAction), StrPtr(targetURL), 0&, 0&, SW_SHOWNORMAL
End Sub

'Quick and sloppy mechanism for parsing the domain name from a URL.  Not well-tested, and used within PD
' for error reporting only (specifically, errors related to loading images from online sources).
Public Function GetDomainName(ByRef srcAddress As String) As String
    
    'Slash direction is always problematic on Windows, so start by normalizing against forward-slashes
    Dim strOutput As String
    If (InStr(1, srcAddress, "\", vbBinaryCompare) <> 0) Then
        strOutput = Replace(srcAddress, "\", "/", 1, -1, vbBinaryCompare)
    Else
        strOutput = srcAddress
    End If
    
    'Look for a "://"; if one is found, remove it and everything to the left of it
    Dim charPos As Long
    charPos = InStr(1, strOutput, "://", vbBinaryCompare)
    If (charPos <> 0) Then strOutput = Right$(strOutput, Len(strOutput) - (charPos + 2))
    
    'Look for another "/"; if one exists, remove it and everything to the right of it
    charPos = InStr(1, strOutput, "/", vbBinaryCompare)
    If (charPos <> 0) Then strOutput = Left$(strOutput, charPos - 1)
    
    GetDomainName = strOutput
    
End Function

Public Sub MapImageLocation()

    If (Not PDImages.GetActiveImage.ImgMetadata.HasGPSMetadata) Then
        PDMsgBox "This image does not contain any GPS metadata.", vbOKOnly Or vbInformation, "No GPS data found"
        Exit Sub
    End If
    
    Dim gMapsURL As String, latString As String, lonString As String
    If PDImages.GetActiveImage.ImgMetadata.FillLatitudeLongitude(latString, lonString) Then
        
        'Build a valid Google maps URL (you can use Google to see what the various parameters mean)
                        
        'Note: I find a zoom of 18 ideal, as that is a common level for switching to an "aerial"
        ' view instead of a satellite view.  Much higher than that and you run the risk of not
        ' having data available at that high of zoom.
        gMapsURL = "https://maps.google.com/maps?f=q&z=18&t=h&q=" & latString & "%2c+" & lonString
        
        'As a convenience, request Google Maps in the current language
        If g_Language.TranslationActive Then
            gMapsURL = gMapsURL & "&hl=" & g_Language.GetCurrentLanguage()
        Else
            gMapsURL = gMapsURL & "&hl=en"
        End If
        
        'Launch Google maps in the user's default browser
        Web.OpenURL gMapsURL
        
    End If
    
End Sub

Private Function CanonicalizeUrl(ByRef srcURL As String, ByRef dstUrl As String, Optional ByVal decodeInstead As Boolean = False) As Boolean
    
    'Flags vary based on encode/decode requirement
    Dim dwFlags As InternetCanonicalizeUrlFlags
    dwFlags = ICU_BROWSER_MODE
    If decodeInstead Then dwFlags = dwFlags Or ICU_DECODE Or ICU_NO_ENCODE
    
    'Retrieve the required destination buffer size
    Const ERROR_INSUFFICIENT_BUFFER As Long = 122
    
    Dim bufSize As Long
    If (InternetCanonicalizeUrlW(StrPtr(srcURL), StrPtr(dstUrl), bufSize, dwFlags) = ERROR_INSUFFICIENT_BUFFER) Then
        If (bufSize > 0) Then
            
            'Prep the buffer and retrieve the canonical URL
            dstUrl = String$(bufSize + 1, 0)
            CanonicalizeUrl = (InternetCanonicalizeUrlW(StrPtr(srcURL), StrPtr(dstUrl), bufSize, dwFlags) = 0)
            If CanonicalizeUrl Then dstUrl = Left$(dstUrl, bufSize)
            
        End If
    End If
    
End Function

'Download the contents of a given URL to a temporary file.  Progress reports will be automatically provided via the
' program progress bar.
'
'If successful, the program will return the full path to the temp file used.  If unsuccessful, a blank string will
' be returned.  Use LenB(returnString) = 0 to check for failure state.
'
'Note that *the calling function* is responsible for cleaning up the temp file!
Public Function DownloadURLToTempFile(ByRef URL As String, Optional ByVal suppressErrorMsgs As Boolean = False) As String
    
    'pdFSO is used for Unicode-compatible destination files
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'Normally changing the cursor is handled by the software processor, but because this function routes
    ' internally, we'll make an exception and change it here. Note that everywhere this function can
    ' terminate (and it's many places - a lot can go wrong while downloading) - the cursor needs to be reset.
    Screen.MousePointer = vbHourglass
    
    'Open an Internet session and assign it a handle
    Dim hInternetSession As Long
    
    Message "Attempting to connect to the Internet..."
    hInternetSession = InternetOpenW(StrPtr("Chromium"), INTERNET_OPEN_TYPE_PRECONFIG, 0, 0, 0)
    
    If (hInternetSession = 0) Then
        If (Not suppressErrorMsgs) Then PDMsgBox "PhotoDemon cannot reach the Internet.  Please double-check your connection and try again.", vbExclamation Or vbOKOnly, "Error"
        DownloadURLToTempFile = vbNullString
        Screen.MousePointer = 0
        Exit Function
    End If
    
    'Using the new Internet session, attempt to find the URL; if found, assign it a handle.
    Message "Verifying URL (this may take a moment)..."
    
    Dim hUrl As Long
    hUrl = InternetOpenUrlW(hInternetSession, StrPtr(URL), 0, 0, INTERNET_FLAG_RELOAD, 0)
    
    If (hUrl = 0) Then
        If (Not suppressErrorMsgs) Then PDMsgBox "PhotoDemon could not reach the target URL (%1).  If the problem persists, try downloading the file manually using your Internet browser.", vbExclamation Or vbOKOnly, "Error", URL
        If (hInternetSession <> 0) Then InternetCloseHandle hInternetSession
        DownloadURLToTempFile = vbNullString
        Screen.MousePointer = 0
        Exit Function
    End If
    
    'Check the size of the file to be downloaded...
    Dim downloadSize As Long, tmpStrBuffer As String
    tmpStrBuffer = String$(256, 0)
    If (HttpQueryInfoW(hUrl, HTTP_QUERY_CONTENT_LENGTH, StrPtr(tmpStrBuffer), LenB(tmpStrBuffer), ByVal 0&) <> 0) Then downloadSize = Val(Strings.TrimNull(tmpStrBuffer)) Else downloadSize = 0
    SetProgBarVal 0
    
    If (downloadSize <> 0) Then SetProgBarMax downloadSize
    
    'We need a temporary file to house the file; generate it automatically, using the extension of the original file.
    PDDebug.LogAction "URL validated.  Creating temporary download file..."
    
    Dim tmpFilename As String
    tmpFilename = cFile.MakeValidWindowsFilename(Files.FileGetName(URL))
    
    'As an added convenience, replace %20 indicators in the filename with actual spaces.
    ' (TODO: move to a full-featured URL encode/decode solution here.)
    If (InStr(1, tmpFilename, "%20", vbBinaryCompare) <> 0) Then tmpFilename = Replace$(tmpFilename, "%20", " ")
    
    Dim tmpFile As String
    tmpFile = UserPrefs.GetTempPath & tmpFilename
    
    'Open the temporary file and begin downloading the image to it
    Message "Downloading..."
        
    Dim hFile As Long
    If cFile.FileCreateHandle(tmpFile, hFile, True, True, OptimizeSequentialAccess) Then
    
        'Prepare a receiving buffer (this will be used to hold chunks of the image)
        Const DEFAULT_BUFFER_SIZE As Long = 2 ^ 16  '65 kb, or TCP/IP packet size upper limit
        Dim tmpBuffer() As Byte
        ReDim tmpBuffer(0 To DEFAULT_BUFFER_SIZE - 1) As Byte
   
        'We will verify each chunk as they're downloaded
        Dim chunkOK As Boolean, numOfBytesRead As Long
        
        'How many bytes of the entire file we've downloaded (so far)
        Dim totalBytesRead As Long
        totalBytesRead = 0
                
        Do
   
            'Read the next chunk of the image
            numOfBytesRead = 0
            chunkOK = (InternetReadFile(hUrl, VarPtr(tmpBuffer(0)), DEFAULT_BUFFER_SIZE, numOfBytesRead) <> 0)
   
            'If something goes horribly wrong, terminate the download
            If (Not chunkOK) Then
                
                If (Not suppressErrorMsgs) Then PDMsgBox "PhotoDemon lost Internet access. Please double-check your connection.", vbExclamation Or vbOKOnly, "Error"
                
                If Files.FileExists(tmpFile) Then
                    cFile.FileCloseHandle hFile
                    Files.FileDelete tmpFile
                End If
                
                If (hUrl <> 0) Then InternetCloseHandle hUrl
                If (hInternetSession <> 0) Then InternetCloseHandle hInternetSession
                
                SetProgBarVal 0
                ReleaseProgressBar
                DownloadURLToTempFile = vbNullString
                Screen.MousePointer = 0
                
                Exit Function
                
            End If
   
            'If the file is done, exit this loop
            If (numOfBytesRead = 0) Then Exit Do
            
            'If we've made it this far, assume we've received legitimate data.  Place that data into the temp file.
            cFile.FileWriteData hFile, VarPtr(tmpBuffer(0)), numOfBytesRead
               
            totalBytesRead = totalBytesRead + numOfBytesRead
            
            If (downloadSize <> 0) Then
            
                SetProgBarVal totalBytesRead
                
                'Display a download update in the message area, but do not log it in the debugger (as there may be
                ' many such notifications, and we don't want to inflate the log unnecessarily)
                If UserPrefs.GenerateDebugLogs Then
                    Message "Downloading (%1 of %2 bytes)...", Format$(totalBytesRead, "#,#0"), Format$(downloadSize, "#,#0"), "DONOTLOG"
                Else
                    Message "Downloading (%1 of %2 bytes)...", Format$(totalBytesRead, "#,#0"), Format$(downloadSize, "#,#0")
                End If
                
            End If
            
        'Carry on
        Loop
        
    End If
    
    'Close the temporary file
    If (hFile <> 0) Then cFile.FileCloseHandle hFile
    
    'Close this URL and Internet session
    If (hUrl <> 0) Then InternetCloseHandle hUrl
    If (hInternetSession <> 0) Then InternetCloseHandle hInternetSession
    
    Message "Download complete. Verifying file integrity..."
    
    'Check to make sure the image downloaded; if the size is unreasonably small, we can assume the site
    ' prevented our download.  (Direct downloads are sometimes treated as hotlinking; similarly, some sites
    ' prevent scraping, which a direct download like this may seem to be.)
    If (totalBytesRead < 20) Then
        
        Message "Download canceled.  (Remote server denied access.)"
        
        Dim domainName As String
        domainName = Web.GetDomainName(URL)
        If (Not suppressErrorMsgs) Then PDMsgBox "Unfortunately, %1 prevented PhotoDemon from directly downloading this file. (Direct downloads are sometimes mistaken as hotlinking by misconfigured servers.)" & vbCrLf & vbCrLf & "You will need to manually download this file using your Internet browser.", vbExclamation Or vbOKOnly, "Error", domainName
        
        Files.FileDeleteIfExists tmpFile
        If (hUrl <> 0) Then InternetCloseHandle hUrl
        If (hInternetSession <> 0) Then InternetCloseHandle hInternetSession
        
        SetProgBarVal 0
        ReleaseProgressBar
        Screen.MousePointer = 0
        
        DownloadURLToTempFile = vbNullString
        Exit Function
        
    End If
    
    'If we made it all the way here, the file was downloaded successfully (most likely... :P)
    SetProgBarVal 0
    ReleaseProgressBar
    Screen.MousePointer = 0
    
    'Return the temp file location
    DownloadURLToTempFile = tmpFile

End Function
