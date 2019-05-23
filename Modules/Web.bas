Attribute VB_Name = "Web"
'***************************************************************************
'Internet helper functions
'Copyright 2001-2019 by Tanner Helland
'Created: 6/12/01
'Last updated: 12/July/17
'Last update: reorganize the Files module to place web-related stuff here.
'
'PhotoDemon doesn't provide much Internet interop, but when it does, the required functions can be found here.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function ShellExecuteW Lib "shell32" (ByVal hWnd As Long, ByVal ptrToOperationString As Long, ByVal ptrToFileString As Long, ByVal ptrToParameters As Long, ByVal ptrToDirectory As Long, ByVal nShowCmd As Long) As Long
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

Private Declare Function InternetCanonicalizeUrlW Lib "wininet" (ByVal lpszUrl As Long, ByVal lpszBuffer As Long, ByRef lpdwBufferLength As Long, ByVal dwFlags As InternetCanonicalizeUrlFlags) As Long
Private Declare Function HttpQueryInfoW Lib "wininet" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByVal ptrSBuffer As Long, ByRef lBufferLength As Long, ByRef lIndex As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Long
Private Declare Function InternetOpenW Lib "wininet" (ByVal lpszAgent As Long, ByVal dwAccessType As Long, ByVal lpszProxyName As Long, ByVal lpszProxyBypass As Long, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrlW Lib "wininet" (ByVal hInternetSession As Long, ByVal lpszUrl As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal ptrToBuffer As Long, ByVal dwNumberOfBytesToRead As Long, ByRef lNumberOfBytesRead As Long) As Long

Private Const HTTP_QUERY_CONTENT_LENGTH As Long = 5
Private Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

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

'Retrieve a given URL as a string.  Returns TRUE if successful; FALSE otherwise.  If FALSE is returned,
' check the debug log for details.
Public Function GetURLAsString(ByRef srcUrl As String, ByRef dstText As String) As Boolean
    
    'Open an Internet session
    Dim hInternetSession As Long
    hInternetSession = InternetOpenW(StrPtr(App.EXEName), INTERNET_OPEN_TYPE_PRECONFIG, 0, 0, 0)
    
    If (hInternetSession = 0) Then
        PDDebug.LogAction "GetURLAsString failed to retrieve a session handle."
        GetURLAsString = False
        Exit Function
    End If
    
    'Using the new Internet session, attempt to find the URL; if found, assign it a handle.
    Dim hUrl As Long
    hUrl = InternetOpenUrlW(hInternetSession, StrPtr(srcUrl), 0, 0, INTERNET_FLAG_RELOAD, 0)
    
    If (hUrl = 0) Then
        PDDebug.LogAction "GetURLAsString failed to find the requested URL: " & srcUrl
        If (hInternetSession <> 0) Then InternetCloseHandle hInternetSession
        GetURLAsString = False
        Exit Function
    End If
    
    'If you wanted to, you could check file size here; we currently skip this step as this function
    ' is not intended for user-facing tasks.
    Dim cFSO As pdFSO
    Set cFSO = New pdFSO
    
    'Prep a temporary filename
    Dim tmpFilename As String, canonUrl As String
    If (Not CanonicalizeUrl(srcUrl, canonUrl)) Then canonUrl = srcUrl
    Debug.Print srcUrl
    Debug.Print canonUrl
    tmpFilename = cFSO.MakeValidWindowsFilename(Files.FileGetName(canonUrl))
    
    Dim tmpFile As String
    tmpFile = UserPrefs.GetTempPath & tmpFilename
    
    'Initiate download
    Dim hFile As Long
    If cFSO.FileCreateHandle(tmpFile, hFile, True, True, OptimizeSequentialAccess) Then
    
        'Prepare a receiving buffer (this will be used to hold chunks of the image)
        Const DEFAULT_BUFFER_SIZE As Long = 2 ^ 16  '65 kb, or TCP/IP packet size upper limit
        Dim dataBuffer() As Byte
        ReDim dataBuffer(0 To DEFAULT_BUFFER_SIZE - 1) As Byte
   
        'Verify each chunk as it's downloaded
        Dim chunkOK As Boolean, numOfBytesRead As Long
        
        'How many bytes of the file we've downloaded (so far)
        Dim totalBytesRead As Long
        totalBytesRead = 0
                
        Do
   
            'Read the next chunk of the image
            numOfBytesRead = 0
            chunkOK = (InternetReadFile(hUrl, VarPtr(dataBuffer(0)), DEFAULT_BUFFER_SIZE, numOfBytesRead) <> 0)
   
            'If something goes horribly wrong, terminate the download
            If (Not chunkOK) Then
                
                PDDebug.LogAction "GetURLAsString failed; Internet access lost."
                
                If Files.FileExists(tmpFile) Then
                    cFSO.FileCloseHandle hFile
                    Files.FileDelete tmpFile
                End If
                
                If (hUrl <> 0) Then InternetCloseHandle hUrl
                If (hInternetSession <> 0) Then InternetCloseHandle hInternetSession
                GetURLAsString = False
                
                Exit Function
                
            End If
   
            'If the file is done, exit this loop
            If (numOfBytesRead = 0) Then Exit Do
            
            'If we've made it this far, assume we've received legitimate data.  Place that data into the temp file.
            cFSO.FileWriteData hFile, VarPtr(dataBuffer(0)), numOfBytesRead
               
            totalBytesRead = totalBytesRead + numOfBytesRead
            
        'Carry on
        Loop
        
    End If
    
    'Close the temporary file and free our buffer
    Erase dataBuffer
    If (hFile <> 0) Then cFSO.FileCloseHandle hFile
    
    'Close this URL and Internet session
    If (hUrl <> 0) Then InternetCloseHandle hUrl
    If (hInternetSession <> 0) Then InternetCloseHandle hInternetSession
    
    'Check to make sure the file downloaded; if the size is unreasonably small, we can assume the host
    ' prevented our download.  (Direct downloads are sometimes treated as hotlinking; similarly, some sites
    ' prevent scraping, which a direct download like this may seem to be.)
    If (totalBytesRead < 20) Then
        
        PDDebug.LogAction "GetURLAsString failed (remote server denied access)"
        
        Files.FileDeleteIfExists tmpFile
        If (hUrl <> 0) Then InternetCloseHandle hUrl
        If (hInternetSession <> 0) Then InternetCloseHandle hInternetSession
        GetURLAsString = False
        
        Exit Function
        
    End If
    
    'If we made it all the way here, the file was downloaded successfully (most likely... with web stuff,
    ' it's always possible that an esoteric error occurred, but we did our due diligence in attempting
    ' the download!)
    
    'Retrieve the temp file into a string and return it.
    GetURLAsString = Files.FileLoadAsString(tmpFile, dstText)
    Files.FileDeleteIfExists tmpFile
    
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

Private Function CanonicalizeUrl(ByRef srcUrl As String, ByRef dstUrl As String, Optional ByVal decodeInstead As Boolean = False) As Boolean
    
    'Flags vary based on encode/decode requirement
    Dim dwFlags As InternetCanonicalizeUrlFlags
    dwFlags = ICU_BROWSER_MODE
    If decodeInstead Then dwFlags = dwFlags Or ICU_DECODE Or ICU_NO_ENCODE
    
    'Retrieve the required destination buffer size
    Const ERROR_INSUFFICIENT_BUFFER As Long = 122
    
    Dim bufSize As Long
    If (InternetCanonicalizeUrlW(StrPtr(srcUrl), StrPtr(dstUrl), bufSize, dwFlags) = ERROR_INSUFFICIENT_BUFFER) Then
        If (bufSize > 0) Then
            
            'Prep the buffer and retrieve the canonical URL
            dstUrl = String$(bufSize + 1, 0)
            CanonicalizeUrl = (InternetCanonicalizeUrlW(StrPtr(srcUrl), StrPtr(dstUrl), bufSize, dwFlags) = 0)
            If CanonicalizeUrl Then dstUrl = Left$(dstUrl, bufSize)
            
        End If
    End If
    
End Function
