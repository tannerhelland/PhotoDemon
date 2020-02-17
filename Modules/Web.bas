Attribute VB_Name = "Web"
'***************************************************************************
'Internet helper functions
'Copyright 2001-2020 by Tanner Helland
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
