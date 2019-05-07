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
