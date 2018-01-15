Attribute VB_Name = "Web"
'***************************************************************************
'Internet helper functions
'Copyright 2001-2018 by Tanner Helland
'Created: 6/12/01
'Last updated: 12/July/17
'Last update: reorganize the Files module to place web-related stuff here.
'
'PhotoDemon doesn't provide much Internet interop, but when it does, the required functions can be found here.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Open a string as a hyperlink in the user's default browser
Public Sub OpenURL(ByVal targetURL As String)
    Dim targetAction As String: targetAction = "Open"
    ShellExecute FormMain.hWnd, StrPtr(targetAction), StrPtr(targetURL), 0&, 0&, SW_SHOWNORMAL
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
