Attribute VB_Name = "Plugin_WebP"
'***************************************************************************
'WebP Library Interface
'Copyright 2021-2021 by Tanner Helland
'Created: 22/September/21
'Last updated: 22/September/21
'Last update: initial build
'
'Per its documentation (available at https://github.com/webmproject/libwebp/), WebP is...
'
'"WebP codec: library to encode and decode images in WebP format. This package contains the library
' that can be used in other programs to add WebP support..."
'
'LibWebP is BSD-licensed and actively maintained by Google.  Fortunately for PhotoDemon, the developers
' also provide a robust C interface and legacy compilation options, enabling support all the way back
' to Windows XP (hypothetically - testing XP is still TODO).
'
'PhotoDemon historically used FreeImage to manage WebP files, but using libwebp directly allows for
' better performance and feature support, including animated WebP files (which do not work via FreeImage).
'
'Note that all features in this module rely on the libwebp binaries that ship with PhotoDemon.
' These features will not work if libwebp cannot be located.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Note that regular WebP features reside in libwebp; animation features are in libwebpde/mux
Private Declare Function WebPGetDecoderVersion Lib "libwebp" () As Long

'Library handle will be non-zero if each library is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE
Private m_hLibWebP As Long, m_hLibWebPDemux As Long, m_hLibWebPMux As Long, m_LibAvailable As Boolean

'Forcibly disable libwebp interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetVersion() As String
    
    'Byte version numbers get packed into a long
    Dim versionAsInt(0 To 3) As Byte
    
    Dim tmpLong As Long
    tmpLong = WebPGetDecoderVersion()
    PutMem4 VarPtr(versionAsInt(0)), tmpLong
    
    'Want to ensure we retrieved the correct values?  Use this:
    'Debug.Print versionAsInt(0), versionAsInt(1), versionAsInt(2), versionAsInt(3)
    GetVersion = versionAsInt(0) & "." & versionAsInt(1) & "." & versionAsInt(2) & "." & versionAsInt(3)
    
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim strLibPath As String
    strLibPath = pathToDLLFolder & "libwebp.dll"
    m_hLibWebP = VBHacks.LoadLib(strLibPath)
    strLibPath = pathToDLLFolder & "libwebpdemux.dll"
    m_hLibWebPDemux = VBHacks.LoadLib(strLibPath)
    strLibPath = pathToDLLFolder & "libwebpmux.dll"
    m_hLibWebPMux = VBHacks.LoadLib(strLibPath)
    
    m_LibAvailable = (m_hLibWebP <> 0) And (m_hLibWebPDemux <> 0) And (m_hLibWebPMux <> 0)
    InitializeEngine = m_LibAvailable
    
    If (Not InitializeEngine) Then PDDebug.LogAction "WARNING!  LoadLibraryW failed to load one or more WebP libraries.  Last DLL error: " & Err.LastDllError
    
End Function

Public Function IsWebPEnabled() As Boolean
    IsWebPEnabled = m_LibAvailable
End Function

Public Sub ReleaseEngine()
    If (m_hLibWebP <> 0) Then
        VBHacks.FreeLib m_hLibWebP
        m_hLibWebP = 0
    End If
    If (m_hLibWebPDemux <> 0) Then
        VBHacks.FreeLib m_hLibWebPDemux
        m_hLibWebPDemux = 0
    End If
    If (m_hLibWebPMux <> 0) Then
        VBHacks.FreeLib m_hLibWebPMux
        m_hLibWebPMux = 0
    End If
End Sub

