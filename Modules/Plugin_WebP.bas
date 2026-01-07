Attribute VB_Name = "Plugin_WebP"
'***************************************************************************
'WebP Library Interface
'Copyright 2021-2026 by Tanner Helland
'Created: 22/September/21
'Last updated: 25/September/23
'Last update: add code for new libwebp dependency (libsharpyuv)
'
'Per its documentation (available at https://github.com/webmproject/libwebp/), libwebp is...
'
'"WebP codec: library to encode and decode images in WebP format. This package contains the library
' that can be used in other programs to add WebP support..."
'
'LibWebP is BSD-licensed and actively maintained by Google.  Fortunately for PhotoDemon, the developers
' also provide a robust C interface and legacy compilation options, enabling support all the way back
' to Windows XP (hypothetically - testing XP is still TODO).
'
'PhotoDemon historically used FreeImage to manage WebP files, but using libwebp directly allows for
' better performance and feature support, including animated WebP support (which do not work via FreeImage).
'
'Note that all features in this module rely on the libwebp binaries that ship with PhotoDemon.
' These features will not work if libwebp cannot be located.  Look in the pdWebP class for details
' on various APIs; they are all declared there.  (This module just provides basic library initialization
' and termination.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'WebP API declares are inside pdWebP; this module only handles initializing and releasing the
' underlying libwebp instanceonly version-checking is handled here.

'Library handle will be non-zero if each library is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE
Private m_hLibWebP As Long, m_hLibWebPDemux As Long, m_hLibWebPMux As Long, m_LibAvailable As Boolean

'As of v1.3.2 (possibly earlier; this was when I updated from 1.2.4), libwebp has an extra dependency on
' libsharpyuv (bundled with PD).  We must load this *prior* to loading any other webp dlls.
Private m_hLibSharpYUV As Long

'Forcibly disable libwebp interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetHandle_LibWebP() As Long
    GetHandle_LibWebP = m_hLibWebP
End Function

Public Function GetHandle_LibWebPDemux() As Long
    GetHandle_LibWebPDemux = m_hLibWebPDemux
End Function

Public Function GetHandle_LibWebPMux() As Long
    GetHandle_LibWebPMux = m_hLibWebPMux
End Function

Public Function GetVersion() As String
        
    If (m_hLibWebP = 0) Or (Not m_LibAvailable) Then Exit Function
        
    'Byte version numbers get packed into a long
    Dim versionAsInt(0 To 3) As Byte
    
    Dim tmpLong As Long, cWebP As pdWebP
    Set cWebP = New pdWebP
    tmpLong = cWebP.GetLibraryVersion()
    PutMem4 VarPtr(versionAsInt(0)), tmpLong
    
    'Want to ensure we retrieved the correct values?  Use this:
    GetVersion = versionAsInt(2) & "." & versionAsInt(1) & "." & versionAsInt(0) & "." & versionAsInt(3)
    
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean
    
    'Initialize all webp libraries
    Dim strLibPath As String
    
    'New to 1.3.2 is the sharpyuv lib which must be loaded *first*
    strLibPath = pathToDLLFolder & "libsharpyuv.dll"
    m_hLibSharpYUV = VBHacks.LoadLib(strLibPath)
    
    'libwebp can now resolve dependencies correctly...
    strLibPath = pathToDLLFolder & "libwebp.dll"
    m_hLibWebP = VBHacks.LoadLib(strLibPath)
    strLibPath = pathToDLLFolder & "libwebpdemux.dll"
    m_hLibWebPDemux = VBHacks.LoadLib(strLibPath)
    strLibPath = pathToDLLFolder & "libwebpmux.dll"
    m_hLibWebPMux = VBHacks.LoadLib(strLibPath)
    m_LibAvailable = (m_hLibSharpYUV <> 0) And (m_hLibWebP <> 0) And (m_hLibWebPDemux <> 0) And (m_hLibWebPMux <> 0)
    InitializeEngine = m_LibAvailable
    
    If (Not InitializeEngine) Then PDDebug.LogAction "WARNING!  LoadLibraryW failed to load one or more WebP libraries.  Last DLL error: " & Err.LastDllError
    
End Function

Public Function IsWebPEnabled() As Boolean
    IsWebPEnabled = m_LibAvailable
End Function

Public Sub ReleaseEngine()
    
    'For extra safety, free in reverse order from loading
    If (m_hLibWebPMux <> 0) Then
        VBHacks.FreeLib m_hLibWebPMux
        m_hLibWebPMux = 0
    End If
    If (m_hLibWebPDemux <> 0) Then
        VBHacks.FreeLib m_hLibWebPDemux
        m_hLibWebPDemux = 0
    End If
    If (m_hLibWebP <> 0) Then
        VBHacks.FreeLib m_hLibWebP
        m_hLibWebP = 0
    End If
    If (m_hLibSharpYUV <> 0) Then
        VBHacks.FreeLib m_hLibSharpYUV
        m_hLibSharpYUV = 0
    End If
    
End Sub

'Import/Export/Validation functions follow

'Test for WebP format.  Fast because it does not rely on libwebp at all; instead, it's simply a check
' for the fixed WebP "magic numbers".
Public Function IsWebP(ByRef srcFile As String) As Boolean
    
    IsWebP = False
    
    'The first 12 bytes of the file have two magic numbers (bytes [0, 3] and [8, 11]) we can use
    ' to reliably identify WebM files.
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_FileBacked, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
        
        'Check the first magic number
        Const WEBM_MAGIC_NO_1 As Long = &H52494646
        If (cStream.ReadLong_BE() = WEBM_MAGIC_NO_1) Then
            
            'Skip the next int (which is file size), then check the following one for the second magic number
            cStream.ReadLong
            Const WEBM_MAGIC_NO_2 As Long = &H57454250
            IsWebP = (cStream.ReadLong_BE() = WEBM_MAGIC_NO_2)
        
        '/first magic no. failed
        End If
    
    '/couldn't access file
    End If
    
End Function
