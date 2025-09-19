Attribute VB_Name = "Plugin_OpenJPEG"
'***************************************************************************
'OpenJPEG (JPEG-2000) Library Interface
'Copyright 2025-2025 by Tanner Helland
'Created: 19/September/25
'Last updated: 19/September/25
'Last update: initial build
'
'Per its documentation (available at https://www.openjpeg.org/), OpenJPEG is...
'
' "...an open-source JPEG 2000 codec written in C language. It has been developed in order to promote
'  the use of JPEG 2000, a still-image compression standard from the Joint Photographic Experts Group
'  (JPEG). Since may 2015, it is officially recognized by ISO/IEC and ITU-T as a JPEG 2000 Reference
'  Software."
'
'OpenJPEG is BSD-licensed and actively maintained.
'
'PhotoDemon originally used OpenJPEG via FreeImage, but FreeImage's abandonment made it impossible
' to continue down that route.  As such, in 2025 I wrote a new, direct OpenJPEG interface from scratch.
' (The only one of its kind in VB6, AFAIK.)
'
'This interface was designed against v2.5.3 of the library (released 09 Dec 2024).  It should work fine
' with any version that maintains ABI compatibility with that release.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Library handle will be non-zero if CharLS is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE
Private m_LibHandle As Long, m_LibAvailable As Boolean

'Official OpenJPEG builds use stdcall
Private Declare Function opj_version Lib "openjp2" Alias "_opj_version@0" () As Long

'Forcibly disable CharLS interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetVersion() As String

    If (m_LibHandle <> 0) And m_LibAvailable Then
        
        Dim ptrVersion As Long
        ptrVersion = opj_version()
        
        GetVersion = Strings.StringFromCharPtr(ptrVersion, False)
        
    End If
    
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim strLibPath As String
    strLibPath = pathToDLLFolder & "openjp2.dll"
    m_LibHandle = VBHacks.LoadLib(strLibPath)
    m_LibAvailable = (m_LibHandle <> 0)
    InitializeEngine = m_LibAvailable
    
    If (Not InitializeEngine) Then PDDebug.LogAction "WARNING!  LoadLibraryW failed to load OpenJPEG.  Last DLL error: " & Err.LastDllError
    
End Function

Public Function IsOpenJPEGEnabled() As Boolean
    IsOpenJPEGEnabled = m_LibAvailable
End Function

Public Sub ReleaseEngine()
    If (m_LibHandle <> 0) Then
        VBHacks.FreeLib m_LibHandle
        m_LibHandle = 0
    End If
End Sub

Private Sub InternalError(Optional ByRef funcName As String = vbNullString, Optional ByRef errString As String = vbNullString)
    PDDebug.LogAction "OpenJPEG error in " & funcName & "(): " & errString, PDM_External_Lib
End Sub
