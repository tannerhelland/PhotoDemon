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

'Enable at DEBUG-TIME for verbose logging
Private Const J2K_DEBUG_VERBOSE As Boolean = True

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

'**********************************************
' / END GENERIC 3RD-PARTY LIBRARY BOILERPLATE
'**********************************************

'Verify J2K file signature.  Doesn't require OpenJPEG.
Public Function IsFileJ2K(ByRef srcFile As String) As Boolean

    Const FUNC_NAME As String = "IsFileJ2K"
    IsFileJ2K = False
    
    If Files.FileExists(srcFile) Then
        
        'Pull the first 12 bytes only
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
        
            Dim bFirst12() As Byte
            If cStream.ReadBytes(bFirst12, 12, True) Then
                
                'Various different signatures are valid, based on the container used.
                ' (Magic numbers taken from https://github.com/uclouvain/openjpeg/blob/41c25e3827c68a39b9e20c1625a0b96e49955445/src/bin/jp2/opj_decompress.c#L532)
                '#define JP2_RFC3745_MAGIC "\x00\x00\x00\x0c\x6a\x50\x20\x20\x0d\x0a\x87\x0a"
                '#define JP2_MAGIC "\x0d\x0a\x87\x0a"
                '#define J2K_CODESTREAM_MAGIC "\xff\x4f\xff\x51"
                Const JP2_RFC3745_MAGIC_1 As Long = &HC000000
                Const JP2_RFC3745_MAGIC_2 As Long = &H2020506A
                Const JP2_RFC3745_MAGIC_3 As Long = &HA870A0D
                 
                If VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(JP2_RFC3745_MAGIC_1), 4) Then
                    If VBHacks.MemCmp(VarPtr(bFirst12(4)), VarPtr(JP2_RFC3745_MAGIC_2), 4) Then
                        IsFileJ2K = VBHacks.MemCmp(VarPtr(bFirst12(8)), VarPtr(JP2_RFC3745_MAGIC_3), 4)
                        If IsFileJ2K And J2K_DEBUG_VERBOSE Then PDDebug.LogAction "File is JP2_RFC3745"
                    End If
                End If
                
                If (Not IsFileJ2K) Then
                    
                    Const JP2_MAGIC As Long = &HA870A0D
                    Const J2K_CODESTREAM_MAGIC As Long = &H51FF4FFF
                    
                    IsFileJ2K = VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(JP2_MAGIC), 4)
                    If IsFileJ2K And J2K_DEBUG_VERBOSE Then PDDebug.LogAction "File is JP2 stream"
                    If (Not IsFileJ2K) Then
                        IsFileJ2K = VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(J2K_CODESTREAM_MAGIC), 4)
                        If IsFileJ2K And J2K_DEBUG_VERBOSE Then PDDebug.LogAction "File is J2K stream"
                    End If
                    
                End If
                
            End If
        End If
        
        Set cStream = Nothing
        
    End If
    
End Function

'Load a J2K image.  Will validate the file prior to loading.  Requires OpenJPEG.
Public Function LoadJ2K(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    Const FUNC_NAME As String = "LoadJ2K"
    LoadJ2K = False
    
    'Failsafe check; this function is pointless if OpenJPEG doesn't exist
    If (Not Plugin_OpenJPEG.IsOpenJPEGEnabled()) Then Exit Function
    
    'Failsafe check; validate file signature (hopefully the caller did this, but you never know)
    If (Not Plugin_OpenJPEG.IsFileJ2K(srcFile)) Then Exit Function
    
End Function
