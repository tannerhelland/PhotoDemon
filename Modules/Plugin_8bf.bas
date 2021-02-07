Attribute VB_Name = "Plugin_8bf"
'***************************************************************************
'8bf Plugin Interface
'Copyright 2021-2021 by Tanner Helland
'Created: 07/February/21
'Last updated: 07/February/21
'Last update: initial build
'
'8bf files are 3rd-party Adobe Photoshop plugins that implement one or more "filters".  These are
' basically DLL files with special interfaces for communicating with a parent Photoshop instance.
'
'We attempt to support these plugins in PhotoDemon, with PD standing in for Photoshop as the
' "host" of the plugins.
'
'This feature relies on the 3rd-party "pspihost" library by Sinisa Petric.  This library is
' MIT-licensed and available from GitHub (link good as of Feb 2020):
' https://github.com/spetric/Photoshop-Plugin-Host/blob/master/LICENSE
'
'Thank you to Sinisa for their great work.
'
'Note that the pspihost library must be modified to work with a VB6 project like PhotoDemon.
' VB6 only understands stdcall calling convention, particularly with callbacks (which are used
' heavily by the 8bf format).  You cannot use a default pspihost release as-is and expect it to
' work.  (The pspihost copy that ships with PD has obviously been modified to work with PD;
' I mention this only for intrepid developers who attempt to compile it themselves.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Private Declare Function pspiGetVersion Lib "pspiHost.dll" Alias "_pspiGetVersion@0" () As Long

Private m_LibHandle As Long, m_LibAvailable As Boolean

Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetPspiVersion() As String
    Dim ptrVersion As Long
    ptrVersion = pspiGetVersion()
    If (ptrVersion <> 0) Then GetPspiVersion = Strings.StringFromCharPtr(ptrVersion, False, 3, True) & ".0"
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim strLibPath As String
    strLibPath = pathToDLLFolder & "pspiHost.dll"
    m_LibHandle = VBHacks.LoadLib(strLibPath)
    m_LibAvailable = (m_LibHandle <> 0)
    InitializeEngine = m_LibAvailable
    
    If (Not InitializeEngine) Then
        PDDebug.LogAction "WARNING!  LoadLibraryW failed to load pspiHost.  Last DLL error: " & Err.LastDllError
    End If
    
End Function

Public Function IsPspiEnabled() As Boolean
    IsPspiEnabled = m_LibAvailable
End Function

Public Sub ReleaseEngine()
    If (m_LibHandle <> 0) Then
        VBHacks.FreeLib m_LibHandle
        m_LibHandle = 0
    End If
End Sub

