Attribute VB_Name = "LittleCMS"
'***************************************************************************
'LittleCMS Interface
'Copyright 2016-2016 by Tanner Helland
'Created: 21/April/16
'Last updated: 21/April/16
'Last update: initial build
'
'Module for handling all LittleCMS interfacing.  This module is pointless without the accompanying
' LittleCMS plugin, which will be in the App/PhotoDemon/Plugins subdirectory as "lcms2.dll".
'
'LittleCMS is a free, open-source color management library.  You can learn more about it here:
'
' http://www.littlecms.com/
'
'PhotoDemon has been designed against v 2.7.0.  It may not work with other versions.
' Additional documentation regarding the use of LittleCMS is available as part of the official LittleCMS library,
' available from https://github.com/mm2/Little-CMS.
'
'LittleCMS is available under the MIT license.  Please see the App/PhotoDemon/Plugins/lcms2-LICENSE.txt file
' for questions regarding copyright or licensing.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Return the current library version as a Long, e.g. "2.7" is returned as "2070"
Private Declare Function cmsGetEncodedCMMversion Lib "lcms2.dll" () As Long

'A single LittleCMS handle is maintained for the life of a PD instance; see InitializeLCMS and ReleaseLCMS, below.
Private m_LCMSHandle As Long

'Initialize LittleCMS.  Do not call this until you have verified the LCMS plugin's existence
' (typically via the PluginManager module)
Public Function InitializeLCMS() As Boolean
    
    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    Dim lcmsPath As String
    lcmsPath = g_PluginPath & "lcms2.dll"
    m_LCMSHandle = LoadLibrary(StrPtr(lcmsPath))
    InitializeLCMS = CBool(m_LCMSHandle <> 0)
    
    #If DEBUGMODE = 1 Then
        If (Not InitializeLCMS) Then
            pdDebug.LogAction "WARNING!  LoadLibrary failed to load LittleCMS.  Last DLL error: " & Err.LastDllError
            pdDebug.LogAction "(FYI, the attempted path was: " & lcmsPath & ")"
        End If
    #End If
    
End Function

'When PD closes, make sure to release our library handle
Public Sub ReleaseLCMS()
    If (m_LCMSHandle <> 0) Then FreeLibrary m_LCMSHandle
    g_LCMSEnabled = False
End Sub

'After LittleCMS has been initialized, you can call this function to retrieve its current version.
' The version will always be formatted as "Major.Minor.0.0".
Public Function GetLCMSVersion() As String
    Dim versionAsLong As Long
    versionAsLong = cmsGetEncodedCMMversion()
    
    'Split the version by zeroes
    Dim versionAsString() As String
    versionAsString = Split(CStr(versionAsLong), "0", , vbBinaryCompare)
    
    If VB_Hacks.IsArrayInitialized(versionAsString) Then
        If (UBound(versionAsString) >= 1) Then
            GetLCMSVersion = versionAsString(0) & "." & versionAsString(1) & ".0.0"
        Else
            GetLCMSVersion = "0.0.0.0"
        End If
    Else
        GetLCMSVersion = "0.0.0.0"
    End If
    
End Function
