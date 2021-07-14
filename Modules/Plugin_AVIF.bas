Attribute VB_Name = "Plugin_AVIF"
'***************************************************************************
'libavif Interface
'Copyright 2021-2021 by Tanner Helland
'Created: 13/July/21
'Last updated: 13/July/21
'Last update: initial build
'
'Module for handling all libavif interfacing (via avifdec/enc.exe).  This module is pointless without
' those exes, which need to be placed in the App/PhotoDemon/Plugins subdirectory.
'
'libavif is a free, open-source portable-C implementation of the AV1 AVIF still image extension.
' You can learn more about it here:
'
' https://github.com/AOMediaCodec/libavif
'
'PhotoDemon has been designed against v0.9.0 (22 Feb '21).  It may not work with other versions.
' Additional documentation regarding the use of libavif is available as part of the official library,
' downloadable from https://github.com/AOMediaCodec/libavif.  You can also run the exe files manually
' with the -h extension for more details on how they work.
'
'libavif is available under a BSD license.  Please see the App/PhotoDemon/Plugins/avif-LICENSE.txt file
' for questions regarding copyright or licensing.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

'Because libavif only targets x64 targets, we interface with its .exe builds.  This means that
' decoding and encoding support exist separately (i.e. just because the import library exists
' at run-time, doesn't mean the export library also exists; users may only install one or none).
Private m_avifImportAvailable As Boolean, m_avifExportAvailable As Boolean

Public Function GetVersion(ByVal testExportLibrary As Boolean) As String
    
    GetVersion = vbNullString
    
    Dim okToCheck As Boolean
    If testExportLibrary Then
        okToCheck = PluginManager.IsPluginCurrentlyInstalled(CCP_AvifExport)
    Else
        okToCheck = PluginManager.IsPluginCurrentlyInstalled(CCP_AvifImport)
    End If
    
    If okToCheck Then
        
        Dim pluginPath As String
        If testExportLibrary Then
            pluginPath = PluginManager.GetPluginPath & "avifenc.exe"
        Else
            pluginPath = PluginManager.GetPluginPath & "avifdec.exe"
        End If
        
        Dim outputString As String, shellOK As Boolean
        If testExportLibrary Then
            shellOK = ShellExecuteCapture(pluginPath, "avifenc.exe -v", outputString)
        Else
            shellOK = ShellExecuteCapture(pluginPath, "avifdec.exe -v", outputString)
        End If
        
        If shellOK Then
        
            'The output string is potentially quite large, and not stable between releases.
            ' For now, just blindly search for the text "Version: "
            Dim vPos As Long, targetString As String
            targetString = "Version: "
            vPos = InStr(1, outputString, targetString, vbTextCompare)
            
            If (vPos <> 0) Then
                
                'Look for a space, linebreak, or end of string
                vPos = vPos + Len(targetString)
                
                On Error GoTo BadVersion
                Do While (vPos < Len(targetString)) And (Mid$(outputString, vPos, 1) <> " ")
                    vPos = vPos + 1
                Loop
                
                Dim ePos As Long
                ePos = InStr(vPos, outputString, " ", vbBinaryCompare)
                If (ePos < 0) Then ePos = InStr(vPos, outputString, vbLf, vbBinaryCompare)
                If (ePos < 0) Then ePos = Len(outputString)
                
                Dim verString As String
                verString = "???"
                verString = Trim$(Mid$(outputString, vPos, ePos - vPos))
                
BadVersion:
                GetVersion = verString
            
            'Failure to return version number is a bad sign, but this isn't the place to handle it.
            Else
                PDDebug.LogAction "WARNING: couldn't retrieve version number of libavif."
            End If
            
        End If
        
    End If
    
End Function

Public Function InitializeEngines(ByRef pathToDLLFolder As String) As Boolean
    
    'Before doing anything else, make sure the OS supports 64-bit apps.
    ' (libavif does not natively support x86 targets)
    If (Not OS.OSSupports64bitExe()) Then
        m_avifExportAvailable = False
        m_avifImportAvailable = False
        InitializeEngines = False
        PDDebug.LogAction "WARNING!  AVIF support not available; system is only 32-bit"
        Exit Function
    End If
    
    'Test import and export support separately
    Dim importPath As String, exportPath As String
    importPath = pathToDLLFolder & "avifdec.exe"
    exportPath = pathToDLLFolder & "avifenc.exe"
    
    m_avifExportAvailable = Files.FileExists(exportPath)
    m_avifImportAvailable = Files.FileExists(importPath)
    
    InitializeEngines = m_avifImportAvailable Or m_avifExportAvailable
    
    If (Not InitializeEngines) Then
        PDDebug.LogAction "WARNING!  AVIF support not available; plugins missing"
    End If
    
End Function

Public Function IsAVIFExportAvailable() As Boolean
    IsAVIFExportAvailable = m_avifExportAvailable
End Function

Public Function IsAVIFImportAvailable() As Boolean
    IsAVIFImportAvailable = m_avifImportAvailable
End Function
