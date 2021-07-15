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

'Using PNG or JPEG as an intermediary format is a tough call.  PNG is lossless, which should make
' it the obvious preference... but damn, PNG encoding is slow.  (I've been spoiled by PD's internal
' encoder lol.)  JPEG as an intermediary format can be 10-30x faster on large images, so I'm inclined
' to use it at present.  You can toggle this constant at compile-time to require PNG only.
Private Const REQUIRE_LOSSLESS_INTERMEDIARY As Boolean = False

Public Function ConvertAVIFtoStandardImage(ByRef srcFile As String, ByRef dstFile As String, Optional ByRef outputPDIF As PD_IMAGE_FORMAT = PDIF_PNG) As Boolean
    
    Const funcName As String = "ConvertAVIFtoStandardImage"
    
    'Safety checks on plugin
    If (Not m_avifImportAvailable) Then
        InternalError funcName, "libavif broken or missing"
        Exit Function
    End If
    
    Dim pluginPath As String
    pluginPath = PluginManager.GetPluginPath & "avifdec.exe"
    If (Not Files.FileExists(pluginPath)) Then
        InternalError funcName, "libavif missing"
        Exit Function
    End If
    
    'Safety checks on source file
    If (Not Files.FileExists(srcFile)) Then
        InternalError funcName, "source file doesn't exist"
        Exit Function
    End If
    
    'If the destination file isn't specified, generate a random temp file name
    If (Not Files.FileExists(dstFile)) Then dstFile = OS.UniqueTempFilename()
    
    'Ensure destination file has an appropriate extension (this is how the decoder
    ' figures out which format to use)
    Dim reqExtension As String
    If REQUIRE_LOSSLESS_INTERMEDIARY Then
        reqExtension = "png"
        outputPDIF = PDIF_PNG
    Else
        reqExtension = "jpg"
        outputPDIF = PDIF_JPEG
    End If
    If Strings.StringsNotEqual(Files.FileGetExtension(dstFile), reqExtension, True) Then dstFile = dstFile & "." & reqExtension
    
    'Shell plugin and wait for return
    Dim shellCmd As pdString
    Set shellCmd = New pdString
    shellCmd.Append "avifdec.exe "
    
    'Use all available cores for decoding
    shellCmd.Append "-j "
    shellCmd.Append Trim$(Str$(OS.LogicalCoreCount())) & " "
    
    'If using a lossy intermediary (jpeg), attempt to maximize quality
    If (Not REQUIRE_LOSSLESS_INTERMEDIARY) Then shellCmd.Append "-q 100 "
    
    'Append space-safe source image
    shellCmd.Append """"
    shellCmd.Append srcFile
    shellCmd.Append """ "
    
    'Append space-safe destination image
    shellCmd.Append """"
    shellCmd.Append dstFile
    shellCmd.Append """"
    
    'Shell plugin and capture output for analysis
    Dim outputString As String
    If ShellExecuteCapture(pluginPath, shellCmd.ToString(), outputString) Then
    
        'Shell appears successful.  The output string will have two easy-to-check flags if
        ' the conversion was successful.  Don't return success unless we find both.
        Dim targetStringSrc As String, targetStringDst As String
        targetStringSrc = "Image decoded: " & srcFile
        
        If REQUIRE_LOSSLESS_INTERMEDIARY Then
            targetStringDst = "Wrote PNG: "
        Else
            targetStringDst = "Wrote JPEG: "
        End If
        targetStringDst = targetStringDst & dstFile
        
        ConvertAVIFtoStandardImage = (Strings.StrStrBM(outputString, targetStringSrc, 1, True) > 0)
        ConvertAVIFtoStandardImage = ConvertAVIFtoStandardImage And (Strings.StrStrBM(outputString, targetStringDst, 1, True) > 0)
        
        'Want to review the output string manually?  Print it here:
        'PDDebug.LogAction outputString
        
        'Record full details of failures
        If ConvertAVIFtoStandardImage Then
            PDDebug.LogAction "libavif reports success; transferring image to internal parser..."
        Else
            InternalError funcName, "load failed; output follows:"
            PDDebug.LogAction outputString
        End If
        
    Else
        InternalError funcName, "shell failed"
    End If
    
End Function

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

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String)
    PDDebug.LogAction "WARNING! libavif error reported in " & funcName & "(): " & errDescription
End Sub
