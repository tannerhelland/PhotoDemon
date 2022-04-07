Attribute VB_Name = "Plugin_AVIF"
'***************************************************************************
'libavif Interface
'Copyright 2021-2022 by Tanner Helland
'Created: 13/July/21
'Last updated: 06/April/21
'Last update: use PD-specific repo for avifdec/enc since the libavif authors keep changing release availability;
'             also, enable fast PNG writing thanks to an issue I submitted!
'
'Module for handling all libavif interfacing (via avifdec/enc.exe).  This module is pointless without
' those exes, which need to be placed in the App/PhotoDemon/Plugins subdirectory.  (PD will automatically
' download these for you if you attempt to interact with AVIF files.)
'
'libavif is a free, open-source portable-C implementation of the AV1 AVIF still image extension.
' You can learn more about it here:
'
' https://github.com/AOMediaCodec/libavif
'
'PhotoDemon has been designed against v0.10.0 (06 April '22).  It may not work with other versions.
' Additional documentation regarding the use of libavif is available as part of the official library,
' downloadable from https://github.com/AOMediaCodec/libavif.  You can also run the exe files manually
' with the -h extension for details on how they work.
'
'libavif is available under a BSD license.  Please see the App/PhotoDemon/Plugins/avif-LICENSE.txt file
' for questions regarding copyright or licensing.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Because libavif only targets x64 targets, we interface with its .exe builds.  This means that
' decoding and encoding support exist separately (i.e. just because the import library exists
' at run-time, doesn't mean the export library also exists; users may only install one or none).
Private m_avifImportAvailable As Boolean, m_avifExportAvailable As Boolean

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
    reqExtension = "png"
    outputPDIF = PDIF_PNG
    If Strings.StringsNotEqual(Files.FileGetExtension(dstFile), reqExtension, True) Then dstFile = dstFile & "." & reqExtension
    
    'Shell plugin and wait for return
    Dim shellCmd As pdString
    Set shellCmd = New pdString
    shellCmd.Append "avifdec.exe "
    
    'Use all available cores for decoding
    shellCmd.Append "-j "
    shellCmd.Append Trim$(Str$(OS.LogicalCoreCount())) & " "
    
    'In April 2022 a new version of libavif finally dropped, meaning I can *finally* request uncompressed PNGs
    ' (see https://github.com/AOMediaCodec/libavif/issues/706 for my feature request on this point)
    If (GetVersion(True) <> "0.9.0") Then shellCmd.Append "--png-compress 0 "
    
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
        
        If (outputPDIF = PDIF_PNG) Then
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

Public Function ConvertStandardImageToAVIF(ByRef srcFile As String, ByRef dstFile As String, Optional ByVal encoderQuality As Long = -1, Optional ByVal encoderSpeed As Long = -1) As Boolean
    
    Const funcName As String = "ConvertStandardImageToAVIF"
    
    'Safety checks on plugin
    If (Not m_avifExportAvailable) Then
        InternalError funcName, "libavif broken or missing"
        Exit Function
    End If
    
    Dim pluginPath As String
    pluginPath = PluginManager.GetPluginPath & "avifenc.exe"
    If (Not Files.FileExists(pluginPath)) Then
        InternalError funcName, "libavif missing"
        Exit Function
    End If
    
    'Safety checks on source and destination files
    If (Not Files.FileExists(srcFile)) Then
        InternalError funcName, "source file doesn't exist"
        Exit Function
    End If
    
    'Start constructing the full shell string
    Dim shellCmd As pdString
    Set shellCmd = New pdString
    shellCmd.Append "avifenc.exe "
    
    'Assign encoding thread count (one per core seems reasonable for initial testing)
    shellCmd.Append "-j "
    shellCmd.Append Trim$(Str$(OS.LogicalCoreCount())) & " "
    
    'Lossless encoding is its own parameter, and note that it supercedes a bunch of other parameters
    ' (because lossless encoding has unique constraints)
    Dim useLossless As Boolean
    useLossless = (encoderQuality = 0)
    
    If useLossless Then
        shellCmd.Append "-l "
    
    'Lossless encoding provides much more granular control over a billion different settings
    Else
        
        'Encoder speed can now be specified; default is 6 (per ./avifenc.exe -h).  Lower = slower.
        ' Negative values indicate "use the current avifenc default".
        If (encoderSpeed >= 0) Then
            If (encoderSpeed > 10) Then encoderSpeed = 10
            shellCmd.Append "--speed " & CStr(encoderSpeed) & " "
        End If
            
        'To simplify the UI, we don't expose min/max quality values (which are used by the encoder
        ' as part of a variable bit-rate approach to encoding).  Instead, we automatically generate
        ' a maximum quality value based on the user-supplied value (which is treated as a minimum
        ' target, where libavif quality=0=lossless ).  This makes the quality process somewhat more
        ' analogous to how otherformats (e.g. JPEG) do it.
        If (encoderQuality >= 0) Then
            If (encoderQuality > 63) Then encoderQuality = 63
            
            shellCmd.Append "--min " & CStr(encoderQuality) & " "
            
            'Treat 0 as lossless; anything else as variable quality
            Dim maxQuality As Long
            maxQuality = encoderQuality
            If (encoderQuality > 0) Then maxQuality = maxQuality + 10
            If (maxQuality > 63) Then maxQuality = 63
            shellCmd.Append "--max " & CStr(maxQuality) & " "
            
        End If
        
    End If
    
    'PD uses premultiplied alpha internally, so signal that to the encoder as well
    'shellCmd.Append "--premultiply "
    
    'Append properly delimited source image
    shellCmd.Append """"
    shellCmd.Append srcFile
    shellCmd.Append """ "
    
    'Append properly delimited destination image
    shellCmd.Append """"
    shellCmd.Append dstFile
    shellCmd.Append """"
    
    'Final step - if destination file exists, kill it.
    ' (TODO: convert to safe save approach)
    Files.FileDeleteIfExists dstFile
    
    Debug.Print shellCmd.ToString()
    
    'Shell plugin and capture output for analysis
    Dim outputString As String
    If ShellExecuteCapture(pluginPath, shellCmd.ToString(), outputString) Then
    
        'Shell appears successful.  The output string will have two easy-to-check flags if
        ' the conversion was successful.  Don't return success unless we find both.
        Dim targetStringSrc As String, targetStringDst As String
        targetStringSrc = "Successfully loaded: " & srcFile
        targetStringDst = "Wrote AVIF: " & dstFile
        
        ConvertStandardImageToAVIF = (Strings.StrStrBM(outputString, targetStringSrc, 1, True) > 0)
        ConvertStandardImageToAVIF = ConvertStandardImageToAVIF And (Strings.StrStrBM(outputString, targetStringDst, 1, True) > 0)
        
        'Want to review the output string manually?  Print it here:
        PDDebug.LogAction outputString
        
        'Record full details of failures
        If ConvertStandardImageToAVIF Then
            PDDebug.LogAction "libavif reports success!"
        Else
            InternalError funcName, "save failed; output follows:"
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

'Notify the user that PD can automatically download and configure AVIF support for them.
'
'Returns TRUE if PD successfully downloaded (and initialized) all required plugins
Public Function PromptForLibraryDownload(Optional ByVal targetIsImportLib As Boolean = True) As Boolean
    
    On Error GoTo BadDownload
    
    'Only attempt download if the current Windows install is 64-bit
    If OS.OSSupports64bitExe() Then
    
        'Ask the user for permission
        Dim uiMsg As pdString
        Set uiMsg = New pdString
        uiMsg.AppendLine g_Language.TranslateMessage("AVIF is a modern image format developed by the Alliance for Open Media.  PhotoDemon does not natively support AVIF images, but it can download a free, open-source plugin that permanently enables AVIF support.")
        uiMsg.AppendLineBreak
        uiMsg.AppendLine g_Language.TranslateMessage("The Alliance for Open Media provides free, open-source 64-bit AVIF encoder and decoder libraries.  These libraries are roughly ~%1 mb each (~%2 mb total).  Once downloaded, they will allow PhotoDemon to import and export AVIF files on any 64-bit system.", 20, 40)
        uiMsg.AppendLineBreak
        uiMsg.Append g_Language.TranslateMessage("Would you like PhotoDemon to download these libraries to your PhotoDemon plugin folder?")
        
        Dim msgReturn As VbMsgBoxResult
        msgReturn = PDMsgBox(uiMsg.ToString, vbInformation Or vbYesNoCancel, "Download required")
        If (msgReturn <> vbYes) Then
            
            'On a NO response, provide additional feedback.
            If (msgReturn = vbNo) Then
                uiMsg.Reset
                uiMsg.AppendLine g_Language.TranslateMessage("PhotoDemon will not download the AVIF libraries at this time.")
                uiMsg.AppendLineBreak
                uiMsg.AppendLine g_Language.TranslateMessage("To manually enable AVIF support, you can download the latest copies of the free ""%1"" and ""%2"" programs and place them into your PhotoDemon plugin folder:", "avifdec.exe", "avifenc.exe")
                uiMsg.AppendLine PluginManager.GetPluginPath()
                uiMsg.AppendLineBreak
                uiMsg.AppendLine g_Language.TranslateMessage("These free libraries are always available at the Alliance for Open Media libavif release page:")
                uiMsg.Append "https://github.com/AOMediaCodec/libavif/releases"
                PDMsgBox uiMsg.ToString, vbInformation Or vbOKOnly, "Download canceled"
            End If
            
            PromptForLibraryDownload = False
            Exit Function
            
        End If
        
        'The user said YES!  Attempt to download the latest libavif release now.
        Dim srcURL As String, dstFileDecoder As String
        
        'Before downloading anything, ensure we have write access on the plugin folder.
        dstFileDecoder = PluginManager.GetPluginPath()
        If Not Files.PathExists(dstFileDecoder, True) Then
            PDMsgBox g_Language.TranslateMessage("You have placed PhotoDemon in a restricted system folder.  Because PhotoDemon does not have administrator access, it cannot download files for you.  Please move PhotoDemon to an unrestricted folder and try again."), vbOKOnly Or vbApplicationModal Or vbCritical, g_Language.TranslateMessage("Error")
            PromptForLibraryDownload = False
            Exit Function
        End If
        
        'We need to download both the import and export library.  Steps are the same for both.
        
        'Start with import.
        srcURL = "https://github.com/tannerhelland/PhotoDemon-Updates-v2/releases/download/libavif-plugins-9.0-alpha.1/avifdec.exe"
        dstFileDecoder = PluginManager.GetPluginPath() & "avifdec.exe"
        
        'If the destination file does exist, kill it (maybe it's broken or bad)
        Files.FileDeleteIfExists dstFileDecoder
        
        'Download
        Dim tmpFile As String
        tmpFile = Web.DownloadURLToTempFile(srcURL, False)
        
        If Files.FileExists(tmpFile) Then Files.FileCopyW tmpFile, dstFileDecoder
        Files.FileDeleteIfExists tmpFile
        
        'Repeat for the encoder
        Dim dstFileEncoder As String
        srcURL = "https://github.com/tannerhelland/PhotoDemon-Updates-v2/releases/download/libavif-plugins-9.0-alpha.1/avifenc.exe"
        dstFileEncoder = PluginManager.GetPluginPath() & "avifenc.exe"
        Files.FileDeleteIfExists dstFileEncoder
        
        tmpFile = vbNullString
        tmpFile = Web.DownloadURLToTempFile(srcURL, False)
        
        If Files.FileExists(tmpFile) Then Files.FileCopyW tmpFile, dstFileEncoder
        Files.FileDeleteIfExists tmpFile
        
        'Attempt to initialize both plugins
        PluginManager.LoadPluginGroup False
        
        If targetIsImportLib Then
            PromptForLibraryDownload = PluginManager.IsPluginCurrentlyEnabled(CCP_AvifImport)
        Else
            PromptForLibraryDownload = PluginManager.IsPluginCurrentlyEnabled(CCP_AvifExport)
        End If
        
    End If
    
    Exit Function
    
BadDownload:
    PromptForLibraryDownload = False
    Exit Function

End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String)
    PDDebug.LogAction "WARNING! libavif error reported in " & funcName & "(): " & errDescription
End Sub
