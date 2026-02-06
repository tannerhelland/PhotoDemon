Attribute VB_Name = "Plugin_AVIF"
'***************************************************************************
'libavif Interface
'Copyright 2021-2026 by Tanner Helland
'Created: 13/July/21
'Last updated: 11/March/25
'Last update: update to the latest libavif (1.2.0)
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
'PhotoDemon has been designed against v1.2.0 (Feb 2025).  It may not work with other versions.
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

'Version numbers are only retrieved once, then cached.  (We need to check version numbers before
' communicating with libavif, because some optional settings only work on specific library versions.)
Private m_inputVersion As String, m_outputVersion As String

'Convert an AVIF file to some other image format.  Currently, PD converts AVIF files to uncompressed PNGs,
' then imports those PNGs directly.  Theoretically, you could use other intermediary formats, such as JPEG,
' if that's better for your usage scenario...
Public Function ConvertAVIFtoStandardImage(ByRef srcFile As String, ByRef dstFile As String, Optional ByVal allowErrorPopups As Boolean = False) As Boolean
    
    Const funcName As String = "ConvertAVIFtoStandardImage"
    
    'Safety checks on plugin existence
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
    
    'Ensure destination file has an appropriate extension (this is how the decoder knows which format to use)
    Dim outputPDIF As PD_IMAGE_FORMAT
    outputPDIF = PDIF_PNG
    
    Dim reqExtension As String
    reqExtension = "png"
    
    If Strings.StringsNotEqual(Files.FileGetExtension(dstFile), reqExtension, True) Then dstFile = dstFile & "." & reqExtension
    
    'Shell plugin and wait for return
    Dim shellCmd As pdString
    Set shellCmd = New pdString
    shellCmd.Append "avifdec.exe "
    
    'Use all available cores for decoding
    shellCmd.Append "-j all "
    
    'Use 8-bit PNG output (16-bit is also available; in the future, this may be a worthwhile switch for incoming HDR images)
    shellCmd.Append "-d 8 "
    
    'In April 2022 a new version of libavif finally dropped, meaning I can *finally* request uncompressed PNGs
    ' (see https://github.com/AOMediaCodec/libavif/issues/706 for my feature request on this point)
    If (GetVersion(True) <> "0.9.0") Then shellCmd.Append "--png-compress 0 "
    
    'To improve compatibility with more AVIF decoders, in 2025 I explicitly request non-strict mode
    ' (without this, many AVIF files "in the wild" won't work).
    shellCmd.Append "--no-strict "
    
    'Explicitly mark the end of options
    shellCmd.Append " -- "
    
    'Append space-safe source image
    shellCmd.Append """"
    shellCmd.Append srcFile
    shellCmd.Append """ "
    
    'Append space-safe destination image
    shellCmd.Append """"
    shellCmd.Append dstFile
    shellCmd.Append """"
    
    'Shell plugin and capture output for analysis
    Dim cShell As pdPipeSync
    Set cShell = New pdPipeSync
    
    If cShell.RunAndCaptureOutput(pluginPath, shellCmd.ToString(), False) Then
        
        Dim outputString As String
        outputString = cShell.GetStdOutDataAsString()
        
        'Shell appears successful.  The output string will have two easy-to-check flags if
        ' the conversion was successful.  Don't return success unless we find both.
        Dim targetStringSrc As String, targetStringDst As String
        targetStringSrc = "Image decoded: "
        
        If (outputPDIF = PDIF_PNG) Then
            targetStringDst = "Wrote PNG: "
        Else
            targetStringDst = "Wrote JPEG: "
        End If
        
        ConvertAVIFtoStandardImage = (Strings.StrStrBM(outputString, targetStringSrc, 1, True) > 0)
        ConvertAVIFtoStandardImage = ConvertAVIFtoStandardImage And (Strings.StrStrBM(outputString, targetStringDst, 1, True) > 0)
        
        'Want to review the output string manually?  Print it here:
        'PDDebug.LogAction outputString
        
        'Record full details of failures
        If ConvertAVIFtoStandardImage Then
            PDDebug.LogAction "libavif reports success; transferring image to internal parser..."
        
        'Conversion failed
        Else
            
            InternalError funcName, "load failed; output follows:"
            PDDebug.LogAction outputString
            
            Dim avifStdErr As String
            avifStdErr = cShell.GetStdErrDataAsString()
            PDDebug.LogAction "For reference, here's stderr: " & avifStdErr
            
            'Store any problems in the central plugin error tracker
            PluginManager.NotifyPluginError CCP_libavif, avifStdErr, Files.FileGetName(srcFile, False)
            
        End If
        
    Else
        InternalError funcName, "shell failed"
    End If
    
End Function

Public Function ConvertStandardImageToAVIF(ByRef srcFile As String, ByRef dstFile As String, Optional ByVal encoderQuality As Long = -1, Optional ByVal encoderSpeed As Long = -1) As Boolean
    
    Const FUNC_NAME As String = "ConvertStandardImageToAVIF"
    
    'Safety checks on plugin
    If (Not m_avifExportAvailable) Then
        InternalError FUNC_NAME, "libavif broken or missing"
        Exit Function
    End If
    
    Dim pluginPath As String
    pluginPath = PluginManager.GetPluginPath & "avifenc.exe"
    If (Not Files.FileExists(pluginPath)) Then
        InternalError FUNC_NAME, "libavif missing"
        Exit Function
    End If
    
    'Safety checks on source and destination files
    If (Not Files.FileExists(srcFile)) Then
        InternalError FUNC_NAME, "source file doesn't exist"
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
    
    'PD uses premultiplied alpha internally, so signal that to the encoder as well.
    ' (NOTE: libavif hasn't always handled premultiplication correctly, so suspend this for now and revisit in future builds.)
    'shellCmd.Append "--premultiply "
    
    'Append properly delimited source image
    shellCmd.Append """"
    shellCmd.Append srcFile
    shellCmd.Append """ "
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
    ' the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    'Append properly delimited destination image
    shellCmd.Append """"
    shellCmd.Append tmpFilename
    shellCmd.Append """"
    
    'Want to confirm the shelled command?  See it here:
    'PDDebug.LogAction shellCmd.ToString()
    
    'We are guaranteed that the destination file does not exist, but in case the user somehow (miraculously?)
    ' created a file with that name in the past 0.01 ms, guarantee non-existence.
    Files.FileDeleteIfExists tmpFilename
    
    'Shell plugin and capture output for analysis
    Dim cShell As pdPipeSync
    Set cShell = New pdPipeSync
    
    If cShell.RunAndCaptureOutput(pluginPath, shellCmd.ToString(), False) Then
        
        Dim outputString As String
        outputString = cShell.GetStdOutDataAsString()
    
        'Shell appears successful.  The output string will have two easy-to-check flags if
        ' the conversion was successful.  Don't return success unless we find both.
        Dim targetStringSrc As String, targetStringDst As String
        targetStringSrc = "Successfully loaded: "
        targetStringDst = "Wrote AVIF: "
        
        ConvertStandardImageToAVIF = (Strings.StrStrBM(outputString, targetStringSrc, 1, True) > 0)
        ConvertStandardImageToAVIF = ConvertStandardImageToAVIF And (Strings.StrStrBM(outputString, targetStringDst, 1, True) > 0)
        
        'Want to review the output string manually?  Print it here:
        'PDDebug.LogAction outputString
        
        'Record full details of failures
        If ConvertStandardImageToAVIF Then
            PDDebug.LogAction "libavif reports success!"
        Else
            InternalError FUNC_NAME, "save failed; output follows:"
            PDDebug.LogAction outputString
        End If
        
    Else
        InternalError FUNC_NAME, "shell failed"
    End If
    
    'If the original file already existed, attempt to replace it now
    If ConvertStandardImageToAVIF And Strings.StringsNotEqual(dstFile, tmpFilename) Then
        ConvertStandardImageToAVIF = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
        If (Not ConvertStandardImageToAVIF) Then
            Files.FileDelete tmpFilename
            PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
        End If
    End If
    
End Function

Public Function GetVersion(ByVal testExportLibrary As Boolean) As String
    
    'These libraries are limited to Vista+ and 64-bit OSes only
    GetVersion = vbNullString
    If (Not OS.IsVistaOrLater) Then Exit Function
    
    'The version-checker may have already been called this session;
    ' use cached values from a previous run, if available.
    If testExportLibrary Then
        If (LenB(m_outputVersion) > 0) Then
            GetVersion = m_outputVersion
            Exit Function
        End If
    Else
        If (LenB(m_inputVersion) > 0) Then
            GetVersion = m_inputVersion
            Exit Function
        End If
    End If
    
    Dim targetAvifAppName As String
    If testExportLibrary Then
        targetAvifAppName = "avifenc.exe"
    Else
        targetAvifAppName = "avifdec.exe"
    End If
    
    Dim pluginPath As String
    pluginPath = PluginManager.GetPluginPath & targetAvifAppName
    
    Dim okToCheck As Boolean
    okToCheck = Files.FileExists(PluginManager.GetPluginPath & targetAvifAppName)
    
    If okToCheck Then
        
        Dim cShell As pdPipeSync
        Set cShell = New pdPipeSync
        
        If cShell.RunAndCaptureOutput(pluginPath, targetAvifAppName & " -v", False) Then
            
            Dim outputString As String
            outputString = cShell.GetStdOutDataAsString()
            
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
                
                'Cache version number between calls
                If testExportLibrary Then
                    m_outputVersion = GetVersion
                Else
                    m_inputVersion = GetVersion
                End If
                
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

'Quick file-header check to see if a file is likely AVIF
Public Function IsFilePotentiallyAVIF(ByRef srcFile As String) As Boolean
    
    IsFilePotentiallyAVIF = False
    
    'AVIF files have a few potential IDs in the first 8 bytes; check those to see if it's worth offering
    ' a full load to libavif (initializing that library requires a full download of libavif).
    If Files.FileExists(srcFile) Then
        
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
            
            'Normally we would just read the first 8-bytes as a string, but some AVIF files have whitespace
            ' padding at the start, so we need to grab extra and then look for known AVIF headers.
            Dim strID As String
            strID = cStream.ReadString_ASCII(16)
            cStream.StopStream
            
            '(By design, this list includes some HEIC/HEIF headers because they can contain embedded AVIF data.)
            IsFilePotentiallyAVIF = InStr(1, strID, "ftypavif", vbBinaryCompare)
            IsFilePotentiallyAVIF = IsFilePotentiallyAVIF Or InStr(1, strID, "ftypavis", vbBinaryCompare)
            IsFilePotentiallyAVIF = IsFilePotentiallyAVIF Or InStr(1, strID, "ftypmif1", vbBinaryCompare)
            IsFilePotentiallyAVIF = IsFilePotentiallyAVIF Or InStr(1, strID, "ftypmsf1", vbBinaryCompare)
            IsFilePotentiallyAVIF = IsFilePotentiallyAVIF Or InStr(1, strID, "ftypheic", vbBinaryCompare)
            
        End If
        
    End If
    
End Function

Public Function QuickLoadPotentialAVIFToDIB(ByRef srcFile As String, ByRef dstDIB As pdDIB, Optional ByRef tmpPDImage As pdImage = Nothing) As Boolean
    
    If Plugin_AVIF.IsAVIFImportAvailable() Then
        
        'The separate AVIF apps convert AVIF to intermediary formats; we use PNG currently
        Dim tmpFile As String
        QuickLoadPotentialAVIFToDIB = Plugin_AVIF.ConvertAVIFtoStandardImage(srcFile, tmpFile, False)
        
        If QuickLoadPotentialAVIFToDIB Then
            Dim cPNG As pdPNG
            Set cPNG = New pdPNG
            If tmpPDImage Is Nothing Then Set tmpPDImage = New pdImage
            QuickLoadPotentialAVIFToDIB = (cPNG.LoadPNG_Simple(tmpFile, tmpPDImage, dstDIB) < png_Failure)
            Set cPNG = Nothing
        End If
        
        'Free the intermediary file before continuing
        Files.FileDeleteIfExists tmpFile
        If (Not dstDIB.GetAlphaPremultiplication) Then dstDIB.SetAlphaPremultiplication True
        
    End If
    
End Function

'Returns TRUE if the installed version of libavif is >= the expected version of libavif.
' By design, this function also returns TRUE if libavif is NOT installed - this is purposeful because
' I don't want to raise "library out of date" warnings if the library doesn't even exist (there's a
' separate code pathway for downloading the library for the first time).
Public Function CheckAVIFVersionAndOfferUpdates(Optional ByVal targetIsImportLib As Boolean = True) As Boolean
    
    'By design, this function returns TRUE if libavif doesn't exist.
    Dim libavifNotInstalled As Boolean
    If targetIsImportLib Then
        libavifNotInstalled = (Not IsAVIFImportAvailable)
    Else
        libavifNotInstalled = (Not IsAVIFExportAvailable)
    End If
    
    If libavifNotInstalled Then
        CheckAVIFVersionAndOfferUpdates = True
        Exit Function
    End If
    
    'Still here?  libavif exists in this install.  Let's pull its version and compare it to the expected version
    ' (for this build of PhotoDemon).
    Dim curVersion As String
    curVersion = Plugin_AVIF.GetVersion(targetIsImportLib)
    
    Dim expectedVersion As String
    expectedVersion = PluginManager.ExpectedPluginVersion(CCP_libavif)
    
    If Updates.IsNewVersionHigher(curVersion, expectedVersion) Then
        
        'The installed copy of libavif is out-of-date.  Offer to download a new copy.
        CheckAVIFVersionAndOfferUpdates = False
        
        Dim okToDownload As VbMsgBoxResult
        okToDownload = Updates.OfferPluginUpdate("libavif", curVersion, expectedVersion)
        If (okToDownload = vbYes) Then CheckAVIFVersionAndOfferUpdates = DownloadLatestLibAVIF()
        
    Else
        CheckAVIFVersionAndOfferUpdates = True
    End If
    
End Function

'Notify the user that PD can automatically download and configure AVIF support for them.
'
'Returns TRUE if PD successfully downloaded (and initialized) all required plugins
Public Function PromptForLibraryDownload_AVIF(Optional ByVal targetIsImportLib As Boolean = True) As Boolean
    
    On Error GoTo BadDownload
    
    'Only attempt download if the current Windows install is 64-bit
    If OS.OSSupports64bitExe() Then
    
        'Ask the user for permission
        Dim uiMsg As pdString
        Set uiMsg = New pdString
        uiMsg.AppendLine g_Language.TranslateMessage("AVIF is a modern image format developed by the Alliance for Open Media.  PhotoDemon does not natively support AVIF images, but it can download a free, open-source plugin that permanently enables AVIF support.")
        uiMsg.AppendLineBreak
        uiMsg.AppendLine g_Language.TranslateMessage("The Alliance for Open Media provides free, open-source 64-bit AVIF encoder and decoder libraries.  These libraries are roughly ~%1 mb each (~%2 mb total).  Once downloaded, they will allow PhotoDemon to import and export AVIF files on any 64-bit system.", 12, 24)
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
            
            PromptForLibraryDownload_AVIF = False
            Exit Function
            
        End If
        
        'The user said YES!  Attempt to download the latest libavif release now.
        PromptForLibraryDownload_AVIF = DownloadLatestLibAVIF()
        
    End If
    
    Exit Function
    
BadDownload:
    PromptForLibraryDownload_AVIF = False
    Exit Function

End Function

'Attempt to download the latest libavif copy to this PC.
Private Function DownloadLatestLibAVIF() As Boolean
    
    ' Currently we expect these files in the package:
    ' - avifdec.exe (for decoding)
    ' - avifenc.exe (for encoding)
    ' - avif-LICENSE.txt (copyright and license info)
    Const EXPECTED_NUM_FILES As Long = 3
    
    'Current libavif build is 1.2.0, downloaded from https://github.com/AOMediaCodec/libavif/releases/
    Const EXPECTED_TOTAL_EXTRACT_SIZE As Long = 24539456
    Const UPDATE_URL As String = "https://github.com/tannerhelland/PhotoDemon-Updates-v2/releases/download/libavif-plugins-1.2.0/libavif-1.2.0.pdz"
    DownloadLatestLibAVIF = Updates.DownloadPluginUpdate(CCP_libavif, UPDATE_URL, EXPECTED_NUM_FILES, EXPECTED_TOTAL_EXTRACT_SIZE)
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String)
    PDDebug.LogAction "WARNING! libavif error reported in " & funcName & "(): " & errDescription
End Sub
