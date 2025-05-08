Attribute VB_Name = "Plugin_DDS"
'***************************************************************************
'DirectXTex (DDS) Interface
'Copyright 2025-2025 by Tanner Helland
'Created: 28/April/25
'Last updated: 08/May/25
'Last update: use texdiag.exe (from DirectXTex) to pull relevant DDS attributes prior to import
'
'Module for handling all DirectXTex interfacing (via texconv.exe).  This module is pointless without
' that exe, which needs to be placed in the App/PhotoDemon/Plugins subdirectory.
'
'DirectXTex is a free, open-source, Microsoft-sponsored interface for DDS (DirectDraw Surface) texture files.
' You can learn more about it here:
'
' https://github.com/microsoft/DirectXTex
'
'PhotoDemon was designed against the October 2024 release, which is the last release to support Win 7.
' It may also work with newer (or older) versions.  You can also run the exe file manually with the -h
' extension for (extensive) details on how it works.
'
'DirectXTex is available under an MIT license.  Please see the App/PhotoDemon/Plugins/DirectXTex-LICENSE.txt
' file for questions regarding copyright or licensing.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Because DirectXTex ships x64 builds by default, we limit DDS support to 64-bit OS versions.
Private m_DirectXTexAvailable As Boolean

'Version number is only retrieved once, then cached.
Private m_LibFullPath As String, m_LibVersion As String

'PD ships with the additional "texdiag.exe" command-line app, which we can use to query some DDS info
' (and set up better texconv.exe flags during import).
Private m_pathToTexDiag As String

'Convert a DDS file to some other image format.  Currently, PD converts DDS files to PNGs,
' then imports those PNGs directly.  Theoretically, you could use other intermediary formats,
' such as JPEG, if that's better for your usage scenario...
Public Function ConvertDDStoStandardImage(ByRef srcFile As String, ByRef dstFile As String, Optional ByVal allowErrorPopups As Boolean = False) As Boolean
    
    Const funcName As String = "ConvertDDStoStandardImage"
    
    'Safety checks on plugin existence
    If (Not m_DirectXTexAvailable) Then
        InternalError funcName, "directxtex broken or missing"
        Exit Function
    End If
    
    Dim pluginPath As String
    pluginPath = PluginManager.GetPluginPath & "texconv.exe"
    If (Not Files.FileExists(pluginPath)) Then
        InternalError funcName, "directxtex missing"
        Exit Function
    End If
    
    'Safety checks on source file
    If (Not Files.FileExists(srcFile)) Then
        InternalError funcName, "source file doesn't exist"
        Exit Function
    End If
    
    'Pull basic attributes from the DDS file.
    ' (We can use these to assemble a better request for texdiag.)
    Dim cAttributes As pdDictionary, attributesExist As Boolean
    attributesExist = GetDDSAttributes(srcFile, cAttributes)
    
    'Ensure destination file has an appropriate extension (this is how the decoder knows which format to use)
    Dim outputPDIF As PD_IMAGE_FORMAT
    outputPDIF = PDIF_PNG
    
    'Note that we *can't* define a destination file.  DirectXTex automatically generates the destination file for us
    ' (using the target extension).
    Dim reqExtension As String
    reqExtension = "png"
    dstFile = Files.FileGetName(srcFile, True) & "." & reqExtension
    
    'TODO: what if the destination file already exists???
    
    'Next we want to shell the plugin and wait for return.
    
    'Start by assembling the shell command.
    Dim shellCmd As pdString
    Set shellCmd = New pdString
    shellCmd.Append "texconv.exe "
    
    'Use 8-bit RGBA PNG output
    shellCmd.Append "--file-type " & reqExtension & " "
    
    'Use relaxed permissions (this salvages some weird DDS variants)
    shellCmd.Append "--permissive "
    
    'Additional options can improve support for legacy formats
    shellCmd.Append "--expand-luminance "
    
    'Sometimes sRGB output is ideal; sometimes it isn't.  To use it correctly, we need to match it
    ' to images that actually contain sRGB data.
    Dim srcIsSRGB As Boolean
    If attributesExist And (Not cAttributes Is Nothing) Then
        Dim srcColorFormat As String
        srcColorFormat = cAttributes.GetEntry_String("format", vbNullString)
        srcIsSRGB = Strings.StringsEqualRight(srcColorFormat, "srgb", True)
        If srcIsSRGB Then shellCmd.Append "--srgb-out "
    End If
    
    'Be explicit about output format (this allows expansiong of e.g. RG to RGBA)
    If srcIsSRGB Then
        shellCmd.Append "--format R8G8B8A8_UNORM_SRGB "
    Else
        shellCmd.Append "--format R8G8B8A8_UNORM "
    End If
    
    'Tone-mapping might be appropriate for high-bit-depth images?
    'shellCmd.Append "--tonemap "
    
    'Multipage (doesn't work with PNG, could possibly try with TIFF)
    'shellCmd.Append "--wic-multiframe "
    
    'Append space-safe source image
    shellCmd.Append """"
    shellCmd.Append srcFile
    shellCmd.Append """ "
    
    'Shell plugin and capture output for analysis
    Dim cShell As pdPipeSync
    Set cShell = New pdPipeSync
    If cShell.RunAndCaptureOutput(pluginPath, shellCmd.ToString(), False) Then
        
        Dim outputString As String
        outputString = cShell.GetStdOutDataAsString()
        
        'Shell appears successful.  If output fails, the output string will contain a line that starts with "ERROR:"
        ConvertDDStoStandardImage = (Strings.StrStrBM(outputString, "ERROR:", 1, True) = 0) And (Strings.StrStrBM(outputString, "FAILED (", 1, True) = 0)
        
        'Want to review the output string manually?  Print it here:
        'PDDebug.LogAction outputString
        
        'Record full details of failures
        If ConvertDDStoStandardImage Then
            PDDebug.LogAction "directxtex reports success; transferring image to internal parser..."
            PDDebug.LogAction outputString
            
        'Conversion failed
        Else
            
            InternalError funcName, "load failed; output follows:"
            PDDebug.LogAction outputString
            
            'texconv.exe doesn't use stderr, to my knowledge
            'Dim srcStdErr As String
            'srcStdErr = cShell.GetStdErrDataAsString()
            'PDDebug.LogAction "For reference, here's stderr: " & srcStdErr
            
            'However, if the destination file exists, we can probably just load it?
            If Files.FileExists(dstFile) Then
                PDDebug.LogAction "Destination file still exists, so we'll allow the caller to use it."
                ConvertDDStoStandardImage = True
            Else
            
                'Store any problems in the central plugin error tracker
                PluginManager.NotifyPluginError CCP_DirectXTex, outputString, Files.FileGetName(srcFile, False)
                
            End If
            
            'You can also report errors directly to the user here:
            'If (Macros.GetMacroStatus <> MacroBATCH) And allowErrorPopups Then PluginManager.GenericLibraryError CCP_DirectXTex, cShell.GetStdErrDataAsString()
            
        End If
        
    Else
        InternalError funcName, "shell failed"
    End If
    
End Function

Public Function ConvertStandardImageToDDS(ByRef srcFile As String, ByRef dstFile As String) As Boolean
    
    Const FUNC_NAME As String = "ConvertStandardImageToDDS"
    
    'Safety checks on plugin
    If (Not m_DirectXTexAvailable) Then
        InternalError FUNC_NAME, "directxtex broken or missing"
        Exit Function
    End If
    
    Dim pluginPath As String
    pluginPath = PluginManager.GetPluginPath & "texconv.exe"
    If (Not Files.FileExists(pluginPath)) Then
        InternalError FUNC_NAME, "directxtex missing"
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
    shellCmd.Append "texconv.exe "
'
'    'Assign encoding thread count (one per core seems reasonable for initial testing)
'    shellCmd.Append "-j "
'    shellCmd.Append Trim$(Str$(OS.LogicalCoreCount())) & " "
'
'    'Lossless encoding is its own parameter, and note that it supercedes a bunch of other parameters
'    ' (because lossless encoding has unique constraints)
'    Dim useLossless As Boolean
'    useLossless = (encoderQuality = 0)
'
'    If useLossless Then
'        shellCmd.Append "-l "
'
'    'Lossless encoding provides much more granular control over a billion different settings
'    Else
'
'        'Encoder speed can now be specified; default is 6 (per ./avifenc.exe -h).  Lower = slower.
'        ' Negative values indicate "use the current avifenc default".
'        If (encoderSpeed >= 0) Then
'            If (encoderSpeed > 10) Then encoderSpeed = 10
'            shellCmd.Append "--speed " & CStr(encoderSpeed) & " "
'        End If
'
'        'To simplify the UI, we don't expose min/max quality values (which are used by the encoder
'        ' as part of a variable bit-rate approach to encoding).  Instead, we automatically generate
'        ' a maximum quality value based on the user-supplied value (which is treated as a minimum
'        ' target, where libavif quality=0=lossless ).  This makes the quality process somewhat more
'        ' analogous to how otherformats (e.g. JPEG) do it.
'        If (encoderQuality >= 0) Then
'            If (encoderQuality > 63) Then encoderQuality = 63
'
'            shellCmd.Append "--min " & CStr(encoderQuality) & " "
'
'            'Treat 0 as lossless; anything else as variable quality
'            Dim maxQuality As Long
'            maxQuality = encoderQuality
'            If (encoderQuality > 0) Then maxQuality = maxQuality + 10
'            If (maxQuality > 63) Then maxQuality = 63
'            shellCmd.Append "--max " & CStr(maxQuality) & " "
'
'        End If
'
'    End If
'
'    'PD uses premultiplied alpha internally, so signal that to the encoder as well.
'    ' (NOTE: libavif hasn't always handled premultiplication correctly, so suspend this for now and revisit in future builds.)
'    'shellCmd.Append "--premultiply "
    
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
        targetStringDst = "Wrote DDS: "
        
        ConvertStandardImageToDDS = (Strings.StrStrBM(outputString, targetStringSrc, 1, True) > 0)
        ConvertStandardImageToDDS = ConvertStandardImageToDDS And (Strings.StrStrBM(outputString, targetStringDst, 1, True) > 0)
        
        'Want to review the output string manually?  Print it here:
        'PDDebug.LogAction outputString
        
        'Record full details of failures
        If ConvertStandardImageToDDS Then
            PDDebug.LogAction "directxtex reports success!"
        Else
            InternalError FUNC_NAME, "save failed; output follows:"
            PDDebug.LogAction outputString
        End If
        
    Else
        InternalError FUNC_NAME, "shell failed"
    End If
    
    'If the original file already existed, attempt to replace it now
    If ConvertStandardImageToDDS And Strings.StringsNotEqual(dstFile, tmpFilename) Then
        ConvertStandardImageToDDS = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
        If (Not ConvertStandardImageToDDS) Then
            Files.FileDelete tmpFilename
            PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
        End If
    End If
    
End Function

'Forcibly disable plugininteractions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_DirectXTexAvailable = newState
End Sub

Public Function GetVersion() As String
    
    'Version string is cached on first access
    If (LenB(m_LibVersion) <> 0) Then
        GetVersion = m_LibVersion
    Else
        
        Const FUNC_NAME As String = "GetVersion"
        
        Dim cFSO As pdFSO
        Set cFSO = New pdFSO
        If cFSO.FileExists(m_LibFullPath) Then
            cFSO.FileGetVersionAsString m_LibFullPath, m_LibVersion
        End If
        
        If (LenB(m_LibVersion) = 0) Then
            InternalError FUNC_NAME, "couldn't retrieve version"
            m_LibVersion = "unknown"
        End If
        
        GetVersion = m_LibVersion
        
    End If
    
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean
    
    'Before doing anything else, make sure the OS supports 64-bit apps.
    ' (PD ships a 64-bit binary of texconv.exe)
    If (Not OS.OSSupports64bitExe()) Or (Not OS.IsWin7OrLater) Then
        m_DirectXTexAvailable = False
        InitializeEngine = False
        PDDebug.LogAction "WARNING!  DDS support not available; system is only 32-bit"
        Exit Function
    End If
    
    m_LibFullPath = pathToDLLFolder & "texconv.exe"
    m_DirectXTexAvailable = Files.FileExists(m_LibFullPath)
    InitializeEngine = m_DirectXTexAvailable
    
    'While here, see if we also have access to additional DDS support libraries
    m_pathToTexDiag = pathToDLLFolder & "texdiag.exe"
    If (Not Files.FileExists(m_pathToTexDiag)) Then m_pathToTexDiag = vbNullString
    
    If (Not InitializeEngine) Then
        PDDebug.LogAction "WARNING!  DDS support not available; plugins missing"
    End If
    
End Function

Public Function IsDirectXTexAvailable() As Boolean
    IsDirectXTexAvailable = m_DirectXTexAvailable
End Function

'Quick file-header check to see if a file is likely DDS
Public Function IsFilePotentiallyDDS(ByRef srcFile As String) As Boolean
    
    IsFilePotentiallyDDS = False
    
    'DDS files have a clear ID in the first 4 bytes; check those to see if it's worth offering
    ' a full load via DirectXTex.
    If Files.FileExists(srcFile) Then
        
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
            
            'First 4-bytes should be the magic number "DDS "
            Dim strID As String
            strID = cStream.ReadString_ASCII(4)
            cStream.StopStream
            IsFilePotentiallyDDS = Strings.StringsEqualLeft(strID, "DDS ", True)
            
        End If
        
    End If
    
End Function

Public Function QuickLoadPotentialDDSToDIB(ByRef srcFile As String, ByRef dstDIB As pdDIB, Optional ByRef tmpPDImage As pdImage = Nothing) As Boolean
    
    If Plugin_DDS.IsDirectXTexAvailable() Then
        
        'The separate AVIF apps convert AVIF to intermediary formats; we use PNG currently
        Dim tmpFile As String
        QuickLoadPotentialDDSToDIB = Plugin_DDS.ConvertDDStoStandardImage(srcFile, tmpFile, False)
        
        If QuickLoadPotentialDDSToDIB Then
            Dim cPNG As pdPNG
            Set cPNG = New pdPNG
            If tmpPDImage Is Nothing Then Set tmpPDImage = New pdImage
            QuickLoadPotentialDDSToDIB = (cPNG.LoadPNG_Simple(tmpFile, tmpPDImage, dstDIB) < png_Failure)
            Set cPNG = Nothing
        End If
        
        'Free the intermediary file before continuing
        Files.FileDeleteIfExists tmpFile
        If (Not dstDIB.GetAlphaPremultiplication) Then dstDIB.SetAlphaPremultiplication True
        
    End If
    
End Function

'Use texdiag.exe to retrieve basic DDS file attributes.  Returns TRUE if attributes were received successfully.
Private Function GetDDSAttributes(ByRef srcFile As String, ByRef dstAttributes As pdDictionary) As Boolean
    
    GetDDSAttributes = False
    Set dstAttributes = New pdDictionary
    
    'Failsafe check for the helper app
    If (LenB(m_pathToTexDiag) <> 0) Then
        If (Not Files.FileExists(srcFile)) Then Exit Function
    Else
        Exit Function
    End If
    
    'Run the helper app and poll stdout
    Dim shellCmd As pdString
    Set shellCmd = New pdString
    shellCmd.Append "texconv.exe info "
    shellCmd.Append """" & srcFile & """"
    
    Dim cShell As pdPipeSync
    Set cShell = New pdPipeSync
    If cShell.RunAndCaptureOutput(m_pathToTexDiag, shellCmd.ToString(), False) Then
        
        Dim outputString As String
        outputString = cShell.GetStdOutDataAsString()
        
        'Look for the filename + the text "FAILED" in the output
        If (InStr(1, outputString, Files.FileGetName(srcFile, False) & " FAILED", vbTextCompare) = 0) Then
            
            'Split the output string into lines
            Dim cLines As pdStringStack
            Set cLines = New pdStringStack
            If cLines.CreateFromMultilineString(outputString) Then
                
                'If we've made it this far, we (probably?) have valid attributes for the target file
                GetDDSAttributes = True
                
                'We now want to parse each line for key+value pairs.
                Dim srcLine As String
                Do While cLines.PopString(srcLine)
                    
                    Const EQUAL_SIGN As String = "="
                    Dim eqPos As Long
                    eqPos = InStr(1, srcLine, EQUAL_SIGN, vbBinaryCompare)
                    If (eqPos > 0) Then
                        
                        Dim sKey As String, sValue As String
                        sKey = Trim$(Left$(srcLine, eqPos - 1))
                        sValue = Trim$(Right$(srcLine, Len(srcLine) - eqPos))
                        dstAttributes.AddEntry sKey, sValue
                        
                        'To review attributes as they're parsed, use this:
                        'pdDebug.LogAction sKey & ":" & sValue
                        
                    End If
                    
                Loop
                
            End If
            
        End If
        
    End If
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String)
    PDDebug.LogAction "WARNING! DirectXTex error reported in " & funcName & "(): " & errDescription
End Sub
