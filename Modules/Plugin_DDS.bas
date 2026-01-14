Attribute VB_Name = "Plugin_DDS"
'***************************************************************************
'DirectXTex (DDS) Interface
'Copyright 2025-2026 by Tanner Helland
'Created: 28/April/25
'Last updated: 22/May/25
'Last update: continued workarounds for lack of output file parameters in texconv
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
'Thank you also to Nicholas Hayes (https://github.com/0xC0000054), author of Paint.NET's DDS plugin
' (https://github.com/0xC0000054/pdn-ddsfiletype-plus), who first pointed me to DirectXTex and whose work
' export dialog features I shamelessly copied when building PhotoDemon's DDS export dialog.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'For verbose debug output, set this to TRUE.
' LEAVE AS FALSE IN PRODUCTION BUILDS.
Private Const DDS_DEBUG_VERBOSE As Boolean = False

'Because DirectXTex ships x64 builds by default, we limit DDS support to 64-bit OS versions.
Private m_DirectXTexAvailable As Boolean

'Version number is only retrieved once, then cached.
Private m_LibFullPath As String, m_LibVersion As String

'PD ships with the additional "texdiag.exe" command-line app, which we can use to query some DDS info
' (and set up better texconv.exe flags during import).
Private m_pathToTexDiag As String

'Convert a DDS file to some other image format.  Currently, PD converts DDS files to PNGs,
' then imports those PNGs directly.
'
'IMPORTANT: unlike other conversion functions, this function does not accept a destination filename.
' Instead, it *populates* that param for you.  (DirectXTex doesn't support variable output filename(s),
' so we have to use workarounds involving a temp folder.)
Public Function ConvertDDStoStandardImage(ByRef srcFile As String, ByRef dstFile As String, Optional ByVal allowErrorPopups As Boolean = False) As Boolean
    
    Const funcName As String = "ConvertDDStoStandardImage"
    ConvertDDStoStandardImage = False
    
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
    
    'Now we are (mostly) confident that we have everything we need to load the DDS file correctly.
    
    'Pull basic attributes from the source DDS file.
    ' (We can use these to assemble a better request for texdiag.)
    Dim cAttributes As pdDictionary, attributesExist As Boolean
    attributesExist = GetDDSAttributes(srcFile, cAttributes)
    
    'Ensure destination file has an appropriate extension (this is how the decoder knows which format to use)
    Dim outputPDIF As PD_IMAGE_FORMAT
    outputPDIF = PDIF_PNG
    
    'Note that we *can't* define a destination file.  DirectXTex automatically generates the destination file for us
    ' (using the target extension).
    '
    'Further testing shows that DirectXTex doesn't accept an output folder correctly via command-line if the output
    ' folder contains a space character.  We can't control this in PD as the user is allowed to use whatever folder
    ' they want for temp processing.
    '
    'What we *can* do as a stupid workaround is to use a text file with a list of files to be processed.
    ' Filenames inside the text file *are* processed correctly, even if they have spaces in their filenames.
    '
    'There is still the problem of filenames that exist as both filename.dds and filename.png *in the same folder*.
    ' As we don't want to blindly overwrite filename.png if it exists, we instead need to make of a copy of
    ' the source DDS file to the user's temp folder first.
    '
    'So the order of operations goes like this:
    ' 1) copy the source file to the user's temp folder
    ' 2) create a text file in the temp folder that consists of only the temporary source file copy
    ' 3) pass that text file to DirectXTex and let it work
    ' 4) load the PNG file from the temp folder (created by DirectXTex)
    ' 5) clean-up all the temp files we created
    
    'Start with copying the source file to the user's temp folder.
    Dim dstTmpFileDDS As String
    dstTmpFileDDS = OS.UniqueTempFilename(customExtension:="dds")
    If (Not Files.FileCopyW(srcFile, dstTmpFileDDS)) Then
        InternalError funcName, "couldn't copy DDS to temp folder"
        Exit Function
    End If
    
    'Next, create a temporary text file in the temp folder, and write the name of the source DDS file to it.
    Dim tmpTxtFile As String
    tmpTxtFile = OS.UniqueTempFilename(customExtension:="txt")
    If (Not Files.FileSaveAsText(dstTmpFileDDS, tmpTxtFile, True, False)) Then
        InternalError funcName, "couldn't save input text file"
        Exit Function
    End If
    
    'We now have all the input files we need.  Next we need to figure out where DirectXTex is going
    ' to put the converted PNG file.
    Dim reqExtension As String
    reqExtension = "png"
    dstFile = Files.FileGetPath(dstTmpFileDDS) & Files.FileGetName(dstTmpFileDDS, True) & "." & reqExtension
    If DDS_DEBUG_VERBOSE Then
        PDDebug.LogAction "Source DDS file: " & srcFile
        PDDebug.LogAction "Temp DDS file (in temp folder): " & dstTmpFileDDS
        PDDebug.LogAction "Temp txt file (for texconv input): " & tmpTxtFile
        PDDebug.LogAction "Expected destination PNG file: " & dstFile
    End If
    
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
    
    'Overwriting target file is OK for this usage
    shellCmd.Append "--overwrite "
    
    'Tone-mapping might be appropriate for high-bit-depth images?
    'shellCmd.Append "--tonemap "
    
    'Multipage (doesn't work with PNG, could possibly try with TIFF)
    'shellCmd.Append "--wic-multiframe "
    
    'Append the text file for input
    shellCmd.Append "--file-list "
    shellCmd.Append """"
    shellCmd.Append tmpTxtFile
    shellCmd.Append """ "
    
    'Force writing to the temp folder, and - this is important - intentionally omit the trailing comma.
    ' This deliberate decision is actually workaround for faulty path parsing in texconv.exe.
    ' (If you add the trailing comma, it gets included as part of the path text!)
    shellCmd.Append "-o """
    shellCmd.Append Files.FileGetPath(dstTmpFileDDS)
    'shellCmd.Append ""
    
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
            
            If DDS_DEBUG_VERBOSE Then
                PDDebug.LogAction "directxtex reports success; transferring image to internal parser..."
                PDDebug.LogAction outputString
            End If
            
            'Clean-up all the intermediary files
            Files.FileDeleteIfExists dstTmpFileDDS
            Files.FileDeleteIfExists tmpTxtFile
            
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
            
        End If
        
    Else
        InternalError funcName, "shell failed"
    End If
    
End Function

'Convert a "standard" image file (typically PNG) to DDS format.
Public Function ConvertStandardImageToDDS(ByRef srcFile As String, ByRef dstFile As String, Optional ByRef dxTex_FormatID As String = vbNullString, Optional ByVal dxTex_numMipMaps As Long = 0&, Optional ByVal dxTex_mmFilter As String = vbNullString, Optional ByVal dxTex_BC As String = vbNullString) As Boolean
    
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
    
    'Note that we *can't* define a destination file.  DirectXTex automatically generates the destination file for us
    ' (using the target extension "DDS").
    '
    'Further testing shows that DirectXTex doesn't accept an output folder correctly via command-line if the output
    ' folder contains a space character.  We can't control this in PD as the user is allowed to use whatever folder
    ' they want for temp processing.
    '
    'What we *can* do as a stupid workaround is to use a text file with a list of files to be processed.
    ' Filenames inside the text file *are* processed correctly, even if they have spaces in their filenames.
    '
    'There is still the problem of filenames that exist as both filename.dds and filename.png *in the same folder*.
    ' As we don't want to blindly overwrite filename.dds if it exists, we instead need to make of a copy of
    ' the source PNG file to the user's temp folder first.
    '
    'So the order of operations goes like this:
    ' 1) copy the source PNG file to the user's temp folder
    ' 2) create a text file in the temp folder that consists of only the temporary source file copy
    ' 3) pass that text file to DirectXTex and let it work
    ' 4) clean-up all the temp files we created
    
    'Start with copying the source PNG file to the user's temp folder.
    Dim dstTmpFilePNG As String
    dstTmpFilePNG = OS.UniqueTempFilename(customExtension:="png")
    If (Not Files.FileCopyW(srcFile, dstTmpFilePNG)) Then
        InternalError FUNC_NAME, "couldn't copy PNG to temp folder"
        Exit Function
    End If
    
    'Next, create a temporary text file in the temp folder, and write the name of the source DDS file to it.
    Dim tmpTxtFile As String
    tmpTxtFile = OS.UniqueTempFilename(customExtension:="txt")
    If (Not Files.FileSaveAsText(dstTmpFilePNG, tmpTxtFile, True, False)) Then
        InternalError FUNC_NAME, "couldn't save input text file"
        Exit Function
    End If
    
    'We now have all the input files we need.  Next we need to figure out where DirectXTex is going
    ' to put the converted DDS file.
    Dim reqExtension As String
    reqExtension = "dds"
    dstFile = Files.FileGetPath(dstTmpFilePNG) & Files.FileGetName(dstTmpFilePNG, True) & "." & reqExtension
    If DDS_DEBUG_VERBOSE Then
        PDDebug.LogAction "Source PNG file: " & srcFile
        PDDebug.LogAction "Temp PNG file (in temp folder): " & dstTmpFilePNG
        PDDebug.LogAction "Temp txt file (for texconv input): " & tmpTxtFile
        PDDebug.LogAction "Expected destination DDS file: " & dstFile
    End If
    
    'Start constructing the full shell string
    Dim shellCmd As pdString
    Set shellCmd = New pdString
    shellCmd.Append "texconv.exe "
    
    'Use DDS output (obviously) in whatever format the caller specified
    shellCmd.Append "--file-type DDS "
    
    shellCmd.Append "--format "
    If (LenB(dxTex_FormatID) > 0) Then
        If DDS_DEBUG_VERBOSE Then PDDebug.LogAction "Using DDS format ID: " & dxTex_FormatID
        shellCmd.Append dxTex_FormatID
    Else
        shellCmd.Append "R8G8B8A8_UNORM"
    End If
    shellCmd.Append " "
    
    'Some block-compression algorithms support additional settings; it's up to the caller to validate these
    If (LenB(dxTex_BC) <> 0) Then
        shellCmd.Append "--block-compress "
        shellCmd.Append dxTex_BC
        shellCmd.Append " "
    End If
    
    'Mipmaps default to "all" (a value of zero), but the user can also request a specific amount.
    If (dxTex_numMipMaps <> 0) Then
        shellCmd.Append "-m "
        shellCmd.Append Trim$(Str$(dxTex_numMipMaps))
        shellCmd.Append " "
    End If
    
    'Filter only applies if mipmaps are being generated (e.g. if # of mipmaps > 1)
    If (dxTex_numMipMaps <> 1) And (LenB(dxTex_mmFilter) <> 0) Then
        shellCmd.Append "--image-filter "
        shellCmd.Append UCase$(dxTex_mmFilter)
        shellCmd.Append " "
    End If
    
    'For sRGB output, mark the *incoming* image as also being sRGB (to prevent unwanted auto-adjustment from linear to sRGB)
    If Strings.StringsEqualRight(dxTex_FormatID, "srgb", True) Then shellCmd.Append "--srgb-in "
    
    'Mipmaps?  Scaling?  DirectX version for header?
    
    'Overwrite destination file is OK
    shellCmd.Append "--overwrite "
    
    'Append the text file for input
    shellCmd.Append "--file-list "
    shellCmd.Append """"
    shellCmd.Append tmpTxtFile
    shellCmd.Append """ "
    
    'Force writing to the temp folder, and - this is important - intentionally omit the trailing comma.
    ' This deliberate decision is actually workaround for faulty path parsing in texconv.exe.
    ' (If you add the trailing comma, it gets included as part of the path text!)
    shellCmd.Append "-o """
    shellCmd.Append Files.FileGetPath(dstTmpFilePNG)
    'shellCmd.Append ""
    
    'Shell plugin and capture output for analysis
    Dim cShell As pdPipeSync
    Set cShell = New pdPipeSync
    If cShell.RunAndCaptureOutput(pluginPath, shellCmd.ToString(), False) Then
        
        Dim outputString As String
        outputString = cShell.GetStdOutDataAsString()
    
        'Shell appears successful. Ensure the destination file exists.
        ConvertStandardImageToDDS = Files.FileExists(dstFile)
        
        'Erase the temporary source image copy and text file we generated
        Files.FileDeleteIfExists dstTmpFilePNG
        Files.FileDeleteIfExists tmpTxtFile
        
        'Record full details of failures
        If (Not ConvertStandardImageToDDS) Then
            InternalError FUNC_NAME, "save failed; output follows:"
            PDDebug.LogAction outputString
        End If
        
    Else
        InternalError FUNC_NAME, "shell failed"
        PDDebug.LogAction "FYI, shell string follows: "
        PDDebug.LogAction shellCmd.ToString()
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

'User-friendly names and DirectXTex-specific IDs for each compression option.
' (You can pull a full list of supported IDs by running "./texconv.exe -h";
'  by design, PD doesn't expose all possible destination formats.)
' Returned INT is the number of items (1-based) added to each list, and it is guaranteed identical for both lists.
Public Function GetListOfFormatNamesAndIDs(ByRef dstNames As pdStringStack, ByRef dstIDs As pdStringStack) As Long

    If (dstNames Is Nothing) Then Set dstNames = New pdStringStack Else dstNames.ResetStack
    If (dstIDs Is Nothing) Then Set dstIDs = New pdStringStack Else dstIDs.ResetStack
    
    dstIDs.AddString "BC1_UNORM"
    dstNames.AddString "BC1 (Linear, DXT1)"

    dstIDs.AddString "BC1_UNORM_SRGB"
    dstNames.AddString "BC1 (sRGB, DX 10+)"

    dstIDs.AddString "BC2_UNORM"
    dstNames.AddString "BC2 (Linear, DXT3)"

    dstIDs.AddString "BC2_UNORM_SRGB"
    dstNames.AddString "BC2 (sRGB, DX 10+)"

    dstIDs.AddString "BC3_UNORM"
    dstNames.AddString "BC3 (Linear, DXT5)"
    
    dstIDs.AddString "BC3_UNORM_SRGB"
    dstNames.AddString "BC3 (sRGB, DX 10+)"

    dstIDs.AddString "BC4_SNORM"
    dstNames.AddString "BC4 (Linear, Signed)"

    dstIDs.AddString "BC4_UNORM"
    dstNames.AddString "BC4 (Linear, Unsigned)"
    
    dstIDs.AddString "BC5_SNORM"
    dstNames.AddString "BC5 (Linear, Signed)"

    dstIDs.AddString "BC5_UNORM"
    dstNames.AddString "BC5 (Linear, Unsigned)"

    dstIDs.AddString "BC6H_SF16"
    dstNames.AddString "BC6H (Linear, Signed, DX 11+)"
    
    dstIDs.AddString "BC6H_UF16"
    dstNames.AddString "BC6H (Linear, Unsigned, DX 11+)"
    
    dstIDs.AddString "BC7_UNORM"
    dstNames.AddString "BC7 (Linear, DX 11+)"

    dstIDs.AddString "BC7_UNORM_SRGB"
    dstNames.AddString "BC7 (sRGB, DX 11+)"
    
    dstIDs.AddString "B8G8R8A8_UNORM"
    dstNames.AddString "B8G8R8A8 (Linear, A8R8G8B8)"

    dstIDs.AddString "B8G8R8A8_UNORM_SRGB"
    dstNames.AddString "B8G8R8A8 (sRGB, DX 10+)"

    dstIDs.AddString "B8G8R8X8_UNORM"
    dstNames.AddString "B8G8R8X8 (Linear, X8R8G8B8)"

    dstIDs.AddString "B8G8R8X8_UNORM_SRGB"
    dstNames.AddString "B8G8R8X8 (sRGB, DX 10+)"
    
    dstIDs.AddString "R8G8B8A8_UNORM"
    dstNames.AddString "R8G8B8A8 (Linear, A8B8G8R8)"

    dstIDs.AddString "R8G8B8A8_UNORM_SRGB"
    dstNames.AddString "R8G8B8A8 (sRGB, DX 10+)"

    dstIDs.AddString "B4G4R4A4_UNORM"
    dstNames.AddString "B4G4R4A4 (Linear, A4R4G4B4)"

    dstIDs.AddString "B5G5R5A1_UNORM"
    dstNames.AddString "B5G5R5A1 (Linear, A1R5G5B5)"

    dstIDs.AddString "B5G6R5_UNORM"
    dstNames.AddString "B5G6R5 (Linear, R5G6B5)"

    dstIDs.AddString "R8_UNORM"
    dstNames.AddString "R8 (Linear, Unsigned, L8)"
    
    dstIDs.AddString "R8G8_SNORM"
    dstNames.AddString "R8G8 (Linear, Signed, V8U8)"

    dstIDs.AddString "R16_FLOAT"
    dstNames.AddString "R16 (Linear, Float)"

    dstIDs.AddString "R32_FLOAT"
    dstNames.AddString "R32 (Linear, Float)"

    GetListOfFormatNamesAndIDs = dstIDs.GetNumOfStrings()
    
End Function

Public Function DoesFormatSupportAlpha(ByRef srcFormatId As String) As Boolean
    
    DoesFormatSupportAlpha = True
    
    Select Case srcFormatId
        Case "BC1_UNORM"
            DoesFormatSupportAlpha = True
        Case "BC1_UNORM_SRGB"
            DoesFormatSupportAlpha = True
        Case "BC2_UNORM"
            DoesFormatSupportAlpha = True
        Case "BC2_UNORM_SRGB"
            DoesFormatSupportAlpha = True
        Case "BC3_UNORM"
            DoesFormatSupportAlpha = True
        Case "BC3_UNORM_SRGB"
            DoesFormatSupportAlpha = True
        Case "BC4_SNORM"
            DoesFormatSupportAlpha = False
        Case "BC4_UNORM"
            DoesFormatSupportAlpha = False
        Case "BC5_SNORM"
            DoesFormatSupportAlpha = False
        Case "BC5_UNORM"
            DoesFormatSupportAlpha = False
        Case "BC6H_SF16"
            DoesFormatSupportAlpha = False
        Case "BC6H_UF16"
            DoesFormatSupportAlpha = False
        Case "BC7_UNORM"
            DoesFormatSupportAlpha = True
        Case "BC7_UNORM_SRGB"
            DoesFormatSupportAlpha = True
        Case "B8G8R8A8_UNORM"
            DoesFormatSupportAlpha = True
        Case "B8G8R8A8_UNORM_SRGB"
            DoesFormatSupportAlpha = True
        Case "B8G8R8X8_UNORM"
            DoesFormatSupportAlpha = False
        Case "B8G8R8X8_UNORM_SRGB"
            DoesFormatSupportAlpha = False
        Case "R8G8B8A8_UNORM"
            DoesFormatSupportAlpha = True
        Case "R8G8B8A8_UNORM_SRGB"
            DoesFormatSupportAlpha = True
        Case "B4G4R4A4_UNORM"
            DoesFormatSupportAlpha = True
        Case "B5G5R5A1_UNORM"
            DoesFormatSupportAlpha = True
        Case "B5G6R5_UNORM"
            DoesFormatSupportAlpha = False
        Case "R8_UNORM"
            DoesFormatSupportAlpha = False
        Case "R8G8_SNORM"
            DoesFormatSupportAlpha = False
        Case "R16_FLOAT"
            DoesFormatSupportAlpha = False
        Case "R32_FLOAT"
            DoesFormatSupportAlpha = False
    End Select
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String)
    PDDebug.LogAction "WARNING! DirectXTex error reported in " & funcName & "(): " & errDescription
End Sub
