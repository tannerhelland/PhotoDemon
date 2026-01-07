Attribute VB_Name = "Plugin_jxl"
'***************************************************************************
'JPEG-XL Reference Library (libjxl) Interface
'Copyright 2022-2026 by Tanner Helland
'Created: 28/September/22
'Last updated: 21/February/24
'Last update: offer automatic updates if the installed copy of libjxl lags the latest officially supported version
'
'libjxl (available at https://github.com/libjxl/libjxl) is the official reference library implementation
' for the modern JPEG-XL format.  Support for this format was added during the PhotoDemon 10.0 release cycle.
'
'I initially tried to work directly with libjxl as a library, but ongoing stability issues and a very complex
' build process eventually led me to switch to interfacing with libjxl via separate apps (cjxl/djxl.exe).
' This module is pointless without those exes, which need to be placed in the App/PhotoDemon/Plugins subdirectory.
' (PD will automatically download these for you if you attempt to interact with JPEG XL files.)
'
'Unfortunately for Windows XP users, libjxl currently requires Windows Vista or later.  PhotoDemon will
' detect this automatically and gracefully hide JPEG XL support for XP users.
'
'PhotoDemon tries to support most JPEG XL features, but esoteric ones (like animation) remain a WIP.
' If you encounter any issues with JPEG XL images, please file an issue on GitHub and attach the image(s)
' in question so I can investigate further.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'DO NOT enable verbose logging in production builds
Private Const JXL_DEBUG_VERBOSE As Boolean = True

'Because libjxl is nearly impossible to build as a portable 32-bit library, we interface with its .exe builds.
' This means that decoding and encoding support exist separately (i.e. just because the import library exists
' at run-time, doesn't mean the export library also exists; users may only install one or none).
Private m_jxlImportAvailable As Boolean, m_jxlExportAvailable As Boolean

'Initialize the library.  Do not call this until you have verified its existence (typically via the PluginManager module)
Public Function InitializeLibJXL(ByRef pathToDLLFolder As String) As Boolean
    
    InitializeLibJXL = False
    m_jxlExportAvailable = False
    m_jxlImportAvailable = False
    
    'libjxl cannot currently be built in an XP-compatible way.
    ' As a result, its support is limited to Win Vista and above.
    If (Not OS.IsVistaOrLater) Then
        DebugMsg "libjxl does not currently work on Windows XP."
        Exit Function
    End If
    
    'Test import and export support separately
    Dim importPath As String, exportPath As String
    importPath = pathToDLLFolder & "djxl.exe"
    exportPath = pathToDLLFolder & "cjxl.exe"
    
    m_jxlExportAvailable = Files.FileExists(exportPath)
    m_jxlImportAvailable = Files.FileExists(importPath)
    
    'Both cjxl and djxl require a host of support files.
    Dim supportFilesOK As Boolean
    supportFilesOK = Files.FileExists(pathToDLLFolder & "jxlinfo.exe")
    
    m_jxlExportAvailable = m_jxlExportAvailable And supportFilesOK
    m_jxlImportAvailable = m_jxlImportAvailable And supportFilesOK
    
    InitializeLibJXL = m_jxlImportAvailable Or m_jxlExportAvailable
    
    If (Not InitializeLibJXL) Then
        PDDebug.LogAction "WARNING! JPEG XL support not available; plugins missing"
    End If
    
End Function

'Forcibly disable library interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_jxlExportAvailable = newState
    m_jxlImportAvailable = newState
End Sub

Public Function GetLibJXLVersion() As String
    
    Const FUNC_NAME As String = "GetLibJXLVersion"
    
    'Do not attempt to retrieve version info unless the library actually exists
    If Files.FileExists(PluginManager.GetPluginPath & "djxl.exe") And OS.IsVistaOrLater Then
        
        Dim pluginExeAndPath As String
        pluginExeAndPath = PluginManager.GetPluginPath() & "djxl.exe"
        If (Not Files.FileExists(pluginExeAndPath)) Then Exit Function
        
        'Start constructing a command-line string
        Dim shellCmd As pdString
        Set shellCmd = New pdString
        shellCmd.Append "djxl.exe --version"
        
        'Shell the JPEG XL decompressor and simply request its version info
        Dim cShell As pdPipeSync
        Set cShell = New pdPipeSync
        
        If cShell.RunAndCaptureOutput(pluginExeAndPath, shellCmd.ToString(), False) Then
            
            'libjxl writes to stderr for reasons unknown
            Dim outputString As String
            outputString = cShell.GetStdOutDataAsString()
            
            'Look for the library name first; version typically follows, as in:
            ' "djxl v0.8.1 c27d499 [SSE4,SSSE3,Unknown]"
            Dim libNamePos As Long
            libNamePos = InStr(1, outputString, "djxl", vbBinaryCompare)
            
            If (libNamePos > 0) Then
                
                'Advance pointer past the space that follows "djxl" (e.g. to the first char past "djxl v")
                libNamePos = libNamePos + 6
                
                Dim libNameEndPos As Long
                libNameEndPos = InStr(libNamePos, outputString, " ", vbBinaryCompare)
                
                If (libNameEndPos > libNamePos) Then
                    GetLibJXLVersion = Mid$(outputString, libNamePos, libNameEndPos - libNamePos)
                Else
                    InternalError FUNC_NAME, "bad version parse"
                End If
            
            Else
                InternalError FUNC_NAME, "bad lib name"
            End If
                
        Else
            InternalError FUNC_NAME, "failed to shell djxl with --version"
        End If
        
    End If
        
End Function

Public Function IsJXLExportAvailable() As Boolean
    IsJXLExportAvailable = m_jxlExportAvailable
End Function

Public Function IsJXLImportAvailable() As Boolean
    IsJXLImportAvailable = m_jxlImportAvailable
End Function

Public Function ConvertImageFileToJXL(ByRef srcFile As String, ByRef dstFile As String, Optional ByRef convertParams As String = vbNullString, Optional ByVal isLivePreview As Boolean = False) As Boolean

    Const FUNC_NAME As String = "ConvertImageFileToJXL"
    ConvertImageFileToJXL = False
    On Error GoTo ConvertFailed
    
    'Failsafe check
    If (Not Plugin_jxl.IsJXLExportAvailable()) Then Exit Function
    
    'Second failsafe check
    Dim pluginExeAndPath As String
    pluginExeAndPath = PluginManager.GetPluginPath() & "cjxl.exe"
    If (Not Files.FileExists(pluginExeAndPath)) Then Exit Function
    
    'Failsafe check on input
    If (Not Files.FileExists(srcFile)) Then Exit Function
    
    'Ensure the source filename includes a recognizable format; if it does not, libjxl will choke
    If (Files.FileGetExtension(srcFile) <> "png") And (Files.FileGetExtension(srcFile) <> "apng") And (Files.FileGetExtension(srcFile) <> "jpg") Then
        InternalError FUNC_NAME, "bad extension"
        Exit Function
    End If
    
    'Start by constructing a command-line string
    Dim shellCmd As pdString
    Set shellCmd = New pdString
    shellCmd.Append "cjxl.exe "
    
    'Input first (note the use of quotes to ensure safety with space-containing paths.)
    shellCmd.Append """"
    shellCmd.Append srcFile
    shellCmd.Append """ "
    
    'Output next
    shellCmd.Append """"
    shellCmd.Append dstFile
    shellCmd.Append """"
    
    'Retrieve parameters from incoming string.  Magic-number constants are taken directly from libjxl via "cjxl.exe -h"
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString convertParams
    
    Dim jxlParamLossless As Boolean, jxlParamQuality As Single, jxlParamEffort As Long
    jxlParamLossless = cParams.GetBool("jxl-lossless", False, True)
    jxlParamQuality = cParams.GetSingle("jxl-lossy-quality", 90!, True)
    jxlParamEffort = cParams.GetLong("jxl-effort", 7)
    
    'Sanity check inputs.  Again, magic-number constants are taken directly from libjxl via "cjxl.exe -h"
    If (jxlParamQuality < 0!) Then jxlParamQuality = 0!
    If (jxlParamQuality > 100!) Then jxlParamQuality = 100!
    
    If (jxlParamEffort < 1) Then jxlParamEffort = 1
    If (jxlParamEffort > 9) Then jxlParamEffort = 9
    
    'Append parameters to shell string
    shellCmd.Append " "
    If jxlParamLossless Then
        shellCmd.Append "-q 100"
    Else
        shellCmd.Append "-q " & Trim$(Str$(jxlParamQuality))
    End If
    
    shellCmd.Append " "
    If isLivePreview Then
        shellCmd.Append " -e 1"
    Else
        shellCmd.Append " -e " & Trim$(Str$(jxlParamEffort))
    End If
    
    If JXL_DEBUG_VERBOSE Then PDDebug.LogAction "Shelling libjxl..."
    PDDebug.LogAction shellCmd.ToString()
    
    'Shell plugin and capture output for analysis
    Dim outputString As String
    
    Dim cShell As pdPipeSync
    Set cShell = New pdPipeSync
    If cShell.RunAndCaptureOutput(pluginExeAndPath, shellCmd.ToString(), False) Then
        
        'For reasons I do not fathom, libjxl writes all state data to stderr only
        ' (including normal success reporting *facepalm*)
        outputString = cShell.GetStdErrDataAsString()
        If JXL_DEBUG_VERBOSE Then PDDebug.LogAction "cjxl.exe returned: " & outputString
        
        'On a successful export, a line should appear in the output like:
        ' "Compressed to 1234 bytes (0.123 bpp)."
        ConvertImageFileToJXL = (InStr(1, outputString, "Compressed to ", vbTextCompare) > 0)
    
    End If
    
    If (Not ConvertImageFileToJXL) And JXL_DEBUG_VERBOSE Then InternalError FUNC_NAME, "failed"
    Exit Function
    
ConvertFailed:
    ConvertImageFileToJXL = False
    InternalError FUNC_NAME, "terminating due to error"
    
End Function

'Caller assumes all responsibility for destination file being valid and writable
Public Function ConvertJXLtoImageFile(ByRef srcFile As String, ByRef dstFile As String) As Boolean

    Const FUNC_NAME As String = "ConvertJXLtoImageFile"
    ConvertJXLtoImageFile = False
    On Error GoTo ConvertFailed
    
    'Failsafe check
    If (Not Plugin_jxl.IsJXLImportAvailable()) Then Exit Function
    
    'Second failsafe check
    Dim pluginExeAndPath As String
    pluginExeAndPath = PluginManager.GetPluginPath() & "djxl.exe"
    If (Not Files.FileExists(pluginExeAndPath)) Then Exit Function
    
    'Next, we need to validate the file format as JPEG-XL.
    If Plugin_jxl.IsFileJXL(srcFile) Then
        
        If JXL_DEBUG_VERBOSE Then DebugMsg "JXL format found.  Proceeding with conversion..."
        
        'Ensure the destination filename includes a recognizable format; if it does not, libjxl will choke
        If (Files.FileGetExtension(dstFile) <> "png") And (Files.FileGetExtension(dstFile) <> "apng") And (Files.FileGetExtension(dstFile) <> "jpg") Then
            InternalError FUNC_NAME, "bad extension"
            Exit Function
        End If
        
        'Start by constructing a command-line string
        Dim shellCmd As pdString
        Set shellCmd = New pdString
        shellCmd.Append "djxl.exe "
        
        'Input first (note the use of quotes to ensure safety with space-containing paths.)
        shellCmd.Append """"
        shellCmd.Append srcFile
        shellCmd.Append """ "
        
        'Output next
        shellCmd.Append """"
        shellCmd.Append dstFile
        shellCmd.Append """"
        
        If JXL_DEBUG_VERBOSE Then PDDebug.LogAction "Shelling libjxl..."
        PDDebug.LogAction shellCmd.ToString()
        
        'Shell plugin and capture output for analysis
        Dim outputString As String
        
        Dim cShell As pdPipeSync
        Set cShell = New pdPipeSync
        
        If cShell.RunAndCaptureOutput(pluginExeAndPath, shellCmd.ToString(), False) Then
            
            'For reasons I do not fathom, libjxl writes all state data to stderr only
            ' (including normal success reporting *facepalm*)
            outputString = cShell.GetStdErrDataAsString()
            
            If JXL_DEBUG_VERBOSE Then
                PDDebug.LogAction "libjxl returned.  Analyzing output..."
                PDDebug.LogAction "(Output follows)" & vbCrLf & outputString
            End If
            
            'Look for success
            Const JXL_DECODER_SUCCESS_TEXT As String = "Decoded to pixels."
            ConvertJXLtoImageFile = (Strings.StrStrBM(outputString, JXL_DECODER_SUCCESS_TEXT, 1, False) > 0)
            
            If ConvertJXLtoImageFile Then
                If JXL_DEBUG_VERBOSE Then PDDebug.LogAction "libjxl returned success!"
            Else
                InternalError FUNC_NAME, "couldn't find success in output string?"
            End If
            
        'Plugin error
        Else
            InternalError FUNC_NAME, "failed to shell decoder (djxl.exe)"
            ConvertJXLtoImageFile = False
        End If
        
    '/File is not JXL format
    Else
        Exit Function
    End If
    
    If (Not ConvertJXLtoImageFile) And JXL_DEBUG_VERBOSE Then InternalError FUNC_NAME, "failed"
    Exit Function
    
ConvertFailed:
    ConvertJXLtoImageFile = False
    InternalError FUNC_NAME, "terminating due to error"
    
End Function

'Check if a file is JXL without relying on libjxl (only an initial magic number is validated; nothing else)
Public Function IsFileJXL_NoExternalLibrary(ByRef srcFile As String) As Boolean
    
    IsFileJXL_NoExternalLibrary = False
    If Files.FileExists(srcFile) Then
        
        'Pull the first 12 bytes only
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
            Dim bFirst12() As Byte
            If cStream.ReadBytes(bFirst12, 12, True) Then
                
                'Two different signatures are valid, based on the container used.
                Const JXL_MAGIC_NUMBER_1 As Integer = &HAFF
                IsFileJXL_NoExternalLibrary = VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(JXL_MAGIC_NUMBER_1), 2)
                If (Not IsFileJXL_NoExternalLibrary) Then
                    
                    Const JXL_MAGIC_NUMBER_2 As Long = &HC000000
                    Const JXL_MAGIC_NUMBER_3 As Long = &H204C584A
                    Const JXL_MAGIC_NUMBER_4 As Long = &HA870A0D
                    IsFileJXL_NoExternalLibrary = VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(JXL_MAGIC_NUMBER_2), 4)
                    If IsFileJXL_NoExternalLibrary Then IsFileJXL_NoExternalLibrary = VBHacks.MemCmp(VarPtr(bFirst12(4)), VarPtr(JXL_MAGIC_NUMBER_3), 4)
                    If IsFileJXL_NoExternalLibrary Then IsFileJXL_NoExternalLibrary = VBHacks.MemCmp(VarPtr(bFirst12(8)), VarPtr(JXL_MAGIC_NUMBER_4), 4)
                    
                End If
                
            End If
        End If
        
        Set cStream = Nothing
        
    End If
    
End Function

'Import/Export functions follow
Public Function IsFileJXL(ByRef srcFile As String) As Boolean
    
    IsFileJXL = False
    
    'Do an initial check that doesn't rely on libjxl
    If (Not IsFileJXL_NoExternalLibrary(srcFile)) Then Exit Function
    
    'A better check relies on libjxl, but if it's not available, rely on whatever we discovered in our "quick and dirty" check
    If (Not Plugin_jxl.IsJXLImportAvailable()) Then Exit Function
    
    'Second failsafe check for the separate JXL info executable
    Dim pluginExeAndPath As String
    pluginExeAndPath = PluginManager.GetPluginPath() & "jxlinfo.exe"
    If (Not Files.FileExists(pluginExeAndPath)) Then Exit Function
    
    'Start constructing a command-line string
    Dim shellCmd As pdString
    Set shellCmd = New pdString
    shellCmd.Append "jxlinfo.exe "
    
    'For basic format detection, all we need to append is the target filename.
    ' (Note the use of quotes to ensure safety with space-containing paths.)
    shellCmd.Append """"
    shellCmd.Append srcFile
    shellCmd.Append """"
    
    'Shell plugin and capture output for analysis
    Dim cShell As pdPipeSync
    Set cShell = New pdPipeSync
    
    If cShell.RunAndCaptureOutput(pluginExeAndPath, shellCmd.ToString(), False) Then
        
        Dim outputString As String
        outputString = cShell.GetStdOutDataAsString()
        
        'Look for potential decoder errors; if none are found, assume the file is worth investigating as JXL
        Const JXL_DECODER_ERROR_TEXT As String = "Decoder error"
        Const JXL_DECODER_ERROR_TEXT_ALT As String = "Error reading file"
        
        IsFileJXL = (Strings.StrStrBM(outputString, JXL_DECODER_ERROR_TEXT, 1, False) = 0)
        IsFileJXL = IsFileJXL And (Strings.StrStrBM(outputString, JXL_DECODER_ERROR_TEXT_ALT, 1, False) = 0)
        
    End If
    
    If JXL_DEBUG_VERBOSE And IsFileJXL Then
        PDDebug.LogAction "IsFileJXL returned " & UCase$(CStr(IsFileJXL)) & " for " & srcFile
    End If
    
End Function

'Load a JPEG XL file from disk.  srcFile must be a fully qualified path.  In the case of animated images,
' dstImage will be populated with all embedded frames, one frame per layer.
Public Function LoadJXL(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "LoadJXL"
    LoadJXL = False
    
    'Failsafe check
    If (Not Plugin_jxl.IsJXLImportAvailable()) Then Exit Function
    
    'Second failsafe check
    Dim pluginExeAndPath As String
    pluginExeAndPath = PluginManager.GetPluginPath() & "djxl.exe"
    If (Not Files.FileExists(pluginExeAndPath)) Then Exit Function
    
    'Next, we need to validate the file format as JPEG-XL.
    If Plugin_jxl.IsFileJXL(srcFile) Then
        
        If JXL_DEBUG_VERBOSE Then DebugMsg "JXL format found.  Proceeding with load..."
        If (dstImage Is Nothing) Then Set dstImage = New pdImage
        
        'Ask the decoder to convert the image to a temporary a/png file.
        ' (This intermediary format allows us to support animated JXL files too.)
        Dim tmpFilePNG As String
        tmpFilePNG = OS.UniqueTempFilename(customExtension:="png")
        
        'Start by constructing a command-line string
        Dim shellCmd As pdString
        Set shellCmd = New pdString
        shellCmd.Append "djxl.exe "
        
        'Input first (note the use of quotes to ensure safety with space-containing paths.)
        shellCmd.Append """"
        shellCmd.Append srcFile
        shellCmd.Append """ "
        
        'Output next
        shellCmd.Append """"
        shellCmd.Append tmpFilePNG
        shellCmd.Append """"
        
        If JXL_DEBUG_VERBOSE Then PDDebug.LogAction "Shelling libjxl..."
        PDDebug.LogAction shellCmd.ToString()
        
        'Shell plugin and capture output for analysis
        Dim outputString As String
        
        Dim cShell As pdPipeSync
        Set cShell = New pdPipeSync
        
        If cShell.RunAndCaptureOutput(pluginExeAndPath, shellCmd.ToString(), False) Then
            
            'For reasons I do not fathom, libjxl writes all state data to stderr only
            ' (including normal success reporting *facepalm*)
            outputString = cShell.GetStdErrDataAsString()
            
            If JXL_DEBUG_VERBOSE Then
                PDDebug.LogAction "libjxl returned.  Analyzing output..."
                PDDebug.LogAction "(Output follows)" & vbCrLf & outputString
            End If
            
            'Look for success
            Const JXL_DECODER_SUCCESS_TEXT As String = "Decoded to pixels."
            LoadJXL = (Strings.StrStrBM(outputString, JXL_DECODER_SUCCESS_TEXT, 1, False) > 0)
            
            If LoadJXL Then
                
                If JXL_DEBUG_VERBOSE Then PDDebug.LogAction "libjxl returned success!  Loading converted PNG..."
                
                'The image now exists as a standalong a/png file.  We can use our internal parser
                ' to rapidly(ish) decompress the image to internal PDI format.
                Dim cPNG As pdPNG
                Set cPNG = New pdPNG
                LoadJXL = (cPNG.LoadPNG_Simple(tmpFilePNG, dstImage, dstDIB, False) <= png_Warning)
                
                If LoadJXL Then
                
                    'If we've experienced one or more warnings during the load process, dump them out to the debug file.
                    If (cPNG.Warnings_GetCount() > 0) Then cPNG.Warnings_DumpToDebugger
                    
                    'Because color-management has already been handled (if applicable), this is a great time to premultiply alpha
                    dstDIB.SetAlphaPremultiplication True
                    
                    'Note the original file format as JXL (*not* PNG, which is relevant because we are using
                    ' PNG as an intermediary format and other load functions may mistakenly operate on PNG assumptions)
                    dstImage.SetOriginalFileFormat PDIF_JXL
                    
                    'Migrate the filled DIB into the destination image object, and initialize it as the base layer
                    Dim newLayerName As String
                    newLayerName = Layers.GenerateInitialLayerName(srcFile, vbNullString, cPNG.IsAnimated(), dstImage, dstDIB)
                    
                    'Create the new layer in the target image, and pass our created name to it
                    Dim newLayerID As Long
                    newLayerID = dstImage.CreateBlankLayer
                    dstImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, newLayerName, dstDIB, False
                    dstImage.UpdateSize
                    
                    'If the JXL converter wrote an animated PNG, import remaining frames now
                    If cPNG.IsAnimated() Then
                        
                        LoadJXL = (cPNG.ImportStage7_LoadRemainingFrames(dstImage) < png_Failure)
                        
                        'Hide all frames except the first
                        Dim pageTracker As Long
                        For pageTracker = 1 To dstImage.GetNumOfLayers - 1
                            dstImage.GetLayerByIndex(pageTracker).SetLayerVisibility False
                        Next pageTracker
                        
                        dstImage.SetActiveLayerByIndex 0
                        
                        'Also tag the image as being animated; we use this to activate some contextual UI bits
                        dstImage.SetAnimated True
                        
                    Else
                        dstImage.SetAnimated False
                    End If
                    
                    'Only *now* do we relay any useful state information to the destination image object.
                    ' (Note that these settings are PNG-specific, not JXL-specific, so e.g. a 12-bit JXL file
                    ' will use a 16-bit intermediary PNG - that's okay for our purposes!)
                    dstImage.SetOriginalColorDepth cPNG.GetBitsPerPixel()
                    dstImage.SetOriginalGrayscale (cPNG.GetColorType = png_Greyscale) Or (cPNG.GetColorType = png_GreyscaleAlpha)
                    dstImage.SetOriginalAlpha cPNG.HasAlpha()
                    
                    'We are now finished with the PNG interface; free it as it may be quite large
                    Set cPNG = Nothing
                    
                End If
                
                'Regardless of success or failure, kill the temporary PNG file (if it exists)
                Files.FileDeleteIfExists tmpFilePNG
                
            Else
                InternalError FUNC_NAME, "couldn't find success in output string?"
            End If
            
        'Plugin error
        Else
            InternalError FUNC_NAME, "failed to shell decoder (djxl.exe)"
            LoadJXL = False
        End If
        
    '/File is not JXL format
    Else
        Exit Function
    End If
    
    If (Not LoadJXL) And JXL_DEBUG_VERBOSE Then DebugMsg "Plugin_jxl.LoadJXL failed."
    Exit Function
    
LoadFailed:
    LoadJXL = False
    InternalError FUNC_NAME, "terminating due to error"
    
End Function

'Preview a single frame as compressed to JXL format, using the passed compression settings.
' This is typically used to generate previews in export dialogs.  Speed is emphasized wherever possible.
' (Per the name, do *not* pass an animated source file to this function!)
Public Function PreviewSingleFrameAsJXL(ByRef srcFile As String, ByRef dstDIB As pdDIB, ByRef srcOptions As String) As Boolean

    Const FUNC_NAME As String = "PreviewSingleFrameAsJXL "
    PreviewSingleFrameAsJXL = False
    On Error GoTo PreviewFailed
    
    'Failsafe check
    If (Not Plugin_jxl.IsJXLExportAvailable()) Then Exit Function
    
    'Second failsafe check
    Dim pluginExeAndPath As String
    pluginExeAndPath = PluginManager.GetPluginPath() & "cjxl.exe"
    If (Not Files.FileExists(pluginExeAndPath)) Then Exit Function
    
    'Failsafe check on input
    If (Not Files.FileExists(srcFile)) Then Exit Function
    
    'Ensure the source filename includes a recognizable format; if it does not, libjxl will choke
    If (Files.FileGetExtension(srcFile) <> "png") And (Files.FileGetExtension(srcFile) <> "apng") And (Files.FileGetExtension(srcFile) <> "jpg") Then
        InternalError FUNC_NAME, "bad extension"
        Exit Function
    End If
    
    'Start by constructing a command-line string
    Dim shellCmd As pdString
    Set shellCmd = New pdString
    shellCmd.Append "cjxl.exe "
    
    'Input first (note the use of quotes to ensure safety with space-containing paths.)
    shellCmd.Append """"
    shellCmd.Append srcFile
    shellCmd.Append """ "
    
    'Output next
    Dim dstFile As String
    dstFile = OS.UniqueTempFilename(customExtension:="jxl")
    
    shellCmd.Append """"
    shellCmd.Append dstFile
    shellCmd.Append """"
    
    'Retrieve parameters from incoming string.  Magic-number constants are taken directly from libjxl via "cjxl.exe -h"
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString srcOptions
    
    Dim jxlParamLossless As Boolean, jxlParamQuality As Single, jxlParamEffort As Long
    jxlParamLossless = cParams.GetBool("jxl-lossless", False, True)
    jxlParamQuality = cParams.GetSingle("jxl-lossy-quality", 90!, True)
    
    'Normally, we would pull effort like this...
    'jxlParamEffort = cParams.GetLong("jxl-effort", 7)
    '...but for previews, we want minimal effort to improve speed
    jxlParamEffort = 1
    
    'Sanity check inputs.  Again, magic-number constants are taken directly from libjxl via "cjxl.exe -h"
    If (jxlParamQuality < 0!) Then jxlParamQuality = 0!
    If (jxlParamQuality > 100!) Then jxlParamQuality = 100!
    
    If (jxlParamEffort < 1) Then jxlParamEffort = 1
    If (jxlParamEffort > 9) Then jxlParamEffort = 9
    
    'Append parameters to shell string
    shellCmd.Append " "
    If jxlParamLossless Then
        shellCmd.Append "-q 100"
    Else
        shellCmd.Append "-q " & Trim$(Str$(jxlParamQuality))
    End If
    
    shellCmd.Append " "
    shellCmd.Append " -e " & Trim$(Str$(jxlParamEffort))
    
    If JXL_DEBUG_VERBOSE Then PDDebug.LogAction "Shelling libjxl..."
    PDDebug.LogAction shellCmd.ToString()
    
    'Shell plugin and capture output for analysis
    Dim outputString As String
    
    Dim cShell As pdPipeSync
    Set cShell = New pdPipeSync
    If cShell.RunAndCaptureOutput(pluginExeAndPath, shellCmd.ToString(), False) Then
        
        'For reasons I do not fathom, libjxl writes all state data to stderr only
        ' (including normal success reporting *facepalm*)
        outputString = cShell.GetStdErrDataAsString()
        If JXL_DEBUG_VERBOSE Then PDDebug.LogAction "cjxl.exe returned: " & outputString
        
        'On a successful export, a line should appear in the output like:
        ' "Compressed to 1234 bytes (0.123 bpp)."
        PreviewSingleFrameAsJXL = (InStr(1, outputString, "Compressed to ", vbTextCompare) > 0)
    
    End If
    
    If (Not PreviewSingleFrameAsJXL) And JXL_DEBUG_VERBOSE Then
        InternalError FUNC_NAME, "failed to generate jxl file"
        Exit Function
    End If
    
    'If we're still here, we now have a JXL file with the compression settings applied.
    ' We now need to convert that file to some other standardized format (currently PNG)
    Dim tmpFilename As String
    tmpFilename = OS.UniqueTempFilename(customExtension:="png")
    PreviewSingleFrameAsJXL = ConvertJXLtoImageFile(dstFile, tmpFilename)
    
    'Hopefully that worked...
    If (Not PreviewSingleFrameAsJXL) And JXL_DEBUG_VERBOSE Then
        InternalError FUNC_NAME, "failed to decode jxl file"
        Exit Function
    End If
    
    'Free the temporary JXL file
    Files.FileDeleteIfExists dstFile
    
    'Load the final PNG file to a pdDIB object
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB Else dstDIB.ResetDIB 0
    PreviewSingleFrameAsJXL = Loading.QuickLoadImageToDIB(tmpFilename, dstDIB, False, False)
    Files.FileDeleteIfExists tmpFilename
    
    Exit Function
    
PreviewFailed:
    PreviewSingleFrameAsJXL = False
    InternalError FUNC_NAME, "terminating due to error"
    
End Function

'Save an arbitrary pdImage object to a standalone JPEG XL file.
Public Function SaveJXL_ToFile(ByRef srcImage As pdImage, ByRef srcOptions As String, ByRef dstFile As String) As Boolean

    Const FUNC_NAME As String = "SaveJXL_ToFile"
    SaveJXL_ToFile = False
    
    'Retrieve the composited pdImage object.
    Dim finalDIB As pdDIB
    srcImage.GetCompositedImage finalDIB, False
    
    'We don't actually need any special options here; we just need to save a PNG, then pass that PNG off to
    ' libjxl for final conversion.
    Dim tmpPngFile As String
    tmpPngFile = OS.UniqueTempFilename(customExtension:="png")
    
    Dim cPNG As pdPNG
    Set cPNG = New pdPNG
    SaveJXL_ToFile = (cPNG.SavePNG_ToFile(tmpPngFile, finalDIB, srcImage, png_AutoColorType, 0, 0, vbNullString, True) = png_Success)
    
    If SaveJXL_ToFile Then
        
        'Convert the saved PNG to JXL
        SaveJXL_ToFile = Plugin_jxl.ConvertImageFileToJXL(tmpPngFile, dstFile, srcOptions, False)
        
        'Regardless of success/failure, delete the temporary PNG
        Files.FileDeleteIfExists tmpPngFile
        
    Else
        Files.FileDeleteIfExists tmpPngFile
        InternalError FUNC_NAME, "tmp png failed"
    End If
    
    Exit Function
    
FatalEncoderError:
    SaveJXL_ToFile = False
    InternalError FUNC_NAME, "VB error # " & Err.Number

End Function

'Save a full image stack as an animated JPEG XL file (using APNG as an intermediary format).
Public Function SaveJXL_ToFile_Animated(ByRef srcImage As pdImage, ByRef srcOptions As String, ByRef dstFile As String) As Boolean

    Const FUNC_NAME As String = "SaveJXL_ToFile_Animated"
    SaveJXL_ToFile_Animated = False
    
    'We don't actually need any special options here; we just need to save a PNG, then pass that PNG off to
    ' libjxl for final conversion.
    Dim tmpPngFile As String
    tmpPngFile = OS.UniqueTempFilename(customExtension:="apng")
    
    Dim cPNG As pdPNG
    Set cPNG = New pdPNG
    SaveJXL_ToFile_Animated = (cPNG.SaveAPNG_ToFile(tmpPngFile, srcImage, png_AutoColorType, 0, 0, vbNullString) = png_Success)
    
    If SaveJXL_ToFile_Animated Then
        
        'Convert the saved PNG to JXL
        SaveJXL_ToFile_Animated = Plugin_jxl.ConvertImageFileToJXL(tmpPngFile, dstFile, srcOptions, False)
        
        'Regardless of success/failure, delete the temporary PNG
        Files.FileDeleteIfExists tmpPngFile
        
    Else
        Files.FileDeleteIfExists tmpPngFile
        InternalError FUNC_NAME, "tmp png failed"
    End If
    
    Exit Function
    
FatalEncoderError:
    SaveJXL_ToFile_Animated = False
    InternalError FUNC_NAME, "VB error # " & Err.Number

End Function

'Returns TRUE if the installed version of libjxl is >= the expected version of libjxl.
' By design, this function also returns TRUE if libjxl is NOT installed - this is purposeful because
' I don't want to raise "library out of date" warnings if the library doesn't even exist (there's a
' separate code pathway for downloading the library for the first time).
Public Function CheckJXLVersionAndOfferUpdates(Optional ByVal targetIsImportLib As Boolean = True) As Boolean
    
    'By design, this function returns TRUE if libavif doesn't exist.
    Dim libNotInstalled As Boolean
    If targetIsImportLib Then
        libNotInstalled = (Not IsJXLImportAvailable)
    Else
        libNotInstalled = (Not IsJXLExportAvailable)
    End If
    
    If libNotInstalled Then
        CheckJXLVersionAndOfferUpdates = True
        Exit Function
    End If
    
    'Still here?  libavif exists in this install.  Let's pull its version and compare it to the expected version
    ' (for this build of PhotoDemon).
    Dim curVersion As String
    curVersion = Plugin_jxl.GetLibJXLVersion()
    
    Dim expectedVersion As String
    expectedVersion = PluginManager.ExpectedPluginVersion(CCP_libjxl)
    
    If Updates.IsNewVersionHigher(curVersion, expectedVersion) Then
        
        'The installed copy of libavif is out-of-date.  Offer to download a new copy.
        CheckJXLVersionAndOfferUpdates = False
        
        Dim okToDownload As VbMsgBoxResult
        okToDownload = Updates.OfferPluginUpdate("libjxl", curVersion, expectedVersion)
        If (okToDownload = vbYes) Then
            
            CheckJXLVersionAndOfferUpdates = DownloadLatestLibjxl()
            
            'On a successful download, clean-up any old libjxl files remaining in the plugin folder
            If CheckJXLVersionAndOfferUpdates Then
                Files.FileDeleteIfExists PluginManager.GetPluginPath() & "brotlicommon.dll"
                Files.FileDeleteIfExists PluginManager.GetPluginPath() & "brotlidec.dll"
                Files.FileDeleteIfExists PluginManager.GetPluginPath() & "brotlienc.dll"
                Files.FileDeleteIfExists PluginManager.GetPluginPath() & "libjxl.dll"
                Files.FileDeleteIfExists PluginManager.GetPluginPath() & "jxl.dll"
                Files.FileDeleteIfExists PluginManager.GetPluginPath() & "jxl_threads.dll"
            End If
            
        End If
        
    Else
        CheckJXLVersionAndOfferUpdates = True
    End If
    
End Function

'Notify the user that PD can automatically download and configure JPEG XL support for them.
'
'Returns TRUE if PD successfully downloaded (and initialized) all required plugins.
Public Function PromptForLibraryDownload_JXL(Optional ByVal targetIsImportLib As Boolean = True) As Boolean
    
    On Error GoTo BadDownload
    
    'Like most modern libraries, libjxl requires at least Win 7
    If OS.IsWin7OrLater() Then
    
        'Ask the user for permission to (attempt) download
        Dim uiMsg As pdString
        Set uiMsg = New pdString
        uiMsg.AppendLine g_Language.TranslateMessage("JPEG XL (JXL) is a modern replacement for the JPEG image format.  PhotoDemon does not natively support JPEG XL images, but it can download a free, open-source plugin that adds JPEG XL support.")
        uiMsg.AppendLineBreak
        uiMsg.AppendLine g_Language.TranslateMessage("The libjxl library provides free, open-source JPEG XL compatibility.  A portable copy of libjxl will require ~%1 mb of disk space.  Once downloaded, PhotoDemon can use libjxl to load and save JPEG XL images (including animations).", 8)
        uiMsg.AppendLineBreak
        uiMsg.Append g_Language.TranslateMessage("Would you like PhotoDemon to download libjxl to your PhotoDemon plugin folder?")
        
        Dim msgReturn As VbMsgBoxResult
        msgReturn = PDMsgBox(uiMsg.ToString, vbInformation Or vbYesNoCancel, "Download required")
        If (msgReturn <> vbYes) Then
            
            'On a NO response, provide additional feedback.
            If (msgReturn = vbNo) Then
                uiMsg.Reset
                uiMsg.AppendLine g_Language.TranslateMessage("PhotoDemon will not download libjxl at this time.")
                PDMsgBox uiMsg.ToString, vbInformation Or vbOKOnly, "Download canceled"
            End If
            
            PromptForLibraryDownload_JXL = False
            Exit Function
            
        End If
        
        'The user said YES!  Attempt to download the latest libavif release now.
        PromptForLibraryDownload_JXL = DownloadLatestLibjxl()
        
    End If
    
    Exit Function
    
BadDownload:
    PromptForLibraryDownload_JXL = False
    Exit Function

End Function

'Attempt to download the latest libjxl copy to this PC.
Private Function DownloadLatestLibjxl() As Boolean
    
    ' Currently we expect these files in the package:
    ' - cjxl.exe (for encoding)
    ' - djxl.exe (for decoding)
    ' - jxlinfo.exe (for determining if a given file is a valid JXL image)
    ' - libjxl-LICENSE.txt (copyright and license info)
    Const EXPECTED_NUM_FILES As Long = 4
    
    'Current libjxl build is 0.11.1, downloaded from https://github.com/libjxl/libjxl/releases/tag/v0.11.1
    Const EXPECTED_TOTAL_EXTRACT_SIZE As Long = 8592383
    Const UPDATE_URL As String = "https://github.com/tannerhelland/PhotoDemon-Updates-v2/releases/download/libjxl-plugins-0.11.1/libjxl-0.11.1.pdz"
    DownloadLatestLibjxl = Updates.DownloadPluginUpdate(CCP_libjxl, UPDATE_URL, EXPECTED_NUM_FILES, EXPECTED_TOTAL_EXTRACT_SIZE)
    
End Function

'The following two functions are for logging errors (always active) and/or informational processing messages
' (only when JXL_DEBUG_VERBOSE = True).
'
' To use these functions outside PhotoDemon, substitute PDDebug.LogAction with your own logger.
Private Sub DebugMsg(ByRef msgText As String)
    PDDebug.LogAction msgText, PDM_External_Lib, True
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then DebugMsg "Plugin_jxl." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "Plugin_jxl." & funcName & "() reported an error: " & errDescription
    End If
End Sub
