Attribute VB_Name = "Plugin_Management"
'***************************************************************************
'Core Plugin Manager
'Copyright 2014-2015 by Tanner Helland
'Created: 30/August/15
'Last updated: 30/August/15
'Last update: migrate a ton of scattered plugin management code to this singular module
'
'As PD grows, it's more and more difficult to supply the functionality we need through WAPI alone.  To that end,
' a number of third-party plugins are now required for proper program operation.
'
'To simplify the management of these plugins, I've created this singular module.  My hope is that future plugins
' will be easier to add and maintain thanks to this.
'
'When adding a new plugin, please make sure to read the declarations at the top of the class, particularly the
' CORE_PLUGINS enum and associated CORE_PLUGIN_COUNT constant.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Currently supported core plugins.  These values are arbitrary and can be changed without consequence, but THEY MUST
' ALWAYS BE SEQUENTIAL, STARTING WITH ZERO, because the enum is iterated using For loops (e.g. during initialization).
Public Enum CORE_PLUGINS
    CCP_FreeImage = 0
    CCP_zLib = 1
    CCP_ExifTool = 2
    CCP_EZTwain = 3
    CCP_PNGQuant = 4
End Enum

#If False Then
    Private Const CCP_FreeImage = 0, CCP_zLib = 1, CCP_ExifTool = 2, CCP_EZTwain = 3, CCP_PNGQuant = 4
#End If

'This constant is used to iterate all core plugins (as listed under the CORE_PLUGINS enum), so if you add or remove
' a plugin, make sure to update this!
Private Const CORE_PLUGIN_COUNT As Long = 5

'To simplify handling throughout this module, plugin existence, allowance, and successful initialization are tracked internally.
' Note that these values ARE NOT EXTERNALLY AVAILABLE, by design; external callers should use the global plugin trackers
' (e.g. g_ZLibEnabled, g_ExifToolEnabled, etc) because those trackers encompass the combination of all these factors, and are
' thus preferable for high-performance code paths.
Private pluginExists() As Boolean
Private pluginAllowed() As Boolean
Private pluginInitialized() As Boolean

'This subroutine handles the detection and installation of all core plugins. required for an optimal PhotoDemon
' experience: zLib, EZTwain32, and FreeImage.  For convenience' sake, it also checks for GDI+ availability.
Public Sub LoadAllPlugins()
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "LoadAllPlugins() called.  Attempting to initialize core plugins..."
    #End If
    
    'Reset all plugin trackers
    ReDim pluginExists(0 To CORE_PLUGIN_COUNT - 1) As Boolean
    ReDim pluginAllowed(0 To CORE_PLUGIN_COUNT - 1) As Boolean
    ReDim pluginInitialized(0 To CORE_PLUGIN_COUNT - 1) As Boolean
    
    'Plugin files are located in the \Data\Plugins subdirectory
    g_PluginPath = g_UserPreferences.getAppPath & "Plugins\"
    
    'Make sure the plugin path exists
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If Not cFile.FolderExist(g_PluginPath) Then cFile.CreateFolder g_PluginPath, True
        
    'Plugin loading is handled in a loop.  This loop will call several helper functions, passing each sequential plugin
    ' index as defined by the CORE_PLUGINS enum (and matching CORE_PLUGIN_COUNT const).  Some initialization steps are
    ' shared among all plugins (e.g. checking for the plugin's existence), while some require custom initializations.
    ' This behavior is all carefully documented in the functions called by the initialization loop.
    Dim i As Long
    For i = 0 To CORE_PLUGIN_COUNT - 1
    
        'Before doing anything else, see if the plugin file actually exists.
        pluginExists(i) = doesPluginFileExist(i)
        
        'If the plugin file exists, see if the user has forcibly disabled it.  If they have, we can skip initialization.
        ' we can initialize it.  (Some plugins may not require this step; that's okay.)
        If pluginExists(i) Then pluginAllowed(i) = isPluginAllowed(i)
        
        'If the user has allowed a plugin to exist, attempt to initialize it.
        If pluginAllowed(i) Then pluginInitialized(i) = initializePlugin(i)
        
        'We now know enough to set global initialization flags.  (This step is technically optional; see comments in the matching sub.)
        setGlobalPluginFlags i, pluginInitialized(i)
                
        'Finally, if a plugin affects UI or other user-exposed bits, that's the last thing we set before exiting.
        ' (This step is optional; plugins do not need to support it.)
        finalizePluginInitialization i, pluginInitialized(i)
        
    Next i
    
    'Initialization complete!  In debug builds, write out some plugin debug information.
    #If DEBUGMODE = 1 Then
        
        Dim successfulPluginCount As Long
        successfulPluginCount = 0
        
        For i = 0 To CORE_PLUGIN_COUNT - 1
            If pluginInitialized(i) Then
                successfulPluginCount = successfulPluginCount + 1
            Else
                pdDebug.LogAction "WARNING!  Plugin ID#" & i & " (" & getPluginFilename(i) & ") was not initialized."
            End If
        Next i
        
        pdDebug.LogAction CStr(successfulPluginCount) & "/" & CStr(CORE_PLUGIN_COUNT) & " plugins initialized successfully."
        
    #End If
    
End Sub

'Given a plugin enum value, return a string of the core plugin's filename.  Note that this (obviously) does not include helper files,
' like README or LICENSE files - just the core DLL or EXE for the plugin.
Private Function getPluginFilename(ByVal pluginEnumID As CORE_PLUGINS) As String
    
    Select Case pluginEnumID
    
        Case CCP_ExifTool
            getPluginFilename = "exiftool.exe"
        
        Case CCP_EZTwain
            getPluginFilename = "eztw32.dll"
        
        Case CCP_FreeImage
            getPluginFilename = "FreeImage.dll"
        
        Case CCP_PNGQuant
            getPluginFilename = "pngquant.exe"
        
        Case CCP_zLib
            getPluginFilename = "zlibwapi.dll"
    
    End Select
    
End Function

'Given a plugin enum value, return a string stack of any non-essential files associated with the plugin.  This includes things like
' README or LICENSE files, and it can be EMPTY if no helper files exist.
'
'Returns TRUE if one or more helper files exist; FALSE if none exist.  This should make it easier for the caller to know if the
' string stack needs to be processed further.
Private Function getNonEssentialPluginFiles(ByVal pluginEnumID As CORE_PLUGINS, ByRef dstStringStack As pdStringStack) As Boolean
    
    If dstStringStack Is Nothing Then Set dstStringStack = New pdStringStack
    dstStringStack.resetStack
    
    Select Case pluginEnumID
    
        Case CCP_ExifTool
            dstStringStack.AddString "exiftool-README.txt"
                    
        Case CCP_EZTwain
            dstStringStack.AddString "eztwain-README.txt"
        
        Case CCP_FreeImage
            dstStringStack.AddString "freeimage-LICENSE.txt"
        
        Case CCP_PNGQuant
            dstStringStack.AddString "pngquant-README.txt"
        
        Case CCP_zLib
            dstStringStack.AddString "zlib-README.txt"
    
    End Select
    
    getNonEssentialPluginFiles = CBool(dstStringStack.getNumOfStrings <> 0)
    
End Function

'The Plugin Manager dialog allows the user to forcibly disable plugins.  This can be very helpful when testing bugs and crashes,
' but generally isn't relevant for a casual user.  Regardless, the plugin loader will check this value prior to initializing a plugin.
Private Function isPluginAllowed(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    
    Select Case pluginEnumID
    
        Case CCP_ExifTool
            isPluginAllowed = Not g_UserPreferences.GetPref_Boolean("Plugins", "Force ExifTool Disable", False)
                    
        Case CCP_EZTwain
            isPluginAllowed = Not g_UserPreferences.GetPref_Boolean("Plugins", "Force EZTwain Disable", False)
        
        Case CCP_FreeImage
            isPluginAllowed = Not g_UserPreferences.GetPref_Boolean("Plugins", "Force FreeImage Disable", False)
        
        Case CCP_PNGQuant
            isPluginAllowed = Not g_UserPreferences.GetPref_Boolean("Plugins", "Force PNGQuant Disable", False)
        
        Case CCP_zLib
            isPluginAllowed = Not g_UserPreferences.GetPref_Boolean("Plugins", "Force ZLib Disable", False)
    
    End Select
    
End Function

'If a function requires specialized initialization steps, add them here.  Do NOT add any user-facing interactions (e.g. UI) here,
' and DO NOT account for user preferences here.  (User preferences are handled via isPluginAllowed(), above).
'
'This step purely exists to handle custom initialization of plugins, when the master plugin file is known to exist in the
 'official plugin folder, and the user has not forcibly disabled a given plugin.
'
'Returns TRUE if the plugin was initialized successfully; FALSE otherwise.
Private Function initializePlugin(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    
    'Because this function has variable complexity (depending on the plugin), an intermediary value is used to track success/failure.
    ' At the end of the function, the function return will simply copy this value, so make sure it's set correctly before the
    ' function ends.
    Dim initializationSuccessful As Boolean
    
    Select Case pluginEnumID
    
        Case CCP_ExifTool
            'Unlike most plugins, ExifTool is an .exe file.  Because we interact with it asynchronously, we start it now, then leave
            ' it in "wait" mode.
            
            'Crashes (or IDE stop button use) can result in stranded ExifTool instances.  As a convenience to the caller, we attempt
            ' to kill any stranded instances before starting new ones.
            If Not peekLastShutdownClean Then
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Previous PhotoDemon session terminated unexpectedly.  Performing plugin clean-up..."
                #End If
                Plugin_ExifTool_Interface.killStrandedExifToolInstances
            End If
            
            'Attempt to shell a new ExifTool instance.  If shell fails (for whatever reason), the function will return FALSE.
            initializationSuccessful = Plugin_ExifTool_Interface.startExifTool()
                    
        Case CCP_EZTwain
            'The ezTwain module provides a function called "isEZTwainAvailable()", but all it does is check if the EZTwain DLL exists.
            ' This is redundant, so skip that check and forcibly return TRUE.
            initializationSuccessful = True
        
        Case CCP_FreeImage
            'FreeImage maintains a program-wide handle for the life of the program, which we attempt to generate now.
            initializationSuccessful = Plugin_FreeImage_Interface.initializeFreeImage()
            
        Case CCP_PNGQuant
            'The ezTwain module provides a function called "isPNGQuantAvailable()", but all it does is check if the PNGquant exe exists.
            ' This is redundant, so skip that check and forcibly return TRUE.
            initializationSuccessful = True
        
        Case CCP_zLib
            'zLib maintains a program-wide handle for the life of the program, which we attempt to generate now.
            initializationSuccessful = Plugin_zLib_Interface.initializeZLib()
            
    End Select

    initializePlugin = initializationSuccessful

End Function

'Most plugins provide a single global "is plugin available" flag, which spares the program from having to plow through all these
' verification steps when it needs to do something with a plugin.  This step is technically optional, although I prefer global flags
' because they let me use plugins in performance-sensitive areas without worry.
Private Sub setGlobalPluginFlags(ByVal pluginEnumID As CORE_PLUGINS, ByVal pluginState As Boolean)
    
    Select Case pluginEnumID
    
        Case CCP_ExifTool
            g_ExifToolEnabled = pluginState
                    
        Case CCP_EZTwain
            g_ScanEnabled = pluginState
        
        Case CCP_FreeImage
            g_ImageFormats.FreeImageEnabled = pluginState
            
        Case CCP_PNGQuant
            g_ImageFormats.pngQuantEnabled = pluginState
        
        Case CCP_zLib
            g_ZLibEnabled = pluginState
            
    End Select
    
End Sub

'This final plugin initialization step is OPTIONAL.
'
'It provides a catch-all for custom initialization behavior (e.g. modifying UI bits to reflect plugin-related features).
' New plugins do not need to make use of this functionality.
Private Sub finalizePluginInitialization(ByVal pluginEnumID As CORE_PLUGINS, ByVal pluginState As Boolean)

    Select Case pluginEnumID
                
        Case CCP_EZTwain
            'EZTwain is currently the only supported method for scanners.  (I hope to fix this in the future.)
            ' As such, availability of the scanner UI is based on EZTwain's successful initialization.
            FormMain.MnuScanImage.Visible = pluginState
            FormMain.MnuSelectScanner.Visible = pluginState
            FormMain.MnuImportSepBar1.Visible = pluginState
        
        Case CCP_FreeImage
            'As of v6.4, PD uses a callback function to track and report any internal FreeImage errors.
            If pluginState Then
                #If DEBUGMODE = 1 Then
                    Outside_FreeImageV3.FreeImage_InitErrorHandler
                #End If
            End If
            
            'Also, at present, the arbitrary rotation option wraps FreeImage's internal rotate functions.  I've been too lazy to
            ' add fallbacks for this, but I may revisit in a future release.
            FormMain.MnuRotate(3).Visible = pluginState
            
        Case Else
        
    End Select
    
End Sub

'NOTE: the following function is PLUGIN AGNOSTIC.  You do not need to modify it when adding a new plugin to the program.
'
'This function performs several tasks:
' 1) If the requested plugin file exists in the target folder, great; it returns TRUE and exits.
' 2) If the requested plugin file does NOT exist in the target folder, it scans the program folder to see if it can find a hit there.
' 3) If it finds a missing plugin in the program folder, it will automatically move the file to the plugin folder, including any
'     helper files (README, LICENSE, etc).
' 4) If the move is successful, it will return TRUE and exit.
Private Function doesPluginFileExist(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    
    'Start by getting the filename of the plugin in question
    Dim pluginFilename As String
    pluginFilename = getPluginFilename(pluginEnumID)
    
    'pdFSO is used for all file interactions
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'See if the file exists.  If it does, great!  We can exit immediately.
    If cFile.FileExist(g_PluginPath & pluginFilename) Then
        doesPluginFileExist = True
    
    'The plugin file is missing.  Let's see if we can find it.
    Else
    
        Dim extraFiles As pdStringStack
        Set extraFiles = New pdStringStack
    
        'See if the plugin file exists in the base PD folder.  This can happen if a user unknowingly extracts the PD .zip without
        ' folders preserved.
        If cFile.FileExist(g_UserPreferences.getProgramPath & pluginFilename) Then
        
            'Move the plugin file to the proper folder
            If cFile.CopyFile(g_UserPreferences.getProgramPath & pluginFilename, g_PluginPath & pluginFilename) Then
            
                'Kill the old plugin instance
                cFile.KillFile g_UserPreferences.getProgramPath & pluginFilename
                
                'Finally, move any associated files to their new home in the plugin folder
                If getNonEssentialPluginFiles(pluginEnumID, extraFiles) Then
                    
                    Dim tmpFilename As String
                    
                    Do While extraFiles.PopString(tmpFilename)
                        
                        If cFile.CopyFile(g_UserPreferences.getProgramPath & tmpFilename, g_PluginPath & tmpFilename) Then
                            cFile.KillFile g_UserPreferences.getProgramPath & tmpFilename
                        End If
                        
                    Loop
                    
                End If
                
                'Return success!
                doesPluginFileExist = True
            
            'The file couldn't be moved.  There's probably write issues with the folder structure, in which case the program
            ' as a whole is pretty much doomed.  Exit now.
            Else
                doesPluginFileExist = False
            End If
        
        'If the plugin file doesn't exist in the base folder either, we're SOL.  Exit now.
        Else
            doesPluginFileExist = False
        End If
    
    End If
    
End Function
