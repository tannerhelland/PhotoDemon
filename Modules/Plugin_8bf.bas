Attribute VB_Name = "Plugin_8bf"
'***************************************************************************
'8bf Plugin Interface
'Copyright 2021-2026 by Tanner Helland
'Created: 07/February/21
'Last updated: 22/January/26
'Last update: continued hardening against run-time errors
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
'Thank you to Sinisa for their work.
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

Option Explicit

'Verbose debug logging; turn OFF is production builds
Private Const DEBUG_VERBOSE As Boolean = False

'During the transitionary period (from pspihost to native code), I've added a toggle to switch
' between pspihost and our own native code when triggering plugin behavior(s).
Private Const USE_NATIVE_INTERFACE As Boolean = True

Private Enum PSPI_Result
    PSPI_OK = 0
    PSPI_ERR_FILTER_NOT_LOADED = 1
    PSPI_ERR_FILTER_BAD_PROC = 2
    PSPI_ERR_FILTER_ABOUT_ERROR = 3
    PSPI_ERR_FILTER_DUMMY_PROC = 4
    PSPI_ERR_FILTER_CANCELED = 5
    PSPI_ERR_FILTER_CRASHED = 6
    PSPI_ERR_FILTER_INVALID = 7
    PSPI_ERR_IMAGE_INVALID = 10
    PSPI_ERR_MEMORY_ALLOC = 11
    PSPI_ERR_INIT_PATH_EMPTY = 12
    PSPI_ERR_WORK_PATH_EMPTY = 13
    PSPI_ERR_BAD_PARAM = 14
    PSPI_ERR_BAD_IMAGE_TYPE = 15
End Enum

#If False Then
    Private Const PSPI_OK = 0, PSPI_ERR_FILTER_NOT_LOADED = 1, PSPI_ERR_FILTER_BAD_PROC = 2, PSPI_ERR_FILTER_ABOUT_ERROR = 3, PSPI_ERR_FILTER_DUMMY_PROC = 4, PSPI_ERR_FILTER_CANCELED = 5, PSPI_ERR_FILTER_CRASHED = 6, PSPI_ERR_FILTER_INVALID = 7
    Private Const PSPI_ERR_IMAGE_INVALID = 10, PSPI_ERR_MEMORY_ALLOC = 11, PSPI_ERR_INIT_PATH_EMPTY = 12, PSPI_ERR_WORK_PATH_EMPTY = 13, PSPI_ERR_BAD_PARAM = 14, PSPI_ERR_BAD_IMAGE_TYPE = 15
#End If

'Supported color formats
Private Enum PSPI_ImgType
    PSPI_IMG_TYPE_BGR = 0
    PSPI_IMG_TYPE_BGRA
    PSPI_IMG_TYPE_RGB
    PSPI_IMG_TYPE_RGBA
    PSPI_IMG_TYPE_GRAY
    PSPI_IMG_TYPE_GRAYA
End Enum

#If False Then
    Private Const PSPI_IMG_TYPE_BGR = 0, PSPI_IMG_TYPE_BGRA = 1, PSPI_IMG_TYPE_RGB = 2, PSPI_IMG_TYPE_RGBA = 3, PSPI_IMG_TYPE_GRAY = 4, PSPI_IMG_TYPE_GRAYA = 5
#End If

'Initialization
Private Declare Function pspiGetVersion Lib "pspiHost.dll" Alias "_pspiGetVersion@0" () As Long
Private Declare Function pspiSetPath Lib "pspiHost.dll" Alias "_pspiSetPath@4" (ByVal strPtrFilterFolder As Long) As PSPI_Result

'Callbacks
'Private Declare Function pspiPlugInEnumerate Lib "pspiHost.dll" Alias "_pspiPlugInEnumerate@8" (ByVal addressOfCallback As Long, Optional ByVal bRecurseSubFolders As Long = 1) As PSPI_Result
Private Declare Function pspiSetProgressCallBack Lib "pspiHost.dll" Alias "_pspiSetProgressCallBack@4" (ByVal addressOfCallback As Long) As PSPI_Result

'Execute various plugin functions
Private Declare Function pspiPlugInLoad Lib "pspiHost.dll" Alias "_pspiPlugInLoad@4" (ByVal ptrStrFilterPath As Long) As PSPI_Result
'Private Declare Function pspiPlugInAbout Lib "pspiHost.dll" Alias "_pspiPlugInAbout@4" (ByVal ownerHwnd As Long) As PSPI_Result
Private Declare Function pspiPlugInExecute Lib "pspiHost.dll" Alias "_pspiPlugInExecute@4" (ByVal ownerHwnd As Long) As PSPI_Result

'Prep plugin features and image access
'Plugins support a "region of interest" in the source image
Private Declare Function pspiSetRoi Lib "pspiHost.dll" Alias "_pspiSetRoi@16" (Optional ByVal roiTop As Long = 0, Optional ByVal roiLeft As Long = 0, Optional ByVal roiBottom As Long = 0, Optional ByVal roiRight As Long = 0) As PSPI_Result
'// set image using contiguous memory buffer pointer
'// note: source image is shared - do not destroy source image in your host program before plug-in is executed
Private Declare Function pspiSetImage Lib "pspiHost.dll" Alias "_pspiSetImage@28" (ByVal tImgType As PSPI_ImgType, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal ptrImageBuffer As Long, ByVal imgStride As Long, Optional ByVal ptrAlphaBuffer As Long = 0, Optional ByVal ptrAlphaStride As Long = 0) As PSPI_Result
'// set mask using contiguous memory buffer pointer
'// note: source mask is shared - do not destroy source mask in your host program before plug-in is executed
' Additional notes from README on final bool parameter:
' "This value tells if the mask will be used by plug-in (if supported by plug-in) or only for blending filtered
'  and source image. Calling pspiSetMask without parameters releases mask (internal scanline pointers). This holds
'  also when mask is set by scanline addition."
Private Declare Function pspiSetMask Lib "pspiHost.dll" Alias "_pspiSetMask@20" (Optional ByVal maskWidth As Long = 0, Optional ByVal maskHeight As Long = 0, Optional ByVal ptrMaskBuffer As Long = 0, Optional ByVal maskStride As Long = 0, Optional ByVal bPluginCanUseMask As Long = 0) As PSPI_Result
'// release all images - all image buffers including mask are released when application is closed, but sometimes
' it's necessary to free memory (big images)
Private Declare Function pspiReleaseAllImages Lib "pspiHost.dll" Alias "_pspiReleaseAllImages@0" () As PSPI_Result

'On a successful call to pspiPlugInEnumerate, an array will be filled with plugin names and paths.
Private Type PD_Plugin8bf
    plugCategory As String
    plugName As String
    'plugEntryPoint As String
    plugLocationOnDisk As String
    plugSortKey As String
End Type

Private m_Plugins() As PD_Plugin8bf, m_numPlugins As Long

'Library handle will be non-zero if pspi is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE
Private m_LibHandle As Long, m_LibAvailable As Boolean

'Index of currently loaded plugin, if any (vbNullString if no plugin loaded)
Private m_Active8bf As String

'To handle progress callbacks, we need to distinguish the first progress event (because that's when we
' load and display a progress bar on the main screen).
Private m_HasSeenProgressEvent As Boolean, m_LastProgressAmount As Long, m_TimeOfLastProgEvent As Currency

'High-res time stamp when the first progress callback is hit
Private m_FirstTimeStamp As Currency

'Selection mask contents, if any
'Private m_MaskCopy() As Byte

'When enumerating plugins, the user can pass an (optional) progress bar.  We'll update the bar as plugins
' are found and loaded.
Private m_EnumProgressBar As pdProgressBar

'DEC 2025: APIs for manually enumerating and handling 8bf filters
Private Declare Function EnumResourceNamesW Lib "kernel32" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function FindResourceW Lib "kernel32" (ByVal hModule As Long, ByVal lpName As Long, ByVal lpType As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hModule As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hModule As Long, ByVal hResInfo As Long) As Long

'Photoshop SDK code follows
Private Type PiPLResource
    resSignature As Integer
    resVersion As Long
    resCount As Long
    resPData As Long
End Type

' /** Definition of a PiPL property. Plug-in property structures (or properties) are
' * the basic units of information stored in a property list. Properties are variable
' * length data structures, which are uniquely identified by a vendor code, property key,
' * and ID number. PiPL properties are stored in a list.  See \c PIPropertyList.
' */
Public Type PIProperty

    '/** The vendor defining this property type. This allows vendors to define
    '* their own properties in a way that does not conflict with either Adobe or other
    '* vendors. It is recommended that a registered application creator code be used
    '* for the vendorID to ensure uniqueness. All Photoshop properties use the vendorID
    '* '8BIM'.
    '*/
    VendorId As String * 4
    
    '/// Identification key for this property type. See @ref PiPLKeys "Property Keys".
    propertyKey As String * 4
    
    '/// Index within this property type. Must be unique for properties of
    '/// a given type in a PiPL.
    propertyID As Long
    
    '/// Length of propertyData. Does not include any padding bytes
    '/// to achieve four byte alignment. May be zero.
    propertyLength As Long
    
    '/// Variable length field containing contents of this property.
    '/// Any values may be contained. Must be padded to achieve four
    '/// byte alignment.
    pPropertyData() As Byte
    
    'For PD only: textual properties (like name, category) will store their processed string name here
    propertyAsString As String

End Type

'Properties of the current file
Private m_numProperties As Long, m_Properties() As PIProperty

'Current plugin file being scanned
Private m_CurrentPluginFilename As String

'Array of plugin objects.  These will (someday) launch the actual plugins involved.
Private m_numSafePlugins As Long, m_SafePlugins() As pd8bf

Public Function Execute8bf(ByVal ownerHwnd As Long, ByRef pluginCanceled As Boolean, Optional ByVal catchProgress As Boolean = True) As Boolean
    
    Const FUNC_NAME As String = "Execute8bf"
    
    If (Not m_LibAvailable) Then
        InternalError FUNC_NAME, "pspihost unavailable"
        Exit Function
    End If
    
    Dim retPspi As PSPI_Result
    
    'Before executing a plugin, we want to queue up a progress callback
    If catchProgress Then
        retPspi = pspiSetProgressCallBack(AddressOf Plugin_8bf.Progress8bfCallback)
        If (retPspi <> PSPI_OK) Then InternalError FUNC_NAME, "couldn't set progress callback", retPspi
        m_HasSeenProgressEvent = False
    End If
    
    retPspi = pspiPlugInExecute(ownerHwnd)
    Execute8bf = (retPspi = PSPI_OK)
    pluginCanceled = (retPspi = PSPI_ERR_FILTER_CANCELED)
    
    If (Not Execute8bf) And (Not pluginCanceled) Then
        InternalError FUNC_NAME, "plugin execution failed", retPspi
    End If
    
End Function

Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

'This is a hacky way to "free" a loaded plugin, but it ensures that FreeLibrary gets called on the
' currently loaded 8bf (if any)
Public Sub Free8bf()
    
    Const FUNC_NAME As String = "Free8bf"
    
    If (Not m_LibAvailable) Then
        InternalError FUNC_NAME, "pspihost unavailable"
        Exit Sub
    End If
    
    m_Active8bf = vbNullString
    Dim retPspi As PSPI_Result
    retPspi = pspiPlugInLoad(StrPtr(""))    'Cannot be null string, must be *empty* string!
    
End Sub

Public Sub FreeImageResources()
    
    Const FUNC_NAME As String = "FreeImageResources"
    
    If (Not m_LibAvailable) Then
        InternalError FUNC_NAME, "pspihost unavailable"
        Exit Sub
    End If
    
    pspiSetMask 0, 0, 0, 0, 0   'See documentation; null parameters frees mask pointers and associated resources
    'Erase m_MaskCopy            'Mask is no longer passed to pspihost; it frequently misuses it and crashes
    pspiReleaseAllImages        'pspi will auto-free upon close, but PD also needs to free unsafe pointers to temporary structs
    
End Sub

'Return value is the number of plugins found by this enumeration instance (e.g. the set produced by the
' last call to EnumerateAvailable8bf).  Note that all strings are appended to the existing stacks, so if
' you already have strings in there, *those strings will not be removed*, by design.
Public Function GetEnumerationResults(ByRef catNames As pdStringStack, ByRef plgNames As pdStringStack, ByRef plgFiles As pdStringStack) As Long

    GetEnumerationResults = m_numPlugins
    If (GetEnumerationResults > 0) Then
        
        If (catNames Is Nothing) Then Set catNames = New pdStringStack
        If (plgNames Is Nothing) Then Set plgNames = New pdStringStack
        If (plgFiles Is Nothing) Then Set plgFiles = New pdStringStack
        
        Dim i As Long
        For i = 0 To m_numPlugins - 1
            catNames.AddString m_Plugins(i).plugCategory
            plgNames.AddString m_Plugins(i).plugName
            plgFiles.AddString m_Plugins(i).plugLocationOnDisk
        Next i
        
    End If

End Function

Public Function GetInitialEffectTimestamp() As Currency
    GetInitialEffectTimestamp = m_FirstTimeStamp
End Function

Public Function GetPspiVersion() As String
    
    Const FUNC_NAME As String = "GetPspiVersion"
    
    If (Not m_LibAvailable) Then
        InternalError FUNC_NAME, "pspihost unavailable"
        Exit Function
    End If
    
    Dim ptrVersion As Long
    ptrVersion = pspiGetVersion()
    If (ptrVersion <> 0) Then GetPspiVersion = Strings.StringFromCharPtr(ptrVersion, False, 3, True) & ".0"
    
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean
    
    Const FUNC_NAME As String = "InitializeEngine"
    
    Dim strLibPath As String
    strLibPath = pathToDLLFolder & "pspiHost.dll"
    
    'Ensure the plugin exists before attempting further load steps
    If (Not Files.FileExists(strLibPath)) Then
        m_LibHandle = 0
        m_LibAvailable = False
        InitializeEngine = False
        InternalError FUNC_NAME, "pspihost.dll missing"
        Exit Function
    End If
    
    m_LibHandle = VBHacks.LoadLib(strLibPath)
    m_LibAvailable = (m_LibHandle <> 0)
    InitializeEngine = m_LibAvailable
    
    If (Not InitializeEngine) Then
        InternalError FUNC_NAME, "LoadLibraryW failed: " & Err.LastDllError
    End If
    
End Function

Public Function IsPspiEnabled() As Boolean
    IsPspiEnabled = m_LibAvailable
End Function

Public Function Load8bf(ByRef fullPathToPlugin As String) As Boolean
    
    Const FUNC_NAME As String = "Load8bf"
    Load8bf = False
    
    If (Not m_LibAvailable) Then
        InternalError FUNC_NAME, "pspihost unavailable"
        Exit Function
    End If
    
    If (LenB(fullPathToPlugin) = 0) Then
        InternalError FUNC_NAME, "null plugin path"
        Exit Function
    End If
    
    If (Not Files.FileExists(fullPathToPlugin)) Then
        InternalError FUNC_NAME, "bad plugin path: " & fullPathToPlugin
        Exit Function
    End If
    
    Dim retPspi As PSPI_Result
    retPspi = pspiPlugInLoad(StrPtr(fullPathToPlugin))
    Load8bf = (retPspi = PSPI_OK)
    
    If Load8bf Then
        m_Active8bf = fullPathToPlugin
    Else
        m_Active8bf = vbNullString
        InternalError FUNC_NAME, "couldn't load plugin", retPspi
    End If
    
End Function

Public Sub Progress8bfCallback(ByVal amtDone As Long, ByVal amtTotal As Long)
    
    'Sometimes weird stuff can happen in callbacks, possibly a result of unstable plugin code.
    ' We don't want VB to freak out, so if something goes wrong, just exit immediately - this
    ' is just a progress update, and there's no harm if we exit prematurely.
    On Error GoTo ExitCallback
    
    'If this is the first progress event, activate the main screen's progress bar
    If (Not m_HasSeenProgressEvent) Then
        ProgressBars.SetProgBarMax amtTotal
        m_HasSeenProgressEvent = True
        
        Message "Applying plugin..."
        Processor.MarkProgramBusyState True, True
        
        VBHacks.GetHighResTime m_FirstTimeStamp
        
    'If this is *not* the first progress event, throttle events as necessary to minimize delays
    ' caused by on-screen progress rendering.
    Else
        If (VBHacks.GetTimerDifferenceNow(m_TimeOfLastProgEvent) < 0.1) Or (m_LastProgressAmount = amtDone) Then Exit Sub
    End If
    
    'Debug.Print "progress callback", amtDone, amtTotal
    
    'Update progress
    ProgressBars.SetProgBarVal amtDone
    
    'Note the time that this event occurred; we'll use this to throttle excessive progress requests
    VBHacks.GetHighResTime m_TimeOfLastProgEvent
    m_LastProgressAmount = amtDone
    
ExitCallback:
    
End Sub

Public Sub ReleaseEngine()
    If (m_LibHandle <> 0) Then
        VBHacks.FreeLib m_LibHandle
        m_LibHandle = 0
    End If
End Sub

Public Sub ResetPluginCollection()
    m_numPlugins = 0
    m_numSafePlugins = 0
    Erase m_SafePlugins     'Force free any underlying 8bf classes
End Sub

'Short-hand function for automatically setting the plugin image to PD's active working image.  Note that
' the image is *shared* with the plugin, so you can't free the image without first freeing the plugin
' without things going horribly wrong!
Public Function SetImage_CurrentWorkingImage(Optional ByVal pspiMaskOK As Boolean = False) As Boolean
    
    Const funcName As String = "SetImage_CurrentWorkingImage"
    
    If (Not m_LibAvailable) Then
        InternalError funcName, "pspihost unavailable"
        Exit Function
    End If
    
    'Failsafes
    If (LenB(m_Active8bf) = 0) Then Exit Function
    
    'Create a standard PD working copy of the image
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, ignoreSelection:=pspiMaskOK
    
    'Notify the plugin of the shared image
    Dim retPspi As PSPI_Result
    retPspi = pspiSetImage(PSPI_IMG_TYPE_BGRA, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, workingDIB.GetDIBPointer, workingDIB.GetDIBStride, 0, 0)
    SetImage_CurrentWorkingImage = (retPspi = PSPI_OK)
    If (retPspi <> PSPI_OK) Then InternalError funcName, "pspiSetImage failed", retPspi
    
    retPspi = pspiSetRoi(0, 0, workingDIB.GetDIBHeight - 1, workingDIB.GetDIBWidth - 1)
    If (retPspi <> PSPI_OK) Then InternalError funcName, "pspiSetRoifailed", retPspi
    
    SetImage_CurrentWorkingImage = SetImage_CurrentWorkingImage And (retPspi = PSPI_OK)
    If (retPspi <> PSPI_OK) Then InternalError funcName, "pspiSetImageOrientation failed", retPspi
    
    'TODO: set mask if selection is active
    
End Function

'JAN 2026: this function has been marked for removal, since it doesn't work 90+% of the time.
' (Full removal will take place after I've produced a working version in native VB6 code.)
''Short-hand function for automatically setting the plugin mask to PD's active selection.  Note that
'' the mask is *shared* with the plugin (but we actually host it), so we must not free our mask copy
'' until the appropriate plugin shutdown functions are called.
'Public Function SetMask_CurrentSelectionMask() As Boolean
'
'    Const funcName As String = "SetMask_CurrentSelectionMask"
'
'    'Make sure we're being called correctly
'    If PDImages.GetActiveImage.IsSelectionActive And PDImages.GetActiveImage.MainSelection.IsLockedIn Then
'
'        'Retrieve a copy of the mask
'        Dim tmpDIB As pdDIB
'        Set tmpDIB = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB()
'
'        'Retrieve just the alpha channel
'        DIBs.RetrieveSingleChannel tmpDIB, m_MaskCopy, 0
'
'        'Notify pspi
'        Dim retPspi As PSPI_Result
'        retPspi = pspiSetMask(tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, VarPtr(m_MaskCopy(0, 0)), tmpDIB.GetDIBWidth, 1)
'        SetMask_CurrentSelectionMask = (retPspi = PSPI_OK)
'        If (retPspi <> PSPI_OK) Then InternalError funcName, "pspiSetMask failed", retPspi
'
'    '/no mask; do nothing
'    End If
'
'End Function

'Show an 8bf plugin's About dialog.  Note that not all plugins expose an About dialog, and they'll often
' return "success" even though they do nothing.  There's no way to circumvent this behavior, alas -
' so calling this function is basically the tech equivalent of a prayer and a wish.
Public Sub ShowAboutDialog(ByRef fullPathToPlugin As String, Optional ByVal ownerHwnd As Long = 0)
    
    Const funcName As String = "ShowAboutDialog"
    If (ownerHwnd = 0) Then ownerHwnd = FormMain.hWnd
    
    'Native VB6
    If USE_NATIVE_INTERFACE Then
        
        If (m_numSafePlugins <= 0) Then Exit Sub
        
        'Find the matching plugin object
        Dim idxTarget As Long
        idxTarget = -1
        
        Dim i As Long
        For i = 0 To m_numSafePlugins - 1
            If (Not m_SafePlugins(i) Is Nothing) Then
                If Strings.StringsEqual(m_SafePlugins(i).GetFilename, fullPathToPlugin, True) Then
                    idxTarget = i
                    Exit For
                End If
            End If
        Next i
        
        If (idxTarget >= 0) Then
            Dim retShow As Boolean
            retShow = m_SafePlugins(i).ShowAboutDialog(ownerHwnd)
        End If
        
    'pspihost (disabled for now)
    Else
'
'        If Plugin_8bf.Load8bf(fullPathToPlugin) Then
'
'            'Display about dialog.  Note that this function may return "dummy proc" which is expected and OK
'            Dim retPspi As PSPI_Result
'            retPspi = pspiPlugInAbout(FormMain.hWnd)
'            If (retPspi <> PSPI_OK) And (retPspi <> PSPI_ERR_FILTER_DUMMY_PROC) Then
'                InternalError funcName, "couldn't show About dialog", retPspi
'            End If
'
'            'Free the plugin
'            Plugin_8bf.Free8bf
'
'        End If
'
    End If
    
End Sub

'PD-specific function to display a UI for plugin selection
Public Sub ShowPluginDialog()

    Dim tmpForm As FormEffects8bf
    Set tmpForm = New FormEffects8bf
    
ShowDialogAgain:
    Interface.ShowPDDialog vbModal, tmpForm, True
    
    'Regardless of what happened, free the progress bar and restore default UI behavior
    ProgressBars.ReleaseProgressBar
    Interface.EnableUserInput
    FormMain.MousePointer = vbDefault
    
    'If the plugin was canceled, show the dialog again
    If tmpForm.RestoreDialog() Then GoTo ShowDialogAgain
    
    'Because pspihost is buggy, I've found improved reliability from forcibly freeing any loaded plugins
    ' (regardless of what happened with this dialog) after user interaction.
    Plugin_8bf.ResetPluginCollection
    
    Unload tmpForm
    Set tmpForm = Nothing
    
End Sub

'Produce a sorted list of 8bf plugins (sorted by category, then function name; path is not considered)
Public Sub SortAvailable8bf()
    
    'Failsafe check for null/single plugin lists
    If (m_numPlugins < 2) Then Exit Sub
    
    'Given the number of plugins a typical user has (asymptotically approaching 0 lol),
    ' a quick insertion sort works fine.
    Dim i As Long, j As Long
    Dim tmpSort As PD_Plugin8bf, searchCont As Boolean
    i = 1
    
    Do While (i < m_numPlugins)
        tmpSort = m_Plugins(i)
        j = i - 1
        
        'Because VB6 doesn't short-circuit And statements, we have to split this check into separate parts.
        searchCont = False
        If (j >= 0) Then searchCont = (Strings.StrCompSortPtr(StrPtr(m_Plugins(j).plugSortKey), StrPtr(tmpSort.plugSortKey)) > 0)
        
        Do While searchCont
            m_Plugins(j + 1) = m_Plugins(j)
            j = j - 1
            searchCont = False
            If (j >= 0) Then searchCont = (Strings.StrCompSortPtr(StrPtr(m_Plugins(j).plugSortKey), StrPtr(tmpSort.plugSortKey)) > 0)
        Loop
        
        m_Plugins(j + 1) = tmpSort
        i = i + 1
        
    Loop
            
End Sub

Private Sub InternalError(ByRef errFuncName As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    PDDebug.LogAction "WARNING!  Problem in Plugin_8bf." & errFuncName & ": " & errDescription
    If (errNum <> 0) Then PDDebug.LogAction "  (If it helps, an error number was also reported: #" & errNum & ")"
End Sub

'******************************************************
' UPDATE DEC 2025: due to ongoing stability issues, I want to reimplement as much of this plugin's behavior
' myself as I can.
'
'First up: enumerating plugin categories and names.  This function returns the number of valid plugins found.
' Inputs:
'   - srcListOfFiles (stack containing a list of candidate 8bf files)
' Returns:
'   - the net count of 32-bit, validation-passed 8bf plugins we found.
'
'To retrieve the actual list of validated failes, call GetEnumerationResults() after this function returned
' a non-zero value.
Public Function EnumeratePlugins_PD(ByRef srcListOfFiles As pdStringStack, Optional ByRef prgUpdate As pdProgressBar) As Long
    
    Const FUNC_NAME As String = "EnumeratePlugins_PD"
    
    EnumeratePlugins_PD = 0
    m_numSafePlugins = 0
    ReDim m_SafePlugins(0) As pd8bf
    
    If (srcListOfFiles Is Nothing) Then Exit Function
    If (srcListOfFiles.GetNumOfStrings <= 0) Then Exit Function
    
    Dim targetFile As String, numFilesExamined As Long
    Do While srcListOfFiles.PopString(targetFile)
        
        numFilesExamined = numFilesExamined + 1
        If (Not prgUpdate Is Nothing) Then
            prgUpdate.Value = numFilesExamined
        End If
        
        'Failsafe check to ensure we weren't passed bad files (yes, this syntax is how you do this in VB6)
        If (Not Files.FileExists(targetFile)) Then GoTo ContinueWithNextFile
        
        'Prior to loading, ensure the plugin is...
        ' 1) a DLL, and
        ' 2) ...a 32-bit x86 DLL (64-bit is unsupported until a TB version of PD matures)
        If (OS.GetDLLBitness(targetFile) <> 32) Then
            If DEBUG_VERBOSE Then InternalError FUNC_NAME, "not 32-bit: " & targetFile
            GoTo ContinueWithNextFile
        End If
        
        If DEBUG_VERBOSE Then PDDebug.LogAction "Attempting PiPL reads for " & targetFile
        m_CurrentPluginFilename = targetFile
        
        'Attempt to load the library.  (We must pass the returned handle to the resource enumerator.)
        ' Note that we're not going to execute any code in the library on this pass - we simply want to
        ' query plugin properties.  To improve performance, load the data as a read-only data resource.
        Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2&
        
        'Per MSDN (https://learn.microsoft.com/en-us/windows/win32/api/libloaderapi/nf-libloaderapi-loadlibraryexw):
        ' "If this value is used, the system maps the file into the calling process's virtual address space
        '  as if it were a data file. Nothing is done to execute or prepare to execute the mapped file.
        '  Therefore, you cannot call functions like GetModuleFileName, GetModuleHandle or GetProcAddress with
        '  this DLL. Using this value causes writes to read-only memory to raise an access violation. Use this
        '  flag when you want to load a DLL only to extract messages or resources from it."
        Dim dwFlags As Long
        dwFlags = LOAD_LIBRARY_AS_DATAFILE
        
        'Per MSDN (https://learn.microsoft.com/en-us/windows/win32/api/libloaderapi/nf-libloaderapi-loadlibraryexw):
        ' "If this value is used, the system maps the file into the process's virtual address space as an image file.
        '  However, the loader does not load the static imports or perform the other usual initialization steps.
        '  Use this flag when you want to load a DLL only to extract messages or resources from it.  Unless the
        '  application depends on the file having the in-memory layout of an image, this value should be used with
        '  either LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE or LOAD_LIBRARY_AS_DATAFILE.  Windows Server 2003 and Windows XP:
        '  This value is not supported until Windows Vista."
        Const LOAD_LIBRARY_AS_IMAGE_RESOURCE As Long = &H20&
        If OS.IsVistaOrLater Then dwFlags = dwFlags Or LOAD_LIBRARY_AS_IMAGE_RESOURCE
        
        'Attempt loading the plugin, again as a READ-ONLY data resource
        Dim hLib As Long
        hLib = VBHacks.LoadLibExW(targetFile, dwFlags)
        If (hLib = 0) Then GoTo ContinueWithNextFile
        
        'Next we want to pull the resource block and "walk" individual resources, querying as we go.
        ' (NOTE: the enumerator callback may be called multiple times, if multiple filters exist
        '  inside a single plugin.)
        Dim resName As String
        resName = "PiPL"
        If (EnumResourceNamesW(hLib, StrPtr(resName), AddressOf EnumResNameProcW, 0&) <> 0) Then
            'The enumeration returned successfully.
        End If
        
        'Make sure we free the library before continuing!
        VBHacks.FreeLib hLib
        hLib = 0
        
        'EnumResourceNamesW
ContinueWithNextFile:
    Loop
    
    'TODO - TEMPORARY SOLUTION:
    '
    'While we're stuck with this weird half-us-half-pspihost solution, just copy over our enumeration results
    ' into the structs used for the old pspihost interface.  This code can be rewritten once pspihost is excised.
    
    'Ensure at least one plugin was found
    If (m_numSafePlugins <= 0) Then
        m_numPlugins = 0
        EnumeratePlugins_PD = 0
        Exit Function
    
    'Still here?  Attempt to retrieve source strings.
    Else
        
        m_numPlugins = m_numSafePlugins
        ReDim Preserve m_Plugins(0 To m_numSafePlugins - 1) As PD_Plugin8bf
        
        Dim i As Long
        For i = 0 To m_numSafePlugins - 1
            
            With m_Plugins(i)
            
                .plugCategory = m_SafePlugins(i).Get8bfCategory()
                .plugName = m_SafePlugins(i).Get8bfName()
                .plugLocationOnDisk = m_SafePlugins(i).GetFilename()
                
                'Prep a convenient sort key
                .plugSortKey = m_SafePlugins(i).Get8bfSortKey()
                
                'Curious about contents?  See 'em here:
                'Debug.Print .plugCategory, .plugName, .plugLocationOnDisk
                
            End With
            
        Next i
        
        EnumeratePlugins_PD = m_numSafePlugins
        
    End If
    
End Function

'Callback for the EnumResourceNamesW API (used to pull data from 8bf files)
Private Function EnumResNameProcW(ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal lUserParam As Long) As Long
    
    Const FUNC_NAME As String = "EnumResNameProcW"
    
    'Failsafes for bad data
    If (hModule = 0) Or (lpType = 0) Or (lpName = 0) Then
        EnumResNameProcW = 0
        Exit Function
    End If
    
    'Next, we want to Find -> Load -> Size -> Lock the target PiPL resource.
    ' Note that (per MSDN) resources are returned as hGlobals for backwards compatibility, but they must not
    ' be accessed or freed using Global* calls.  The data is auto-freed when the target library is freed.
    Dim hResource As Long
    hResource = FindResourceW(hModule, lpName, lpType)
    If (hResource = 0) Then
        EnumResNameProcW = 0
        Exit Function
    End If
    
    Dim hGlobalNotReally As Long
    hGlobalNotReally = LoadResource(hModule, hResource)
    If (hGlobalNotReally = 0) Then
        EnumResNameProcW = 0
        Exit Function
    End If
    
    'As a failsafe, grab resource size in case the plugin is malformed
    Dim resSize As Long
    resSize = SizeofResource(hModule, hResource)
    If (resSize = 0) Then
        EnumResNameProcW = 0
        Exit Function
    End If
    
    'Lock the resource, then we can access it!
    Dim hLocked As Long
    hLocked = LockResource(hGlobalNotReally)
    If (hLocked = 0) Then
        EnumResNameProcW = 0
        Exit Function
    End If
    
    'With the resource ready, we can now "walk" it, noting properties as we go.
    
    'Start by pointing a stream at the target resource.  This simplifies reading arbitrary data types
    ' from an arbitrary pointer in VB6.
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, resSize, hLocked) Then
    
        Dim thisResource As PiPLResource
        
        'Load the resource header
        thisResource.resSignature = cStream.ReadInt()
        thisResource.resVersion = cStream.ReadLong()
        thisResource.resCount = cStream.ReadLong()
        
        'The first property starts here!  Don't read it yet; VB6 limitations mean we want to leave the
        ' stream pointer where it is, and we'll manually load struct members.
        'thisResource.resPData = cStream.ReadLong()
        
        'Validate the header.  The signature is typically "1" (I don't enforce this because I don't
        ' think it matters) but what we really care about is the version being "0".  This is a hard
        ' requirement by PS.  From the SDK:
        '/// Current Plug-in Property List version
        Const kCurrentPiPLVersion As Long = 0
        If (thisResource.resVersion <> kCurrentPiPLVersion) Then
            InternalError FUNC_NAME, m_CurrentPluginFilename & " bad PiPL version: " & thisResource.resVersion
            EnumResNameProcW = 1
            Exit Function
        End If
        
        'Zero-count resource lists are useless to us
        If (thisResource.resCount <= 0) Then
            EnumResNameProcW = 1
            Exit Function
        End If
        
        'With the header pulled, we can now walk through individual properties
        If DEBUG_VERBOSE Then PDDebug.LogAction "Walking " & thisResource.resCount & " resources..."
        
        'Reset the module-level property tracker
        m_numProperties = thisResource.resCount
        ReDim m_Properties(0 To m_numProperties - 1) As PIProperty
        
        'Iterate properties one-by-one.  As always, use a stream to help.
        Dim i As Long
        For i = 0 To m_numProperties - 1
        
            Dim initOffset As Long
            initOffset = cStream.GetPosition()
            
            'Pull the prop header out of the stream
            With m_Properties(i)
                
                'Some struct members are 4-byte strings, but because of endianness we're going to first read them
                ' to a temporary int, *then* read them as a string
                Dim tmpKey As Long
                tmpKey = cStream.ReadLong_BE()
                .VendorId = Strings.StringFromCharPtr(VarPtr(tmpKey), False, 4, True)
                
                tmpKey = cStream.ReadLong_BE()
                .propertyKey = Strings.StringFromCharPtr(VarPtr(tmpKey), False, 4, True)
                
                'ID and length are just uints
                .propertyID = cStream.ReadLong()
                .propertyLength = cStream.ReadLong()
                
                'Also pull the property data in; these tend to be small (< 100 bytes) and their
                ' data may be useful to the initializer
                ReDim .pPropertyData(0 To .propertyLength - 1) As Byte
                cStream.ReadBytesToBarePointer VarPtr(.pPropertyData(0)), .propertyLength
                
                'Want to see the data?  Dump it to debug here:
                'Debug.Print .VendorId, .propertyKey, .propertyID, .propertyLength
                
            End With
            
            'Property length is always reported as-is, but must be manually padded to 4-byte alignment
            ' before advancing to the next property.
            Dim paddedPropLength As Long
            paddedPropLength = (m_Properties(i).propertyLength + 3) And &H7FFFFFFC
            
            'Advance the pointer to the next property.  The hard-coded 16& is the size of the
            ' property header (not included in the property length value).
            cStream.SetPosition initOffset + 16& + paddedPropLength, FILE_BEGIN
            
        Next i
        
        'Turn off the stream before exiting, or it will crash on class termination
        cStream.StopStream
        
        'All properties in this file have now been read and cached at module-level.
        
        'Next, we need to validate this plugin's properties before displaying it to the user.
        ' (Maybe it only operates on weird color spaces, or it's actually a 64-bit DLL, etc.)
        '
        'Instantiate a new plugin class for this file+action combination, and note that the
        ' initializer will return TRUE if the plugin is compatible with PD.
        If (m_numSafePlugins > UBound(m_SafePlugins)) Then ReDim Preserve m_SafePlugins(0 To m_numSafePlugins * 2 - 1) As pd8bf
        Set m_SafePlugins(m_numSafePlugins) = New pd8bf
        If m_SafePlugins(m_numSafePlugins).Initialize8bf_FromFile(m_CurrentPluginFilename, m_numProperties, m_Properties) Then
            m_numSafePlugins = m_numSafePlugins + 1
        Else
            If DEBUG_VERBOSE Then PDDebug.LogAction "WARNING: initialization failed for " & m_CurrentPluginFilename
        End If
        
        'We are done processing this plugin file, but this function may get called again for a
        ' different plugin in the same file (files can contain multiple plugins).
            
    Else
        EnumResNameProcW = 1
        Exit Function
    End If
    
    'Return success before exiting (return type is a BOOL)
    EnumResNameProcW = 1
    
End Function
