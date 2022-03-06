Attribute VB_Name = "Plugin_8bf"
'***************************************************************************
'8bf Plugin Interface
'Copyright 2021-2022 by Tanner Helland
'Created: 07/February/21
'Last updated: 10/February/21
'Last update: add better UI support (progress bar tracking) during plugin enumeration
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
'Thank you to Sinisa for their great work.
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

'Windows uses bottom-up DIBs by default, so it's helpful to have a toggle for scanline orientation
Private Enum PSPI_ImgOrientation
    PSPI_IMG_ORIENTATION_ASIS = 0
    PSPI_IMG_ORIENTATION_INVERT
End Enum

#If False Then
    Private Const PSPI_IMG_ORIENTATION_ASIS = 0, PSPI_IMG_ORIENTATION_INVERT = 1
#End If

'Initialization
Private Declare Function pspiGetVersion Lib "pspiHost.dll" Alias "_pspiGetVersion@0" () As Long
Private Declare Function pspiSetPath Lib "pspiHost.dll" Alias "_pspiSetPath@4" (ByVal strPtrFilterFolder As Long) As PSPI_Result

'Callbacks
Private Declare Function pspiPlugInEnumerate Lib "pspiHost.dll" Alias "_pspiPlugInEnumerate@8" (ByVal addressOfCallback As Long, Optional ByVal bRecurseSubFolders As Long = 1) As PSPI_Result
Private Declare Function pspiSetProgressCallBack Lib "pspiHost.dll" Alias "_pspiSetProgressCallBack@4" (ByVal addressOfCallback As Long) As PSPI_Result

'TODO:
'Private Declare Function pspiSetColorPickerCallBack Lib "pspiHost.dll" Alias "_pspiSetcolorPickerCallBack@4" (ByVal addressOfCallback As Long) As PSPI_Result

'Execute various plugin functions
Private Declare Function pspiPlugInLoad Lib "pspiHost.dll" Alias "_pspiPlugInLoad@4" (ByVal ptrStrFilterPath As Long) As PSPI_Result
Private Declare Function pspiPlugInAbout Lib "pspiHost.dll" Alias "_pspiPlugInAbout@4" (ByVal ownerHwnd As Long) As PSPI_Result
Private Declare Function pspiPlugInExecute Lib "pspiHost.dll" Alias "_pspiPlugInExecute@4" (ByVal ownerHwnd As Long) As PSPI_Result

'Prep plugin features and image access
'Plugins support a "region of interest" in the source image
Private Declare Function pspiSetRoi Lib "pspiHost.dll" Alias "_pspiSetRoi@16" (Optional ByVal roiTop As Long = 0, Optional ByVal roiLeft As Long = 0, Optional ByVal roiBottom As Long = 0, Optional ByVal roiRight As Long = 0) As PSPI_Result
'// set image orientation
'Private Declare Function pspiSetImageOrientation Lib "pspiHost.dll" Alias "_pspiSetImageOrientation@4" (ByVal newOrientation As PSPI_ImgOrientation) As PSPI_Result
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
'// block for adding image scanlines (possibly non-contiguous image)
'// note: source image is shared - do not destroy source image in your host program before plug-in is executed
'private declare function pspiStartImageSL Lib "pspiHost.dll" Alias "" (TImgType type, int width, int height, bool externalAlpha = false) As PSPI_Result
'private declare function pspiAddImageSL Lib "pspiHost.dll" Alias "" (void *imageScanLine, void *alphaScanLine = 0) As PSPI_Result
'private declare function pspiFinishImageSL Lib "pspiHost.dll" Alias "" (int imageStride = 0, int alphaStride = 0) As PSPI_Result
'// block dor addding mask scanlines  Lib "pspiHost.dll" Alias "" (possibly non-contiguous mask)
'// note: source mask is shared - do not destroy source maske in your host program before plug-in is executed
'private declare function pspiStartMaskSL Lib "pspiHost.dll" Alias "" (int width, int height, bool useMaskByPi = true) As PSPI_Result
'private declare function pspiAddMaskSL Lib "pspiHost.dll" Alias "" (void *maskScanLine) As PSPI_Result
'private declare function pspiFinishMaskSL Lib "pspiHost.dll" Alias "" (int maskStride = 0) As PSPI_Result
'// release all images - all image buffers including mask are released when application is closed, but sometimes
' it's necessary to free memory (big images)
Private Declare Function pspiReleaseAllImages Lib "pspiHost.dll" Alias "_pspiReleaseAllImages@0" () As PSPI_Result

'On a successful call to pspiPlugInEnumerate, an array will be filled with plugin names and paths.
Private Type PD_Plugin8bf
    plugCategory As String
    plugName As String
    plugEntryPoint As String
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
Private m_MaskCopy() As Byte

'When enumerating plugins, the user can pass an (optional) progress bar.  We'll update the bar as plugins
' are found and loaded.
Private m_EnumProgressBar As pdProgressBar

'Returns the number of discovered 8bf plugins; 0 means no plugins found.  Note that you can call this
' function back-to-back with different folders, and it will just keep appending discoveries to a central list.
' This makes it convenient to do a single list sort before calling GetEnumerateResults(), below.
Public Function EnumerateAvailable8bf(ByVal srcPath As String, Optional ByRef dstProgressBar As pdProgressBar = Nothing) As Long
    
    Const funcName As String = "EnumerateAvailable8bf"
    
    'Failsafes
    If (Not m_LibAvailable) Then Exit Function
    If (LenB(srcPath) = 0) Then Exit Function
    
    Dim retPspi As PSPI_Result
    
    'Ensure path exists
    srcPath = Files.PathAddBackslash(srcPath)
    If Files.PathExists(srcPath, False) Then
        retPspi = pspiSetPath(StrPtr(srcPath))
        If (retPspi <> PSPI_OK) Then
            InternalError funcName, "pspiSetPath error: " & srcPath, retPspi
            Exit Function
        End If
    Else
        InternalError funcName, "path doesn't exist: " & srcPath
        Exit Function
    End If
    
    'Prepare a default size for the enum array
    Const INIT_COLLECTION_SIZE As Long = 16
    If (m_numPlugins = 0) Then ReDim m_Plugins(0 To INIT_COLLECTION_SIZE - 1) As PD_Plugin8bf
    
    'We will report the number of new plugins added in this enumeration, only
    Dim numPluginsAtStart As Long
    numPluginsAtStart = m_numPlugins
    
    'Note the target progress bar, if any
    Set m_EnumProgressBar = dstProgressBar
    
    'Call the enumerator and hope for the best
    retPspi = pspiPlugInEnumerate(AddressOf Enumerate8bfCallback, 1)
    If (retPspi = PSPI_OK) Then
        EnumerateAvailable8bf = m_numPlugins - numPluginsAtStart
    Else
        InternalError funcName, "pspi failed", retPspi
    End If
    
End Function

Private Sub Enumerate8bfCallback(ByVal ptrCategoryA As Long, ByVal ptrNameA As Long, ByVal ptrEntryPointA As Long, ByVal ptrLocationW As Long)
    
    Const funcName As String = "Enumerate8bfCallback"
    
    'Failsafe checks
    If (ptrCategoryA = 0) Or (ptrNameA = 0) Or (ptrEntryPointA = 0) Or (ptrLocationW = 0) Then
        InternalError funcName, "bad char *"
        Exit Sub
    End If
    
    'Still here?  Attempt to retrieve source strings.
    If (m_numPlugins > UBound(m_Plugins)) Then ReDim Preserve m_Plugins(0 To m_numPlugins * 2 - 1) As PD_Plugin8bf
    With m_Plugins(m_numPlugins)
    
        .plugCategory = Strings.StringFromCharPtr(ptrCategoryA, False)
        .plugName = Strings.StringFromCharPtr(ptrNameA, False)
        .plugEntryPoint = Strings.StringFromCharPtr(ptrEntryPointA, False)
        .plugLocationOnDisk = Strings.StringFromCharPtr(ptrLocationW, True)
        
        'Prep a convenient sort key
        .plugSortKey = .plugCategory & "_" & .plugName
        
        'Curious about contents?  See 'em here:
        'Debug.Print .plugCategory, .plugName, .plugEntryPoint, .plugLocationOnDisk
        
    End With
    
    m_numPlugins = m_numPlugins + 1
    
    'Update the target progress bar (if one exists)
    If (Not m_EnumProgressBar Is Nothing) Then m_EnumProgressBar.Value = m_numPlugins
    
End Sub

Public Function Execute8bf(ByVal ownerHwnd As Long, ByRef pluginCanceled As Boolean, Optional ByVal catchProgress As Boolean = True) As Boolean
    
    Const funcName As String = "Execute8bf"
    
    Dim retPspi As PSPI_Result
    
    'Before executing a plugin, we want to queue up a progress callback
    If catchProgress Then
        retPspi = pspiSetProgressCallBack(AddressOf Plugin_8bf.Progress8bfCallback)
        If (retPspi <> PSPI_OK) Then InternalError funcName, "couldn't set progress callback", retPspi
        m_HasSeenProgressEvent = False
    End If
    
    retPspi = pspiPlugInExecute(ownerHwnd)
    Execute8bf = (retPspi = PSPI_OK)
    pluginCanceled = (retPspi = PSPI_ERR_FILTER_CANCELED)
    
    If (Not Execute8bf) And (Not pluginCanceled) Then
        InternalError funcName, "plugin execution failed", retPspi
    End If
    
End Function

Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

'This is a hacky way to "free" a loaded plugin, but it ensures that FreeLibrary gets called on the
' currently loaded 8bf (if any)
Public Sub Free8bf()
    m_Active8bf = vbNullString
    Dim retPspi As PSPI_Result
    retPspi = pspiPlugInLoad(StrPtr(""))    'Cannot be null string, must be *empty* string!
End Sub

Public Sub FreeImageResources()
    pspiSetMask 0, 0, 0, 0, 0   'See documentation; null parameters frees mask pointers and associated resources
    Erase m_MaskCopy
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
    Dim ptrVersion As Long
    ptrVersion = pspiGetVersion()
    If (ptrVersion <> 0) Then GetPspiVersion = Strings.StringFromCharPtr(ptrVersion, False, 3, True) & ".0"
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim strLibPath As String
    strLibPath = pathToDLLFolder & "pspiHost.dll"
    m_LibHandle = VBHacks.LoadLib(strLibPath)
    m_LibAvailable = (m_LibHandle <> 0)
    InitializeEngine = m_LibAvailable
    
    If (Not InitializeEngine) Then
        PDDebug.LogAction "WARNING!  LoadLibraryW failed to load pspiHost.  Last DLL error: " & Err.LastDllError
    End If
    
End Function

Public Function IsPspiEnabled() As Boolean
    IsPspiEnabled = m_LibAvailable
End Function

Public Function Load8bf(ByRef fullPathToPlugin As String) As Boolean
    
    Const funcName As String = "Load8bf"
    Load8bf = False
    
    If (LenB(fullPathToPlugin) = 0) Then Exit Function
    
    Dim retPspi As PSPI_Result
    retPspi = pspiPlugInLoad(StrPtr(fullPathToPlugin))
    Load8bf = (retPspi = PSPI_OK)
    
    If Load8bf Then
        m_Active8bf = fullPathToPlugin
    Else
        m_Active8bf = vbNullString
        InternalError funcName, "couldn't load plugin", retPspi
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
End Sub

'Short-hand function for automatically setting the plugin image to PD's active working image.  Note that
' the image is *shared* with the plugin, so you can't free the image without first freeing the plugin
' without things going horribly wrong!
Public Function SetImage_CurrentWorkingImage(Optional ByVal pspiMaskOK As Boolean = False) As Boolean
    
    Const funcName As String = "SetImage_CurrentWorkingImage"
    
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

'Short-hand function for automatically setting the plugin mask to PD's active selection.  Note that
' the mask is *shared* with the plugin (but we actually host it), so we must not free our mask copy
' until the appropriate plugin shutdown functions are called.
Public Function SetMask_CurrentSelectionMask() As Boolean
    
    Const funcName As String = "SetMask_CurrentSelectionMask"
    
    'Make sure we're being called correctly
    If PDImages.GetActiveImage.IsSelectionActive And PDImages.GetActiveImage.MainSelection.IsLockedIn Then
    
        'Retrieve a copy of the mask
        Dim tmpDIB As pdDIB
        Set tmpDIB = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB()
        
        'Retrieve just the alpha channel
        DIBs.RetrieveSingleChannel tmpDIB, m_MaskCopy, 0
        
        'Notify pspi
        Dim retPspi As PSPI_Result
        retPspi = pspiSetMask(tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, VarPtr(m_MaskCopy(0, 0)), tmpDIB.GetDIBWidth, 1)
        SetMask_CurrentSelectionMask = (retPspi = PSPI_OK)
        If (retPspi <> PSPI_OK) Then InternalError funcName, "pspiSetMask failed", retPspi
        
    '/no mask; do nothing
    End If

End Function

'Experimental only; show a plugin's About dialog
Public Sub ShowAboutDialog(ByRef fullPathToPlugin As String, Optional ByVal ownerHwnd As Long = 0)
    
    Const funcName As String = "ShowAboutDialog"
    If (ownerHwnd = 0) Then ownerHwnd = FormMain.hWnd
    
    If Plugin_8bf.Load8bf(fullPathToPlugin) Then
        
        'Display about dialog.  Note that this function may return "dummy proc" which is expected and OK
        Dim retPspi As PSPI_Result
        retPspi = pspiPlugInAbout(FormMain.hWnd)
        If (retPspi <> PSPI_OK) And (retPspi <> PSPI_ERR_FILTER_DUMMY_PROC) Then
            InternalError funcName, "couldn't show About dialog", retPspi
        End If
        
        'Free the plugin
        Plugin_8bf.Free8bf
    
    End If
    
End Sub

'PD-specific function to display a UI for plugin selection
Public Sub ShowPluginDialog()

    Dim tmpForm As FormEffects8bf
    Set tmpForm = New FormEffects8bf
    
ShowDialogAgain:
    ShowPDDialog vbModal, tmpForm, True
    
    'Regardless of what happened, free the progress bar and restore default UI behavior
    ProgressBars.ReleaseProgressBar
    Interface.EnableUserInput
    FormMain.MousePointer = vbDefault
    
    'If the plugin was canceled, show the dialog again
    If tmpForm.RestoreDialog() Then GoTo ShowDialogAgain
    
    'TEMPORARY FIX UNTIL CACHING IS IMPLEMENTED:
    Plugin_8bf.ResetPluginCollection
    
    Unload tmpForm
    Set tmpForm = Nothing
    
End Sub

'Produce a sorted list (by category, then function name)
Public Sub SortAvailable8bf()
    
    'Failsafe checks
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
