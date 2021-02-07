Attribute VB_Name = "Plugin_8bf"
'***************************************************************************
'8bf Plugin Interface
'Copyright 2021-2021 by Tanner Helland
'Created: 07/February/21
'Last updated: 07/February/21
'Last update: initial build
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
Private Declare Function pspiSetColorPickerCallBack Lib "pspiHost.dll" Alias "_pspiSetcolorPickerCallBack@4" (ByVal addressOfCallback As Long) As PSPI_Result

'Execute various plugin functions
Private Declare Function pspiPlugInLoad Lib "pspiHost.dll" Alias "_pspiPlugInLoad@4" (ByVal ptrStrFilterPath As Long) As PSPI_Result
Private Declare Function pspiPlugInAbout Lib "pspiHost.dll" Alias "_pspiPlugInAbout@4" (ByVal ownerhWnd As Long) As PSPI_Result
Private Declare Function pspiPlugInExecute Lib "pspiHost.dll" Alias "_pspiPlugInExecute@4" (ByVal ownerhWnd As Long) As PSPI_Result

'Prep plugin features and image access
'Plugins support a "region of interest" in the source image
Private Declare Function pspiSetRoi Lib "pspiHost.dll" Alias "_pspiSetRoi@16" (Optional ByVal roiTop As Long = 0, Optional ByVal roiLeft As Long = 0, Optional ByVal roiBottom As Long = 0, Optional ByVal roiRight As Long = 0) As PSPI_Result
'// set image orientation
Private Declare Function pspiSetImageOrientation Lib "pspiHost.dll" Alias "_pspiSetImageOrientation@4" (ByVal newOrientation As PSPI_ImgOrientation) As PSPI_Result
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

'Path to 8bf plugins
Private m_8bfPath As String

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

'Returns the number of discovered 8bf plugins; 0 means no plugins found
Public Function EnumerateAvailable8bf() As Long
    
    Const funcName As String = "EnumerateAvailable8bf"
    
    'Failsafes
    If (Not m_LibAvailable) Then Exit Function
    If (LenB(m_8bfPath) = 0) Then Exit Function
    
    'Prepare a default size for the enum array
    Const INIT_COLLECTION_SIZE As Long = 16
    ReDim m_Plugins(0 To INIT_COLLECTION_SIZE - 1) As PD_Plugin8bf
    m_numPlugins = 0
    
    'Call the enumerator and hope for the best
    Dim retPspi As PSPI_Result
    retPspi = pspiPlugInEnumerate(AddressOf Enumerate8bfCallback, 1)
    If (retPspi = PSPI_OK) Then
        EnumerateAvailable8bf = m_numPlugins
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
        Debug.Print .plugCategory, .plugName, .plugEntryPoint, .plugLocationOnDisk
    End With
    
    m_numPlugins = m_numPlugins + 1
    
End Sub

Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

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

Public Sub ReleaseEngine()
    If (m_LibHandle <> 0) Then
        VBHacks.FreeLib m_LibHandle
        m_LibHandle = 0
    End If
End Sub

Public Function Set8bfPath(ByRef dstPath As String) As Boolean
    
    Const funcName As String = "Set8bfPath"
    
    m_8bfPath = Files.PathAddBackslash(dstPath)
    If Files.PathExists(dstPath, False) Then
    
        Dim retPspi As PSPI_Result
        retPspi = pspiSetPath(StrPtr(m_8bfPath))
        Set8bfPath = (retPspi = PSPI_OK)
        If (Not Set8bfPath) Then InternalError funcName, "pspi error", retPspi
        
    Else
        InternalError funcName, "path doesn't exist: " & dstPath
    End If
    
End Function

'Experimental only; show a plugin's About dialog
Public Sub ShowAboutDialog(ByVal plgIndex As Long)
    
    Const funcName As String = "ShowAboutDialog"
    
    If (plgIndex < 0) Or (plgIndex >= m_numPlugins) Then Exit Sub
    
    'Load plugin
    Dim retPspi As PSPI_Result
    retPspi = pspiPlugInLoad(StrPtr(m_Plugins(plgIndex).plugLocationOnDisk))
    If (retPspi <> PSPI_OK) Then
        InternalError funcName, "couldn't load plugin", retPspi
        Exit Sub
    End If
    
    'Display about dialog.  Note that this function may return "dummy proc" which is expected and OK
    retPspi = pspiPlugInAbout(FormMain.hWnd)
    If (retPspi <> PSPI_OK) And (retPspi <> PSPI_ERR_FILTER_DUMMY_PROC) Then
        InternalError funcName, "couldn't show About dialog", retPspi
    End If
    
    'There is no explict "unload plugin" function, alas
    
End Sub

'Produce a sorted list (by category, then function name)
Public Sub SortAvailable8bf()
    
    'Failsafe checks
    If (m_numPlugins < 2) Then Exit Sub
    
    Dim i As Long, j As Long
    Dim tmpSort As PD_Plugin8bf, searchCont As Boolean
    i = 1
    
    Do While (i < m_numPlugins)
        tmpSort = m_Plugins(i)
        j = i - 1
        
        'Because VB6 doesn't short-circuit And statements, we have to split this check into separate parts.
        searchCont = False
        If (j >= 0) Then searchCont = (Strings.StrCompSortPtr(StrPtr(m_Plugins(j).plugSortKey), StrPtr(tmpSort.plugSortKey)) > 0)
        '(m_GradientPoints(j).PointPosition > tmpSort.PointPosition)
        
        Do While searchCont
            m_Plugins(j + 1) = m_Plugins(j)
            j = j - 1
            searchCont = False
            If (j >= 0) Then searchCont = (Strings.StrCompSortPtr(StrPtr(m_Plugins(j).plugSortKey), StrPtr(tmpSort.plugSortKey)) > 0)
        Loop
        
        m_Plugins(j + 1) = tmpSort
        i = i + 1
        
    Loop
    
    'List the final collection
    For i = 0 To m_numPlugins - 1
        Debug.Print m_Plugins(i).plugCategory & " > " & m_Plugins(i).plugName
    Next i
            
End Sub

Private Sub InternalError(ByRef errFuncName As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    PDDebug.LogAction "WARNING!  Problem in Plugin_8bf." & errFuncName & ": " & errDescription
    If (errNum <> 0) Then PDDebug.LogAction "  (If it helps, an error number was also reported: #" & errNum & ")"
End Sub
