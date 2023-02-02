VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileDialogVista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'You may download the original version of this class from the following link (good as of May '16):
' http://www.vbforums.com/showthread.php?785837-VB6-Common-Dialog-Replacement-IFileDialog-Vista-%28No-TLBs%29

'This class was originally written by vbforums.com user LaVolpe.  To the best of my knowledge, the code has
' been released into the public domain.

'Many thanks to LaVolpe for this non-TLB implementation of the new IFileDialog interface available in Vista+.
' PhotoDemon currently uses its updated PathBrowseDialog implementation, which is significantly better than
' the old XP-style browse box.


Option Explicit

' About this class. Can only be run on Vista or better.

' This class exposes the new IFileDialog, IFileOpenDialog & IFileSaveDialogs.
' To get events from the dialog while it is displayed, you will also need to include the
'   IFileDialogEvents class I've written. If you do not want any events from the dialog,
'   then you do not need to include that other class. However, if you don't include it,
'   you will need to rem out the code in the AttachEvents() method. That event makes a
'   hard-reference to the IFileDialogEvents class. The hard-reference is needed if the
'   class is being used. And if the class is not being used, then VB won't let you compile
'   with the reference coded. Rem'ing out the code in that routine has no ill effects when
'   this class calls that function.

' Ok the tricky part is covered. Now the plus side of things
' 1) The older file save/open dialogs, if you subclassed it, then Windows gave you the XP
'       look of the dialog and removed the Vista+ look. No longer the case here. If you
'       want to subclass the window, you will want to use the IFileDialogEvents class.
'       Then when you receive the DialogOnInt event, you can call this class' DialogHWND
'       property to get the dialog's hWnd. There are several built in events where you
'       may not even need to subclass it.
' 2) You have the option of adding controls to the dialog. Several types of controls can be
'       added: combobox,textbox,option buttons,command buttons,labels,checkbox,menu-like controls
'       If you add these, you will want to use the IFileDialogEvents class in order to receive
'       events related to those controls. Spend time reviewing comments I've added throughout.
'       Controls are referenced by the Key you assign when you create the control
'       *** Keys CANNOT contain any leading spaces
' 3) The IFileDialogEvents class. If used in your project, here's how to get the events:
'    a) At top of your form (class/usercontrol/whatever): Implements IFileDialogEvents
'       VB will add about 10 events for you to play with.
'    b) When wanting to display the dialog, create an instance of this class, then
'       - set up the properties of this class as needed
'       - call the Customize_CreateGroup & Customize_CreateNonGroupItem & Customize_OpenVirtualContainer
'           to create controls. Optionally, for menu, option buttons & combobox items, you can use
'           this method in a loop, if desired: Customize_AppendGroupItem
'    c) After you set your properties & added any controls, call: AttachEvents(Me, boolean, boolean, boolean)
'    d) Call DialogShow
' 4) This class does not maintain any properties or settings after the dialog closes, other than the
'       selected folder/file(s). Those are destroyed once the class closes or DialogShow called again
' 5) I've tried to keep the class organized as such:
'       Design-time properties that can be set before the dialog is displayed are prefixed with: prop
'       Run-time properties/methods are prefixed with: Dialog
'       Run-time IShell methods are prefixed with: IShell
'       Any other methods are intended for after the dialog is closed
' 6) Useful references:
'    Shared Open/Save Dialog methods/properties - http://msdn.microsoft.com/en-us/library/windows/desktop/bb775966%28v=vs.85%29.aspx
'    Specific Open Dialog methods/properties - http://msdn.microsoft.com/en-us/library/windows/desktop/bb775834%28v=vs.85%29.aspx
'    Specific Save Dialog methods/properites - http://msdn.microsoft.com/en-us/library/windows/desktop/bb775688%28v=vs.85%29.aspx
'    Dialog Events - http://msdn.microsoft.com/en-us/library/windows/desktop/bb775876%28v=vs.85%29.aspx
'    FileDialog Customization - http://msdn.microsoft.com/en-us/library/windows/desktop/bb775910%28v=vs.85%29.aspx
'    Customization Events - http://msdn.microsoft.com/en-us/library/windows/desktop/bb775936%28v=vs.85%29.aspx

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDst As Any, ByRef lpSrc As Any, ByVal byteLength As Long)
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
Private Declare Function CoCreateInstance Lib "ole32.dll" (ByVal rClsID As Long, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByVal rIID As Long, ByRef ppv As Long) As Long
Private Const CLSCTX_INPROC_SERVER As Long = 1

Private Const clsID_FileOpenDialog As String = "{DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7}"
Private Const IID_IFileOpenDialog As String = "{d57c7288-d4ad-4768-be02-9d969532d960}"
Private Const clsID_FileSaveDialog As String = "{C0B4E2F3-BA21-4773-8DBA-335EC946EB8B}"
Private Const IID_IFileSaveDialog As String = "{84bccd23-5fde-4cdb-aea4-af64b83d78ab}"
Private Const IID_IFileDialog As String = "{42F85136-DB7E-439C-85F1-E4075D135FC8}"
Private Const IID_IFileDialogCustomize As String = "{8016b7b3-3d49-4504-a0aa-2a37494e606f}"
Private Const IID_IShellItem As String = "{43826d1e-e718-42ee-bc55-a1e261c37bfe}"
Private Const IID_IShellItemArray As String = "{B63EA76D-1F85-456F-A19C-48159EFA858B}"

Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, ByVal lpiid As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Long, ByVal pClsID As Long) As Long

Private Enum InterfaceMethodOffsets                 ' helper for me during design
    ifUnknown_QueryInterface = 0                    ' pass 2 params
    ifUnknown_AddRef = 4                            ' no params
    ifUnknown_Release = 8                           ' no params
    ifOleWindow_GetWindow = 12                      ' pass 1 param
    ifModalWindow_Show = 12                         ' pass 1 param
    ifFileDialog_SetFileTypes = 16                  ' pass 2 params
    ifFileDialog_SetFileTypeIndex = 20              ' pass 1 param
    ifFileDialog_GetFileTypeIndex = 24              ' pass 1 param
    ifFileDialog_Advise = 28                        ' pass 2 params
    ifFileDialog_Unadvise = 32                      ' pass 1 param
    ifFileDialog_SetOptions = 36                    ' pass 1 param
    ifFileDialog_GetOptions = 40                    ' pass 1 param
    ifFileDialog_SetDefaultFolder = 44              ' pass 1 param
    ifFileDialog_SetFolder = 48                     ' pass 1 param
    ifFileDialog_FileGetPath = 52                     ' pass 1 param
    ifFileDialog_GetCurrentSelection = 56           ' pass 1 param
    ifFileDialog_SetFileName = 60                   ' pass 1 param
    ifFileDialog_GetFileName = 64                   ' pass 1 param
    ifFileDialog_SetTitle = 68                      ' pass 1 param
    ifFileDialog_SetOkButtonLabel = 72              ' pass 1 param
    ifFileDialog_SetFileNameLabel = 76              ' pass 1 param
    ifFileDialog_GetResult = 80                     ' pass 1 param
    ifFileDialog_AddPlace = 84                      ' pass 2 params
    ifFileDialog_SetDefaultExtension = 88           ' pass 1 param
    ifFileDialog_Close = 92                         ' pass 1 param
    ifFileDialog_SetClientGuid = 96                 ' pass 1 param
    ifFileDialog_ClearClientData = 100              ' no params
    ifFileDialog_SetFilter = 104                    ' pass 1 param
    ifFileDialogOpen_GetResults = 108               ' pass 1 param
    ifFileDialogOpen_GetSelectedItems = 112         ' pass 1 param
    ifFileDialogSave_SetSaveAsItem = 108            ' pass 1 param
    ifFileDialogSave_SetProperties = 112            ' pass 1 param
    ifFileDialogSave_SetCollectedProperties = 116   ' pass 2 params
    ifFileDialogSave_GetProperties = 120            ' pass 1 param
    ifFileDialogSave_ApplyProperties = 124          ' pass 4 params
    ifFileDlgCustomize_EnableOpenDropDown = 12      ' pass 1 param
    ifFileDlgCustomize_AddMenu = 16                 ' pass 2 params
    ifFileDlgCustomize_AddPushButton = 20           ' pass 2 params
    ifFileDlgCustomize_AddComboBox = 24             ' pass 1 param
    ifFileDlgCustomize_AddRadioButtonList = 28      ' pass 1 param
    ifFileDlgCustomize_AddCheckButton = 32          ' pass 3 params
    ifFileDlgCustomize_AddEditBox = 36              ' pass 2 params
    ifFileDlgCustomize_AddSeparator = 40            ' pass 1 param
    ifFileDlgCustomize_AddText = 44                 ' pass 1 param
    ifFileDlgCustomize_SetControlLabel = 48         ' pass 2 params
    ifFileDlgCustomize_GetControlState = 52         ' pass 2 params
    ifFileDlgCustomize_SetControlState = 56         ' pass 2 params
    ifFileDlgCustomize_GetEditBoxText = 60          ' pass 2 params
    ifFileDlgCustomize_SetEditBoxText = 64          ' pass 2 params
    ifFileDlgCustomize_GetCheckButtonState = 68     ' pass 2 params
    ifFileDlgCustomize_SetCheckButtonState = 72     ' pass 2 params
    ifFileDlgCustomize_AddControlItem = 76          ' pass 3 params
    ifFileDlgCustomize_RemoveControlItem = 80       ' pass 2 params
    ifFileDlgCustomize_RemoveAllControlItems = 84   ' pass 1 param
    ifFileDlgCustomize_GetControlItemState = 88     ' pass 3 params
    ifFileDlgCustomize_SetControlItemState = 92     ' pass 3 params
    ifFileDlgCustomize_GetSelectedControlItem = 96  ' pass 2 params
    ifFileDlgCustomize_SetSelectedControlItem = 100 ' pass 2 params
    ifFileDlgCustomize_StartVisualGroup = 104       ' pass 2 params
    ifFileDlgCustomize_EndVisualGroup = 108         ' no params
    ifFileDlgCustomize_MakeProminent = 112          ' pass 1 param
    ifShellItem_BindToHandler = 12                  ' pass 4 params
    ifShellItem_GetParent = 16                      ' pass 1 param
    ifShellItem_GetDisplayName = 20                 ' pass 2 params
    ifShellItem_GetAttributes = 24                  ' pass 2 params
    ifShellItem_Compare = 28                        ' pass 2 params
    ifShellItemArray_BindToHandler = 12             ' pass 4 params
    ifShellItemArray_GetPropertyStore = 16          ' pass 3 params
    ifShellItemArray_GetPropertyDescriptionList = 20 ' pass 3 params
    ifShellItemArray_GetAttributes = 24             ' pass 3 params
    ifShellItemArray_GetCount = 28                  ' pass 1 param
    ifShellItemArray_GetItemAt = 32                 ' pass 2 params
    ifShellItemArray_EnumItems = 36                 ' pass 1 param
End Enum

'(**) in the comments indicates that the control does not naturally have a label/caption to
'   identify to the user the purpose of the control. If you want to add a label for the control,
'   the control must be added to a container (ctlTypeOf_Container). Containers have labels and
'   the label will be left of the control, can't be modified. Consider the ctlTypeOf_Container
'   object to be similar to a VB frame
Public Enum CustomControlGroupTypeEnum              ' these contain sub-item objects
    ctlTypeOf_ComboBoxes = 1                        ' subitems are the list items (**)
    ctlTypeOf_RadioButtons = 2                      ' subitems are individual option buttons (**)
    ctlTypeOf_Menus = 3                             ' subitems are submenu items
    ctlTypeOf_OKSplitButton = 4                     ' subitems are submenu items
    ctlTypeOf_Container = 5                         ' subitems are 1 or more CustomControlTypeEnum
End Enum
Public Enum CustomControlTypeEnum                   ' these are stand-alone objects
    ctlTypeOf_CheckBox = 10                         ' includes a caption
    ctlTypeOf_TextBox = 11                          ' width is pretty small, no events (**)
    ctlTypeOf_CommandButton = 12                    ' includes a caption
    ctlTypeOf_StaticLabel = 13                      ' includes a caption, no events
    ctlTypeOf_Separator = 14                        ' drawn diabled, no events, no caption
End Enum


Public Enum CustomControlState
    ctlState_Enabled = 1
    ctlState_Visible = 2
    ctlState_EnabledVisible = 3
End Enum
Public Enum PropertyPathType
    ppType_NotSet = 0
    ppType_AsString = 1
    ppType_AsPIDL = 2
    ppType_AsIShellItemObject = 3
    ppType_AsGUID = 4
End Enum

' http://msdn.microsoft.com/en-us/library/windows/desktop/bb762544%28v=vs.85%29.aspx
Public Enum IShellDisplayNameFormat
  SIGDN_NORMALDISPLAY = &H0&
  SIGDN_PARENTRELATIVEPARSING = &H80018001
  SIGDN_DESKTOPABSOLUTEPARSING = &H80028000
  SIGDN_PARENTRELATIVEEDITING = &H80031001
  SIGDN_DESKTOPABSOLUTEEDITING = &H8004C000
  SIGDN_FILESYSPATH = &H80058000
  SIGDN_URL = &H80068000
  SIGDN_PARENTRELATIVEFORADDRESSBAR = &H8007C001
  SIGDN_PARENTRELATIVE = &H80080001
  SIGDN_PARENTRELATIVEFORUI = &H80094001
'                                       normal file                             virtual folder                              item in virtual folder
'    SIGDN_URL                          file:///C:/WINDOWS/system32/msvcp50.dll ::{26EE0668-A00A-44D7-9371-BEB064C98683}    ::{21EC2020-3AEA-1069-A2DD-08002B30309D}\{E2E7934B-DCE5-43C4-9576-7FE4F75E7480}
'    SIGDN_PARENTRELATIVEPARSING        msvcp50.dll                             ::{26EE0668-A00A-44D7-9371-BEB064C98683}    {E2E7934B-DCE5-43C4-9576-7FE4F75E7480}
'    SIGDN_PARENTRELATIVEFORUI          msvcp50.dll                             Control Panel                               Date and Time
'    SIGDN_PARENTRELATIVEFORADDRESSBAR  msvcp50.dll                             Control Panel                               Date and Time
'    SIGDN_PARENTRELATIVEEDITING        msvcp50.dll                             Control Panel                               Date and Time
'    SIGDN_PARENTRELATIVE               msvcp50.dll                             Control Panel                               Date and Time
'    SIGDN_NORMALDISPLAY                msvcp50.dll                             Control Panel                               Date and Time
'    SIGDN_FILESYSPATH                  C:\WINDOWS\system32\msvcp50.dll         nullstring                                  nullstring
'    SIGDN_DESKTOPABSOLUTEPARSING       C:\WINDOWS\system32\msvcp50.dll         Control Panel                               Date and Time
'    SIGDN_DESKTOPABSOLUTEEDITING       C:\WINDOWS\system32\msvcp50.dll         Control Panel                               Date and Time
End Enum

' http://msdn.microsoft.com/en-us/library/windows/desktop/dn457282%28v=vs.85%29.aspx
Public Enum FileOpenDialogOptions           ' Open/Save pre-Vista flags     PathBrowseDialog
  FOS_OVERWRITEPROMPT = &H2&                ' OFN_OVERWRITEPROMPT
  FOS_STRICTFILETYPES = &H4&                ' OFN_EXTENSIONDIFFERENT
  FOS_NOCHANGEDIR = &H8&                    ' OFN_NOCHANGEDIR
  FOS_PICKFOLDERS = &H20&                   ' BIF_RETURNONLYFSDIRS
  FOS_FORCEFILESYSTEM = &H40&               ' was OFN_ENABLETEMPLATE
  FOS_ALLNONSTORAGEITEMS = &H80&            ' was OFN_ENABLETEMPLATEHANDLE
  FOS_NOVALIDATE = &H100&                   ' OFN_NOVALIDATE
  FOS_ALLOWMULTISELECT = &H200&             ' OFN_ALLOWMULTISELECT
  FOS_PATHMUSTEXIST = &H800&                ' OFN_PATHMUSTEXIST
  FOS_FILEMUSTEXIST = &H1000&               ' OFN_FILEMUSTEXIST
  FOS_CREATEPROMPT = &H2000&                ' OFN_CREATEPROMPT
  FOS_SHAREAWARE = &H4000&                  ' OFN_SHAREAWARE
  FOS_NOREADONLYRETURN = &H8000&            ' OFN_NOREADONLYRETURN
  FOS_NOTESTFILECREATE = &H10000            ' OFN_NOTESTFILECREATE
  FOS_HIDEMRUPLACES = &H20000               ' was OFN_NONETWORKBUTTON
  FOS_HIDEPINNEDPLACES = &H40000            ' was OFN_NOLONGNAMES
  FOS_NODEREFERENCELINKS = &H100000         ' OFN_NODEREFERENCELINKS
  FOS_DONTADDTORECENT = &H2000000           ' OFN_DONTADDTORECENT
  FOS_FORCESHOWHIDDEN = &H10000000          ' OFN_FORCESHOWHIDDEN
  FOS_DEFAULTNOMINIMODE = &H20000000        ' new
  FOS_FORCEPREVIEWPANEON = &H40000000       ' new
  FOS_SUPPORTSTREAMABLEITEMS = &H80000000   ' new
  FOS__SaveFileDefaults = FOS_OVERWRITEPROMPT Or FOS_PATHMUSTEXIST Or FOS_NOREADONLYRETURN
  FOS__OpenFileDefaults = FOS_PATHMUSTEXIST Or FOS_FILEMUSTEXIST
  FOS__BrowseFoldersDefaults = FOS_PICKFOLDERS Or FOS__OpenFileDefaults
End Enum

Public Enum FileDialogModeEnum
    FDLG_FILEOPEN = 0
    FDLG_FILESAVE = 1
    FDLG_BROWSEFOLDERS = 2
End Enum

Private Type ControlInfoStruct
    Text As String
    State As Long
End Type
Private Const CTRLIDOFFSET As Long = 100&
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Any)
Private Declare Function SHCreateItemFromParsingName Lib "shell32.dll" (ByVal pszPath As Long, ByVal pBC As Long, ByVal rIID As Long, pUnk As Long) As Long
Private Declare Function SHParseDisplayName Lib "shell32.dll" (ByVal pszName As Long, ByVal pBC As Long, ByRef ppidl As Long, ByVal sfgaoIn As Long, ByRef psfgaoOut As Long) As Long

'Edit by Tanner: SHILCreateFromPath is deprecated; see my replacement, below
'Private Declare Function SHILCreateFromPath Lib "shell32.dll" (ByVal pszPath As Long, ByRef ppidl As Long, ByRef rgflnOut As Long) As Long
Private Declare Function SHCreateItemFromIDList Lib "shell32.dll" (ByVal pIDL As Long, ByVal rIID As Long, ByRef pUnk As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long

Private m_State As Long                 ' Events requested by client
Private m_Flags As Long                 ' dialog flags: FileOpenDialogOptions enum
Private m_FDlg As IUnknown              ' the IFileDialog object
Private m_FDCustomize As IUnknown       ' the IFileDialogCustomize object
Private m_Events As Object              ' the client, as IFileDialogsEvents implementation
Private m_Result As IUnknown            ' the client-selected files/folders from the dialog
Private m_EventsCookie As Long          ' token used to establish dialog events
Private m_Filters As String             ' dialog filter property, i.e., *.exe;*.dll
Private m_FilterIndex As Long           ' which filter to use if multiple filters exist
Private m_DefaultFolder As Variant      ' which folder to initially display if app's 1st time running
Private m_InitDir As Variant            ' which filter to display (overrides m_DefaultFolder)
Private m_InitFile As String            ' what to initialize the dialog's edit box with
Private m_FileLabel As String           ' the text associated with the dialog's edit box
Private m_OKCaption As String           ' the Open/Save/Select button's caption (nullString = default)
Private m_DefExt As String              ' what default extension to use when saving a file
Private m_ClientGUID As String          ' ability to store/reset cached dialog state (size, position, last folder)
Private m_AddPlaces As Variant          ' a folder that can be added to the dialog's navigation pane
Private m_CtrlSetup() As ControlInfoStruct ' custom control setup info; erased before dialog is displayed
Private m_IDtoKey As Collection         ' custom control keys
Private m_KeyToID As Collection         ' custom control IDs; user never sees these

'--------------------------------------------------------------------------------------------------
' See comments above. The code between the slashes MUST be remarked out if you have NOT
'   added the IFileDialogEvents class to your project. Suggest not deleting the code
'   in case you decide you want to use it later down the road.
' Why rem it out? Because IFileDialogEvents is hard-coded and cannot be declared as Object/IUknown.
' Why can't it be declared as Object? Because we need to access its Friend methods/properties
'   and Friend declarations are not exposed via Object declarations
Public Function AttachEvents(EventReceiver As Object, Optional ByVal Key As String, _
                        Optional ByVal WantMainDialogEvents As Boolean = True, _
                        Optional ByVal WantCustomControlEvents As Boolean = False, _
                        Optional ByVal WantFilterEvents As Boolean) As Boolean
' If events are wanted, this must be called before calling DialogShow
' EventReciever: You must have declared: Implements IFileDialogEvents. Pass ME as the parameter.
' The Key is passed to optional and useful if you call the dialog from multiple points in your
'   project. For example, keys like: "SaveConfig","GetConfig",etc
'   The key is always available from any of the events: Browser.Key
' WantMainDialogEvents: receive all DialogOn[xxx] events from IFileDialogEvents
' WantCustomControlEvents: receive all CustomControl[xxx] events from IFileDialogEvents
' WantFIlterEvents: receive the FilterRequest event (note: depricated as of Win7)
'//////////////////////////////////////////////////////////////////////////////////////////////////
'    Static userKey As String
'    If EventReceiver Is Nothing Then Exit Function
'    If Not (TypeOf EventReceiver Is IFileDialogEvents) Then Exit Function
'    If EventReceiver Is m_Events Then
'        If Not m_Events Is Nothing Then
'            Dim IEvents As IFileDialogEvents
'            If WantMainDialogEvents = True Then
'                Set IEvents = New IFileDialogEvents
'                IEvents.Key = userKey
'                IEvents.InitThunks (m_State And 1), (m_State And 2), (m_State And 4)
'                IEvents.AttachDialog Me, m_Events, m_IDtoKey, m_KeyToID
'                Set m_Events = IEvents
'                If (m_State And 3) Then
'                    pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_Advise, IEvents.GetInterfacePointer(True), VarPtr(m_EventsCookie)
'                End If
'                If (m_State And 4) Then
'                    pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetFilter, IEvents.GetInterfacePointer(, , True)
'                End If
'                userKey = vbNullString
'            ElseIf m_EventsCookie Then
'                Set IEvents = m_Events
'                If m_EventsCookie Then pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_Unadvise, m_EventsCookie
'                If (m_State And 4) Then pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetFilter, 0&
'                Call IEvents.Detach
'                Set m_Events = Nothing
'                m_EventsCookie = 0&
'            End If
'            Set IEvents = Nothing
'        End If
'    ElseIf Not m_Events Is Nothing Then
'        Exit Function
'    ElseIf (WantMainDialogEvents Or WantCustomControlEvents Or WantFilterEvents) = True Then
'        If WantMainDialogEvents Then m_State = m_State Or 1
'        If WantCustomControlEvents Then m_State = m_State Or 2
'        If WantFilterEvents Then m_State = m_State Or 4
'        userKey = Key
'        Set m_Events = EventReceiver
'        AttachEvents = True
'    End If
'//////////////////////////////////////////////////////////////////////////////////////////////////
End Function

'--------------------------------------------------------------------------------------------------
Public Function DialogShow(ByVal ownerHwnd As Long, _
                            Optional ByVal Mode As FileDialogModeEnum = FDLG_FILEOPEN, _
                            Optional dialogTitle As String) As Long
    
    'Added by Tanner: notify the UI subsystem that an OS-owned dialog is active
    Interface.NotifySystemDialogState True
    
    ' The user-selected item(s) are passed via the Result property of this class
    '   Results are in the form of an IShellItem object passed to you
    '       There are a few IShell_[xxx] methods provided so that you can get
    '       the file name and other attributes from the interface. You also have
    '       a variety of choices for how the selected item's text is returned
    '   If user pressed cancel button: return result from this function will be: -2147023673 (&H800704C7)
    '   If dialog failed to be created then the return result from this function will be: -1
    
    ' OwnerHwnd ensures dialog is modal the project. Always pass this as a form hWnd
    '   If passed as zero, user has full access to your project & can even close it while dialog is open,
    '   but when dialog closes, likely your project will re-load but not be shown. Pass the hWnd.

    If Not m_FDlg Is Nothing Then                   ' already opened
        DialogShow = -1&
        Exit Function
    End If
    Set m_Result = Nothing

    Dim hR As Long, lPtr As Long, pIDL As Long
    Dim oIShellItem As IUnknown, pReturn As Long
    Dim aGuid() As Long

    ReDim aGuid(0 To 7)
    ' remove incompatible flags
    If m_Flags And FOS_FORCEFILESYSTEM Then m_Flags = m_Flags And Not FOS_ALLNONSTORAGEITEMS
    
    If Mode = FDLG_BROWSEFOLDERS Then
        If m_Flags = 0& Then m_Flags = FOS__BrowseFoldersDefaults
        m_Flags = m_Flags Or FOS_PICKFOLDERS
        CLSIDFromString StrPtr(clsID_FileOpenDialog), VarPtr(aGuid(0))
        If (m_Flags And FOS_ALLOWMULTISELECT) Then
            IIDFromString StrPtr(IID_IFileOpenDialog), VarPtr(aGuid(4))
        Else ' IID_IFileDialog does not support retrieving multi-selections
            IIDFromString StrPtr(IID_IFileDialog), VarPtr(aGuid(4))
        End If
    Else
        m_Flags = m_Flags And Not FOS_PICKFOLDERS
        If Mode = FDLG_FILEOPEN Then
            If m_Flags = 0& Then m_Flags = FOS__OpenFileDefaults
            CLSIDFromString StrPtr(clsID_FileOpenDialog), VarPtr(aGuid(0))
            IIDFromString StrPtr(IID_IFileOpenDialog), VarPtr(aGuid(4))
        ElseIf Mode = FDLG_FILESAVE Then
            m_Flags = m_Flags And Not FOS_ALLOWMULTISELECT
            If m_Flags = 0& Then m_Flags = FOS__SaveFileDefaults
            CLSIDFromString StrPtr(clsID_FileSaveDialog), VarPtr(aGuid(0))
            IIDFromString StrPtr(IID_IFileSaveDialog), VarPtr(aGuid(4))
        Else
            Call Me.Clear
            DialogShow = -1&
            Exit Function
        End If
    End If
    
    hR = CoCreateInstance(VarPtr(aGuid(0)), 0&, CLSCTX_INPROC_SERVER, VarPtr(aGuid(4)), lPtr)
    If Not hR = 0& Then
        Call Me.Clear
        DialogShow = hR
        Exit Function
    End If
    
    ' option of setting/resetting cached dialog state data
    If Not m_ClientGUID = vbNullString Then
        If m_ClientGUID = "{00000000-0000-0000-0000-000000000000}" Then
            pvCallFunction_COM lPtr, ifFileDialog_ClearClientData
        ElseIf Len(m_ClientGUID) = 38& Or (Len(m_ClientGUID) = 39& And Left$(m_ClientGUID, 1) = "-") Then
            If Left$(m_ClientGUID, 1) = "-" Then
                hR = IIDFromString(StrPtr(m_ClientGUID) + 2&, VarPtr(aGuid(4)))
            Else
                hR = IIDFromString(StrPtr(m_ClientGUID), VarPtr(aGuid(4)))
            End If
            If hR = 0& Then
                pvCallFunction_COM lPtr, ifFileDialog_SetClientGuid, VarPtr(aGuid(4))
                If Left$(m_ClientGUID, 1) = "-" Then pvCallFunction_COM lPtr, ifFileDialog_ClearClientData
            End If
        End If
    End If
    
    Set m_FDlg = pvPointerToIUnknown(lPtr, True)     ' create VB ref to Dialog
    
    If Not m_IDtoKey Is Nothing Then                ' custom controls?
        If m_IDtoKey.Count > 0& Then                ' ensure interface supports it
            IIDFromString StrPtr(IID_IFileDialogCustomize), VarPtr(aGuid(4))
            If pvCallFunction_COM(ObjPtr(m_FDlg), ifUnknown_QueryInterface, VarPtr(aGuid(4)), VarPtr(lPtr)) = 0& Then
                Set m_FDCustomize = pvPointerToIUnknown(lPtr, True) ' create VB ref to customizer
                Call pvInitializeCustomControls     ' add the controls
            End If
        End If
    End If
    If m_FDCustomize Is Nothing Then
        Erase m_CtrlSetup()
        m_State = m_State And Not 2 ' remove callback options for custom controls
    End If
    Erase aGuid()
    
    Dim sFilter() As String
    If m_Filters = vbNullString Then m_Filters = "All Files|*.*"
    sFilter = Split(m_Filters, "|")
    If UBound(sFilter) > 0& Then
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetFileTypes, CLng((UBound(sFilter) + 1) \ 2&), VarPtr(sFilter(0))
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetFileTypeIndex, m_FilterIndex + 1&
    End If
    Erase sFilter()
    
    pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetOptions, m_Flags
    
    Set oIShellItem = pvGetIShellItem(m_DefaultFolder)
    If Not oIShellItem Is Nothing Then
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetDefaultFolder, ObjPtr(oIShellItem)
        Set oIShellItem = Nothing
    End If
    
    Set oIShellItem = pvGetIShellItem(m_InitDir)
    If Not oIShellItem Is Nothing Then
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetFolder, ObjPtr(oIShellItem)
        Set oIShellItem = Nothing
    End If
    
    Set oIShellItem = pvGetIShellItem(m_AddPlaces)
    If Not oIShellItem Is Nothing Then
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_AddPlace, ObjPtr(oIShellItem), (m_State And &H10&) \ &H10&
        Set oIShellItem = Nothing
    End If
    
    If Not m_InitFile = vbNullString Then _
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetFileName, StrPtr(m_InitFile)
    
    If Not dialogTitle = vbNullString Then _
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetTitle, StrPtr(dialogTitle)
    
    If Not m_OKCaption = vbNullString Then _
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetOkButtonLabel, StrPtr(m_OKCaption)
    
    If Not m_FileLabel = vbNullString Then _
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetFileNameLabel, StrPtr(m_FileLabel)
        
    If Not m_DefExt = vbNullString Then _
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_SetDefaultExtension, StrPtr(m_DefExt)
    
    ' create the events callbacks
    If (m_State And 7) Then Call AttachEvents(m_Events, (m_State And 7))
    ' display the dialog
    hR = pvCallFunction_COM(ObjPtr(m_FDlg), ifModalWindow_Show, CLng(ownerHwnd))
    
    ' release the callbacks
    Call AttachEvents(m_Events, , False)
    
    DialogShow = hR
    If hR = 0& Then
        If m_Flags And FOS_ALLOWMULTISELECT Then
            hR = pvCallFunction_COM(ObjPtr(m_FDlg), ifFileDialogOpen_GetResults, VarPtr(pReturn))
            If hR = 0& Then     ' if only 1 item was selected, use IShellItem vs IShellItemArray
                If pvCallFunction_COM(pReturn, ifShellItemArray_GetCount, VarPtr(pIDL)) = 0& Then
                    If pIDL = 1& Then
                        lPtr = 0&
                        pvCallFunction_COM pReturn, ifShellItemArray_GetItemAt, 0&, VarPtr(lPtr)
                        If lPtr Then
                            pvCallFunction_COM pReturn, ifUnknown_Release
                            pReturn = lPtr
                        End If
                    End If
                End If
            End If
        Else
            hR = pvCallFunction_COM(ObjPtr(m_FDlg), ifFileDialog_GetResult, VarPtr(pReturn))
        End If
    End If
    Set m_FDCustomize = Nothing
    Set m_FDlg = Nothing
    Call Me.Clear
    If pReturn Then Set m_Result = pvPointerToIUnknown(pReturn, True)
    
    'Added by Tanner: notify the UI subsystem that an OS-owned dialog is inactive
    Interface.NotifySystemDialogState False
    
End Function
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' User-selected files/folder. Can be more than 1 if multi-selection was opted for
Public Property Get results() As IUnknown
    Set results = m_Result
End Property
' This will return number of selections, if any
Public Property Get ResultsCount() As Long
    If pvIsTypeOf(ObjPtr(m_Result), IID_IShellItemArray) Then
        pvCallFunction_COM ObjPtr(m_Result), ifShellItemArray_GetCount, VarPtr(ResultsCount)
    ElseIf Not m_Result Is Nothing Then
        ResultsCount = 1&
    End If
End Property
' return a specific selection, 1-based, 1-n
Public Property Get ResultsItem(itemIndex As Long) As IUnknown
    Dim lCount As Long, pIShellItem As Long
    lCount = Me.ResultsCount
    If Not (itemIndex < 1& Or itemIndex > lCount) Then
        If lCount > 1& Then
            pvCallFunction_COM ObjPtr(m_Result), ifShellItemArray_GetItemAt, itemIndex - 1&, VarPtr(pIShellItem)
            If pIShellItem Then Set ResultsItem = pvPointerToIUnknown(pIShellItem, True)
        Else
            Set ResultsItem = m_Result
        End If
    End If
End Property
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' when you are retrieving IShellItem interfaces and need to access their methods, these can help
' http://msdn.microsoft.com/en-us/library/windows/desktop/bb761144%28v=vs.85%29.aspx
' Note: The optional ValidatePointer parameters are for safety sake. If you know 100% that the
'   pointers are to IShellItem or IShellItem2 interfaces, then you can set those parameters to
'   false. If you are unsure, do not. The validation verifies that the pointer you passed is of
'   the expected interface. If it wasn't, then we would be calling an expected method on some other
'   interface and a crash is 99% likely.
Public Function IShellItem_GetDisplayName(ByVal oIShellItem As Long, ByVal DisplayFormat As IShellDisplayNameFormat, _
                                          Optional ByVal ValidatePointer As Boolean = True) As String
    If ValidatePointer Then
        If pvIsTypeOf(oIShellItem, IID_IShellItem) = False Then Exit Function
    End If
    Dim lPtr As Long
    If pvCallFunction_COM(oIShellItem, ifShellItem_GetDisplayName, DisplayFormat, VarPtr(lPtr)) = 0& Then
        IShellItem_GetDisplayName = pvComStrPtrToVBstring(lPtr, True)
    ElseIf pvCallFunction_COM(oIShellItem, ifShellItem_GetDisplayName, CLng(SIGDN_NORMALDISPLAY), VarPtr(lPtr)) = 0& Then
        IShellItem_GetDisplayName = pvComStrPtrToVBstring(lPtr, True)
    End If
End Function
Public Function IShellItem_GetParent(ByVal oIShellItem As Long, Optional ByVal ValidatePointer As Boolean = True) As IUnknown
    If ValidatePointer Then
        If pvIsTypeOf(oIShellItem, IID_IShellItem) = False Then Exit Function
    End If
    Dim lPtr As Long
    pvCallFunction_COM oIShellItem, ifShellItem_GetParent, VarPtr(lPtr)
    If lPtr Then Set IShellItem_GetParent = pvPointerToIUnknown(lPtr, True)
End Function
Public Function IShellItem_GetAttributes(ByVal ptrIShellItem As Long, ByVal ShellAttrMask As Long, Optional ByVal ValidatePointer As Boolean = True) As Long
    If ValidatePointer Then
        If pvIsTypeOf(ptrIShellItem, IID_IShellItem) = False Then Exit Function
    End If
    pvCallFunction_COM ptrIShellItem, ifShellItem_GetAttributes, ShellAttrMask, VarPtr(IShellItem_GetAttributes)
End Function
Public Function IShellItem_AreEqual(ByVal ptrIShellItem As Long, ByVal ptrIShellItemCmp As Long, ByVal Hint As Long, Optional ByVal ValidatePointers As Boolean = True) As Boolean
    If ValidatePointers Then
        If pvIsTypeOf(ptrIShellItem, IID_IShellItem) = False Then Exit Function
        If pvIsTypeOf(ptrIShellItemCmp, IID_IShellItem) = False Then Exit Function
    End If
    Dim lResult As Long
    pvCallFunction_COM ptrIShellItem, ifShellItem_Compare, ptrIShellItemCmp, Hint, VarPtr(lResult)
    IShellItem_AreEqual = (lResult = 0&)
End Function
Public Function IShellItem_PointerToVBObject(ByVal ptrIShellItem As Long, Optional ByVal CallReleaseOnPointer As Boolean = True) As IUnknown
    Set IShellItem_PointerToVBObject = pvPointerToIUnknown(ptrIShellItem, CallReleaseOnPointer)
End Function
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' sets the OK button caption: OK Button is the Open/Save/Select button
Public Property Get propOKButtonCaption() As String
    propOKButtonCaption = m_OKCaption
End Property
Public Property Let propOKButtonCaption(newValue As String)
    m_OKCaption = newValue      ' set to null to reset
End Property
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' this is the label to the left of the file selection text box
Public Property Get propFileLabelCaption() As String
    propFileLabelCaption = m_FileLabel
End Property
Public Property Let propFileLabelCaption(newValue As String)
    m_FileLabel = vbNullString      ' set to null to reset
End Property
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' set the folder to be displayed if your app has never ever been run before for this customer
Public Sub propDefaultFolder_Set(Path As Variant, Optional ByVal PathType As PropertyPathType)
    ' can pass: String path, PIDL to path, GUID, or IShellItem as an object
    '   example for Control Panel: {26EE0668-A00A-44D7-9371-BEB064C98683}
    ' you are responsible for destroying any passed PIDLs
    Select Case VarType(Path)
    Case vbDataObject, vbLong, vbEmpty
        m_DefaultFolder = Path
    Case vbString
        If Left$(Path, 1) = "{" Then
            If Right$(Path, 1) = "}" And Len(Path) = 38 Then m_DefaultFolder = Path
        Else
            m_DefaultFolder = Path
        End If
    End Select
End Sub
Public Function propDefaultFolder_Get(Optional PathType As PropertyPathType) As Variant
    propDefaultFolder_Get = m_DefaultFolder
    Select Case VarType(m_DefaultFolder)
        Case vbDataObject: PathType = ppType_AsIShellItemObject
        Case vbString:
            If Left$(m_DefaultFolder, 1) = "{" Then PathType = ppType_AsGUID Else PathType = ppType_AsString
        Case vbLong: PathType = ppType_AsPIDL
        Case Else: PathType = ppType_NotSet
    End Select
End Function
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' set the startup folder for the dialog
Public Sub propStartupFolder_Set(Path As Variant, Optional ByVal PathType As PropertyPathType)
    ' can pass: String path, PIDL to path, GUID, or IShellItem as an object
    '   example for Control Panel: {26EE0668-A00A-44D7-9371-BEB064C98683}
    ' you are responsible for destroying any passed PIDLs
    Select Case VarType(Path)
    Case vbDataObject, vbLong, vbEmpty
        m_InitDir = Path
    Case vbString
        If Left$(Path, 1) = "{" Then
            If Right$(Path, 1) = "}" And Len(Path) = 38 Then m_InitDir = Path
        Else
            m_InitDir = Path
        End If
    End Select
End Sub
Public Function propStartupFolder_Get(Optional PathType As PropertyPathType) As Variant
    propStartupFolder_Get = m_InitDir
    Select Case VarType(m_InitDir)
        Case vbDataObject: PathType = ppType_AsIShellItemObject
        Case vbString
            If Left$(m_InitDir, 1) = "{" Then PathType = ppType_AsGUID Else PathType = ppType_AsString
        Case vbLong: PathType = ppType_AsPIDL
        Case Else: PathType = ppType_NotSet
    End Select
End Function
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Add a "place" to navgiation panel, optionally at the top or bottom of that panel
Public Sub propAddPlace_SetFolder(Path As Variant, Optional ByVal PathType As PropertyPathType, _
                                  Optional ByVal PlaceAtTop As Boolean = True)
    ' can pass: String path, PIDL to path, GUID, or IShellItem as an object
    '   example for Control Panel: {26EE0668-A00A-44D7-9371-BEB064C98683}
    ' you are responsible for destroying any passed PIDLs
    Select Case VarType(Path)
    Case vbDataObject, vbLong, vbEmpty
        m_AddPlaces = Path
    Case vbString
        If Left$(Path, 1) = "{" Then
            If Right$(Path, 1) = "}" And Len(Path) = 38 Then m_AddPlaces = Path
        Else
            m_AddPlaces = Path
        End If
    End Select
    m_State = (m_State And &HFFFFFF0F) Or Abs(PlaceAtTop) * &H10
End Sub
Public Function propAddPlace_FileGetPath(Optional ByRef PlaceAtTop As Boolean, Optional ByVal PathType As PropertyPathType) As Variant
    propAddPlace_FileGetPath = m_AddPlaces
    PlaceAtTop = CBool(m_State And &H10)
    Select Case VarType(m_AddPlaces)
        Case vbDataObject: PathType = ppType_AsIShellItemObject
        Case vbString
            If Left$(m_AddPlaces, 1) = "{" Then PathType = ppType_AsGUID Else PathType = ppType_AsString
        Case vbLong: PathType = ppType_AsPIDL
        Case Else: PathType = ppType_NotSet
    End Select
End Function

'--------------------------------------------------------------------------------------------------
' you can store dialog-specific instances (i.e., last folder, size, posistion, etc). If you opt
'   for this, you must use the same GUID for each save state. By default Windows stores this info
'   under the app's name. But you may want to store this info per instance, i.e., Import, Export,
'   Open, Save, etc.
'--------------------------------------------------------------------------------------------------
' To clear any previous settings you may have had, set this property to the GUID that will be
'   reset, but precede it with a minus sign (-). If you were not using your own GUIDs but still
'   want to reset the data, set this property to: {00000000-0000-0000-0000-000000000000}
Public Property Get propDialogStateGUID() As String
    propDialogStateGUID = m_ClientGUID
End Property
Public Property Let propDialogStateGUID(newValue As String)
    m_ClientGUID = newValue
End Property
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' set dialog restrictions/flags
Public Property Get propFlags() As FileOpenDialogOptions
    propFlags = m_Flags
End Property
Public Property Let propFlags(newValue As FileOpenDialogOptions)
    m_Flags = newValue
End Property
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' this is what appears in the dialog's edit box
Public Property Get propInitialFileName() As String
    propInitialFileName = m_InitFile
End Property
Public Property Let propInitialFileName(newValue As String)
    m_InitFile = newValue
End Property
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' filters are delimited with |
' example: All Files|*.*|Bitmaps|*.bmp|Icons|*.ico;*.cur
Public Property Get propFilters() As String
    propFilters = m_Filters
End Property
Public Property Let propFilters(newValue As String)
    m_Filters = newValue
End Property
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' pass 0-bound filter indexes. Class uses 1-bound but will convert it
Public Property Get propFilterIndex() As Long
    propFilterIndex = m_FilterIndex
End Property
Public Property Let propFilterIndex(newValue As Long)
    m_FilterIndex = newValue
End Property
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' when supplying this DO NOT include the dot (.), i.e., jpg not .jpg
Public Property Get propDefaultExtension() As String
    propDefaultExtension = m_DefExt
End Property
Public Property Let propDefaultExtensions(newValue As String)
    m_DefExt = newValue
End Property
'--------------------------------------------------------------------------------------------------

Public Sub Clear()
    If m_FDlg Is Nothing Then
        Call Me.Customize_ClearCustomizations
        m_State = 0&
        m_Flags = 0&
        m_EventsCookie = 0&
        m_Filters = vbNullString
        m_FilterIndex = 0&
        m_DefaultFolder = Empty
        m_InitDir = Empty
        m_InitFile = vbNullString
        m_FileLabel = vbNullString
        m_OKCaption = vbNullString
        Set m_Result = Nothing
        m_DefExt = vbNullString
        m_ClientGUID = vbNullString
        m_AddPlaces = Empty
    End If
End Sub

Public Sub DialogClose(ByVal ErrorCodeReturned As Boolean)
    ' close the dialog from within called events. Run-time only
    ' if the error code is non-zero, nothing will be assigned to this class' Results property.
    ' However, since you are closing this, the Results property should be empty and passing an
    '   error code is recommended. Your app can process that code as the return value of DialogShow.
    If Not m_FDlg Is Nothing Then pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_Close, ErrorCodeReturned
End Sub

Public Function DialogGetCurrentSelection() As IUnknown
    ' can be called during events to see what the user has curretly selected (file/folder/edit box)
    If Not m_FDlg Is Nothing Then
        Dim pIShellItem As Long
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_GetCurrentSelection, VarPtr(pIShellItem)
        If pIShellItem Then Set DialogGetCurrentSelection = pvPointerToIUnknown(pIShellItem, True)
    End If
End Function

Public Function DialogGetCurrentFolder() As IUnknown
    ' called during events
    If Not m_FDlg Is Nothing Then
        Dim pIShellItem As Long
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_FileGetPath, VarPtr(pIShellItem)
        If pIShellItem Then Set DialogGetCurrentFolder = pvPointerToIUnknown(pIShellItem, True)
    End If
End Function

Public Function DialogGetEditBoxText() As String
    ' called during events
    If Not m_FDlg Is Nothing Then
        Dim lPtr As Long
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_GetFileName, VarPtr(lPtr)
        If lPtr Then DialogGetEditBoxText = pvComStrPtrToVBstring(lPtr, True)
    End If
End Function

Public Function DialogGetFilterIndex() As Long
    ' called during events
    If Not m_FDlg Is Nothing Then
        Dim lItem As Long
        pvCallFunction_COM ObjPtr(m_FDlg), ifFileDialog_GetFileTypeIndex, VarPtr(lItem)
        DialogGetFilterIndex = lItem - 1&
    End If
End Function

Public Function DialogGetHWnd() As Long
    ' Returns the dialog's window handle should you need to subclass any controls on it
    ' May want to call this in your callback DialogOnInit event. Run-time only
    If (Not m_FDlg Is Nothing) Then
        Const IID_OleWindow As String = "{00000114-0000-0000-C000-000000000046}"
        Dim aGuid(0 To 3) As Long, pOLE As Long
        IIDFromString StrPtr(IID_OleWindow), VarPtr(aGuid(0))
        pvCallFunction_COM ObjPtr(m_FDlg), ifUnknown_QueryInterface, VarPtr(aGuid(0)), VarPtr(pOLE)
        If pOLE Then
            pvCallFunction_COM pOLE, ifOleWindow_GetWindow, VarPtr(DialogGetHWnd)
            pvCallFunction_COM pOLE, ifUnknown_Release
        End If
    End If
End Function

Public Function Customize_CreateContainer(ByVal Key As String, ByVal Label As String) As Boolean
' All future Customize_Create[xxx] methods are applied to this container until Customize_CloseContainer called
' Changing the state of a container, changes the state on all controls contained in the container.
    If Not m_FDlg Is Nothing Then Exit Function     ' design-time only
    If pvValidateKey(Key, ctlTypeOf_Container, 0) Then
        m_CtrlSetup(m_IDtoKey.Count).Text = Label
        Customize_CreateContainer = True
    End If
End Function
Public Function Customize_CloseContainer(ByVal ContainerKey As String) As Boolean
    ' call to stop appending controls to this container
    If Not m_FDlg Is Nothing Then Exit Function     ' design-time only
    Dim lType As Long, dwCtrlID As Long
    dwCtrlID = pvIDfromKey(ContainerKey, , lType)
    If lType = ctlTypeOf_Container Then
        If pvValidateKey(" " & ContainerKey, 0, 0, True) Then
            m_CtrlSetup(m_IDtoKey.Count).State = dwCtrlID
        End If
    End If
End Function

Public Function Customize_CreateGroup(ByVal GroupType As CustomControlGroupTypeEnum, ByVal GroupKey As String, _
                                     ByVal MenuCaption As String, ParamArray GroupItems() As Variant) As Boolean
    ' GroupKey is required and must not have been used for any other controls
    ' GroupType will be combobox, menu, option button group
    ' MenuCaption applies to menu-controls only
    '     Labels can have mnemonics
    ' GroupItems are a comma-delimited list of string values
    If Not m_FDlg Is Nothing Then Exit Function     ' design-time only
    If GroupType < ctlTypeOf_ComboBoxes Or GroupType > ctlTypeOf_OKSplitButton Then Exit Function
    
    Dim ListCount As Long, lItem As Long, sPath As String, pRevert As Long
    Dim lID As Long, sID As String * 10
    If IsMissing(GroupItems) = False Then ListCount = UBound(GroupItems) + 1
    If pvValidateKey(GroupKey, GroupType, ListCount) = False Then Exit Function
    On Error GoTo ExitRoutine
    pRevert = m_IDtoKey.Count
    sID = "&H" & Hex$(pRevert + CTRLIDOFFSET Or GroupType * &H1000000) & "&"
    sPath = sID
    For lItem = 0& To ListCount - 1&
        lID = pRevert + CTRLIDOFFSET + lItem + 1&: sID = "&H" & Hex$(lID) & "&"
        m_IDtoKey.Add GroupKey, CStr(lID)    ' ID > parent key
        sPath = sPath & sID    ' Add child key in children list
        m_CtrlSetup(m_IDtoKey.Count).Text = GroupItems(lItem)
    Next
    m_KeyToID.Remove GroupKey
    m_KeyToID.Add sPath, GroupKey    ' add parent/children list
    With m_CtrlSetup(pRevert)
        If GroupType = ctlTypeOf_Menus Then .Text = GroupKey
        .Text = MenuCaption
    End With
    Customize_CreateGroup = True
ExitRoutine:
    If Err Then
        Err.Clear
        For pRevert = m_IDtoKey.Count To pRevert Step -1
            m_IDtoKey.Remove pRevert
        Next
    End If
End Function

Public Function Customize_AppendGroupItem(ByVal GroupKey As String, ByVal Item As String) As Boolean
' Run/Design-time appending of combobox,menu,radio group items
    Dim sPath As String, sID As String * 10
    On Error GoTo ExitRoutine            ' should bad key be passed
    If VarType(m_KeyToID(GroupKey)) = vbString Then
        If pvValidateKey(vbNullChar, 0&, 0&) = False Then Exit Function
        m_IDtoKey.Add GroupKey, CStr(m_IDtoKey.Count + CTRLIDOFFSET + 1&)
        sID = "&H" & Hex$(m_IDtoKey.Count + CTRLIDOFFSET) & "&"
        sPath = m_KeyToID(GroupKey) & sID
        m_KeyToID.Remove GroupKey
        m_KeyToID.Add sPath, GroupKey
        If m_FDCustomize Is Nothing Then
            ReDim Preserve m_CtrlSetup(1 To m_IDtoKey.Count)
            m_CtrlSetup(m_IDtoKey.Count).Text = Item
        Else ' run-time add it now
            pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_AddControlItem, CLng(Val(Left$(sPath, 10))) And &HFFFFFF, m_IDtoKey.Count + CTRLIDOFFSET, StrPtr(Item)
            ' container ID is Val(Left$(sPath,10)), itemID is m_IDtoKey.Count+100
        End If
        Customize_AppendGroupItem = True
    End If
ExitRoutine:
End Function

Public Function Customize_RemoveGroupItem(ByVal GroupKey As String, ByVal itemIndex As Long) As Boolean
' Run-time removing of combobox items & submenu items
    If m_FDCustomize Is Nothing Then Exit Function
    If itemIndex < 0& Then Exit Function
    
    Dim sPath As String, pIndex As Long
    On Error GoTo ExitRoutine            ' should bad key be passed
    If VarType(m_KeyToID(GroupKey)) = vbString Then
        sPath = m_KeyToID(GroupKey)
        If itemIndex < Len(sPath) \ 10 Then
            pIndex = Val(Mid$(sPath, (itemIndex + 1&) * 10, 10))
            sPath = Left$(sPath, (itemIndex + 1&) * 10& - 1&) & Mid$(sPath, (itemIndex + 2&) * 10&)
            m_KeyToID.Remove GroupKey
            m_KeyToID.Add sPath, GroupKey
            pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_RemoveControlItem, CLng(Val(Left$(sPath, 10))) And &HFFFFFF, pIndex
            Customize_RemoveGroupItem = True
        End If
    End If
ExitRoutine:
End Function

Public Function Customize_CreateNonGroupItem(ByVal ItemType As CustomControlTypeEnum, ByVal itemKey As String, _
                    Optional ByVal ItemTextOrCaption As String, Optional ByVal isChecked As Boolean) As Boolean
    ' ItemKey is required and must not have been used for any other controls
    ' ItemType will be cmdbutton,textbox,label,checkbox,separator
    ' ItemTextOrCaption can be blank. Ignored for separators. It is the Caption or Text property value
    ' Labels can have mnemonics
    If Not m_FDlg Is Nothing Then Exit Function     ' design-time only
    If ItemType < ctlTypeOf_CheckBox Or ItemType > ctlTypeOf_Separator Then Exit Function
    If pvValidateKey(itemKey, ItemType, 0) = False Then Exit Function
    On Error GoTo ExitRoutine
    With m_CtrlSetup(m_IDtoKey.Count)
        .Text = ItemTextOrCaption
        If ItemType = ctlTypeOf_CheckBox Then .State = Abs(isChecked)
    End With
    Customize_CreateNonGroupItem = True
ExitRoutine:
End Function

Public Property Let Customize_MakeControlProminent(ByVal ControlKey As String)
    ' the prominent control (can only be one) is placed next to the Open/Save button
    If Not m_FDlg Is Nothing Then Exit Property     ' design-time only
    Dim dwCtrlID As Long, lType As Long
    dwCtrlID = pvIDfromKey(ControlKey, , lType)
    If dwCtrlID Then
        Select Case lType   ' per MSDN, only these types of control can become prominent
        Case ctlTypeOf_CheckBox, ctlTypeOf_ComboBoxes, ctlTypeOf_CommandButton, ctlTypeOf_Menus
            With m_CtrlSetup(dwCtrlID - CTRLIDOFFSET)
                .State = .State Or &H10000000
            End With
        Case ctlTypeOf_Container ' or a container if it has just 1 of the above
            With m_CtrlSetup(pvIDfromKey(" " & ControlKey) - CTRLIDOFFSET)
                .State = .State Or &H10000000
            End With
        End Select
    End If
End Property

Public Function CustomControl_GetTextBoxText(ByVal Key As String) As String
' no other control types have supported methods to return captions or other text
    If m_FDCustomize Is Nothing Then Exit Function
    Dim dwCtrlID As Long, lType As Long
    dwCtrlID = pvIDfromKey(Key, , lType)
    If dwCtrlID <> 0& And lType = ctlTypeOf_TextBox Then
        lType = 0&
        pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_GetEditBoxText, dwCtrlID, VarPtr(lType)
        If lType Then CustomControl_GetTextBoxText = pvComStrPtrToVBstring(lType)
    End If
End Function
Public Function CustomControl_SetTextBoxText(ByVal Key As String, ByVal Text As String) As String
' no other control types have supported methods to set control captions or other text
    If m_FDCustomize Is Nothing Then Exit Function
    Dim dwCtrlID As Long, lType As Long
    dwCtrlID = pvIDfromKey(Key, , lType)
    If dwCtrlID <> 0& And lType = ctlTypeOf_TextBox Then
        pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_SetEditBoxText, dwCtrlID, StrPtr(Text)
    End If
End Function

Public Function CustomControl_SetControlCaption(ByVal Key As String, ByVal Caption As String) As String
' applies to: menu (not submenus), static label, checkbox, command button
    If m_FDCustomize Is Nothing Then Exit Function
    Dim dwCtrlID As Long, lType As Long
    dwCtrlID = pvIDfromKey(Key, , lType)
    If dwCtrlID <> 0& Then        ' call SetControlLabel
        pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_SetControlLabel, dwCtrlID, StrPtr(Caption)
    End If
End Function

Public Function CustomControl_GetState(ByVal Key As String, Optional ByVal itemIndex As Long = -1&) As CustomControlState
    If m_FDCustomize Is Nothing Then Exit Function
    Dim dwCtrlID As Long, lType As Long
    dwCtrlID = pvIDfromKey(Key, itemIndex, lType)
    If dwCtrlID <> 0& Then
        If itemIndex < 0& Then
            pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_GetControlState, dwCtrlID, VarPtr(CustomControl_GetState)
            ' call GetControlState
        Else
            pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_GetControlItemState, dwCtrlID, itemIndex, VarPtr(CustomControl_GetState)
            ' call GetControlItemState
        End If
    End If
End Function
Public Function CustomControl_SetState(ByVal Key As String, ByVal Enabled As Boolean, ByVal Visible As Boolean, _
                                        Optional ByVal itemIndex As Long = -1&) As Boolean
    ' Note: for setting radio buttons, function is returning E_NOTIMPL
    If m_FDCustomize Is Nothing Then Exit Function
    Dim dwCtrlID As Long, lType As Long
    dwCtrlID = pvIDfromKey(Key, itemIndex, lType)
    If dwCtrlID <> 0& Then
        lType = Abs(Enabled)
        If Visible Then lType = lType Or ctlState_Visible
        If itemIndex < 0& Then
            CustomControl_SetState = (pvCallFunction_COM(ObjPtr(m_FDCustomize), ifFileDlgCustomize_SetControlState, dwCtrlID, lType) = 0&)
        Else
            CustomControl_SetState = (pvCallFunction_COM(ObjPtr(m_FDCustomize), ifFileDlgCustomize_SetControlItemState, dwCtrlID, itemIndex, lType) = 0&)
        End If
    End If
End Function

Public Function CustomControl_GetSelected(ByVal Key As String) As Long
    ' return checkbox, which combobox item, or radio button index is selected (not menus)
    ' custom menus selections must be trapped in the control events, not dialog events
    If m_FDCustomize Is Nothing Then Exit Function
    Dim dwCtrlID As Long, lType As Long
    dwCtrlID = pvIDfromKey(Key, , lType)
    If dwCtrlID = 0& Then
        CustomControl_GetSelected = -1&
    Else
        If lType = ctlTypeOf_RadioButtons Or lType = ctlTypeOf_ComboBoxes Or lType = ctlTypeOf_OKSplitButton Then
            lType = 0&
            If pvCallFunction_COM(ObjPtr(m_FDCustomize), ifFileDlgCustomize_GetSelectedControlItem, dwCtrlID, VarPtr(lType)) = 0& Then
                Dim lPos As Long
                Key = m_KeyToID(Key)
                For lPos = 11 To Len(Key) Step 10
                    If Val(Mid$(Key, lPos, 10)) = lType Then
                        CustomControl_GetSelected = lPos \ 10 - 1&
                        Exit Function
                    End If
                Next
            End If
            CustomControl_GetSelected = -1&
        ElseIf ctlTypeOf_CheckBox Then
            pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_GetCheckButtonState, dwCtrlID, VarPtr(CustomControl_GetSelected)
        End If
    End If
End Function
Public Function CustomControl_SetSelected(ByVal Key As String, ByVal itemIndex As Long, Optional ByVal ChkBoxSelected As Boolean) As Long
    ' set checkbox, which combobox item, or radio button index is selected (not menus)
    ' custom menus selections must be trapped in the control events, not dialog events
    If m_FDCustomize Is Nothing Then Exit Function
    Dim dwCtrlID As Long, lType As Long
    dwCtrlID = pvIDfromKey(Key, itemIndex, lType)
    If Not dwCtrlID = 0& Then
        If lType = ctlTypeOf_RadioButtons Or lType = ctlTypeOf_ComboBoxes Then
            pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_SetSelectedControlItem, dwCtrlID, itemIndex
        ElseIf ctlTypeOf_CheckBox Then
            pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_SetCheckButtonState, dwCtrlID, Abs(ChkBoxSelected)
        End If
    End If
End Function

Public Sub Customize_ClearCustomizations()
    ' Start over. Ignored during runtime
    If m_FDCustomize Is Nothing Then
        Set m_IDtoKey = Nothing
        Set m_KeyToID = Nothing
        Erase m_CtrlSetup()
    End If
End Sub

Private Function pvIDfromKey(Key As String, Optional ListIndex As Long = -1&, _
                            Optional CtrlType As Long) As Long
' called from within the browser to return dwCtrlID from Key and/or ListIndex
' method used when user wants to let/get control properties
    On Error GoTo ExitRoutine  ' should bad Key be passed
    If VarType(m_KeyToID(Key)) = vbString Then
        Dim sData As String
        sData = m_KeyToID(Key)
        pvIDfromKey = Val(Left$(sData, 10))
        If ListIndex > -1& Then
            If ListIndex < Len(sData) \ 10 Then
                ListIndex = Val(Mid$(sData, (ListIndex + 1) * 10 + 1, 10))
            Else
                ListIndex = -1&
            End If
        End If
    Else
        ListIndex = -1&
        pvIDfromKey = m_KeyToID(Key)
    End If
    CtrlType = pvIDfromKey \ &H1000000
    pvIDfromKey = pvIDfromKey And &HFFFFFF
ExitRoutine:
End Function

Private Function pvValidateKey(Key As String, CtrlType As Long, SubItems As Long, Optional AllowLeadingSpace As Boolean = False) As Boolean
    On Error Resume Next
    If Key = vbNullString Then Exit Function
    If Left$(Key, 1) = " " Then
        If AllowLeadingSpace = False Then Exit Function
    End If
    If m_KeyToID Is Nothing Then
        Set m_KeyToID = New Collection
        Set m_IDtoKey = New Collection
    ElseIf m_IDtoKey.Count = &HFFFFFF Then
        Exit Function    ' maxed out
    End If
    If Key = vbNullChar Then
        pvValidateKey = True
    Else
        m_KeyToID.Add (m_IDtoKey.Count + CTRLIDOFFSET + 1&) Or CtrlType * &H1000000, Key
        If Err Then
            Err.Clear
        Else
            m_IDtoKey.Add Key, CStr(m_IDtoKey.Count + CTRLIDOFFSET + 1&)
            ReDim Preserve m_CtrlSetup(1 To m_IDtoKey.Count + SubItems)
            pvValidateKey = True
        End If
    End If
End Function

Private Sub pvInitializeCustomControls()
    Dim c As Long, lID As Long, lType As Long, lpID As Long
    Dim sPath As String, p As Long, vCount As Long
    
    If m_KeyToID Is Nothing Then Exit Sub
    For c = 1 To m_KeyToID.Count
        sPath = m_KeyToID.Item(c)
        If Left$(sPath, 1) = "&" Then      ' container
            lID = Val(Left$(sPath, 10))
        Else
            lID = Val(sPath)
        End If
        lType = lID \ &H1000000
        lID = lID And &HFFFFFF
        With m_CtrlSetup(lID - CTRLIDOFFSET)
            Select Case lType
            Case 0
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_EndVisualGroup
                vCount = vCount - 1&
                lID = .State And &HFFFFFF
            Case ctlTypeOf_Container
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_StartVisualGroup, lID, StrPtr(.Text)
                vCount = vCount + 1&
            Case ctlTypeOf_CheckBox
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_AddCheckButton, lID, StrPtr(.Text), .State
            Case ctlTypeOf_ComboBoxes
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_AddComboBox, lID
            Case ctlTypeOf_CommandButton
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_AddPushButton, lID, StrPtr(.Text)
            Case ctlTypeOf_Separator
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_AddSeparator, lID
            Case ctlTypeOf_StaticLabel
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_AddText, lID, StrPtr(.Text)
            Case ctlTypeOf_RadioButtons
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_AddRadioButtonList, lID
            Case ctlTypeOf_TextBox
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_AddEditBox, lID, StrPtr(.Text)
            Case ctlTypeOf_Menus
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_AddMenu, lID, StrPtr(.Text)
            Case ctlTypeOf_OKSplitButton
                pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_EnableOpenDropDown, lID
            End Select
            If (.State And &H10000000) Then pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_MakeProminent, lID
        End With
        If Left$(sPath, 1) = "&" Then
            For p = 11& To Len(sPath) Step 10&
                lpID = Val(Mid$(sPath, p, 10))
                With m_CtrlSetup(lpID - CTRLIDOFFSET)
                   pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_AddControlItem, lID, lpID, .Text
                End With
            Next
        End If
    Next
    For vCount = 1& To vCount
        pvCallFunction_COM ObjPtr(m_FDCustomize), ifFileDlgCustomize_EndVisualGroup
    Next
    Erase m_CtrlSetup()
End Sub
Private Function pvCallFunction_COM(ByVal InterfacePointer As Long, ByVal VTableOffset As InterfaceMethodOffsets, _
                            ParamArray FunctionParameters() As Variant) As Variant
                            
' Used to call active-x or COM objects, not standard dlls
' Return value. Will be a variant containing a value of FunctionReturnType
'   If this method fails, the return value will always be Empty. This can be verified by checking
'       the Err.LastDLLError value. It will be non-zero if the function failed else zero.
'   If the method succeeds, there is no guarantee that the Interface function you called succeeded. The
'       success/failure of that function would be indicated by this method's return value.
'       Typically, success is returned as S_OK (zero) and any other value is an error code.
'   If calling a sub vs function & this method succeeds, the return value will be zero.
'   Summarizing: if method fails to execute, Err.LastDLLError value will be non-zero
'       If method executes ok, if the return value is zero, method succeeded else return is error code

' Parameters:
'   InterfacePointer. A pointer to an object/class, i.e., ObjPtr(IPicture)
'       Passing invalid pointers likely to result in crashes
'   VTableOffset. The offset from the passed InterfacePointer where the virtual function exists.
'       The value is in bytes. These offsets are generally in multiples of 4. Value cannot be negative.
'       Example: to call IUnknown:Release, CallFunction_COM InterfacePointer, 8&, CR_LONG, CC_STDCALL

    '// minimal sanity check for these 3 parameters:
    If VTableOffset < 0& Then Exit Function

    Dim pIndex As Long, pCount As Long
    Dim vParamPtr() As Long, vParamType() As Integer
    Dim vRtn As Variant, vParams() As Variant
    Const CallConvention As Long = 4&                   ' STDCALL
    
    vParams() = FunctionParameters()                    ' copy passed parameters, if any
    pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
    If pCount = 0& Then                                 ' no return value (sub vs function)
        ReDim vParamPtr(0 To 0)
        ReDim vParamType(0 To 0)
    Else
        ReDim vParamPtr(0 To pCount - 1&)               ' need matching array of parameter types
        ReDim vParamType(0 To pCount - 1&)              ' and pointers to the parameters
        For pIndex = 0& To pCount - 1&
            vParamPtr(pIndex) = VarPtr(vParams(pIndex))
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next
    End If
                                                        ' call the function now
    pIndex = DispCallFunc(InterfacePointer, VTableOffset, CallConvention, vbLong, _
                          pCount, vParamType(0), vParamPtr(0), vRtn)
    If pIndex = 0& Then                                 ' 0 = S_OK
        pvCallFunction_COM = vRtn                       ' return result
    Else
        SetLastError pIndex                             ' set error & return Empty
    End If

End Function

Private Function pvComStrPtrToVBstring(pString As Long, Optional FreePointer As Boolean = True) As String
    Dim lLen As Long
    If Not pString = 0& Then
        lLen = lstrlenW(pString)
        If lLen > 0& Then
            pvComStrPtrToVBstring = String$(lLen, vbNullChar)
            CopyMemoryStrict StrPtr(pvComStrPtrToVBstring), pString, lLen * 2&
        End If
        If FreePointer Then CoTaskMemFree pString
    End If
End Function

Private Function pvGetIShellItem(Source As Variant) As IUnknown
    ' Source can contain a path in 1 of 4 formats:
    ' 1) plain text path
    ' 2) PIDL to the path
    ' 3) an IShellItem VB.IUnknown object
    ' 4) a GUID representing a virtual folder
    On Error GoTo ExitRoutine
    If IsEmpty(Source) Then Exit Function
    If IsObject(Source) Then
        Set pvGetIShellItem = Source
    Else
        Dim aGuid(0 To 3) As Variant, pIDL As Long, lPtr As Long
        IIDFromString StrPtr(IID_IShellItem), VarPtr(aGuid(0))
        If VarType(Source) = vbString Then
            If Left$(Source, 1) = "{" And Right$(Source, 1) = "}" And Len(Source) = 38& Then
                If SHCreateItemFromParsingName(StrPtr("::" & Source), 0&, VarPtr(aGuid(0)), lPtr) = 0& Then
                    Set pvGetIShellItem = pvPointerToIUnknown(lPtr, True)
                End If
                Exit Function
            ElseIf Not Len(Source) = 0& Then
                'Call SHILCreateFromPath(StrPtr(CStr(Source)), pIDL, ByVal 0&)
                'Modified by Tanner: SHILCreateFromPath is deprecated; use the suggested replacement function, instead
                Call SHParseDisplayName(StrPtr(CStr(Source)), 0&, pIDL, 0&, ByVal 0&)
            End If
        ElseIf VarType(Source) = vbLong Then
            pIDL = Source
        End If
        If pIDL Then
            SHCreateItemFromIDList pIDL, VarPtr(aGuid(0)), lPtr
            Set pvGetIShellItem = pvPointerToIUnknown(lPtr, True)
            If VarType(Source) = vbString Then CoTaskMemFree pIDL
        End If
    End If
ExitRoutine:
End Function

Private Function pvPointerToIUnknown(pObject As Long, CallRelease As Boolean) As IUnknown
    If pObject = 0& Then Exit Function
    Dim tObj As IUnknown
    CopyMemory tObj, pObject, 4&
    Set pvPointerToIUnknown = tObj
    CopyMemory tObj, 0&, 4&
    If CallRelease Then pvCallFunction_COM pObject, ifUnknown_Release
End Function

Private Function pvIsTypeOf(pPointer As Long, isWhatIID As String) As Boolean
    If pPointer = 0& Then Exit Function
    Dim aGuid(0 To 3) As Long, lPtr As Long
    IIDFromString StrPtr(isWhatIID), VarPtr(aGuid(0))
    If pvCallFunction_COM(pPointer, ifUnknown_QueryInterface, VarPtr(aGuid(0)), VarPtr(lPtr)) = 0& Then
        pvCallFunction_COM lPtr, ifUnknown_Release
        pvIsTypeOf = (pPointer = lPtr)
    End If
End Function

Private Sub Class_Terminate()
    Call Me.Clear
End Sub
