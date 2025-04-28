VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUnicodeBrowseFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'You may download the original version of this class from the following link (good as of May '16):
' http://www.vbforums.com/showthread.php?585762-VB6-Unicode-File-Open-Save-Dialog

'This class was originally written by vbforums.com user LaVolpe.  To the best of my knowledge, the code has
' been released into the public domain.

'Many thanks to LaVolpe for this comprehensive folder browsing interface, which PhotoDemon uses on XP installs.

Option Explicit

'  -----======== PURPOSE: A unicode aware/compatible browse-for-folder dialog class ========-----
' Class Instancing = MultiUse (can be created by users)

'//////////////////// APIs for displaying Browse for Folder dialog \\\\\\\\\\\\\\\\\\\\\\\
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private Declare Function SHBrowseForFolderA Lib "shell32" (lpBrowseInfo As BrowseInfo) As Long
Private Declare Function SHBrowseForFolderW Lib "shell32" (ByVal lpBrowseInfo As Long) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pidList As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As Long) As Long
Private Declare Function SHILCreateFromPath Lib "shell32" (ByVal pszPath As Long, ByRef ppidl As Long, ByRef rgflnOut As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long

'//// Required: API declarations needed for creating callback thunk as needed (see pvInitStructure)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDst As Any, ByRef lpSrc As Any, ByVal byteLength As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
'////////////////////////////////////////////////////////////////////////////
Private Const WM_USER As Long = &H400&
Private Const BFFM_INITIALIZED As Long = 1&
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Private Const MAX_PATH As Long = 260&
Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfnCallback As Long
   lpCallbackData As Long
   iImage As Long
End Type

Public Enum FolderBrowseFlagsEnum
    ' full descriptions: http://msdn.microsoft.com/en-us/library/windows/desktop/bb773205%28v=vs.85%29.aspx
    BIF_RETURNONLYFSDIRS = &H1
        'Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
    BIF_DONTGOBELOWDOMAIN = &H2
        'Do not include network folders below the domain level in the dialog box's tree view control.
    BIF_STATUSTEXT = &H4
        'Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box
    BIF_RETURNFSANCESTORS = &H8
        'Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy
    BIF_EDITBOX = &H10
        'Include an edit control in the browse dialog box that allows the user to type the name of an item.
    BIF_VALIDATE = &H20
        'If the user types an invalid name into the edit box, the browse dialog box calls the application's BrowseCallbackProc
    BIF_NEWDIALOGSTYLE = &H40
        'Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized.
    BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
        'Use the new user interface, including an edit box
    BIF_BROWSEINCLUDEURLS = &H80
        'The browse dialog box can display URLs
    BIF_UAHINT = &H100
        'When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box, in place of the edit box.
    BIF_NONEWFOLDERBUTTON = &H200
        'Do not include the New Folder button in the browse dialog box
    BIF_NOTRANSLATETARGETS = &H400
        'When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target.
    BIF_BROWSEFORCOMPUTER = &H1000
        'Only return computers. If the user selects anything other than a computer, the OK button is grayed
    BIF_BROWSEFORPRINTER = &H2000
        'Only allow the selection of printers. If the user selects anything other than a printer, the OK button is grayed.
    BIF_BROWSEINCLUDEFILES = &H4000
        'The browse dialog box displays files as well as folders.
    BIF_SHAREABLE = &H8000&
        'The browse dialog box can display shareable resources on remote systems
    BIF_BROWSEFILEJUNCTIONS = &H10000
        'Windows 7 and later. Allow folder junctions such as a library or a compressed file with a .zip file name extension to be browsed.
End Enum

Private obif As BrowseInfo      ' current structure used for dialog
Private m_State As Long         ' class-only flags/settings
' 0x00000001                        unicode system
' 0x00000002                        cancel error property true
' 0x00000004                        ShowBrowseForFolder:ReleasePIDL=false
' 0x00000008                        PIDL used for Root folder is owned by this class
Private m_PIDL As Long          ' return value of the ShowBrowseForFolder function
Private m_File As String        ' file name returned by ShowBrowseForFolder function, if any
Private m_ThunkAddr As Long     ' allocated memory address if callback thunk was created
Private m_InitPath As String    ' startup path for the dialog, either string or PIDL

Public Function ShowBrowseForFolder(ByVal ownerHwnd As Long, Optional releasePIDL As Boolean = True) As Boolean
    
    'Added by Tanner: notify the UI subsystem that an OS-owned dialog is active
    Interface.NotifySystemDialogState True
    
    ' Depending on what flags you set, you may get a returned path and/or file name or nothing.
    ' If the user makes a selection to a virtual path/object, you will not get a path returned.
    
    ' The API this function calls returns a PIDL:
    ' However, in all cases where the function does not fail, a PIDL will be returned, this PIDL
    '   may be to a virtual path/object like the Control Panel for example. The PIDL will remain
    '   in this class until class is destroyed or this function is called again or you destroy the PIDL.
    ' If ReleasePIDL parameter is False, then you MUST destroy the PIDL before this class is destroyed
    '   or this function is called again. If that parameter is True, you should NOT destroy the PIDL.
    '   If you are required to destroy the PIDL, you can destroy it manually or by calling this class'
    '   DestroyPIDL method and passing the PIDL value.
    ' The PIDL can be retrieved via the SelectedPIDL() property. That property has an optional parameter
    '   that will override the ReleasePIDL parameter passed to this function. If that optional parameter
    '   is set to True, you own the PIDL regardless of the ReleasePIDL parameter. This can be useful in
    '   this scenario: User selected a virtual folder but you passed the ReleasePIDL parameter as True.
    '   This means next time you call this function, it will be destroyed. But you want to set the
    '   InitialDirectory property to the PIDL that was returned here. Since you opted to release it in
    '   in the previous call, it is destroyed as soon as this method is called. By calling
    '   SelectedPIDL(True), you take ownership of the PIDL and this class will not destroy it. However,
    '   you are not responsible for destroying it at some point.
    
    Dim lLen As Long
   
    Call pvInitStructure(ownerHwnd)                 ' ensure structure is configured correctly
    If releasePIDL = False Then m_State = m_State Or 4& ' set flag whether to manage the returned PIDL
    
    If (m_State And 1&) = 1& Then                   ' call unicode or ANSI method as needed
        m_PIDL = SHBrowseForFolderW(VarPtr(obif))
    Else
        m_PIDL = SHBrowseForFolderA(obif)
    End If
    
    If m_PIDL = 0& Then                             ' failure or user canceled
        If (m_State And 2&) = 2& Then Err.Raise 18 + vbObjectError, "ShowBrowseForFolder", "Canceled by user"
    Else
        m_File = String$(MAX_PATH + 2&, vbNullChar) ' get the returned path if any (may be PIDL only return)
        If (m_State And 1&) = 1& Then
            lLen = SHGetPathFromIDListW(m_PIDL, StrPtr(m_File))
        Else
            lLen = SHGetPathFromIDListA(m_PIDL, m_File)
        End If
        If lLen Then
            lLen = lstrlenW(StrPtr(m_File))
            m_File = Left$(m_File, lLen)
        Else
            m_File = vbNullString       ' virtual path? PIDL valid, but no path returned, i.e., Control Panel
        End If
        ShowBrowseForFolder = True
    End If
    If Not m_ThunkAddr = 0& Then        ' if we created a callback procedure, remove it
        Const CB_MEM_DECOMMIT  As Long = &H4000&
        Const CB_MEM_RELEASE   As Long = &H8000&
        Const FUNK_LEN         As Long = 48&
        If Not VirtualFree(m_ThunkAddr, FUNK_LEN, CB_MEM_DECOMMIT) = 0& Then
          Call VirtualFree(m_ThunkAddr, 0&, CB_MEM_RELEASE)
          m_ThunkAddr = 0&
        End If
        obif.lpCallbackData = 0&
        obif.lpfnCallback = 0&
    End If
    
    'Added by Tanner: notify the UI subsystem that an OS-owned dialog is inactive
    Interface.NotifySystemDialogState False
    
End Function

' forces an error to occur if user cancels/closes dialog without selecting anything
Public Property Get CancelError() As Boolean
    CancelError = (m_State And 2&) = 2&
End Property
Public Property Let CancelError(ByVal cErr As Boolean)
    m_State = (m_State And &HFFFFFFFD) Or Abs(cErr * 2&)
End Property

Public Sub Clear()
    ' clears all properties and resets to default
    Dim blankOFN As BrowseInfo
    Me.SetRootFolderPIDL 0&
    If Not m_PIDL = 0& Then
        If (m_State And 4&) = 0& Then Me.DestroyPIDL m_PIDL
    End If
    obif = blankOFN
    m_InitPath = vbNullString
    m_State = m_State And 1&                    ' remove all flags except Unicode system
End Sub

' returns/sets custom value sent to hook procedure
Public Property Get CustomHookData() As Long
    CustomHookData = obif.lpCallbackData
End Property
Public Property Let CustomHookData(ByVal nPointer As Long)
    ' If you want to use your own Callback/subclass routine for this browser, then
    '   this value is what is passed to to the final parameter of the callback procedure
    ' This value is only applicable if the CustomHookProc property is also set.
    ' The callback procedure definition can be found here:
    '   http://msdn.microsoft.com/en-us/library/windows/desktop/bb762598%28v=vs.85%29.aspx
    obif.lpCallbackData = nPointer
End Property

' returns/sets the hook procedure address: pass result of AddressOf()
' ref: msdn.microsoft.com/en-us/library/windows/desktop/bb762598(v=vs.85).aspx
Public Property Get CustomHookProc() As Long
    CustomHookProc = obif.lpfnCallback
End Property
Public Property Let CustomHookProc(ByVal nProc As Long)
    ' If you want to use your own Callback/subclass routine for this browser, then
    '   this value must be the function address/pointer of the callback procedure.
    ' If set, the InitialDirectory & InitialDirectoryPIDL property settings are ignored
    ' The callback procedure definition can be found here:
    '   http://msdn.microsoft.com/en-us/library/windows/desktop/bb762598%28v=vs.85%29.aspx
    obif.lpfnCallback = nProc
    
    ' how to get the AddressOf() into a Long value?
    ' use a helper function like so:
    
'       ' following would be placed in module along with your BrowseCallbackProc function
'       ' the BrowseCallbackProc function must also be Public for this usage
'        Public Function GetFunctionAddress(inAddr As Long) As Long
'            GetFunctionAddress = inAddr
'        End Function
'
'       ' now a sample call might look like:
'       myFolderBrowser.CustomHookProc = GetFunctionAddress(AddressOf YourBrowseCallbackProc)

'   Note that you will most likely want to or need to set the CustomHookData property also

End Property

' returns/sets the dialog window caption/title
Public Property Get dialogTitle() As String
    dialogTitle = obif.lpszTitle
End Property
Public Property Let dialogTitle(ByVal dTitle As String)
    obif.lpszTitle = dTitle
End Property

' returns the folder/filename the dialog returns
' This may return blank if user selected a virutual object: Control Panel for example
' If this is blank, then the FolderDisplayName will not be blank if browser wasn't canceled
Public Property Get SelectedFolder() As String
    SelectedFolder = m_File
End Property

' set the startup directory for the browser. If a virtual folder is requested,
'   then call InitialDirectoryPIDL instead.
' This property is ignored if you set the CustomHookProc
' This property overwrites the InitialDirectoryPIDL value if it was set
Public Property Let InitialDirectory(newValue As String)
    m_InitPath = newValue
End Property
Public Property Get InitialDirectory() As String
    If Not m_InitPath = vbNullString Then
        If IsNumeric(m_InitPath) = False Then InitialDirectory = m_InitPath
    End If
End Property

' set the startup directory for the browser as a PIDL reference.
' You are responsible for destroying the PIDL
' This property is ignored if you set the CustomHookProc
' This property overwrites the InitialDirectory value if it was set
Public Property Let InitialDirectoryPIDL(newValue As Long)
    If newValue = 0& Then
        m_InitPath = vbNullString
    ElseIf Not newValue = m_PIDL Then       ' cannot use the previously returned PIDL unless took ownership
        m_InitPath = CStr(newValue)
    End If
End Property
Public Property Get InitialDirectoryPIDL() As Long
    InitialDirectoryPIDL = Val(m_InitPath)
End Property

' reutrns the display name of the item selected
' this only returns blank if the ShowBrowseForFolder failed
Public Property Get FolderDisplayName() As String
    Dim lLen As Long
    lLen = lstrlenW(StrPtr(obif.pszDisplayName))
    If lLen Then FolderDisplayName = Left$(obif.pszDisplayName, lLen)
End Property

' returns/sets the flags used for the dialog
Public Property Get Flags() As FolderBrowseFlagsEnum
    Flags = obif.ulFlags
End Property
Public Property Let Flags(ByVal dFlags As FolderBrowseFlagsEnum)
    obif.ulFlags = dFlags
End Property

Public Property Get SelectedPIDL(Optional ByVal ReleaseOwnerShip As Boolean) As Long
' Returns the PIDL selected by the user after ShowBrowseForFolder is called
' PIDL is only valid until next ShowBrowseForFolder call or Clear() is called
' YOU ARE RESPONSIBLE FOR DESTROYING THE PIDL if you chose to
'   not release it in ShowBrowseForFolder or pass ReleaseOwnerShip here as True;
'   otherwise you must not destroy the PIDL
    SelectedPIDL = m_PIDL
    If ReleaseOwnerShip = True Then m_PIDL = 0&
End Property

' Note that this property seems to always return zero when the
' BIF_NEWDIALOGSTYLE is used
Public Property Get SysImageListIconIndex() As Long
    SysImageListIconIndex = obif.iImage
End Property

Public Sub DestroyPIDL(pIDL As Long)
    ' courtesy method provided. Allows you to destroy any PIDL without declaring APIs yourself
    If pIDL Then CoTaskMemFree pIDL
    If pIDL = m_PIDL Then m_PIDL = 0&
End Sub

Public Function CreatePIDLfromFolder(thePath As String) As Long
    ' can return a PIDL that could be useful for a BrowseCallbackProc function
    ' YOU ARE RESPONSIBLE FOR DESTROYING THE PIDL
    Call SHILCreateFromPath(StrPtr(thePath), CreatePIDLfromFolder, ByVal 0&)
End Function

' sets the top level folder where browsing will begin from
Public Function SetRootFolder(thePath As String) As Boolean
    If obif.pIDLRoot Then       ' destroy PIDL if we created it
        If (m_State And 8&) = 8& Then Call CoTaskMemFree(obif.pIDLRoot)
        obif.pIDLRoot = 0&
        m_State = m_State And &HFFFFFFF7    ' remove PIDL ownership flag
    End If
    If Len(thePath) Then
        Call SHILCreateFromPath(StrPtr(thePath), obif.pIDLRoot, ByVal 0&)
        If Not (obif.pIDLRoot = 0&) Then
            m_State = m_State Or 8&         ' add PIDL ownership flag
            SetRootFolder = True
        End If
    Else
        SetRootFolder = True
    End If
End Function

Public Function SetRootFolderPIDL(ByVal thePIDL As Long) As Boolean
    If obif.pIDLRoot Then       ' destroy PIDL if we created it
        If (m_State And 8&) = 8& Then Call CoTaskMemFree(obif.pIDLRoot)
        obif.pIDLRoot = 0&
        m_State = m_State And &HFFFFFFF7    ' remove PIDL ownership flag
    End If
    If Not thePIDL = m_PIDL Then            ' cannot use previously selected PIDL unless you took ownership
        obif.pIDLRoot = thePIDL
        SetRootFolderPIDL = True
    End If
End Function

Private Sub pvInitStructure(hWnd As Long)

    Const CB_PAGE_RWX    As Long = &H40
    Const CB_MEM_COMMIT  As Long = &H1000
    Const FUNK_LEN       As Long = 48&
    Dim lLastMethod As Long, cb() As Long
    
    If Not m_PIDL = 0& Then                     ' destroy PIDL from previous call, as needed
        If (m_State And 4&) = 0& Then Me.DestroyPIDL m_PIDL
    End If
    m_State = m_State And &HFFFFFFFD            ' remove PIDL release option
    With obif
         ' add/remove any flags that are in conflict with the dialog mode (save/open)
         If (.ulFlags And BIF_BROWSEINCLUDEURLS) Then
             ' The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set
             .ulFlags = .ulFlags Or BIF_USENEWUI Or BIF_BROWSEINCLUDEFILES
         End If
         If (.ulFlags And BIF_SHAREABLE) Then
             ' The BIF_NEWDIALOGSTYLE flag must also be set.
             .ulFlags = .ulFlags Or BIF_NEWDIALOGSTYLE
         End If
         
         .pszDisplayName = String$(MAX_PATH + 2&, vbNullChar)
         .hWndOwner = hWnd
         .iImage = 0&
     End With
     m_File = vbNullString
     If Me.CustomHookProc = 0& Then                     ' no custom hook procedure set by user
        If Not m_InitPath = vbNullString Then           ' do we have a startup pidl/path
            
            'Thunks are not required in TwinBasic; use built-in language features instead!
            ' (Thank you to @fafalone for this fix: https://github.com/tannerhelland/PhotoDemon/issues/665)
            #If TWINBASIC Then
                obif.lpfnCallback = AddressOf pvBrowseCallbackProc
            #Else
                If m_ThunkAddr = 0& Then
                    lLastMethod = pvFindLastMethod()
                    m_ThunkAddr = VirtualAlloc(0&, FUNK_LEN, CB_MEM_COMMIT, CB_PAGE_RWX)
                    ReDim cb(0 To FUNK_LEN \ 4& - 1&)           ' Allocate executable memory
                    cb(0) = ObjPtr(Me)                          ' thunks call back here
                    cb(2) = &HBB60E089: cb(4) = &HE883C589: cb(5) = &HB9509004: cb(7) = &H74FF06E3
                    cb(8) = &HFAE2008D: cb(9) = &H53FF33FF: cb(10) = &HC2906104
                    
                    CopyMemory cb(1), ByVal lLastMethod, 4&
                    cb(3) = m_ThunkAddr
                    cb(6) = 4&
                    cb(11) = cb(6) * 4&
                    CopyMemory ByVal m_ThunkAddr, cb(0), FUNK_LEN
                End If
                obif.lpfnCallback = m_ThunkAddr + 8&
            #End If
            
            If IsNumeric(m_InitPath) Then
                obif.lpCallbackData = Val(m_InitPath)
            Else
                obif.lpCallbackData = StrPtr(m_InitPath)
            End If
        End If
    End If
End Sub

Private Function pvFindLastMethod() As Long
    '/// This method must not be modified. Method is only called once
    ' Return the address of the specified ordinal private method, 1 = last private method, 2 = second last private method, etc
  Dim bSub  As Byte                         ' Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                         ' Address of the vTable
  Dim i     As Long                         ' Loop index
  Dim j     As Long                         ' Loop limit
  Dim tblLoc As Long
    CopyMemory nAddr, ByVal ObjPtr(Me), 4&  ' Get the address of this object instance
    If pvProbeClassMethods(nAddr + &H1C, i, bSub) = 0 Then Exit Function ' Probe for a Class method
    i = i + 4&                              ' Bump to the next entry
    j = i + 1024&                           ' Set a reasonable limit, scan 256 vTable entries
    Do While i < j
      CopyMemory nAddr, ByVal i, 4&         ' Get the address stored in this vTable entry
      If IsBadCodePtr(nAddr) Then           ' Is the entry an invalid code address?
        tblLoc = i                          ' Cache the vTable end-point
        GoTo Found                          ' Bad method signature, quit loop
      End If
      CopyMemory bVal, ByVal nAddr, 1&      ' Get the byte pointed to by the vTable entry
      If bVal <> bSub Then                  ' If the byte doesn't match the expected value...
        tblLoc = i                          ' Cache the vTable end-point
        GoTo Found                          ' Bad method signature, quit loop
      End If
      i = i + 4&                            ' Next vTable entry
    Loop
    Exit Function                           ' Final method not found
Found:                                      ' Return the specified vTable entry address
  pvFindLastMethod = tblLoc - 4&
End Function

Private Function pvProbeClassMethods(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  '/// This method must not be modified.  Method is only called once
  'Probe at the specified start address for a method signature
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  nAddr = nStart                            ' Start address
  nLimit = nAddr + 32&                      ' Probe eight entries
  Do While nAddr < nLimit                   ' While we've not reached our probe depth
    CopyMemory nEntry, ByVal nAddr, 4&      ' Get the vTable entry
    If nEntry <> 0 Then                     ' If not an implemented interface
      CopyMemory bVal, ByVal nEntry, 1&     ' Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then    ' Check for a native or pcode method signature
        nMethod = nAddr                     ' Store the vTable entry
        bSub = bVal                         ' Store the found method signature
        pvProbeClassMethods = True          ' Indicate success
        Exit Function                       ' Return
      End If
    End If
    nAddr = nAddr + 4&                      ' Next vTable entry
  Loop
End Function

Private Sub Class_Initialize()
    If Not (IsWindowUnicode(GetDesktopWindow) = 0&) Then m_State = 1&
End Sub

Private Sub Class_Terminate()
    Call Me.Clear
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////////
'   This pvBrowseCallbackProc must be the final method of this class
'//////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function pvBrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal lpData As Long) As Long
    Select Case uMsg
       Case BFFM_INITIALIZED
            ' http://msdn.microsoft.com/en-us/library/aa452875.aspx
            ' the 3rd parameter of SendMessage is either zero or non-zero
            '   zero :: lParam is a PIDL (as used below)
            '   non-zero :: lParam is yourDesiredInitialDirectoryPath string
            If IsNumeric(m_InitPath) Then lp = 0& Else lp = 1&
            If (m_State And 1&) = 1& Then
                Call SendMessageA(hWnd, BFFM_SETSELECTIONW, lp, ByVal lpData)
            Else
                Call SendMessageA(hWnd, BFFM_SETSELECTIONA, lp, ByVal lpData)
            End If

        ' other Case BFFM_xxxx messages as desired

   End Select
   pvBrowseCallbackProc = 0&
End Function
'//////////////////////////////////////////////////////////////////////////////////////////////////////
'   DO NOT PLACE ANY EXECUTABLE CODE AFTER THE p0vBrowseCallbackProc
'//////////////////////////////////////////////////////////////////////////////////////////////////////
