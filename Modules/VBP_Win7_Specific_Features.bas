Attribute VB_Name = "OS_Win7_8_Features"
'***************************************************************************
'Handler for features specific to Windows 7+
'Copyright ©2012-2013 by Tanner Helland
'Created: 21/November/13
'Last updated: 21/November/13
'Last update: initial build
'
'Windows 7 exposes some neat features (like progress bars overlaying the taskbar), and PhotoDemon tries to make
' use of them when relevant.  All Win7-specific features are handled from this module.  If Win7 is not present,
' calling these functions has no effect.
'
'I owe many thanks to AndRAY and his VB project located here, which inspired many of these features:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=72856&lngWId=1
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'***************************************************************************
'Functions required to initialize an OLE interface from within VB

Private Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Declare Sub OleInitialize Lib "ole32" (pvReserved As Any)
Private Declare Sub OleUninitialize Lib "ole32" ()
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As String, pclsid As Guid) As Long
Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As String, lpiid As Guid) As Long
Private Declare Function CoCreateInstance Lib "ole32" (rclsid As Guid, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Guid, ppv As Any) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PutMem2 Lib "msvbvm60" (ByVal pWORDDst As Long, ByVal NewValue As Long) As Long
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal pDWORDDst As Long, ByVal NewValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByVal pDWORDSrc As Long, ByVal pDWORDDst As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Const GMEM_FIXED As Long = &H0
Private Const asmPUSH_imm32 As Byte = &H68
Private Const asmRET_imm16 As Byte = &HC2
Private Const asmCALL_rel32 As Byte = &HE8

Private Const unk_QueryInterface As Long = 0
Private Const unk_AddRef As Long = 1
Private Const unk_Release As Long = 2

Private Const CLSID_TaskbarList As String = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
Private Const IID_ITaskbarList3 As String = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"
Private Enum ITaskbarList3Members
                                '/* ITaskbarList methods */
    HrInit_ = 3                 'STDMETHOD( HrInit )( THIS ) PURE;
    AddTab_ = 4                 'STDMETHOD( AddTab )( THIS_ HWND ) PURE;
    DeleteTab_ = 5              'STDMETHOD( DeleteTab )( THIS_ HWND ) PURE;
    ActivateTab_ = 6            'STDMETHOD( ActivateTab )( THIS_ HWND ) PURE;
    SetActiveAlt_ = 7           'STDMETHOD( SetActiveAlt )( THIS_ HWND ) PURE;
                                '/* ITaskbarList2 methods */
    MarkFullscreenWindow_ = 8   'STDMETHOD( MarkFullscreenWindow )( THIS_ HWND, BOOL ) PURE;
                                '/* ITaskbarList3 methods */
    SetProgressValue_ = 9       'STDMETHOD( SetProgressValue )( THIS_ HWND, ULONGLONG, ULONGLONG ) PURE;
    SetProgressState_ = 10      'STDMETHOD( SetProgressState )( THIS_ HWND, TBPFLAG ) PURE;
    RegisterTab_ = 11           'STDMETHOD( RegisterTab )( THIS_ HWND, HWND ) PURE;
    UnregisterTab_ = 12         'STDMETHOD( UnregisterTab )( THIS_ HWND ) PURE;
    SetTabOrder_ = 13           'STDMETHOD( SetTabOrder )( THIS_ HWND, HWND ) PURE;
    SetTabActive_ = 14          'STDMETHOD( SetTabActive )( THIS_ HWND, HWND, DWORD ) PURE;
    ThumbBarAddButtons_ = 15    'STDMETHOD( ThumbBarAddButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
    ThumbBarUpdateButtons_ = 16 'STDMETHOD( ThumbBarUpdateButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
    ThumbBarSetImageList_ = 17  'STDMETHOD( ThumbBarSetImageList )( THIS_ HWND, HIMAGELIST ) PURE;
    SetOverlayIcon_ = 18        'STDMETHOD( SetOverlayIcon )( THIS_ HWND, HICON, LPCWSTR ) PURE;
    SetThumbnailTooltip_ = 19   'STDMETHOD( SetThumbnailTooltip )( THIS_ HWND, LPCWSTR ) PURE;
    SetThumbnailClip_ = 20      'STDMETHOD( SetThumbnailClip )( THIS_ HWND, RECT * ) PURE;
'                                '/* ITaskbarList4 methods */
'    SetTabProperties_ = 21      'STDMETHOD( SetTabProperties )( THIS_ HWND, STPFLAG ) PURE;
End Enum

'
'***************************************************************************

'Possible task bar progress states.  PD is primarily interested in NOPROGRESS and NORMAL
Public Const TBPF_NOPROGRESS = 0
Public Const TBPF_INDETERMINATE = 1
Public Const TBPF_NORMAL = 2
Public Const TBPF_ERROR = 4
Public Const TBPF_PAUSED = 8

'The handle to the OLE interface we create
Private objHandle As Long

'If this module is enabled, this will be set to TRUE
Private win7FeaturesAllowed As Boolean


'Request an OLE interface from within VB.  I apologize for a lack of comments in this function, but I did not write it.
' For additional details, please see the original project, available here: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=72856&lngWId=1
Private Function CallInterface(ByVal pInterface As Long, ByVal Member As Long, ByVal ParamsCount As Long, Optional ByVal p1 As Long = 0, Optional ByVal p2 As Long = 0, Optional ByVal p3 As Long = 0, Optional ByVal p4 As Long = 0, Optional ByVal p5 As Long = 0, Optional ByVal p6 As Long = 0, Optional ByVal p7 As Long = 0, Optional ByVal p8 As Long = 0, Optional ByVal p9 As Long = 0, Optional ByVal p10 As Long = 0) As Long
  
    Dim i As Long, t As Long
    Dim hGlobal As Long, hGlobalOffset As Long
    
    If ParamsCount < 0 Then Err.Raise 5
    If pInterface = 0 Then Err.Raise 5
    
    hGlobal = GlobalAlloc(GMEM_FIXED, 5 * ParamsCount + 5 + 5 + 3 + 1)
    If hGlobal = 0 Then Err.Raise 7
    hGlobalOffset = hGlobal
    
    If ParamsCount > 0 Then
        t = VarPtr(p1)
        
        For i = ParamsCount - 1 To 0 Step -1
          PutMem2 hGlobalOffset, asmPUSH_imm32
          hGlobalOffset = hGlobalOffset + 1
          GetMem4 t + i * 4, hGlobalOffset
          hGlobalOffset = hGlobalOffset + 4
        Next
      
    End If
    
    PutMem2 hGlobalOffset, asmPUSH_imm32
    hGlobalOffset = hGlobalOffset + 1
    PutMem4 hGlobalOffset, pInterface
    hGlobalOffset = hGlobalOffset + 4
    
    PutMem2 hGlobalOffset, asmCALL_rel32
    hGlobalOffset = hGlobalOffset + 1
    GetMem4 pInterface, VarPtr(t)
    GetMem4 t + Member * 4, VarPtr(t)
    PutMem4 hGlobalOffset, t - hGlobalOffset - 4
    hGlobalOffset = hGlobalOffset + 4
      
    PutMem4 hGlobalOffset, &H10C2&
    
    CallInterface = CallWindowProc(hGlobal, 0, 0, 0, 0)
    
    GlobalFree hGlobal
  
End Function

'If desired, a custom state can be set for the taskbar.  Normally this is handled by the SetTaskbarProgressValue function,
' but it can also be done custom here.
Public Function SetTaskbarProgressState(ByVal tbpFlags As Long) As Long
    If win7FeaturesAllowed Then
        SetTaskbarProgressState = CallInterface(objHandle, SetProgressState_, 2, FormMain.hWnd, tbpFlags)
    End If
End Function

Public Function SetTaskbarProgressValue(ByVal amtCompleted As Long, ByVal amtTotal As Long) As Long
    If win7FeaturesAllowed Then
        If amtCompleted = 0 Then
            SetTaskbarProgressState TBPF_NOPROGRESS
        Else
            SetTaskbarProgressState TBPF_NORMAL
            SetTaskbarProgressValue = CallInterface(objHandle, SetProgressValue_, 5, FormMain.hWnd, amtCompleted, 0, amtTotal, 0)
        End If
    End If
End Function

'If the OS is detected as Windows 7+, this function will be called.  It will prepare a handle to the OLE interface
' we use for Win7-specific features.
Public Sub prepWin7Features()

    'Note that Win 7 features are allowed
    win7FeaturesAllowed = True
    
    Dim CLSID As Guid, InterfaceGuid As Guid
    Call CLSIDFromString(StrConv(CLSID_TaskbarList, vbUnicode), CLSID)
    Call IIDFromString(StrConv(IID_ITaskbarList3, vbUnicode), InterfaceGuid)
    Call CoCreateInstance(CLSID, 0, 1, InterfaceGuid, objHandle)
    
End Sub

'Make sure to release the interface when we are done with it!
Public Sub releaseWin7Features()
    If win7FeaturesAllowed Then CallInterface objHandle, unk_Release, 0
End Sub
