VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSelfSubHookCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'To my knowledge, this class is Copyright 2007 Paul Caton and LaVolpe

'You may download the original version of this code from the following link (good as of February '13):
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=68737&lngWId=1

'Many thanks to these fine gentlemen for their work on IDE-safe subclassing in VB6.


Option Explicit

'========================================
' COMMON-USE ITEMS USED IN ALL THUNKS
'========================================
  ' Local variables/constants: must declare these regardless if using subclassing, hooking, callbacks
    Private z_scFunk            As Collection   'hWnd/thunk-address collection
    Private z_hkFunk            As Collection   'hook/thunk-address collection
    Private z_cbFunk            As Collection   'callback/thunk-address collection
    Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
    Private Const IDX_PREVPROC  As Long = 9     'Thunk data index of the original WndProc
    Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table for messages
    Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table for messages
    Private Const IDX_CALLBACKORDINAL As Long = 36 ' Ubound(callback thunkdata)+1, index of the callback

  ' Declarations:
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Public Enum eThunkType
        SubclassThunk = 0
        HookThunk = 1
        CallbackThunk = 2
    End Enum

'===============================
' TO USE IDE-SAFE SUBCLASSING...
'===============================
  ' Special Notes:
  '   a. ssc_Subclass uses function ordinals. Last routine in the class/form/uc/ppg is Ordinal #1, second to last is Ordinal #2, etc
  '         IMPORTANT: Routines used for subclass procedure MUST NOT be Public, can be Private or Friend
  '   b. VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
  '   c. Do not terminate application while stopped in the subclass procedure
  '   d. At a minimum, review ssc_Subclass, ssc_Unsubclass & ssc_Terminate
    
    Public Enum eMsgWhen                                                   'When to callback
      MSG_BEFORE = 1                                                        'Callback before the original WndProc
      MSG_AFTER = 2                                                         'Callback after the original WndProc
      MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
    End Enum
    
    Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
    Private Const IDX_UNICODE   As Long = 107   'Must be UBound(subclass thunkdata)+1; index for unicode support
    Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows
    
    Public Enum eAllMessages
        ALL_MESSAGES = -1     'All messages will callback
    End Enum
    
    Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'-------------------------------------------------------------------------------------------------

'========================================================================================================
' TO USE IDE-SAFE HOOKING...
'==============================
  ' Special Notes:
  '   a. shk_SetHook uses function ordinals. Last routine in the class/form/uc is Ordinal #1, second to last is Ordinal #2, etc
  '         IMPORTANT: Routines used for hooks MUST NOT be Public, can be Private or Friend
  '   b. VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
  '   c. When END executed, app does not crash, but hook may not be released immediately. Should you notice this,
  '      simply do something that would trigger the hook (i.e., move  mouse, keystroke, depending on hook type).
  '      The hook then will release.
  '   d. Do not terminate application while stopped in the hook procedure
  '   e. At a minimum, review shk_SetHook, shk_Unhook & shk_TerminateHooks
  
    '-SelfHook specific declarations----------------------------------------------------------------------------
    Private Declare Function SetWindowsHookExA Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
    Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
        Private Declare Function GetDesktopWindow Lib "user32" () As Long
    
    Public Enum eHookType  ' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
      WH_MSGFILTER = -1
      WH_JOURNALRECORD = 0
      WH_JOURNALPLAYBACK = 1
      WH_KEYBOARD = 2
      WH_GETMESSAGE = 3
      WH_CALLWNDPROC = 4
      WH_CBT = 5
      WH_SYSMSGFILTER = 6
      WH_MOUSE = 7
      WH_DEBUG = 9
      WH_SHELL = 10
      WH_FOREGROUNDIDLE = 11
      WH_CALLWNDPROCRET = 12
      WH_KEYBOARD_LL = 13       ' NT/2000/XP+ only, Global hook only
      WH_MOUSE_LL = 14          ' NT/2000/XP+ only, Global hook only
    End Enum
    
    
'-------------------------------------------------------------------------------------------------
'========================================================================================================
' TO USE IDE-SAFE CALLBACKS...
'==============================
  ' Special notes:
  '  a. scb_SetCallbackAddr uses function ordinals. Last routine in the class/form/uc/ppg is Ordinal #1, second to last is Ordinal #2, etc
  '         IMPORTANT: Routines used for callbacks MUST NOT be Public, can be Private or Friend
  '  b. The callback function in your class/form/uc/ppg must be a Function that returns Long.
  '     For example. The callback routine for a timer has no return value per MSDN, but when that timer
  '     procedure is coded, the callback function must return a long. The return value is not important in that case
  '  c. VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
  '  d. Do not terminate application while stopped in the callback procedure
  '  e. Never call scb_ReleaseCallback from within the callback procedure
  '  f. At a minimum, review scb_SetCallbackAddr, scb_ReleaseCallback & scb_TerminateCallbacks
'-------------------------------------------------------------------------------------------------

Private Sub Class_Terminate()

  ssc_Terminate      '(add this to Unload or Terminate event if you are subclassing)
  shk_TerminateHooks  '(add this to Unload or Terminate event if you are hooking)
  scb_TerminateCallbacks '(add this to Unload or Terminate event if you are using callbacks)
  
End Sub

'-SelfSub code------------------------------------------------------------------------------------
'-The following routines are exclusively for the ssc_Subclass routines----------------------------
Public Function ssc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False, _
                    Optional ByVal bIsAPIwindow As Boolean = False) As Boolean 'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. There is not reason to set this to False
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '* bIsAPIwindow - Optional, if True DestroyWindow will be called if IDE ENDs
    '*****************************************************************************************
    '** Subclass.asm - subclassing thunk
    '**
    '** Paul_Caton@hotmail.com
    '** Copyright free, use and abuse as you see fit.
    '**
    '** v2.0 Re-write by LaVolpe, based mostly on Paul Caton's original thunks....... 20070720
    '** .... Reorganized & provided following additional logic
    '** ....... Unsubclassing only occurs after thunk is no longer recursed
    '** ....... Flag used to bypass callbacks until unsubclassing can occur
    '** ....... Timer used as delay mechanism to free thunk memory afer unsubclassing occurs
    '** .............. Prevents crash when one window subclassed multiple times
    '** .............. More END safe, even if END occurs within the subclass procedure
    '** ....... Added ability to destroy API windows when IDE terminates
    '** ....... Added auto-unsubclass when WM_NCDESTROY received
    '*****************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    
    Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name
    Const CODE_LEN      As Long = 4 * IDX_UNICODE + 4  'Thunk length in bytes
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H60                 'Thunk offset to the WndProc execution address
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)) 'Bytes to allocate per thunk, data + code + msg tables
    
  ' This is the complete listing of thunk offset values and what they point/relate to.
  ' Those rem'd out are used elsewhere or are initialized in Declarations section
  
  'Const IDX_RECURSION  As Long = 0     'Thunk data index of callback recursion count
  'Const IDX_SHUTDOWN   As Long = 1     'Thunk data index of the termination flag
  'Const IDX_INDEX      As Long = 2     'Thunk data index of the subclassed hWnd
   Const IDX_EBMODE     As Long = 3     'Thunk data index of the EbMode function address
   Const IDX_CWP        As Long = 4     'Thunk data index of the CallWindowProc function address
   Const IDX_SWL        As Long = 5     'Thunk data index of the SetWindowsLong function address
   Const IDX_FREE       As Long = 6     'Thunk data index of the VirtualFree function address
   Const IDX_BADPTR     As Long = 7     'Thunk data index of the IsBadCodePtr function address
   Const IDX_OWNER      As Long = 8     'Thunk data index of the Owner object's vTable address
  'Const IDX_PREVPROC   As Long = 9     'Thunk data index of the original WndProc
   Const IDX_CALLBACK   As Long = 10    'Thunk data index of the callback method address
  'Const IDX_BTABLE     As Long = 11    'Thunk data index of the Before table
  'Const IDX_ATABLE     As Long = 12    'Thunk data index of the After table
  'Const IDX_PARM_USER  As Long = 13    'Thunk data index of the User-defined callback parameter data index
   Const IDX_DW         As Long = 14    'Thunk data index of the DestroyWinodw function address
   Const IDX_ST         As Long = 15    'Thunk data index of the SetTimer function address
   Const IDX_KT         As Long = 16    'Thunk data index of the KillTimer function address
   Const IDX_EBX_TMR    As Long = 20    'Thunk code patch index of the thunk data for the delay timer
   Const IDX_EBX        As Long = 26    'Thunk code patch index of the thunk data
  'Const IDX_UNICODE    As Long = xx    'Must be UBound(subclass thunkdata)+1; index for unicode support
    
    Dim z_ScMem       As Long           'Thunk base address
    Dim nAddr         As Long
    Dim nID           As Long
    Dim nMyID         As Long
    Dim bIDE          As Boolean

    If IsWindow(lng_hWnd) = 0 Then      'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID              'Get the process ID associated with the window handle
    If nID <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If z_ScMem <> 0 Then                                  'Ensure the allocation succeeded
    
      If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
      On Error GoTo CatchDoubleSub                              'Catch double subclassing
        z_scFunk.Add z_ScMem, "h" & lng_hWnd                    'Add the hWnd/thunk-address to the collection
      On Error GoTo 0
      
   'z_Sc (0) thru z_Sc(17) are used as storage for the thunks & IDX_ constants above relate to these thunk positions which are filled in below
    z_Sc(18) = &HD231C031: z_Sc(19) = &HBBE58960: z_Sc(21) = &H21E8F631: z_Sc(22) = &HE9000001: z_Sc(23) = &H12C&: z_Sc(24) = &HD231C031: z_Sc(25) = &HBBE58960: z_Sc(27) = &H3FFF631: z_Sc(28) = &H75047339: z_Sc(29) = &H2873FF23: z_Sc(30) = &H751C53FF: z_Sc(31) = &HC433913: z_Sc(32) = &H53FF2274: z_Sc(33) = &H13D0C: z_Sc(34) = &H18740000: z_Sc(35) = &H875C085: z_Sc(36) = &H820443C7: z_Sc(37) = &H90000000: z_Sc(38) = &H87E8&: z_Sc(39) = &H22E900: z_Sc(40) = &H90900000: z_Sc(41) = &H2C7B8B4A: z_Sc(42) = &HE81C7589: z_Sc(43) = &H90&: z_Sc(44) = &H75147539: z_Sc(45) = &H6AE80F: z_Sc(46) = &HD2310000: z_Sc(47) = &HE8307B8B: z_Sc(48) = &H7C&: z_Sc(49) = &H7D810BFF: z_Sc(50) = &H8228&: z_Sc(51) = &HC7097500: z_Sc(52) = &H80000443: z_Sc(53) = &H90900000: z_Sc(54) = &H44753339: z_Sc(55) = &H74047339: z_Sc(56) = &H2473FF3F: z_Sc(57) = &HFFFFFC68
    z_Sc(58) = &H2475FFFF: z_Sc(59) = &H811453FF: z_Sc(60) = &H82047B: z_Sc(61) = &HC750000: z_Sc(62) = &H74387339: z_Sc(63) = &H2475FF07: z_Sc(64) = &H903853FF: z_Sc(65) = &H81445B89: z_Sc(66) = &H484443: z_Sc(67) = &H73FF0000: z_Sc(68) = &H646844: z_Sc(69) = &H56560000: z_Sc(70) = &H893C53FF: z_Sc(71) = &H90904443: z_Sc(72) = &H10C261: z_Sc(73) = &H53E8&: z_Sc(74) = &H3075FF00: z_Sc(75) = &HFF2C75FF: z_Sc(76) = &H75FF2875: z_Sc(77) = &H2473FF24: z_Sc(78) = &H891053FF: z_Sc(79) = &H90C31C45: z_Sc(80) = &H34E30F8B: z_Sc(81) = &H1078C985: z_Sc(82) = &H4C781: z_Sc(83) = &H458B0000: z_Sc(84) = &H75AFF228: z_Sc(85) = &H90909023: z_Sc(86) = &H8D144D8D: z_Sc(87) = &H8D503443: z_Sc(88) = &H75FF1C45: z_Sc(89) = &H2C75FF30: z_Sc(90) = &HFF2875FF: z_Sc(91) = &H51502475: z_Sc(92) = &H2073FF52: z_Sc(93) = &H902853FF: z_Sc(94) = &H909090C3: z_Sc(95) = &H74447339: z_Sc(96) = &H4473FFF7
    z_Sc(97) = &H4053FF56: z_Sc(98) = &HC3447389: z_Sc(99) = &H89285D89: z_Sc(100) = &H45C72C75: z_Sc(101) = &H800030: z_Sc(102) = &H20458B00: z_Sc(103) = &H89145D89: z_Sc(104) = &H81612445: z_Sc(105) = &H4C4&: z_Sc(106) = &H1862FF00

    ' cache callback related pointers & offsets
      z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
      z_Sc(IDX_EBX_TMR) = z_ScMem                                             'Patch the thunk data address
      z_Sc(IDX_INDEX) = lng_hWnd                                              'Store the window handle in the thunk data
      z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
      z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
      z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
      z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
      z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
      
      ' validate unicode request & cache unicode usage
      If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
      z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
      
      ' get function pointers for the thunk
      If bIdeSafety = True Then                                               'If the user wants IDE protection
          Debug.Assert zInIDE(bIDE)
          If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
                                                        '^^ vb5 users, change vba6 to vba5
      End If
      If bIsAPIwindow Then                                                    'If user wants DestroyWindow sent should IDE end
          z_Sc(IDX_DW) = zFnAddr("user32", "DestroyWindow", bUnicode)
      End If
      z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
      z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
      z_Sc(IDX_ST) = zFnAddr("user32", "SetTimer", bUnicode)                  'Store the SetTimer function address in the thunk data
      z_Sc(IDX_KT) = zFnAddr("user32", "KillTimer", bUnicode)                 'Store the KillTimer function address in the thunk data
      
      If bUnicode Then
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      Else
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      End If
      If z_Sc(IDX_PREVPROC) = 0 Then                                          'Ensure the new WndProc was set correctly
          zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
      End If
      'Store the original WndProc address in the thunk data
      RtlMoveMemory z_ScMem + IDX_PREVPROC * 4, VarPtr(z_Sc(IDX_PREVPROC)), 4&
      ssc_Subclass = True                                                     'Indicate success
      
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
        
    End If

 Exit Function                                                                'Exit ssc_Subclass
    
CatchDoubleSub:
 zError SUB_NAME, "Window handle is already subclassed"
      
ReleaseMemory:
      VirtualFree z_ScMem, 0, MEM_RELEASE                                     'ssc_Subclass has failed after memory allocation, so release the memory
      
End Function

'Terminate all subclassing
Public Sub ssc_Terminate()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    zTerminateThunks SubclassThunk
End Sub

'UnSubclass the specified window handle
Public Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public, can be removed & zUnthunk can be called instead
    zUnThunk lng_hWnd, SubclassThunk
End Sub

Public Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' can be made public, can be removed if you will not use this in your window procedure
    Dim z_ScMem       As Long         'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)
    If z_ScMem Then                 'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE, z_ScMem) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
End Function


'Add the message value to the window handle's specified callback table
Public Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    
    Dim z_ScMem       As Long         'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)
    If z_ScMem Then                     'Ensure that the thunk hasn't already released its memory
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)
        Select Case VarType(Messages(M))                                        ' ensure no strings, arrays, doubles, objects, etc are passed
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                                         'If the message is to be added to the before original WndProc table...
              If zAddMsg(Messages(M), IDX_BTABLE, z_ScMem) = False Then                 'Add the message to the before table
                When = (When And Not MSG_BEFORE)
              End If
            End If
            If When And MSG_AFTER Then                                          'If message is to be added to the after original WndProc table...
              If zAddMsg(Messages(M), IDX_ATABLE, z_ScMem) = False Then                 'Add the message to the after table
                When = (When And Not MSG_AFTER)
              End If
            End If
        End Select
      Next
    End If
End Sub

'Delete the message value from the window handle's specified callback table
Public Sub ssc_DelMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    
    Dim z_ScMem       As Long         'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)
    If z_ScMem Then                'Ensure that the thunk hasn't already released its memory
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)                                  ' ensure no strings, arrays, doubles, objects, etc are passed
        Select Case VarType(Messages(M))
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                                             'If the message is to be removed from the before original WndProc table...
              zDelMsg Messages(M), IDX_BTABLE, z_ScMem                                      'Remove the message to the before table
            End If
            If When And MSG_AFTER Then                                              'If message is to be removed from the after original WndProc table...
              zDelMsg Messages(M), IDX_ATABLE, z_ScMem                                      'Remove the message to the after table
            End If
        End Select
      Next
    End If
End Sub

'Let the subclasser lParamUser callback parameter
Public Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType, ByVal NewValue As Long)
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long         'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                 'Ensure that the thunk hasn't already released its memory
          zData(IDX_PARM_USER, z_ScMem) = NewValue                                         'Set the lParamUser callback parameter
        End If
    End If
End Sub

'Add the message to the specified table of the window handle
Private Function zAddMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long) As Boolean
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long                                                        'Remember z_ScMem
      Dim i      As Long                                                        'Loop index
    
      zAddMsg = True
      nBase = zData(nTable, z_ScMem)                                                 'Map zData() to the specified table
      
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
      Else
        
        nCount = zData(0, nBase)                                                      'Get the current table entry count
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = 0 Then                                                 'If the element is free...
            zData(i, nBase) = uMsg                                                    'Use this element
            GoTo Bail                                                           'Bail
          ElseIf zData(i, nBase) = uMsg Then                                          'If the message is already in the table...
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
    
        nCount = i                                                             'On drop through: i = nCount + 1, the new table entry count
        If nCount > MSG_ENTRIES Then                                           'Check for message table overflow
          zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
          zAddMsg = False
          GoTo Bail
        End If
        
        zData(nCount, nBase) = uMsg                                                   'Store the message in the appended table entry
      End If
    
      zData(0, nBase) = nCount                                                        'Store the new table entry count
Bail:

End Function

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long)
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long                                                        'Remember z_ScMem
      Dim i      As Long                                                        'Loop index
    
      nBase = zData(nTable, z_ScMem)                                                 'Map zData() to the specified table
    
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0, nBase) = 0                                                           'Zero the table entry count
      Else
        nCount = zData(0, nBase)                                                      'Get the table entry count
        
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = uMsg Then                                              'If the message is found...
            zData(i, nBase) = 0                                                       'Null the msg value -- also frees the element for re-use
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
        
'        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
      End If
      
Bail:

End Sub

'-SelfHook code------------------------------------------------------------------------------------
'-The following routines are exclusively for the shk_SetHook routines----------------------------
Public Function shk_SetHook(ByVal HookType As eHookType, _
                    Optional ByVal bGlobal As Boolean, _
                    Optional ByVal When As eMsgWhen = MSG_BEFORE, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False) As Boolean 'Setting specified hook

    '*************************************************************************************************
    '* HookType - One of the eHookType enumerators
    '* bGlobal - If False, then hook applies to app's thread else it applies Globally (only supported by WH_KEYBOARD_LL & WH_MOUSE_LL)
    '* When - either MSG_AFTER, MSG_BEFORE or MSG_BEFORE_AFTER
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time hooking
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '*************************************************************************************************
    ' Hook procedure must be declared identical to the one near the end of this class (Sample at Ordinal #2)

    ' \\LaVolpe - The ASM for this procedure rewritten to mirror Paul Caton's SelfSub ASM
    '       Therefore, it appears to be crash proof and allows a choice of whether you want
    '       hook messages before and/or after the VB gets the message
    
    Dim z_Sc(0 To 66) As Long                   'Thunk machine-code initialised here
    Const MEM_LEN      As Long = 4 * 67         'Thunk length in bytes (last # must be = UBound zSc() + 1)
    
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const IDX_EBMODE    As Long = 3             'Thunk data index of the EbMode function address
    Const IDX_CNH       As Long = 4             'Thunk data index of the CallNextHook function address
    Const IDX_UNW       As Long = 5             'Thunk data index of the UnhookWindowsEx function address
    Const IDX_OBJCHK    As Long = 6             'Thunk data index of the callback validation token
    Const IDX_BADPTR    As Long = 7             'Thunk data index of the IsBadCodePtr function address
    Const IDX_OWNER     As Long = 8             'Thunk data index of the Owner object's vTable address
    Const IDX_CALLBACK  As Long = 10            'Thunk data index of the callback method address
    Const IDX_BTABLE    As Long = 11            'Thunk data index of the Before flag
    Const IDX_ATABLE    As Long = 12            'Thunk data index of the After flag
    Const IDX_EBX       As Long = 16            'Thunk code patch index of the thunk data
    Const PROC_OFF      As Long = &H38          'Thunk offset to the HookProc execution address
    Const SUB_NAME      As String = "shk_SetHook" 'This routine's name
    
    Dim nAddr           As Long
    Dim nID             As Long
    Dim bIDE            As Boolean
    Dim z_ScMem         As Long                 'Thunk base address

    If oCallback Is Nothing Then Set oCallback = Me 'If the user hasn't specified the callback owner
      
    nAddr = zAddressOf(oCallback, nOrdinal)         'Get the address of the specified ordinal method
    If nAddr = 0 Then                               'Ensure that we've found the ordinal method
      zError SUB_NAME, "Callback method not found"
      Exit Function
    End If
       
    If Not bGlobal Then nID = App.ThreadID                      ' thread ID to be used if not global hook
      
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)    'Allocate executable memory
    
    If z_ScMem <> 0 Then                                        'Ensure the allocation succeeded
      
        If z_hkFunk Is Nothing Then Set z_hkFunk = New Collection   'If this is the first time through, do the one-time initialization
        On Error GoTo CatchDoubleSub                                'Catch double subclassing
          z_hkFunk.Add z_ScMem, "h" & HookType                      'Add the hook/thunk-address to the collection
        On Error GoTo 0
        
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H3FFF631: z_Sc(18) = &H75047339: z_Sc(19) = &H2873FF2F: z_Sc(20) = &H751C53FF: z_Sc(21) = &HC43391F: z_Sc(22) = &H7B8B2E74: z_Sc(23) = &H8478B20: z_Sc(24) = &H75184339: z_Sc(25) = &HC53FF0F: z_Sc(26) = &H13D&: z_Sc(27) = &H85197400: z_Sc(28) = &H900975C0: z_Sc(29) = &H443C7: z_Sc(30) = &H90000080: z_Sc(31) = &H47E8&: z_Sc(32) = &H2AE900: z_Sc(33) = &H90900000: z_Sc(34) = &H742C7339: z_Sc(35) = &H75894A0F: z_Sc(36) = &H46E81C: z_Sc(37) = &H75390000: z_Sc(38) = &H90157514: z_Sc(39) = &H27E8&: z_Sc(40) = &H30733900: z_Sc(41) = &HD2310A74: z_Sc(42) = &H2FE8&: z_Sc(43) = &H90909000: z_Sc(44) = &H33390BFF: z_Sc(45) = &H73390E75: z_Sc(46) = &HFF097404: z_Sc(47) = &H53FF2473: z_Sc(48) = &H90909014: z_Sc(49) = &HCC261: z_Sc(50) = &HFF2C75FF: z_Sc(51) = &H75FF2875: z_Sc(52) = &H2473FF24
        z_Sc(53) = &H891053FF: z_Sc(54) = &H90C31C45: z_Sc(55) = &H2873FF52: z_Sc(56) = &H5A1C53FF: z_Sc(57) = &H438D2275: z_Sc(58) = &H144D8D34: z_Sc(59) = &H1C458D50: z_Sc(60) = &HFF0873FF: z_Sc(61) = &H75FF2C75: z_Sc(62) = &H2475FF28: z_Sc(63) = &HFF525150: z_Sc(64) = &H53FF2073: z_Sc(65) = &H90909028: z_Sc(66) = &HC3&

        z_Sc(IDX_EBX) = z_ScMem                         'Patch the thunk data address
        z_Sc(IDX_INDEX) = HookType                      'Store the hook type in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)             'Store the callback owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                      'Store the callback address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                'Store the lParamUser callback parameter in the thunk data
        
        ' get a piece of the oCallback to use as a validation token
        RtlMoveMemory VarPtr(z_Sc(IDX_OBJCHK)), z_Sc(IDX_OWNER) + 8, 4&

        ' validate unicode request & cache unicode usage
        If bUnicode Then bUnicode = (IsWindowUnicode(GetDesktopWindow) <> 0&)
        
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)    'Store the IsBadCodePtr function address in the thunk data
        z_Sc(IDX_CNH) = zFnAddr("user32", "CallNextHookEx", bUnicode)       'Store CallWindowProc function address in the thunk data
        z_Sc(IDX_UNW) = zFnAddr("user32", "UnhookWindowsHookEx", bUnicode)  'Store the SetWindowLong function address in the thunk data
        
        If bIdeSafety = True Then                                           'If the user wants IDE protection
            Debug.Assert zInIDE(bIDE)
            If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
        End If
        
        If (When And MSG_BEFORE) = MSG_BEFORE Then z_Sc(IDX_BTABLE) = 1     ' non-zero flag if Before messages desired
        If (When And MSG_AFTER) = MSG_AFTER Then z_Sc(IDX_ATABLE) = 1       ' non-zero flag if After messages desired
        
        RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), MEM_LEN                     'Copy the thunk code/data to the allocated memory
        'Set the new HookProc, return the address of the original HookProc
        If bUnicode Then
            z_Sc(IDX_PREVPROC) = SetWindowsHookExW(HookType, z_ScMem + PROC_OFF, App.hInstance, nID)
        Else
            z_Sc(IDX_PREVPROC) = SetWindowsHookExA(HookType, z_ScMem + PROC_OFF, App.hInstance, nID)
        End If
        
        If z_Sc(IDX_PREVPROC) = 0 Then                                              'Ensure the new HookProc was set correctly
          zError SUB_NAME, "SetWindowsHookEx failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
        End If
        RtlMoveMemory z_ScMem + IDX_PREVPROC * 4, VarPtr(z_Sc(IDX_PREVPROC)), 4&    'Store the callback address
            
        shk_SetHook = True                                                          'Indicate success
      Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
      End If
      
      Exit Function
    
CatchDoubleSub:
      zError SUB_NAME, "Hook is already established"
      
ReleaseMemory:
      VirtualFree z_ScMem, 0, MEM_RELEASE                                     'shk_SetHook has failed after memory allocation, so release the memory
    
End Function

Public Function shk_UnHook(ByVal HookType As eHookType) As Boolean
    ' can be made public, can be removed & zUnThunk can be called instead
    zUnThunk HookType, HookThunk
End Function

Public Sub shk_TerminateHooks()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    zTerminateThunks HookThunk
End Sub


'-The following routines are used for each of the three types of thunks ----------------------------

'Maps zData() to the memory address for the specified thunk type
Private Function zMap_VFunction(vFuncTarget As Long, _
                                vType As eThunkType, _
                                Optional oCallback As Object, _
                                Optional bIgnoreErrors As Boolean) As Long
    
    '\\LaVolpe - Redone to be shared/used by Hook, Subclass & Callback routines
    
    Dim thunkCol As Collection
    Dim colID As String
    Dim z_ScMem       As Long         'Thunk base address
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
        If oCallback Is Nothing Then Set oCallback = Me
        colID = "h" & ObjPtr(oCallback) & "." & vFuncTarget
    ElseIf vType = HookThunk Then
        Set thunkCol = z_hkFunk
        colID = "h" & vFuncTarget
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
        colID = "h" & vFuncTarget
    Else
        zError "zMap_Vfunction", "Invalid thunk type passed"
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        zError "zMap_VFunction", "Thunk hasn't been initialized"
    Else
        If thunkCol.Count Then
            On Error GoTo Catch
            z_ScMem = thunkCol(colID)               'Get the thunk address
            If IsBadCodePtr(z_ScMem) Then z_ScMem = 0&
            zMap_VFunction = z_ScMem
            'If z_ScMem = 0 Then Debug.Print "already erased : "; z_ScMem
        End If
    End If
    Exit Function                                               'Exit returning the thunk address
    
Catch:
    ' error ignored when zUnThunk is called, error handled there
    If Not bIgnoreErrors Then zError "zMap_VFunction", "Thunk type for " & vType & " does not exist"
End Function

Private Property Get zData(ByVal nIndex As Long, ByVal thunkPtr As Long) As Long
  RtlMoveMemory VarPtr(zData), thunkPtr + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal thunkPtr As Long, ByVal nValue As Long)
  RtlMoveMemory thunkPtr + (nIndex * 4), VarPtr(nValue), 4
End Property

'Error handler
Private Sub zError(ByRef sRoutine As String, ByVal sMsg As String)
  ' \\LaVolpe -  Note. These two lines can be rem'd out if you so desire. But don't remove the routine
  'App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  'pdMsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
  
  'Tanner's modification: do not display message boxes on errors.  Debug output is preferred.
  Debug.Print ">> Subclasser returned the following message: " & sRoutine & ": " & sMsg
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
  If asUnicode Then
    zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
  Else
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
  End If
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      ' \\LaVolpe - Added propertypage offset
      If Not zProbe(nAddr + &H710, i, bSub) Then                            'Probe for a PropertyPage method
        If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
            Exit Function                                                   'Bail...
        End If
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  'J = I + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  j = i + 1024 * 2                                                     'Set a reasonable limit, scan 512 vTable entries
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                               'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Do                                                             'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Function zInIDE(ByRef bIDE As Boolean) As Boolean
    ' only called in IDE, never called when compiled
    bIDE = True
    zInIDE = bIDE
End Function

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType, Optional ByVal oCallback As Object)

    ' thunkID, depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback

    Const IDX_SHUTDOWN  As Long = 1
    Const MEM_RELEASE As Long = &H8000&             'Release allocated memory flag
    Dim z_ScMem       As Long         'Thunk base address
    
    z_ScMem = zMap_VFunction(thunkID, vType, oCallback, True)
    Select Case vType
    Case SubclassThunk
        If z_ScMem Then       'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN, z_ScMem) = 1            'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE, z_ScMem   'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE, z_ScMem   'Delete all after messages
        End If
        
        'EDIT BY TANNER: check to see if the thunk exists before attempting to remove it.  Not sure why
        ' this has suddenly become necessary, but I was getting errors from this line before implementing
        ' a collection check.
        If ExistsNonObject(z_scFunk, "h" & thunkID) Then z_scFunk.Remove "h" & thunkID        'Remove the specified thunk from the collection
        
        
    Case HookThunk
        If z_ScMem Then       'Ensure that the thunk hasn't already released its memory
            ' if not unhooked, then unhook now
            If zData(IDX_SHUTDOWN, z_ScMem) = 0 Then UnhookWindowsHookEx zData(IDX_PREVPROC, z_ScMem)
            If zData(0, z_ScMem) = 0 Then                   ' not recursing then
                VirtualFree z_ScMem, 0, MEM_RELEASE 'Release allocated memory
                z_hkFunk.Remove "h" & thunkID       'Remove the specified thunk from the collection
            Else
                zData(IDX_SHUTDOWN, z_ScMem) = 1            ' Set the shutdown indicator
                zData(IDX_ATABLE, z_ScMem) = 0              ' want no more After messages
                zData(IDX_BTABLE, z_ScMem) = 0              ' want no more Before messages
                
                ' when zTerminate is called this thunk's memory will be released
                
                'ADDED BY TANNER: let the debugger know that the hook is still technically active!
                Debug.Print "cSelfSubHookCallback couldn't release hook immediately, as another hook is stil active.  Will retry on next _Terminate event."
                
            End If
        Else
            z_hkFunk.Remove "h" & thunkID       'Remove the specified thunk from the collection
        End If
    Case CallbackThunk
        If z_ScMem Then       'Ensure that the thunk hasn't already released its memory
            VirtualFree z_ScMem, 0, MEM_RELEASE 'Release allocated memory
        End If
        z_cbFunk.Remove "h" & ObjPtr(oCallback) & "." & thunkID           'Remove the specified thunk from the collection
    End Select

End Sub

Public Function isHookingActive() As Boolean

    If (z_hkFunk Is Nothing) Then
        isHookingActive = False
    Else
        isHookingActive = True
    End If

End Function

Private Sub zTerminateThunks(ByVal vType As eThunkType)

    Dim i As Long
    Dim oCallback As Object
    Dim thunkCol As Collection
    Dim z_ScMem       As Long         'Thunk base address
    Const INDX_OWNER As Long = 0
    
    Select Case vType
    Case SubclassThunk
        Set thunkCol = z_scFunk
    Case HookThunk
        Set thunkCol = z_hkFunk
    Case CallbackThunk
        Set thunkCol = z_cbFunk
    Case Else
        Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
      With thunkCol
        For i = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
          z_ScMem = .Item(i)                          'Get the thunk address
          If IsBadCodePtr(z_ScMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
            Select Case vType
                Case SubclassThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), SubclassThunk    'Unsubclass
                Case HookThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), HookThunk            'Unhook
                Case CallbackThunk
                    ' zUnThunk expects object not pointer, convert pointer to object
                    RtlMoveMemory VarPtr(oCallback), VarPtr(zData(INDX_OWNER, z_ScMem)), 4&
                    zUnThunk zData(IDX_CALLBACKORDINAL, z_ScMem), CallbackThunk, oCallback ' release callback
                    ' remove the object pointer reference
                    RtlMoveMemory VarPtr(oCallback), VarPtr(INDX_OWNER), 4&
            End Select
          End If
        Next i                                        'Next member of the collection
      End With
      Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If


End Sub

'-SelfCallback code------------------------------------------------------------------------------------
'-The following routines are exclusively for the scb_SetCallbackAddr routines----------------------------
Public Function scb_SetCallbackAddr(ByVal nParamCount As Long, _
                     Optional ByVal nOrdinal As Long = 1, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True, _
                     Optional ByVal bIsTimerCallback As Boolean) As Long   'Return the address of the specified callback thunk
    '*************************************************************************************************
    '* nParamCount  - The number of parameters that will callback
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety   - Optional, set to false to disable IDE protection.
    '* bIsTimerCallback - optional, set to true for extra protection when used as a SetTimer callback
    '       If True, timer will be destroyed when IDE/app terminates. See scb_ReleaseCallback.
    '*************************************************************************************************
    ' Callback procedure must return a Long even if, per MSDN, the callback procedure is a Sub vs Function
    ' The number of parameters and their types are dependent on the individual callback procedures
    
    Const MEM_LEN     As Long = IDX_CALLBACKORDINAL * 4 + 4     'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                         'Commit allocated memory
    Const SUB_NAME      As String = "scb_SetCallbackAddr"       'This routine's name
    Const INDX_OWNER    As Long = 0                             'Thunk data index of the Owner object's vTable address
    Const INDX_CALLBACK As Long = 1                             'Thunk data index of the EbMode function address
    Const INDX_EBMODE   As Long = 2                             'Thunk data index of the IsBadCodePtr function address
    Const INDX_BADPTR   As Long = 3                             'Thunk data index of the IsBadCodePtr function address
    Const INDX_KT       As Long = 4                             'Thunk data index of the KillTimer function address
    Const INDX_EBX      As Long = 6                             'Thunk code patch index of the thunk data
    Const INDX_PARAMS   As Long = 18                            'Thunk code patch index of the number of parameters expected in callback
    Const INDX_PARAMLEN As Long = 24                            'Thunk code patch index of the bytes to be released after callback
    Const PROC_OFF      As Long = &H14                          'Thunk offset to the callback execution address

    Dim z_ScMem       As Long                                   'Thunk base address
    Dim z_Cb()    As Long                                       'Callback thunk array
    Dim nCallback As Long
    Dim bIDE      As Boolean
      
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    If z_cbFunk Is Nothing Then
        Set z_cbFunk = New Collection           'If this is the first time through, do the one-time initialization
    Else
        On Error Resume Next                    'Catch already initialized?
        z_ScMem = z_cbFunk.Item("h" & ObjPtr(oCallback) & "." & nOrdinal) 'Test it
        If Err = 0 Then
            scb_SetCallbackAddr = z_ScMem + PROC_OFF  'we had this one, just reference it
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    If nParamCount < 0 Then                     ' validate parameters
        zError SUB_NAME, "Invalid Parameter count"
        Exit Function
    End If
    
    nCallback = zAddressOf(oCallback, nOrdinal)         'Get the callback address of the specified ordinal
    If nCallback = 0 Then
        zError SUB_NAME, "Callback address not found."
        Exit Function
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
        
    If z_ScMem = 0& Then
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError  ' oops
        Exit Function
    End If
    z_cbFunk.Add z_ScMem, "h" & ObjPtr(oCallback) & "." & nOrdinal 'Add the callback/thunk-address to the collection
        
    ReDim z_Cb(0 To IDX_CALLBACKORDINAL) As Long          'Allocate for the machine-code array
    
    ' Create machine-code array
    z_Cb(5) = &HBB60E089: z_Cb(7) = &H73FFC589: z_Cb(8) = &HC53FF04: z_Cb(9) = &H59E80A74: z_Cb(10) = &HE9000000
    z_Cb(11) = &H30&: z_Cb(12) = &H87B81: z_Cb(13) = &H75000000: z_Cb(14) = &H9090902B: z_Cb(15) = &H42DE889: z_Cb(16) = &H50000000: z_Cb(17) = &HB9909090: z_Cb(19) = &H90900AE3
    z_Cb(20) = &H8D74FF: z_Cb(21) = &H9090FAE2: z_Cb(22) = &H53FF33FF: z_Cb(23) = &H90909004: z_Cb(24) = &H2BADC261: z_Cb(25) = &H3D0853FF: z_Cb(26) = &H1&: z_Cb(27) = &H23DCE74: z_Cb(28) = &H74000000: z_Cb(29) = &HAE807
    z_Cb(30) = &H90900000: z_Cb(31) = &H4589C031: z_Cb(32) = &H90DDEBFC: z_Cb(33) = &HFF0C75FF: z_Cb(34) = &H53FF0475: z_Cb(35) = &HC310&

    z_Cb(INDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", False)
    z_Cb(INDX_OWNER) = ObjPtr(oCallback)                  'Set the Owner
    z_Cb(INDX_CALLBACK) = nCallback                       'Set the callback address
    z_Cb(IDX_CALLBACKORDINAL) = nOrdinal                  'Cache ordinal used for zTerminateThunks
      
    If bIdeSafety = True Then                                           'If the user wants IDE protection
        Debug.Assert zInIDE(bIDE)
        If bIDE = True Then z_Cb(INDX_EBMODE) = zFnAddr("vba6", "EbMode", False) 'Store the EbMode function address in the thunk data
    End If
    If bIsTimerCallback Then
        z_Cb(INDX_KT) = zFnAddr("user32", "KillTimer", False)
    End If
        
    z_Cb(INDX_PARAMS) = nParamCount                         'Set the parameter count
    RtlMoveMemory VarPtr(z_Cb(INDX_PARAMLEN)) + 2, VarPtr(nParamCount * 4), 2&
      
    z_Cb(INDX_EBX) = z_ScMem                                'Set the data address relative to virtual memory pointer
      
    RtlMoveMemory z_ScMem, VarPtr(z_Cb(INDX_OWNER)), MEM_LEN 'Copy thunk code to executable memory
    scb_SetCallbackAddr = z_ScMem + PROC_OFF                       'Thunk code start address

End Function

Public Sub scb_ReleaseCallback(ByVal nOrdinal As Long, Optional ByVal oCallback As Object)
    ' can be made public, can be removed & zUnThunk can be called instead
    ' NEVER call this from within the callback routine itself
    
    ' oCallBack is the object containing nOrdinal to be released
    ' if oCallback was already closed (say it was a class or form), then you won't be
    '   able to release it here, but it will be released when zTerminateThunks is
    '   eventually called
    
    ' Special Warning. If the callback thunk is used for a recurring callback (i.e., Timer),
    ' then ensure you terminate what is using the callback before releasing the thunk,
    ' otherwise you are subject to a crash
    zUnThunk nOrdinal, CallbackThunk, oCallback
End Sub
Public Sub scb_TerminateCallbacks()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    zTerminateThunks CallbackThunk
End Sub

'This handy function comes from http://stackoverflow.com/questions/40651/check-if-a-record-exists-in-a-vb6-collection
Private Function ExistsNonObject(col, Index) As Boolean
    On Error GoTo ExistsNonObjectErrorHandler
    Dim v As Variant

    v = col(Index)
    ExistsNonObject = True
    Exit Function

ExistsNonObjectErrorHandler:
    ExistsNonObject = False
End Function
