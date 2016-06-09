VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUniversalDLLCalls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'You may download the original version of this class from the following link (good as of May '16):
' http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)

'This class was originally written by vbforums.com user LaVolpe.  To the best of my knowledge, the code has
' been released into the public domain.

'Many thanks to LaVolpe for this helpful DLL calling code workaround, which PhotoDemon uses for messy things
' like faking cdecl versions of VB callbacks (which are forcibly implemented as stdcall).

Option Explicit

' for documentation on the main API DispCallFunc... http://msdn.microsoft.com/en-us/library/windows/desktop/ms221473%28v=vs.85%29.aspx
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long

' APIs used for _CDecl callback workarounds. See ThunkFor_CDeclCallbackToVB & ThunkRelease_CDECL
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Public Enum CALLINGCONVENTION_ENUM
  ' http://msdn.microsoft.com/en-us/library/system.runtime.interopservices.comtypes.callconv%28v=vs.110%29.aspx
  CC_FASTCALL = 0&
  CC_CDECL
  CC_PASCAL
  CC_MACPASCAL
  CC_STDCALL                        ' typical windows APIs
  CC_FPFASTCALL
  CC_SYSCALL
  CC_MPWCDECL
  CC_MPWPASCAL
End Enum
Public Enum CALLRETURNTUYPE_ENUM
    CR_None = vbEmpty
    CR_LONG = vbLong
    CR_BYTE = vbByte
    CR_INTEGER = vbInteger
    CR_SINGLE = vbSingle
    CR_DOUBLE = vbDouble
    CR_CURRENCY = vbCurrency
    ' if the value you need isn't in above list, you can pass the value manually to the
    ' CallFunction_DLL method below. For additional values, see:
    ' http://msdn.microsoft.com/en-us/library/cc237865.aspx
End Enum
Public Enum STRINGPARAMS_ENUM
    STR_NONE = 0&
    STR_ANSI
    STR_UNICODE
End Enum

Private m_DLLname As String     ' track last DLL loaded by this class
Private m_Mod As Long           ' reference to loaded module
Private m_Release As Boolean    ' whether or not we unload the module/dll

Public Function CallFunction_DLL(ByVal LibName As String, ByVal FunctionName As String, _
                            ByVal HasStringParams As STRINGPARAMS_ENUM, _
                            ByVal FunctionReturnType As CALLRETURNTUYPE_ENUM, _
                            ByVal CallConvention As CALLINGCONVENTION_ENUM, _
                            ParamArray FunctionParameters() As Variant) As Variant
                            
' Used to call standard dlls, not active-x or COM objects

' Return value. Will be a variant containing a value of FunctionReturnType
'   If this method fails, the return value will always be Empty. This can be verified by checking
'       the Err.LastDLLError value. It will be non-zero if the function failed else zero.
'   If the method succeeds, there is no guarantee that the function you called succeeded. The
'       success/failure of that function would be indicated by this method's return value.
'   If calling a sub vs function & this method succeeds, the return value will be zero.
'   Summarizing: if method fails to execute, Err.LastDLLError value will be non-zero
'       If method executes ok, return value is from the DLL you called

' Parameters:
'   LibName. The dll name. You should always include the extension else DLL is used
'       See LoadLibrary documentation for more: http://msdn.microsoft.com/en-us/library/windows/desktop/ms684175%28v=vs.85%29.aspx
'   FunctionName. The DLL function to call. This is case-senstiive
'       To call a function by ordinal, prefix it with a hash symbol, i.e., #124
'   HasStringParams. Provide one of the 3 available values
'       STR_NONE. No parameters are strings or all strings are passed via StrPtr()
'       STR_UNICODE. Any passed string values are for a Unicode function, i.e., SetWindowTextW
'       STR_ANSI. Any passed string values are for an ANSI function, i.e., SetWindowTextA
'       Important: If you pass one of FunctionParameters a String variable, you must include
'           STR_UNICODE or STR_ANSI depending on what version function you are calling
'           See the FunctionParameters section below for more
'   FunctionReturnType. Describes what variant type the called function returns
'       If calling a subroutine that does not return a value, use CR_None
'   CallConvention. One of various DLL calling conventions
'       You must know the calling convention of the function you are calling and the number
'           of parameters, along with the parameter variable type
'   FunctionParameters. The values and variant type for each value as required by the function
'       you are calling. This is important. Passing incorrect variable types can cause crashes.
'       There is no auto-conversion like VB would do for you if you were to call an API function.
'       To ensure you pass the correct variable type, use VBs conversion routines:
'           Passing a Long? CLng(10), CLng(x). Passing an Integer? CInt(10), CInt(x)
'       Special cases:
'           UDTs (structures). Pass these using VarPtr(), i.e., VarPtr(uRect)
'               If UDT members contain static size strings, you should declare those string members
'               as Byte arrays instead. When array is filled in by the function you called,
'               you can use StrConv() to convert array to string.
'               If UDT members contain dynamic size strings, you should declare those as Long.
'               When the function returns, you can use built-in functions within this class to
'               retrieve the string from the pointer provided to your UDT.
'           Arrays. DO NOT pass the array. Pass only a pointer to the first member of the array,
'               i.e., VarPtr(myArray(0)), VarPtr(myArray(0,0)), etc
'           Strings for ANSI functions.
'               1) Passing by variable name or value? i.e., strContent, "Edit", etc
'                   The string needs to be converted to ANSI, and this class will do that for you
'                   if you also pass HasStringParams as STR_ANSI. Otherwise, do NOT pass strings
'                   for ANSI functions by variable name or value. When passed by variable name,
'                   the variable contents are changed to 1 byte per character. To prevent this,
'                   pass the variable name inside parentheses, i.e., (myVariable)
'               2) Passing by StrPtr()? i.e, StrPtr(strContent), StrPtr("Edit")
'                   If the function you are calling needs the string contents, then do NOT pass
'                   the string this way. You must first convert it to ANSI. Else, you could
'                   pass it as desribed in #1 above.
'               Rule-of-Thumb. If string is just a buffer, pass it by StrPtr(), then on return,
'                   use VB's StrConv() to convert it from ANSI to unicode. Otherwise, pass the
'                   string by variable name or value
'           Strings for Unicode functions
'               1) Passing by variable name or value? i.e., strContent, "Edit", etc
'                   Internally, the string must be passed to the function ByVal via StrPtr().
'                   This class will do that, but it is faster (less code) if you pass all strings
'                   for unicode functions via StrPtr()
'               2) Passing by StrPtr()? i.e, StrPtr(strContent), StrPtr("Edit")
'                   Less code required, fastest method, no conversions required at all
'               Rule-of-Thumb. All strings for unicode functions should be passed via StrPtr()
'           Numeric values vs. variables. Be aware of the variable type of the number you pass.
'               Depending on the value of the number, it may be Integer, Long, Double, etc.
'               Numbers in range -32768 to 32767 are Integer, from -2147483648 to 2147483647 are Long
'               Fractional/decimal numbers are Double
'               If function parameter expects Long, don't pass just 5, pass 5& or CLng(5)
'           Numbers as variables. Be sure the variable type matches the parameter type, i.e.,
'               dont pass variables declared as Variant to a function expecting Long

    '// minimal sanity check for these 4 parameters:
    If LibName = vbNullString Then Exit Function
    If FunctionName = vbNullString Then Exit Function
    If Not (FunctionReturnType And &HFFFF0000) = 0& Then Exit Function ' can only be 4 bytes
    If HasStringParams < STR_NONE Or HasStringParams > STR_UNICODE Then Exit Function
    
    Dim sText As String, lStrPtr As Long, lValue As Long
    Const VT_BYREF As Long = &H4000&
    
    Dim hMod As Long, fPtr As Long
    Dim pIndex As Long, pCount As Long
    
    Dim vParamPtr() As Long, vParamType() As Integer
    Dim vRtn As Variant, vParams() As Variant
    
    '// determine if we will be loading this or already loaded
    If LibName = m_DLLname Then
        hMod = m_Mod                                    ' already loaded
    Else
        If Not m_Mod = 0& Then                          ' reset m_Mod & m_Release
            If m_Release = True Then FreeLibrary m_Mod
            m_Mod = 0&: m_Release = False
        End If
        hMod = GetModuleHandle(LibName)                 ' loaded in process already?
        If hMod = 0& Then                               ' if not, load it now
            hMod = LoadLibrary(LibName)
            If hMod = 0& Then Exit Function
            m_Release = True                            ' need to use FreeLibrary at some point
        End If
        m_Mod = hMod                                    ' cache hMod & LibName
        m_DLLname = LibName
    End If
    fPtr = GetProcAddress(hMod, FunctionName)           ' get the function pointer (Case-Sensitive)
    If fPtr = 0& Then Exit Function                     ' abort if failure
    
    vParams() = FunctionParameters()                    ' copy passed parameters, if any
    pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
    If HasStringParams > STR_NONE Then                  ' patch to ensure Strings passed as handles
        For pIndex = 0& To pCount - 1&                  ' for each string param, get its StrPtr
            If VarType(FunctionParameters(pIndex)) = vbString Then
                CopyMemory lValue, ByVal VarPtr(FunctionParameters(pIndex)), 2&
                If (lValue And VT_BYREF) = 0& Then      ' else variant has pointer to StrPtr
                    lValue = VarPtr(FunctionParameters(pIndex)) + 8&
                Else
                    CopyMemory lValue, ByVal VarPtr(FunctionParameters(pIndex)) + 8&, 4&
                End If
                CopyMemory lStrPtr, ByVal lValue, 4&    ' get the StrPtr
                If lStrPtr > 0& Then                    ' if not null then
                    If HasStringParams = STR_ANSI Then  ' convert Unicode to ANSI
                        sText = FunctionParameters(pIndex) ' then re-write the passd String to ANSI
                        FillMemory ByVal lStrPtr, LenB(sText), 0
                        sText = StrConv(sText, vbFromUnicode)
                        CopyMemory ByVal lStrPtr, ByVal StrPtr(sText), LenB(sText)
                    End If
                End If
                vParams(pIndex) = lStrPtr               ' reference the StrPtr
            End If
        Next
    End If
                                                        ' fill in rest of APIs parameters
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
    lValue = DispCallFunc(0&, fPtr, CallConvention, FunctionReturnType, _
                         pCount, vParamType(0), vParamPtr(0), vRtn)
        
    If lValue = 0& Then                                 ' 0 = S_OK
        If FunctionReturnType = CR_None Then
            CallFunction_DLL = lValue
        Else
            CallFunction_DLL = vRtn                     ' return result
        End If
    Else
        SetLastError lValue                             ' set error & return Empty
    End If
    
End Function

Public Function CallFunction_COM(ByVal InterfacePointer As Long, ByVal VTableOffset As Long, _
                            ByVal FunctionReturnType As CALLRETURNTUYPE_ENUM, _
                            ByVal CallConvention As CALLINGCONVENTION_ENUM, _
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
'       These offsets are generally in multiples of 4. Value cannot be negative.
'   For the remaining parameters, see the details withn the CallFunction_DLL method.
'       They are the same with one exception: strings. Pass the string variable name or value

    '// minimal sanity check for these 4 parameters:
    If InterfacePointer = 0& Or VTableOffset < 0& Then Exit Function
    If Not (FunctionReturnType And &HFFFF0000) = 0& Then Exit Function ' can only be 4 bytes

    Dim pIndex As Long, pCount As Long
    Dim vParamPtr() As Long, vParamType() As Integer
    Dim vRtn As Variant, vParams() As Variant
    
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
    pIndex = DispCallFunc(InterfacePointer, VTableOffset, CallConvention, FunctionReturnType, _
                          pCount, vParamType(0), vParamPtr(0), vRtn)
        
    If pIndex = 0& Then                                 ' 0 = S_OK
        CallFunction_COM = vRtn                         ' return result
    Else
        SetLastError pIndex                             ' set error & return Empty
    End If

End Function

Public Function PointerToStringA(ByVal ANSIpointer As Long) As String
    ' courtesy function provided for your use as needed
    ' ANSIpointer must be a pointer to an ANSI string (1 byte per character)
    Dim lSize As Long, sANSI As String
    If Not ANSIpointer = 0& Then
        lSize = lstrlenA(ANSIpointer)
        If lSize > 0& Then
            sANSI = String$(lSize \ 2& + 1&, vbNullChar)
            CopyMemory ByVal StrPtr(sANSI), ByVal ANSIpointer, lSize
            PointerToStringA = Left$(StrConv(sANSI, vbUnicode), lSize)
        End If
    End If
End Function

Public Function PointerToStringW(ByVal UnicodePointer As Long) As String
    ' courtesy function provided for your use as needed
    ' UnicodePointer must be a pointer to an unicode string (2 bytes per character)
    Dim lSize As Long
    If Not UnicodePointer = 0& Then
        lSize = lstrlenW(UnicodePointer)
        If lSize > 0& Then
            PointerToStringW = Space$(lSize)
            CopyMemory ByVal StrPtr(PointerToStringW), ByVal UnicodePointer, lSize * 2&
        End If
    End If
End Function

Public Function ThunkFor_CDeclCallbackToVB(ByVal VBcallbackPointer As Long, _
                            ByVal CallbackParamCount As Long) As Long
                            
    ' this method is a workaround for cases where you are calling a CDECL function that requests
    ' a callback function address in CDECL calling convention.
    ' Ex: qsort in msvcrt20.dll uses such a callback & qsort function description found here:
    ' http://msdn.microsoft.com/en-us/library/zes7xw0h.aspx
    
    ' Important notes:
    ' 1) DO NOT USE this workaround when any function accepting a callback pointer,
    '       uses stdCall calling convention to that pointer. DO NOT USE this function
    '       for other than CDECL functions calling back to VB
    ' 2) This method's return value MUST BE RELEASED via a call to ThunkRelease_CDECL method
    ' 3) The VB callback function must be a function vs. sub, even if the the callback
    '       definition describes it as a sub, i.e., returns no value, void
    ' 4) The thunk prevents VB's stack cleaning by copying first, then replacing it after VB returns
    
    ' Parameters:
    '   VBcallbackPointer: the VB callback address. If function exists in a bas module, then
    '       this would be the return value of your AddressOf call. If using thunks to get addresses
    '       from class methods, then pass that thunk address as appropriate
    '   CallbackParamCount: Number of parameters your VB method accepts. This cannot be dynamic
    
    ' sample call: assume that vbCallBackFunction is a Public function within a bas module
    ' -------------------------------------------------------------------------------------
    ' Dim lCallback As Long, lThunkAddress As Long, lResult As Long
    '   lCallback = thisClass.ThunkFor_CDeclCallbackToVB(AddressOf vbCallBackFunction, 2&, lThunkAddress)
    '   ' now call your CDECL function, passing lCallback as the required callback address paramter,
    '   '    in whatever param position it is required
    '   lResult = thisClass.CallFunction_DLL("someCDECL.dll", "functionName", STR_NONE, CR_LONG, _
    '                                        CC_CDECL, params, lCallback)
    '   ' destroy the thunk when no longer needed
    '   Call thisClass.ThunkRelease_CDECL(lThunkAddress)
    
    
    ' sanity checks on passed parameters
    If VBcallbackPointer = 0& Or CallbackParamCount < 0& Or CallbackParamCount > 63& Then Exit Function
    ' FYI: Why is 63 the max count? CallbackParamCount stored in the thunk as unsigned byte: 63*4 =252
    
    Dim fPtr As Long, tCode(0 To 2) As Currency
    
    fPtr = VirtualAlloc(0&, 28&, &H1000&, &H40&)            ' reserve memory for our virtual function
  
    tCode(0) = 465203369712025.6232@                        ' thunk code is small, 28 bytes
    tCode(1) = -140418483381718.8329@
    tCode(2) = -4672484613390.9419@
    CopyMemory ByVal fPtr, ByVal VarPtr(tCode(0)), 24&      ' copy to virt memmory
    CopyMemory ByVal fPtr + 24&, &HC30672, 4&               ' copy final 4 bytes also

    ' thunk uses relative address to VB function address, calc relative address & patch the thunk
    CopyMemory ByVal fPtr + 10&, VBcallbackPointer - fPtr - 14&, 4&
    CopyMemory ByVal fPtr + 16&, CallbackParamCount * 4&, 1& ' patch thunk's param count (stack adjustment)

    ThunkFor_CDeclCallbackToVB = fPtr
    
    '  FYI: Thunk described below. Paul Caton's work found here:
    '  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=49776&lngWId=1
    '==============================================================================
    '                                ;FASM syntax
    '        use32                   ;32bit
    '        call    L1              ;Call the next instruction
    '    L1: pop eax                 ;Pop the return address into eax (eax = L1)
    '        pop dword [eax+(L3-L1)] ;Pop the calling cdecl function's return address to the save location
    '        db  0E8h                ;Op-code for a relative address call
    '        dd  55555555h           ;Address of target vb callback function, patched at run-time
    '        sub esp, 55h            ;Unfix the stack, our caller expects to do it, patched at runtime
    '        call    L2              ;Call the next instruction
    '    L2: pop edx                 ;Pop the return address into edx (edx = L2)
    '        push dword [edx+(L3-L2)];Push the saved return address, the stack is now as it was on entry to callback_wrapper
    '        ret                     ;Return to caller
    '        db  0                   ;Alignment pad
    '    L3: dd  0                   ;Return address of the cdecl caller saved here
    '==============================================================================
End Function

Public Sub ThunkRelease_CDECL(ByVal ThunkCallBackAddress As Long)
    ' Used to release memory created during a call to the ThunkFor_CDeclCallbackToVB method.
    ' The parameter passed here must be the return value of the ThunkFor_CDeclCallbackToVB method
    If Not ThunkCallBackAddress = 0& Then VirtualFree ThunkCallBackAddress, 0&, &H8000&
End Sub

Private Sub Class_Terminate()
    If Not m_Mod = 0& Then
        If m_Release = True Then FreeLibrary m_Mod
    End If
End Sub


