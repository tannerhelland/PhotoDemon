Attribute VB_Name = "VBHacks"
'***************************************************************************
'Misc VB6 Hacks
'Copyright 2016-2019 by Tanner Helland
'Created: 06/January/16
'Last updated: 26/January/19
'Last update: helper function for fast endian swapping (16-bit boundaries); this is used constantly in PSD parsing
'
'PhotoDemon relies on a lot of "not officially sanctioned" VB6 behavior to enable various optimizations and C-style
' code techniques. If a function's primary purpose is a VB6-specific workaround, I prefer to move it here, so I
' don't clutter up purposeful modules with obscure, VB-specific hackery.
'
'Note that some code here may seem redundant (e.g. identical functions suffixed by data type, instead of declared
' "As Any") but that's by design - e.g. to improve safety since these techniques are almost always crash-prone if
' used incorrectly or imprecisely.
'
'A number of the techniques in this module were devised with help from Karl E. Peterson's work at http://vb.mvps.org/
' Thank you, Karl!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Type winMsg
    hWnd As Long
    sysMsg As Long
    wParam As Long
    lParam As Long
    msgTime As Long
    ptX As Long
    ptY As Long
End Type

'Some APIs are used *so* frequently throughout PD that we declare them publicly
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDst As Any, ByRef lpSrc As Any, ByVal byteLength As Long)
Public Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal numOfBytes As Long, ByVal fillValue As Byte)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByVal dstPointer As Long, ByVal numOfBytes As Long)

Public Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (ptr() As Any) As Long

Public Declare Sub GetMem1 Lib "msvbvm60" (ByVal ptrSrc As Long, ByRef dstByte As Byte)
Public Declare Sub GetMem1_Ptr Lib "msvbvm60" Alias "GetMem1" (ByVal ptrSrc As Long, ByVal ptrDst1 As Long)
Public Declare Sub GetMem2 Lib "msvbvm60" (ByVal ptrSrc As Long, ByRef dstInteger As Integer)
Public Declare Sub GetMem2_Ptr Lib "msvbvm60" Alias "GetMem2" (ByVal ptrSrc As Long, ByVal ptrDst2 As Long)
Public Declare Sub GetMem4 Lib "msvbvm60" (ByVal ptrSrc As Long, ByRef dstValue As Long)
Public Declare Sub GetMem4_Ptr Lib "msvbvm60" Alias "GetMem4" (ByVal ptrSrc As Long, ByVal ptrDst4 As Long)
Public Declare Sub GetMem8 Lib "msvbvm60" (ByVal ptrSrc As Long, ByRef dstCurrency As Currency)
Public Declare Sub GetMem8_Ptr Lib "msvbvm60" Alias "GetMem8" (ByVal ptrSrc As Long, ByVal ptrDst8 As Long)
Public Declare Sub PutMem1 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Byte)
Public Declare Sub PutMem2 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Integer)
Public Declare Sub PutMem4 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Long)
Public Declare Sub PutMem8 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Currency)

'Private declares follow:

'We use Karl E. Peterson's approach of declaring subclass functions by ordinal, per the documentation at http://vb.mvps.org/samples/HookXP/
Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function RtlCompareMemory Lib "ntdll" (ByVal ptrSource1 As Long, ByVal ptrSource2 As Long, ByVal Length As Long) As Long

Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As VbVarType, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Sub SafeArrayLock Lib "oleaut32" (ByVal ptrToSA As Long)
Private Declare Sub SafeArrayUnlock Lib "oleaut32" (ByVal ptrToSA As Long)

Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, ByRef hGlobal As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long

Private Declare Function DispatchMessageA Lib "user32" (ByRef lpMsg As winMsg) As Long
Private Declare Function DispatchMessageW Lib "user32" (ByRef lpMsg As winMsg) As Long
Private Declare Function PeekMessageA Lib "user32" (ByRef lpMsg As winMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function PeekMessageW Lib "user32" (ByRef lpMsg As winMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (ByRef lpMsg As winMsg) As Long

Private Declare Function htonl Lib "Ws2_32" (ByVal srcLong As Long) As Long
Private Declare Function htons Lib "Ws2_32" (ByVal srcShort As Integer) As Integer

Private Const GMEM_MOVEABLE As Long = &H2&
Public Const WM_NCDESTROY As Long = &H82&
Private Const WH_KEYBOARD As Long = 2

'Unsigned arithmetic helpers
Private Const SIGN_BIT As Long = &H80000000

'Higher-performance timing functions are also handled by this class.  Note that you *must* initialize the timer engine
' before requesting any time values, or crashes will occurs because the frequency timer is 0.
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
Private m_TimerFrequency As Currency

'Because AddressOf doesn't work in classes, we have to jump through some hoops to allow class-based keyboard hooking
Private m_EditBoxRef As pdEditBoxW
Private m_AcceleratorRef As pdAccelerator
Private m_PDIKRef As pdInputKeyboard

'Point an internal 1D VB array at some other arbitrary 1D array.  The new array should *NOT* be initialized
' or it will leak memory.  Any arrays aliased this way must be freed via Unalias1DArray or VB will crash.
Public Sub Alias1DArray_Byte(ByRef orig1DArray() As Byte, ByRef new1DArray() As Byte, ByRef newArraySA As SafeArray1D)
    
    'Retrieve a copy of the original 2D array's SafeArray struct
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig1DArray()), ptrSrc
    CopyMemory ByVal VarPtr(newArraySA), ByVal ptrSrc, LenB(newArraySA)
    
    'newArraySA now contains the full SafeArray of the original array.  Copy this over our current array.
    CopyMemory ByVal VarPtrArray(new1DArray()), VarPtr(newArraySA), 4&
    
    'Add a lock to the original array, to prevent potential crashes from unknowing users.  (Thanks to @Kroc for this tip.)
    SafeArrayLock ptrSrc
    
End Sub

'Point an internal 2D array at some other 2D array.  Any arrays aliased this way must be freed via Unalias2DArray,
' or VB will crash.
Public Sub Alias2DArray_Byte(ByRef orig2DArray() As Byte, ByRef new2DArray() As Byte, ByRef newArraySA As SafeArray2D)
    
    'Retrieve a copy of the original 2D array's SafeArray struct
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig2DArray()), ptrSrc
    CopyMemory ByVal VarPtr(newArraySA), ByVal ptrSrc, LenB(newArraySA)
    
    'newArraySA now contains the full SafeArray of the original array.  Copy this over our current array.
    CopyMemory ByVal VarPtrArray(new2DArray()), VarPtr(newArraySA), 4&
    
    'Add a lock to the original array, to prevent potential crashes from unknowing users.  (Thanks to @Kroc for this tip.)
    SafeArrayLock ptrSrc
    
End Sub

Public Sub Alias2DArray_Integer(ByRef orig2DArray() As Integer, ByRef new2DArray() As Integer, ByRef newArraySA As SafeArray2D)
    
    'Retrieve a copy of the original 2D array's SafeArray struct
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig2DArray()), ptrSrc
    CopyMemory ByVal VarPtr(newArraySA), ByVal ptrSrc, LenB(newArraySA)
    
    'newArraySA now contains the full SafeArray of the original array.  Copy this over our current array.
    CopyMemory ByVal VarPtrArray(new2DArray()), VarPtr(newArraySA), 4&
    
    'Add a lock to the original array, to prevent potential crashes from unknowing users.  (Thanks to @Kroc for this tip.)
    SafeArrayLock ptrSrc
    
End Sub

Public Sub Alias2DArray_Long(ByRef orig2DArray() As Long, ByRef new2DArray() As Long, ByRef newArraySA As SafeArray2D)
    
    'Retrieve a copy of the original 2D array's SafeArray struct
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig2DArray()), ptrSrc
    CopyMemory ByVal VarPtr(newArraySA), ByVal ptrSrc, LenB(newArraySA)
    
    'newArraySA now contains the full SafeArray of the original array.  Copy this over our current array.
    CopyMemory ByVal VarPtrArray(new2DArray()), VarPtr(newArraySA), 4&
    
    'Add a lock to the original array, to prevent potential crashes from unknowing users.  (Thanks to @Kroc for this tip.)
    SafeArrayLock ptrSrc
    
End Sub

'Counterpart to Alias1DArray_ functions, above.  Do NOT call this function on arrays that were not originally
' processed by an Alias21Array_ function.
Public Sub Unalias1DArray_Byte(ByRef orig1DArray() As Byte, ByRef new1DArray() As Byte)
    
    'Wipe the array pointer
    CopyMemory ByVal VarPtrArray(new1DArray), 0&, 4&
    
    'Remove a lock from the original array; this allows the user to safely release the array on their own terms
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig1DArray()), ptrSrc
    SafeArrayUnlock ptrSrc
    
End Sub

'Counterparts to Alias2DArray_ functions, above.  Do NOT call this function on arrays that were not originally
' processed by an Alias2DArray_ function.
Public Sub Unalias2DArray_Byte(ByRef orig2DArray() As Byte, ByRef new2DArray() As Byte)
    
    'Wipe the array pointer
    CopyMemory ByVal VarPtrArray(new2DArray), 0&, 4&
    
    'Remove a lock from the original array; this allows the user to safely release the array on their own terms
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig2DArray()), ptrSrc
    SafeArrayUnlock ptrSrc
    
End Sub

Public Sub Unalias2DArray_Integer(ByRef orig2DArray() As Integer, ByRef new2DArray() As Integer)
    
    'Wipe the array pointer
    CopyMemory ByVal VarPtrArray(new2DArray), 0&, 4&
    
    'Remove a lock from the original array; this allows the user to safely release the array on their own terms
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig2DArray()), ptrSrc
    SafeArrayUnlock ptrSrc
    
End Sub

Public Sub Unalias2DArray_Long(ByRef orig2DArray() As Long, ByRef new2DArray() As Long)
    
    'Wipe the array pointer
    CopyMemory ByVal VarPtrArray(new2DArray), 0&, 4&
    
    'Remove a lock from the original array; this allows the user to safely release the array on their own terms
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig2DArray()), ptrSrc
    SafeArrayUnlock ptrSrc
    
End Sub

'Because we can't use the AddressOf operator inside a class module, timer classes will cheat and AddressOf this
' function instead.  The unique TimerID we specify is actually a handle to the timer instance.
' (Thank you to Karl Peterson for suggesting this excellent trick: http://vb.mvps.org/samples/TimerObj/)
Public Sub StandInTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal cTimer As pdTimer, ByVal dwTime As Long)
    If (Not cTimer Is Nothing) Then cTimer.TimerEventArrived
End Sub

'This beautiful little function comes courtesy of coder Merri:
' http://www.vbforums.com/showthread.php?536960-RESOLVED-how-can-i-see-if-the-object-is-array-or-not
Public Function InControlArray(Ctl As Object) As Boolean
    InControlArray = Not Ctl.Parent.Controls(Ctl.Name) Is Ctl
End Function

'Given a VB array, return an IStream containing the array's data.  We use this frequently in PD to move arrays into
' streams that libraries like GDI+ can work with.  You can also pass a null pointer to generate an empty stream.
' (Note that the returned stream is self-cleaning, so you do not have to worry about manually releasing it.)
Public Function GetStreamFromVBArray(ByVal ptrToFirstArrayElement As Long, ByVal streamLength As Long, Optional ByVal createStreamForNullPointer As Boolean = False) As IUnknown

    On Error GoTo StreamDied
     
    'Null pointers return an empty stream
    If (ptrToFirstArrayElement = 0) Then
        If createStreamForNullPointer Then CreateStreamOnHGlobal 0&, 1&, GetStreamFromVBArray
    Else
        
        'Make sure the length is valid
        If (streamLength <> 0) Then
        
            Dim hGlobalHandle As Long
            hGlobalHandle = GlobalAlloc(GMEM_MOVEABLE, streamLength)
            If (hGlobalHandle <> 0) Then
            
                Dim ptrGlobal As Long
                ptrGlobal = GlobalLock(hGlobalHandle)
                If (ptrGlobal <> 0) Then
                    CopyMemoryStrict ptrGlobal, ptrToFirstArrayElement, streamLength
                    GlobalUnlock ptrGlobal
                    CreateStreamOnHGlobal hGlobalHandle, 1&, GetStreamFromVBArray
                Else
                    PDDebug.LogAction "WARNING!  GetStreamFromVBArray() failed to retrieve a pointer to its hGlobal data!"
                End If
            
            Else
                PDDebug.LogAction "WARNING!  GetStreamFromVBArray() failed to create a valid hGlobal!"
            End If
            
        Else
            PDDebug.LogAction "WARNING!  GetStreamFromVBArray() requires a valid stream length!"
        End If
        
    End If
    
    Exit Function
    
StreamDied:
    PDDebug.LogAction "WARNING!  GetStreamFromVBArray() failed for unknown reasons.  Please investigate!"
    
End Function

'Given an IStream, use its native functionality to write its contents into a VB array.  This should work regardless of
' the IStream's original source (hGlobal, mapped file, whatever).
'
'Note that this function requires you to know the write length in advance.  We could dynamically request a size from
' the IStream itself, but the manual use of DispCallFunc makes this tedious and time-consuming, and PD typically knows
' the size in advance anyway - so please provide that length in advance!
Public Function ReadIStreamIntoVBArray(ByVal ptrSrcStream As Long, ByRef dstArray() As Byte, ByVal dstLength As Long) As Boolean

    On Error GoTo StreamConversionFailed
    
    ReadIStreamIntoVBArray = False
    
    'Null streams are pointless; ignore them completely!
    If (ptrSrcStream <> 0) And (dstLength > 0) Then
        
        ReDim dstArray(0 To dstLength - 1) As Byte
        
        'Prep a manual DispCallFunc invocation
        Dim lRead As Long, varRtn As Variant
        Dim Vars(0 To 3) As Variant, pVars(0 To 3) As Long, pVartypes(0 To 3) As Integer
        pVartypes(0) = vbLong: pVartypes(1) = vbLong: pVartypes(2) = vbLong
        Vars(0) = VarPtr(dstArray(0)): Vars(1) = dstLength: Vars(2) = VarPtr(lRead)
        pVars(0) = VarPtr(Vars(0)): pVars(1) = VarPtr(Vars(1)): pVars(2) = VarPtr(Vars(2))
        
        Const ISTREAM_READ As Long = 12
        Const CC_STDCALL As Long = 4
        
        If (DispCallFunc(ptrSrcStream, ISTREAM_READ, CC_STDCALL, vbLong, 3&, pVartypes(0), pVars(0), varRtn) = 0) Then
            ReadIStreamIntoVBArray = True
        Else
            PDDebug.LogAction "WARNING!  ReadIStreamIntoVBArray() failed to initiate a successful DispCallFunc-based IStream read."
        End If
        
    Else
        PDDebug.LogAction "WARNING!  ReadIStreamIntoVBArray() was passed a null stream pointer and/or size!"
    End If
    
    Exit Function
    
StreamConversionFailed:
    PDDebug.LogAction "WARNING!  ReadIStreamIntoVBArray() failed for unknown reasons.  Please investigate!"
    
End Function

'Check array initialization.  All array types supported.  Thank you to http://stackoverflow.com/questions/183353/how-do-i-determine-if-an-array-is-initialized-in-vb6
Public Function IsArrayInitialized(ByRef arr As Variant) As Boolean
    Dim saAddress As Long
    GetMem4 VarPtr(arr) + 8, saAddress
    GetMem4 saAddress, saAddress
    IsArrayInitialized = (saAddress <> 0)
    If IsArrayInitialized Then IsArrayInitialized = (UBound(arr) >= LBound(arr))
End Function

Public Sub EnableHighResolutionTimers()
    QueryPerformanceFrequency m_TimerFrequency
    If (m_TimerFrequency = 0) Then m_TimerFrequency = 1 Else m_TimerFrequency = 1# / m_TimerFrequency
End Sub

Public Function GetTimerDifference(ByVal startTime As Currency, ByVal stopTime As Currency) As Double
    GetTimerDifference = (stopTime - startTime) * m_TimerFrequency
End Function

Public Function GetTimeDiffAsString(ByVal startTime As Currency, ByVal stopTime As Currency) As String
    Dim tmpDouble As Double
    tmpDouble = (stopTime - startTime) * m_TimerFrequency
    GetTimeDiffAsString = Format$(tmpDouble * 1000#, "0.0") & " ms"
End Function

Public Function GetTimerDifferenceNow(ByVal startTime As Currency) As Double
    Dim tmpTime As Currency
    QueryPerformanceCounter tmpTime
    GetTimerDifferenceNow = (tmpTime - startTime) * m_TimerFrequency
End Function

Public Function GetTimeDiffNowAsString(ByVal startTime As Currency) As String
    Dim tmpTime As Currency:    QueryPerformanceCounter tmpTime
    Dim tmpDouble As Double:    tmpDouble = (tmpTime - startTime) * m_TimerFrequency
    GetTimeDiffNowAsString = Format$(tmpDouble * 1000#, "#,##0.0") & " ms"
End Function

Public Function GetTotalTimeAsString(ByVal netTime As Currency) As String
    Dim tmpDouble As Double
    tmpDouble = netTime * m_TimerFrequency
    GetTotalTimeAsString = Format$(tmpDouble * 1000#, "#,##0.0") & " ms"
End Function

Public Sub GetHighResTime(ByRef dstTime As Currency)
    QueryPerformanceCounter dstTime
End Sub

Public Function GetHighResTimeEx() As Currency
    QueryPerformanceCounter GetHighResTimeEx
End Function

Public Sub GetHighResTimeInMS(ByRef dstTimeInMS As Currency)
    QueryPerformanceCounter dstTimeInMS
    dstTimeInMS = dstTimeInMS * (m_TimerFrequency * 1000@)
End Sub

Public Function GetHighResTimeInMSEx() As Currency
    QueryPerformanceCounter GetHighResTimeInMSEx
    GetHighResTimeInMSEx = GetHighResTimeInMSEx * (m_TimerFrequency * 1000@)
End Function

Public Function MemCmp(ByVal ptr1 As Long, ByVal ptr2 As Long, ByVal bytesToCompare As Long) As Boolean
    Dim bytesEqual As Long
    bytesEqual = RtlCompareMemory(ptr1, ptr2, bytesToCompare)
    MemCmp = (bytesEqual = bytesToCompare)
End Function

'This function mimicks DoEvents, but instead of processing all messages for all windows on all threads (slow! error-prone!),
' it only processes messages for the supplied hWnd.
Public Sub DoEvents_SingleHwnd(ByVal srcHwnd As Long)
    Dim tmpMsg As winMsg
    Do While PeekMessageW(tmpMsg, srcHwnd, 0&, 0&, &H1&)
        TranslateMessage tmpMsg
        DispatchMessageW tmpMsg
    Loop
End Sub

'PD sometimes wants to yield for asynchronous timers (we use pipes in a number of places to communicate with
' 3rd-party libraries), and rather than use DoEvents and risk all kinds of havoc, we simply yield for timer
' events only.
Public Sub DoEventsTimersOnly()
    Dim tmpMsg As winMsg
    Const WM_TIMER As Long = &H113
    Do While PeekMessageA(tmpMsg, 0&, WM_TIMER, WM_TIMER, &H1&)
        TranslateMessage tmpMsg
        DispatchMessageA tmpMsg
    Loop
End Sub

Public Sub PurgeTimerMessagesByID(ByVal nIDEvent As Long)
    Dim tmpMsg As winMsg
    Const WM_TIMER As Long = &H113
    Do While PeekMessageA(tmpMsg, 0&, WM_TIMER, WM_TIMER, &H1&)
        If (tmpMsg.wParam <> nIDEvent) Then
            TranslateMessage tmpMsg
            DispatchMessageA tmpMsg
        End If
    Loop
End Sub

Public Sub PurgeInputMessages(ByVal srcHwnd As Long)
    
    Const QS_MOUSEMOVE = &H2
    Const QS_MOUSEBUTTON = &H4
    Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
    Const QS_KEY = &H1
    Const QS_INPUT = (QS_MOUSE Or QS_KEY)
    Const PM_QS_INPUT = (QS_INPUT * (2 ^ 16))
    
    Dim tmpMsg As winMsg
    Do While PeekMessageW(tmpMsg, srcHwnd, 0&, 0&, &H1& Or PM_QS_INPUT)
    Loop
    
End Sub

Public Function FreeLib(ByVal hLib As Long) As Boolean
    If (hLib = 0) Then
        FreeLib = True
    Else
        FreeLib = (FreeLibrary(hLib) <> 0)
        If FreeLib Then hLib = 0
    End If
End Function

Public Function LoadLib(ByRef libPathAndName As String) As Long
    LoadLib = LoadLibraryW(StrPtr(libPathAndName))
End Function

Public Function SendMsgW(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    SendMsgW = SendMessageW(hWnd, wMsg, wParam, lParam)
End Function

Public Sub SleepAPI(ByVal sleepTimeInMS As Long)
    Sleep sleepTimeInMS
End Sub

'Make certain the length of the source array is an even number (e.g. the UBound is odd) before calling;
' this function does not attempt to verify otherwise
Public Sub SwapEndianness16(ByRef srcData() As Byte)
    Dim i As Long, tmpValue As Long
    For i = 0 To UBound(srcData) Step 2
        tmpValue = srcData(i)
        srcData(i) = srcData(i + 1)
        srcData(i + 1) = tmpValue
    Next i
End Sub

'Make certain the length of the source array is a multiple of 4 before calling;
' this function does not attempt to verify otherwise.
Public Sub SwapEndianness32(ByRef srcData() As Byte)
    Dim i As Long, tmpValue As Long, tmpIndex As Long
    For i = 0 To UBound(srcData) Step 4
        tmpIndex = i + 3
        tmpValue = srcData(tmpIndex)
        srcData(tmpIndex) = srcData(i)
        srcData(i) = tmpValue
        tmpIndex = i + 2
        tmpValue = srcData(tmpIndex)
        srcData(tmpIndex) = srcData(i + 1)
        srcData(i + 1) = tmpValue
    Next i
End Sub

'Safe unsigned addition, regardless of compilation options (e.g. compiling to native code with
' overflow ignored negates the need for this, but we sometimes use it "just in case").
' With thanks to vbforums user Krool for the original implementation: http://www.vbforums.com/showthread.php?698563-CommonControls-(Replacement-of-the-MS-common-controls)
Public Function UnsignedAdd(ByVal baseValue As Long, ByVal amtToAdd As Long) As Long
    UnsignedAdd = ((baseValue Xor SIGN_BIT) + amtToAdd) Xor SIGN_BIT
End Function

'Wrap an array of [type] around an arbitrary pointer.  This is currently used by the PSD parser to
' accelerate some otherwise tedious pointer math.
Public Sub WrapArrayAroundPtr_Int(ByRef dstInts() As Integer, ByRef dstSA1D As SafeArray1D, ByVal srcPtr As Long, ByVal srcLenInBytes As Long)
    With dstSA1D
        .cbElements = 2
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcLenInBytes \ 2
        .pvData = srcPtr
    End With
    CopyMemory ByVal VarPtrArray(dstInts()), VarPtr(dstSA1D), 4&
End Sub

Public Sub UnwrapArrayFromPtr_Int(ByRef dstInts() As Integer)
    PutMem4 VarPtrArray(dstInts), 0&
End Sub

'Subclassing helper functions follow
Public Function StartSubclassing(ByVal hWnd As Long, ByVal Thing As ISubclass, Optional dwRefData As Long) As Boolean
    StartSubclassing = CBool(SetWindowSubclass(hWnd, AddressOf SubclassProc, ObjPtr(Thing), dwRefData))
End Function

Public Function StopSubclassing(ByVal hWnd As Long, ByVal Thing As ISubclass) As Boolean
    StopSubclassing = CBool(RemoveWindowSubclass(hWnd, AddressOf SubclassProc, ObjPtr(Thing)))
End Function

Public Function DefaultSubclassProc(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    DefaultSubclassProc = DefSubclassProc(hWnd, uiMsg, wParam, lParam)
End Function

'As a failsafe against client negligence, this function will automatically remove subclassing when WM_NCDESTROY
' is received.  (PD assumes automatic teardown behavior in a number of places, so *do not* remove the WM_NCDESTROY
' check in this function!)  Note that there is no problem if the caller manually unsubclasses prior to returning;
' the API will simply return FALSE because the hWnd/key pair doesn't exist in the object table.
Public Function SubclassProc(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As ISubclass, ByVal dwRefData As Long) As Long
   SubclassProc = uIdSubclass.WindowMsg(hWnd, uiMsg, wParam, lParam, dwRefData)
   If (uiMsg = WM_NCDESTROY) Then StopSubclassing hWnd, uIdSubclass
End Function

'When a pdInputKeyboard class needs to hook the keyboard, it notifies us via this function.  This function returns the
' hook ID generated by the hook request.
Public Function NotifyPDIKHookNeeded(ByRef srcPDIK As pdInputKeyboard) As Long
    Set m_PDIKRef = srcPDIK
    NotifyPDIKHookNeeded = SetWindowsHookExW(WH_KEYBOARD, AddressOf KeyboardHookProcPDIK, App.hInstance, App.ThreadID)
End Function

'If we mistakenly call a class that has unhooked itself, it can notify us via this function.  We'll release
' our reference if it hasn't already been stolen by another pdIK instance.
Public Sub NotifyPDIKHookNotNeeded(ByVal objPointer As Long)
    If (ObjPtr(m_PDIKRef) = objPointer) Then Set m_PDIKRef = Nothing
End Sub

'Hooked keyboard events happen here; we simply relay the results to last-referenced pdInputKeyboard class
Public Function KeyboardHookProcPDIK(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (Not m_PDIKRef Is Nothing) Then
        KeyboardHookProcPDIK = m_PDIKRef.PDIKKeyboardProc(nCode, wParam, lParam)
    End If
End Function

'Same idea as the three previous functions, but for the main pdAccelerator instance on FormMain
Public Function NotifyAcceleratorHookNeeded(ByRef srcAccelerator As pdAccelerator) As Long
    Set m_AcceleratorRef = srcAccelerator
    NotifyAcceleratorHookNeeded = SetWindowsHookExW(WH_KEYBOARD, AddressOf KeyboardHookProcAccelerator, App.hInstance, App.ThreadID)
End Function

Public Sub NotifyAcceleratorHookNotNeeded(ByVal objPointer As Long)
    If (ObjPtr(m_AcceleratorRef) = objPointer) Then Set m_AcceleratorRef = Nothing
End Sub

Public Function KeyboardHookProcAccelerator(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (Not m_AcceleratorRef Is Nothing) Then
        KeyboardHookProcAccelerator = m_AcceleratorRef.KeyboardHookProcAccelerator(nCode, wParam, lParam)
    End If
End Function

'Same idea as the three previous functions, but for individual pdEditBox instances
Public Function NotifyEditBoxHookNeeded(ByRef srcPDEditBox As pdEditBoxW) As Long
    Set m_EditBoxRef = srcPDEditBox
    NotifyEditBoxHookNeeded = SetWindowsHookExW(WH_KEYBOARD, AddressOf KeyboardHookProcEditBox, App.hInstance, App.ThreadID)
End Function

Public Sub NotifyEditBoxHookNotNeeded(ByVal objPointer As Long)
    If (ObjPtr(m_EditBoxRef) = objPointer) Then Set m_EditBoxRef = Nothing
End Sub

Public Function KeyboardHookProcEditBox(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (Not m_EditBoxRef Is Nothing) Then
        KeyboardHookProcEditBox = m_EditBoxRef.EditBoxKeyboardProc(nCode, wParam, lParam)
    End If
End Function

'If you have any hack-related cleanup that needs to be performed at shutdown time, use this function.
Public Sub ShutdownCleanup()
    Set m_EditBoxRef = Nothing
    Set m_AcceleratorRef = Nothing
    Set m_PDIKRef = Nothing
End Sub
