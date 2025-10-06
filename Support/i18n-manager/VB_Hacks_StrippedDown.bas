Attribute VB_Name = "VBHacks"
'***************************************************************************
'Misc VB6 Hacks (stripped down version of the PD original)
'Copyright 2016-2022 by Tanner Helland
'Created: 06/January/16
'Last updated: 27/May/22
'Last update: move various incarnations of Get/SetBitFlag functions here, and build a fixed-size flag table for perf reasons
'
'PhotoDemon relies on a lot of "not officially sanctioned" VB6 behavior to enable various optimizations
' and C-style code techniques. If a function's primary purpose is a VB6-specific workaround, I prefer to
' move it here instead of cluttering purposeful modules with obscure, VB-specific hackery.
'
'Note that some code here may seem redundant (e.g. identical functions suffixed by data type, instead of
' declared "As Any") but that's by design - e.g. to improve safety since these techniques are almost
' always crash-prone if used incorrectly or imprecisely.
'
'A number of the techniques in this module were devised with help from Karl E. Peterson's work at
' http://vb.mvps.org/ - thank you, Karl!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Some of these are copied from elsewhere in PD to allow this module version to stand alone
Public Const LONG_MAX As Long = 2147483647
Public Const DOUBLE_MAX As Double = 1.79769313486231E+308
Public Const SINGLE_MAX As Single = 3.402823E+38!

Public Type RGBQuad
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

'SafeArray types for pointing VB arrays at arbitrary memory locations (in our case, bitmap data)
Public Type SafeArrayBound
    cElements As Long
    lBound   As Long
End Type

Public Type SafeArray2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SafeArrayBound
End Type

Public Type SafeArray1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    cElements As Long
    lBound   As Long
End Type

Private Type DROPFILES
    pFiles As Long
    ptX As Long
    ptY As Long
    fNC As Long
    fWide As Long
End Type

Private Type FORMATETC
    cfFormat As Long
    pDVTARGETDEVICE As Long
    dwAspect As Long
    lIndex As Long
    TYMED As Long
End Type

Private Type STGMEDIUM
    TYMED As Long
    Data As Long
    pUnkForRelease As Long
End Type

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
'Public Declare Sub GetMem8 Lib "msvbvm60" (ByVal ptrSrc As Long, ByRef dstCurrency As Currency)
Public Declare Sub GetMem8_Ptr Lib "msvbvm60" Alias "GetMem8" (ByVal ptrSrc As Long, ByVal ptrDst8 As Long)
Public Declare Sub PutMem1 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Byte)
Public Declare Sub PutMem2 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Integer)
Public Declare Sub PutMem4 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Long)
'Public Declare Sub PutMem8 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Currency)

Public Const WM_NCDESTROY As Long = &H82&

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
Private Declare Function lstrlenW Lib "kernel32" (ByVal ptrToFirstChar As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function RtlCompareMemory Lib "ntdll" (ByVal ptrSource1 As Long, ByVal ptrSource2 As Long, ByVal Length As Long) As Long

Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As VbVarType, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
Private Declare Sub ReleaseStgMedium Lib "ole32" (ByVal ptrToStgMedium As Long)

Private Declare Function DispatchMessageA Lib "user32" (ByRef lpMsg As winMsg) As Long
Private Declare Function DispatchMessageW Lib "user32" (ByRef lpMsg As winMsg) As Long
Private Declare Function PeekMessageA Lib "user32" (ByRef lpMsg As winMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function PeekMessageW Lib "user32" (ByRef lpMsg As winMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (ByRef lpMsg As winMsg) As Long

Private Const CC_STDCALL As Long = 4
Private Const DVASPECT_CONTENT As Long = 1
Private Const GMEM_MOVEABLE As Long = &H2&
Private Const IDataObjVTable_GetData As Long = 12   '12 is an offset to the GetData function (e.g. the 4th VTable entry)
Private Const SIGN_BIT As Long = &H80000000
Private Const TYMED_HGLOBAL As Long = 1
Private Const WH_KEYBOARD As Long = 2

'Higher-performance timing functions are also handled by this class.  Note that you *must* initialize the timer engine
' before requesting any time values, or crashes will occurs because the frequency timer is 0.
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
Private m_TimerFrequency As Currency

Private m_BitFlags(0 To 31) As Long, m_BitFlagsReady As Boolean

'"Wrap" an arbitrary VB array at some other arbitrary array.  The new array must *NOT* be initialized
' or its memory will leak.
'
'The destination array inherits all properties of the source array.  This is useful for very specific
' performance scenarios - i.e. instead of passing arrays to functions inside a tight loop, a single
' alias can be performed before the loop, and when the loop completes, unalias the array.)
'
'Any arrays aliased this way must be freed via UnaliasArbitraryArray or VB will crash (double-free).
Public Sub AliasArbitraryArray(ByVal source_VarPtrArray As Long, ByVal destination_VarPtrArray As Long)
    GetMem4 source_VarPtrArray, ByVal destination_VarPtrArray
End Sub

'Point an internal 2D array at some other 2D array.  Any arrays aliased this way must be freed via Unalias2DArray,
' or VB will crash.
Public Sub Alias2DArray_Byte(ByRef orig2DArray() As Byte, ByRef new2DArray() As Byte)
    GetMem4 VarPtrArray(orig2DArray()), ByVal VarPtrArray(new2DArray())
End Sub

Public Sub Alias2DArray_Integer(ByRef orig2DArray() As Integer, ByRef new2DArray() As Integer)
    GetMem4 VarPtrArray(orig2DArray()), ByVal VarPtrArray(new2DArray())
End Sub

'Counterpart to AliasArbitraryArray(), above.  Do NOT call this function on arrays that
' were not originally processed by AliasArbitraryArray().  (Note that the source pointer
' is not used - the function is left this way, by design, so it's symmetrical with the
' original AliasArbitraryArray call!)
Public Sub UnaliasArbitraryArray(ByVal source_VarPtrArray As Long, ByVal destination_VarPtrArray As Long)
    PutMem4 destination_VarPtrArray, 0&
End Sub

'Counterparts to Alias2DArray_ functions, above.  Do NOT call this function on arrays that were not originally
' processed by an Alias2DArray_ function.
Public Sub Unalias2DArray_Byte(ByRef orig2DArray() As Byte, ByRef new2DArray() As Byte)
    PutMem4 VarPtrArray(new2DArray), 0&
End Sub

Public Sub Unalias2DArray_Integer(ByRef orig2DArray() As Integer, ByRef new2DArray() As Integer)
    PutMem4 VarPtrArray(new2DArray), 0&
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
    GetTimeDiffNowAsString = Format$(tmpDouble * 1000#, "#,#0.0") & " ms"
End Function

Public Function GetTotalTimeAsString(ByVal netTime As Currency) As String
    Dim tmpDouble As Double
    tmpDouble = netTime * m_TimerFrequency
    GetTotalTimeAsString = Format$(tmpDouble * 1000#, "#,#0.0") & " ms"
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
Public Sub DoEvents_SingleHwnd(ByVal srcHWnd As Long)
    Dim tmpMsg As winMsg
    Do While PeekMessageW(tmpMsg, srcHWnd, 0&, 0&, &H1&)
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

'PD sometimes wants to yield for paint events (e.g. updating a status bar) without risking
' reentrancy from input events.
Public Sub DoEvents_PaintOnly(ByVal targetHWnd As Long, Optional ByVal alsoPurgeInputEvents As Boolean = True)
    
    Dim tmpMsg As winMsg
    Const QS_PAINT As Long = &H20&
    Const PM_QS_PAINT As Long = (QS_PAINT * (2& ^ 16&))
    Do While PeekMessageA(tmpMsg, 0&, 0&, 0&, &H1& Or PM_QS_PAINT)
        TranslateMessage tmpMsg
        DispatchMessageA tmpMsg
    Loop
    
    If alsoPurgeInputEvents Then VBHacks.PurgeInputMessages targetHWnd
    
End Sub

Public Sub PurgeTimerMessagesByID(ByVal nIDEvent As Long)
    Dim tmpMsg As winMsg
    Const WM_TIMER As Long = &H113&
    Do While PeekMessageA(tmpMsg, 0&, WM_TIMER, WM_TIMER, &H1&)
        If (tmpMsg.wParam <> nIDEvent) Then
            TranslateMessage tmpMsg
            DispatchMessageA tmpMsg
        End If
    Loop
End Sub

Public Sub PurgeInputMessages(ByVal srcHWnd As Long)
    
    Const QS_MOUSEMOVE = &H2
    Const QS_MOUSEBUTTON = &H4
    Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
    Const QS_KEY = &H1
    Const QS_INPUT = (QS_MOUSE Or QS_KEY)
    Const PM_QS_INPUT = (QS_INPUT * (2& ^ 16&))
    
    Dim tmpMsg As winMsg
    Do While PeekMessageW(tmpMsg, srcHWnd, 0&, 0&, &H1& Or PM_QS_INPUT)
    Loop
    
End Sub

Private Sub BuildBitFlagTable()
    Dim i As Long
    For i = 0 To 30
        m_BitFlags(i) = 2 ^ i
    Next i
    m_BitFlags(31) = &H80000000
    m_BitFlagsReady = True
End Sub

'Retrieve an arbitrary 1-bit position [0-31] in a Long-type value.
' Position 0 is the LEAST-SIGNIFICANT BIT, and Position 31 is the SIGN BIT for a standard VB Long.
'
' Inputs:
'  1) position of the flag, which must be in the range [0, 31]
'  2) The Long from which you want the bit retrieved
Public Function GetBitFlag_Long(ByVal flagPosition As Long, ByVal srcLong As Long) As Boolean
    If (flagPosition >= 0) And (flagPosition <= 31) Then
        If (Not m_BitFlagsReady) Then BuildBitFlagTable
        GetBitFlag_Long = (Int(srcLong And m_BitFlags(flagPosition)) <> 0)
    End If
End Function

'Set an arbitrary 1-bit position [0-31] in a Long-type value to either 1 or 0.
' Position 0 is the LEAST-SIGNIFICANT BIT, and Position 31 is the SIGN BIT for a standard VB Long.
' Inputs:
'  1) position of the flag, which must be in the range [0, 31]
'  2) value of the flag, TRUE for 1, FALSE for 0
'  3) The Long-type value where you want the flag placed
Public Sub SetBitFlag_Long(ByVal flagPosition As Long, ByVal flagValue As Boolean, ByRef dstLong As Long)
    If (flagPosition >= 0) And (flagPosition <= 31) Then
        If (Not m_BitFlagsReady) Then BuildBitFlagTable
        If flagValue Then
            dstLong = dstLong Or m_BitFlags(flagPosition)
        Else
            dstLong = dstLong And Not m_BitFlags(flagPosition)
        End If
    End If
End Sub

Public Function FreeLib(ByRef hLib As Long) As Boolean
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

'Shuffle a source array (should be float data crammed or aliased into a byte array) into a new array,
' but perform the equivalent of a planar-to-chunky conversion.  This is useful for floating-point data
' in particular, as it can greatly improve compression ratios by placing all 1st bytes, 2nd bytes, etc
' next to each to eachother as there is a far higher probability of repeatable sequences.
'
'Note that THE CALLER needs to ensure that BOTH the source and destination points are to valid,
' correctly allocated arrays (type doesn't matter - we'll allocate at byte level as necessary).
' and that the number of LONGS (not bytes) to shuffle is accurate.
Public Sub ShuffleBytes_4(ByVal srcPtr As Long, ByVal dstPtr As Long, ByVal numLongsToShuffle As Long)
    
    'Alias an RGBQuad around the source array.  Note that this is an arbitrary struct;
    ' we simply need a VB-convenient way to access individual bytes, and PD already uses
    ' this struct everywhere.
    Dim srcQuads() As RGBQuad, sfArrayQuad As SafeArray1D
    With sfArrayQuad
        .cbElements = 4
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = numLongsToShuffle
        .pvData = srcPtr
    End With
    PutMem4 VarPtrArray(srcQuads()), VarPtr(sfArrayQuad)
    
    'Next, alias an arbitrary byte array around the destination
    Dim dstBytes() As Byte, sfArrayByte As SafeArray1D
    With sfArrayByte
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = numLongsToShuffle * 4
        .pvData = dstPtr
    End With
    PutMem4 VarPtrArray(dstBytes()), VarPtr(sfArrayByte)
    
    'Shuffle!
    Dim i As Long
    For i = 0 To numLongsToShuffle - 1
        dstBytes(i) = srcQuads(i).Alpha
        dstBytes(numLongsToShuffle + i) = srcQuads(i).Red
        dstBytes(numLongsToShuffle * 2 + i) = srcQuads(i).Green
        dstBytes(numLongsToShuffle * 3 + i) = srcQuads(i).Blue
    Next i
    
    PutMem4 VarPtrArray(srcQuads()), 0&
    PutMem4 VarPtrArray(dstBytes()), 0&

End Sub

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

Public Sub SwapEndianness32(ByRef srcData() As Byte)
    Dim i As Long, tmpValue As Long
    For i = 0 To UBound(srcData) Step 4
        tmpValue = srcData(i)
        srcData(i) = srcData(i + 3)
        srcData(i + 3) = tmpValue
        tmpValue = srcData(i + 1)
        srcData(i + 1) = srcData(i + 2)
        srcData(i + 2) = tmpValue
    Next i
End Sub

Public Sub SwapEndianness64(ByRef srcData() As Byte)
    Dim i As Long, tmpValue As Long
    For i = 0 To UBound(srcData) Step 8
        tmpValue = srcData(i)
        srcData(i) = srcData(i + 7)
        srcData(i + 7) = tmpValue
        tmpValue = srcData(i + 1)
        srcData(i + 1) = srcData(i + 6)
        srcData(i + 6) = tmpValue
        tmpValue = srcData(i + 2)
        srcData(i + 2) = srcData(i + 5)
        srcData(i + 5) = tmpValue
        tmpValue = srcData(i + 3)
        srcData(i + 3) = srcData(i + 4)
        srcData(i + 4) = tmpValue
    Next i
End Sub

'Un-shuffle data that was previously shuffled by the ShuffleBytes_4 function.
' (Look there for comments.)
'As always, make sure your source and destination pointers are accurate!
Public Sub UnshuffleBytes_4(ByVal srcPtr As Long, ByVal dstPtr As Long, ByVal numLongsToShuffle As Long)
    
    'Basically, do everything in reverse from ShuffleBytes_4.
    Dim dstQuads() As RGBQuad, sfArrayQuad As SafeArray1D
    With sfArrayQuad
        .cbElements = 4
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = numLongsToShuffle
        .pvData = dstPtr
    End With
    PutMem4 VarPtrArray(dstQuads()), VarPtr(sfArrayQuad)
    
    Dim srcBytes() As Byte, sfArrayByte As SafeArray1D
    With sfArrayByte
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = numLongsToShuffle * 4
        .pvData = srcPtr
    End With
    PutMem4 VarPtrArray(srcBytes()), VarPtr(sfArrayByte)
    
    Dim i As Long
    For i = 0 To numLongsToShuffle - 1
        With dstQuads(i)
            .Alpha = srcBytes(i)
            .Red = srcBytes(numLongsToShuffle + i)
            .Green = srcBytes(numLongsToShuffle * 2 + i)
            .Blue = srcBytes(numLongsToShuffle * 3 + i)
        End With
    Next i
    
    PutMem4 VarPtrArray(dstQuads()), 0&
    PutMem4 VarPtrArray(srcBytes()), 0&

End Sub

'Make certain the length of the source array is a multiple of 4 before calling;
' this function does not attempt to verify otherwise.
' (NOTE: this function is currently unused in PD.)
'Public Sub SwapEndianness32(ByRef srcData() As Byte)
'    Dim i As Long, tmpValue As Long, tmpIndex As Long
'    For i = 0 To UBound(srcData) Step 4
'        tmpIndex = i + 3
'        tmpValue = srcData(tmpIndex)
'        srcData(tmpIndex) = srcData(i)
'        srcData(i) = tmpValue
'        tmpIndex = i + 2
'        tmpValue = srcData(tmpIndex)
'        srcData(tmpIndex) = srcData(i + 1)
'        srcData(i + 1) = tmpValue
'    Next i
'End Sub

'Safe unsigned addition, regardless of compilation options (e.g. compiling to native code with
' overflow ignored negates the need for this, but we sometimes use it "just in case").
' With thanks to vbforums user Krool for the original implementation: http://www.vbforums.com/showthread.php?698563-CommonControls-(Replacement-of-the-MS-common-controls)
Public Function UnsignedAdd(ByVal baseValue As Long, ByVal amtToAdd As Long) As Long
    UnsignedAdd = ((baseValue Xor SIGN_BIT) + amtToAdd) Xor SIGN_BIT
End Function

'Wrap an array of [type] around an arbitrary pointer.
Public Sub WrapArrayAroundPtr_Byte(ByRef dstBytes() As Byte, ByRef dstSA1D As SafeArray1D, ByVal srcPtr As Long, ByVal srcLenInBytes As Long)
    With dstSA1D
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcLenInBytes
        .pvData = srcPtr
    End With
    PutMem4 VarPtrArray(dstBytes()), VarPtr(dstSA1D)
End Sub

Public Sub UnwrapArrayFromPtr_Byte(ByRef dstBytes() As Byte)
    PutMem4 VarPtrArray(dstBytes), 0&
End Sub

Public Sub WrapArrayAroundPtr_Int(ByRef dstInts() As Integer, ByRef dstSA1D As SafeArray1D, ByVal srcPtr As Long, ByVal srcLenInBytes As Long)
    With dstSA1D
        .cbElements = 2
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcLenInBytes \ 2
        .pvData = srcPtr
    End With
    PutMem4 VarPtrArray(dstInts()), VarPtr(dstSA1D)
End Sub

Public Sub UnwrapArrayFromPtr_Int(ByRef dstInts() As Integer)
    PutMem4 VarPtrArray(dstInts), 0&
End Sub

Public Sub WrapArrayAroundPtr_Long(ByRef dstLongs() As Long, ByRef dstSA1D As SafeArray1D, ByVal srcPtr As Long, ByVal srcLenInBytes As Long)
    With dstSA1D
        .cbElements = 4
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcLenInBytes \ 4
        .pvData = srcPtr
    End With
    PutMem4 VarPtrArray(dstLongs()), VarPtr(dstSA1D)
End Sub

Public Sub UnwrapArrayFromPtr_Long(ByRef dstLongs() As Long)
    PutMem4 VarPtrArray(dstLongs), 0&
End Sub

Public Sub WrapArrayAroundPtr_Float(ByRef dstFloats() As Single, ByRef dstSA1D As SafeArray1D, ByVal srcPtr As Long, ByVal srcLenInBytes As Long)
    With dstSA1D
        .cbElements = 4
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcLenInBytes \ 4
        .pvData = srcPtr
    End With
    PutMem4 VarPtrArray(dstFloats()), VarPtr(dstSA1D)
End Sub

Public Sub UnwrapArrayFromPtr_Float(ByRef dstFloats() As Single)
    PutMem4 VarPtrArray(dstFloats), 0&
End Sub

'PD can use standard VB6 OLEDragDrop/Over events, but we need to perform some hackery to prevent
' VB from downsampling paths and filenames from 16-bits per char to 8.  Use this function to
' do the hackery for you.
'
'Note that - by default - this function validates the existence of each file before returning it.
' This doesn't spare you from also needing to validate the list (as file existence can obviously
' change over time!) but it does ensure that this function will never return a non-existent file.
'
'The Unicode-friendly drag/drop interpreter code comes courtesy LaVolpe, c/o
' http://cyberactivex.com/UnicodeTutorialVb.htm#Filenames_via_DragDrop_or_Paste (retrieved on 27/Feb/2016).
' IMPORTANT NOTE: this function has been modified for use inside PD.  If using it in your own project,
' I strongly recommend downloading the original version, as it includes additional helper code
' and explanations.
'
'Returns: TRUE if the drag/drop object contains a list of files, *and* at least one valid path exists.
Public Function GetDragDropFileListW(ByRef OLEDragDrop_DataObject As DataObject, ByRef dstStringStack As pdStringStack) As Boolean

    On Error GoTo DragDropFilesFailed
    
    GetDragDropFileListW = False
    
    'Note that I always validate data objects before passing them, but better safe than sorry
    If (OLEDragDrop_DataObject Is Nothing) Then Exit Function
    If (Not OLEDragDrop_DataObject.GetFormat(vbCFFiles)) Then Exit Function
    
    'This function basically handles the task of hacking around the DataObject's VTable,
    ' and manually retrieving pointers to the original, unmodified Unicode filenames in
    ' the Data object.
    Dim fmtEtc As FORMATETC
    With fmtEtc
        .cfFormat = vbCFFiles
        .lIndex = -1                  ' -1 means "we want everything"
        .TYMED = TYMED_HGLOBAL        ' TYMED_HGLOBAL means we want to use "hGlobal" as the transfer medium
        .dwAspect = DVASPECT_CONTENT  ' dwAspect is used to request extra metadata (like an icon representation) - we want the actual data
    End With
    
    'The IDataObject pointer appears 16 bytes past VBs DataObject
    Dim IID_IDataObject As Long
    CopyMemoryStrict VarPtr(IID_IDataObject), ObjPtr(OLEDragDrop_DataObject) + 16&, 4&
    
    'The objPtr of the IDataObject interface also tells us where the interface's VTable begins.
    ' Since we know the VTable address and we know which function index we want, we can call it
    ' directly using DispCallFunc. (You could also do this using a TLB, obviously.)
    
    ' In particular, we want the GetData function which is #4 in the VTable, per
    ' http://msdn2.microsoft.com/en-us/library/ms688421.aspx
    
    'Next, we need to populate the input values required by the OLE API
    ' (http://msdn2.microsoft.com/en-us/library/ms221473.aspx)
    Dim pVartypes(0 To 1) As Integer, Vars(0 To 1) As Variant, pVars(0 To 1) As Long
    pVartypes(0) = vbLong: Vars(0) = VarPtr(fmtEtc): pVars(0) = VarPtr(Vars(0))
    
    Dim pMedium As STGMEDIUM
    pVartypes(1) = vbLong: Vars(1) = VarPtr(pMedium): pVars(1) = VarPtr(Vars(1))
    
    'Manually invoke the desired interface
    Dim varRtn As Variant
    If (DispCallFunc(IID_IDataObject, IDataObjVTable_GetData, CC_STDCALL, vbLong, 2, pVartypes(0), pVars(0), varRtn) = 0) Then
        
        'Make sure we received a non-null hGlobal pointer (nothing needs to be freed at this point, FYI)
        If (pMedium.Data = 0) Then Exit Function
        
        'Remember that this data object doesn't point directly at the files themselves,
        ' but to a 20-byte DROPFILES structure
        Dim hDrop As Long
        CopyMemoryStrict VarPtr(hDrop), pMedium.Data, 4&
        
        'Technically we should never retrieve have received a null-pointer, but again,
        ' better safe than sorry.
        If (hDrop <> 0) Then
            
            'Convert the hDrop into a usable DROPFILES struct
            Dim dFiles As DROPFILES
            CopyMemoryStrict VarPtr(dFiles), hDrop, 20&
            
            'To my knowledge, we'll never get an ANSI listing from this (convoluted) approach,
            ' at least not on XP or later - but check just in case.
            If (dFiles.fWide <> 0) Then
                
                'Use the pFiles member to track the filename's offsets
                dFiles.pFiles = dFiles.pFiles + hDrop
                
                'Prep a string stack.  Some of the passed files may not be valid, so we want to
                ' validate them in turn before sending them off
                Set dstStringStack = New pdStringStack
                
                Dim lLen As Long, tmpString As String
                
                'Now we're going to iterate through each of the original files, and copy the C-style
                ' strings into BSTRs
                Dim i As Long
                For i = 0 To OLEDragDrop_DataObject.Files.Count - 1
                
                    'Retrieve this filename's length, prep a buffer, then copy over the w-char bytes
                    lLen = lstrlenW(dFiles.pFiles) * 2
                    
                    If (lLen <> 0) Then
                        tmpString = String$(lLen \ 2, 0&)
                        CopyMemoryStrict StrPtr(tmpString), dFiles.pFiles, lLen
                    
                        'Any valid files get added to the collection.
                        If Files.FileExists(tmpString) Then dstStringStack.AddString tmpString
                        
                    End If
                        
                    'Manually move the pointer to the next file, and note that we add two extra
                    ' bytes because of the double-null delimiter between filenames
                    dFiles.pFiles = dFiles.pFiles + lLen + 2
                    
                Next i
                
                'We've got what we need from the hGlobal pointer, so go ahead and free it
                ' (if we're responsible for the drop - note that a NULL value for pUnkForRelease
                ' means that we must free the data; non-NULL means the caller will do it.)
                If (pMedium.pUnkForRelease = 0) Then ReleaseStgMedium VarPtr(pMedium)
                
                'The cFiles string stack now contains all the valid filenames from the dropped list.
                GetDragDropFileListW = (dstStringStack.GetNumOfStrings > 0)
                
            '/End failsafe check for wchar strings
            End If
            
        '/End non-zero hDrop
        End If

    '/End DispCallFunc success
    End If
    
    Exit Function
    
DragDropFilesFailed:
    PDDebug.LogAction "WARNING!  VBHacks.GetDragDropFileListW() experienced error #" & Err.Number & ": " & Err.Description
    
End Function

'Return the minimum of three integer values.  (PD commonly uses this for colors, hence the RGB notation.)
Public Function Min3Int(ByVal rR As Long, ByVal rG As Long, ByVal rB As Long) As Long
    If (rR < rG) Then
        If (rR < rB) Then Min3Int = rR Else Min3Int = rB
    Else
        If (rB < rG) Then Min3Int = rB Else Min3Int = rG
    End If
End Function
