Attribute VB_Name = "VB_Hacks"
'***************************************************************************
'Misc VB6 Hacks
'Copyright 2016-2016 by Tanner Helland
'Created: 06/January/16
'Last updated: 06/January/16
'Last update: started moving VB6 hacks to a dedicated home; this should make various other modules more readable,
'             since VB6 hacks often involve obscure trickery whose purpose isn't always obvious to non-VB6 coders.
'
'PhotoDemon relies on a lot of "not officially sanctioned" VB6 behavior to enable various optimizations and C-style
' code techniques. If a function's primary purpose is a VB6-specific workaround, I prefer to move it here, so I
' don't clutter up purposeful modules with obscure, VB-specific hackery.
'
'Note that some code here may seem redundant (e.g. identical functions suffixed by data type, instead of declared
' "As Any") but that's by design - e.g. to improve safety since these techniques are almost always crash-prone if
' used incorrectly or imprecisely.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Sub SafeArrayLock Lib "oleaut32" (ByVal ptrToSA As Long)
Private Declare Sub SafeArrayUnlock Lib "oleaut32" (ByVal ptrToSA As Long)
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal newValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByVal Addr As Long, ByRef dstValue As Long) As Long
Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDestPtr As Long, ByVal lpvSourcePtr As Long, ByVal cbCopy As Long)
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, ByRef hGlobal As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED As Long = &H0&
Private Const GMEM_MOVEABLE As Long = &H2&

'A system info class is used to retrieve ThunderMain's hWnd, if required
Private m_SysInfo As pdSystemInfo

'Point an internal 2D array at some other 2D array.  Any arrays aliased this way must be freed via Unalias2DArray,
' or VB will crash.
Public Sub Alias2DArray_Byte(ByRef orig2DArray() As Byte, ByRef new2DArray() As Byte, ByRef newArraySA As SAFEARRAY2D)
    
    'Retrieve a copy of the original 2D array's SafeArray struct
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig2DArray()), ptrSrc
    CopyMemory ByVal VarPtr(newArraySA), ByVal ptrSrc, LenB(newArraySA)
    
    'newArraySA now contains the full SafeArray of the original array.  Copy this over our current array.
    CopyMemory ByVal VarPtrArray(new2DArray()), VarPtr(newArraySA), 4&
    
    'Add a lock to the original array, to prevent potential crashes from unknowing users.  (Thanks to @Kroc for this tip.)
    SafeArrayLock ptrSrc
    
End Sub

Public Sub Alias2DArray_Integer(ByRef orig2DArray() As Integer, ByRef new2DArray() As Integer, ByRef newArraySA As SAFEARRAY2D)
    
    'Retrieve a copy of the original 2D array's SafeArray struct
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig2DArray()), ptrSrc
    CopyMemory ByVal VarPtr(newArraySA), ByVal ptrSrc, LenB(newArraySA)
    
    'newArraySA now contains the full SafeArray of the original array.  Copy this over our current array.
    CopyMemory ByVal VarPtrArray(new2DArray()), VarPtr(newArraySA), 4&
    
    'Add a lock to the original array, to prevent potential crashes from unknowing users.  (Thanks to @Kroc for this tip.)
    SafeArrayLock ptrSrc
    
End Sub

Public Sub Alias2DArray_Long(ByRef orig2DArray() As Long, ByRef new2DArray() As Long, ByRef newArraySA As SAFEARRAY2D)
    
    'Retrieve a copy of the original 2D array's SafeArray struct
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig2DArray()), ptrSrc
    CopyMemory ByVal VarPtr(newArraySA), ByVal ptrSrc, LenB(newArraySA)
    
    'newArraySA now contains the full SafeArray of the original array.  Copy this over our current array.
    CopyMemory ByVal VarPtrArray(new2DArray()), VarPtr(newArraySA), 4&
    
    'Add a lock to the original array, to prevent potential crashes from unknowing users.  (Thanks to @Kroc for this tip.)
    SafeArrayLock ptrSrc
    
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

'Retrieve a handle to ThunderMain.  Works in the IDE as well, but the usual caveats apply.
Public Function GetThunderMainHWnd() As Long
    If m_SysInfo Is Nothing Then Set m_SysInfo = New pdSystemInfo
    GetThunderMainHWnd = m_SysInfo.GetPhotoDemonMasterHWnd()
End Function

'Because we can't use the AddressOf operator inside a class module, timer classes will cheat and AddressOf this
' function instead.  The unique TimerID we specify is actually a handle to the timer instance.
' (Thank you to Karl Peterson for suggesting this excellent trick: http://vb.mvps.org/samples/TimerObj/)
Public Sub StandInTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal cTimer As pdTimer, ByVal dwTime As Long)
    cTimer.TimerEventArrived
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
    If ptrToFirstArrayElement = 0 Then
        If createStreamForNullPointer Then CreateStreamOnHGlobal 0&, 1&, GetStreamFromVBArray
    Else
        
        'Make sure the length is valid
        If streamLength <> 0 Then
        
            Dim hGlobalHandle As Long
            hGlobalHandle = GlobalAlloc(GMEM_MOVEABLE, streamLength)
            If hGlobalHandle <> 0 Then
            
                Dim ptrGlobal As Long
                ptrGlobal = GlobalLock(hGlobalHandle)
                If ptrGlobal <> 0 Then
                    CopyMemoryStrict ptrGlobal, ptrToFirstArrayElement, streamLength
                    GlobalUnlock ptrGlobal
                    CreateStreamOnHGlobal hGlobalHandle, 1&, GetStreamFromVBArray
                Else
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "WARNING!  GetStreamFromVBArray() failed to retrieve a pointer to its hGlobal data!"
                    #End If
                End If
            
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  GetStreamFromVBArray() failed to create a valid hGlobal!"
                #End If
            End If
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  GetStreamFromVBArray() requires a valid stream length!"
            #End If
        End If
        
    End If
    
    Exit Function
    
StreamDied:
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  GetStreamFromVBArray() failed for unknown reasons.  Please investigate!"
    #End If
End Function

'Given an IStream, return its contents as a VB array.  (This implementation is pretty darn similar to its partner function, above.)
Public Function GetVBArrayFromStream(ByVal ptrSrcStream As Long, ByRef dstArray() As Byte) As Boolean

    On Error GoTo StreamConversionFailed
    
    GetVBArrayFromStream = False
    
    'Null streams are pointless; ignore them completely!
    If (ptrSrcStream <> 0) Then
        
        'Get an hGlobal that points to the stream's data
        Dim hGlobalHandle As Long
        If GetHGlobalFromStream(ptrSrcStream, hGlobalHandle) = 0 Then
            
            'Make sure the stream contains at least one usable byte
            Dim streamSize As Long
            streamSize = GlobalSize(hGlobalHandle)
            If streamSize > 0 Then
                
                'Get a raw pointer to the data
                Dim ptrGlobal As Long
                ptrGlobal = GlobalLock(hGlobalHandle)
                If ptrGlobal <> 0 Then
                    
                    'Copy the data, free the hGlobal, then bail
                    ReDim arrayBytes(0 To streamSize - 1) As Byte
                    CopyMemoryStrict VarPtr(arrayBytes(0)), ptrGlobal, streamSize
                    
                    GlobalUnlock hGlobalHandle
                    GetVBArrayFromStream = True
                    
                End If
                
            End If
            
        End If
        
    End If
    
    Exit Function
    
StreamConversionFailed:
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  GetVBArrayFromStream() failed for unknown reasons.  Please investigate!"
    #End If
End Function
