VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Clipboard Interface
'Copyright 2013-2015 by Tanner Helland
'Created: 06/November/15 (formally split off from a heavily modified cCustomClipboard by Steve McMahon)
'Last updated: 15/November/15
'Last update: rewrite the damn thing from scratch
'
'For many years, PD used vbAccelerator's "cCustomClipboard" class to simplify Windows clipboard interactions:
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Clipboard/Customising_Clipboard_Use/article.asp
'
'Starting in May 2013 (https://github.com/tannerhelland/PhotoDemon/commit/c61951048c8a411e5e100cba9010dad8d81fc096),
' I rewrote much of the class to add features, and solve some glaring stability issues.  Unfortunately, over time I
' have come to realize that the class is extremely bug-prone.  API returns are not always checked correctly, and when
' they are, they are handled out-of-order in various If/Then branches, so failure states result in things like memory
' objects not being released.  Worse still, the class mixes a bunch of different coding styles together (e.g. 1-based
' arrays in some places 0-based in others), which makes interop prone to all kinds of failures.  The class is also
' not Unicode-compatible, it uses a bunch of deprecated constants for 16-bit support (!!!), and it's missing support
' for new system-defined formats that have been around since XP (like DIBv5).
'
'Because clipboard interactions are such a crucial part of a graphics program, a ground-up rewrite seemed long overdue.
' Hence this new class.  It uses a completely different methodology for enumerating available clipboard formats,
' and many small bugs have been ironed out.  The coding style also fits with the rest of PD (so everything is 0-based),
' event names are now representative of what they actually do, many dead and/or broken functions are gone, and a bunch
' of previously missing functionality has been plugged in.
'
'Despite my bitching, I am still very grateful to Steve McMahon for his original implementation, which served PD well
' for several years.  It's still a fine reference for beginners, and you can find the original article and code here
' (good as of November '15):
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Clipboard/Customising_Clipboard_Use/article.asp
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'When delayed rendering is active, this class raises three potential events
Public Event ClipboardRenderFormat(ByVal FormatID As PredefinedClipboardFormatConstants)
Public Event ClipboardRenderAllFormats()
Public Event ClipboardDestroyStashedData()

'Clipboard functions:
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardFormatNameW Lib "user32" (ByVal wFormat As Long, ByVal lpStringW As Long, ByVal nMaxCount As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function RegisterClipboardFormatW Lib "user32" (ByVal lpStringW As Long) As Long
Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function CountClipboardFormats Lib "user32" () As Long
Private Declare Function GetClipboardOwner Lib "user32" () As Long

'Global memory functions (all clipboard data is placed on/taken from the global heap):
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_MOVEABLE As Long = &H2

'Basic memory copying:
Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDestPtr As Long, ByVal lpvSourcePtr As Long, ByVal cbCopy As Long)

'Current ID and name values currently on the clipboard.  Update these values via UpdateClipboardLists().
Private m_ClipboardIDs() As Long
Private m_ClipboardNames() As String
Private m_NumOfClipboardFormats As Long

'Tracking variables; certain clipboard APIs require an owner hWnd, so we cache one here when the clipboard is first opened.
Private m_IsClipboardOpen As Boolean
Private m_OwnerHwnd As Long

'Windows defines some clipboard formats for us.  These never change, regardless of OS version.
Public Enum PredefinedClipboardFormatConstants
    CF_MIN = 0
    CF_TEXT = 1
    CF_BITMAP = 2
    CF_METAFILEPICT = 3
    CF_SYLK = 4
    CF_DIF = 5
    CF_TIFF = 6
    CF_OEMTEXT = 7
    CF_DIB = 8
    CF_PALETTE = 9
    CF_PENDATA = 10
    CF_RIFF = 11
    CF_WAVE = 12
    CF_UNICODETEXT = 13
    CF_ENHMETAFILE = 14
    CF_HDROP = 15
    CF_LOCALE = 16
    CF_DIBV5 = 17
    CF_MAX = 18
    
    CF_OWNERDISPLAY = &H80
    CF_DSPTEXT = &H81
    CF_DSPBITMAP = &H82
    CF_DSPMETAFILEPICT = &H83
    CF_DSPENHMETAFILE = &H8E
    
    'IMPORTANT NOTE!  Data on the clipboard in this range is *not* owned by the clipboard, but by the application who stuck the data there.
    ' As such, attempting to call GlobalFree() will cause major issues.  Don't do it.
    CF_PRIVATEFIRST = &H200
    CF_PRIVATELAST = &H2FF
    
    'PD defines some internal formats, which is awesome for Cut/Copy/Paste within PD herself.
    CF_PD_DIB = CF_PRIVATEFIRST + 1
     
    'IMPORTANT NOTE!  These have nothing to do with GDI objects, despite the stupid names.  They are simply private-range objects
    ' that the Clipboard itself is responsible for freeing.  Per MSDN:
    ' "An application can place data handles on the clipboard by defining a private format in the range CF_GDIOBJFIRST through CF_GDIOBJLAST.
    '  When using values in this range, the data handle is not a handle to a Windows Graphics Device Interface (GDI) object, but is a handle
    '  allocated by the GlobalAlloc function with the GMEM_MOVEABLE flag. When the clipboard is emptied the system automatically deletes the
    '  object using the GlobalFree function."
    ' So basically, these are identical to CF_PRIVATEFIRST/LAST, except the clipboard is responsible for freeing them - not the owner.
    CF_GDIOBJFIRST = &H300
    CF_GDIOBJLAST = &H3FF

End Enum

#If False Then
    Private Const CF_MIN = 0, CF_TEXT = 1, CF_BITMAP = 2, CF_METAFILEPICT = 3, CF_SYLK = 4, CF_DIF = 5, CF_TIFF = 6, CF_OEMTEXT = 7
    Private Const CF_DIB = 8, CF_PALETTE = 9, CF_PENDATA = 10, CF_RIFF = 11, CF_WAVE = 12, CF_UNICODETEXT = 13, CF_ENHMETAFILE = 14
    Private Const CF_HDROP = 15, CF_LOCALE = 16, CF_DIBV5 = 17, CF_MAX = 18, CF_OWNERDISPLAY = &H80, CF_DSPTEXT = &H81
    Private Const CF_DSPBITMAP = &H82, CF_DSPMETAFILEPICT = &H83, CF_DSPENHMETAFILE = &H8E, CF_PRIVATEFIRST = &H200, CF_PRIVATELAST = &H2FF
    Private Const CF_GDIOBJFIRST = &H300, CF_GDIOBJLAST = &H3FF, CF_PD_DIB = 0
#End If

'Helper function for dealing with file lists
Private Declare Function DragQueryFile Lib "shell32" Alias "DragQueryFileW" (ByVal hDrop As Long, ByVal iFile As Long, ByVal lpStrW As Long, ByVal sizeOfStringBuffer As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type DROPFILES
    pFiles As Long
    Pt As POINTAPI
    fNC As Long
    fWide As Long
End Type

Private Const MAX_PATH As Long = 260
Private Const MAX_PATH_UNICODE As Long = 1024

'If the caller activates delayed rendering, we need to subclass a window and listen for messages
Private Const WM_RENDERFORMAT As Long = &H305
Private Const WM_RENDERALLFORMATS As Long = &H306
Private Const WM_DESTROYCLIPBOARD As Long = &H307
Private m_Subclass As cSelfSubHookCallback
Attribute m_Subclass.VB_VarHelpID = -1

'If delayed rendering is active, we need to know if we are responsible for the current clipboard data.  Prior to program shutdown,
' if we're *still* responsible for clipboard data, we must render all supported formats prior to shutdown (otherwise, when another
' program asks us for a copy of our data, we'll already be gone!)
Private m_DelayedRenderingActive As Boolean

'Register a new clipboard format.  This step is important, because multiple programs may register the same format (e.g. PNG).
' Using native clipboard functions guarantees that all programs registering the same format name get the same ID in return.
'
'RETURNS: clipboard format ID if successful; 0 otherwise.
Public Function AddClipboardFormat(ByVal sName As String) As Long
    
    Dim wFormat As Long
    wFormat = RegisterClipboardFormatW(StrPtr(sName))
    
    'Per MSDN, "Registered clipboard formats are identified by values in the range 0xC000 through 0xFFFF."
    If (wFormat >= &HC000&) Then AddClipboardFormat = wFormat Else AddClipboardFormat = 0&
    
End Function

'Number of formats currently located on the clipboard
Public Function GetClipboardFormatCount() As Long
    GetClipboardFormatCount = CountClipboardFormats()
End Function

'Return a comma-delimited list of currently available string formats.  This function exists primarily for debugging, and is not intended to
' be used for actual clipboard enumerations (hence the weird return format).
Public Function GetListOfAvailableFormatNames() As String

    If m_NumOfClipboardFormats > 0 Then
        
        Dim finalList As String
        finalList = m_ClipboardNames(0) & " (" & m_ClipboardIDs(0) & ")"
        
        If m_NumOfClipboardFormats > 1 Then
            Dim i As Long
            For i = 1 To m_NumOfClipboardFormats - 1
                finalList = finalList & ", " & m_ClipboardNames(i) & " (" & m_ClipboardIDs(i) & ")"
            Next i
        End If
        
        GetListOfAvailableFormatNames = finalList
        
    Else
        GetListOfAvailableFormatNames = ""
    End If
    
End Function

'Given an index into our module-level clipboard ID array, return the corresponding ID value
Public Function GetFormatIDByIndex(ByVal lIndex As Long) As Long
    If lIndex >= 0 And lIndex < m_NumOfClipboardFormats Then
        GetFormatIDByIndex = m_ClipboardIDs(lIndex)
    Else
        GetFormatIDByIndex = 0
    End If
End Function

'Given an index into our module-level clipboard name array, return the corresponding format name
Public Function GetFormatNameByIndex(ByVal lIndex As Long) As String
    If lIndex >= 0 And lIndex < m_NumOfClipboardFormats Then
        GetFormatNameByIndex = m_ClipboardNames(lIndex)
    Else
        GetFormatNameByIndex = ""
    End If
End Function

'See if a given clipboard format ID exists on the current clipboard.
' IMPORTANT NOTE: you must have previously opened the clipboard (in order to call UpdateClipboardLists) for this function to work.
Public Function DoesClipboardHaveFormatID(ByVal srcFormatId As PredefinedClipboardFormatConstants) As Boolean
    
    DoesClipboardHaveFormatID = False
    
    If m_NumOfClipboardFormats > 0 Then
    
        'Search for the given ID
        Dim i As Long
        For i = 0 To m_NumOfClipboardFormats - 1
            If (m_ClipboardIDs(i) = srcFormatId) Then
                DoesClipboardHaveFormatID = True
                Exit For
            End If
        Next i
        
    End If
    
End Function

'Shortcut function to check for HTML text.  HTML text is handled weirdly, because it's system-defined (see
' https://msdn.microsoft.com/en-us/library/windows/desktop/ms649015%28v=vs.85%29.aspx), but it doesn't use a hard-coded ID.
' IMPORTANT NOTE: you must have previously opened the clipboard (in order to call UpdateClipboardLists) for this function to work.
Public Function DoesClipboardHaveHTML() As Boolean
    DoesClipboardHaveHTML = Me.DoesClipboardHaveFormatName("HTML Format")
End Function

'Shortcut function to check for one or more file paths.  File lists are handled weirdly, via an HDROP handle.
' IMPORTANT NOTE: you must have previously opened the clipboard (in order to call UpdateClipboardLists) for this function to work.
Public Function DoesClipboardHaveFiles() As Boolean
    DoesClipboardHaveFiles = Me.DoesClipboardHaveFormatID(CF_HDROP)
End Function

'Shortcut function to check for text.  Text, OEM text, and Unicode text all cause a valid return.
' IMPORTANT NOTE: you must have previously opened the clipboard (in order to call UpdateClipboardLists) for this function to work.
Public Function DoesClipboardHaveText() As Boolean
    DoesClipboardHaveText = Me.DoesClipboardHaveFormatID(CF_TEXT) Or Me.DoesClipboardHaveFormatID(CF_UNICODETEXT) Or Me.DoesClipboardHaveFormatID(CF_OEMTEXT)
End Function

'Shortcut function to check for bitmaps and DIBs.  Note that the OS auto-converts between CF_BITMAP, CF_DIB, and CF_DIBv5 for you.
' IMPORTANT NOTE: you must have previously opened the clipboard (in order to call UpdateClipboardLists) for this function to work.
Public Function DoesClipboardHaveBitmapImage() As Boolean
    DoesClipboardHaveBitmapImage = Me.DoesClipboardHaveFormatID(CF_BITMAP) Or Me.DoesClipboardHaveFormatID(CF_DIB) Or Me.DoesClipboardHaveFormatID(CF_DIBV5)
End Function

'See if a given clipboard format name exists on the current clipboard.  The comparison is not case-sensitive.
' IMPORTANT NOTE: you must have previously opened the clipboard (in order to call UpdateClipboardLists) for this function to work.
Public Function DoesClipboardHaveFormatName(ByVal srcFormatName As String) As Boolean
    
    DoesClipboardHaveFormatName = False
    
    If m_NumOfClipboardFormats > 0 Then
    
        'Search for the given name
        Dim i As Long
        For i = 0 To m_NumOfClipboardFormats - 1
            If StrComp(m_ClipboardNames(i), srcFormatName, vbTextCompare) = 0 Then
                DoesClipboardHaveFormatName = True
                Exit For
            End If
        Next i
        
    End If
    
End Function

'Return the name for a given clipboard format ID.  (Format IDs for custom formats can be obtained via AddClipboardFormat(), above.)
' By design, this function does not support translations.  It is for internal PD use only.
Public Function GetFormatNameFromID(ByVal clipFormatID As PredefinedClipboardFormatConstants) As String
    
    'Predefined clipboard format names can be supplied without relying on the OS
    If (clipFormatID > CF_MIN And clipFormatID < CF_MAX) Then
        
        Select Case clipFormatID
        
            Case CF_TEXT
                GetFormatNameFromID = "Text/default"
                
            Case CF_BITMAP
                GetFormatNameFromID = "Bitmap image"
                
            Case CF_METAFILEPICT
                GetFormatNameFromID = "Metafile image"
            
            Case CF_SYLK
                GetFormatNameFromID = "Microsoft Symbolic Link (SYLK) data."
                
            Case CF_DIF
                GetFormatNameFromID = "Software Arts' Data Interchange information."
                
            Case CF_TIFF = 6
                GetFormatNameFromID = "Tiff image"
                
            Case CF_OEMTEXT
                GetFormatNameFromID = "Text/OEM"
                
            Case CF_DIB
                GetFormatNameFromID = "DIB image"
                
            Case CF_PALETTE
                GetFormatNameFromID = "Color palette"
                
            Case CF_PENDATA
                GetFormatNameFromID = "Pen data"
                
            Case CF_RIFF
                GetFormatNameFromID = "RIFF audio"
                
            Case CF_WAVE
                GetFormatNameFromID = "Wave audio"
                
            Case CF_UNICODETEXT
                GetFormatNameFromID = "Text/Unicode"
                
            Case CF_ENHMETAFILE
                GetFormatNameFromID = "Enhanced Metafile image"
            
            Case CF_HDROP
                GetFormatNameFromID = "File list"
                
            Case CF_LOCALE
                GetFormatNameFromID = "Text Locale Identifier"
                
            Case CF_DIBV5
                GetFormatNameFromID = "DIBv5 image"
                
        End Select
        
    'Custom clipboard formats require us to retrieve names at run-time
    Else
    
        'Size of the name buffer.  Per MSDN comments (https://msdn.microsoft.com/en-us/library/windows/desktop/ms649040):
        ' "Testing finds that RegisterClipboardFormat() returns error 87 ERROR_INVALID_PARAMETER if lpszGetClipboardFormatName is 256 characters or longer.
        '  Attempts to register a zero length name result in error 123 ERROR_INVALID_NAME. In other words, the lpszGetClipboardFormatName can be from 1 to 255
        '  characters plus the terminating NUL."
        Dim largestNameSize As Long
        largestNameSize = 255
        
        Dim tmpString As String
        tmpString = String$(largestNameSize, 0&)
        
        'The returned value of the API name function is the actual size of the string, not counting the terminating null-char.
        Dim actualLength As Long
        actualLength = GetClipboardFormatNameW(clipFormatID, StrPtr(tmpString), largestNameSize)
        
        'If a non-zero-length name is returned, trim it automatically
        If (actualLength <> 0) Then GetFormatNameFromID = Left$(tmpString, actualLength) Else GetFormatNameFromID = ""
        
    End If
    
End Function

'Given a named clipboard format, return the matching ID.
' IMPORTANT NOTE: the clipboard must be open for this to work.
Public Function GetFormatIDFromName(ByVal sName As String) As Long
    
    GetFormatIDFromName = 0
    
    'Update our current clipboard format list
    If Me.UpdateClipboardLists() Then
    
        'Search for a matching format ID
        Dim i As Long
        For i = 0 To m_NumOfClipboardFormats - 1
            If (sName = m_ClipboardNames(i)) Then
                GetFormatIDFromName = m_ClipboardIDs(i)
                Exit For
            End If
        Next i
        
    End If
    
End Function

'Given a clipboard format ID, return the corresponding data chunk as a raw byte array.  It's up to the caller to parse the data further.
' Returns: TRUE if successful; FALSE if something goes horribly wrong.
Public Function GetClipboardBinaryData(ByVal clipFormatID As PredefinedClipboardFormatConstants, ByRef dstData() As Byte) As Boolean
    
    GetClipboardBinaryData = False
    
    'Try to retrieve a handle to the binary blob associated with this format.
    Dim hGlobalHandle As Long
    hGlobalHandle = GetClipboardMemoryHandle(clipFormatID)
    If (hGlobalHandle <> 0) Then
        
        'Make sure the blob has a non-zero size
        Dim blobSize As Long
        blobSize = GlobalSize(hGlobalHandle)
        If (blobSize <> 0) Then
            
            'Attempt to lock the memory
            Dim ptrBlob As Long
            ptrBlob = GlobalLock(hGlobalHandle)
            If ptrBlob <> 0 Then
                
                'Copy the data locally, then immediately unlock the handle.  (Note that the clipboard is still responsible
                ' for the data, so we don't must not free it.)
                ReDim dstData(0 To blobSize - 1) As Byte
                CopyMemoryStrict VarPtr(dstData(0)), ptrBlob, blobSize
                GlobalUnlock hGlobalHandle
                
                GetClipboardBinaryData = True
                
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  pdClipboard.GetClipboardBinaryData() couldn't lock the clipboard data.  Clipboard operation abandoned."
                #End If
            End If
        
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  pdClipboard.GetClipboardBinaryData() couldn't grab the clipboard data, because blob size = 0."
            #End If
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdClipboard.GetClipboardBinaryData() couldn't grab the clipboard data, because the hGlobal handle = 0."
        #End If
    End If
    
End Function

'Get a handle (HGLOBAL, specifically) to the clipboard data associated with a given format ID
Public Function GetClipboardMemoryHandle(ByVal clipFormatID As PredefinedClipboardFormatConstants) As Long
    
    GetClipboardMemoryHandle = 0
    
    'Make sure the clipboard is open, or handle retrieval will fail
    If m_IsClipboardOpen Then
        
        'Attempt to retrieve a handle to the data; zero indicates failure
        GetClipboardMemoryHandle = GetClipboardData(clipFormatID)
        
        #If DEBUGMODE = 1 Then
            If GetClipboardMemoryHandle = 0 Then pdDebug.LogAction "WARNING!  pdClipboard.GetClipboardMemoryHandle() returned zero.  API Error code: " & Err.LastDllError
        #End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  You can't retrieve clipboard memory handles (via GetClipboardMemoryHandle) unless the clipboard is open!"
        #End If
    End If
    
End Function

'Stick new data on the clipboard.  An associated clipboard format is required.
' (IMPORTANT NOTE: the data placed on the clipboard is owned by the clipboard - NOT you - so we don't need to worry about freeing it later.)
' (IMPORTANT NOTE: the passed array must be initialized, or the function will fail.)
Public Function SetClipboardBinaryData(ByVal clipFormatID As PredefinedClipboardFormatConstants, ByRef srcData() As Byte) As Boolean
    
    SetClipboardBinaryData = False
    
    On Error GoTo SetClipboardFailure
    
    'Make sure the clipboard is actually open.  (Note that delayed rendering requires to explicitly *not* open the clipboard
    ' prior to rendering data, as another program owns the clipboard as part of their Paste action - this is the sole exception.)
    If m_IsClipboardOpen Or m_DelayedRenderingActive Then
        
        'Determine how much binary data we need to write
        Dim blobSize As Long
        blobSize = (UBound(srcData) - LBound(srcData)) + 1
        
        'Ask the system for a global chunk of (moveable!) memory.  The moveable part is important; clipboard data requires it, so the system
        ' can move resources around as necessary.
        Dim hGlobalHandle As Long
        hGlobalHandle = GlobalAlloc(GMEM_MOVEABLE, blobSize)
        If (hGlobalHandle <> 0) Then
            
            'Get a raw pointer to our new global memory object
            Dim ptrGlobal As Long
            ptrGlobal = GlobalLock(hGlobalHandle)
            If ptrGlobal <> 0 Then
            
                'Copy our data over, then unlock the memory block.  (Remember that the clipboard is going to assume ownership over our
                ' hGlobal object.)
                CopyMemoryStrict ptrGlobal, VarPtr(srcData(LBound(srcData))), blobSize
                GlobalUnlock hGlobalHandle
                        
                'Hand the data over to the clipboard.  If we succeed, the data is no longer ours.
                SetClipboardBinaryData = CBool(SetClipboardData(clipFormatID, hGlobalHandle) <> 0)
                
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  SetClipboardBinaryData couldn't allocate a valid global memory handle; clipboard operation abandoned."
                #End If
            End If
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  SetClipboardBinaryData couldn't allocate a valid global memory handle; clipboard operation abandoned."
            #End If
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  You can't stick stuff on the clipboard (via SetClipboardBinaryData) unless the clipboard is open, or delayed rendering is active!"
        #End If
    End If
    
    Exit Function
    
SetClipboardFailure:
    SetClipboardBinaryData = False

End Function

'If the caller's already allocated a global object, they can use this shortcut function to set that object into the clipboard.
' (NOTE: CF_BITMAP is an exception here, as you can pass this function a local DDB handle, and the clipboard will automatically
'        create its own copy, without any intervention from you.)
Public Function SetClipboardMemoryHandle(ByVal clipFormatID As PredefinedClipboardFormatConstants, ByVal hMem As Long) As Boolean
    SetClipboardMemoryHandle = CBool(SetClipboardData(clipFormatID, hMem) <> 0)
End Function

'Helper function, for when the caller wants to upload custom stuff to the clipboard.  We already have all the global memory APIs
' declared and implemented here, so this saves the caller some trouble.
Public Function GetGlobalMemoryHandle(ByVal memSize As Long) As Long
    GetGlobalMemoryHandle = GlobalAlloc(GMEM_MOVEABLE, memSize)
End Function
    
'Helper function to GetGlobalMemoryHandle(), above.
'IMPORTANT NOTE:  Make sure to call the corresponding FinishedWithGlobalMemoryPtr function when you're done.
Public Function GetPointerFromGlobalMemory(ByVal srcHGlobal As Long) As Long
    If (srcHGlobal <> 0) Then GetPointerFromGlobalMemory = GlobalLock(srcHGlobal)
End Function

'Helper function to GetPointerFromGlobalMemory, above.
'IMPORTANT NOTE:  You must pass this function the original hGlobal handle, NOT the pointer generated by GlobalLock.
Public Sub FinishedWithGlobalMemoryPtr(ByVal srcHGlobal As Long)
    If srcHGlobal <> 0 Then GlobalUnlock srcHGlobal
End Sub

'Retrieve text from the clipboard.  Format doesn't matter; PD always tries to grab Unicode text, if possible, with the system implicitly
' applying any necessary conversions.
Public Function GetClipboardText(ByRef dstString As String, Optional ClipboardFormatID As PredefinedClipboardFormatConstants = CF_UNICODETEXT) As Boolean
    
    GetClipboardText = False
    
    'Make sure the clipboard has data in the target format.  Note that for the default CF_UNICODETEXT ID, any ASCII-format text will
    ' be auto-converted to UTF-16 by the system.
    If Me.DoesClipboardHaveFormatID(ClipboardFormatID) Then
        
        'Retrieve the data into a byte array
        Dim tmpData() As Byte
        If GetClipboardBinaryData(ClipboardFormatID, tmpData()) Then
            
            'Use a pdUnicode object to apply format conversion
            Dim cUnicode As pdUnicode
            Set cUnicode = New pdUnicode
            
            If (ClipboardFormatID = CF_UNICODETEXT) Or (ClipboardFormatID = CF_TEXT) Then
                dstString = cUnicode.ConvertUTF16PointerToVBString(VarPtr(tmpData(0)), UBound(tmpData) + 1)
                GetClipboardText = CBool(Len(dstString) <> 0)
            Else
                GetClipboardText = cUnicode.ConvertUnknownBytesToString(tmpData, dstString)
            End If
            
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdClipboard.GetClipboardText() failed, because the requested format ID doesn't exist on the clipboard."
        #End If
    End If
    
End Function

'Retrieve HTML text from the clipboard.  You should verify the presence of HTML text before using this function, as HTML text will
' likely be in UTF-8 format, and this function will apply format heuristics to try and verify this.  If you simply need plain text,
' skip the need for heuristics, and use the normal GetClipboardText() function, above.
Public Function GetClipboardHTML(ByRef dstString As String) As Boolean
    
    GetClipboardHTML = False
    
    'Make sure the clipboard has data in the target format.  Note that for the default CF_UNICODETEXT ID, any ASCII-format text will
    ' be auto-converted to UTF-16 by the system.
    Dim htmlID As Long
    htmlID = Me.GetFormatIDFromName("HTML Format")
    If Me.DoesClipboardHaveFormatID(htmlID) Then
        
        'Retrieve the data into a byte array
        Dim tmpData() As Byte
        If GetClipboardBinaryData(htmlID, tmpData()) Then
            
            'Use a pdUnicode object to apply format conversion
            Dim cUnicode As pdUnicode
            Set cUnicode = New pdUnicode
            GetClipboardHTML = cUnicode.ConvertUnknownBytesToString(tmpData, dstString)
            
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdClipboard.GetClipboardHTML() failed, because the clipboard doesn't have HTML text."
        #End If
    End If
    
End Function

'If the clipboard contains one or more file paths, you can retrieve them via this function.  Note that you must pass a string array,
' even if there is only a single file on the clipboard.
Public Function GetFileList(ByRef dstFileList() As String, ByRef numOfFiles As Long) As Boolean
    
    GetFileList = False
    
    'Make sure files actually exist on the clipboard
    If Me.DoesClipboardHaveFiles() Then
    
        'Retrieve a handle to the hDrop (drop group struct)
        Dim hDrop As Long
        hDrop = GetClipboardData(CF_HDROP)
        If hDrop <> 0 Then
            
            'Get the number of files in the list
            numOfFiles = DragQueryFile(hDrop, -1&, 0&, 0)
            If numOfFiles > 0 Then
                
                'Prep the target array, and a temporary buffer for each (null-terminated) return
                ReDim dstFileList(0 To numOfFiles - 1) As String
                Dim tmpFilename As String, strLength As Long
                tmpFilename = String$(MAX_PATH_UNICODE, 0&)
                
                'Retrieve each filename in the list
                Dim i As Long
                For i = 0 To numOfFiles - 1
                   
                    'Get the length of this string
                    strLength = DragQueryFile(hDrop, i, 0&, 0&)
                    If DragQueryFile(hDrop, i, StrPtr(tmpFilename), MAX_PATH_UNICODE) <> 0 Then
                        dstFileList(i) = Left$(tmpFilename, strLength)
                    End If
                   
                Next i
                
                GetFileList = True
   
            Else
                numOfFiles = 0
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  pdClipboard.GetFileList() failed, because the hDrop contained zero valid files."
                #End If
            End If

        Else
            numOfFiles = 0
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  pdClipboard.GetFileList() failed, because the clipboard didn't return a valid hDrop."
            #End If
        End If
        
    Else
        numOfFiles = 0
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdClipboard.GetFileList() failed, because the clipboard doesn't contains any filenames."
        #End If
    End If
    
End Function

'If a program places any one of CF_BITMAP, CF_DIB, or CF_DIBv5 on the clipboard, Windows will automatically list all three formats
' as available.  Other programs can then request whichever format they like best, and Windows will silently convert between them
' as necessary.  In PD, we prefer to know which format the caller *actually* supplied, because that lets us know whether to take
' things like ICC profiles and color endpoints in the CF_DIBv5 entry seriously.
'
'Call this function to find out which BITMAP format the caller actually placed on the clipboard, and which ones were auto-created
' for them by Windows.
Public Function GetPriorityBitmapFormat() As PredefinedClipboardFormatConstants
    
    'CF_MIN is a placeholder value that equals (first_valid_entry - 1).  For this function, it's considered a failure value,
    ' as subsequenct select statements would fail to hit it.
    GetPriorityBitmapFormat = CF_MIN
    
    If m_NumOfClipboardFormats > 0 Then
    
        'Per MSDN (https://msdn.microsoft.com/en-us/library/windows/desktop/ms649013%28v=vs.85%29.aspx):
        ' If the system provides an automatic type conversion for a particular clipboard format, and you call EnumClipboardFormats
        ' to enumerate the clipboard data formats, the system first enumerates the format that is on the clipboard, followed by the
        ' formats to which it can be converted.
        '
        'This lets us know which bitmap format the caller *actually* placed on the clipboard, vs which ones are auto-created.
        ' (And since we cache the results of EnumClipboardFormats when the clipboard is opened, retrieving the first hit is easy!)
        Dim i As Long
        For i = 0 To m_NumOfClipboardFormats - 1
            If (m_ClipboardIDs(i) = CF_BITMAP) Or (m_ClipboardIDs(i) = CF_DIB) Or (m_ClipboardIDs(i) = CF_DIBV5) Then
                GetPriorityBitmapFormat = m_ClipboardIDs(i)
                Exit For
            End If
        Next i
        
    End If
    
End Function

'If a program places any one of CF_TEXT, CF_UNICODETEXT, or CF_OEMTEXT on the clipboard, Windows will automatically list all three
' formats as available.  Other programs can then request whichever format they like best, and Windows will silently convert between
' them as necessary.  In PD, we sometimes prefer to know which format the caller *actually* supplied.
'
'Call this function to find out which TEXT format the caller actually placed on the clipboard, and which ones were auto-created
' for them by Windows.
Public Function GetPriorityTextFormat() As PredefinedClipboardFormatConstants
    
    'CF_MIN is a placeholder value that equals (first_valid_entry - 1).  For this function, it's considered a failure value,
    ' as subsequenct select statements would fail to hit it.
    GetPriorityTextFormat = CF_MIN
    
    If m_NumOfClipboardFormats > 0 Then
    
        'Per MSDN (https://msdn.microsoft.com/en-us/library/windows/desktop/ms649013%28v=vs.85%29.aspx):
        ' If the system provides an automatic type conversion for a particular clipboard format, and you call EnumClipboardFormats
        ' to enumerate the clipboard data formats, the system first enumerates the format that is on the clipboard, followed by the
        ' formats to which it can be converted.
        '
        'This lets us know which text format the caller *actually* placed on the clipboard, vs which ones are auto-created.
        ' (And since we cache the results of EnumClipboardFormats when the clipboard is opened, retrieving the first hit is easy!)
        Dim i As Long
        For i = 0 To m_NumOfClipboardFormats - 1
            If (m_ClipboardIDs(i) = CF_TEXT) Or (m_ClipboardIDs(i) = CF_UNICODETEXT) Or (m_ClipboardIDs(i) = CF_OEMTEXT) Then
                GetPriorityTextFormat = m_ClipboardIDs(i)
                Exit For
            End If
        Next i
        
    End If
    
End Function

'Prior to writing data, you *must* clear the clipboard.  Otherwise, any unwritten formats will remain, potentially fucking up other
' programs who preferentially read clipboard formats in a different order than yours.
Public Function ClearClipboard() As Boolean
    If m_IsClipboardOpen Then
        ClearClipboard = CBool(EmptyClipboard <> 0)
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdClipboard.ClearClipboard doesn't work unless the clipboard is open!"
        #End If
    End If
End Function

'Only one program at a time can open the clipboard, so control needs to be relinquished as quickly as possible.
Public Sub ClipboardClose()
    If m_IsClipboardOpen Then
        CloseClipboard
        m_IsClipboardOpen = False
    End If
End Sub

'Open the clipboard.  You must supply an owner hWnd, as some clipboard APIs rely on that hWnd to function.  (Custom clipboard formats that
' consume tons of space can be rendered dynamically, via window messages, rather than stuffing a blob into memory - hence the hWnd requirement.)
Public Function ClipboardOpen(ByVal hWndOwner As Long) As Boolean
    
    If (OpenClipboard(hWndOwner) <> 0) Then
        
        'Cache the owner hWnd and clipboard state, so we don't have to request them again
        m_OwnerHwnd = hWndOwner
        m_IsClipboardOpen = True
        
        'Update the list of available clipboard formats, then exit
        Me.UpdateClipboardLists
        ClipboardOpen = True
        
    Else
        m_OwnerHwnd = 0
        m_IsClipboardOpen = False
        ClipboardOpen = False
    End If
    
End Function

'Enumerate the names and IDs of all items currently on the clipboard.  Values will stores in the module-level m_Clipboard_() arrays.
' IMPORTANT NOTE: this action is performed automatically whenever the clipboard is opened (via ClipboardOpen()).  You should never
' need to call it manually.
'
'RETURNS: TRUE if successful and at least one format exists on the clipboard.  FALSE otherwise.
Public Function UpdateClipboardLists() As Boolean
    
    'Make sure the clipboard open
    If m_IsClipboardOpen Then
    
        'Reset the module-level clipboard trackers
        m_NumOfClipboardFormats = 0
        
        'Initialize the tracking arrays to the number of objects on the clipboard
        Dim lR As Long
        lR = CountClipboardFormats()
        
        If (lR <> 0) Then
        
            ReDim m_ClipboardIDs(0 To lR - 1) As Long
            ReDim m_ClipboardNames(0 To lR - 1) As String
        
            'Enumerate each format in return.
            lR = EnumClipboardFormats(0&)
        
            If (lR <> 0) Then
                
                Do
                    
                    'As a failsafe, make sure our destination arrays are large enough to hold all clipboard formats
                    If m_NumOfClipboardFormats > UBound(m_ClipboardIDs) Then
                        ReDim Preserve m_ClipboardIDs(1 To m_NumOfClipboardFormats) As Long
                        ReDim Preserve m_ClipboardNames(1 To m_NumOfClipboardFormats) As String
                    End If
                    
                    'Retrieve the new ID and name values
                    m_ClipboardIDs(m_NumOfClipboardFormats) = lR
                    m_ClipboardNames(m_NumOfClipboardFormats) = GetFormatNameFromID(lR)
                    lR = EnumClipboardFormats(m_ClipboardIDs(m_NumOfClipboardFormats))
                    
                    'Increment the clipboard format count and carry on
                    m_NumOfClipboardFormats = m_NumOfClipboardFormats + 1
                    
                Loop While lR <> 0
                
            End If
            
        End If
        
        UpdateClipboardLists = True
        
    Else
        UpdateClipboardLists = False
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  You can't enumerate clipboard formats without first opening the clipboard!"
        #End If
    End If
    
End Function

'Set clipboard data in delayed rendering mode.  This does not require source data; instead, it simply notifies the clipboard that we're
' the owner of the current clipboard data.  If/when someone else requests clipboard data, this class will raise a matching event -
' at THAT point you must upload the actual clipboard data, in the requested format.
Public Function SetClipboardData_DelayedRendering(ByVal clipFormatID As PredefinedClipboardFormatConstants) As Boolean
        
    SetClipboardData_DelayedRendering = False
    
    On Error GoTo SetClipboardDRFailure
    
    'Make sure the clipboard is actually open
    If m_IsClipboardOpen Then
        
        'Delayed rendering requires substantially different behavior elsewhere in the class, so we mark its state at class-level
        m_DelayedRenderingActive = True
        
        'If we haven't already, subclass the main window and start listening for three clipboard-related messages
        If (m_Subclass Is Nothing) Then
            Set m_Subclass = New cSelfSubHookCallback
            If (m_OwnerHwnd <> 0) Then
                m_Subclass.ssc_Subclass m_OwnerHwnd, 0, 1, Me
                m_Subclass.ssc_AddMsg m_OwnerHwnd, MSG_BEFORE, WM_RENDERFORMAT, WM_RENDERALLFORMATS, WM_DESTROYCLIPBOARD
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  pdClipboard.SetClipboardData_DelayedRendering failed to initialized the subclasser, because you haven't given it an hWnd to subclass!"
                #End If
            End If
        End If
        
        'Set a null-copy of the current data to the clipboard.  This tells the clipboard that we're willing to provide this data
        ' on an "as-needed" basis.
        SetClipboardData_DelayedRendering = CBool(SetClipboardData(clipFormatID, 0&) <> 0&)
        
    End If
    
    Exit Function
    
SetClipboardDRFailure:
    SetClipboardData_DelayedRendering = False
        
End Function

'Want to know if PD is responsible for the current clipboard data?  (You may want to do this prior to the program shutting down,
' for example, since WM_RENDERALLFORMATS doesn't play nicely with the way VB destroys windows.)
Public Function IsOurDataOnTheClipboard() As Boolean
    IsOurDataOnTheClipboard = CBool(GetClipboardOwner() = m_OwnerHwnd)
End Function

'On termination, we do a failsafe clipboard close, but it *really* should be closed manually as soon as you're done with it!
Private Sub Class_Terminate()
    If m_IsClipboardOpen Then ClipboardClose
    If Not (m_Subclass Is Nothing) Then
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Terminating clipboard subclasser.  I hope all clipboard formats havee been rendered already..."
        #End If
        
        Set m_Subclass = Nothing
        
        'Not that it matters, but also remove the delayed rendering flag, as a reminder that we no longer have access to that clipboard data
        m_DelayedRenderingActive = False
        
    End If
End Sub

'All events subclassed by this window are processed here.
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************

    'This class subclasses three relevant clipboard messages:
    Select Case uMsg
    
        'Another program wants to paste our data; the relevant format is in the wParam
        Case WM_RENDERFORMAT
            RaiseEvent ClipboardRenderFormat(wParam)
            bHandled = True
            lReturn = 0
        
        'PD is going down; we must render ALL possible formats to the clipboard, in case someone else wants them later
        Case WM_RENDERALLFORMATS
            RaiseEvent ClipboardRenderAllFormats
            bHandled = True
            lReturn = 0
        
        'Another program was used to copy data, meaning our stash is now irrelevant.  Feel free to clear it out.
        Case WM_DESTROYCLIPBOARD
            If Not IsOurDataOnTheClipboard() Then
                RaiseEvent ClipboardDestroyStashedData
                m_DelayedRenderingActive = False
            End If
            bHandled = True
            lReturn = 0
            
    End Select
    

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub

