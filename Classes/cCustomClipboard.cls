VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCustomClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'This class was originally written by Steve McMahon.

'You may download the original version of this code from the following link (good as of May '13):
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Clipboard/Customising_Clipboard_Use/article.asp

'To the best of my knowledge, this code is released under a CC-BY-1.0 license.  (Assumed from the footer text of vbaccelerator.com: "All contents of this web site are licensed under a Creative Commons Licence, except where otherwise noted.")
' You may access a complete copy of this license at the following link:
' http://creativecommons.org/licenses/by/1.0/

'Many thanks to Steve McMahon and vbaccelerator.com for this excellent clipboard interface.

Option Explicit

' =================================================================
' File:     cCustomClipboard
' Author:   SP McMahon 9 March 1998
' Requires: Self contained
' Version:  2.0
'
' Allows complete access to the clipboard.
'   1) Define custom clipboard formats
'   2) Read binary or text data from formats other
'      than those provided with VB as well as the VB
'      ones.
'   3) Enumerate types of data on the clipboard.
'
' Note when Copying to the Clipboard, using this class, you
' muse follow this order:
'
'   ' Get access to the clipboard:
'   .ClipboardOpen Me.hWnd
'       ' Become the clipboard owner:
'       .ClearClipboard
'
'       ' Do copying in all formats here:
'
'   ' Give clipboard control over allocated memory:
'   .ClipboardClose
'
' If you do not follow this order, GPF may result, particularly
' under NT4.
'
' =================================================================

' Clipboard functions:
Private Declare Function OpenClipboard Lib "user32" _
    (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" _
    () As Long
Private Declare Function GetClipboardData Lib "user32" _
    (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" _
    (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" _
    (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" _
    () As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" _
    (ByVal lpString As String) As Long
Private Declare Function EnumClipboardFormats Lib "user32" _
    (ByVal wFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" _
    (ByVal wFormat As Long) As Long
Private Declare Function CountClipboardFormats Lib "user32" () As Long

' Memory functions:
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

' Members:
Private m_lId() As Long
Private m_sName() As String
Private m_iCount As Long
Private m_bClipboardIsOpen As Boolean
Private m_hWnd As Long

'/*
' * Predefined Clipboard Formats
' */
Public Enum EPredefinedClipboardFormatConstants
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
''#if(WINVER >= 0x0400)
     CF_HDROP = 15
     CF_LOCALE = 16
     CF_MAX = 17
'#endif /* WINVER >= 0x0400 */
     CF_OWNERDISPLAY = &H80
     CF_DSPTEXT = &H81
     CF_DSPBITMAP = &H82
     CF_DSPMETAFILEPICT = &H83
     CF_DSPENHMETAFILE = &H8E
'/*
' * "Private" formats don't get GlobalFree()'d
' */
     CF_PRIVATEFIRST = &H200
     CF_PRIVATELAST = &H2FF
'/*
' * "GDIOBJ" formats do get DeleteObject()'d
' */
     CF_GDIOBJFIRST = &H300
     CF_GDIOBJLAST = &H3FF

End Enum

' New shell-oriented clipboard formats
'Private Const CFSTR_SHELLIDLIST As String = "Shell IDList Array"
'Private Const CFSTR_SHELLIDLISTOFFSET As String = "Shell Object Offsets"
'Private Const CFSTR_NETRESOURCES As String = "Net Resource"
'Private Const CFSTR_FILEDESCRIPTOR As String = "FileGroupDescriptor"
'Private Const CFSTR_FILECONTENTS As String = "FileContents"
'Private Const CFSTR_FILENAME As String = "FileName"
'Private Const CFSTR_PRINTERGROUP As String = "PrinterFriendlyName"
'Private Const CFSTR_FILENAMEMAP As String = "FileNameMap"

Public Enum ECustomClipboardErrorConstant
    eccErrorBase = vbObjectError + 1048 + 521
    eccClipboardNotOpen
    eccCantOpenClipboard
End Enum

Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal Ch As Long) As Long

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


Public Function AddFormat( _
        ByVal sName As String _
    ) As Long
' Adds a custom clipboard format and returns its
' ID if successful, otherwise returns 0.
Dim wFormat As Long
    wFormat = RegisterClipboardFormat(sName & Chr$(0))
    If (wFormat > &HC000&) Then
        AddFormat = wFormat
    End If
End Function
Property Get FormatCount() As Integer
' Returns the number of formats available on the
' clipboard:
    FormatCount = CountClipboardFormats()
End Property
Public Function GetCurrentFormats(ByVal hWndOwner As Long) As Long
' Enumerates all the names and IDs of items currently
' on the clipboard, and returns the number of items:
Dim lR As Long

    Erase m_lId
    Erase m_sName
    m_iCount = 0
    If (ClipboardOpen(hWndOwner)) Then
        lR = EnumClipboardFormats(0)
        If (lR <> 0) Then
            Do
                m_iCount = m_iCount + 1
                ReDim Preserve m_lId(1 To m_iCount) As Long
                ReDim Preserve m_sName(1 To m_iCount) As String
                m_lId(m_iCount) = lR
                m_sName(m_iCount) = FormatName(lR)
                lR = EnumClipboardFormats(m_lId(m_iCount))
            Loop While lR <> 0
        End If
    End If
    GetCurrentFormats = m_iCount
    ClipboardClose
End Function
Public Property Get GetCurrentFormatID(ByVal lIndex As Long)
' Returns the clipboard format id obtained by GetCurrentFormats
' at 1 based position lIndex
    GetCurrentFormatID = m_lId(lIndex)
End Property
Public Property Get GetCurrentFormatName(ByVal lIndex As Long)
' Returns the clipboard format name obtained by GetCurrentFormats
' at 1 based position lIndex
    GetCurrentFormatName = m_sName(lIndex)
End Property
Public Property Get HasCurrentFormat(ByVal lFormatId As Long) As Boolean
Dim iFormat As Long
    For iFormat = 1 To m_iCount
        If (m_lId(iFormat) = lFormatId) Then
            HasCurrentFormat = True
            Exit For
        End If
    Next iFormat
End Property
Public Property Get FormatName( _
        ByVal lFormatId As Long _
    ) As String
' Returns the format name for a clipboard format id:
Dim lSize As Long
Dim sBuf As String
Dim lR As Long
    
    If (lFormatId >= 1 And lFormatId <= 17) Then
        ' For pre-defined formats, we have to make the text
        ' up ourselves:
        Select Case lFormatId
        Case CF_TEXT
            FormatName = "Text"
        Case CF_BITMAP
            FormatName = "Bitmap Picture"
        Case CF_METAFILEPICT
            FormatName = "Meta-File Picture"
        Case CF_SYLK
            FormatName = "Microsoft Symbolic Link (SYLK) data."
        Case CF_DIF
            FormatName = "Software Arts' Data Interchange information."
        Case CF_TIFF = 6
            FormatName = "Tagged Image File Format (TIFF) Picture"
        Case CF_OEMTEXT
            FormatName = "Text (OEM)"
        Case CF_DIB
            FormatName = "DIB Bitmap Picture"
        Case CF_PALETTE
            FormatName = "Colour Palette"
        Case CF_PENDATA
            FormatName = "Pen Data"
        Case CF_RIFF
            FormatName = "RIFF Audio data"
        Case CF_WAVE
            FormatName = "Wave File"
        Case CF_UNICODETEXT
            FormatName = "Text (Unicode)"
        Case CF_ENHMETAFILE
            FormatName = "Enhanced Meta-File Picture"
''#if(WINVER >= 0x0400)
         Case CF_HDROP
            FormatName = "File List"
        Case CF_LOCALE = 16
            FormatName = "Text Locale Identifier"
        End Select
    Else
        ' For custom formats, we can ask the Clipboard for
        ' the registered name:
        lSize = 255
        sBuf = String$(lSize, 0)
        lR = GetClipboardFormatName(lFormatId, sBuf, lSize)
        If (lR <> 0) Then
            FormatName = Left$(sBuf, lR)
        End If
    End If
End Property
Public Property Get FormatIDForName( _
        ByVal hWndOwner As Long, _
        ByVal sName As String _
    ) As Long
' Searches for the Name sName on the Clipboard, and returns the
' format ID for it (or 0 if the item cannot be found)
Dim i As Integer
Dim lId As Long

    GetCurrentFormats hWndOwner
    For i = 1 To m_iCount
        If (sName = m_sName(i)) Then
            lId = m_lId(i)
            Exit For
        End If
    Next i
    FormatIDForName = lId
    
End Property
Property Get IsDataAvailableForFormatName( _
        ByVal hWndOwner As Long, _
        ByVal sFormatName As String _
    ) As Boolean
' Returns whether data is available for a given format name:
Dim lId As Long
Dim lR As Long
    lId = FormatIDForName(hWndOwner, sFormatName)
    If (lId > 0) Then
        lR = IsClipboardFormatAvailable(lId)
        IsDataAvailableForFormatName = (lR <> 0)
    End If
End Property
Property Get IsDataAvailableForFormat( _
        ByVal lFormatId As Long _
    )
' Returns whether data is available for a given format id:
Dim lR As Long
    lR = IsClipboardFormatAvailable(lFormatId)
    IsDataAvailableForFormat = (lR <> 0)
End Property

Public Function GetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean

    ' Returns a byte array containing binary data on the clipboard for
    ' format lFormatID:
    Dim hMem As Long, lSize As Long, lPtr As Long
    
    ' Ensure the return array is clear:
    Erase bData
    
    hMem = GetClipboardMemoryHandle(lFormatId)
    
    ' If success:
    If (hMem <> 0) Then
        
        ' Get the size of this memory block:
        lSize = GlobalSize(hMem)
        
        ' Get a pointer to the memory:
        lPtr = GlobalLock(hMem)
        
        If (lSize > 0) Then
            
            ' Resize the byte array to hold the data:
            ReDim bData(0 To lSize - 1) As Byte
            ' Copy from the pointer into the array:
            CopyMemory bData(0), ByVal lPtr, lSize
        
        Else
            Debug.Print "The size of the PNG object is reported as 0.  Check GlobalSize API instructions."
        End If
        
        ' Unlock the memory block:
        GlobalUnlock hMem
        
        ' Success:
        GetBinaryData = (lSize > 0)
        
        ' Don't free the memory - it belongs to the clipboard.
        
    Else
        Debug.Print "hMem = 0; it's likely something went wrong with GetClipboardMemoryHandle!"
    End If
    
End Function
Public Function GetClipboardMemoryHandle( _
        ByVal lFormatId As Long _
    ) As Long
    
    If pbNotReady() Then
        Debug.Print "pbNotReady; exiting now"
        Exit Function
    End If
    
    ' If the format id is there:
    If (IsDataAvailableForFormat(lFormatId)) Then
    
        ' Get the global memory handle to the clipboard data:
        GetClipboardMemoryHandle = GetClipboardData(lFormatId)
        
        'If 0 is returned, report the last API error
        If GetClipboardMemoryHandle = 0 Then
            Debug.Print "GetClipboardData() returned 0.  Last error was: " & Err.LastDllError
        End If
        
    Else
        Debug.Print "IsDataAvailableForFormat returned 0 for " & lFormatId & "."
    End If
    
End Function
Public Function SetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean
' Puts the binary data contained in bData() onto the clipboard under
' format lFormatID:
Dim lSize As Long
Dim lPtr As Long
Dim hMem As Long

    If pbNotReady() Then Exit Function
    
    ' Determine the size of the binary data to write:
    lSize = UBound(bData) - LBound(bData) + 1
    ' Generate global memory to hold this:
    hMem = GlobalAlloc(GMEM_DDESHARE, lSize)
    If (hMem <> 0) Then
        ' Get pointer to the memory block:
        lPtr = GlobalLock(hMem)
        ' Copy the data into the memory block:
        CopyMemory ByVal lPtr, bData(LBound(bData)), lSize
        ' Unlock the memory block.
        GlobalUnlock hMem
                
        ' Now set the clipboard data:
        If (SetClipboardData(lFormatId, hMem) <> 0) Then
            ' Success:
            SetBinaryData = True
        End If
    End If
    ' We don't free the memory because the clipboard takes
    ' care of that now.

End Function
Public Function SetClipboardMemoryHandle( _
        ByVal lFormatId As Long, _
        ByVal hMem As Long _
    ) As Boolean
    SetClipboardMemoryHandle = (SetClipboardData(lFormatId, hMem) <> 0)
End Function
Public Function GetTextData( _
        ByVal lFormatId As Long, _
        ByRef sTextOut As String _
    ) As Boolean
' Returns a string containing text on the clipboard for
' format lFormatID:
Dim lHwndCache As Long

    If (lFormatId = CF_TEXT) Or (lFormatId = CF_UNICODETEXT) Then
        ' Use VB method, temporarily we close the clipboard:
        If (m_bClipboardIsOpen) Then
            lHwndCache = m_hWnd
            ClipboardClose
        End If
        sTextOut = Clipboard.GetText
        If (lHwndCache <> 0) Then
            ClipboardOpen lHwndCache
        End If
        GetTextData = True
    Else
        Dim bData() As Byte
        If (GetBinaryData(lFormatId, bData())) Then
            sTextOut = StrConv(bData, vbUnicode)
            GetTextData = True
        End If
    End If
End Function
Public Function SetTextData( _
        ByVal lFormatId As Long, _
        ByVal sText As String _
    ) As Boolean
Dim bData() As Byte
Dim i As Long
' Sets the text in sText onto the clipboard under format
' lFormatID:
    If (Len(sText) <> 0) Then
        ReDim bData(0 To Len(sText)) As Byte
        For i = 0 To Len(sText) - 1
            bData(i) = Asc(Mid$(sText, i + 1, 1))
        Next i
        SetTextData = SetBinaryData(lFormatId, bData())
    End If
End Function
Public Function GetFileList( _
      ByRef sFiles() As String, _
      ByRef iFileCount As Long _
   ) As Boolean
Dim hDrop As Long
Dim iFile As Long
Dim sFilename As String
Dim iPos As Long

   ' Get handle to CF_HDROP if any:
   hDrop = GetClipboardMemoryHandle(CF_HDROP)
   If (hDrop <> 0) Then
      iFileCount = DragQueryFile(hDrop, -1&, "", 0)
         
      If (iFileCount > 0) Then
         ' Allocate space for return and working variables.
         ReDim sFiles(1 To iFileCount) As String
         sFilename = String$(MAX_PATH, 0)
   
         ' Retrieve each filename in Dropped Filelist.
         For iFile = 1 To iFileCount
            DragQueryFile hDrop, iFile - 1, sFilename, Len(sFilename)
            iPos = InStr(sFilename, vbNullChar)
            If (iPos <> 0) Then
               sFiles(iFile) = Left$(sFilename, iPos - 1)
            Else
               sFiles(iFile) = sFilename
            End If
         Next iFile
         
         GetFileList = True
      End If
   End If
   
End Function
Public Function SetFileList( _
      ByRef sFiles() As String _
   ) As Boolean
Dim sData As String
Dim tDF As DROPFILES
Dim b() As Byte
Dim i As Long

   If pbNotReady() Then Exit Function

   ' Build double-null terminated list of files.
   For i = LBound(sFiles) To UBound(sFiles)
      sData = sData & sFiles(i) & vbNullChar
   Next i
   sData = sData & vbNullChar

   ' Copy a dropfiles structure then this into a data array:
   ReDim b(0 To Len(tDF) + Len(sData) - 1) As Byte
   tDF.pFiles = Len(tDF)
   CopyMemory b(0), tDF, Len(tDF)
   CopyMemoryStr b(Len(tDF)), sData, Len(sData)
   
   ' Put the data on the clipboard:
   SetFileList = SetBinaryData(CF_HDROP, b())
   
End Function

Private Function pbNotReady() As Boolean
' Determines whether a call to Get or Set Data on the
' clipboard will work.
    If Not (m_bClipboardIsOpen) Or (m_hWnd = 0) Then
        'Debug.Assert (1 = 0)
        'Err.Raise eccClipboardNotOpen, App.EXEName & ".cCustomClipboard", "Attempt to access the clipboard when clipboard not Open."
        pbNotReady = True
    End If
End Function

Public Sub ClearClipboard()
' Clears all data in the clipboard, and also takes ownership
' of the clipboard.  This method will fail
' unless OpenClipboard has been called first.
    If (pbNotReady()) Then Exit Sub
    EmptyClipboard
End Sub
Public Sub ClipboardClose()
' Closes the clipboard if this class has it open:
    If (m_bClipboardIsOpen) Then
        CloseClipboard
        m_bClipboardIsOpen = False
        m_hWnd = 0
    End If
End Sub
Public Function ClipboardOpen( _
        ByVal hWndOwner As Long _
    ) As Boolean
Dim lR As Long
' Opens the clipboard:
    lR = OpenClipboard(hWndOwner)
    If (lR > 0) Then
        m_hWnd = hWndOwner
        m_bClipboardIsOpen = True
        ClipboardOpen = True
        'Message "CLIPBOARD OPEN"
    Else
        m_hWnd = 0
        m_bClipboardIsOpen = False
        'Err.Raise eccCantOpenClipboard, App.EXEName & ".cCustomClipboard", "Unable to Open Clipboard."
    End If
End Function

Private Sub Class_Terminate()
    ' We Shouldn't have the clipboard open here, but if it is,
    ' and you manage to get this far(!) the clipboard will be
    ' closed...
    ClipboardClose

End Sub


