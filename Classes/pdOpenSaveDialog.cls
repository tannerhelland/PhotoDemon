VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdOpenSaveDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PD System Open/Save Dialog Interface
'Copyright 2015 by Tanner Helland
'Created: 05/August/15
'Last updated: 15/August/15
'Last update: finish work on GetSaveFileName()
'
'Common dialogs appear in multiple places throughout PD.  In the beginning, PD relied on a convenient common dialog wrapper
' from vbAccelerator (http://www.vbaccelerator.com/home/VB/Code/Libraries/Common_Dialogs/Code_Only_Common_Dialogs/article.asp),
' but besides being unwieldy and somewhat error-prone, Steve's implementation was ANSI-only.  Due to some major differences in
' string handling (including multifile delimiters), it was non-trivial to overhaul his control, so in 2015, I wrote a new
' common dialog interface, specific to PD.
'
'The eventual goal is to add parallel support for Vista's IFileOpen/SaveDialogs, which would actually allow me to customize
' the program's main dialogs, but that's a ways off.  For now, everything simply wraps the traditional common dialog controls.
'
'For convenience against existing code, the function calls are roughly identical to Steve's old control.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'Unicode interface.  Note that the only thing we pass it is a pointer to the master OPENFILENAME struct; it does the rest.
Private Declare Function GetOpenFileNameW Lib "comdlg32" (ByVal ptrToOFN As Long) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32" (ByVal ptrToOFN As Long) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

'MSDN struct breakdown: https://msdn.microsoft.com/en-us/library/ms646839%28v=vs.85%29.aspx
' Note that some parameter behavior (e.g. lpstrInitialDir) varies under Win 7 and later.
Private Type OPENFILENAME
    lStructSize          As Long
    hWndOwner            As Long
    hInstance            As Long
    lpstrFilter          As String
    lpstrCustomFilter    As String
    nMaxCustFilter       As Long
    nFilterIndex         As Long
    lpstrFile            As String
    nMaxFile             As Long
    lpstrFileTitle       As String
    nMaxFileTitle        As Long
    lpstrInitialDir      As String
    lpstrTitle           As String
    Flags                As Long
    nFileOffset          As Integer
    nFileExtension       As Integer
    lpstrDefExt          As String
    lCustData            As Long
    lpfnHook             As Long
    lpTemplateName       As Long
    pvReserved           As Long
    dwReserved           As Long
    FlagsEx              As Long
End Type

Private Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

#If False Then
    Private Const OFN_READONLY = &H1, OFN_OVERWRITEPROMPT = &H2, OFN_HIDEREADONLY = &H4, OFN_NOCHANGEDIR = &H8, OFN_SHOWHELP = &H10
    Private Const OFN_ENABLEHOOK = &H20, OFN_ENABLETEMPLATE = &H40, OFN_ENABLETEMPLATEHANDLE = &H80, OFN_NOVALIDATE = &H100
    Private Const OFN_ALLOWMULTISELECT = &H200, OFN_EXTENSIONDIFFERENT = &H400, OFN_PATHMUSTEXIST = &H800, OFN_FILEMUSTEXIST = &H1000
    Private Const OFN_CREATEPROMPT = &H2000, OFN_SHAREAWARE = &H4000, OFN_NOREADONLYRETURN = &H8000&, OFN_NOTESTFILECREATE = &H10000
    Private Const OFN_NONETWORKBUTTON = &H20000, OFN_NOLONGNAMES = &H40000, OFN_EXPLORER = &H80000, OFN_NODEREFERENCELINKS = &H100000
    Private Const OFN_LONGNAMES = &H200000
#End If

'MSDN explanation of potential CD errors: https://msdn.microsoft.com/en-us/library/ms646916%28v=vs.85%29.aspx
Public Enum CommonDialogErrorsEnum
    CDERR_DIALOGFAILURE = &HFFFF&
    CDERR_FINDRESFAILURE = &H6
    CDERR_LOADRESFAILURE = &H7
    CDERR_INITIALIZATION = &H2
    CDERR_LOADSTRFAILURE = &H5
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    CDERR_NOHINSTANCE = &H4
    CDERR_NOHOOK = &H8
    CDERR_NOTEMPLATE = &H3
    CDERR_REGISTERMSGFAIL = &HC
    CDERR_STRUCTSIZE = &H1
    FNERR_BUFFERTOOSMALL = &H3003&
    FNERR_INVALIDFILENAME = &H3002&
    FNERR_SUBCLASSFAILURE = &H3001&
    CDERR_CANCELED = vbObjectError
End Enum

'Arbitrary-length return strings are possible, but unlike some of Windows's better APIs, we have to blindly try with an
' arbitrary buffer size, then repeat until an acceptable size is found.
Private Const MAX_PATH_LOAD As Long = 16384
Private Const MAX_PATH_SAVE As Long = 1024
Private Const MAX_FILE As Long = 260

'Display a system-specific Open File dialog.
' RETURNS: TRUE if one or more files was selected; FALSE if canceled.
Public Function GetOpenFileName(ByRef FileName As String, _
                           Optional ByRef FileTitle As String, _
                           Optional ByVal FileMustExist As Boolean = True, _
                           Optional ByVal MultiSelect As Boolean = False, _
                           Optional ByVal Filter As String = "All (*.*)|*.*", _
                           Optional ByRef FilterIndex As Long = 1, _
                           Optional ByRef InitDir As String, _
                           Optional ByRef DlgTitle As String, _
                           Optional ByRef DefaultExt As String, _
                           Optional ByVal Owner As Long = -1) As Boolean
                           
    Dim lReturn As Long, lInitPathSize As Long, Flags As Long
    
    'Prep a relevant OFN struct
    Dim tmpOFN As OPENFILENAME
    
    With tmpOFN
        .lStructSize = Len(tmpOFN)
        
        'OFN_EXPLORER is explicitly flagged to force Explorer mode even when allowing multiselect, and OFN_PATHMUSTEXIST
        ' is used to prevent the user from entering invalid filenames.
        .Flags = OFN_PATHMUSTEXIST Or OFN_LONGNAMES Or OFN_HIDEREADONLY
        If FileMustExist Then .Flags = .Flags Or OFN_FILEMUSTEXIST
        If MultiSelect Then .Flags = .Flags Or OFN_ALLOWMULTISELECT Or OFN_EXPLORER
        
        'Various other obvious flags
        If Owner <> -1 Then .hWndOwner = Owner Else .hWndOwner = 0&
        .lpstrInitialDir = InitDir
        .lpstrDefExt = DefaultExt
        .lpstrTitle = DlgTitle
        
        'The filter string is actually null-delimited, which isn't easy for the caller to do, so we allow the old-style
        ' pipe and colon separators, which are forcibly converted now.
        Dim Ch As String, i As Integer, s As String
        For i = 1 To Len(Filter)
            Ch = Mid$(Filter, i, 1)
            If Ch = "|" Or Ch = ":" Then
                s = s & vbNullChar
            Else
                s = s & Ch
            End If
        Next
        
        'Double-nulls mark final termination
        s = s & vbNullChar & vbNullChar
        .lpstrFilter = s
        .nFilterIndex = FilterIndex
        
        'We have to start with a prepared buffer of some arbitrary size.  If the size proves too small, we can call the function
        ' again with a larger buffer.  (NOT CURRENTLY IMPLEMENTED!)
        lInitPathSize = MAX_PATH_LOAD
        s = FileName & String$(lInitPathSize - Len(FileName), 0)
        .lpstrFile = s
        .nMaxFile = lInitPathSize
        
        s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
        .lpstrFileTitle = s
        .nMaxFileTitle = MAX_FILE
        
        'All other fields are fine as-is
        
    End With
    
    'Turn control over to the common dialog function
    lReturn = GetOpenFileNameW(VarPtr(tmpOFN))
    
    'Before proceeding, add debug info in the special case of a buffer underrun.  This case is problematic, as we can't simply
    ' "try again" without re-raising a modal dialog, so I'm hoping my current default buffer size is always sufficient.
    If (lReturn = 0) And (CommDlgExtendedError() = FNERR_BUFFERTOOSMALL) Then
        
        pdMsgBox "Unfortunately, the number of images you selected exceeds what Windows® allows for a single load operation.  Please try again with a smaller set of selected images.", vbExclamation + vbApplicationModal + vbOKOnly, "Error: too many images"
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING: pdOpenSaveDialog experienced a common dialog buffer size error.", PDM_NORMAL
        #End If
        
    End If
    
    'Non-zero returns are considered successful, but note that a failure state may have occurred due to a too-small buffer.
    ' (We could technically handle buffer overflows by retrying with a larger buffer, but I don't know how to do that
    '  without re-displaying the entire effing common dialog; to that end, I've hard-coded a 16K buffer as a limit,
    '  since PD won't do well with that many loaded images anyway.)
    If lReturn <> 0 Then
        
        'Success!  Parse out the relevant strings.
        
        Dim lLen As Long
        
        If Not MultiSelect Then
        
            If tmpOFN.nFileOffset <> 0 Then
                lLen = InStr(tmpOFN.nFileOffset, tmpOFN.lpstrFile, vbNullChar) - 1&
                If lLen < 1& Then lLen = Len(tmpOFN.lpstrFile)
            Else
                lLen = lstrlenW(StrPtr(tmpOFN.lpstrFile))
            End If
            
            If lLen Then FileName = Left$(tmpOFN.lpstrFile, lLen)
            
        Else
            FileName = tmpOFN.lpstrFile
        End If
        
        lLen = lstrlenW(StrPtr(tmpOFN.lpstrFileTitle))
        If lLen Then FileTitle = Left$(tmpOFN.lpstrFileTitle, lLen)
        
        'Return the filter index, so the caller can store it to preferences as desired
        FilterIndex = tmpOFN.nFilterIndex
        
        GetOpenFileName = True
        
    Else
        
        Message "Load canceled. "
        GetOpenFileName = False
        FileName = ""
        FileTitle = ""
        
    End If
                           
End Function

'Display a system-specific Save File dialog.
' RETURNS: TRUE if a valid save target was selected/created; FALSE if canceled.
Public Function GetSaveFileName(ByRef FileName As String, _
                           Optional ByRef FileTitle As String, _
                           Optional ByVal OverWritePrompt As Boolean = True, _
                           Optional ByRef Filter As String = "All (*.*)|*.*", _
                           Optional ByRef FilterIndex As Long = 1, _
                           Optional ByRef InitDir As String, _
                           Optional ByRef DlgTitle As String, _
                           Optional ByRef DefaultExt As String, _
                           Optional ByVal Owner As Long = -1) As Boolean

    Dim lReturn As Long, lInitPathSize As Long, Flags As Long
    
    'Prep a relevant OFN struct
    Dim tmpOFN As OPENFILENAME
    
    With tmpOFN
        .lStructSize = Len(tmpOFN)
        
        'We forcefully hide the "Open as Read Only" checkbox, as it makes little sense for PD.
        .Flags = OFN_HIDEREADONLY
        If OverWritePrompt Then .Flags = .Flags Or OFN_OVERWRITEPROMPT
        
        'Various other obvious flags
        If Owner <> -1 Then .hWndOwner = Owner
        .lpstrInitialDir = InitDir
        .lpstrDefExt = DefaultExt
        .lpstrTitle = DlgTitle
        
        'The filter string is actually null-delimited, which isn't easy for the caller to do, so we allow the old-style
        ' pipe and colon separators, which are forcibly converted now.
        Dim Ch As String, i As Integer, s As String
        'For i = 1 To Len(Filter)
        '    Ch = Mid$(Filter, i, 1)
        '    If Ch = "|" Or Ch = ":" Then
        '        s = s & vbNullChar
        '    Else
        '        s = s & Ch
        '    End If
        'Next
        
        s = Filter
        If InStr(1, s, "|", vbBinaryCompare) Then s = Replace$(s, "|", vbNullChar)
        If InStr(1, s, ":", vbBinaryCompare) Then s = Replace$(s, ":", vbNullChar)
        
        'Double-nulls mark final termination
        s = s & vbNullChar & vbNullChar
        .lpstrFilter = s
        .nFilterIndex = FilterIndex
        
        'We have to start with a prepared buffer of some arbitrary size.  If the size proves too small, we could technically call
        ' the function again with a larger buffer (not implemented, as PD uses a comically huge buffer).
        lInitPathSize = MAX_PATH_SAVE
        s = FileName & String$(lInitPathSize - Len(FileName), 0)
        .lpstrFile = s
        .nMaxFile = lInitPathSize
        
        s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
        .lpstrFileTitle = s
        .nMaxFileTitle = MAX_FILE
        
        'All other fields are fine as-is
        
    End With
    
    'Turn control over to the common dialog function
    lReturn = GetSaveFileNameW(VarPtr(tmpOFN))
    
    'Non-zero returns are considered successful, but note that a failure state may have occurred due to a too-small buffer.
    ' (We could technically handle buffer overflows by retrying with a larger buffer, but I don't know how to do that
    '  without re-displaying the entire effing common dialog; to that end, I've hard-coded a 1K buffer as a limit,
    '  which should allow for saving to just about anywhere!)
    If lReturn <> 0 Then
        
        'Success!  Parse out the relevant filename string.
        Dim lLen As Long
        If tmpOFN.nFileOffset <> 0 Then
            lLen = InStr(tmpOFN.nFileOffset, tmpOFN.lpstrFile, vbNullChar) - 1&
            If lLen < 1& Then lLen = Len(tmpOFN.lpstrFile)
        Else
            lLen = lstrlenW(StrPtr(tmpOFN.lpstrFile))
        End If
        
        If lLen Then FileName = Left$(tmpOFN.lpstrFile, lLen)
        
        'As a convenience, return the file title too (which is just the filename + extension, with no path data)
        lLen = lstrlenW(StrPtr(tmpOFN.lpstrFileTitle))
        If lLen Then FileTitle = Left$(tmpOFN.lpstrFileTitle, lLen)
        
        'Return the filter index, so the caller can store it to preferences as desired
        FilterIndex = tmpOFN.nFilterIndex
        
        GetSaveFileName = True
        
    Else
        
        Message "Save canceled. "
        GetSaveFileName = False
        
    End If
    
End Function
