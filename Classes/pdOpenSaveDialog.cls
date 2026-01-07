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
'Copyright 2015-2026 by Tanner Helland
'Created: 05/August/15
'Last updated: 29/May/22
'Last update: add convenience helper for multi-select, returning the results as a pdStringStack object
'
'Common dialogs appear in multiple places throughout PD.  In the beginning, PD relied on a convenient
' common dialog wrapper from vbAccelerator...
'
'http://www.vbaccelerator.com/home/VB/Code/Libraries/Common_Dialogs/Code_Only_Common_Dialogs/article.html
'
'...but besides being unwieldy and somewhat error-prone, Steve's implementation was ANSI-only.  Due to major
' differences in string handling (including multifile delimiters), it was non-trivial to overhaul his code,
' so in 2015 I wrote a new common dialog interface, specific to PhotoDemon.
'
'The eventual goal is to add parallel support for Vista's IFileOpen/SaveDialogs, which would allow me to
' (finally) customize the program's main dialogs, but as of today that remains TODO.  For now, PD simply
' wraps the old-school common dialog APIs.
'
'(Also: for convenience against existing code, parameter order is roughly identical to Steve's old implementation.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Unicode interface.  Note that the only thing we pass it is a pointer to an OPENFILENAME struct.
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

Private Enum Win32_CommonDialogFlags
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
Public Enum Win32_CommonDialogErrors
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

'Arbitrary-length return strings are possible, but because we have to re-raise the dialog upon failure (ugh), it's not
' really practical to implement a "try again with larger buffer" feature.  Fortunately, post-XP operatings systems have
' no upper limit, so we can start with an arbitrarily large string.
Private Const MAX_PATH_LOAD_XP As Long = 16384
Private Const MAX_PATH_LOAD As Long = 262144
Private Const MAX_PATH_SAVE As Long = 1024
Private Const MAX_FILE As Long = 260

'Display a system-specific Open File dialog.
' RETURNS: TRUE if one or more files was selected; FALSE if canceled.
Friend Function GetOpenFileName(ByRef dstFilename As String, _
                           Optional ByRef fileAndExtensionOnlyNoPath As String, _
                           Optional ByVal fileMustExist As Boolean = True, _
                           Optional ByVal allowMultiSelect As Boolean = False, _
                           Optional ByVal extensionFilters As String = "All (*.*)|*.*", _
                           Optional ByRef idxFilter As Long = 1, _
                           Optional ByRef initialFolder As String, _
                           Optional ByRef titleBarCaption As String, _
                           Optional ByRef defaultExtension As String, _
                           Optional ByVal hWndOwner As Long = 0&) As Boolean
    
    'Lock various UI bits and stop responding to certain window messages
    Interface.NotifySystemDialogState True
    
    'Prep a relevant OFN struct
    Dim tmpOFN As OPENFILENAME
    
    With tmpOFN
        .lStructSize = Len(tmpOFN)
        
        'OFN_EXPLORER is explicitly flagged to force Explorer mode even when allowing multiselect,
        ' and OFN_PATHMUSTEXIST is used to prevent the user from entering invalid filenames.
        .Flags = OFN_PATHMUSTEXIST Or OFN_LONGNAMES Or OFN_HIDEREADONLY
        If fileMustExist Then .Flags = .Flags Or OFN_FILEMUSTEXIST
        If allowMultiSelect Then .Flags = .Flags Or OFN_ALLOWMULTISELECT Or OFN_EXPLORER
        
        'Various other obvious flags
        .hWndOwner = hWndOwner
        .lpstrInitialDir = initialFolder
        .lpstrDefExt = defaultExtension
        .lpstrTitle = titleBarCaption
        
        'The filter string must be passed in null-delimited format, but this is non-trivial to do in VB6.
        ' So we ask callers to use pipe separators, then manually convert to nulls before invoking the dialog.
        Dim tmpString As String
        tmpString = extensionFilters
        If (InStr(1, tmpString, "|", vbBinaryCompare) <> 0) Then tmpString = Replace$(tmpString, "|", vbNullChar)
        
        'Double-nulls mark final termination
        .lpstrFilter = tmpString & vbNullChar & vbNullChar
        .nFilterIndex = idxFilter
        
        'Because we allocate our own return string, we must pass the common dialog an allocated buffer
        ' of some arbitrary size.  If the size proves too small to hold all files selected by the user,
        ' the call to GetOpenFilename will fail, and we can't retry without re-raising the dialog -
        ' so make sure the initial size is as big as possible!
        Dim initPathSize As Long
        If OS.IsVistaOrLater Then
            initPathSize = MAX_PATH_LOAD
        Else
            initPathSize = MAX_PATH_LOAD_XP
        End If
        
        'Null-pad strings before passing them to the common dialog
        .nMaxFile = initPathSize
        .lpstrFile = dstFilename & String$(initPathSize - Len(dstFilename), 0)
        
        .nMaxFileTitle = MAX_FILE
        .lpstrFileTitle = fileAndExtensionOnlyNoPath & String$(MAX_FILE - Len(fileAndExtensionOnlyNoPath), 0)
        
        'All other fields are fine as-is
        
    End With
    
    'Turn control over to the common dialog
    Dim lReturn As Long
    lReturn = GetOpenFileNameW(VarPtr(tmpOFN))
    
    'Before proceeding, add debug info in the special case of a buffer underrun.  This case is problematic,
    ' as we can't simply "try again" without re-raising a modal dialog, so I'm hoping my current default
    ' buffer size is always sufficient.
    If (lReturn = 0) And (CommDlgExtendedError() = FNERR_BUFFERTOOSMALL) Then
        PDMsgBox "Unfortunately, the number of images you selected exceeds what Windows allows for a single load operation.  Please try again with a smaller set of images.", vbExclamation Or vbOKOnly, "Error"
        PDDebug.LogAction "WARNING: pdOpenSaveDialog experienced a common dialog buffer size error.", PDM_Normal
    End If
    
    If (lReturn <> 0) Then
        
        'Success!  Parse out the relevant strings.
        Dim lLen As Long
        
        'In multi-select mode, parsing out individual filenames is left to the caller
        If allowMultiSelect Then
            dstFilename = tmpOFN.lpstrFile
        
        'In single-file mode, we handle filename assembly for the caller (as a convenience)
        Else
            
            If (tmpOFN.nFileOffset <> 0) Then
                lLen = InStr(tmpOFN.nFileOffset, tmpOFN.lpstrFile, vbNullChar) - 1
                If (lLen < 1) Then lLen = Len(tmpOFN.lpstrFile)
            Else
                lLen = lstrlenW(StrPtr(tmpOFN.lpstrFile))
            End If
            
            If (lLen <> 0) Then dstFilename = Left$(tmpOFN.lpstrFile, lLen)
            
        End If
        
        lLen = lstrlenW(StrPtr(tmpOFN.lpstrFileTitle))
        If (lLen <> 0) Then fileAndExtensionOnlyNoPath = Left$(tmpOFN.lpstrFileTitle, lLen)
        
        'Return the filter index, so the caller can store it to preferences as desired
        idxFilter = tmpOFN.nFilterIndex
        
        GetOpenFileName = True
    
    'A return value of 0 means the user canceled the dialog
    Else
        Message "Load canceled. "
        GetOpenFileName = False
        dstFilename = vbNullString
        fileAndExtensionOnlyNoPath = vbNullString
    End If
    
    'Re-enable various hooks, subclassers, and UI bits
    Interface.NotifySystemDialogState False
                           
End Function

'Display a system-specific Save File dialog.
' RETURNS: TRUE if a valid save target was selected/created; FALSE if canceled.
Friend Function GetSaveFileName(ByRef dstFilename As String, _
                           Optional ByRef fileAndExtensionOnlyNoPath As String, _
                           Optional ByVal showOverWritePrompt As Boolean = True, _
                           Optional ByRef extensionFilters As String = "All (*.*)|*.*", _
                           Optional ByRef idxFilter As Long = 1, _
                           Optional ByRef initialFolder As String, _
                           Optional ByRef titleBarCaption As String, _
                           Optional ByRef defaultExtension As String, _
                           Optional ByVal hWndOwner As Long = 0&) As Boolean

    Interface.NotifySystemDialogState True
    
    'Prep a relevant OFN struct
    Dim tmpOFN As OPENFILENAME
    
    With tmpOFN
        .lStructSize = Len(tmpOFN)
        
        'We forcefully hide the "Open as Read Only" checkbox, as it makes little sense for PD.
        .Flags = OFN_HIDEREADONLY
        If showOverWritePrompt Then .Flags = .Flags Or OFN_OVERWRITEPROMPT
        
        'Various other obvious flags
        .hWndOwner = hWndOwner
        .lpstrInitialDir = initialFolder
        .lpstrDefExt = defaultExtension
        .lpstrTitle = titleBarCaption
        
        'The filter string must be passed in null-delimited format, but this is non-trivial to do in VB6.
        ' So we ask callers to use pipe separators, then manually convert to nulls before invoking the dialog.
        Dim tmpString As String
        tmpString = extensionFilters
        If (InStr(1, tmpString, "|", vbBinaryCompare) <> 0) Then tmpString = Replace$(tmpString, "|", vbNullChar)
        
        'Double-nulls mark final termination
        .lpstrFilter = tmpString & vbNullChar & vbNullChar
        .nFilterIndex = idxFilter
        
        'We have to start with a prepared buffer of some arbitrary size.  If the size proves too small,
        ' we could technically call the function again with a larger buffer (not implemented, as PD uses
        ' a comically huge buffer).
        Dim initPathSize As Long
        initPathSize = MAX_PATH_SAVE
        
        'Null-pad all strings to their max size
        .lpstrFile = dstFilename & String$(initPathSize - Len(dstFilename), 0)
        .nMaxFile = initPathSize
        
        .lpstrFileTitle = fileAndExtensionOnlyNoPath & String$(MAX_FILE - Len(fileAndExtensionOnlyNoPath), 0)
        .nMaxFileTitle = MAX_FILE
        
        'All other fields are fine as-is
        
    End With
    
    'Turn control over to the common dialog function
    Dim lReturn As Long
    lReturn = GetSaveFileNameW(VarPtr(tmpOFN))
    
    'Non-zero returns are considered successful, but note that a failure state may have occurred due to a too-small buffer.
    ' (We could technically handle buffer overflows by retrying with a larger buffer, but I don't know how to do that
    '  without re-displaying the entire effing common dialog; to that end, I've hard-coded a 1K buffer as a limit,
    '  which should allow for saving to just about anywhere!)
    If (lReturn <> 0) Then
        
        'Success!  Parse out the relevant filename string.
        Dim lLen As Long
        If (tmpOFN.nFileOffset <> 0) Then
            lLen = InStr(tmpOFN.nFileOffset, tmpOFN.lpstrFile, vbNullChar) - 1
            If (lLen < 1) Then lLen = Len(tmpOFN.lpstrFile)
        Else
            lLen = lstrlenW(StrPtr(tmpOFN.lpstrFile))
        End If
        
        If (lLen <> 0) Then dstFilename = Left$(tmpOFN.lpstrFile, lLen)
        
        'As a convenience, return the file title too (which is just the filename + extension, with no path data)
        lLen = lstrlenW(StrPtr(tmpOFN.lpstrFileTitle))
        If (lLen <> 0) Then fileAndExtensionOnlyNoPath = Left$(tmpOFN.lpstrFileTitle, lLen)
        
        'Return the filter index, so the caller can store it to preferences as desired
        idxFilter = tmpOFN.nFilterIndex
        
        GetSaveFileName = True
        
    Else
        Message "Save canceled. "
        GetSaveFileName = False
    End If
    
    'Re-enable various hooks, subclassers, and UI bits
    Interface.NotifySystemDialogState False
    
End Function

'Display a system-specific Open File dialog with allowMultiSelect = TRUE, and return the list of selected file(s)
' as as pdStringStack object.
' RETURNS: TRUE if one or more files was selected; FALSE if canceled.
Friend Function GetOpenFileNames_AsStringStack(ByRef dstStringStack As pdStringStack, ByRef dstFilename As String, _
                           Optional ByRef fileAndExtensionOnlyNoPath As String, _
                           Optional ByVal fileMustExist As Boolean = True, _
                           Optional ByVal extensionFilters As String = "All (*.*)|*.*", _
                           Optional ByRef idxFilter As Long = 1, _
                           Optional ByRef initialFolder As String, _
                           Optional ByRef titleBarCaption As String, _
                           Optional ByRef defaultExtension As String, _
                           Optional ByVal hWndOwner As Long = 0&) As Boolean

    GetOpenFileNames_AsStringStack = Me.GetOpenFileName(dstFilename, fileAndExtensionOnlyNoPath, fileMustExist, True, extensionFilters, idxFilter, initialFolder, titleBarCaption, defaultExtension, hWndOwner)
    
    If GetOpenFileNames_AsStringStack Then
    
        If (dstStringStack Is Nothing) Then Set dstStringStack = New pdStringStack
        
        'Take the return string (a null-delimited list of filenames) and split it into an array
        Dim listOfFiles() As String
        listOfFiles = Split(dstFilename, vbNullChar)
        
        Dim i As Long
        
        'Due to buffering required by the API call, uBound(listOfFiles) should ALWAYS > 0 but
        ' check it anyway (just to be safe)
        If (UBound(listOfFiles) > 0) Then
        
            'Remove all empty strings from the array (which are a byproduct of the aforementioned buffering)
            For i = UBound(listOfFiles) To 0 Step -1
                If (LenB(listOfFiles(i)) <> 0) Then Exit For
            Next
            
            'With all empty strings removed, only legitimate file paths remain
            ReDim Preserve listOfFiles(0 To i) As String
            
        End If
        
        'If multiple files were selected, we need to perform additional processing
        If (UBound(listOfFiles) > 0) Then
        
            'The common dialog function returns a unique array. Index (0) contains the folder path
            ' (without a trailing backslash), so first things first - add a trailing backslash
            Dim basePath As String
            basePath = Files.PathAddBackslash(listOfFiles(0))
            
            'The remaining indices contain filenames within that base folder.  To get the full filename,
            ' we must append the path from (0) to the start of each filename.  This simplifies things for
            ' the caller - it can simply loop through full paths, loading files as it goes.
            For i = 1 To UBound(listOfFiles)
                dstStringStack.AddString basePath & listOfFiles(i)
            Next i
            
        'If there is only one file in the array (e.g. the user only opened one image), we don't need to
        ' perform extra processing - just save the new directory to the preferences file
        Else
            dstStringStack.AddString listOfFiles(0)
        End If
        
        'Success requires at least one valid return
        GetOpenFileNames_AsStringStack = (dstStringStack.GetNumOfStrings > 0)
    
    End If

End Function
