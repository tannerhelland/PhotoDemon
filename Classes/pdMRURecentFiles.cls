VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdMRURecentFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'RecentFiles MRU List
'Copyright 2005-2015 by Tanner Helland
'Created: 16-02-2015
'Last updated: 16-02-2015
'Last updated by: Raj
'Last update: LAST UPDATE DETAILS HERE
'
'This class is a specific implementation of the IMRUList Interface.
'  It works with pdMRUManager to create and maintain the Recent
'  Files list.
'This is designed to be compatible to the old pdRecentFiles class,
'  so old MRU data should not be lost.
'Thumbnails are automatically saved (in PNG format) for each entry in the recent files list.  Thumbnail filenames are
' hashed as a privacy precaution.  PhotoDemon automatically erases unneeded thumbnail files as files are "pushed" off
' the recent file stack, and as an extra precaution, when the program exits, it scans the thumbnail directory for any
' orphaned thumbnails.  (Orphaned thumbnails are possible if you run multiple copies of PD in parallel.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit
Implements IMRUList

'Because we need to hash MRU names to generate icon save locations, and hashing is computationally expensive, store all
' calculated hashes in a table.
Private Type mruHash
    mruInitPath As String
    mruHashPath As String
End Type

Private m_mruHashes() As mruHash
Private m_numOfMRUHashes As Long

Private Function doesMRUHashExist(ByVal filePath As String) As String

    'Check to see if this file has been requested before.  If it has, return our previous
    ' hash instead of recalculating one from scratch.  If it does not exist, return "".
    If m_numOfMRUHashes > 0 Then
    
        'Loop through all previous hashes from this session
        Dim i As Long
        For i = 0 To m_numOfMRUHashes - 1
        
            'If this file path matches one we've already calculated, return that instead of calculating it again
            If StrComp(m_mruHashes(i).mruInitPath, filePath, vbTextCompare) = 0 Then
                doesMRUHashExist = m_mruHashes(i).mruHashPath
                Exit Function
            End If
        
        Next i
    
    End If
    
    doesMRUHashExist = ""

End Function


'Return a 16-character hash of a specific MRU entry.  (This is used to generate unique menu icon filenames.)
Private Function getMRUHash(ByVal filePath As String) As String
    
    'Check to see if this hash already exists
    Dim prevHash As String
    prevHash = doesMRUHashExist(filePath)
    
    'If it does, return it.
    If Len(prevHash) <> 0 Then
        getMRUHash = prevHash
        Exit Function
    
    'If no correlating hash was found, calculate one from scratch.
    Else
    
        'Prepare an SHA-256 hash calculator
        Dim cSHA2 As CSHA256
        Set cSHA2 = New CSHA256
            
        Dim hString As String
        hString = cSHA2.SHA256(filePath)
                
        'The SHA-256 function returns a 64 character string (256 / 8 = 32 bytes, but 64 characters due to hex representation).
        ' This is too long for a filename, so take only the first sixteen characters of the hash.
        hString = Left$(hString, 16)
        
        'Save this hash to our hashes array
        m_mruHashes(m_numOfMRUHashes).mruInitPath = filePath
        m_mruHashes(m_numOfMRUHashes).mruHashPath = hString
        m_numOfMRUHashes = m_numOfMRUHashes + 1
        ReDim Preserve m_mruHashes(0 To m_numOfMRUHashes) As mruHash
        
        'Return this as the hash value
        getMRUHash = hString
    
    End If
    
End Function

Private Function getMRUThumbnailPath(imagePath As String) As String
    getMRUThumbnailPath = g_UserPreferences.getIconPath & getMRUHash(imagePath) & ".png"
End Function

'Saves a thumbnail PNG of a pdImage object.  The thumbnail is saved to the /Data/Icons directory
Private Sub saveMRUThumbnail(ByRef imagePath As String, ByRef tImage As pdImage)

    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Saving recent file menu thumbnail..."
    #End If
    
    'First, generate a path at which to save the file in question
    Dim sFilename As String
    sFilename = getMRUThumbnailPath(imagePath)
    
    'Request a thumbnail from the current image, and store it in a temporary DIB.  Note that we can use large (64x64)
    ' thumbnails on Vista, but we are limited to 16x16 on XP.
    Dim mruThumb As pdDIB
    Set mruThumb = New pdDIB
    If g_IsVistaOrLater Then
        tImage.requestThumbnail mruThumb, 64
    Else
        tImage.requestThumbnail mruThumb, 16
    End If
    
    'Write the PNG file.  (FreeImage and GDI+ are both supported by this function, FYI)
    Saving.QuickSaveDIBAsPNG sFilename, mruThumb
    
    'Regardless of how our thumbnail was saved, we can now delete the temporary thumbnail DIB
    mruThumb.eraseDIB
    Set mruThumb = Nothing

End Sub

Private Sub IMRUList_BeforeListSave(manager As pdMRUManager, Optional cancel As Boolean = False)

End Sub

Private Sub IMRUList_BeforeNewFileAdded(manager As pdMRUManager, newFile As String, Optional cancel As Boolean = False)

End Sub

Private Property Get IMRUList_SupportsThumbNails() As Boolean
    IMRUList_SupportsThumbNails = True
End Property

Private Function IMRUList_GetThumbNailPath(manager As pdMRUManager, Index As Long) As Variant
    IMRUList_GetThumbNailPath = getMRUThumbnailPath(manager.getSpecificMRU(Index))
End Function

Private Sub IMRUList_SaveThumbNail(filePath As String, srcImage As pdImage)
    saveMRUThumbnail filePath, srcImage
End Sub


Private Sub IMRUList_AfterListCleared()
    'Unload all recent file menu entries
    If FormMain.mnuRecDocs.Count > 1 Then
        Dim i As Long
        For i = FormMain.mnuRecDocs.Count - 1 To 1 Step -1
            Unload FormMain.mnuRecDocs(i)
        Next i
    End If
End Sub

Private Sub IMRUList_AfterListSaved(manager As pdMRUManager)
    'Finally, scan the MRU icon directory to make sure there are no orphaned PNG files.  (Multiple instances of PhotoDemon
    ' running simultaneously can lead to this.)  Delete any PNG files that don't correspond to current MRU entries.
    Dim chkFile As String
    chkFile = Dir(g_UserPreferences.getIconPath & "*.png", vbNormal)
    
    Dim fileOK As Boolean
    
    Do While Len(chkFile) <> 0
        
        fileOK = False
        
        Dim i As Long
        Dim numEntries As Long
        numEntries = manager.MRU_ReturnCount
        'Compare this file to the hash for all current MRU entries
        If numEntries > 0 Then
            For i = 0 To numEntries - 1
                
                'If this hash matches one on file, mark it as OK.
                If StrComp(g_UserPreferences.getIconPath & chkFile, manager.getMRUThumbnailPath(i), vbTextCompare) = 0 Then
                    fileOK = True
                    Exit For
                End If
                
            Next i
        Else
            fileOK = False
        End If
        
        'If an MRU hash does not exist for this file, delete it
        If Not fileOK Then
            If FileExist(g_UserPreferences.getIconPath & chkFile) Then Kill g_UserPreferences.getIconPath & chkFile
        End If
    
        'Retrieve the next file and repeat
        chkFile = Dir
    
    Loop
End Sub

Private Sub IMRUList_AfterNewFileAdded(manager As pdMRUManager, newFile As String, mruCaption As String)
    
    'Redraw the MRU menu based on the updated list
    If Not FormMain.mnuRecDocs(0).Enabled Then
        FormMain.mnuRecDocs(0).Enabled = True
        FormMain.MnuRecentSepBar1.Visible = True
        FormMain.MnuLoadAllMRU.Visible = True
        FormMain.MnuClearMRU.Visible = True
    End If
    
    'Based on the user's preference, display just the filename or the entire file path (up to the max character length)
    FormMain.mnuRecDocs(0).Caption = mruCaption
    
    'Display the corresponding accelerator (hotkey).
    FormMain.mnuRecDocs(0).Caption = FormMain.mnuRecDocs(0).Caption & vbTab & g_Language.TranslateMessage("Ctrl") & "+0"
    
    Dim i As Long
    Dim numEntries As Long
    numEntries = manager.MRU_ReturnCount()
    'If there are more entries on the list besides this one, update their menus now.
    If numEntries > 1 Then
    
        'Unload existing menus...
        For i = FormMain.mnuRecDocs.Count - 1 To 1 Step -1
            Unload FormMain.mnuRecDocs(i)
        Next i
        
        'Load new menus...
        For i = 1 To numEntries - 1
        
            Load FormMain.mnuRecDocs(i)
            
            'Based on the user's preference, display just the filename or the entire file path (up to the max character length)
            FormMain.mnuRecDocs(i).Caption = manager.getSpecificMRUCaption(i)
            
            'If this entry is less than #11, give it an accelerator as well
            If i < 10 Then FormMain.mnuRecDocs(i).Caption = FormMain.mnuRecDocs(i).Caption & vbTab & g_Language.TranslateMessage("Ctrl") & "+" & i
            
        Next i
    End If
    
    'The icons in the MRU sub-menu need to be reset after this action
    Icon_and_Cursor_Handler.resetMenuIcons
End Sub

Private Sub IMRUList_BeforeListLoad(manager As pdMRUManager, Optional ByRef cancel As Boolean = False)
    'Reset the MRU hash storage.  Hashes are used to save recent file thumbnails (for use in the Recent Files menu),
    ' without revealing the actual filenames, which may contain sensitive information.
    ReDim m_mruHashes(0) As mruHash
    m_numOfMRUHashes = 0
End Sub



Private Sub IMRUList_OnItemLoad(manager As pdMRUManager, Index As Long, mruCaption As String)
    'Note that we do not need to load mnuRecDocs(0) - it is always loaded.  Only load menu items greater than this.
    If Index > 0 Then
        Load FormMain.mnuRecDocs(Index)
    Else
        FormMain.mnuRecDocs(Index).Enabled = True
    End If
    
    FormMain.mnuRecDocs(Index).Caption = mruCaption
    
    'For the first 10 menu entries, display a corresponding shortcut
    If Index < 10 Then FormMain.mnuRecDocs(Index).Caption = FormMain.mnuRecDocs(Index).Caption & vbTab & g_Language.TranslateMessage("Ctrl") & "+" & Index
End Sub

Private Sub IMRUList_AfterListLoaded(manager As pdMRUManager)
    'Make sure the "Clear MRU" option is visible and enabled
    FormMain.MnuRecentSepBar1.Visible = True
    FormMain.MnuLoadAllMRU.Visible = True
    FormMain.MnuClearMRU.Visible = True
End Sub

Private Sub IMRUList_OnListIsEmpty()
    FormMain.mnuRecDocs(0).Caption = g_Language.TranslateMessage("Empty")
    FormMain.mnuRecDocs(0).Enabled = False
    FormMain.MnuRecentSepBar1.Visible = False
    FormMain.MnuClearMRU.Visible = False
    FormMain.MnuLoadAllMRU.Visible = False
End Sub

Private Sub IMRUList_OnMaxLimitChanged(manager As pdMRUManager)
    Icon_and_Cursor_Handler.resetMenuIcons
End Sub


Private Sub IMRUList_TrimListAtPosition(manager As pdMRUManager, Index As Long)
    Dim thumbnailpath As String
    thumbnailpath = getMRUThumbnailPath(manager.getSpecificMRU(Index))
    If FileExist(thumbnailpath) Then Kill thumbnailpath
End Sub

Private Property Get IMRUList_XMLFileName() As String
    IMRUList_XMLFileName = g_UserPreferences.getPresetPath & "Program_RecentFiles.xml"
End Property

