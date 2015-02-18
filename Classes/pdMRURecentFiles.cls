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
'Recent Files (MRU) handler for PD's primary image list (File > Open Recent)
'Copyright 2005-2015 by Raj Chaudhuri and Tanner Helland
'Created: 16/February/15
'Last updated: 17/February/15
'Last updated by: Raj
'Last update: Corrected IMRUList_GetThumbnailPath to return String
'Dependencies: - everything listed in IMRUList
'              - pdFSO (for simpler Unicode file interactions)
'
'This class is a specific implementation of the IMRUList Interface.  It works alongside pdMRUManager to manage the UI for
' PD's primary recent file list (File > Open Recent menu).
'
'In 2015, Raj Chaudhuri kindly rewrote this class as a replacement for PD's previous pdRecentFiles class.  Data written by
' the old class is still compatible with this implementation.  Many thanks to Raj for this great improvement!  (Thanks also
' to Raj for rewriting this class as an implementation of IMRUList, so PD can now support recent file lists in multiple
' places throughout the program.)
'
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

'Because we need to hash MRU names to generate icon save locations, and hashing is computationally expensive, we store all
' calculated hashes in a table.
Private Type MRUEntry
    actualFilePath As String
    hashedFilename As String
End Type

Private m_Entries() As MRUEntry
Private m_NumOfEntries As Long

'Unicode file interactions are made easier by using a pdFSO instance
Private cFile As pdFSO

'If a hash currently exists for a given file, return it so we can avoid hashing it again
Private Function doesMRUHashExist(ByVal filePath As String) As String

    'Check to see if this file has been requested before.  If it has, return our previous
    ' hash instead of recalculating one from scratch.  If it does not exist, return "".
    If m_NumOfEntries > 0 Then
    
        'Loop through all previous hashes from this session
        Dim i As Long
        For i = 0 To m_NumOfEntries - 1
        
            'If this file path matches one we've already calculated, return that instead of calculating it again
            If StrComp(m_Entries(i).actualFilePath, filePath, vbTextCompare) = 0 Then
                doesMRUHashExist = m_Entries(i).hashedFilename
                Exit Function
            End If
        
        Next i
    
    End If
    
    'If we made it all the way here, this file does not exist in the current list.  A new hash must be generated.
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
            
        'Retrieve a hash of this filename
        Dim hString As String
        hString = cSHA2.SHA256(filePath)
                
        'The SHA-256 function returns a 64 character string (256 / 8 = 32 bytes, but 64 characters due to hex representation).
        ' This is too long for a filename, so take only the first sixteen characters of the hash.
        hString = Left$(hString, 16)
        
        'Save this hash so we don't have to calculate it again in the future
        m_Entries(m_NumOfEntries).actualFilePath = filePath
        m_Entries(m_NumOfEntries).hashedFilename = hString
        m_NumOfEntries = m_NumOfEntries + 1
        ReDim Preserve m_Entries(0 To m_NumOfEntries) As MRUEntry
        
        'Return the calculated hash
        getMRUHash = hString
    
    End If
    
End Function

'Given a file path, return the matching thumbnail path.  (PD stores all thumbnails to the same subfolder, currently /Data/Icons.)
Private Function getMRUThumbnailPath(ByVal imagePath As String) As String
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

Private Sub Class_Initialize()

    'Initialize our pdFSO instance
    Set cFile = New pdFSO

End Sub

Private Property Get IMRUList_SupportsThumbnails() As Boolean
    IMRUList_SupportsThumbnails = True
End Property

Private Function IMRUList_GetThumbnailPath(ByRef mruManager As pdMRUManager, ByVal Index As Long) As String
    IMRUList_GetThumbnailPath = getMRUThumbnailPath(mruManager.getSpecificMRU(Index))
End Function

Private Sub IMRUList_SaveThumbnailImage(ByRef filePath As String, ByRef srcImage As pdImage)
    saveMRUThumbnail filePath, srcImage
End Sub

'When a list clear is requested, this class needs to perform potential maintenance on the File > Open Recent submenu
Private Sub IMRUList_AfterListCleared()

    'Unload all recent file menu entries
    If FormMain.mnuRecDocs.Count > 1 Then
        
        Dim i As Long
        For i = FormMain.mnuRecDocs.Count - 1 To 1 Step -1
            Unload FormMain.mnuRecDocs(i)
        Next i
        
    End If
    
End Sub

'After saving this master recent file list (typically done at program termination), scan the icon folder and look for orphaned
' PNG files.  This will only happen after program crashes, or when running multiple PD instances in parallel.
Private Sub IMRUList_AfterListSave(ByRef mruManager As pdMRUManager)

    'Use a pdFSO instance to enumerate PNG files for us.
    Dim pngList As pdStringStack
    
    If cFile.retrieveAllFiles(g_UserPreferences.getIconPath, pngList, False, False, "png") Then
    
        Dim chkFile As String, fileOK As Boolean
        Dim i As Long
        
        Dim numEntries As Long
        numEntries = mruManager.MRU_ReturnCount
        
        'Enumerate all PNGs in the folder
        Do While pngList.PopString(chkFile)
            
            'Compare this file to the hash for all current MRU entries
            If numEntries > 0 Then
                For i = 0 To numEntries - 1
                    
                    'If this hash matches one on file, mark it as OK.
                    If StrComp(chkFile, mruManager.getMRUThumbnailPath(i), vbTextCompare) = 0 Then
                        fileOK = True
                        Exit For
                    End If
                    
                Next i
            Else
                fileOK = False
            End If
        
            'If an MRU hash does not exist for this file, delete the file
            If Not fileOK Then cFile.KillFile chkFile
            
        Loop
    
    End If
    
End Sub

'When a new file is loaded into PD, we need to perform a bunch of maintenance on the File > Recent Files submenu
Private Sub IMRUList_AfterFileAdded(ByRef mruManager As pdMRUManager, ByVal newFilePath As String, ByVal mruCaption As String)
    
    'If the Open Recent submenu is currently empty, start by enabling all relevant menu objects
    If Not FormMain.mnuRecDocs(0).Enabled Then
        FormMain.mnuRecDocs(0).Enabled = True
        FormMain.MnuRecentSepBar1.Visible = True
        FormMain.MnuLoadAllMRU.Visible = True
        FormMain.MnuClearMRU.Visible = True
    End If
    
    'Add the filename and corresponding accelerator (hotkey) to the base menu
    FormMain.mnuRecDocs(0).Caption = mruCaption
    FormMain.mnuRecDocs(0).Caption = FormMain.mnuRecDocs(0).Caption & vbTab & g_Language.TranslateMessage("Ctrl") & "+0"
    
    'Ask the manager object how many entries it currently maintains
    Dim numEntries As Long
    numEntries = mruManager.MRU_ReturnCount()
    
    'If there are more entries on the list besides this first one, update all corresponding menus now
    If numEntries > 1 Then
    
        'Start by unloading any/all existing menus
        Dim i As Long
        For i = FormMain.mnuRecDocs.Count - 1 To 1 Step -1
            Unload FormMain.mnuRecDocs(i)
        Next i
        
        'Next, load all relevant menus in turn
        For i = 1 To numEntries - 1
        
            Load FormMain.mnuRecDocs(i)
            
            'Add the caption, and if the entry is less than #10, give it an accelerator, too
            FormMain.mnuRecDocs(i).Caption = mruManager.getSpecificMRUCaption(i)
            If i < 10 Then FormMain.mnuRecDocs(i).Caption = FormMain.mnuRecDocs(i).Caption & vbTab & g_Language.TranslateMessage("Ctrl") & "+" & i
            
        Next i
        
    End If
    
    'Note that thumbnail icons in the sub-menu need to be reset after this action; this is handled by PD's separate menu icon manager
    Icon_and_Cursor_Handler.resetMenuIcons
    
End Sub

Private Sub IMRUList_BeforeListLoad(ByRef mruManager As pdMRUManager, Optional ByRef Cancel As Boolean = False)
    
    'Reset our internal file and hash list
    ReDim m_Entries(0) As MRUEntry
    m_NumOfEntries = 0
    
End Sub

'When loading an entry from file, the setup process is somewhat more streamlined, as we don't need to unload existing
' menu entries.
Private Sub IMRUList_OnItemLoad(ByRef mruManager As pdMRUManager, ByVal Index As Long, ByVal mruCaption As String)
    
    'Note that we do not need to load mnuRecDocs(0) - it is always loaded.  We only ever load menu items greater than this.
    If Index > 0 Then
        Load FormMain.mnuRecDocs(Index)
    Else
        FormMain.mnuRecDocs(Index).Enabled = True
    End If
    
    FormMain.mnuRecDocs(Index).Caption = mruCaption
    
    'For the first 10 menu entries, display a corresponding shortcut
    If Index < 10 Then FormMain.mnuRecDocs(Index).Caption = FormMain.mnuRecDocs(Index).Caption & vbTab & g_Language.TranslateMessage("Ctrl") & "+" & Index
    
End Sub

Private Sub IMRUList_AfterListLoad(ByRef mruManager As pdMRUManager)
    
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

'Whenever the maximum file limit is changed, menu icons need to be reset
Private Sub IMRUList_OnMaxLimitChanged(Manager As pdMRUManager)
    Icon_and_Cursor_Handler.resetMenuIcons
End Sub

'Individual entries can be cleared via this function
Private Sub IMRUList_EraseEntryAtPosition(ByRef mruManager As pdMRUManager, ByVal Index As Long)
    
    'Make sure the requested index is valid
    If (Index >= 0) And (Index < m_NumOfEntries) Then
    
        'Kill the thumbnail at this position
        cFile.KillFile getMRUThumbnailPath(mruManager.getSpecificMRU(Index))
    
    End If
    
End Sub

'PD's main file list (File > Open Recent menu) is saved to the preset path with a hard-coded name
Private Property Get IMRUList_XMLFileName() As String
    IMRUList_XMLFileName = g_UserPreferences.getPresetPath & "Program_RecentFiles.xml"
End Property

'Unneeded interface methods are collected here
Private Sub IMRUList_BeforeListSave(Manager As pdMRUManager, Optional Cancel As Boolean = False)

End Sub

Private Sub IMRUList_BeforeFileAdded(ByRef mruManager As pdMRUManager, ByVal newFile As String, Optional ByRef Cancel As Boolean = False)

End Sub

