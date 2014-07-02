VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdRecentFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'MRU (Most Recently Used) List Handler
'Copyright ©2005-2014 by Tanner Helland
'Created: 22/May/05
'Last updated: 20/May/14
'Last update: replace custom PNG export code with a newly created "QuickSaveDIBAsPNG" function in the Saving module.
'
'This class is responsible for the creation and maintenance of the program's Recent Files list, which is commonly
' referred to as an MRU or "Most Recently Used" list.
'
'Previously, recent file data was stored in PhotoDemon's main user preferences file.  As of 31 Oct '13 this data is
' now stored in its own XML file in the /Data/Presets subfolder.  This makes it much easier to dynamically change
' the max recent file count, as well as perform MRU maintenance (like clearing the list).  Users who upgrade to this
' new system will lose their old MRU data, I'm afraid.
'
'Thumbnails are automatically saved (in PNG format) for each entry in the recent files list.  Thumbnail filenames are
' hashed as a privacy precaution.  PhotoDemon automatically erases unneeded thumbnail files as files are "pushed" off
' the recent file stack, and as an extra precaution, when the program exits, it scans the thumbnail directory for any
' orphaned thumbnails.  (Orphaned thumbnails are possible if you run multiple copies of PD in parallel.)
'
'Special thanks to Randy Birch for the original version of the path shrinking code used in this class. You can
' download his original code from this link (good as of Nov 2012):
' http://vbnet.mvps.org/index.html?code/fileapi/pathcompactpathex.htm
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'MRUlist will contain string entries of all the most recently used files
Private MRUlist() As String

'Current number of entries in the MRU list
Private numEntries As Long

'These functions are used to shrink long path entries to a max number of characters
Private Declare Function PathCompactPathEx Lib "shlwapi" Alias "PathCompactPathExA" (ByVal pszOut As String, ByVal pszSrc As String, ByVal cchMax As Long, ByVal dwFlags As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Const MAX_PATH As Long = 260
Private Const maxMRULength As Long = 64

'Because we need to hash MRU names to generate icon save locations, and hashing is computationally expensive, store all
' calculated hashes in a table.
Private Type mruHash
    mruInitPath As String
    mruHashPath As String
End Type

Private mruHashes() As mruHash
Private numOfMRUHashes As Long

'XML handling (used to save/load actual MRU entries) is handled through a specialized class
Private xmlEngine As pdXML

'The file where we'll store the MRU data when the program is closed.  This file will be saved in the
' /Data/Presets/ folder.
Private xmlFilename As String

'Return the path to an MRU thumbnail file (in PNG format)
Public Function getMRUThumbnailPath(ByVal mruIndex As Long) As String
    If numEntries > 0 Then
        If (mruIndex >= 0) And (mruIndex <= numEntries) Then
            getMRUThumbnailPath = g_UserPreferences.getIconPath & getMRUHash(MRUlist(mruIndex)) & ".png"
        Else
            getMRUThumbnailPath = ""
        End If
    Else
        getMRUThumbnailPath = ""
    End If
End Function

Private Function doesMRUHashExist(ByVal FilePath As String) As String

    'Check to see if this file has been requested before.  If it has, return our previous
    ' hash instead of recalculating one from scratch.  If it does not exist, return "".
    If numOfMRUHashes > 0 Then
    
        'Loop through all previous hashes from this session
        Dim i As Long
        For i = 0 To numOfMRUHashes - 1
        
            'If this file path matches one we've already calculated, return that instead of calculating it again
            If StrComp(mruHashes(i).mruInitPath, FilePath, vbTextCompare) = 0 Then
                doesMRUHashExist = mruHashes(i).mruHashPath
                Exit Function
            End If
        
        Next i
    
    End If
    
    doesMRUHashExist = ""

End Function

'Return a 16-character hash of a specific MRU entry.  (This is used to generate unique menu icon filenames.)
Private Function getMRUHash(ByVal FilePath As String) As String
    
    'Check to see if this hash already exists
    Dim prevHash As String
    prevHash = doesMRUHashExist(FilePath)
    
    'If it does, return it.
    If Len(prevHash) > 0 Then
        getMRUHash = prevHash
        Exit Function
    
    'If no correlating hash was found, calculate one from scratch.
    Else
    
        'Prepare an SHA-256 hash calculator
        Dim cSHA2 As CSHA256
        Set cSHA2 = New CSHA256
            
        Dim hString As String
        hString = cSHA2.SHA256(FilePath)
                
        'The SHA-256 function returns a 64 character string (256 / 8 = 32 bytes, but 64 characters due to hex representation).
        ' This is too long for a filename, so take only the first sixteen characters of the hash.
        hString = Left$(hString, 16)
        
        'Save this hash to our hashes array
        mruHashes(numOfMRUHashes).mruInitPath = FilePath
        mruHashes(numOfMRUHashes).mruHashPath = hString
        numOfMRUHashes = numOfMRUHashes + 1
        ReDim Preserve mruHashes(0 To numOfMRUHashes) As mruHash
        
        'Return this as the hash value
        getMRUHash = hString
    
    End If
    
End Function

'Return the MRU entry at a specific location (used to load MRU files)
Public Function getSpecificMRU(ByVal mIndex As Long) As String
    If (mIndex <= numEntries) And (mIndex >= 0) Then
        getSpecificMRU = MRUlist(mIndex)
    Else
        getSpecificMRU = ""
    End If
End Function

'Load the MRU list from file (if it exists)
Public Sub MRU_LoadFromFile()

    'Start by seeing if an XML file with previously saved MRU data exists
    If FileExist(xmlFilename) Then
        
        'Attempt to load and validate the current file; if we can't, create a new, blank XML object
        If Not xmlEngine.loadXMLFile(xmlFilename) Then
            Message "List of recently opened images seems to be invalid.  A new recent images list has been created."
            resetXMLData
        End If
        
    Else
        resetXMLData
    End If
    
    'Reset the MRU hash storage.  Hashes are used to save recent file thumbnails (for use in the Recent Files menu),
    ' without revealing the actual filenames, which may contain sensitive information.
    ReDim mruHashes(0) As mruHash
    numOfMRUHashes = 0
    
    'We are now ready to load the actual MRU data from file.
    
    'The XML engine will do most the heavy lifting for this task.  We pass it a String array, and it fills it with
    ' all values corresponding to the given tag name and attribute.  (We must do this dynamically, because we don't
    ' know how many recent filenames are actually saved - it could be anywhere from 0 to RECENT_FILE_COUNT.)
    Dim allRecentFiles() As String
    If xmlEngine.findAllAttributeValues(allRecentFiles, "mruEntry", "id") Then
        numEntries = UBound(allRecentFiles) + 1
        
        'Make sure the file does not contain more entries than are allowed (shouldn't theoretically be possible,
        ' but it doesn't hurt to check).
        If numEntries > g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10) Then
            numEntries = g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10)
        End If
        
    'No recent file entries were found.  Adjust the Recent Files menu to match
    Else
        numEntries = 0
        FormMain.mnuRecDocs(0).Caption = g_Language.TranslateMessage("Empty")
        FormMain.mnuRecDocs(0).Enabled = False
        FormMain.MnuRecentSepBar1.Visible = False
        FormMain.MnuClearMRU.Visible = False
        FormMain.MnuLoadAllMRU.Visible = False
    End If
    
    'If one or more recent file entries were found, load them now.
    If numEntries > 0 Then
    
        'Remember that the allRecentFiles array was filled with the id values of all "mruEntry" entries in the XML file.
        ' We can use its upper bound to determine how many MRU entries are actually present in the file
        numEntries = UBound(allRecentFiles) + 1
        ReDim MRUlist(0 To numEntries) As String
        
        'Load the actual file paths from the MRU file
        Dim i As Long
        For i = 0 To numEntries - 1
            MRUlist(i) = xmlEngine.getUniqueTag_String("filePath", , , "mruEntry", "id", allRecentFiles(i))
            
            'Note that we do not need to load mnuRecDocs(0) - it is always loaded.  Only load menu items greater than this.
            If i > 0 Then
                Load FormMain.mnuRecDocs(i)
            Else
                FormMain.mnuRecDocs(i).Enabled = True
            End If
            
            'Based on the user's preference for captioning, display either the full path or just the filename
            If g_UserPreferences.GetPref_Long("Interface", "MRU Caption Length", 0) = 0 Then
                FormMain.mnuRecDocs(i).Caption = getFilename(MRUlist(i))
            Else
                FormMain.mnuRecDocs(i).Caption = getShortMRU(MRUlist(i))
            End If
            
            'For the first 10 menu entries, display a corresponding shortcut
            If i < 10 Then FormMain.mnuRecDocs(i).Caption = FormMain.mnuRecDocs(i).Caption & vbTab & "Ctrl+" & i
            
        Next i
    
        'Finally, make sure the "Clear MRU" option is visible and enabled
        FormMain.MnuRecentSepBar1.Visible = True
        FormMain.MnuLoadAllMRU.Visible = True
        FormMain.MnuClearMRU.Visible = True
     
    End If
        
End Sub

'If the user changes their preference regarding the number of recent files we can save, call this sub to rebuild
' the current menu.
Public Sub MRU_NotifyNewMaxLimit()
    
    'Erase any entries above the new limit
    If numEntries > g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10) Then
        numEntries = g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10)
    End If
    
    'Write the current MRU list out to file.
    MRU_SaveToFile
    
    'Unload all recent file menu entries
    If FormMain.mnuRecDocs.Count > 1 Then
        Dim i As Long
        For i = FormMain.mnuRecDocs.Count - 1 To 1 Step -1
            Unload FormMain.mnuRecDocs(i)
        Next i
    End If
    
    'Reload MRU data from file, and refresh any menu icons
    MRU_LoadFromFile
    resetMenuIcons
    
End Sub

'Save the current MRU list to file (currently done at program close)
Public Sub MRU_SaveToFile()

    'Reset whatever XML data we may have stored at present - we will be rewriting the full MRU file from scratch.
    resetXMLData
    
    Dim i As Long
    
    'Only write new entries if MRU data exists for them
    If numEntries > 0 Then
        For i = 0 To numEntries - 1
            xmlEngine.writeTagWithAttribute "mruEntry", "id", Str(i), "", True
            xmlEngine.writeTag "filePath", MRUlist(i)
            xmlEngine.closeTag "mruEntry"
            xmlEngine.writeBlankLine
        Next i
    End If
    
    'With the XML file now complete, write it out to file
    xmlEngine.writeXMLToFile xmlFilename
    
    'Unload all corresponding menu entries.  (This doesn't matter when the program is closing, but we also use this
    ' routine to refresh the MRU list after changing the caption preference - and for that an unload is required.)
    If numEntries > 1 Then
        For i = FormMain.mnuRecDocs.Count - 1 To 1 Step -1
            Unload FormMain.mnuRecDocs(i)
        Next i
    End If
    
    'Finally, scan the MRU icon directory to make sure there are no orphaned PNG files.  (Multiple instances of PhotoDemon
    ' running simultaneously can lead to this.)  Delete any PNG files that don't correspond to current MRU entries.
    Dim chkFile As String
    chkFile = Dir(g_UserPreferences.getIconPath & "*.png", vbNormal)
    
    Dim fileOK As Boolean
    
    Do While Len(chkFile) > 0
        
        fileOK = False
        
        'Compare this file to the hash for all current MRU entries
        If numEntries > 0 Then
            For i = 0 To numEntries - 1
                
                'If this hash matches one on file, mark it as OK.
                If StrComp(g_UserPreferences.getIconPath & chkFile, getMRUThumbnailPath(i), vbTextCompare) = 0 Then
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

'Add another file to the MRU list
Public Sub MRU_AddNewFile(ByVal newFile As String, ByRef srcImage As pdImage)

    'Locators are used to determine if this file already exists in the recent files list.
    ' If it does, we will simply shuffle its position instead of adding it as a new entry.
    Dim alreadyThere As Boolean
    alreadyThere = False
    
    Dim curLocation As Long
    curLocation = -1
    
    Dim i As Long
    
    'First, check to see if this file currently exists in the MRU list
    For i = 0 To numEntries - 1
    
        'This file already exists in the list!  Make a note of its location, then exit.
        If MRUlist(i) = newFile Then
            alreadyThere = True
            curLocation = i
            Exit For
        End If
        
    Next i
        
    'If the file already exists in the recent files list, shuffle its position instead of adding it as a new entry.
    If alreadyThere Then
        
        'If this file is already the most recent file (position 0), we don't need to do anything - but if it appears
        ' elsewhere in the list, shift everything after its position downward.
        If curLocation > 0 Then
            For i = curLocation To 1 Step -1
                MRUlist(i) = MRUlist(i - 1)
            Next i
        End If
    
    'This file doesn't exist in the MRU list, so it must be added at the very top as a new entry.
    Else

        numEntries = numEntries + 1
        
        'Cap the number of MRU files at a certain value (specified by the user in the Preferences menu)
        If numEntries > g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10) Then
            numEntries = g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10)
            
            'Also, because we are about to purge the MRU list, we need to delete the last entry's image (if it exists).
            ' If we don't do this, the icons directory will eventually fill up with icons of old files.
            If FileExist(getMRUThumbnailPath(numEntries - 1)) Then Kill getMRUThumbnailPath(numEntries - 1)
        End If
        
        'Resize the list of MRU entries, which may have grown on account of this new addition.
        ReDim Preserve MRUlist(0 To numEntries) As String
    
        'Shift all existing entries downward
        If numEntries > 1 Then
            For i = numEntries To 1 Step -1
                MRUlist(i) = MRUlist(i - 1)
            Next i
        End If
        
    End If
    
    'Add this entry to the list
    MRUlist(0) = newFile
    
    'Save a thumbnail of this image to file.
    saveMRUThumbnail newFile, srcImage
    
    'Redraw the MRU menu based on the updated list
    If Not FormMain.mnuRecDocs(0).Enabled Then
        FormMain.mnuRecDocs(0).Enabled = True
        FormMain.MnuRecentSepBar1.Visible = True
        FormMain.MnuLoadAllMRU.Visible = True
        FormMain.MnuClearMRU.Visible = True
    End If
    
    'Based on the user's preference, display just the filename or the entire file path (up to the max character length)
    If g_UserPreferences.GetPref_Long("Core", "MRU Caption Length", 0) = 0 Then
        FormMain.mnuRecDocs(0).Caption = getFilename(newFile)
    Else
        FormMain.mnuRecDocs(0).Caption = getShortMRU(newFile)
    End If
    
    'Display the corresponding accelerator (hotkey).
    FormMain.mnuRecDocs(0).Caption = FormMain.mnuRecDocs(0).Caption & vbTab & "Ctrl+0"
    
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
            If g_UserPreferences.GetPref_Long("Core", "MRU Caption Length", 0) = 0 Then
                FormMain.mnuRecDocs(i).Caption = getFilename(MRUlist(i))
            Else
                FormMain.mnuRecDocs(i).Caption = getShortMRU(MRUlist(i))
            End If
            
            'If this entry is less than #11, give it an accelerator as well
            If i < 10 Then FormMain.mnuRecDocs(i).Caption = FormMain.mnuRecDocs(i).Caption & vbTab & "Ctrl+" & i
            
        Next i
    End If
    
    'The icons in the MRU sub-menu need to be reset after this action
    resetMenuIcons

End Sub

'Saves a thumbnail PNG of a pdImage object.  The thumbnail is saved to the /Data/Icons directory
Private Sub saveMRUThumbnail(ByRef imagePath As String, ByRef tImage As pdImage)

    Message "Saving recent file menu thumbnail..."
    
    'First, generate a path at which to save the file in question
    Dim sFilename As String
    sFilename = g_UserPreferences.getIconPath & getMRUHash(imagePath) & ".png"
    
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

'Empty the entire MRU list and clear the menu of all entries
Public Sub MRU_ClearList()
    
    Dim i As Long
    
    'Delete all recent file menu entries except the last one (which we will use to display "no files found")
    If FormMain.mnuRecDocs.Count > 1 Then
        For i = FormMain.mnuRecDocs.Count - 1 To 1 Step -1
            Unload FormMain.mnuRecDocs(i)
        Next i
    End If
    
    'Disable the lone remaining menu entry, and hide the option to clear the recent files list.
    FormMain.mnuRecDocs(0).Caption = g_Language.TranslateMessage("Empty")
    FormMain.mnuRecDocs(0).Enabled = False
    FormMain.MnuRecentSepBar1.Visible = False
    FormMain.MnuLoadAllMRU.Visible = False
    FormMain.MnuClearMRU.Visible = False
    
    'If any recent file thumbnails still exist, delete them.
    Dim tmpFilename As String
    
    For i = 0 To numEntries
        tmpFilename = getMRUThumbnailPath(i)
        If FileExist(tmpFilename) Then Kill tmpFilename
    Next i
    
    'Reset the number of entries in the MRU list
    numEntries = 0
    ReDim MRUlist(0) As String
    
    'Erase any existing XML file, and reset the in-memory version
    If FileExist(xmlFilename) Then Kill xmlFilename
    resetXMLData
    
    'The icons in the MRU sub-menu also need to be reset after this action; call that function now.
    resetMenuIcons

End Sub

'Return how many MRU entries are currently in the menu
Public Function MRU_ReturnCount() As Long
    MRU_ReturnCount = numEntries
End Function

'Truncates a path to a specified number of characters by replacing path components with ellipses.
' (Originally written by Randy Birch @ http://vbnet.mvps.org/index.html?code/fileapi/pathcompactpathex.htm)
Private Function getShortMRU(ByVal sPath As String) As String

    Dim ret As Long
    Dim buff As String
      
    buff = Space$(MAX_PATH)
    ret = PathCompactPathEx(buff, sPath, maxMRULength + 1, 0&)
   
    getShortMRU = TrimNull(buff)
   
End Function

'Remove null characters from a string
Private Function TrimNull(ByVal sString As String) As String
   TrimNull = Left$(sString, lstrlenW(StrPtr(sString)))
End Function

Private Sub Class_Initialize()
    
    'Initialize an XML engine, which we will use to read/write our MRU data to file
    Set xmlEngine = New pdXML
    
    'The XML data will be stored in the Preset path (/Data/Presets)
    xmlFilename = g_UserPreferences.getPresetPath & "Program_RecentFiles.xml"
    
    'If an XML file exists, it will be loaded separately, by the MRU_LoadFromFile() function
    
End Sub

'When creating a new MRU file, or overwriting a corrupt one, use this to initialize the XML header.
Private Sub resetXMLData()

    xmlEngine.prepareNewXML "Recent files"
    xmlEngine.writeBlankLine
    xmlEngine.writeComment "Everything past this point is recent file data.  Entries are sorted in reverse chronological order."
    xmlEngine.writeBlankLine

End Sub
