VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdMRUManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Generic Recent Files (MRU) List Manager
'Copyright 2005-2026 by Raj Chaudhuri and Tanner Helland
'Created: 16/February/15
'Last updated: 17/February/15
'Last updated by: Raj
'Last update: Trimmed fileName in MRU_AddNewFile, and deleted old XML file if Max limit changed.
'
'This class is responsible for the creation and maintenance of MRU or "Most Recently Used" lists.
' It contains functionality for saving and loading MRU lists to files, and updating a single
' menu array. The functionality can be further specialized if required.
'
'Many thanks to Raj Chaudhuri for rewriting this class as a generic manager, so PD can now support
'crecent file lists in multiple places throughout the program.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Long filenames (and/or paths, if the user selects that option in Tools > Preference) are automatically
' shrunk using PathCompact; this constant controls the maximum length, in characters.
Private Const MAX_MRU_LENGTH_IN_CHARS As Long = 64

'MRUlist will contain string entries of all the most recently used files
Private m_FileList() As String

'Current number of entries in the MRU list
Private m_NumEntries As Long

'PD's standard XML engine is used to write the recent file lists out to file
Private m_XMLEngine As pdXML

'The file where this class persistently stores its data.  In PD, the path of this file is controlled by
' the preferences manager (UserPrefs)
Private m_XMLFilename As String

'This class relies on a child class that implements that the IMRUList interface.
Private m_MRUList As IMRUList

'When creating a new MRU file, or overwriting a corrupt one, use this to initialize the XML header.
Private Sub ResetXMLData()
    m_XMLEngine.PrepareNewXML "Recent files"
    m_XMLEngine.WriteBlankLine
    m_XMLEngine.WriteComment "Everything past this point is recent file data.  Entries are sorted in reverse chronological order."
    m_XMLEngine.WriteBlankLine
End Sub

'Return the menu caption of the current recent file entry.  At present, menu caption length is controlled by a global preference
Private Function GetMRUCaption(ByRef newFile As String) As String
    
    'Based on the user's preference, display just the filename or the entire file path (up to the max character length)
    If (UserPrefs.GetPref_Long("Core", "MRU Caption Length", 0) = 0) Then
        GetMRUCaption = Files.FileGetName(newFile)
    Else
        GetMRUCaption = Files.PathCompact(newFile, MAX_MRU_LENGTH_IN_CHARS)
    End If
    
End Function

'Return the actual file path at a given index
Public Function GetSpecificMRU(ByVal mIndex As Long) As String
    If (mIndex >= 0) And (mIndex <= m_NumEntries) Then
        GetSpecificMRU = m_FileList(mIndex)
    Else
        GetSpecificMRU = vbNullString
    End If
End Function

'Return a menu-friendly caption of a given index
Public Function GetSpecificMRUCaption(ByVal mIndex As Long) As String
    If (mIndex >= 0) And (mIndex <= m_NumEntries) Then
        GetSpecificMRUCaption = GetMRUCaption(m_FileList(mIndex))
    Else
        GetSpecificMRUCaption = vbNullString
    End If
End Function

'Return the path to an MRU thumbnail file (if one exists; PD does not currently support this for macros)
Public Function GetMRUThumbnailPath(ByVal mIndex As Long) As String
    
    GetMRUThumbnailPath = vbNullString
    
    'Query the child to see if it supports thumbnail images
    If m_MRUList.SupportsThumbnails Then
        
        'Ignore thumbnail requests if the current list is empty
        If (m_NumEntries > 0) Then
            If (mIndex >= 0) And (mIndex <= m_NumEntries) Then GetMRUThumbnailPath = m_MRUList.GetThumbnailPath(Me, mIndex)
        End If
    
    End If
    
End Function

'Load the MRU list from file (if it exists)
Public Sub MRU_LoadFromFile()

    'Start by seeing if an XML file with previously saved MRU data exists
    If Files.FileExists(m_XMLFilename) Then
        
        'Attempt to load and validate the current file; if we can't, create a new, blank XML object
        If (Not m_XMLEngine.LoadXMLFile(m_XMLFilename)) Then
            PDDebug.LogAction "The Recent Files XML data at " & m_XMLFilename & " didn't validate.  Creating a new list now..."
            ResetXMLData
        End If
        
    Else
        ResetXMLData
    End If
    
    'Allow the child to run any required initialization steps
    Dim childCanceled As Boolean
    childCanceled = False
    m_MRUList.BeforeListLoad Me, childCanceled
    
    'If something goes wrong, the child is allowed to abort the load process.  Do not proceed if cancellation was requested.
    If childCanceled Then Exit Sub
       
    'We are now ready to load the actual MRU data from file.
    
    'The XML engine will do most the heavy lifting for this task.  We pass it a String array, and it fills it with
    ' all values corresponding to the given tag name and attribute.  (We must do this dynamically, because we don't
    ' know how many recent filenames are actually saved - it could be anywhere from 0 to RECENT_FILE_COUNT.)
    Dim allRecentFiles() As String
    If m_XMLEngine.FindAllAttributeValues(allRecentFiles, "mruEntry", "id") Then
        
        m_NumEntries = UBound(allRecentFiles) + 1
        
        'Make sure the file does not contain more entries than are allowed (shouldn't theoretically be possible,
        ' but it doesn't hurt to check).
        If (m_NumEntries > UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)) Then
            m_NumEntries = UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)
        End If
        
    'No recent file entries were found.  Adjust the Recent Files menu to match
    Else
        m_NumEntries = 0
        m_MRUList.OnListIsEmpty
    End If
    
    'If one or more recent file entries were found, load them now.
    If (m_NumEntries > 0) Then
    
        'Prepare our internal file list
        ReDim m_FileList(0 To m_NumEntries) As String
        
        'Load the actual file paths from the MRU file
        Dim i As Long
        For i = 0 To m_NumEntries - 1
        
            m_FileList(i) = m_XMLEngine.GetUniqueTag_String("filePath", , , "mruEntry", "id", allRecentFiles(i))
            
            'Let the child do any UI preparation required on a per-item basis
            m_MRUList.OnItemLoad Me, i, GetMRUCaption(m_FileList(i))
            
        Next i
        
        'Let the child perform an UI preparation after the whole list has been loaded
        m_MRUList.AfterListLoad Me
        
    End If
    
End Sub

'Save the current MRU list to file (currently done at program close)
Public Sub MRU_SaveToFile()

    Dim saveCancel As Boolean
    saveCancel = False
    
    'Allow the child to do any necessary prep work.  They can cancel this operation; if they do, the file will not be saved.
    m_MRUList.BeforeListSave Me, saveCancel
    If saveCancel Then Exit Sub
    
    'Reset whatever XML data we may have stored at present - we will be rewriting the full MRU file from scratch.
    ResetXMLData
    
    'Only write new entries if MRU data exists for them
    If (m_NumEntries > 0) Then
    
        Dim i As Long
        For i = 0 To m_NumEntries - 1
            m_XMLEngine.WriteTagWithAttribute "mruEntry", "id", CStr(i), vbNullString, True
            m_XMLEngine.WriteTag "filePath", m_FileList(i)
            m_XMLEngine.CloseTag "mruEntry"
            m_XMLEngine.WriteBlankLine
        Next i
        
    End If
    
    'With the XML file now complete, write it out to file
    m_XMLEngine.WriteXMLToFile m_XMLFilename
    
    'Unload all corresponding menu entries.  (This doesn't matter when the program is closing, but we also use this
    ' routine to refresh the MRU list after changing the caption preference - and for that an unload is required.)
    m_MRUList.AfterListCleared
    
    'Allow the child to perform any post-save cleanup
    m_MRUList.AfterListSave Me
    
End Sub

'Add another file to the MRU list
Public Sub MRU_AddNewFile(ByVal newFile As String, Optional ByRef srcImage As pdImage = Nothing)
    
    'The filename is sometimes passed via dialogs, and as such may contain nulls
    newFile = Strings.TrimNull(newFile)
    
    'Allow the child to perform any necessary UI prep work.  The child also has the option to cancel this operation.
    Dim childCancel As Boolean
    childCancel = False
    m_MRUList.BeforeFileAdded Me, newFile, childCancel
    If childCancel Then Exit Sub

    'Locators are used to determine if this file already exists in the recent files list.
    ' If it does, we will simply shuffle its position instead of adding it as a new entry.
    Dim alreadyThere As Boolean
    alreadyThere = False
    
    Dim curLocation As Long
    curLocation = -1
    
    Dim i As Long
    
    'First, check to see if this file currently exists in the MRU list
    For i = 0 To m_NumEntries - 1
    
        'This file already exists in the list!  Make a note of its location, then exit.
        If Strings.StringsEqual(m_FileList(i), newFile, True) Then
            alreadyThere = True
            curLocation = i
            Exit For
        End If
        
    Next i
        
    'If the file already exists in the recent files list, shuffle its position instead of adding it as a new entry.
    If alreadyThere Then
        
        'If this file is already the most recent file (position 0), we don't need to do anything - but if it appears
        ' elsewhere in the list, shift everything after its position downward.
        If (curLocation > 0) Then
            For i = curLocation To 1 Step -1
                m_FileList(i) = m_FileList(i - 1)
            Next i
        End If
    
    'This file doesn't exist in the MRU list, so it must be added at the very top as a new entry.
    Else

        m_NumEntries = m_NumEntries + 1
        
        'Cap the number of MRU files at a certain value (specified by the user in the Preferences menu)
        If (m_NumEntries > UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)) Then
            
            m_NumEntries = UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)
            
            'Also, because we are about to purge the MRU list, the last entry may require cleanup
            m_MRUList.EraseEntryAtPosition Me, m_NumEntries - 1
            
        End If
        
        'Resize the list of MRU entries, which may have grown on account of this new addition.
        ReDim Preserve m_FileList(0 To m_NumEntries) As String
    
        'Shift all existing entries downward
        If (m_NumEntries > 1) Then
            For i = m_NumEntries To 1 Step -1
                m_FileList(i) = m_FileList(i - 1)
            Next i
        End If
        
    End If
    
    'Add this entry to the top of the list
    m_FileList(0) = newFile
    
    'Save a thumbnail of this image to file.
    If m_MRUList.SupportsThumbnails Then m_MRUList.SaveThumbnailImage newFile, srcImage
      
    'Based on the user's preference, display just the filename or the entire file path (up to the max character length).
    ' Note that the child handles the actual UI work involved.
    m_MRUList.AfterFileAdded Me, newFile, GetMRUCaption(newFile)
    
End Sub

'If the user changes their preference regarding the number of recent files we can save, call this sub to rebuild
' the current menu.
Public Sub MRU_NotifyNewMaxLimit()
    
    'Erase any entries above the new limit
    If (m_NumEntries > UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)) Then
    
        m_NumEntries = UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)
        
        ' Delete the old MRU file now, because we need to truncate it
        Files.FileDeleteIfExists m_XMLFilename
        
    End If
    
    'Write the current MRU list out to file.
    MRU_SaveToFile
    
    'Unload all recent file menu entries
    m_MRUList.AfterListCleared
    
    'Reload MRU data from file
    MRU_LoadFromFile
    
    'Allow the child perform any related clean-up
    m_MRUList.OnMaxLimitChanged Me
    
End Sub

'Empty the entire MRU list and clear the menu of all entries
Public Sub MRU_ClearList()
    
    'Reset the number of entries in the MRU list
    m_NumEntries = 0
    ReDim m_FileList(0) As String
    
    'Erase any existing XML file, and reset the in-memory version
    Files.FileDeleteIfExists m_XMLFilename
    ResetXMLData
    
    'Allow the child to perform any related UI clean-up
    m_MRUList.AfterListCleared
    m_MRUList.OnListIsEmpty
    
End Sub

'Return the number of MRU entries currently loaded and active
Public Function MRU_ReturnCount() As Long
    MRU_ReturnCount = m_NumEntries
End Function

'Perform any pre-load initiation tasks
Public Sub InitList(specificList As IMRUList)

    'Set a pointer to the child list UI handler
    Set m_MRUList = specificList
    
    'Initialize an XML engine, which we will use to read/write our MRU data to file
    Set m_XMLEngine = New pdXML
    
    'The XML data will be stored in the Preset path (/Data/Presets), but we simply read it from the child's matching property
    m_XMLFilename = m_MRUList.XMLFilename
    
    'If an XML file exists, it will be loaded separately, by the MRU_LoadFromFile() function
    
End Sub
