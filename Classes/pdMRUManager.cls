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
'Generic MRU (Most Recently Used) List Handler
'Copyright 2005-2015 by Tanner Helland
'Created: 16/February/15
'Last updated: 17/February/15
'Last updated by: Raj
'Last update: Trimmed fileName in MRU_AddNewFile, and deleted old XML file if Max limit changed.
'
'This class is responsible for the creation and maintenance of MRU or "Most Recently Used" lists.
'  It contains functionality for saving and loading MRU lists to files, and updating a single
'  menu array. The functionality can be further specialized if required.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************
Option Explicit


'These functions are used to shrink long path entries to a max number of characters
Private Declare Function PathCompactPathEx Lib "shlwapi" Alias "PathCompactPathExW" (ByVal pszOutPointer As Long, ByVal pszSrcPointer As Long, ByVal cchMax As Long, ByVal dwFlags As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Const MAX_PATH As Long = 260
Private Const maxMRULength As Long = 64

'MRUlist will contain string entries of all the most recently used files
Private m_fileList() As String

'Current number of entries in the MRU list
Private m_numEntries As Long

'XML handling (used to save/load actual MRU entries) is handled through a specialized class
Private m_xmlEngine As pdXML

'The file where we'll store the MRU data when the program is closed.
Private m_xmlfilename As String

Private m_MRUList As IMRUList


'When creating a new MRU file, or overwriting a corrupt one, use this to initialize the XML header.
Private Sub resetXMLData()
    m_xmlEngine.prepareNewXML "Recent files"
    m_xmlEngine.writeBlankLine
    m_xmlEngine.writeComment "Everything past this point is recent file data.  Entries are sorted in reverse chronological order."
    m_xmlEngine.writeBlankLine
End Sub

'Remove null characters from a string
Private Function TrimNullW(ByVal sString As String) As String
   TrimNullW = Left$(sString, lstrlenW(StrPtr(sString)))
End Function

'Truncates a path to a specified number of characters by replacing path components with ellipses.
' (Originally written by Randy Birch @ http://vbnet.mvps.org/index.html?code/fileapi/pathcompactpathex.htm)
Private Function getShortMRU(ByVal sPath As String) As String

    Dim ret As Long
    Dim buff As String
      
    buff = Space$(MAX_PATH)
    ret = PathCompactPathEx(StrPtr(buff), StrPtr(sPath), maxMRULength + 1, 0&)
   
    getShortMRU = TrimNullW(buff)
   
End Function

Private Function getMRUCaption(newFile As String) As String
    'Based on the user's preference, display just the filename or the entire file path (up to the max character length)
    If g_UserPreferences.GetPref_Long("Core", "MRU Caption Length", 0) = 0 Then
        getMRUCaption = File_And_Path_Handling.getFilename(newFile)
    Else
        getMRUCaption = getShortMRU(newFile)
    End If
End Function

Public Function getSpecificMRU(ByVal mIndex As Long) As String
    If (mIndex <= m_numEntries) And (mIndex >= 0) Then
        getSpecificMRU = m_fileList(mIndex)
    Else
        getSpecificMRU = ""
    End If
End Function

Public Function getSpecificMRUCaption(Index As Long) As String
    If (Index <= m_numEntries) And (Index >= 0) Then
        getSpecificMRUCaption = getMRUCaption(m_fileList(Index))
    Else
        getSpecificMRUCaption = ""
    End If
End Function

'Return the path to an MRU thumbnail file (in PNG format)
Public Function getMRUThumbnailPath(ByVal Index As Long) As String
    Dim result As String
    result = ""
    If m_MRUList.SupportsThumbNails Then
        If m_numEntries > 0 Then
            If (Index >= 0) And (Index <= m_numEntries) Then
                
                result = m_MRUList.GetThumbNailPath(Me, Index)
            End If
        End If
    End If
    getMRUThumbnailPath = result
End Function

'Load the MRU list from file (if it exists)
Public Sub MRU_LoadFromFile()

    'Start by seeing if an XML file with previously saved MRU data exists
    If FileExist(m_xmlfilename) Then
        
        'Attempt to load and validate the current file; if we can't, create a new, blank XML object
        If Not m_xmlEngine.loadXMLFile(m_xmlfilename) Then
            Message "List of recently opened images seems to be invalid.  A new recent images list has been created."
            resetXMLData
        End If
        
    Else
        resetXMLData
    End If
    
    Dim cancel As Boolean
    cancel = False
    m_MRUList.BeforeListLoad Me, cancel
    If cancel Then Exit Sub
       
    'We are now ready to load the actual MRU data from file.
    
    'The XML engine will do most the heavy lifting for this task.  We pass it a String array, and it fills it with
    ' all values corresponding to the given tag name and attribute.  (We must do this dynamically, because we don't
    ' know how many recent filenames are actually saved - it could be anywhere from 0 to RECENT_FILE_COUNT.)
    Dim allRecentFiles() As String
    If m_xmlEngine.findAllAttributeValues(allRecentFiles, "mruEntry", "id") Then
        m_numEntries = UBound(allRecentFiles) + 1
        
        'Make sure the file does not contain more entries than are allowed (shouldn't theoretically be possible,
        ' but it doesn't hurt to check).
        If m_numEntries > g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10) Then
            m_numEntries = g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10)
        End If
        
    'No recent file entries were found.  Adjust the Recent Files menu to match
    Else
        m_numEntries = 0
        
        m_MRUList.OnListIsEmpty
    End If
    
    'If one or more recent file entries were found, load them now.
    If m_numEntries > 0 Then
    
        'Remember that the allRecentFiles array was filled with the id values of all "mruEntry" entries in the XML file.
        ' We can use its upper bound to determine how many MRU entries are actually present in the file
        m_numEntries = UBound(allRecentFiles) + 1
        ReDim m_fileList(0 To m_numEntries) As String
        
        'Load the actual file paths from the MRU file
        Dim i As Long
        For i = 0 To m_numEntries - 1
        
            m_fileList(i) = m_xmlEngine.getUniqueTag_String("filePath", , , "mruEntry", "id", allRecentFiles(i))
            
            m_MRUList.OnItemLoad Me, i, getMRUCaption(m_fileList(i))
            
        Next i
    
        m_MRUList.AfterListLoaded Me
    End If
End Sub

'Save the current MRU list to file (currently done at program close)
Public Sub MRU_SaveToFile()
    Dim cancel As Boolean
    cancel = False
    m_MRUList.BeforeListSave Me, cancel
    If cancel Then
        Exit Sub
    End If

    On Error Resume Next

    'Reset whatever XML data we may have stored at present - we will be rewriting the full MRU file from scratch.
    resetXMLData
    
    Dim i As Long
    
    'Only write new entries if MRU data exists for them
    If m_numEntries > 0 Then
        For i = 0 To m_numEntries - 1
            m_xmlEngine.writeTagWithAttribute "mruEntry", "id", CStr(i), "", True
            m_xmlEngine.writeTag "filePath", m_fileList(i)
            m_xmlEngine.closeTag "mruEntry"
            m_xmlEngine.writeBlankLine
        Next i
    End If
    
    'With the XML file now complete, write it out to file
    m_xmlEngine.writeXMLToFile m_xmlfilename
    
    'Unload all corresponding menu entries.  (This doesn't matter when the program is closing, but we also use this
    ' routine to refresh the MRU list after changing the caption preference - and for that an unload is required.)
    m_MRUList.AfterListCleared
    
    ' Perform any cleanup
    m_MRUList.AfterListSaved Me
End Sub

'Add another file to the MRU list
Public Sub MRU_AddNewFile(ByVal newFile As String, Optional ByRef srcImage As pdImage = Nothing)
    ' The filename is sometimes passed via dialogs, and as such may contain nulls
    newFile = TrimNullW(newFile)

    Dim cancel As Boolean
    cancel = False
    m_MRUList.BeforeNewFileAdded Me, newFile, cancel
    If cancel Then Exit Sub

    'Locators are used to determine if this file already exists in the recent files list.
    ' If it does, we will simply shuffle its position instead of adding it as a new entry.
    Dim alreadyThere As Boolean
    alreadyThere = False
    
    Dim curLocation As Long
    curLocation = -1
    
    Dim i As Long
    
    'First, check to see if this file currently exists in the MRU list
    For i = 0 To m_numEntries - 1
    
        'This file already exists in the list!  Make a note of its location, then exit.
        If StrComp(m_fileList(i), newFile, vbBinaryCompare) = 0 Then
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
                m_fileList(i) = m_fileList(i - 1)
            Next i
        End If
    
    'This file doesn't exist in the MRU list, so it must be added at the very top as a new entry.
    Else

        m_numEntries = m_numEntries + 1
        
        'Cap the number of MRU files at a certain value (specified by the user in the Preferences menu)
        If m_numEntries > g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10) Then
            m_numEntries = g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10)
            
            'Also, because we are about to purge the MRU list, the last entry may require cleanup
            m_MRUList.TrimListAtPosition Me, m_numEntries - 1
        End If
        
        'Resize the list of MRU entries, which may have grown on account of this new addition.
        ReDim Preserve m_fileList(0 To m_numEntries) As String
    
        'Shift all existing entries downward
        If m_numEntries > 1 Then
            For i = m_numEntries To 1 Step -1
                m_fileList(i) = m_fileList(i - 1)
            Next i
        End If
        
    End If
    
    'Add this entry to the top of the list
    m_fileList(0) = newFile
    
    'Save a thumbnail of this image to file.
    If m_MRUList.SupportsThumbNails Then
        m_MRUList.SaveThumbNail newFile, srcImage
    End If
      
    'Based on the user's preference, display just the filename or the entire file path (up to the max character length)
    m_MRUList.AfterNewFileAdded Me, newFile, getMRUCaption(newFile)
End Sub


'If the user changes their preference regarding the number of recent files we can save, call this sub to rebuild
' the current menu.
Public Sub MRU_NotifyNewMaxLimit()
    
    'Erase any entries above the new limit
    If m_numEntries > g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10) Then
        m_numEntries = g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10)
        ' Delete the old MRU file now, because we need to truncate it
        If File_And_Path_Handling.FileExist(m_xmlfilename) Then
            Kill m_xmlfilename
        End If
    End If
    
    'Write the current MRU list out to file.
    MRU_SaveToFile
    
    'Unload all recent file menu entries
    m_MRUList.AfterListCleared
    
    'Reload MRU data from file
    MRU_LoadFromFile
    
    m_MRUList.OnMaxLimitChanged Me
End Sub

'Empty the entire MRU list and clear the menu of all entries
Public Sub MRU_ClearList()
    
    'Reset the number of entries in the MRU list
    m_numEntries = 0
    ReDim m_fileList(0) As String
    
    'Erase any existing XML file, and reset the in-memory version
    If File_And_Path_Handling.FileExist(m_xmlfilename) Then Kill m_xmlfilename
    resetXMLData
    
    m_MRUList.AfterListCleared
    
    m_MRUList.OnListIsEmpty
End Sub

'Return how many MRU entries are currently in the menu
Public Function MRU_ReturnCount() As Long
    MRU_ReturnCount = m_numEntries
End Function

Public Sub InitList(specificList As IMRUList)
    Set m_MRUList = specificList
    'Initialize an XML engine, which we will use to read/write our MRU data to file
    Set m_xmlEngine = New pdXML
    
    'The XML data will be stored in the Preset path (/Data/Presets)
    m_xmlfilename = m_MRUList.XMLFileName
    
    'If an XML file exists, it will be loaded separately, by the MRU_LoadFromFile() function
End Sub
