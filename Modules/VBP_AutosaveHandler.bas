Attribute VB_Name = "Image_Autosave_Handler"
'***************************************************************************
'Image Autosave Handler
'Copyright ©2013-2014 by Tanner Helland
'Created: 18/January/14
'Last updated: 23/January/14
'Last update: wrap up initial build
'
'PhotoDemon's Autosave engine is closely tied to the pdUndo class, so some understanding of that class is necessary
' to appreciate how this module operates.
'
'All Undo/Redo data is saved to the hard drive, in a temp folder of the user's choosing (the Windows temp folder
' by default).  The data is cleared whenever an image is unloaded, and an extra pass is made at program shutdown
' "just to be safe".
'
'In the event of an unclean shutdown, this module searches the temp folder for any PhotoDemon-specific data.  If
' some is found, the user is given a choice to restore those files.  If the user declines, that data is wiped
' (to prevent future unclean shutdown checks from re-detecting it).
'
'As part of its Autosave functionality, this module also handles the creation and subsequent destruction of a
' "clean shutdown" file.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Binary autosave entries.  These are raw image buffers dumped to the user's temp folder as part of normal
' PD processing.
Public Type autosaveBinary
    fullPath As String
    origImageID As Long
    origUndoID As Long
End Type

'A collection of valid Autosave XML entries in the user's Data\Autosave folder.  In all but the worst-case
' scenarios (e.g. program failure during generating Undo/Redo data), these *should* correspond to raw image
' data in the Undo/Redo list.
Public Type autosaveXML
    xmlPath As String
    friendlyName As String
    idValue As Long
    isDisplayed As Boolean
    latestUndoFound As Long
    latestUndoPath As String
    isBufferOnly As Boolean
End Type

'For performance reasons, we cache the list of Autosave files found during our initial search (if any)
Private m_BinaryEntries() As autosaveBinary
Private m_numOfBinaryFound As Long

'Collection of Autosave XML entries found
Private m_numOfXMLFound As Long
Private m_XmlEntries() As autosaveXML

'Check to make sure the last program shutdown was clean.  If it was, return TRUE (and write out a new safe shutdown file).
' If it was not, return FALSE.
Public Function wasLastShutdownClean() As Boolean

    Dim safeShutdownPath As String
    safeShutdownPath = g_UserPreferences.getPresetPath & "SafeShutdown.xml"
    
    'If a previous program session terminated unexpectedly, its safe shutdown file will still be present
    If FileExist(safeShutdownPath) Then
    
        wasLastShutdownClean = False

    'The previous shutdown was clean.  Write a new safe shutdown file.
    Else
    
        Dim xmlEngine As pdXML
        Set xmlEngine = New pdXML
        
        xmlEngine.prepareNewXML "Safe shutdown"
        
        xmlEngine.writeBlankLine
        xmlEngine.writeComment "This file is used to see if the previous PhotoDemon session terminated unexpectedly."
        xmlEngine.writeBlankLine
        xmlEngine.writeTag "SessionDate", Format$(Now, "Long Date")
        xmlEngine.writeTag "SessionTime", Format$(Now, "h:mm AMPM")
        xmlEngine.writeBlankLine
        
        xmlEngine.writeXMLToFile safeShutdownPath
        
        wasLastShutdownClean = True
    
    End If
    
    
End Function

'If the program has shut itself down without incident, the last thing it does will be notifying this sub.
' (This sub clears the safe shutdown file.)
Public Sub notifyCleanShutdown()
    
    Dim safeShutdownPath As String
    safeShutdownPath = g_UserPreferences.getPresetPath & "SafeShutdown.xml"
    
    If FileExist(safeShutdownPath) Then Kill safeShutdownPath

End Sub

'After an unclean shutdown is detected, this function can be called to search the temp directory for saveable Undo/Redo data.
' It will return a value larger than 0 if Undo/Redo data was found.
Public Function saveableImagesPresent() As Long

    'Search the temporary folder for any files matching PhotoDemon's Undo/Redo file pattern.  The pattern we currently use is:
    ' g_UserPreferences.getTempPath & "~cPDU_" & parentPDImage.imageID & "_" & uIndex & ".pdtmp"
    
    'This is going to change in the future, as some kind of Layer ID will also be necessary.
    
    Dim numOfImagesFound As Long
    numOfImagesFound = 0
    m_numOfBinaryFound = 0
    ReDim m_BinaryEntries(0 To 9) As autosaveBinary
    
    Dim imgIDCheck As Long, undoIDCheck As Long
    
    'Retrieve the first image from the list (if any)
    Dim chkFile As String
    chkFile = Dir(g_UserPreferences.getTempPath & "~cPDU_*_*.pdtmp", vbNormal)
        
    Do While Len(chkFile) > 0
    
        'Do some processing on said file to make sure it is valid; if it is, increment the "images found" counter
        If getIDValuesFromUndoPath(chkFile, imgIDCheck, undoIDCheck) Then
            numOfImagesFound = numOfImagesFound + 1
            
            'Also, cache this path in a module-level array, which we'll use externally to interact with the user
            m_numOfBinaryFound = m_numOfBinaryFound + 1
            
            If (m_numOfBinaryFound - 1) > UBound(m_BinaryEntries) Then
                ReDim Preserve m_BinaryEntries(0 To UBound(m_BinaryEntries) * 2) As autosaveBinary
            End If
            
            m_BinaryEntries(m_numOfBinaryFound - 1).fullPath = g_UserPreferences.getTempPath & chkFile
            m_BinaryEntries(m_numOfBinaryFound - 1).origImageID = imgIDCheck
            m_BinaryEntries(m_numOfBinaryFound - 1).origUndoID = undoIDCheck
            
        End If
        
        'Check the next file in the list
        chkFile = Dir
        
    Loop
    
    saveableImagesPresent = numOfImagesFound

End Function

'If the user declines to restore old AutoSave data, purge it from the system (to prevent it from showing up in future searches).
Public Sub purgeOldAutosaveData()

    Dim i As Long
    
    'Release binary autosave entries first
    If m_numOfBinaryFound > 0 Then
    
        For i = 0 To m_numOfBinaryFound - 1
        
            'Validate each path before removing it from the system (just to be safe!)
            If i < UBound(m_BinaryEntries) Then
            
                If Len(m_BinaryEntries(i).fullPath) > 0 Then
                    If FileExist(m_BinaryEntries(i).fullPath) Then Kill m_BinaryEntries(i).fullPath
                    
                    'Also check for selection data matching this file, and remove it if present
                    If FileExist(m_BinaryEntries(i).fullPath & ".selection") Then Kill m_BinaryEntries(i).fullPath & ".selection"
                    
                End If
                
            End If
        
        Next i
        
        'Release any memory associated with autosaves
        m_numOfBinaryFound = 0
        ReDim m_BinaryEntries(0) As autosaveBinary
    
    End If
    
    'Follow it with XML entries
    If m_numOfXMLFound > 0 Then
    
        For i = 0 To m_numOfXMLFound - 1
        
            'Validate each path before removing it from the system (just to be safe!)
            If i < UBound(m_XmlEntries) Then
                If Len(m_XmlEntries(i).xmlPath) > 0 Then
                    If FileExist(m_XmlEntries(i).xmlPath) Then Kill m_XmlEntries(i).xmlPath
                End If
            End If
        
        Next i
        
        'Release any memory associated with autosaves
        m_numOfXMLFound = 0
        ReDim m_XmlEntries(0) As autosaveXML
    
    End If
    
End Sub

'Given a path to an Undo file, retrieve the image ID and undo ID from it.  Returns TRUE if successful, FALSE if the file in
' question does not appear to be a PD Undo/Redo file after all.
Private Function getIDValuesFromUndoPath(ByVal undoPath As String, ByRef imgID As Long, ByRef undoID As Long) As Boolean

    Dim pCheck() As String
    pCheck = Split(undoPath, "_")
    
    'pCheck() now contains the contents of undoPath, separated by underscore.  Make sure it generated at least three entries.
    If UBound(pCheck) < 2 Then
        getIDValuesFromUndoPath = False
        Exit Function
    End If
    
    'Search from the BACK of the string (as the path could have underscores, and we don't care about those), while checking
    ' to make sure the entry is 1) numeric, and 2) above 0
    If Not checkNumberAtPosition(pCheck, UBound(pCheck) - 1, undoID, -10) Then
        getIDValuesFromUndoPath = False
        Exit Function
    End If
    
    'Repeat the steps above, but for the imageID
    If Not checkNumberAtPosition(pCheck, UBound(pCheck) - 2, imgID, -10) Then
        getIDValuesFromUndoPath = False
        Exit Function
    End If
    
    'If we made it all the way here, the search was successful.  Return TRUE.
    getIDValuesFromUndoPath = True
    
End Function

'Used by getIDValuesFromUndoPath, above, to retrieve numeric identifiers from an Undo string
Private Function checkNumberAtPosition(ByRef srcString() As String, ByRef sArrayPosition As Long, ByRef dstNumber As Long, Optional ByVal lowBound As Long = 0) As Boolean

    If IsNumeric(srcString(sArrayPosition)) Then
        dstNumber = CLng(srcString(sArrayPosition))
        
        If dstNumber < lowBound Then
            checkNumberAtPosition = False
        Else
            checkNumberAtPosition = True
        End If
        
    Else
        checkNumberAtPosition = False
    End If
    
End Function

'External functions can retrieve a copy of the binary autosave entries we've found by using this function.
Public Function getBinaryAutosaveEntries(ByRef autosaveArray() As autosaveBinary, ByRef autosaveCount As Long) As Boolean

    ReDim autosaveArray(0 To m_numOfBinaryFound - 1) As autosaveBinary
    autosaveCount = m_numOfBinaryFound
    
    Dim i As Long
    For i = 0 To autosaveCount - 1
        autosaveArray(i) = m_BinaryEntries(i)
    Next i
    
    getBinaryAutosaveEntries = True
    
End Function

'External functions can retrieve a copy of the XML autosave entries we've found by using this function.
Public Function getXMLAutosaveEntries(ByRef autosaveArray() As autosaveXML, ByRef autosaveCount As Long) As Boolean

    ReDim autosaveArray(0 To m_numOfXMLFound - 1) As autosaveXML
    autosaveCount = m_numOfXMLFound
    
    Dim i As Long
    For i = 0 To autosaveCount - 1
        autosaveArray(i) = m_XmlEntries(i)
    Next i
    
    getXMLAutosaveEntries = True
    
End Function

'Retrieve all XML autosave entries from the user's /Data/Autosave folder.
Public Function findAllAutosaveXML() As Boolean

    m_numOfXMLFound = 0
    ReDim m_XmlEntries(0) As autosaveXML
    
    'Retrieve the first image from the list (if any)
    Dim chkFile As String
    chkFile = Dir(g_UserPreferences.getAutosavePath & "*.xml", vbNormal)
    
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    Do While Len(chkFile) > 0
    
        'Do some processing on said XML to make sure it is valid
        If xmlEngine.loadXMLFile(g_UserPreferences.getAutosavePath & chkFile) Then
        
            'Make sure the XML type is valid, and an ID value is present in the file
            If xmlEngine.isPDDataType("pdImage Backup") And xmlEngine.validateLoadedXMLData("ID") Then
            
                'The file checks out.  Add it to our XML entries array
                With m_XmlEntries(m_numOfXMLFound)
                    .xmlPath = g_UserPreferences.getAutosavePath & chkFile
                    .friendlyName = xmlEngine.getUniqueTag_String("OriginalFileNameAndExtension")
                    .idValue = xmlEngine.getUniqueTag_Long("ID", -1)
                    .isDisplayed = False
                    .latestUndoFound = 0
                    .latestUndoPath = ""
                End With
                
                'Increment the "number found" counter and resize the array as necessary
                m_numOfXMLFound = m_numOfXMLFound + 1
                If m_numOfXMLFound > UBound(m_XmlEntries) Then
                    ReDim Preserve m_XmlEntries(0 To (UBound(m_XmlEntries) + 1) * 2) As autosaveXML
                End If
            
            End If
        
        End If
        
        'Check the next file in the list
        chkFile = Dir
    
    Loop
    
    'All entries have been found.  Return TRUE if more than one entry was discovered.
    If m_numOfXMLFound > 0 Then
        findAllAutosaveXML = True
    Else
        findAllAutosaveXML = False
    End If

End Function

'Once Binary and XML autosave data has been retrieved, this function can be used to align the two.  The latest binary image buffer
' for each XML entry will be marked, and the resulting XML array will contain a full collection of relevant autosave data.
Public Sub alignXMLandBinaryAutosaves()

    Dim i As Long, j As Long
    Dim curImage As Long, curImageExists As Boolean
    
    Dim unfoundImages As Long
    unfoundImages = 0
    
    'For each raw image buffer found, we are now going to attempt to align it with an XML entry.  If we can, we
    ' will load a single copy of that image into the list box.
    For i = 0 To m_numOfBinaryFound - 1
        
        curImage = m_BinaryEntries(i).origImageID
        curImageExists = False
        
        'Find a matching entry in the xmlEntries list
        For j = 0 To m_numOfXMLFound - 1
        
            If m_XmlEntries(j).idValue = curImage Then
                
                'A matching XML entry was found!  See if this entry has already been matched up with a raw
                ' image buffer.
                
                If m_XmlEntries(j).isDisplayed Then
                
                    curImageExists = True
                
                    'This XML entry has already been matched up with a raw buffer.  See if this Undo value is
                    ' more recent than the previous one.
                    If m_XmlEntries(j).latestUndoFound < m_BinaryEntries(i).origUndoID Then
                    
                        'This entry is newer.  Update accordingly.
                        m_XmlEntries(j).latestUndoFound = m_BinaryEntries(i).origUndoID
                        m_XmlEntries(j).latestUndoPath = m_BinaryEntries(i).fullPath
                    
                    End If
                
                Else
                
                    'This XML entry has not yet been matched up with a raw buffer.  Match it now.
                    curImageExists = True
                    m_XmlEntries(j).latestUndoFound = m_BinaryEntries(i).origUndoID
                    m_XmlEntries(j).latestUndoPath = m_BinaryEntries(i).fullPath
                    m_XmlEntries(j).isDisplayed = True
                
                End If
                
            End If
        
            'If we've already found a matching entry, exit the search loop
            If curImageExists Then Exit For
        
        Next j
        
        'If this buffer did not have a corresponding entry in the XML array, let's add one now.  The entry will
        ' be necessarily incomplete due to not knowing things like the file's original name, but at least the
        ' user can recover the raw image data - which is certainly better than nothing at all!
        If Not curImageExists Then
        
            'Add a new spot to the XML array
            m_numOfXMLFound = m_numOfXMLFound + 1
            ReDim Preserve m_XmlEntries(0 To m_numOfXMLFound - 1) As autosaveXML
            
            'Fill the new spot with data corresponding to this set of raw image data
            unfoundImages = unfoundImages + 1
            With m_XmlEntries(m_numOfXMLFound - 1)
                .idValue = m_BinaryEntries(i).origImageID
                .isBufferOnly = True
                .isDisplayed = True
                .latestUndoFound = m_BinaryEntries(i).origUndoID
                .latestUndoPath = m_BinaryEntries(i).fullPath
                .friendlyName = g_Language.TranslateMessage("unknown image %1", CStr(unfoundImages))
            End With
            
        End If
        
    Next i

End Sub

'After any autosave images have been loaded into PD, call this function to replace those images' data (such as "location on disk")
' with information from the Autosave XML files.
Public Sub alignLoadedImageWithAutosave(ByRef srcPDImage As pdImage)

    Dim i As Long
    
    'Make sure the image loaded successfully
    If Not (srcPDImage Is Nothing) Then
    
        If srcPDImage.IsActive Then
        
            'Find a corresponding Autosave XML file for this image (if one exists)
            For i = 0 To m_numOfXMLFound - 1
            
                'If this file's location on disk matches the binary buffer associated with a given XML entry,
                ' ask the pdImage object to rewrite its internal data to match the XML file.
                If StrComp(srcPDImage.locationOnDisk, m_XmlEntries(i).latestUndoPath, vbTextCompare) = 0 Then
                    srcPDImage.readInternalDataFromFile m_XmlEntries(i).xmlPath
                    Exit For
                End If
            
            Next i
        
        End If
    
    End If
    
End Sub
