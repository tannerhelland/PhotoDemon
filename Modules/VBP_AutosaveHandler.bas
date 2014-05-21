Attribute VB_Name = "Image_Autosave_Handler"
'***************************************************************************
'Image Autosave Handler
'Copyright ©2013-2014 by Tanner Helland
'Created: 18/January/14
'Last updated: 20/May/14
'Last update: rewrite everything Autosave-related against PD's new Undo/Redo engine.
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
Public Type AutosaveXML
    xmlPath As String
    parentImageID As Long
    friendlyName As String
    undoStackHeight As Long
    undoStackAbsoluteMaximum As Long
    undoStackPointer As Long
    undoNumAtLastSave As Long
End Type

'Collection of Autosave XML entries found
Private m_numOfXMLFound As Long
Private m_XmlEntries() As AutosaveXML

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

    'Search the temporary folder for any files matching PhotoDemon's Undo/Redo file pattern.  Because PD's Undo/Redo engine
    ' is awesome, it automatically saves very nice Undo XML files that contain key data for each pdImage opened by the program.
    ' In the event of an unsafe shutdown, these XML files help us easily reconstruct any "lost" images.
    
    'Note: the pattern of PhotoDemon's Undo XML summary files is:
    ' g_UserPreferences.GetTempPath & "~PDU_StackSummary_" & parentPDImage.imageID & "_.pdtmp"
    
    'Reset our XML detection arrays
    m_numOfXMLFound = 0
    ReDim m_XmlEntries(0 To 9) As AutosaveXML
    
    'We'll use PD's standard XML engine to validate any discovered autosave entries
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Retrieve the first matching file from the folder (if any)
    Dim chkFile As String
    chkFile = Dir(g_UserPreferences.GetTempPath & "~PDU_StackSummary_*_.pdtmp", vbNormal)
    
    'Continue checking potential autosave XML entries until all have been analyzed
    Do While Len(chkFile) > 0
    
        'First, make sure the file actually contains XML data
        If xmlEngine.loadXMLFile(g_UserPreferences.GetTempPath & chkFile) Then
        
            'If it does, make sure the XML data is valid, and that at least one Undo entry is listed in the file
            If xmlEngine.isPDDataType("Undo stack") And xmlEngine.validateLoadedXMLData("pdUndoVersion") Then
            
                'The file checks out!  Add it to our XML entries array
                With m_XmlEntries(m_numOfXMLFound)
                    .xmlPath = g_UserPreferences.GetTempPath & chkFile
                    .friendlyName = xmlEngine.getUniqueTag_String("friendlyName")
                    .parentImageID = xmlEngine.getUniqueTag_Long("imageID", -1)
                    .undoNumAtLastSave = xmlEngine.getUniqueTag_Long("UndoNumAtLastSave", 0)
                    .undoStackAbsoluteMaximum = xmlEngine.getUniqueTag_Long("StackAbsoluteMaximum", 0)
                    .undoStackHeight = xmlEngine.getUniqueTag_Long("StackHeight", 1)
                    .undoStackPointer = xmlEngine.getUniqueTag_Long("CurrentStackPointer", 0)
                End With
                
                'Increment the "number found" counter and resize the array as necessary
                m_numOfXMLFound = m_numOfXMLFound + 1
                If m_numOfXMLFound > UBound(m_XmlEntries) Then
                    ReDim Preserve m_XmlEntries(0 To (UBound(m_XmlEntries) + 1) * 2) As AutosaveXML
                End If
                
            End If
            
        End If
        
        'Check the next file in the list
        chkFile = Dir
        
    Loop
    
    'Trim the XML array to its smallest relevant size, then return the number of images found
    ReDim Preserve m_XmlEntries(0 To m_numOfXMLFound) As AutosaveXML
    
    saveableImagesPresent = m_numOfXMLFound

End Function

'If the user declines to restore old AutoSave data, purge it from the system (to prevent it from showing up in future searches).
Public Sub purgeOldAutosaveData()
    
    Message "Purging old autosave data..."
    
    'Create a dummy pdUndo object.  This object will help us generate relevant filenames using PD's standard Undo filename formula.
    Dim tmpUndoEngine As pdUndo
    Set tmpUndoEngine = New pdUndo
    
    Dim tmpFilename As String
    Dim i As Long, j As Long
    
    'Loop through all XML files found.  We will not only be deleting the XML files themselves, but also any child
    ' files they may reference
    For i = 0 To m_numOfXMLFound - 1
    
        Debug.Print "Attempting to delete " & m_XmlEntries(i).undoStackAbsoluteMaximum & " files..."
    
        'Delete all possible child references for this image.
        For j = 0 To m_XmlEntries(i).undoStackAbsoluteMaximum
        
            tmpFilename = tmpUndoEngine.generateUndoFilenameExternal(m_XmlEntries(i).parentImageID, j)
        
            'Check image data first...
            If FileExist(tmpFilename) Then Kill tmpFilename
        
            '...followed by layer data
            If FileExist(tmpFilename & ".layer") Then Kill tmpFilename & ".layer"
        
            '...followed by selection data
            If FileExist(tmpFilename & ".selection") Then Kill tmpFilename & ".selection"
        
        Next j
        
        'Finally, kill the Autosave XML file and preview image associated with this entry
        If FileExist(m_XmlEntries(i).xmlPath) Then Kill m_XmlEntries(i).xmlPath
        If FileExist(m_XmlEntries(i).xmlPath & ".asp") Then Kill m_XmlEntries(i).xmlPath & ".asp"
    
    Next i
    
    'As a nice gesture, release any module-level data associated with the Autosave engine
    m_numOfXMLFound = 0
    ReDim m_XmlEntries(0) As AutosaveXML
    
End Sub

'External functions can retrieve a copy of the XML autosave entries we've found by using this function.
Public Function getXMLAutosaveEntries(ByRef autosaveArray() As AutosaveXML, ByRef autosaveCount As Long) As Boolean

    ReDim autosaveArray(0 To m_numOfXMLFound - 1) As AutosaveXML
    autosaveCount = m_numOfXMLFound
    
    Dim i As Long
    For i = 0 To autosaveCount - 1
        autosaveArray(i) = m_XmlEntries(i)
    Next i
    
    getXMLAutosaveEntries = True
    
End Function

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
                If StrComp(srcPDImage.locationOnDisk, m_XmlEntries(i).xmlPath, vbTextCompare) = 0 Then
                    srcPDImage.readExternalData m_XmlEntries(i).xmlPath
                    Exit For
                End If
            
            Next i
        
        End If
    
    End If
    
End Sub
