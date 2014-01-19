Attribute VB_Name = "Image_Autosave_Handler"
'***************************************************************************
'Image Autosave Handler
'Copyright ©2013-2014 by Tanner Helland
'Created: 18/January/14
'Last updated: 18/January/14
'Last update: initial build
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

Private Type autoSaveEntry
    fullPath As String
    origImageID As Long
    origUndoID As Long
End Type

'For performance reasons, we cache the list of Autosave files found during our initial search (if any)
Private m_ListOfAutosaveFiles() As autoSaveEntry
Private m_numOfAutosaveFiles As Long

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
    ' g_UserPreferences.getTempPath & "~cPDU_" & parentPDImage.imageID & "_" & uIndex & ".tmp"
    
    'This is going to change in the future, as some kind of Layer ID will also be necessary.
    
    Dim numOfImagesFound As Long
    numOfImagesFound = 0
    m_numOfAutosaveFiles = 0
    ReDim m_ListOfAutosaveFiles(0 To 9) As autoSaveEntry
    
    Dim imgIDCheck As Long, undoIDCheck As Long
    
    'Retrieve the first image from the list (if any)
    Dim chkFile As String
    chkFile = Dir(g_UserPreferences.getTempPath & "~cPDU_*_*.tmp", vbNormal)
        
    Do While Len(chkFile) > 0
    
        'Do some processing on said file to make sure it is valid; if it is, increment the "images found" counter
        If getIDValuesFromUndoPath(chkFile, imgIDCheck, undoIDCheck) Then
            numOfImagesFound = numOfImagesFound + 1
            
            'Also, cache this path in a module-level array, which we'll use externally to interact with the user
            m_numOfAutosaveFiles = m_numOfAutosaveFiles + 1
            
            If (m_numOfAutosaveFiles - 1) > UBound(m_ListOfAutosaveFiles) Then
                ReDim Preserve m_ListOfAutosaveFiles(0 To UBound(m_ListOfAutosaveFiles) * 2) As autoSaveEntry
            End If
            
            m_ListOfAutosaveFiles(m_numOfAutosaveFiles - 1).fullPath = g_UserPreferences.getTempPath & chkFile
            m_ListOfAutosaveFiles(m_numOfAutosaveFiles - 1).origImageID = imgIDCheck
            m_ListOfAutosaveFiles(m_numOfAutosaveFiles - 1).origUndoID = undoIDCheck
            
        End If
        
        'Check the next file in the list
        chkFile = Dir
        
    Loop
    
    saveableImagesPresent = numOfImagesFound

End Function

'If the user declines to restore old AutoSave data, purge it from the system (to prevent it from showing up in future searches).
Public Sub purgeOldAutosaveData()

    If m_numOfAutosaveFiles > 0 Then
    
        Dim i As Long
        For i = 0 To m_numOfAutosaveFiles - 1
        
            'Validate each path before removing it from the system (just to be safe!)
            If i < UBound(m_ListOfAutosaveFiles) Then
            
                If Len(m_ListOfAutosaveFiles(i).fullPath) > 0 Then
                    If FileExist(m_ListOfAutosaveFiles(i).fullPath) Then Kill m_ListOfAutosaveFiles(i).fullPath
                    
                    'Also check for selection data matching this file, and remove it if present
                    If FileExist(m_ListOfAutosaveFiles(i).fullPath & ".selection") Then Kill m_ListOfAutosaveFiles(i).fullPath & ".selection"
                    
                End If
                
            End If
        
        Next i
        
        'Release any memory associated with autosaves
        m_numOfAutosaveFiles = 0
        ReDim m_ListOfAutosaveFiles(0) As autoSaveEntry
    
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
