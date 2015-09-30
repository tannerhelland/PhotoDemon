Attribute VB_Name = "Macro_Interface"
'***************************************************************************
'PhotoDemon Macro Interface
'Copyright 2001-2015 by Tanner Helland
'Created: 10/21/01
'Last updated: 17/February/15
'Last updated by: Raj
'Last update: Update the macro MRU list when a macro is recorded or played.
'
'This (relatively small) sub handles all macro-related operations.  Macros are simply a recorded list of program operations, which
' can be "played back" to automate complex lists of image processing actions.  To create a macro, the user can "record" themselves
' applying a series of actions to an image.  When finished, they can then save that complete list of actions to file, then re-play
' those actions back at any time in the future.
'
'PhotoDemon's batch processing wizard allows use of macros, so that any combination of actions can be applied to any combination of
' images automatically.  This is a trademark feature of the program.
'
'As of 2014, the macro engine has been rewritten in significant ways.  Macros now rely on PhotoDemon's new string-based param
' design, and all macro settings are saved out to valid XML files.  This makes the human-readable and human-editable, but it also
' means that old macro files are no longer supported.  Users of old macro files are automatically warned of this change if they try
' to load an outdated macro file.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Macro loading information

'The current macro version string, which must be embedded in every saved macro file.
Public Const MACRO_VERSION_2014 As String = "8.2014"

'Macro recording information
Public MacroStatus As Byte
Public Const MacroSTOP As Long = 0
Public Const MacroSTART As Long = 1
Public Const MacroBATCH As Long = 2
Public Const MacroPLAYBACK As Long = 3
Public Const MacroCANCEL As Long = 128
Public MacroMessage As String

Public Sub StartMacro()
    
    'Set the program-wide "recording" flag
    MacroStatus = MacroSTART
    
    'Resize the array that will hold the macro data
    ProcessCount = 1
    ReDim Processes(0 To ProcessCount) As ProcessCall
    
    'Update any related macro UI elements
    Macro_Interface.updateMacroUI True
    
End Sub

'Stop recording the current macro, and offer to save it to file.
Public Sub StopMacro()
    
    'Before stopping the macro, make sure at least one valid, recordable action has occurred.
    Dim i As Long, numOfValidProcesses As Long
    numOfValidProcesses = 0
    
    For i = 0 To ProcessCount
        If (Len(Processes(i).Id) <> 0) And (Not Processes(i).Dialog) And Processes(i).Recorded Then
            numOfValidProcesses = numOfValidProcesses + 1
        End If
    Next i
    
    Dim msgReturn As VbMsgBoxResult
    
    If numOfValidProcesses = 0 Then
    
        'Warn the user that this macro won't be saved unless they keep recording
        msgReturn = PDMsgBox("This macro does not contain any recordable actions.  Are you sure you want to stop recording?" & vbCrLf & vbCrLf & "(Press No to continue recording.)", vbApplicationModal + vbExclamation + vbYesNo, "Warning: invalid macro")
        
        If msgReturn = vbYes Then
            
            'Update any related macro UI elements
            Macro_Interface.updateMacroUI False
            
            'Reset the macro engine and exit
            MacroStatus = MacroSTOP
            ProcessCount = 0
            Message "Macro abandoned."
            Exit Sub
        
        'If the user clicks anything but "yes", exit without making changes (e.g. let them continue recording).
        Else
            Exit Sub
        End If
        
    End If
    
    MacroStatus = MacroSTOP
    
    'Update any related macro UI elements
    Macro_Interface.updateMacroUI False
    
    'Automatically launch the save macro data routine
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
        
    Dim sFile As String
    
    Dim cdFilter As String
    cdFilter = PROGRAMNAME & " " & g_Language.TranslateMessage("Macro") & " (." & MACRO_EXT & ")|*." & MACRO_EXT
            
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Save macro data")
    
    'If the user cancels the save dialog, we'll raise a warning to tell them that the macro will be lost for good.
    ' That dialog gives them an option to return to the save dialog, which will bring us back to this line of code.
SaveMacroAgain:
     
    'If we get the data we want, save the information
    If saveDialog.GetSaveFileName(sFile, , True, cdFilter, 1, g_UserPreferences.getMacroPath, cdTitle, "." & MACRO_EXT, GetModalOwner().hWnd) Then
        
        'Save this macro's directory as the default macro path
        g_UserPreferences.setMacroPath sFile
        
        'Create a pdXML class, which will help us assemble the macro file
        Dim xmlEngine As pdXML
        Set xmlEngine = New pdXML
        xmlEngine.prepareNewXML "Macro"
        
        'Write out the XML version we're using for this macro
        xmlEngine.writeTag "pdMacroVersion", MACRO_VERSION_2014
        
        'We now want to count the number of actual processes that we will be writing to file.  A valid process meets
        ' the following criteria:
        ' 1) It isn't blank/empty
        ' 2) It doesn't display a dialog
        ' 3) It was not specifically marked as "DO_NOT_RECORD"
        
        'Due to the previous check at the top of this function, we already know how many valid functions are in the process list,
        ' and this value is guaranteed to be non-zero.
        
        'Write out the number of valid processes in the macro
        xmlEngine.writeTag "processCount", CStr(numOfValidProcesses)
        xmlEngine.writeBlankLine
        
        'Now, write out each macro entry in the current process list
        numOfValidProcesses = 0
        
        For i = 0 To ProcessCount
            
            'We only want to write out valid processes, using the same criteria as the original counting loop above.
            If (Len(Processes(i).Id) <> 0) And (Not Processes(i).Dialog) And Processes(i).Recorded Then
                numOfValidProcesses = numOfValidProcesses + 1
                
                'Start each process entry with a unique identifier
                xmlEngine.writeTagWithAttribute "processEntry", "index", numOfValidProcesses, "", True
                
                'Write out all the properties of this entry
                xmlEngine.writeTag "ID", Processes(i).Id
                xmlEngine.writeTag "Parameters", Processes(i).Parameters
                xmlEngine.writeTag "MakeUndo", Str(Processes(i).MakeUndo)
                xmlEngine.writeTag "Tool", Str(Processes(i).Tool)
                
                'Note that the Dialog and Recorded properties are not written to file.  There is no need to remember
                ' them, as we know their values must be FALSE and TRUE, respectively, per the check above.
            
                'Close this process entry
                xmlEngine.closeTag "processEntry"
                xmlEngine.writeBlankLine
            End If
            
        Next i
        
        'With all tags successfully written, we can now close the XML data and write it out to file.
        xmlEngine.writeXMLToFile sFile
        
        Message "Macro saved successfully."
        
        'At this point, the macro should be added to the Recent Macros list
        g_RecentMacros.MRU_AddNewFile sFile
        
    Else
        
        msgReturn = PDMsgBox("If you do not save this macro, all actions recorded during this session will be permanently lost.  Are you sure you want to cancel?" & vbCrLf & vbCrLf & "(Press No to return to the Save Macro screen.  Note that you can always delete this macro later if you decide you don't want it.)", vbApplicationModal + vbExclamation + vbYesNo, "Warning: last chance to save macro")
        If msgReturn = vbNo Then GoTo SaveMacroAgain
        
        Message "Macro abandoned."
        
    End If
            
    ProcessCount = 0
    
End Sub

'All macro-related UI instructions should be placed here, as PD can terminate a macro recording session for any number of reasons,
' and it needs a uniform way to wipe macro-related UI changes).
Private Sub updateMacroUI(ByVal recordingIsActive As Boolean)

    If recordingIsActive Then
    
        'Notify the user that recording has begun
        Message "Macro recording started."
        toolbar_Toolbox.lblRecording.Visible = True
        
        'Disable "start recording", and enable "stop recording"
        FormMain.MnuRecordMacro(0).Enabled = False
        FormMain.MnuRecordMacro(1).Enabled = True
    
    Else
        Message "Macro recording stopped."
        toolbar_Toolbox.lblRecording.Visible = False
        FormMain.MnuRecordMacro(0).Enabled = True
        FormMain.MnuRecordMacro(1).Enabled = False
    End If

End Sub

Public Sub PlayMacro()

    'Disable user input until the dialog closes
    Interface.DisableUserInput

    'Automatically launch the load Macro data routine
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    Dim sFile As String
        
    Dim cdFilter As String
    cdFilter = PROGRAMNAME & " " & g_Language.TranslateMessage("Macro") & " (." & MACRO_EXT & ")|*." & MACRO_EXT & ";*.thm"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Open Macro File")
        
    'If we get a path, load that file
    If openDialog.GetOpenFileName(sFile, , True, , cdFilter, 1, g_UserPreferences.getMacroPath, cdTitle, "." & MACRO_EXT, GetModalOwner().hWnd) Then
        
        Message "Loading macro data..."
        
        'Save this macro's folder as the default macro path
        g_UserPreferences.setMacroPath sFile
                
        PlayMacroFromFile sFile
        
    Else
        Message "Macro load canceled."
    End If
    
    'Re-enable user input
    Interface.EnableUserInput
        
End Sub

'Given a valid macro file, play back its recorded actions.
Public Function PlayMacroFromFile(ByVal MacroPath As String) As Boolean
    
    'Create a pdXML class, which will help us load and parse the source file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Load the XML file into memory
    xmlEngine.loadXMLFile MacroPath
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon macro file
    If xmlEngine.isPDDataType("Macro") And xmlEngine.validateLoadedXMLData("pdMacroVersion") Then
    
        'Next, check the macro's version number, and make sure it's still supported
        Dim verCheck As String
        verCheck = xmlEngine.getUniqueTag_String("pdMacroVersion")
        
        Select Case verCheck
        
            'The current macro version (e.g. the first draft of the new XML format)
            Case MACRO_VERSION_2014
            
                'Retrieve the number of processes in this macro
                ProcessCount = xmlEngine.getUniqueTag_Long("processCount")
                
                If ProcessCount > 0 Then
                
                    ReDim Processes(0 To ProcessCount - 1) As ProcessCall
                    
                    'Start retrieving individual process data from the file
                    Dim i As Long
                    For i = 1 To ProcessCount
                    
                        'Start by finding the location of the tag we want
                        Dim tagPosition As Long
                        tagPosition = xmlEngine.getLocationOfTagPlusAttribute("processEntry", "index", i)
                        
                        If tagPosition > 0 Then
                        
                            'Use that tag position to retrieve the processor parameters we need.
                            With Processes(i - 1)
                                .Id = xmlEngine.getUniqueTag_String("ID", , tagPosition)
                                .Parameters = xmlEngine.getUniqueTag_String("Parameters", , tagPosition)
                                .MakeUndo = xmlEngine.getUniqueTag_Long("MakeUndo", , tagPosition)
                                .Tool = xmlEngine.getUniqueTag_Long("Tool", , tagPosition)
                                
                                'These two attributes can be assigned automatically, as we know what their values must be.
                                .Dialog = False
                                .Recorded = True
                            End With
                            
                        Else
                            Debug.Print "Expected macro entry could not be found!"
                        End If
                    
                    Next i
                    
                'This macro file contains no valid actions.  It's no longer possible to create a macro like this, so this is basically
                ' a failsafe for faulty old versions of PD.
                Else
                    
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "WARNING!  ProcessCount is zero!  Macro file is technically valid, but there's nothing to see here..."
                    #End If
                    
                    Message "Macro complete!"
                    PlayMacroFromFile = True
                    Exit Function
                    
                End If
            
            Case Else
                Message "Incompatible macro version found.  Macro playback abandoned."
                PlayMacroFromFile = False
                Exit Function
        
        End Select
        
        'Mark the load as successful and continue
        PlayMacroFromFile = True
        
    Else
    
        PDMsgBox "Unfortunately, this macro file is no longer supported by the current version of PhotoDemon." & vbCrLf & vbCrLf & "In version 6.0, PhotoDemon macro files were redesigned to support new features, improve performance, and solve some long-standing reliability issues.  Unfortunately, this means that macros recorded prior to version 6.0 are no longer compatible.  You will need to re-record these macros from scratch." & vbCrLf & vbCrLf & "(Note that any old macro files will still work in old versions of PhotoDemon, if you absolutely need to access them.)", vbInformation + vbOKOnly, "Unsupported macro file"
        PlayMacroFromFile = False
        Exit Function
        
    End If
    
    'Now we run a loop through the macro structure, calling the software processor with all the necessary information for each action
    Message "Processing macro data..."
    
    MacroStatus = MacroPLAYBACK
    
    Dim tProc As Long
    For tProc = 0 To ProcessCount - 1
        Process Processes(tProc).Id, Processes(tProc).Dialog, Processes(tProc).Parameters, Processes(tProc).MakeUndo, Processes(tProc).Tool, Processes(tProc).Recorded
    Next tProc
    
    MacroStatus = MacroSTOP
    
    'Some processor requests may not manually update the screen; as such, perform a manual update now
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    'Our work here is complete!
    Message "Macro complete!"
    
    'After playing, the macro should be added to the Recent Macros list
    g_RecentMacros.MRU_AddNewFile MacroPath
    
End Function
