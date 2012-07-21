Attribute VB_Name = "MacroAndBatchConversion"
'***************************************************************************
'Macro and Batch Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 10/21/01
'Last updated: 30/July/06
'Last update: Uses new process call format and can convert old 2002 files
'             to the new 2006 format
'
'Routines for automated image processing.  These are based off recorded macros,
' which the user must supply by recording actions from within PhotoDemon.
' Similarly, PhotoDemon's batch processing functionality relies heavily on
' this module.
'
'***************************************************************************

Option Explicit

'Macro loading information
Public Const MACRO_IDENTIFIER As String * 4 = "DSmf"
Public Const MACRO_VERSION_2006 As Long = &H2006

'OLD outdated macro versions (included only to preserve functionality)
Public Const MACRO_VERSION_2002 As Long = &H80000000

'The 2002 version was discontinued 30/July/06 in favor of additional opcodes
'(required for image levels and possibly future hDC-handling routines)
Public Type ProcessCall2002
    MainType As Long
    pOPCODE As Variant
    pOPCODE2 As Variant
    pOPCODE3 As Variant
    pOPCODE4 As Variant
    LoadForm As Boolean
    RecordAction As Boolean
End Type


'Macro recording information
Public MacroStatus As Byte
Public Const MacroSTOP As Long = 0
Public Const MacroSTART As Long = 1
Public Const MacroBATCH As Long = 2
Public Const MacroCANCEL As Long = 128
Public MacroMessage As String

Public Sub StartMacro()
    'Easy - set the flag and start recording
    MacroStatus = MacroSTART
    CurrentCall = 1
    'Transfer the sub data into an array for tracking
    ReDim Calls(0 To CurrentCall) As ProcessCall
    Message "Macro recording started."
    FormMain.lblRecording.Visible = True
    FormMain.Line3.Y1 = 328
    FormMain.Line3.Y2 = 328
End Sub

Public Sub StopMacro()
    MacroStatus = MacroSTOP
    'Automatically launch the save macro data routine
    Dim CC As cCommonDialog
    Dim sFile As String
    Set CC = New cCommonDialog
    
    'Get the last macro-related path from the INI file
    Dim tempPathString As String
    tempPathString = GetFromIni("Program Paths", "Macro")
    
    'If the user cancels the save dialog, give them another chance to save - just in case
    Dim mReturn As VbMsgBoxResult
     
SaveMacroAgain:
     
    'If we get the data we want, save the information
    If CC.VBGetSaveFileName(sFile, , True, PROGRAMNAME & " Macro Data (." & MACRO_EXT & ")|*." & MACRO_EXT, , tempPathString, "Save macro data", "." & MACRO_EXT, FormMain.HWnd, 0) Then
        
        'Save the new directory as the default path for future usage
        tempPathString = sFile
        StripDirectory tempPathString
        WriteToIni "Program Paths", "Macro", tempPathString

        'Delete any existing file (overwrite) and dump the info to file
        If FileExist(sFile) = True Then Kill sFile
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        Open sFile For Binary As #fileNum
            Put #fileNum, 1, MACRO_IDENTIFIER
            Put #fileNum, , MACRO_VERSION_2006
            'Remove the last call (stop macro recording - redundant to save that)
            CurrentCall = CurrentCall - 1
            Put #fileNum, , CurrentCall
            ReDim Preserve Calls(CurrentCall) As ProcessCall
            Put #fileNum, , Calls
        Close #fileNum
    Else
        
        mReturn = MsgBox("If you do not save this macro, all actions recorded during this session will be permanently lost.  Are you sure you want to cancel?" & vbCrLf & vbCrLf & "(Press No to return to the Save Macro screen.  Note that you can always delete this macro later if you decide you don't want it.)", vbApplicationModal + vbCritical + vbYesNo, "Warning: Last Chance to Save Macro")
        If mReturn = vbNo Then GoTo SaveMacroAgain
            
    End If
    Message "Macro recording stopped."
    FormMain.lblRecording.Visible = False
    FormMain.Line3.Y1 = 304
    FormMain.Line3.Y2 = 304
    CurrentCall = 0
End Sub

Public Sub PlayMacro()
    'Automatically launch the load Macro data routine
    Dim CC As cCommonDialog
    Dim sFile As String
    Set CC = New cCommonDialog
    
    'Get the last macro-related path from the INI file
    Dim tempPathString As String
    tempPathString = GetFromIni("Program Paths", "Macro")
   
    'If we get a path, load that file
    If CC.VBGetOpenFileName(sFile, , , , , True, PROGRAMNAME & " Macro Data (." & MACRO_EXT & ")|*." & MACRO_EXT & "|All files|*.*", , tempPathString, "Open Macro File", "." & MACRO_EXT, FormMain.HWnd, OFN_HIDEREADONLY) Then
        Message "Loading macro data..."
        
        'Save the new directory as the default path for future usage
        tempPathString = sFile
        StripDirectory tempPathString
        WriteToIni "Program Paths", "Macro", tempPathString
        
        PlayMacroFromFile sFile
        
    Else
        Message "Macro load canceled."
    End If
    
End Sub

'Need to convert this to a FUNCTION that returns a boolean for SUCCESSFUL or NOT
Public Sub PlayMacroFromFile(ByVal macroToPlay As String)
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open macroToPlay For Binary As #fileNum
        'Check to make sure this is actually a macro file
        Dim Macro_ID As String * 4
        Get #fileNum, 1, Macro_ID
        If (Macro_ID <> MACRO_IDENTIFIER) Then
            Close #fileNum
            Message "Invalid macro file."
            MsgBox macroToPlay & " is not a valid macro file.", vbOKOnly + vbCritical + vbApplicationModal, PROGRAMNAME & " Macro Error"
            Exit Sub
        End If
        'Now check to make sure that the version number is supported
        Dim Macro_Version As Long
        Get #fileNum, , Macro_Version
        'Check macro version incompatibility
        If (Macro_Version <> MACRO_VERSION_2006) Then
            'Attempt to save 2002 version macros
            If (Macro_Version = MACRO_VERSION_2002) Then
                Message "Converting outdated macro format..."
                Get #fileNum, , CurrentCall
                ReDim Calls(0 To CurrentCall) As ProcessCall
                'Temporary structure for playing old macros
                Dim OldCalls() As ProcessCall2002
                ReDim OldCalls(0 To CurrentCall) As ProcessCall2002
                Get #fileNum, , OldCalls
                'Loop through and copy our old macro structure into
                'the new format
                For x = 0 To CurrentCall
                    Calls(x).MainType = OldCalls(x).MainType
                    Calls(x).pOPCODE = OldCalls(x).pOPCODE
                    Calls(x).pOPCODE2 = OldCalls(x).pOPCODE2
                    Calls(x).pOPCODE3 = OldCalls(x).pOPCODE3
                    Calls(x).pOPCODE4 = OldCalls(x).pOPCODE4
                    Calls(x).LoadForm = OldCalls(x).LoadForm
                    Calls(x).RecordAction = OldCalls(x).RecordAction
                Next x
                'Once complete, close the old file, then copy it over
                'with a new version
                Close #fileNum
                
                Kill macroToPlay
                
                Dim newFileNum As Integer
                newFileNum = FreeFile
                
                Open macroToPlay For Binary As #newFileNum
                    Put #newFileNum, 1, MACRO_IDENTIFIER
                    Put #newFileNum, , MACRO_VERSION_2006
                    Put #newFileNum, , CurrentCall
                    Put #newFileNum, , Calls
                Close #newFileNum
                'Now this is a pretty stupid method for doing this,
                'but oh well: REOPEN the file and reorient the file pointer
                'correctly, allowing the routine to continue normally
                Open macroToPlay For Binary As #fileNum
                    Get #fileNum, 1, Macro_ID
                    Get #fileNum, , Macro_Version
                'Leave the If() block and continue normally
                Message "Macro converted successfully!  Continuing..."
            'If we make it here, we have an INCOMPATIBLE macro version
            Else
                Close #fileNum
                Message "Invalid macro version."
                MsgBox macroToPlay & " is no longer a supported macro version (#" & Macro_Version & ").", vbOKOnly + vbCritical + vbApplicationModal, PROGRAMNAME & " Macro Error"
                Exit Sub
            End If
        End If
    
        Get #fileNum, , CurrentCall
        ReDim Calls(0 To CurrentCall) As ProcessCall
        Get #fileNum, , Calls
        
    Close #fileNum
        
    'Now we run a loop through the macro structure, calling the software
    'processor with all the necessary information for each effect
    Message "Processing macro data..."
    Dim tCall As Long
    For tCall = 1 To CurrentCall
        If (Calls(tCall).LoadForm = False) Then
            Process Calls(tCall).MainType, Calls(tCall).pOPCODE, Calls(tCall).pOPCODE2, Calls(tCall).pOPCODE3, Calls(tCall).pOPCODE4, Calls(tCall).pOPCODE5, Calls(tCall).pOPCODE6, Calls(tCall).pOPCODE7, Calls(tCall).pOPCODE8, Calls(tCall).pOPCODE9, Calls(tCall).LoadForm, Calls(tCall).RecordAction
            Do
                DoEvents
            Loop While Processing = True
        End If
    Next tCall
    Message "Macro complete!"
    
End Sub
