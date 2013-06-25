Attribute VB_Name = "Macro_and_Batch_Handler"
'***************************************************************************
'Macro and Batch Handler
'Copyright ©2000-2013 by Tanner Helland
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Macro loading information

'OLD outdated macro versions (included only to preserve functionality)
Public Const MACRO_IDENTIFIER As String * 4 = "DSmf"
Public Const MACRO_VERSION_2013 As Long = &H2013
Public Const MACRO_VERSION_2006 As Long = &H2006
Public Const MACRO_VERSION_2002 As Long = &H80000000

Public Type OldProcessCall
    MainType As Long
    pOPCODE As Variant
    pOPCODE2 As Variant
    pOPCODE3 As Variant
    pOPCODE4 As Variant
    pOPCODE5 As Variant
    pOPCODE6 As Variant
    pOPCODE7 As Variant
    pOPCODE8 As Variant
    pOPCODE9 As Variant
    LoadForm As Boolean
    recordAction As Boolean
End Type

'Array of processor calls - tracks what is going on
'Public OldCalls() As OldProcessCall

'Macro recording information
Public MacroStatus As Byte
Public Const MacroSTOP As Long = 0
Public Const MacroSTART As Long = 1
Public Const MacroBATCH As Long = 2
Public Const MacroCANCEL As Long = 128
Public MacroMessage As String

Public Sub StartMacro()

    'REMOVE BEFORE 5.6
    Dim cancelRecording As VbMsgBoxResult
    cancelRecording = pdMsgBox("Macro recording is currently being overhauled in preparation for PhotoDemon 5.6.  As such, it is very unstable, and I DO NOT recommend recording macros with this development build." & vbCrLf & vbCrLf & "If you want to risk it, press OK to continue.  Otherwise, press CANCEL to return to regular editing mode.", vbOKCancel + vbExclamation + vbApplicationModal, "Macros unstable in this build")

    If cancelRecording = vbCancel Then Exit Sub
    
    'Set the program-wide "recording" flag
    MacroStatus = MacroSTART
    
    'Resize the array that will hold the macro data
    ProcessCount = 1
    ReDim Processes(0 To ProcessCount) As ProcessCall
    
    'Notify the user that recording has begun
    Message "Macro recording started."
    FormMain.lblRecording.Visible = True
    
    FormMain.MnuStartMacroRecording.Enabled = False
    FormMain.MnuStopMacroRecording.Enabled = True

End Sub

Public Sub StopMacro()

    MacroStatus = MacroSTOP
    Message "Macro recording stopped."
    
    FormMain.lblRecording.Visible = False
    FormMain.MnuStartMacroRecording.Enabled = True
    FormMain.MnuStopMacroRecording.Enabled = False
    
    'Automatically launch the save macro data routine
    Dim CC As cCommonDialog
    Dim sFile As String
    Set CC = New cCommonDialog
    
    Dim cdFilter As String
    cdFilter = PROGRAMNAME & " " & g_Language.TranslateMessage("Macro Data") & " (." & MACRO_EXT & ")|*." & MACRO_EXT
            
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Save macro data")
    
    'If the user cancels the save dialog, give them another chance to save - just in case
    Dim mReturn As VbMsgBoxResult
     
SaveMacroAgain:
     
    'If we get the data we want, save the information
    If CC.VBGetSaveFileName(sFile, , True, cdFilter, , g_UserPreferences.getMacroPath, cdTitle, "." & MACRO_EXT, FormMain.hWnd, 0) Then
        
        'Save this macro's directory as the default macro path
        g_UserPreferences.setMacroPath sFile

        'Delete any existing file (overwrite) and dump the info to file
        If FileExist(sFile) Then Kill sFile
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        Open sFile For Binary As #fileNum
            Put #fileNum, 1, MACRO_IDENTIFIER
            Put #fileNum, , MACRO_VERSION_2013
            'Remove the last call (stop macro recording - redundant to save that)
            ProcessCount = ProcessCount - 1
            Put #fileNum, , ProcessCount
            ReDim Preserve Processes(ProcessCount) As ProcessCall
            Put #fileNum, , Processes
        Close #fileNum
        
        Message "Macro saved successfully."
        
    Else
        
        mReturn = pdMsgBox("If you do not save this macro, all actions recorded during this session will be permanently lost.  Are you sure you want to cancel?" & vbCrLf & vbCrLf & "(Press No to return to the Save Macro screen.  Note that you can always delete this macro later if you decide you don't want it.)", vbApplicationModal + vbExclamation + vbYesNo, "Warning: Last Chance to Save Macro")
        If mReturn = vbNo Then GoTo SaveMacroAgain
        
        Message "Macro abandoned."
        
    End If
        
    ProcessCount = 0
    
End Sub

Public Sub PlayMacro()

    'Automatically launch the load Macro data routine
    Dim CC As cCommonDialog
    Dim sFile As String
    Set CC = New cCommonDialog
    
    Dim cdFilter As String
    cdFilter = PROGRAMNAME & " " & g_Language.TranslateMessage("Macro Data") & " (." & MACRO_EXT & ")|*." & MACRO_EXT
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Open Macro File")
    
    'If we get a path, load that file
    If CC.VBGetOpenFileName(sFile, , , , , True, cdFilter, , g_UserPreferences.getMacroPath, cdTitle, "." & MACRO_EXT, FormMain.hWnd, OFN_HIDEREADONLY) Then
        
        Message "Loading macro data..."
        
        'Save this macro's folder as the default macro path
        g_UserPreferences.setMacroPath sFile
                
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
            pdMsgBox "%1 is not a valid macro file.", vbOKOnly + vbExclamation + vbApplicationModal, " Macro Error", macroToPlay
            Exit Sub
        End If
        
        'Now check to make sure that the version number is supported
        Dim Macro_Version As Long
        Get #fileNum, , Macro_Version
        
        'Check macro version incompatibility
        If (Macro_Version <> MACRO_VERSION_2013) Then
        
            'Attempt to save 2006 version macros
            If (Macro_Version = MACRO_VERSION_2006) Then
                Message "Converting outdated macro format..."
                Get #fileNum, , ProcessCount
                ReDim Processes(0 To ProcessCount) As ProcessCall
            
                'Temporary structure for playing old macros
                Dim OldProcesses() As OldProcessCall
                ReDim OldProcesses(0 To ProcessCount) As OldProcessCall
                Get #fileNum, , OldProcesses
            
                'Loop through and copy the old macro structure into the new format
                Dim x As Long
            
                For x = 0 To ProcessCount
                    Processes(x).ID = GetNameOfProcess(OldProcesses(x).MainType)
                    Processes(x).Dialog = OldProcesses(x).LoadForm
                    Processes(x).Recorded = OldProcesses(x).recordAction
                    Processes(x).MakeUndo = Not OldProcesses(x).LoadForm
                    Processes(x).Parameters = buildParams(OldProcesses(x).pOPCODE, OldProcesses(x).pOPCODE2, OldProcesses(x).pOPCODE3, OldProcesses(x).pOPCODE4, OldProcesses(x).pOPCODE5, , OldProcesses(x).pOPCODE6, , OldProcesses(x).pOPCODE7, , OldProcesses(x).pOPCODE8, , OldProcesses(x).pOPCODE9)
                Next x
            
                'Once complete, close the old file, then copy it over with a new version
                Close #fileNum
            
                Kill macroToPlay
            
                Dim newFileNum As Integer
                newFileNum = FreeFile
            
                Open macroToPlay For Binary As #newFileNum
                    Put #newFileNum, 1, MACRO_IDENTIFIER
                    Put #newFileNum, , MACRO_VERSION_2013
                    Put #newFileNum, , ProcessCount
                    Put #newFileNum, , Processes
                Close #newFileNum
            
                'Now this is a pretty stupid method for doing this, but oh well: REOPEN the file and reorient the
                ' file pointer correctly, allowing the routine to continue normally
                Open macroToPlay For Binary As #fileNum
                    Get #fileNum, 1, Macro_ID
                    Get #fileNum, , Macro_Version
            
                'Leave the If() block and continue normally
                Message "Old macro converted successfully!  Continuing..."
                
            'If we make it here, we have an INCOMPATIBLE macro version
            Else
                Close #fileNum
                Message "Invalid macro version."
                pdMsgBox "%1 is no longer a supported macro version (#%2).", vbOKOnly + vbExclamation + vbApplicationModal, " Macro Error", macroToPlay, Macro_Version
                Exit Sub
            End If
        End If
    
        Get #fileNum, , ProcessCount
        ReDim Processes(0 To ProcessCount) As ProcessCall
        Get #fileNum, , Processes
        
    Close #fileNum
        
    'Now we run a loop through the macro structure, calling the software
    'processor with all the necessary information for each effect
    Message "Processing macro data..."
    Dim tProc As Long
    For tProc = 1 To ProcessCount
        If Not Processes(tProc).Dialog Then
            Process Processes(tProc).ID, Processes(tProc).Dialog, Processes(tProc).Parameters, Processes(tProc).MakeUndo, Processes(tProc).Tool, Processes(tProc).Recorded
            'Do
            '    DoEvents
            'Loop While Processing = True
        End If
    Next tProc
    Message "Macro complete!"
    
End Sub

'Return a string with a human-readable name of a given process ID.

' This function was necessary in old versions of PhotoDemon to generate custom text for things like the Undo menu
' (e.g. "Undo <name here>").  As of version 5.6, PhotoDemon uses only strings as process IDs, so this function is
' no longer necessary.

' However, old PhotoDemon macros stored only numerical process IDs, so in order to keep the software backwards
' compatible, this function is still included to translate old process ID values into new ones.
Private Function GetNameOfProcess(ByVal processID As Long) As String

    Select Case processID
    
        'Main functions (not used for image editing); numbers 1-99
        Case 1
            GetNameOfProcess = "Open"
        Case 2
            GetNameOfProcess = "Save"
        Case 3
            GetNameOfProcess = "Save as"
        Case 10
            GetNameOfProcess = "Screen capture"
        Case 20
            GetNameOfProcess = "Copy to clipboard"
        Case 21
            GetNameOfProcess = "Paste as new image"
        Case 22
            GetNameOfProcess = "Empty clipboard"
        Case 30
            GetNameOfProcess = "Undo"
        Case 31
            GetNameOfProcess = "Redo"
        Case 40
            GetNameOfProcess = "Start macro recording"
        Case 41
            GetNameOfProcess = "Stop macro recording"
        Case 42
            GetNameOfProcess = "Play macro"
        Case 50
            GetNameOfProcess = "Select scanner or camera"
        Case 51
            GetNameOfProcess = "Scan image"
            
        'Histogram functions; numbers 100-199
        Case 100
            GetNameOfProcess = "Display histogram"
        Case 101
            GetNameOfProcess = "Stretch histogram"
        Case 102
            GetNameOfProcess = "Equalize"
        Case 104
            GetNameOfProcess = "White Balance"
            
        'Black/White conversion; numbers 200-299
        Case 200
            GetNameOfProcess = "Color to monochrome"
        Case 201
            GetNameOfProcess = "Monochrome to grayscale"
            
        'Grayscale conversion; numbers 300-399
        Case 300
            GetNameOfProcess = "Desaturate"
        Case 301
            GetNameOfProcess = "Grayscale (ITU standard)"
        Case 302
            GetNameOfProcess = "Grayscale (average)"
        Case 303
            GetNameOfProcess = "Grayscale (custom # of colors)"
        Case 304
            GetNameOfProcess = "Grayscale (custom dither)"
        Case 305
            GetNameOfProcess = "Grayscale (decomposition)"
        Case 306
            GetNameOfProcess = "Grayscale (single channel)"
        
        'Area filters; numbers 400-499
        Case 400
            GetNameOfProcess = "Gaussian blur"
        Case 401
            GetNameOfProcess = "Gaussian blur"
        Case 402
            GetNameOfProcess = "Gaussian blur"
        Case 403
            GetNameOfProcess = "Gaussian blur"
        Case 404
            GetNameOfProcess = "Sharpen"
        Case 405
            GetNameOfProcess = "Sharpen more"
        Case 406
            GetNameOfProcess = "Unsharp mask"
        Case 407
            GetNameOfProcess = "Diffuse"
        Case 408
            GetNameOfProcess = "Diffuse"
        Case 409
            GetNameOfProcess = "Diffuse"
        Case 410
            GetNameOfProcess = "Pixelate"
        Case 412
            GetNameOfProcess = "Dilate (maximum rank)"
        Case 411
            GetNameOfProcess = "Erode (minimum rank)"
        Case 415
            GetNameOfProcess = "Grid blur"
        Case 416
            GetNameOfProcess = "Gaussian blur"
        Case 417
            GetNameOfProcess = "Smart blur"
        Case 418
            GetNameOfProcess = "Box blur"
        
        'Edge filters; numbers 500-599
        Case 500
            GetNameOfProcess = "Emboss"
        Case 501
            GetNameOfProcess = "Engrave"
        Case 504
            GetNameOfProcess = "Pencil drawing"
        Case 505
            GetNameOfProcess = "Relief"
        Case 506
            GetNameOfProcess = "Find Edges (Prewitt Horizontal)"
        Case 507
            GetNameOfProcess = "Find Edges (Prewitt Vertical)"
        Case 508
            GetNameOfProcess = "Find Edges (Sobel Horizontal)"
        Case 509
            GetNameOfProcess = "Find Edges (Sobel Vertical)"
        Case 510
            GetNameOfProcess = "Find Edges (Laplacian)"
        Case 511
            GetNameOfProcess = "Artistic Contour"
        Case 512
            GetNameOfProcess = "Find Edges (Hilite)"
        Case 513
            GetNameOfProcess = "Find Edges (PhotoDemon Linear)"
        Case 514
            GetNameOfProcess = "Find Edges (PhotoDemon Cubic)"
        Case 515
            GetNameOfProcess = "Edge Enhance"
        Case 516
            GetNameOfProcess = "Trace Contour"
            
        'Color operations; numbers 600-699
        Case 600
            GetNameOfProcess = "Rechannel"
        'Rechannel Green and Red are only included for legacy reasons
        Case 601
            GetNameOfProcess = "Rechannel"
        Case 602
            GetNameOfProcess = "Rechannel"
        '-------
        Case 603
            GetNameOfProcess = "Shift Colors (Left)"
        Case 604
            GetNameOfProcess = "Shift Colors (Right)"
        Case 605
            GetNameOfProcess = "Brightness and Contrast"
        Case 606
            GetNameOfProcess = "Gamma"
        Case 607
            GetNameOfProcess = "Invert RGB"
        Case 608
            GetNameOfProcess = "Invert Hue"
        Case 609
            GetNameOfProcess = "Film Negative"
        Case 617
            GetNameOfProcess = "Compound Invert"
        Case 610
            GetNameOfProcess = "Auto-Enhance Contrast"
        Case 611
            GetNameOfProcess = "Auto-Enhance Highlights"
        Case 612
            GetNameOfProcess = "Auto-Enhance Midtones"
        Case 613
            GetNameOfProcess = "Auto-Enhance Shadows"
        Case 614
            GetNameOfProcess = "Levels"
        Case 615
            GetNameOfProcess = "Colorize"
        Case 616
            GetNameOfProcess = "Reduce Colors"
        Case 618
            GetNameOfProcess = "Color Temperature"
        Case 619
            GetNameOfProcess = "Hue and Saturation"
        Case 620
            GetNameOfProcess = "Color Balance"
        Case 621
            GetNameOfProcess = "Shadows and Highlights"
            
        'Coordinate filters/transformations; numbers 700-799
        Case 700
            GetNameOfProcess = "Resize"
        Case 701
            GetNameOfProcess = "Flip"
        Case 702
            GetNameOfProcess = "Mirror"
        Case 703
            GetNameOfProcess = "Rotate 90° Clockwise"
        Case 704
            GetNameOfProcess = "Rotate 180°"
        Case 705
            GetNameOfProcess = "Rotate 90° Counter-Clockwise"
        Case 706
            GetNameOfProcess = "Arbitrary Rotation"
        Case 707
            GetNameOfProcess = "Isometric Conversion"
        Case 708
            GetNameOfProcess = "Tile"
        Case 709
            GetNameOfProcess = "Crop"
        Case 710
            GetNameOfProcess = "Remove alpha channel"
        Case 711
            GetNameOfProcess = "Add alpha channel"
        Case 712
            GetNameOfProcess = "Swirl"
        Case 713
            GetNameOfProcess = "Apply lens distortion"
        Case 714
            GetNameOfProcess = "Correct lens distortion"
        Case 715
            GetNameOfProcess = "Ripple"
        Case 716
            GetNameOfProcess = "Pinch and whirl"
        Case 717
            GetNameOfProcess = "Waves"
        Case 718
            GetNameOfProcess = "Figured glass"
        Case 719
            GetNameOfProcess = "Kaleidoscope"
        Case 720
            GetNameOfProcess = "Polar conversion"
        Case 721
            GetNameOfProcess = "Autocrop"
        Case 722
            GetNameOfProcess = "Shear"
        Case 723
            GetNameOfProcess = "Squish"
        Case 724
            GetNameOfProcess = "Perspective"
        Case 725
            GetNameOfProcess = "Pan and Zoom"
            
        'Miscellaneous filters; numbers 800-899
        Case 803
            GetNameOfProcess = "Fade"
        Case 807
            GetNameOfProcess = "Unfade"
        Case 808
            GetNameOfProcess = "Atmosphere"
        Case 809
            GetNameOfProcess = "Freeze"
        Case 810
            GetNameOfProcess = "Lava"
        Case 811
            GetNameOfProcess = "Burn"
        Case 813
            GetNameOfProcess = "Water"
        Case 814
            GetNameOfProcess = "Steel"
        Case 815
            GetNameOfProcess = "Dream"
        Case 816
            GetNameOfProcess = "Alien"
        Case 817
            GetNameOfProcess = "Custom Filter"
        Case 818
            GetNameOfProcess = "Antique"
        Case 819
            GetNameOfProcess = "Blacklight"
        Case 820
            GetNameOfProcess = "Posterize"
        Case 821
            GetNameOfProcess = "Radioactive"
        Case 822
            GetNameOfProcess = "Solarize"
        Case 823
            GetNameOfProcess = "Twins"
        Case 824
            GetNameOfProcess = "Synthesize"
        Case 825
            GetNameOfProcess = "Add RGB Noise"
        Case 827
            GetNameOfProcess = "Count Image Colors"
        Case 828
            GetNameOfProcess = "Fog"
        Case 829
            GetNameOfProcess = "Rainbow"
        Case 830
            GetNameOfProcess = "Vibrate"
        Case 831
            GetNameOfProcess = "Despeckle"
        Case 832
            GetNameOfProcess = "Custom Despeckle"
        Case 840
            GetNameOfProcess = "Comic book"
        Case 826
            GetNameOfProcess = "Sepia"
        Case 833
            GetNameOfProcess = "Thermograph (Heat Map)"
        Case 841
            GetNameOfProcess = "Add Film Grain"
        
        Case 900
            GetNameOfProcess = "Repeat Last Action"
        Case 901
            GetNameOfProcess = "Fade last effect"
            
        Case 1000
            GetNameOfProcess = "Create New Selection"
        Case 1001
            GetNameOfProcess = "Clear Active Selection"
        Case 842
            GetNameOfProcess = "Vignetting"
        Case 843
            GetNameOfProcess = "Median"
        Case 844
            GetNameOfProcess = "Modern art"
            
        'This "Else" statement should never trigger, but if it does, return an empty string
        Case Else
            GetNameOfProcess = ""
            
    End Select
    
    GetNameOfProcess = g_Language.TranslateMessage(GetNameOfProcess)
    
End Function

