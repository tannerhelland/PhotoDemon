VERSION 5.00
Begin VB.Form FormPatch 
   BackColor       =   &H80000005&
   Caption         =   " PhotoDemon Update"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   9360
      Top             =   120
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "FormPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH_LEN = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH_LEN
End Type

'APIs for making sure PhotoDemon.exe has terminated
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'APIs for restarting PhotoDemon.exe when we're done
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperationStr As Long, ByVal lpFileStr As Long, ByVal lpParametersStr As Long, ByVal lpDirectoryStr As Long, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

'Two paths are required for the update process: a path to PD's folder, and PD's update subfolder
Private m_PDPath As String, m_PDUpdatePath As String

'Once PhotoDemon.exe can no longer be detected as an active process, this will be set to TRUE
Private m_PDClosed As Boolean

'If the user asked PD to restart after patching, this app will be notified of the decision via command-line
Private m_RestartWhenDone As Boolean

'If the patch was successful, this will be set to TRUE
Private m_PatchSuccessful As Boolean

'PhotoDemon passes some values to us via command line:
Private m_StartPosition As Long, m_EndPosition As Long  'Start and end position of the relevant update track in the update XML file

'This program starts working as soon as it loads.  No user interaction is expected or handled.
Private Sub Form_Load()
    
    'Position the output text box
    txtOut.Width = FormPatch.ScaleWidth - txtOut.Left * 2
    
    'Check relevant command-line params
    parseCommandLine
    
    'Replace the crappy default VB icon
    SetIcon Me.hWnd, "AAA", True
    
    'Wait for PD to close; when it does, the timer will initiate the rest of the patch process.
    txtOut.Text = "Waiting for PhotoDemon to terminate..."
    m_PDClosed = False
    tmrCheck.Enabled = True
    
End Sub

'Parse the command line for all relevant instructions.  PD handles some update tasks for us, and it relays its findings through
' the command line.
Private Sub parseCommandLine()
    
    'Split params according to spaces
    Dim allParams() As String
    allParams = Split(Command$, " ")
    
    Dim curLine As Long
    curLine = LBound(allParams)
    
    'Iterate through the params, looking for meaningful entries as we go
    Do While curLine <= UBound(allParams)
        
        'Start checking instructions of interest
        If stringsEqual(allParams(curLine), "/restart") Then
            m_RestartWhenDone = True
        
        ElseIf stringsEqual(allParams(curLine), "/start") Then
            
            'Retrieve the start position
            curLine = curLine + 1
            m_StartPosition = CLng(allParams(curLine))
            
        ElseIf stringsEqual(allParams(curLine), "/end") Then
            
            'Retrieve the start position
            curLine = curLine + 1
            m_EndPosition = CLng(allParams(curLine))
        
        End If
        
        'Increment to the next line and continue checking params
        curLine = curLine + 1
        
    Loop
    
End Sub

'Shortcut function for checking string equality
Private Function stringsEqual(ByVal strOne As String, ByVal strTwo As String) As Boolean
    stringsEqual = (StrComp(Trim$(strOne), Trim$(strTwo), vbBinaryCompare) = 0)
End Function

Private Sub tmrCheck_Timer()

    'Check to see if PD has closed.
    If (Not m_PDClosed) Then
    
        Dim pdFound As Boolean
        pdFound = False
        
        'Prepare to iterate through all running processes
        Const TH32CS_SNAPPROCESS As Long = 2&
        Const PROCESS_ALL_ACCESS = 0
        Dim uProcess As PROCESSENTRY32
        Dim rProcessFound As Long, hSnapshot As Long, myProcess As Long
        Dim szExename As String
        Dim i As Long
        
        On Local Error GoTo PDDetectionError
    
        'Prepare a generic process reference
        uProcess.dwSize = Len(uProcess)
        hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
        rProcessFound = ProcessFirst(hSnapshot, uProcess)
        
        'Iterate through all running processes, looking for PhotoDemon instances
        Do While rProcessFound
    
            'Retrieve the EXE name of this process
            i = InStr(1, uProcess.szexeFile, Chr(0))
            szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            
            'If the process name is "exiftool.exe", terminate it
            If Right$(szExename, Len("PhotoDemon.exe")) = "PhotoDemon.exe" Then
                
                pdFound = True
                Exit Do
                 
            End If
            
            'Find the next process, then continue
            rProcessFound = ProcessNext(hSnapshot, uProcess)
        
        Loop
    
        'Release our generic process snapshot
        CloseHandle hSnapshot
    
        'If PD was found, do nothing.  Otherwise, start patching the program.
        If Not pdFound Then
            
            'Disable this timer
            tmrCheck.Enabled = False
            
            'Start the patch process
            m_PDClosed = True
            startPatching
            
        End If
    
        Exit Sub
    
    End If
    
PDDetectionError:

    textOut "Unknown error occurred while waiting for PhotoDemon to close.  Checking again..."

End Sub

'Start the patch process
Private Sub startPatching()
    
    textOut "PhotoDemon shutdown detected.  Starting patch process."
    
    'This update patcher will have been extracted to PD's root folder.
    m_PDPath = App.Path
    
    If StrComp(Right$(m_PDPath, 1), "\", vbBinaryCompare) <> 0 Then m_PDPath = m_PDPath & "\"
    m_PDUpdatePath = m_PDPath & "Data\Updates\"
    
    'Retrieve the patch XML file from its hard-coded location
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    If xmlEngine.loadXMLFile(m_PDUpdatePath & "patch.xml") Then
        
        If xmlEngine.isPDDataType("Program version") Then
        
            m_PatchSuccessful = True
            
        Else
            textOut "Update XML file doesn't contain patch data.  Patching cannot proceed.", False
            m_PatchSuccessful = False
        End If
        
    Else
        textOut "Update XML file wasn't found.  Patching cannot proceed.", False
        m_PatchSuccessful = False
    End If
    
    'Regardless of outcome, perform some clean-up afterward.
    finishPatching
    
End Sub

'Regardless of patch success or failure, this function is called.  If the user wants us to restart PD, we do so now.
Private Sub finishPatching()

    If m_RestartWhenDone Then
        
        Dim actionString As String, fileString As String, pathString As String, paramString As String
        actionString = "open"
        fileString = "PhotoDemon.exe"
        pathString = m_PDPath
        paramString = ""
        
        ShellExecute 0&, StrPtr(actionString), StrPtr(fileString), 0&, StrPtr(pathString), SW_SHOWNORMAL
    
    End If

End Sub

'Display basic update text
Public Sub textOut(ByVal newText As String, Optional ByVal appendEllipses As Boolean = True)
    
    If appendEllipses Then
    
        If StrComp(Right$(newText, 1), ".", vbBinaryCompare) = 0 Then
            txtOut.Text = txtOut.Text & vbCrLf & newText & ".."
        Else
            txtOut.Text = txtOut.Text & vbCrLf & newText & "..."
        End If
        
    Else
        txtOut.Text = txtOut.Text & vbCrLf & newText
    End If
    
    'Stick the cursor at the end of the text, which looks more natural IMO
    txtOut.SelStart = Len(txtOut.Text)
    
End Sub
