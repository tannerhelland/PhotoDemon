VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "VB6 full project search-and-replace"
   ClientHeight    =   9690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   646
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "I accept full responsibility for whatever happens next.  Apply search and replace."
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   8760
      Width           =   9255
   End
   Begin VB.CheckBox chkAllow 
      BackColor       =   &H80000005&
      Caption         =   "User controls"
      Height          =   255
      Index           =   3
      Left            =   6960
      TabIndex        =   15
      Top             =   6000
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkAllow 
      BackColor       =   &H80000005&
      Caption         =   "Classes"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   14
      Top             =   6000
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkAllow 
      BackColor       =   &H80000005&
      Caption         =   "Modules"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   13
      Top             =   6000
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkAllow 
      BackColor       =   &H80000005&
      Caption         =   "Forms"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   12
      Top             =   6000
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkMatchCase 
      BackColor       =   &H80000005&
      Caption         =   "Match case"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   6840
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.TextBox txtReplace 
      Height          =   1695
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Text            =   "frmMain.frx":0000
      Top             =   3840
      Width           =   8055
   End
   Begin VB.TextBox txtFind 
      Height          =   1695
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "frmMain.frx":001D
      Top             =   2040
      Width           =   8055
   End
   Begin VB.CommandButton cmdLoadProject 
      Caption         =   "Load and validate VBP file"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox txtVBP 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Text            =   "C:\PhotoDemon v4\PhotoDemon\PhotoDemon.vbp"
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton cmdSelectVBP 
      Caption         =   "..."
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmMain.frx":0035
      ForeColor       =   &H005D4FFF&
      Height          =   1095
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   7560
      Width           =   9285
   End
   Begin VB.Line lnSeparator 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   8
      X2              =   624
      Y1              =   496
      Y2              =   496
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search options:"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   11
      Top             =   6480
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File types to search:"
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   10
      Top             =   5640
      Width           =   1725
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2070
      Width           =   405
   End
   Begin VB.Label lblReady 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Search and replace is NOT READY.  Please load a valid VB6 project."
      ForeColor       =   &H002818E8&
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   9255
   End
   Begin VB.Line lnSeparator 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   8
      X2              =   624
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Target VBP:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'List of files inside the VBP.  These files will be searched and replaced.
Private m_ListOfFiles As pdStringStack

'Type of each file.  This is easier to parse during the load step, so we simply store it in a mirrored stack.
Private m_TypeOfFiles As pdStringStack

Private Sub cmdLoadProject_Click()
    
    'Load the file into a string array, split by line delimiter
    Dim cFSO As pdFSO
    Set cFSO = New pdFSO
    
    Dim vbpPath As String, vbpBaseFolder As String
    vbpPath = txtVBP.Text
    vbpBaseFolder = cFSO.GetPathOnly(vbpPath)
    
    Dim vbpContents As String
    If cFSO.LoadTextFileAsString(vbpPath, vbpContents) Then
        
        'Split the incoming text into individual lines
        Dim vbpText() As String
        vbpText = Split(vbpContents, vbCrLf)
        
        If m_ListOfFiles Is Nothing Then Set m_ListOfFiles = New pdStringStack Else m_ListOfFiles.ResetStack
        If m_TypeOfFiles Is Nothing Then Set m_TypeOfFiles = New pdStringStack Else m_TypeOfFiles.ResetStack
        
        Dim numOfFiles As Long
        numOfFiles = 0
        
        'Extract only relevant file paths from the VBP
        Dim i As Long
        For i = 0 To UBound(vbpText)
        
            'Check for forms
            If InStr(1, vbpText(i), "Form=", vbBinaryCompare) = 1 Then
                m_ListOfFiles.AddString vbpBaseFolder & Right$(vbpText(i), Len(vbpText(i)) - 5)
                m_TypeOfFiles.AddString "Form"
                numOfFiles = numOfFiles + 1
            End If
            
            'Check for user controls
            If InStr(1, vbpText(i), "UserControl=", vbBinaryCompare) = 1 Then
                m_ListOfFiles.AddString vbpBaseFolder & Right$(vbpText(i), Len(vbpText(i)) - 12)
                m_TypeOfFiles.AddString "UserControl"
                numOfFiles = numOfFiles + 1
            End If
            
            'Check for modules
            If InStr(1, vbpText(i), "Module=", vbBinaryCompare) = 1 Then
                m_ListOfFiles.AddString vbpBaseFolder & Trim$(Right$(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), ";")))
                m_TypeOfFiles.AddString "Module"
                numOfFiles = numOfFiles + 1
            End If
            
            'Check for classes
            If InStr(1, vbpText(i), "Class=", vbBinaryCompare) = 1 Then
                m_ListOfFiles.AddString vbpBaseFolder & Trim$(Right$(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), ";")))
                m_TypeOfFiles.AddString "Class"
                numOfFiles = numOfFiles + 1
            End If
            
        Next i
        
        lblReady.Caption = "Search is replace is READY!  " & numOfFiles & " files found in this project."
        lblReady.ForeColor = RGB(37, 173, 26)
        
    Else
        lblReady.Caption = "Search is replace is NOT READY.  Please load a valid VB6 project."
        lblReady.ForeColor = RGB(232, 24, 40)
    End If
    
End Sub

Private Sub cmdSelectVBP_Click()
    Dim tmpCommonDialog As pdOpenSaveDialog
    Set tmpCommonDialog = New pdOpenSaveDialog
    
    Dim initFile As String
    initFile = txtVBP.Text
    If tmpCommonDialog.GetOpenFileName(initFile, , True, False, "VBP (*.vbp)|*.vbp") Then
        txtVBP.Text = initFile
    End If
End Sub

Private Sub cmdStart_Click()
    
    Dim cFSO As pdFSO
    Set cFSO = New pdFSO
    
    'Parse each file in turn
    Dim allowedToSearch As Boolean
    Dim i As Long
    Dim curFileName As String, curFileContents As String, strFind As String, strReplace As String
    Dim compareMode As VbCompareMethod
    strFind = txtFind
    strReplace = txtReplace
    If CBool(chkMatchCase) Then compareMode = vbBinaryCompare Else compareMode = vbTextCompare
    
    For i = 0 To m_ListOfFiles.GetNumOfStrings - 1
        
        cmdStart.Caption = "Processing file " & CStr(i + 1) & " of " & m_ListOfFiles.GetNumOfStrings
        
        'See if we're allowed to search this type of file
        allowedToSearch = False
        If CBool(chkAllow(0)) And StrComp(m_TypeOfFiles.GetString(i), "Form", vbBinaryCompare) = 0 Then allowedToSearch = True
        If CBool(chkAllow(1)) And StrComp(m_TypeOfFiles.GetString(i), "Module", vbBinaryCompare) = 0 Then allowedToSearch = True
        If CBool(chkAllow(2)) And StrComp(m_TypeOfFiles.GetString(i), "Class", vbBinaryCompare) = 0 Then allowedToSearch = True
        If CBool(chkAllow(3)) And StrComp(m_TypeOfFiles.GetString(i), "UserControl", vbBinaryCompare) = 0 Then allowedToSearch = True
        
        'Attempt to load the source file
        If allowedToSearch Then
            curFileName = m_ListOfFiles.GetString(i)
            If cFSO.LoadTextFileAsString(curFileName, curFileContents, False) Then
        
                'See if the "find" string occurs at least once
                If InStr(1, curFileContents, strFind, compareMode) > 0 Then
                
                    'Replace all occurrences of the "find" string
                    curFileContents = Replace$(curFileContents, strFind, strReplace, , , compareMode)
                    
                    'VB likes to add trailing linebreaks; remove these, if any
                    curFileContents = RemoveTrailingLinebreaks(curFileContents) & vbCrLf
                    
                    'Overwrite the original file
                    If Not cFSO.SaveStringToTextFile(curFileContents, curFileName, True, False) Then
                        Debug.Print "WARNING!  Failed to save new contents of "; curFileName & ".  Please investigate."
                    End If
                
                End If
            
            End If
        End If
        
    Next i
    
    cmdStart.Caption = "I accept full responsibility for whatever happens next.  Apply search and replace."
    
End Sub

Private Function RemoveTrailingLinebreaks(ByRef srcString As String) As String
    
    Dim keepRemoving As Boolean: keepRemoving = True
    
    Do
        If (InStrRev(srcString, vbCrLf, , vbBinaryCompare) = Len(srcString) - 1) Then
            srcString = Left$(srcString, Len(srcString) - 2)
        Else
            keepRemoving = False
        End If
    Loop While keepRemoving
    
    RemoveTrailingLinebreaks = srcString
    
End Function
