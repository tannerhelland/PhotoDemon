VERSION 5.00
Begin VB.Form FormCustomFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Custom Filter"
   ClientHeight    =   3000
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   2985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   199
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtBias 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2160
      TabIndex        =   30
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   28
      Top             =   2520
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   2520
      Width           =   1125
   End
   Begin VB.TextBox TxtWeight 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   720
      TabIndex        =   26
      Text            =   "1"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   24
      Left            =   2520
      TabIndex        =   24
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   23
      Left            =   1920
      TabIndex        =   23
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   22
      Left            =   1320
      TabIndex        =   22
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   21
      Left            =   720
      TabIndex        =   21
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   20
      Left            =   120
      TabIndex        =   20
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   19
      Left            =   2520
      TabIndex        =   19
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   18
      Left            =   1920
      TabIndex        =   18
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   17
      Left            =   1320
      TabIndex        =   17
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   16
      Left            =   720
      TabIndex        =   16
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   14
      Left            =   2520
      TabIndex        =   14
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   13
      Left            =   1920
      TabIndex        =   13
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   12
      Left            =   1320
      TabIndex        =   12
      Text            =   "1"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   11
      Left            =   720
      TabIndex        =   11
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   9
      Left            =   2520
      TabIndex        =   9
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   8
      Left            =   1920
      TabIndex        =   8
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   6
      Left            =   720
      TabIndex        =   6
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   4
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblBias 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bias:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   1680
      TabIndex        =   29
      Top             =   2085
      Width           =   360
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scale:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   120
      TabIndex        =   25
      Top             =   2085
      Width           =   480
   End
   Begin VB.Menu MnuOpen 
      Caption         =   "Open Filter"
   End
   Begin VB.Menu MnuSave 
      Caption         =   "Save Filter"
   End
End
Attribute VB_Name = "FormCustomFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Custom Filter Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 15/April/01
'Last updated: 05/June/12
'Last update: FilterWeight and FilterBias are now Long-type.  A new filter version (2012)
'             had to be created to accomodate the change.
'
'This form handles creation/loading/saving of user-defined filters.
'
'***************************************************************************

Option Explicit

Const CUSTOM_FILTER_ID As String * 4 = "DScf"
Const CUSTOM_FILTER_VERSION_2003 = &H80000000
Const CUSTOM_FILTER_VERSION_2012 = &H80000001

'When the user clicks OK...
Private Sub CmdOK_Click()
    
    'Before we do anything else, check to make sure every text box has a
    'valid number in it (no range checking is necessary)
    For x = 0 To 24
        If Not NumberValid(TxtF(x)) Then
            AutoSelectText TxtF(x)
            Exit Sub
        End If
    Next x
    If Not NumberValid(TxtWeight) Then
        AutoSelectText TxtWeight
        Exit Sub
    End If
    If Not NumberValid(TxtBias) Then
        AutoSelectText TxtBias
        Exit Sub
    End If
    
    Me.Visible = False
    FormMain.SetFocus
    'Copy the values from the text boxes into an array
    Message "Generating filter data..."
        FilterSize = 5
        ReDim FM(-2 To 2, -2 To 2) As Long
        For x = -2 To 2
        For y = -2 To 2
            FM(x, y) = val(TxtF((x + 2) + (y + 2) * 5))
        Next y
        Next x
'What to divide the final value by
    FilterWeight = val(TxtWeight.Text)
'Any offset value
    FilterBias = val(TxtBias.Text)
'Set that we have created a filter during this program session, and save it accordingly
    HasCreatedFilter = True
    SaveCustomFilter TempPath & "~W096THCF.tmp"
    Process CustomFilter, TempPath & "~W096THCF.tmp"
    Unload Me
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'LOAD form
Private Sub Form_Load()
    
    'If a filter has been used previously, load it from the temp file
    If HasCreatedFilter = True Then OpenCustomFilter TempPath & "~W096THCF.tmp"
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

Private Sub MnuOpen_Click()
    'Simple open dialog
    Dim CC As cCommonDialog
    
    'Get the last "open image" path from the INI file
    Dim tempPathString As String
    tempPathString = GetFromIni("Program Paths", "CustomFilter")
    
    Dim sFile As String
    Set CC = New cCommonDialog
    If CC.VBGetOpenFileName(sFile, , , , , True, PROGRAMNAME & " Filter (." & FILTER_EXT & ")|*." & FILTER_EXT & "|All files|*.*", , tempPathString, "Open a custom filter", , FormCustomFilter.HWnd, 0) Then
        If OpenCustomFilter(sFile) = True Then
            'Save the new directory as the default path for future usage
            tempPathString = sFile
            StripDirectory tempPathString
            WriteToIni "Program Paths", "CustomFilter", tempPathString
            Message "Custom filter loaded successfully."
        Else
            Me.Visible = False
            Message "Custom filter not loaded."
            MsgBox "An error occurred while attempting to load " & sFile & ".  Please verify that the file is a valid custom filter file.", vbOKOnly + vbCritical + vbApplicationModal, PROGRAMNAME & " Custom Filter Error"
            Me.Visible = True
        End If
    End If
End Sub

Private Sub MnuSave_Click()
    'Simple save dialog
    Dim CC As cCommonDialog
    
    'Get the last "open image" path from the INI file
    Dim tempPathString As String
    tempPathString = GetFromIni("Program Paths", "CustomFilter")
    
    Dim sFile As String
    Set CC = New cCommonDialog
    If CC.VBGetSaveFileName(sFile, , True, PROGRAMNAME & " Filter (." & FILTER_EXT & ")|*." & FILTER_EXT & "|All files|*.*", , tempPathString, "Save a custom filter", "." & FILTER_EXT, FormCustomFilter.HWnd, 0) Then
        'Save the new directory as the default path for future usage
        tempPathString = sFile
        StripDirectory tempPathString
        WriteToIni "Program Paths", "CustomFilter", tempPathString
 
        SaveCustomFilter sFile
    End If
End Sub

'This subroutine will load the data from a custom filter file straight into the array
Public Function LoadCustomFilterData(ByRef FilterPath As String) As Boolean
    
    'These are used to load values from the filter file; previously, they were integers, but in
    ' 2012 I changed them to Longs.  PhotoDemon loads both types.
    Dim TmpVal As Integer
    Dim TmpValLong As Long
    
    'Open the specified path
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open FilterPath For Binary As #fileNum
        
        'Verify that the filter is actually a valid filter file
        Dim VerifyID As String * 4
        Get #fileNum, 1, VerifyID
        If (VerifyID <> CUSTOM_FILTER_ID) Then
            Close #fileNum
            LoadCustomFilterData = False
            Exit Function
        End If
        'End verification
       
        'Next get the version number (gotta have this for backwards compatibility)
        Dim VersionNumber As Long
        Get #fileNum, , VersionNumber
        If (VersionNumber <> CUSTOM_FILTER_VERSION_2003) And (VersionNumber <> CUSTOM_FILTER_VERSION_2012) Then
            Message "Unsupported custom filter version."
            Close #fileNum
            LoadCustomFilterData = False
        End If
        'End version check
        
        If VersionNumber = CUSTOM_FILTER_VERSION_2003 Then
            Get #fileNum, , TmpVal
            FilterWeight = TmpVal
            Get #fileNum, , TmpVal
            FilterBias = TmpVal
        ElseIf VersionNumber = CUSTOM_FILTER_VERSION_2012 Then
            Get #fileNum, , TmpValLong
            FilterWeight = TmpValLong
            Get #fileNum, , TmpValLong
            FilterBias = TmpValLong
        End If
        
        'Resize the filter array to fit the default filter size
        FilterSize = 5
        ReDim FM(-2 To 2, -2 To 2) As Long
        'Dim a temporary array from which to load the array data
        Dim tFilterArray(0 To 24) As Long
        
        If VersionNumber = CUSTOM_FILTER_VERSION_2003 Then
            For x = 0 To 24
                Get #fileNum, , TmpVal
                tFilterArray(x) = TmpVal
            Next x
        ElseIf VersionNumber = CUSTOM_FILTER_VERSION_2012 Then
            For x = 0 To 24
                Get #fileNum, , TmpValLong
                tFilterArray(x) = TmpValLong
            Next x
        End If
        
        'Now dump the temporary array into the filter array
        For x = -2 To 2
        For y = -2 To 2
            FM(x, y) = tFilterArray((x + 2) + (y + 2) * 5)
        Next y
        Next x
    'Close the file up
    Close #fileNum
    LoadCustomFilterData = True
End Function

'Subroutine for loading a custom filter
Private Function OpenCustomFilter(ByRef FilterPath As String) As Boolean
    
    Dim TmpVal As Integer
    Dim TmpValLong As Long
    
    'Open the specified path
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open FilterPath For Binary As #fileNum
        
        'Verify that the filter is actually a valid filter file
        Dim VerifyID As String * 4
        Get #fileNum, 1, VerifyID
        If (VerifyID <> CUSTOM_FILTER_ID) Then
            Close #fileNum
            OpenCustomFilter = False
            Exit Function
        End If
        'End verification
       
        'Next get the version number (gotta have this for backwards compatibility)
        Dim VersionNumber As Long
        Get #fileNum, , VersionNumber
        If (VersionNumber <> CUSTOM_FILTER_VERSION_2003) And (VersionNumber <> CUSTOM_FILTER_VERSION_2012) Then
            Message "Unsupported custom filter version."
            Close #fileNum
            OpenCustomFilter = False
            Exit Function
        End If
        'End version check
        
        If VersionNumber = CUSTOM_FILTER_VERSION_2003 Then
            Get #fileNum, , TmpVal
            TxtWeight = TmpVal
            Get #fileNum, , TmpVal
            TxtBias = TmpVal
        ElseIf VersionNumber = CUSTOM_FILTER_VERSION_2012 Then
            Get #fileNum, , TmpValLong
            TxtWeight = TmpValLong
            Get #fileNum, , TmpValLong
            TxtBias = TmpValLong
        End If
        
        If VersionNumber = CUSTOM_FILTER_VERSION_2003 Then
            For x = 0 To 24
                Get #fileNum, , TmpVal
                TxtF(x) = TmpVal
            Next x
        ElseIf VersionNumber = CUSTOM_FILTER_VERSION_2012 Then
            For x = 0 To 24
                Get #fileNum, , TmpValLong
                TxtF(x) = TmpValLong
            Next x
        End If
        
    Close #fileNum
    OpenCustomFilter = True
End Function

'Subroutine for saving a custom filter
Private Function SaveCustomFilter(ByRef FilterPath As String) As Boolean

    If FileExist(FilterPath) Then Kill FilterPath
    
    'Open the specified file
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open FilterPath For Binary As #fileNum
        'Place the identifier
        Put #fileNum, 1, CUSTOM_FILTER_ID
        'Place the current version number
        Put #fileNum, , CUSTOM_FILTER_VERSION_2012
        'Strip the information out of the text boxes and send it
        Dim TmpVal As Long
        TmpVal = val(TxtWeight.Text)
        Put #fileNum, , TmpVal
        TmpVal = val(TxtBias.Text)
        Put #fileNum, , TmpVal
        For x = 0 To 24
            TmpVal = val(TxtF(x).Text)
            Put #fileNum, , TmpVal
        Next x
    Close #fileNum
    SaveCustomFilter = True
    
End Function

Private Sub TxtBias_GotFocus()
    AutoSelectText TxtBias
End Sub

Private Sub TxtF_GotFocus(Index As Integer)
    AutoSelectText TxtF(Index)
End Sub

'This powerful routine runs a filter straight from a custom filter file to the active picture box - note that it has a counterpart in the
' FiltersArea module, which is used for running custom filters from within PhotoDemon (aka not from a file)
Public Sub DoCustomFilterFromFile(ByVal FilterPath As String)
    
    GetImageData
    
    'Load the custom filter data from file
    Message "Attempting to load custom filter file..."
    Dim FilterReturn As Boolean
    FilterReturn = LoadCustomFilterData(FilterPath)
    If FilterReturn = False Then Err.Raise 1024, PROGRAMNAME, "Invalid custom filter file"
    
    Message "Applying custom filter..."
    
    'C and D are like X and Y - they are additional loop variables used for sub-loops
    Dim c As Long, d As Long
    
    'CalcVar determines the size of each sub-loop (so that we don't waste time running a 5x5 matrix on 3x3 filters)
    Dim CalcVar As Long
    CalcVar = (FilterSize \ 2)
    
    'Temporary red, green, and blue values
    Dim tR As Long, tG As Long, tB As Long
    
    'iFM() will hold the contents of FM() - the filter matrix; I don't use FM in case other events want to access it
    Dim iFM() As Long
    
    'Resize iFM according to the size of the filter matrix, then copy over the contents of FM()
    If FilterSize = 3 Then ReDim iFM(-1 To 1, -1 To 1) As Long Else ReDim iFM(-2 To 2, -2 To 2) As Long
    iFM = FM
    
    'FilterWeightA and FilterBiasA are copies of the global FilterWeight and FilterBias variables; again, we don't use the originals in case other events
    ' want to access them
    Dim FilterWeightA As Long, FilterBiasA As Long
    FilterWeightA = FilterWeight
    FilterBiasA = FilterBias
    
    'FilterWeightTemp will be reset for every pixel, and decremented appropriately if
    'attempting to calculate the value for a pixel outside the perimeter of the image
    Dim FilterWeightTemp As Long
    
    'Temporary calculation variables
    Dim CalcX As Long, CalcY As Long
    
    'tData holds the processed image data; at the end of the filter processing it will get copied over the original image data
    ReDim tData(0 To (PicWidthL + 1) * 3 - 1, 0 To PicHeightL + 2) As Byte
    
    'TempRef is like QuickX below, but for sub-loops
    Dim TempRef As Long
    
    SetProgBarMax PicWidthL
    
    Dim QuickVal As Long
    
    'Now that we're ready, loop through the image, calculating pixel values as we go
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    
        'Reset our values upon beginning analysis on a new pixel
        tR = 0
        tG = 0
        tB = 0
        FilterWeightTemp = FilterWeightA
        
        'Run a sub-loop around the current pixel
        For c = x - CalcVar To x + CalcVar
            TempRef = c * 3
        For d = y - CalcVar To y + CalcVar
        
            CalcX = c - x
            CalcY = d - y
            
            'If no filter value is being applied to this pixel, ignore it (GoTo's aren't generally a part of good programming, but they ARE convenient :)
            If iFM(CalcX, CalcY) = 0 Then GoTo NextCustomFilterFromFilePixel
            
            'If this pixel lies outside the image perimeter, ignore it and adjust FilterWeight accordingly
            If c < 0 Or d < 0 Or c > PicWidthL Or d > PicHeightL Then
                FilterWeightTemp = FilterWeightTemp - iFM(CalcX, CalcY)
                GoTo NextCustomFilterFromFilePixel
            End If
            
            'Adjust red, green, and blue according to the values in the filter matrix (FM)
            tR = tR + (ImageData(TempRef + 2, d) * iFM(CalcX, CalcY))
            tG = tG + (ImageData(TempRef + 1, d) * iFM(CalcX, CalcY))
            tB = tB + (ImageData(TempRef, d) * iFM(CalcX, CalcY))

NextCustomFilterFromFilePixel:  Next d
        Next c
        
        'If a weight has been set, apply it now
        If (FilterWeightA <> 1) And (FilterWeightTemp <> 0) Then
            tR = tR \ FilterWeightTemp
            tG = tG \ FilterWeightTemp
            tB = tB \ FilterWeightTemp
        End If
        
        'If a bias has been specified, apply it now
        If FilterBiasA <> 0 Then
            tR = tR + FilterBiasA
            tG = tG + FilterBiasA
            tB = tB + FilterBiasA
        End If
        
        'Make sure all values are between 0 and 255
        ByteMeL tR
        ByteMeL tG
        ByteMeL tB
        'Finally, remember the new value in our tData array
        tData(QuickVal + 2, y) = tR
        tData(QuickVal + 1, y) = tG
        tData(QuickVal, y) = tB
        
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    
    SetProgBarVal cProgBar.Max
    
    'Copy tData over the original pixel data
    TransferImageData
    
    'Draw the updated image to the screen
    SetImageData
    
End Sub

Private Sub TxtWeight_GotFocus()
    AutoSelectText TxtWeight
End Sub
