VERSION 5.00
Begin VB.Form FormCustomFilter 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Custom Filter"
   ClientHeight    =   6540
   ClientLeft      =   150
   ClientTop       =   120
   ClientWidth     =   10395
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   693
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7380
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   8850
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin VB.TextBox txtBias 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   8880
      TabIndex        =   28
      Text            =   "1"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox TxtWeight 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6960
      TabIndex        =   27
      Text            =   "1"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   24
      Left            =   8520
      TabIndex        =   26
      Text            =   "0"
      Top             =   2640
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   23
      Left            =   7920
      TabIndex        =   25
      Text            =   "0"
      Top             =   2640
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   22
      Left            =   7320
      TabIndex        =   24
      Text            =   "0"
      Top             =   2640
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   21
      Left            =   6720
      TabIndex        =   23
      Text            =   "0"
      Top             =   2640
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   20
      Left            =   6120
      TabIndex        =   22
      Text            =   "0"
      Top             =   2640
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   19
      Left            =   8520
      TabIndex        =   21
      Text            =   "0"
      Top             =   2160
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   18
      Left            =   7920
      TabIndex        =   20
      Text            =   "0"
      Top             =   2160
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   17
      Left            =   7320
      TabIndex        =   19
      Text            =   "0"
      Top             =   2160
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   16
      Left            =   6720
      TabIndex        =   18
      Text            =   "0"
      Top             =   2160
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Index           =   15
      Left            =   6120
      TabIndex        =   17
      Text            =   "0"
      Top             =   2160
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   14
      Left            =   8520
      TabIndex        =   16
      Text            =   "0"
      Top             =   1680
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   13
      Left            =   7920
      TabIndex        =   15
      Text            =   "0"
      Top             =   1680
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   12
      Left            =   7320
      TabIndex        =   14
      Text            =   "1"
      Top             =   1680
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   11
      Left            =   6720
      TabIndex        =   13
      Text            =   "0"
      Top             =   1680
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Index           =   10
      Left            =   6120
      TabIndex        =   12
      Text            =   "0"
      Top             =   1680
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   9
      Left            =   8520
      TabIndex        =   11
      Text            =   "0"
      Top             =   1200
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   8
      Left            =   7920
      TabIndex        =   10
      Text            =   "0"
      Top             =   1200
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   7
      Left            =   7320
      TabIndex        =   9
      Text            =   "0"
      Top             =   1200
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   6
      Left            =   6720
      TabIndex        =   8
      Text            =   "0"
      Top             =   1200
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Index           =   5
      Left            =   6120
      TabIndex        =   7
      Text            =   "0"
      Top             =   1200
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   4
      Left            =   8520
      TabIndex        =   6
      Text            =   "0"
      Top             =   720
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   3
      Left            =   7920
      TabIndex        =   5
      Text            =   "0"
      Top             =   720
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   2
      Left            =   7320
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   1
      Left            =   6720
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   540
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   2
      Text            =   "0"
      Top             =   720
      Width           =   540
   End
   Begin PhotoDemon.jcbutton cmdOpen 
      Height          =   615
      Left            =   6120
      TabIndex        =   29
      Top             =   4800
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1085
      ButtonStyle     =   13
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormCustomFilter.frx":0000
      DisabledPictureMode=   1
      CaptionEffects  =   0
      ToolTip         =   "Open a previously saved convolution filter."
      TooltipType     =   1
      TooltipTitle    =   "Open Filter"
   End
   Begin PhotoDemon.jcbutton cmdSave 
      Height          =   615
      Left            =   7200
      TabIndex        =   30
      Top             =   4800
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1085
      ButtonStyle     =   13
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormCustomFilter.frx":1052
      DisabledPictureMode=   1
      CaptionEffects  =   0
      ToolTip         =   "Save the current filter to file.  This allows you to use the filter later, or share the filter with other PhotoDemon users."
      TooltipType     =   1
      TooltipTitle    =   "Save Current Filter"
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   36
      Top             =   5760
      Width           =   10455
   End
   Begin VB.Label lblAdditional 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "additional settings:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   6000
      TabIndex        =   35
      Top             =   3240
      Width           =   2010
   End
   Begin VB.Label lblOffset 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "offset:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   8040
      TabIndex        =   34
      Top             =   3735
      Width           =   645
   End
   Begin VB.Label lblDivisor 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "divisor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   6120
      TabIndex        =   32
      Top             =   3735
      Width           =   675
   End
   Begin VB.Label lblLoadSave 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "load / save filter data:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   6000
      TabIndex        =   33
      Top             =   4320
      Width           =   2325
   End
   Begin VB.Label lblConvolution 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "convolution matrix:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   6000
      TabIndex        =   31
      Top             =   240
      Width           =   2070
   End
End
Attribute VB_Name = "FormCustomFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Custom Filter Handler
'Copyright ©2000-2013 by Tanner Helland
'Created: 15/April/01
'Last updated: 08/September/12
'Last update: several routines from this form have been moved to the Filters_Area module, which is a more sensible place for them.
'
'This form handles creation/loading/saving of user-defined filters.
'
'***************************************************************************

Option Explicit

'When the user clicks OK...
Private Sub cmdOK_Click()
    
    'Before we do anything else, check to make sure every text box has a
    'valid number in it (no range checking is necessary)
    Dim x As Long, y As Long
    
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
    If Not NumberValid(txtBias) Then
        AutoSelectText txtBias
        Exit Sub
    End If
    
    Me.Visible = False
    
    'Copy the values from the text boxes into an array
    Message "Generating filter data..."
        
    g_FilterSize = 5
        
    ReDim g_FM(-2 To 2, -2 To 2) As Long
    
    For x = -2 To 2
    For y = -2 To 2
        g_FM(x, y) = Val(TxtF((x + 2) + (y + 2) * 5))
    Next y
    Next x
        
    'What to divide the final value by
    g_FilterWeight = Val(TxtWeight.Text)
    
    'Any offset value
    g_FilterBias = Val(txtBias.Text)
    
    'Set that we have created a filter during this program session, and save it accordingly
    g_HasCreatedFilter = True
    
    SaveCustomFilter g_UserPreferences.getTempPath & "~PD_CF.tmp"
    Process CustomFilter, g_UserPreferences.getTempPath & "~PD_CF.tmp"
    
    Unload Me
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
        
    'If a filter has been used previously, load it from the temp file
    If g_HasCreatedFilter = True Then OpenCustomFilter g_UserPreferences.getTempPath & "~PD_CF.tmp"
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Render a preview
    updatePreview
    
End Sub

Private Sub cmdOpen_Click()
    'Simple open dialog
    Dim CC As cCommonDialog
        
    Dim sFile As String
    Set CC = New cCommonDialog
    If CC.VBGetOpenFileName(sFile, , , , , True, PROGRAMNAME & " Filter (." & FILTER_EXT & ")|*." & FILTER_EXT & "|All files|*.*", , g_UserPreferences.getFilterPath, "Open a custom filter", , FormCustomFilter.hWnd, 0) Then
        If OpenCustomFilter(sFile) = True Then
            
            'Save the new directory as the default path for future usage
            g_UserPreferences.setFilterPath sFile
            
            'Redraw the preview
            updatePreview
            
        Else
            MsgBox "An error occurred while attempting to load " & sFile & ".  Please verify that the file is a valid custom filter file.", vbOKOnly + vbExclamation + vbApplicationModal, PROGRAMNAME & " Custom Filter Error"
        End If
    End If
    
End Sub

'Provide a save prompt, and use that to trigger a save of this custom filter to file
Private Sub cmdSave_Click()
    
    'Simple save dialog
    Dim CC As cCommonDialog
        
    Dim sFile As String
    Set CC = New cCommonDialog
    If CC.VBGetSaveFileName(sFile, , True, PROGRAMNAME & " Filter (." & FILTER_EXT & ")|*." & FILTER_EXT & "|All files|*.*", , g_UserPreferences.getFilterPath, "Save a custom filter", "." & FILTER_EXT, FormCustomFilter.hWnd, 0) Then
        
        'Save the new directory as the default path for future usage
        g_UserPreferences.setFilterPath sFile
        
        'Write out the file
        SaveCustomFilter sFile
        
    End If
    
End Sub

'Load a custom filter from file
Private Function OpenCustomFilter(ByRef srcFilterPath As String) As Boolean
    
    Dim tmpVal As Integer
    Dim tmpValLong As Long
    
    Dim x As Long
    
    'Open the specified path
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open srcFilterPath For Binary As #fileNum
        
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
            Get #fileNum, , tmpVal
            TxtWeight = tmpVal
            Get #fileNum, , tmpVal
            txtBias = tmpVal
        ElseIf VersionNumber = CUSTOM_FILTER_VERSION_2012 Then
            Get #fileNum, , tmpValLong
            TxtWeight = tmpValLong
            Get #fileNum, , tmpValLong
            txtBias = tmpValLong
        End If
        
        If VersionNumber = CUSTOM_FILTER_VERSION_2003 Then
            For x = 0 To 24
                Get #fileNum, , tmpVal
                TxtF(x) = tmpVal
            Next x
        ElseIf VersionNumber = CUSTOM_FILTER_VERSION_2012 Then
            For x = 0 To 24
                Get #fileNum, , tmpValLong
                TxtF(x) = tmpValLong
            Next x
        End If
        
    Close #fileNum
    
    OpenCustomFilter = True
    
End Function

'Save a custom filter to file
Private Function SaveCustomFilter(ByRef dstFilterPath As String) As Boolean

    If FileExist(dstFilterPath) Then Kill dstFilterPath
    
    Dim x As Long
    
    'Open the specified file
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open dstFilterPath For Binary As #fileNum
        'Place the identifier
        Put #fileNum, 1, CUSTOM_FILTER_ID
        'Place the current version number
        Put #fileNum, , CUSTOM_FILTER_VERSION_2012
        'Strip the information out of the text boxes and send it
        Dim tmpVal As Long
        tmpVal = Val(TxtWeight.Text)
        Put #fileNum, , tmpVal
        tmpVal = Val(txtBias.Text)
        Put #fileNum, , tmpVal
        For x = 0 To 24
            tmpVal = Val(TxtF(x).Text)
            Put #fileNum, , tmpVal
        Next x
    Close #fileNum
    SaveCustomFilter = True
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub TxtBias_GotFocus()
    AutoSelectText txtBias
End Sub

Private Sub txtBias_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtBias, True
    updatePreview
End Sub

Private Sub TxtF_GotFocus(Index As Integer)
    AutoSelectText TxtF(Index)
End Sub

Private Sub TxtF_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    textValidate TxtF(Index), True
    updatePreview
End Sub

Private Sub TxtWeight_GotFocus()
    AutoSelectText TxtWeight
End Sub

Private Sub TxtWeight_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate TxtWeight, True
    updatePreview
End Sub

'When the filter is changed, update the preview to match
Private Sub updatePreview()

    Dim x As Long, y As Long

    'We can only apply the preview if all loaded text boxes are valid
    For x = 0 To 24
        If Not EntryValid(TxtF(x), -1000000, 1000000, False, False) Then Exit Sub
    Next x
    
    If Not EntryValid(TxtWeight, -1000000, 1000000, False, False) Then Exit Sub
    
    If Not EntryValid(txtBias, -1000000, 1000000, False, False) Then Exit Sub
    
    'Copy the values from the text boxes into an array
    g_FilterSize = 5
        
    ReDim g_FM(-2 To 2, -2 To 2) As Long
    
    For x = -2 To 2
    For y = -2 To 2
        g_FM(x, y) = Val(TxtF((x + 2) + (y + 2) * 5))
    Next y
    Next x
        
    'What to divide the final value by
    g_FilterWeight = Val(TxtWeight.Text)
    
    'Offset value
    g_FilterBias = Val(txtBias.Text)
        
    'Apply the preview
    DoFilter "Preview", False, , True, fxPreview
    
End Sub

