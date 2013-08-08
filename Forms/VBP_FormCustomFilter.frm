VERSION 5.00
Begin VB.Form FormCustomFilter 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Custom Filter"
   ClientHeight    =   6540
   ClientLeft      =   150
   ClientTop       =   120
   ClientWidth     =   9960
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
   ScaleWidth      =   664
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.smartCheckBox chkNormalize 
      Height          =   480
      Left            =   6360
      TabIndex        =   36
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   847
      Caption         =   "normalize"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save filter"
      Height          =   1200
      Left            =   8040
      TabIndex        =   35
      ToolTipText     =   "Save the current filter to file.  This allows you to use the filter later, or share the filter with other PhotoDemon users."
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Load filter"
      Height          =   1200
      Left            =   6120
      TabIndex        =   34
      ToolTipText     =   "Open a previously saved convolution filter."
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7020
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   8490
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin VB.TextBox txtOffset 
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
      Left            =   8280
      TabIndex        =   28
      Text            =   "1"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtDivisor 
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
      Left            =   6360
      TabIndex        =   27
      Text            =   "1"
      Top             =   3840
      Width           =   1215
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
      Left            =   8760
      TabIndex        =   26
      Text            =   "0"
      Top             =   2520
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
      Left            =   8160
      TabIndex        =   25
      Text            =   "0"
      Top             =   2520
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
      Left            =   7560
      TabIndex        =   24
      Text            =   "0"
      Top             =   2520
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
      Left            =   6960
      TabIndex        =   23
      Text            =   "0"
      Top             =   2520
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
      Left            =   6360
      TabIndex        =   22
      Text            =   "0"
      Top             =   2520
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
      Left            =   8760
      TabIndex        =   21
      Text            =   "0"
      Top             =   2040
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
      Left            =   8160
      TabIndex        =   20
      Text            =   "0"
      Top             =   2040
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
      Left            =   7560
      TabIndex        =   19
      Text            =   "0"
      Top             =   2040
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
      Left            =   6960
      TabIndex        =   18
      Text            =   "0"
      Top             =   2040
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
      Left            =   6360
      TabIndex        =   17
      Text            =   "0"
      Top             =   2040
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
      Left            =   8760
      TabIndex        =   16
      Text            =   "0"
      Top             =   1560
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
      Left            =   8160
      TabIndex        =   15
      Text            =   "0"
      Top             =   1560
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
      Left            =   7560
      TabIndex        =   14
      Text            =   "1"
      Top             =   1560
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
      Left            =   6960
      TabIndex        =   13
      Text            =   "0"
      Top             =   1560
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
      Left            =   6360
      TabIndex        =   12
      Text            =   "0"
      Top             =   1560
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
      Left            =   8760
      TabIndex        =   11
      Text            =   "0"
      Top             =   1080
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
      Left            =   8160
      TabIndex        =   10
      Text            =   "0"
      Top             =   1080
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
      Left            =   7560
      TabIndex        =   9
      Text            =   "0"
      Top             =   1080
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
      Left            =   6960
      TabIndex        =   8
      Text            =   "0"
      Top             =   1080
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
      Left            =   6360
      TabIndex        =   7
      Text            =   "0"
      Top             =   1080
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
      Left            =   8760
      TabIndex        =   6
      Text            =   "0"
      Top             =   600
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
      Left            =   8160
      TabIndex        =   5
      Text            =   "0"
      Top             =   600
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
      Left            =   7560
      TabIndex        =   4
      Text            =   "0"
      Top             =   600
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
      Left            =   6960
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
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
      Left            =   6360
      TabIndex        =   2
      Text            =   "0"
      Top             =   600
      Width           =   540
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   32
      Top             =   5760
      Width           =   11415
   End
   Begin VB.Label lblOffset 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "offset:"
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
      Left            =   8160
      TabIndex        =   31
      Top             =   3480
      Width           =   675
   End
   Begin VB.Label lblDivisor 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "divisor:"
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
      Left            =   6240
      TabIndex        =   30
      Top             =   3495
      Width           =   795
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
      Left            =   6240
      TabIndex        =   29
      Top             =   120
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
'Copyright ©2001-2013 by Tanner Helland
'Created: 15/April/01
'Last updated: 08/August/13
'Last update: new "normalize" option to automatically populate divisor and offset values
'
'This dialog allows the user to create custom convolution filters.  It also allows the user to save those filters to
' file, or to load previously saved convolution filter files.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Used to render images onto the command buttons
Private cImgCtl As clsControlImage

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Normalizing automatically computes divisor and offset for the user
Private Sub chkNormalize_Click()
    updatePreview
End Sub

'When the user clicks OK...
Private Sub CmdOK_Click()
    
    'Before we do anything else, check to make sure every text box has a
    'valid number in it (no range checking is necessary)
    Dim x As Long, y As Long
    
    For x = 0 To 24
        If Not NumberValid(TxtF(x)) Then
            AutoSelectText TxtF(x)
            Exit Sub
        End If
    Next x
    If Not NumberValid(txtDivisor) Then
        AutoSelectText txtDivisor
        Exit Sub
    End If
    If Not NumberValid(txtOffset) Then
        AutoSelectText txtOffset
        Exit Sub
    End If
    
    Me.Visible = False
    
    'Copy the values from the text boxes into an array
    Message "Generating filter data..."
        
    g_FilterSize = 5
        
    ReDim g_FM(-2 To 2, -2 To 2) As Double
    
    For x = -2 To 2
    For y = -2 To 2
        g_FM(x, y) = Val(TxtF((x + 2) + (y + 2) * 5))
    Next y
    Next x
        
    'What to divide the final value by
    g_FilterWeight = CDbl(txtDivisor.Text)
    If g_FilterWeight = 0 Then g_FilterWeight = 1
    
    'Any offset value
    g_FilterBias = CDbl(txtOffset.Text)
    
    'Set that we have created a filter during this program session, and save it accordingly
    g_HasCreatedFilter = True
    
    saveCustomFilter g_UserPreferences.getTempPath & "~PD_CF.tmp"
    Process "Custom filter", , g_UserPreferences.getTempPath & "~PD_CF.tmp"
    
    Unload Me
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    ' (NOTE: because the icons require manifest theming, they will not appear in the IDE.)
    Set cImgCtl = New clsControlImage
    With cImgCtl
        .LoadImageFromStream cmdOpen.hWnd, LoadResData("LRGOPENDOC", "CUSTOM"), 32, 32
        .LoadImageFromStream cmdSave.hWnd, LoadResData("LRGSAVE", "CUSTOM"), 32, 32
        
        .SetMargins cmdOpen.hWnd, , 12
        .Align(cmdOpen.hWnd) = Icon_Top
        .SetMargins cmdSave.hWnd, , 12
        .Align(cmdSave.hWnd) = Icon_Top
    End With
    
    'If a filter has been used previously, load it from the temp file
    If g_HasCreatedFilter Then openCustomFilter g_UserPreferences.getTempPath & "~PD_CF.tmp"
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Render a preview
    updatePreview
    
End Sub

Private Sub cmdOpen_Click()
    'Simple open dialog
    Dim CC As cCommonDialog
        
    Dim sFile As String
    Set CC = New cCommonDialog
    
    Dim cdFilter As String
    cdFilter = PROGRAMNAME & " " & g_Language.TranslateMessage("Filter") & " (." & FILTER_EXT & ")|*." & FILTER_EXT & ";*.thf|"
    cdFilter = cdFilter & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Open a custom filter")
    
    If CC.VBGetOpenFileName(sFile, , , , , True, cdFilter, , g_UserPreferences.getFilterPath, cdTitle, , FormCustomFilter.hWnd, 0) Then
        
        'Disable normalizing during load, as it may erase saved weight/offset values
        chkNormalize = vbUnchecked
        
        If openCustomFilter(sFile) Then
            
            'Save the new directory as the default path for future usage
            g_UserPreferences.setFilterPath sFile
            
            'Redraw the preview
            updatePreview
            
        Else
            pdMsgBox "An error occurred while attempting to load %1.  Please verify that the file is a valid custom filter file.", vbOKOnly + vbExclamation + vbApplicationModal, "Custom Filter Error", sFile
        End If
        
    End If
    
End Sub

'Provide a save prompt, and use that to trigger a save of this custom filter to file
Private Sub cmdSave_Click()
    
    'Simple save dialog
    Dim CC As cCommonDialog
        
    Dim sFile As String
    Set CC = New cCommonDialog
    
    Dim cdFilter As String
    cdFilter = PROGRAMNAME & " " & g_Language.TranslateMessage("Filter") & " (." & FILTER_EXT & ")|*." & FILTER_EXT
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Save a custom filter")
    
    If CC.VBGetSaveFileName(sFile, , True, cdFilter, , g_UserPreferences.getFilterPath, cdTitle, "." & FILTER_EXT, FormCustomFilter.hWnd, 0) Then
        
        'Save the new directory as the default path for future usage
        g_UserPreferences.setFilterPath sFile
        
        'Write out the file
        saveCustomFilter sFile
        
    End If
    
End Sub

'Load a custom filter from file using the new pdXML class
Private Function openCustomFilter(ByRef srcFilterPath As String) As Boolean
    
    'Create a pdXML class, which will help us load and parse the source file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Load the XML file into memory
    xmlEngine.loadXMLFile srcFilterPath
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon filter file
    If xmlEngine.isPDDataType("Convolution filter") And xmlEngine.validateLoadedXMLData("pdFilterVersion") Then
    
        'Next, check the filter's version number, and make sure it's still supported
        Dim verCheck As String
        verCheck = xmlEngine.getUniqueTag_String("pdFilterVersion")
        
        Select Case verCheck
        
            'The current filter version (e.g. the first draft of the new XML format)
            Case CUSTOM_FILTER_VERSION_2013
            
                'Load the divisor and offset values
                txtDivisor = xmlEngine.getUniqueTag_Double("filterDivisor")
                txtOffset = xmlEngine.getUniqueTag_Double("filterOffset")
                
                'Load the individual text box values
                Dim i As Long
                For i = 0 To TxtF.UBound
                    TxtF(i) = xmlEngine.getUniqueTag_Double("filterEntry_" & i)
                Next i
            
            Case Else
                Message "Incompatible filter version found.  Filter load abandoned."
                openCustomFilter = False
                Exit Function
        
        End Select
        
        'Mark the load as successful and exit
        openCustomFilter = True
        Exit Function
        
    Else
        
        pdMsgBox "Unfortunately, this custom filter file is no longer supported by the current version of PhotoDemon." & vbCrLf & vbCrLf & "In version 6.0, PhotoDemon filter files were redesigned to support new features, improve performance, and solve some long-standing reliability issues.  Unfortunately, this means that custom filters created prior to version 6.0 are no longer compatible.  You will need to re-enter those filters from scratch." & vbCrLf & vbCrLf & "(Note that any old custom filter files will still work in old versions of PhotoDemon, if you absolutely need to access them.)", vbInformation + vbOKOnly, "Unsupported custom filter"
        openCustomFilter = False
        Exit Function
    End If
    
End Function

'Save a custom filter to file using the new pdXML class
Private Function saveCustomFilter(ByRef dstFilterPath As String) As Boolean

    'Create a pdXML class, which will help us assemble the file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    xmlEngine.prepareNewXML "Convolution filter"
    
    'Write out the XML version we're using for this filter
    xmlEngine.writeTag "pdFilterVersion", CUSTOM_FILTER_VERSION_2013
    xmlEngine.writeBlankLine
    
    'Write out the user-supplied divisor and offset values
    xmlEngine.writeTag "filterDivisor", txtDivisor
    xmlEngine.writeTag "filterOffset", txtOffset
    xmlEngine.writeBlankLine
    
    'Write out the contents of the text boxes
    Dim i As Long
    For i = 0 To TxtF.UBound
        xmlEngine.writeTag "filterEntry_" & i, TxtF(i)
    Next i
    
    'Finally, ask the XML file to write itself out to file.
    xmlEngine.writeBlankLine
    xmlEngine.writeXMLToFile dstFilterPath
    
    saveCustomFilter = True
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub txtOffset_GotFocus()
    AutoSelectText txtOffset
End Sub

Private Sub txtOffset_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtOffset, True, True
    updatePreview
End Sub

Private Sub TxtF_GotFocus(Index As Integer)
    AutoSelectText TxtF(Index)
End Sub

Private Sub TxtF_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    textValidate TxtF(Index), True, True
    updatePreview
End Sub

Private Sub txtDivisor_GotFocus()
    AutoSelectText txtDivisor
End Sub

Private Sub txtDivisor_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtDivisor, True, True
    updatePreview
End Sub

'When the filter is changed, update the preview to match
Private Sub updatePreview()

    Dim x As Long, y As Long

    'We can only apply the preview if all loaded text boxes are valid
    For x = 0 To 24
        If Not EntryValid(TxtF(x), -1000000, 1000000, False, False) Then Exit Sub
    Next x
    
    If Not EntryValid(txtDivisor, -1000000, 1000000, False, False) Then Exit Sub
    
    If Not EntryValid(txtOffset, -1000000, 1000000, False, False) Then Exit Sub
    
    'If normalization has been requested, apply it now
    txtDivisor.Enabled = Not CBool(chkNormalize)
    txtOffset.Enabled = Not CBool(chkNormalize)
    
    If CBool(chkNormalize) Then
    
        'Sum up the total of all filter boxes
        Dim filterSum As Double
        filterSum = 0
        For x = 0 To 24
            filterSum = filterSum + CDbl(TxtF(x))
        Next x
        
        'Generate automatic divisor and offset values based on the total.
        If filterSum = 0 Then
            txtDivisor = 1
            txtOffset = 127
        ElseIf filterSum > 0 Then
            txtDivisor = filterSum
            txtOffset = 0
        Else
            txtDivisor = Abs(filterSum)
            txtOffset = 255
        End If
    
    End If
    
    'Copy the values from the text boxes into an array
    g_FilterSize = 5
        
    ReDim g_FM(-2 To 2, -2 To 2) As Double
    
    For x = -2 To 2
    For y = -2 To 2
        g_FM(x, y) = Val(TxtF((x + 2) + (y + 2) * 5))
    Next y
    Next x
        
    'What to divide the final value by
    g_FilterWeight = CDbl(txtDivisor.Text)
    
    'Offset value
    g_FilterBias = CDbl(txtOffset.Text)
        
    'Apply the preview
    DoFilter g_Language.TranslateMessage("Preview"), False, , True, fxPreview
    
End Sub

