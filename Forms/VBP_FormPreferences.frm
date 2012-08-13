VERSION 5.00
Begin VB.Form FormPreferences 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Preferences"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
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
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkConfirmUnsaved 
      Appearance      =   0  'Flat
      Caption         =   "Confirm closing of unsaved images"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Width           =   3375
   End
   Begin VB.PictureBox picCanvasImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5040
      Picture         =   "VBP_FormPreferences.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picCanvasColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   12
      ToolTipText     =   "Click to change color"
      Top             =   840
      Width           =   4815
   End
   Begin VB.ComboBox cmbCanvas 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   435
      Width           =   2655
   End
   Begin VB.ComboBox cmbLargeImages 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1800
      Width           =   4575
   End
   Begin VB.CheckBox ChkLogMessages 
      Appearance      =   0  'Flat
      Caption         =   "Log program messages to file (useful for debugging)"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   $"VBP_FormPreferences.frx":008A
      Top             =   3480
      Width           =   4455
   End
   Begin VB.CommandButton CmdTmpPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   5
      ToolTipText     =   "Click to open a browser-folder dialog"
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox TxtTempPath 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "automatically generated at run-time"
      ToolTipText     =   "Folder used for temporary calculations"
      Top             =   4320
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   5040
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   5040
      Width           =   1125
   End
   Begin VB.CheckBox ChkPromptPluginDownload 
      Appearance      =   0  'Flat
      Caption         =   "Check for missing plugins on program start"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label lblGeneralOptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General Program Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   2130
   End
   Begin VB.Label lblCanvasFX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image window background:"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   1980
   End
   Begin VB.Label lblInterface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interface Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label lblImgOpen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "When large images are opened: "
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   2340
   End
   Begin VB.Label lblTempFolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Temporary file folder (used to hold Undo/Redo data):"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   4575
   End
End
Attribute VB_Name = "FormPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Program Preferences Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 8/November/02
'Last updated: 03/July/12
'Last update: added support for custom canvas backgrounds (checkerboard or colors)
'
'Module for interfacing with the user's desired program preferences.  Handles
'reading from and copying to the program's ".INI" file.
'
'***************************************************************************

Option Explicit

'Used to see if the user physically clicked the canvas combo box, or if VB selected it on its own
Dim userInitiatedColorSelection As Boolean

'Canvas background selection
Private Sub cmbCanvas_Click()
    
    'Only respond to user-generated events
    If userInitiatedColorSelection = False Then Exit Sub
    
    'Redraw the sample picture box value based on the value the user has selected
    Select Case cmbCanvas.ListIndex
        Case 0
            CanvasBackground = -1
        Case 1
            CanvasBackground = vb3DLight
        Case 2
            CanvasBackground = vb3DShadow
        Case 3
            'Prompt with a color selection box
            Dim retColor As Long
    
            Dim CD1 As cCommonDialog
            Set CD1 = New cCommonDialog
    
            retColor = picCanvasColor.BackColor
    
            CD1.VBChooseColor retColor, True, True, False, Me.HWnd
    
            'If a color was selected, change the picture box and associated combo box to match
            If retColor > 0 Then CanvasBackground = retColor Else CanvasBackground = picCanvasColor.BackColor
            
    End Select
    
    DrawSampleCanvasBackground
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Store whether the user wants to be prompted when closing unsaved images
    If chkConfirmUnsaved.Value = vbChecked Then
        ConfirmClosingUnsaved = True
        WriteToIni "General Preferences", "ConfirmClosingUnsaved", 1
    Else
        ConfirmClosingUnsaved = False
        WriteToIni "General Preferences", "ConfirmClosingUnsaved", 0
    End If
    
    'Store whether PhotoDemon is allowed to offer the automatic download of missing core plugins
    If ChkPromptPluginDownload.Value = vbChecked Then
        WriteToIni "General Preferences", "PromptForPluginDownload", 1
    Else
        WriteToIni "General Preferences", "PromptForPluginDownload", 0
    End If
    
    'Store whether we'll log system messages or not
    If ChkLogMessages.Value = vbChecked Then
        LogProgramMessages = True
        WriteToIni "General Preferences", "LogProgramMessages", 1
    Else
        LogProgramMessages = False
        WriteToIni "General Preferences", "LogProgramMessages", 0
    End If
    
    'Store the canvas background preference
    Select Case cmbCanvas.ListIndex
        
        'Checkerboard pattern
        Case 0
            CanvasBackground = -1
            WriteToIni "General Preferences", "CanvasBackground", "-1"
            
        'Color only
        Case Else
            CanvasBackground = picCanvasColor.BackColor
            WriteToIni "General Preferences", "CanvasBackground", CStr(CLng(picCanvasColor.BackColor))
            
    End Select
    
    'Now run a loop to draw the checkerboard effect on every window
    Dim tForm As Form
    Message "Saving preferences..."
    For Each tForm In VB.Forms
        If tForm.Name = "FormImage" Then ScrollViewport tForm
    Next
    Message "Finished."
    
    'Remember whether or not to autozoom large images
    AutosizeLargeImages = cmbLargeImages.ListIndex
    WriteToIni "General Preferences", "AutosizeLargeImages", CStr(AutosizeLargeImages)
    
    'Verify the temporary path
    If LCase(TxtTempPath.Text) <> LCase(TempPath) Then
        WriteToIni "Paths", "TempPath", TxtTempPath.Text
    End If
    
    Unload Me
End Sub

'When the "..." button is clicked, prompt the user with a "browse for folder" dialog
Private Sub CmdTmpPath_Click()
    Dim tString As String
    tString = BrowseForFolder(Me.HWnd)
    If tString <> "" Then TxtTempPath.Text = FixPath(tString)
End Sub

'When the form is loaded, populate the various checkboxes and textboxes with the values from the INI file
Private Sub Form_Load()
    
    Me.Caption = PROGRAMNAME & " Preferences"
    
    'Load in the appropriate values from the INI file
    
    'Start with the canvas background (which also requires populating the canvas background combo box)
    userInitiatedColorSelection = False
    cmbCanvas.AddItem "Checkerboard pattern", 0
    cmbCanvas.AddItem "System theme: light", 1
    cmbCanvas.AddItem "System theme: dark", 2
    cmbCanvas.AddItem "Custom color", 3
    
    'Select the proper combo box value based on the CanvasBackground variable
    If CanvasBackground = -1 Then
        'Checkerboard pattern
        cmbCanvas.ListIndex = 0
    ElseIf CanvasBackground = vb3DLight Then
        'System theme: light
        cmbCanvas.ListIndex = 1
    ElseIf CanvasBackground = vb3DShadow Then
        'System theme: dark
        cmbCanvas.ListIndex = 2
    Else
        'Custom color
        cmbCanvas.ListIndex = 3
    End If
    
    'Draw the current canvas background to the sample picture box
    DrawSampleCanvasBackground
    userInitiatedColorSelection = True
    
    'Assign the check box for logging program messages
    If LogProgramMessages = True Then ChkLogMessages.Value = vbChecked Else ChkLogMessages.Value = vbUnchecked
    
    'Assign the check box for prompting about unsaved images
    If ConfirmClosingUnsaved = True Then chkConfirmUnsaved.Value = vbChecked Else chkConfirmUnsaved.Value = vbUnchecked
    
    'Display the current temporary file path
    TxtTempPath.Text = TempPath
    
    'We have to pull the "offer to download plugins" value from the INI file, since we don't track
    ' it internally (it's only accessed when PhotoDemon is first loaded)
    Dim tmpString As String
    tmpString = GetFromIni("General Preferences", "PromptForPluginDownload")
    If val(tmpString) = 1 Then ChkPromptPluginDownload.Value = vbChecked Else ChkPromptPluginDownload.Value = vbUnchecked
    
    'Populate the "what to do when loading large images" combo box
    cmbLargeImages.AddItem "Automatically zoom out so the images fit on-screen", 0
    cmbLargeImages.AddItem "Load images at 100% zoom regardless of size", 1
    
    tmpString = GetFromIni("General Preferences", "AutosizeLargeImages")
    cmbLargeImages.ListIndex = val(tmpString)
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    setHandCursor picCanvasColor
    
End Sub

'Draw a sample of the current background to the PicCanvasColor picture box
Private Sub DrawSampleCanvasBackground()

    '-1 indicates the user wants a checkboard background pattern
    If CanvasBackground = -1 Then

        Dim stepIntervalX As Long, stepIntervalY As Long
        stepIntervalX = Me.picCanvasImg.ScaleWidth
        stepIntervalY = Me.picCanvasImg.ScaleHeight

        For x = 0 To picCanvasColor.ScaleWidth Step stepIntervalX
        For y = 0 To picCanvasColor.ScaleHeight Step stepIntervalY
            BitBlt Me.picCanvasColor.hDC, x, y, stepIntervalX, stepIntervalY, Me.picCanvasImg.hDC, 0, 0, vbSrcCopy
        Next y
        Next x
        picCanvasColor.Picture = picCanvasColor.Image
        
        Me.picCanvasColor.Enabled = False
        
    'Any other value is treated as an RGB long
    Else
    
        Me.picCanvasColor.Picture = LoadPicture("")
        Me.picCanvasColor.BackColor = CanvasBackground
    
        Me.picCanvasColor.Enabled = True
    
    End If
    
End Sub

'Clicking the sample color box allows the user to pick a new color
Private Sub picCanvasColor_Click()
    
    Dim retColor As Long
    
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    
    retColor = picCanvasColor.BackColor
    
    'Display a Windows color selection box
    CD1.VBChooseColor retColor, True, True, False, Me.HWnd
    
    'If a color was selected, change the picture box and associated combo box to match
    If retColor > 0 Then
    
        CanvasBackground = retColor
        
        userInitiatedColorSelection = False
        If CanvasBackground = vb3DLight Then
            'System theme: light
            cmbCanvas.ListIndex = 1
        ElseIf CanvasBackground = vb3DShadow Then
            'System theme: dark
            cmbCanvas.ListIndex = 2
        Else
            'Custom color
            cmbCanvas.ListIndex = 3
        End If
        userInitiatedColorSelection = True
        
        DrawSampleCanvasBackground
        
    End If
    
End Sub
