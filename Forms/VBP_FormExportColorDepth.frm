VERSION 5.00
Begin VB.Form dialog_ExportColorDepth 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Please Choose Exported Color Depth"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6510
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
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optColorDepth 
      Appearance      =   0  'Flat
      Caption         =   " 32 bpp (16 million colors + full transparency)"
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
      Height          =   345
      Index           =   4
      Left            =   600
      TabIndex        =   7
      Top             =   2640
      Value           =   -1  'True
      Width           =   5775
   End
   Begin VB.OptionButton optColorDepth 
      Appearance      =   0  'Flat
      Caption         =   " 24 bpp (16 million colors)"
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
      Height          =   345
      Index           =   3
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   3615
   End
   Begin VB.OptionButton optColorDepth 
      Appearance      =   0  'Flat
      Caption         =   " 8 bpp (256 colors or shades of gray)"
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
      Height          =   345
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   5175
   End
   Begin VB.OptionButton optColorDepth 
      Appearance      =   0  'Flat
      Caption         =   " 4 bpp (16 shades of gray)"
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
      Height          =   345
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   3975
   End
   Begin VB.OptionButton optColorDepth 
      Appearance      =   0  'Flat
      Caption         =   " 1 bpp (monochrome)"
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
      Height          =   345
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   3600
      Width           =   1245
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "available color depths for this format:"
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
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3990
   End
End
Attribute VB_Name = "dialog_ExportColorDepth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Exported Color Depth Dialog
'Copyright ©2011-2012 by Tanner Helland
'Created: 11/December/12
'Last updated: 11/December/12
'Last update: initial build
'
'Dialog for presenting the user a choice of exported color depths.  I prefer this to be
' handled automatically by the software, but in certain rare cases it may be desirable
' for a user to manually export a certain color depth
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'The desired output format (used to activate available color depths)
Private outputFormat As Long

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'This form can be notified of the image being exported.  This may be used in the future to provide a preview.
Public Property Let imageFormat(ByVal imageFormat As Long)
    outputFormat = imageFormat
End Property

'CANCEL button
Private Sub CmdCancel_Click()
    
    userAnswer = vbCancel
    Me.Hide
    
End Sub

'OK button
Private Sub CmdOK_Click()
        
    'Save the selected color depth to the corresponding global variable (so other functions can access it
    ' after this form is unloaded)
    If optColorDepth(0).Value = True Then g_ColorDepth = 1
    If optColorDepth(1).Value = True Then g_ColorDepth = 4
    If optColorDepth(2).Value = True Then g_ColorDepth = 8
    If optColorDepth(3).Value = True Then g_ColorDepth = 24
    If optColorDepth(4).Value = True Then g_ColorDepth = 32
     
    userAnswer = vbOK
    Me.Hide
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog()

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
        
    'Based on the supplied image format, disable invalid color depths
    If imageFormats.isColorDepthSupported(outputFormat, 1) Then optColorDepth(0).Enabled = True Else optColorDepth(0).Enabled = False
    If imageFormats.isColorDepthSupported(outputFormat, 4) Then optColorDepth(1).Enabled = True Else optColorDepth(1).Enabled = False
    If imageFormats.isColorDepthSupported(outputFormat, 8) Then optColorDepth(2).Enabled = True Else optColorDepth(2).Enabled = False
    If imageFormats.isColorDepthSupported(outputFormat, 24) Then optColorDepth(3).Enabled = True Else optColorDepth(3).Enabled = False
    If imageFormats.isColorDepthSupported(outputFormat, 32) Then optColorDepth(4).Enabled = True Else optColorDepth(4).Enabled = False
        
    'Out of politeness, set the default color depth to the current image's color depth
    If (pdImages(CurrentImage).mainLayer.getLayerColorDepth = 24) And (optColorDepth(3).Enabled) Then
        optColorDepth(3).Value = True
    ElseIf (pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32) And (optColorDepth(4).Enabled) Then
        optColorDepth(4).Value = True
    Else
        'If both 24 and 32bpp are disabled (not possible at present, but whatever), select the highest possible
        ' color depth by default
        Dim i As Long
        For i = optColorDepth.Count - 1 To 0 Step -1
            If optColorDepth(i).Enabled Then
                optColorDepth(i).Value = True
                Exit For
            End If
        Next i
    End If
        
    Message "Waiting for user to specify color depth... "
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Display the dialog
    Me.Show vbModal, FormMain

End Sub

