VERSION 5.00
Begin VB.Form FormScreenCapture 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Screenshot options"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13095
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
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   873
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10140
      TabIndex        =   8
      Top             =   6390
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11610
      TabIndex        =   7
      Top             =   6390
      Width           =   1365
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   6120
      ScaleHeight     =   359
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   455
      TabIndex        =   6
      Top             =   600
      Width           =   6855
   End
   Begin PhotoDemon.smartCheckBox chkMinimize 
      Height          =   300
      Left            =   840
      TabIndex        =   4
      Top             =   1050
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   582
      Caption         =   "minimize PhotoDemon prior to capture"
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
   Begin VB.ListBox lstWindows 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   4815
   End
   Begin PhotoDemon.smartOptionButton optSource 
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   661
      Caption         =   "entire desktop"
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optSource 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   661
      Caption         =   "specific program (listed by window title)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartCheckBox chkChrome 
      Height          =   330
      Left            =   840
      TabIndex        =   5
      Top             =   5640
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   582
      Caption         =   "include window decorations"
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "screenshot preview:"
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
      Index           =   1
      Left            =   5880
      TabIndex        =   10
      Top             =   180
      Width           =   2115
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -120
      TabIndex        =   9
      Top             =   6240
      Width           =   13815
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "screenshot source:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1980
   End
End
Attribute VB_Name = "FormScreenCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Custom Screen Capture Dialog
'Copyright ©2012-2014 by Tanner Helland
'Created: 01/January/14 (approx)
'Last updated: 15/January/14
'Last update: minor bugfixes to account for delays caused by window animations
'
'Basic screen and window capture dialog.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'APIs for listing currently open applications (windows)
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Sub chkChrome_Click()
    updatePreview
End Sub

Private Sub chkMinimize_Click()
    updatePreview
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    
    If optSource(0) Then
        Me.Visible = False
        Process "Screen capture", False, buildParams(True, CBool(chkMinimize), 0, CBool(chkChrome), "Screen capture"), UNDO_NOTHING
    Else
        
        'Make sure the user has selected a window to capture
        If lstWindows.ListIndex = -1 Then
            pdMsgBox "Please select a window to capture.", vbInformation + vbApplicationModal + vbOKOnly, "Target window required"
            Exit Sub
        End If
        
        Me.Visible = False
        Process "Screen capture", False, buildParams(False, CBool(chkMinimize), IIf(lstWindows.ListIndex > -1, lstWindows.ItemData(lstWindows.ListIndex), 0), CBool(chkChrome), IIf(lstWindows.ListIndex > -1, lstWindows.List(lstWindows.ListIndex), "Screen capture")), UNDO_NOTHING
        
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    'Retrieve a list of all currently open programs.  Many thanks to Karl E Peterson for help with this topic, via:
    ' http://vb.mvps.org/articles/ap199902.pdf
    fillListWithOpenApplications lstWindows
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Wait just a moment before continuing, to give the corresponding menu time to animate away (otherwise it may
    ' get caught in the capture preview)
    Sleep 500
    
    'Render a preview of whichever item is currently selected
    updatePreview
    
End Sub

'Given a list box, fill it with a list of open applications.  The .ItemData property will be filled with
' each window's hWnd.
Private Function fillListWithOpenApplications(ByVal dstListbox As ListBox) As Long
    
    dstListbox.Clear
    Call EnumWindows(AddressOf Screen_Capture.EnumWindowsProc, dstListbox.hWnd)
    fillListWithOpenApplications = dstListbox.ListCount
    
End Function

Private Sub lstWindows_Click()
    
    If Not optSource(1) Then optSource(1) = True
    updatePreview
    
End Sub

Private Sub optSource_Click(Index As Integer)
    
    'If the user has selected "specific program", make sure a program is selected
    If Index = 1 Then
        If lstWindows.ListIndex = -1 Then lstWindows.ListIndex = 0
    End If
    updatePreview
    
End Sub

'Live previews of the screen capture are now provided
Private Sub updatePreview()

    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB

    'Full screen capture was requested
    If optSource(0) Then
        Screen_Capture.getDesktopAsDIB tmpDIB
        tmpDIB.renderToPictureBox picPreview
    
    'Specific window capture was requested
    Else
        If lstWindows.ListIndex > -1 Then
            
            'Make sure the function returns successfully; if a window is unloaded after the listbox has been
            ' filled, the function will (obviously) fail to capture the screen contents.
            If Screen_Capture.getHwndContentsAsDIB(tmpDIB, lstWindows.ItemData(lstWindows.ListIndex), chkChrome) Then
                tmpDIB.renderToPictureBox picPreview
            Else
                lstWindows.RemoveItem lstWindows.ListIndex
                displayScreenCaptureError
            End If
            
        End If
    
    End If
    
End Sub

'If the user attempts to capture a window after it's been unloaded, warn them via this function
Private Sub displayScreenCaptureError()

    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createBlank picPreview.ScaleWidth, picPreview.ScaleHeight
    
    Dim notifyFont As pdFont
    Set notifyFont = New pdFont
    notifyFont.setFontFace g_InterfaceFont
    notifyFont.setFontSize 14
    notifyFont.setFontColor 0
    notifyFont.setTextAlignment vbCenter
    notifyFont.createFontObject
    notifyFont.attachToDC tmpDIB.getDIBDC
    
    notifyFont.fastRenderText picPreview.ScaleWidth / 2, picPreview.ScaleHeight / 2 - notifyFont.getHeightOfString("ABCjqy"), g_Language.TranslateMessage("Unfortunately, that program has exited.")
    notifyFont.fastRenderText picPreview.ScaleWidth / 2, picPreview.ScaleHeight / 2, g_Language.TranslateMessage("Please select another one.")
    tmpDIB.renderToPictureBox picPreview
    Set tmpDIB = Nothing

End Sub
