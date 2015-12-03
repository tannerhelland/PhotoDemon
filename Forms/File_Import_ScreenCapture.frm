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
   Begin PhotoDemon.pdLabel lblMinimizedWarning 
      Height          =   495
      Left            =   6120
      Top             =   5640
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   873
      Alignment       =   2
      Caption         =   ""
      FontSize        =   9
      ForeColor       =   2627816
      Layout          =   1
      UseCustomForeColor=   -1  'True
   End
   Begin PhotoDemon.commandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   8
      Top             =   6255
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   6120
      ScaleHeight     =   327
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
      _ExtentY        =   582
      Caption         =   "entire desktop"
      Value           =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton optSource 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   582
      Caption         =   "specific program (listed by window title)"
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
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "preview"
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
      Left            =   6120
      TabIndex        =   7
      Top             =   180
      Width           =   825
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "screenshot source"
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
      Width           =   1890
   End
End
Attribute VB_Name = "FormScreenCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Custom Screen Capture Dialog
'Copyright 2012-2015 by Tanner Helland
'Created: 01/January/12 (approx)
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

'APIs for listing currently open applications (windows)
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Sub chkChrome_Click()
    UpdatePreview
End Sub

Private Sub chkMinimize_Click()
    UpdatePreview
End Sub

Private Sub cmdBarMini_OKClick()
    
    If optSource(0) Then
        Me.Visible = False
        Process "Screen capture", False, buildParams(True, CBool(chkMinimize), 0, CBool(chkChrome), "Screen capture"), UNDO_NOTHING
    Else
        
        'Make sure the user has selected a window to capture
        If lstWindows.ListIndex = -1 Then
            PDMsgBox "Please select a window to capture.", vbInformation + vbApplicationModal + vbOKOnly, "Target window required"
            Exit Sub
        End If
        
        Me.Visible = False
        Process "Screen capture", False, buildParams(False, CBool(chkMinimize), IIf(lstWindows.ListIndex > -1, lstWindows.itemData(lstWindows.ListIndex), 0), CBool(chkChrome), IIf(lstWindows.ListIndex > -1, lstWindows.List(lstWindows.ListIndex), "Screen capture")), UNDO_NOTHING
        
    End If
    
End Sub

Private Sub Form_Load()
        
    'Populate the "window is minimized" warning
    lblMinimizedWarning.Caption = g_Language.TranslateMessage("This program is currently minimized.  Restore it to normal size for best results.")
    If Not (g_Themer Is Nothing) Then
        lblMinimizedWarning.ForeColor = g_Themer.GetThemeColor(PDTC_CANCEL_RED)
    Else
        lblMinimizedWarning.ForeColor = RGB(232, 24, 20)
    End If
        
    'Retrieve a list of all currently open programs.  Many thanks to Karl E Peterson for help with this topic, via:
    ' http://vb.mvps.org/articles/ap199902.pdf
    FillListWithOpenApplications lstWindows
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Wait just a moment before continuing, to give the corresponding menu time to animate away (otherwise it may
    ' get caught in the capture preview)
    Sleep 500
    
    'Render a preview of whichever item is currently selected
    UpdatePreview
    
End Sub

'Given a list box, fill it with a list of open applications.  The .ItemData property will be filled with
' each window's hWnd.
Private Function FillListWithOpenApplications(ByVal dstListbox As ListBox) As Long
    
    dstListbox.Clear
    Call EnumWindows(AddressOf Screen_Capture.EnumWindowsProc, dstListbox.hWnd)
    FillListWithOpenApplications = dstListbox.ListCount
    
End Function

Private Sub lstWindows_Click()
    
    If Not optSource(1) Then optSource(1) = True
    UpdatePreview
    
End Sub

Private Sub optSource_Click(Index As Integer)
    
    'If the user has selected "specific program", make sure a program is selected
    If Index = 1 Then
        If lstWindows.ListIndex = -1 Then lstWindows.ListIndex = 0
    End If
    
    UpdatePreview
    
End Sub

'Live previews of the screen capture are now provided
Private Sub UpdatePreview()

    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB

    'Full screen capture was requested
    If optSource(0) Then
        Screen_Capture.GetDesktopAsDIB tmpDIB
        tmpDIB.RenderToPictureBox picPreview
    
    'Specific window capture was requested
    Else
        If lstWindows.ListIndex > -1 Then
            
            'Make sure the function returns successfully; if a window is unloaded after the listbox has been
            ' filled, the function will (obviously) fail to capture the screen contents.
            Dim minimizeCheck As Boolean
            If Screen_Capture.GetHwndContentsAsDIB(tmpDIB, lstWindows.itemData(lstWindows.ListIndex), chkChrome, minimizeCheck) Then
                tmpDIB.RenderToPictureBox picPreview, , True
                lblMinimizedWarning.Visible = minimizeCheck
            Else
                lstWindows.RemoveItem lstWindows.ListIndex
                DisplayScreenCaptureError
            End If
            
        End If
    
    End If
    
End Sub

'If the user attempts to capture a window after it's been unloaded, warn them via this function
Private Sub DisplayScreenCaptureError()

    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createBlank picPreview.ScaleWidth, picPreview.ScaleHeight
    
    Dim notifyFont As pdFont
    Set notifyFont = New pdFont
    notifyFont.SetFontFace g_InterfaceFont
    notifyFont.SetFontSize 14
    notifyFont.SetFontColor 0
    notifyFont.SetTextAlignment vbCenter
    notifyFont.CreateFontObject
    notifyFont.AttachToDC tmpDIB.getDIBDC
    
    notifyFont.FastRenderText picPreview.ScaleWidth / 2, picPreview.ScaleHeight / 2 - notifyFont.GetHeightOfString("ABCjqy"), g_Language.TranslateMessage("Unfortunately, that program has exited.")
    notifyFont.FastRenderText picPreview.ScaleWidth / 2, picPreview.ScaleHeight / 2, g_Language.TranslateMessage("Please select another one.")
    tmpDIB.RenderToPictureBox picPreview
    notifyFont.ReleaseFromDC
    Set tmpDIB = Nothing
    
    lblMinimizedWarning.Visible = False
    
End Sub
