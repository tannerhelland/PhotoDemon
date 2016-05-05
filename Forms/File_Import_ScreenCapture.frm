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
   Begin PhotoDemon.pdListBox lstWindows 
      Height          =   2895
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Width           =   5055
      _extentx        =   8916
      _extenty        =   5106
   End
   Begin PhotoDemon.pdLabel lblMinimizedWarning 
      Height          =   495
      Left            =   6120
      Top             =   4980
      Visible         =   0   'False
      Width           =   6855
      _extentx        =   12091
      _extenty        =   873
      alignment       =   2
      caption         =   ""
      fontsize        =   9
      forecolor       =   2627816
      layout          =   1
      usecustomforecolor=   -1  'True
   End
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6255
      Width           =   13095
      _extentx        =   23098
      _extenty        =   1323
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   6120
      ScaleHeight     =   287
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   455
      TabIndex        =   5
      Top             =   600
      Width           =   6855
   End
   Begin PhotoDemon.pdCheckBox chkMinimize 
      Height          =   300
      Left            =   840
      TabIndex        =   3
      Top             =   1050
      Width           =   5115
      _extentx        =   9022
      _extenty        =   582
      caption         =   "minimize PhotoDemon prior to capture"
   End
   Begin PhotoDemon.pdRadioButton optSource 
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   5370
      _extentx        =   9472
      _extenty        =   582
      caption         =   "entire desktop"
      value           =   -1  'True
   End
   Begin PhotoDemon.pdRadioButton optSource 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   5370
      _extentx        =   9472
      _extenty        =   582
      caption         =   "specific program (listed by window title)"
   End
   Begin PhotoDemon.pdCheckBox chkChrome 
      Height          =   330
      Left            =   840
      TabIndex        =   4
      Top             =   5040
      Width           =   5115
      _extentx        =   8599
      _extenty        =   582
      caption         =   "include window decorations"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   6120
      Top             =   180
      Width           =   6825
      _extentx        =   0
      _extenty        =   0
      caption         =   "preview"
      fontsize        =   12
      forecolor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   240
      Top             =   180
      Width           =   5730
      _extentx        =   0
      _extenty        =   0
      caption         =   "screenshot source"
      fontsize        =   12
      forecolor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblSecurity 
      Height          =   600
      Left            =   840
      Top             =   5550
      Width           =   12015
      _extentx        =   21193
      _extenty        =   1058
      caption         =   ""
      forecolor       =   4210752
      layout          =   1
   End
End
Attribute VB_Name = "FormScreenCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Custom Screen Capture Dialog
'Copyright 2012-2016 by Tanner Helland
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

'List of open application names and their top-level hWnds
Private m_WindowNames As pdStringStack
Private m_WindowHWnds As pdStringStack

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
        If (lstWindows.ListIndex = -1) Then
            PDMsgBox "Please select a window to capture.", vbInformation + vbApplicationModal + vbOKOnly, "Target window required"
            Exit Sub
        End If
        
        Me.Visible = False
        Process "Screen capture", False, buildParams(False, CBool(chkMinimize.Value), IIf(lstWindows.ListIndex > -1, m_WindowHWnds.GetString(lstWindows.ListIndex), 0&), CBool(chkChrome), IIf(lstWindows.ListIndex > -1, lstWindows.List(lstWindows.ListIndex), "Screen capture")), UNDO_NOTHING
        
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
    
    lblSecurity.Caption = g_Language.TranslateMessage("Some programs (including Windows Store apps) do not allow direct screen captures.  If an application preview appears as a black square, you will need to take a full desktop screenshot, then manually crop the desired window region.")
    
    'Retrieve a list of all currently open programs.  Many thanks to Karl E Peterson for help with this topic, via:
    ' http://vb.mvps.org/articles/ap199902.pdf
    FillListWithOpenApplications lstWindows
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Wait just a moment before continuing, to give the corresponding menu time to animate away (otherwise it may
    ' get caught in the capture preview)
    Sleep 500
    
    'Render a preview of whichever item is currently selected
    UpdatePreview
    
End Sub

'Given a list box, fill it with a list of open applications.  The .ItemData property will be filled with
' each window's hWnd.
Private Function FillListWithOpenApplications(ByVal dstListbox As pdListBox) As Long
    
    dstListbox.Clear
    Call EnumWindows(AddressOf Screen_Capture.EnumWindowsProc, 0&)
    
    'Retrieve the list of window names and hWnds
    Screen_Capture.GetAllWindowNamesAndHWnds m_WindowNames, m_WindowHWnds
    
    'Fill the list box with the retrieved list of window names
    Dim i As Long
    For i = 0 To m_WindowNames.GetNumOfStrings - 1
        dstListbox.AddItem m_WindowNames.GetString(i), i
    Next i
    
    FillListWithOpenApplications = dstListbox.ListCount
    
End Function

Private Sub lstWindows_Click()
    If (Not optSource(1)) Then optSource(1) = True
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
        If (lstWindows.ListIndex > -1) Then
            
            'Make sure the function returns successfully; if a window is unloaded after the listbox has been
            ' filled, the function will (obviously) fail to capture the screen contents.
            Dim minimizeCheck As Boolean
            If Screen_Capture.GetHwndContentsAsDIB(tmpDIB, CLng(m_WindowHWnds.GetString(lstWindows.ListIndex)), chkChrome, minimizeCheck) Then
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
    tmpDIB.CreateBlank picPreview.ScaleWidth, picPreview.ScaleHeight
    
    Dim notifyFont As pdFont
    Set notifyFont = New pdFont
    notifyFont.SetFontFace g_InterfaceFont
    notifyFont.SetFontSize 14
    notifyFont.SetFontColor 0
    notifyFont.SetTextAlignment vbCenter
    notifyFont.CreateFontObject
    notifyFont.AttachToDC tmpDIB.GetDIBDC
    
    notifyFont.FastRenderText picPreview.ScaleWidth / 2, picPreview.ScaleHeight / 2 - notifyFont.GetHeightOfString("ABCjqy"), g_Language.TranslateMessage("Unfortunately, that program has exited.")
    notifyFont.FastRenderText picPreview.ScaleWidth / 2, picPreview.ScaleHeight / 2, g_Language.TranslateMessage("Please select another one.")
    tmpDIB.RenderToPictureBox picPreview
    notifyFont.ReleaseFromDC
    Set tmpDIB = Nothing
    
    lblMinimizedWarning.Visible = False
    
End Sub
