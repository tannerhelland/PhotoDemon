VERSION 5.00
Begin VB.Form FormScreenCapture 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Screenshot options"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13095
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   873
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4335
      Index           =   1
      Left            =   120
      Top             =   1200
      Width           =   5895
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButton cmdResetList 
         Height          =   615
         Left            =   5175
         TabIndex        =   3
         Top             =   3150
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
      End
      Begin PhotoDemon.pdListBox lstWindows 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5530
         Caption         =   "currently available programs (listed by window title):"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdCheckBox chkChrome 
         Height          =   330
         Left            =   240
         TabIndex        =   5
         Top             =   3240
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   582
         Caption         =   "include window decorations"
      End
      Begin PhotoDemon.pdLabel lblMinimizedWarning 
         Height          =   615
         Left            =   240
         Top             =   3660
         Visible         =   0   'False
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   1085
         Alignment       =   2
         Caption         =   ""
         FontSize        =   9
         ForeColor       =   2627816
         Layout          =   1
         UseCustomForeColor=   -1  'True
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4335
      Index           =   0
      Left            =   120
      Top             =   1200
      Width           =   5895
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdCheckBox chkMinimize 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   582
         Caption         =   "minimize PhotoDemon prior to capture"
      End
   End
   Begin PhotoDemon.pdButtonStrip btsSource 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "screenshot source"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6255
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdPictureBox picPreview 
      Height          =   4935
      Left            =   6120
      Top             =   600
      Width           =   6855
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   6120
      Top             =   180
      Width           =   6825
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "preview"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblSecurity 
      Height          =   600
      Left            =   360
      Top             =   5580
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1058
      Caption         =   ""
      ForeColor       =   4210752
      Layout          =   1
   End
End
Attribute VB_Name = "FormScreenCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Custom Screen Capture Dialog
'Copyright 2012-2026 by Tanner Helland
'Created: 01/January/12 (approx)
'Last updated: 13/April/22
'Last update: replace lingering picture box with pdPictureBox and fix some minor annoyances
'
'Basic screen and window capture dialog.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'List of open application names and their top-level hWnds
Private m_WindowNames As pdStringStack, m_WindowHWnds As pdStringStack

'APIs for listing currently open applications (windows)
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'Back buffer for the preview (if any)
Private m_PreviewDIB As pdDIB

'Prevent recursive redraws
Private m_SkipUpdate As Boolean

Private Sub btsSource_Click(ByVal buttonIndex As Long)
    
    UpdateVisibleContainer
    
    'If the user has selected "specific program", make sure a program is selected
    If (buttonIndex = 1) Then
        If (lstWindows.ListIndex = -1) Then lstWindows.ListIndex = 0
    End If
    
    UpdatePreview
    
End Sub

Private Sub UpdateVisibleContainer()
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        picContainer(i).Visible = (i = btsSource.ListIndex)
    Next i
    lblSecurity.Visible = (btsSource.ListIndex = 1)
End Sub

Private Sub chkChrome_Click()
    UpdatePreview
End Sub

Private Sub chkMinimize_Click()
    UpdatePreview
End Sub

Private Sub cmdBarMini_OKClick()
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "wholescreen", (btsSource.ListIndex = 0)
        .AddParam "minimizefirst", chkMinimize.Value
        If (btsSource.ListIndex <> 0) Then
            .AddParam "targethwnd", IIf(lstWindows.ListIndex >= 0, m_WindowHWnds.GetString(lstWindows.ListIndex), 0&)
            .AddParam "chrome", chkChrome.Value
            .AddParam "targetwindowname", IIf(lstWindows.ListIndex >= 0, lstWindows.List(lstWindows.ListIndex), g_Language.TranslateMessage("Screen capture"))
        End If
    End With
    
    'If the user wants a specific window captured, make sure they actually selected one from the list
    If (btsSource.ListIndex = 1) And (lstWindows.ListIndex = -1) Then
        PDMsgBox "Please select a window to capture.", vbInformation Or vbOKOnly, "Target window required"
        Exit Sub
    End If
        
    Me.Visible = False
    Process "Screen capture", False, cParams.GetParamString, UNDO_Nothing
    
End Sub

Private Sub cmdResetList_Click()
    FillListWithOpenApplications lstWindows
End Sub

Private Sub Form_Load()
            
    btsSource.AddItem "entire desktop", 0
    btsSource.AddItem "specific program", 1
    btsSource.ListIndex = 0
    UpdateVisibleContainer
    
    cmdResetList.AssignImage "generic_reset", , Interface.FixDPI(24), Interface.FixDPI(24)
    
    'Populate the "window is minimized" warning
    lblMinimizedWarning.Caption = g_Language.TranslateMessage("This program is currently minimized.  Restore it to normal size for best results.")
    If Not (g_Themer Is Nothing) Then
        lblMinimizedWarning.ForeColor = g_Themer.GetGenericUIColor(UI_ErrorRed)
    Else
        lblMinimizedWarning.ForeColor = RGB(232, 24, 20)
    End If
    
    lblSecurity.Caption = g_Language.TranslateMessage("Some programs (including Windows Store apps) do not allow direct screen captures.  If an application preview appears as a black square, you will need to take a full desktop screenshot, then manually crop the desired window region.")
    
    'Retrieve a list of all running programs (with some caveats; see the function for details)
    FillListWithOpenApplications lstWindows
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Wait just a moment before continuing, to give the corresponding menu time to animate away (otherwise it may
    ' get caught in the capture preview)
    VBHacks.SleepAPI 500
    
    'Render a preview of whichever item is currently selected
    UpdatePreview
    
End Sub

'Given a list box, fill it with a list of open applications.  Each application's name and hWnd is also cached in a
' pdStringStack object.
Private Function FillListWithOpenApplications(ByRef dstListbox As pdListBox) As Long
    
    dstListbox.Clear
    dstListbox.SetAutomaticRedraws False
    EnumWindows AddressOf ScreenCapture.EnumWindowsProc, 0&
    
    'Retrieve the list of window names and hWnds
    ScreenCapture.GetAllWindowNamesAndHWnds m_WindowNames, m_WindowHWnds
    
    'Fill the list box with the retrieved list of window names
    Dim i As Long
    For i = 0 To m_WindowNames.GetNumOfStrings - 1
        dstListbox.AddItem m_WindowNames.GetString(i), i
    Next i
    
    dstListbox.SetAutomaticRedraws True, True
    FillListWithOpenApplications = dstListbox.ListCount
    
End Function

Private Sub lstWindows_Click()
    UpdatePreview
End Sub

'Live previews of the screen capture are now provided
Private Sub UpdatePreview()
    
    'Prevent recursive redraws
    If m_SkipUpdate Then Exit Sub
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    Dim captureOK As Boolean: captureOK = False
    
    'Full screen capture was requested
    If (btsSource.ListIndex = 0) Then
        ScreenCapture.GetDesktopAsDIB tmpDIB
        captureOK = True
        
    'Specific window capture was requested
    Else
    
        If (lstWindows.ListIndex >= 0) Then
            
            'Make sure the function returns successfully; if a window is unloaded after the listbox has been
            ' filled, the function will (obviously) fail to capture the screen contents.
            Dim minimizeCheck As Boolean
            captureOK = ScreenCapture.GetHwndContentsAsDIB(tmpDIB, CLng(m_WindowHWnds.GetString(lstWindows.ListIndex)), chkChrome, minimizeCheck)
            
            If captureOK Then
                lblMinimizedWarning.Visible = minimizeCheck
            Else
                
                m_SkipUpdate = True
                
                'Remove the bad window from the UI and the background window name/hWnd lists
                Dim curIndex As Long
                curIndex = lstWindows.ListIndex
                lstWindows.RemoveItem curIndex
                m_WindowNames.RemoveStringByIndex curIndex
                m_WindowHWnds.RemoveStringByIndex curIndex
                
                m_SkipUpdate = False
                
                lblMinimizedWarning.Visible = False
                DisplayScreenCaptureError
                
            End If
            
        End If
    
    End If
    
    If captureOK Then
    
        If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
        If (Not g_Themer Is Nothing) Then
        
            m_PreviewDIB.CreateBlank picPreview.GetWidth, picPreview.GetHeight, 24, g_Themer.GetGenericUIColor(UI_Background)
            
            Dim dstWidth As Long, dstHeight As Long
            PDMath.ConvertAspectRatio tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, m_PreviewDIB.GetDIBWidth, m_PreviewDIB.GetDIBHeight, dstWidth, dstHeight, True
            
            Dim dstX As Long, dstY As Long
            dstX = (m_PreviewDIB.GetDIBWidth - dstWidth) \ 2
            dstY = (m_PreviewDIB.GetDIBHeight - dstHeight) \ 2
            
            GDI_Plus.GDIPlus_StretchBlt m_PreviewDIB, dstX, dstY, dstWidth, dstHeight, tmpDIB, 0, 0, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, dstCopyIsOkay:=True
            
            Set tmpDIB = Nothing
            picPreview.RequestRedraw True
            
        End If
    
    End If
    
    
End Sub

'If the user attempts to capture a window after it's been unloaded, warn them via this function
Private Sub DisplayScreenCaptureError()
    
    If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
    If (Not g_Themer Is Nothing) Then m_PreviewDIB.CreateBlank picPreview.GetWidth, picPreview.GetHeight, 24, g_Themer.GetGenericUIColor(UI_Background)
    
    Dim notifyFont As pdFont
    Set notifyFont = New pdFont
    notifyFont.SetFontFace Fonts.GetUIFontName()
    notifyFont.SetFontSize 14
    If (Not g_Themer Is Nothing) Then notifyFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextReadOnly)
    notifyFont.SetTextAlignment vbCenter
    notifyFont.CreateFontObject
    notifyFont.AttachToDC m_PreviewDIB.GetDIBDC
    
    notifyFont.FastRenderText picPreview.GetWidth / 2, picPreview.GetHeight / 2 - notifyFont.GetHeightOfString("ABCjqy"), g_Language.TranslateMessage("Unfortunately, that program has exited.")
    notifyFont.FastRenderText picPreview.GetWidth / 2, picPreview.GetHeight / 2, g_Language.TranslateMessage("Please select another one.")
    notifyFont.ReleaseFromDC
    
    picPreview.RequestRedraw True
    
End Sub

Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_PreviewDIB Is Nothing) Then m_PreviewDIB.AlphaBlendToDC targetDC
End Sub
