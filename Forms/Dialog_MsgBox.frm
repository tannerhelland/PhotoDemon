VERSION 5.00
Begin VB.Form dialog_MsgBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
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
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdPictureBox picIcon 
      Height          =   615
      Left            =   240
      Top             =   1080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdContainer pnlBase 
      Height          =   750
      Left            =   0
      Top             =   5400
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1508
      Begin PhotoDemon.pdButton cmdBase 
         Height          =   510
         Index           =   0
         Left            =   7440
         TabIndex        =   1
         Top             =   120
         Width           =   1365
         _ExtentX        =   2910
         _ExtentY        =   1323
      End
      Begin PhotoDemon.pdButton cmdBase 
         Height          =   510
         Index           =   1
         Left            =   5880
         TabIndex        =   2
         Top             =   120
         Width           =   1365
         _ExtentX        =   2910
         _ExtentY        =   1323
      End
      Begin PhotoDemon.pdButton cmdBase 
         Height          =   510
         Index           =   2
         Left            =   4320
         TabIndex        =   0
         Top             =   120
         Width           =   1365
         _ExtentX        =   2910
         _ExtentY        =   1323
      End
   End
   Begin PhotoDemon.pdLabel lblMsg 
      Height          =   4725
      Left            =   1005
      Top             =   390
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8334
      Caption         =   ""
      ForeColor       =   2105376
      Layout          =   1
   End
End
Attribute VB_Name = "dialog_MsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Homebrew Message Box Replacement
'Copyright 2017-2026 by Tanner Helland
'Created: 15/August/17
'Last updated: 05/June/22
'Last update: fix Unicode window captions
'
'Theming a system message box is not worth the trouble.  Instead, we roll our own.  Any calls to PD's central
' "PDMsgBox" function are routed through here.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private m_userAnswer As VbMsgBoxResult

'Current button list.  (We parse this from the incoming VbMsgBoxStyle parameter.)  Note that,
' by design, only options currently used in PD are implemented.
Private Enum PD_MB_Buttons
    mb_OKOnly = 0
    mb_OKCancel = 1
    mb_YesNo = 2
    mb_YesNoCancel = 3
End Enum

#If False Then
    Private Const mb_OKOnly = 0, mb_OKCancel = 1, mb_YesNo = 2, mb_YesNoCancel = 3
#End If

Private Enum PD_MB_Icons
    mb_None = 0
    mb_Error = 1
    mb_Information = 2
    mb_Warning = 3
End Enum

#If False Then
    Private Const mb_None = 0, mb_Error = 1, mb_Information = 2, mb_Warning = 3
#End If

'Parsing flags is obnoxious in VB, so we cache button and icon settings when the dialog is initialized
Private m_CurButtons As PD_MB_Buttons, m_CurIcon As PD_MB_Icons

'Similarly, we cache the relevant icon, as necessary
Private m_iconDIB As pdDIB

'Local list of themable colors.  BY DESIGN, this is an exact copy of the command bar color set.  (This dialog is meant to
' mimic that one as much as possible.)
Private Enum PDCB_COLOR_LIST
    [_First] = 0
    PDCB_Background = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_userAnswer
End Property

'Because this control has to perform some compliated dynamic layout tasks, it also has to use some controls in non-standard ways
' (e.g. using a pdContainer as a command bar stand-in).  This also means we need to grab theme colors at run-time.
Private Sub InitColors()

    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCB_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDCommandBar", colorCount
    m_Colors.LoadThemeColor PDCB_Background, "Background", IDE_GRAY
    pnlBase.SetCustomBackcolor m_Colors.RetrieveColor(PDCB_Background, True)
    
    Dim i As Long
    For i = cmdBase.lBound To cmdBase.UBound
        cmdBase(i).BackgroundColor = m_Colors.RetrieveColor(PDCB_Background, True)
        cmdBase(i).UseCustomBackgroundColor = True
    Next i
    
End Sub

'Button detector called by ShowDialog; look there for details
Private Sub DetectButtonFlag(ByVal pButtons As VbMsgBoxStyle)

    If ((pButtons And &H7) = vbOKOnly) Then
        m_CurButtons = mb_OKOnly
    ElseIf ((pButtons And &H7) = vbOKCancel) Then
        m_CurButtons = mb_OKCancel
    ElseIf ((pButtons And &H7) = vbYesNo) Then
        m_CurButtons = mb_YesNo
    ElseIf ((pButtons And &H7) = vbYesNoCancel) Then
        m_CurButtons = mb_YesNoCancel
    Else
        m_CurButtons = mb_OKOnly
    End If
    
End Sub

'After button flags are parsed, call this function to set button visibility and position accordingly.
' See ShowDialog() for details.
Private Sub SetButtonVisibility(ByRef buttonLeft As Long, ByRef buttonRight As Long)

    buttonRight = cmdBase(0).GetLeft + cmdBase(0).GetWidth
    
    Select Case m_CurButtons
    
        Case mb_OKOnly
            cmdBase(0).Caption = g_Language.TranslateMessage("OK")
            cmdBase(0).Visible = True
            buttonLeft = cmdBase(0).GetLeft
            cmdBase(1).Visible = False
            cmdBase(2).Visible = False
            
        Case mb_OKCancel
            cmdBase(0).Caption = g_Language.TranslateMessage("Cancel")
            cmdBase(0).Visible = True
            cmdBase(1).Caption = g_Language.TranslateMessage("OK")
            cmdBase(1).Visible = True
            buttonLeft = cmdBase(1).GetLeft
            cmdBase(2).Visible = False
            
        Case mb_YesNo
            cmdBase(0).Caption = g_Language.TranslateMessage("No")
            cmdBase(0).Visible = True
            cmdBase(1).Caption = g_Language.TranslateMessage("Yes")
            cmdBase(1).Visible = True
            buttonLeft = cmdBase(1).GetLeft
            cmdBase(2).Visible = False
            
        Case mb_YesNoCancel
            cmdBase(0).Caption = g_Language.TranslateMessage("Cancel")
            cmdBase(0).Visible = True
            cmdBase(1).Caption = g_Language.TranslateMessage("No")
            cmdBase(1).Visible = True
            cmdBase(2).Caption = g_Language.TranslateMessage("Yes")
            cmdBase(2).Visible = True
            buttonLeft = cmdBase(2).GetLeft
    
    End Select
    
End Sub

'Parse the message box flags and look for icon settings.  By design, PD only supports a subset of icon types.
Private Sub SetIconVisibility(ByVal pButtons As VbMsgBoxStyle, ByRef iconActive As Boolean, ByRef dpiIconSize As Long)

    If ((pButtons And &H70&) = vbCritical) Then
        iconActive = True
        m_CurIcon = mb_Error
    
    'vbInformation/vbExclamation result in the same icon
    ElseIf ((pButtons And &H70&) = vbInformation) Then
        iconActive = True
        m_CurIcon = mb_Information
    ElseIf ((pButtons And &H70&) = vbExclamation) Then
        iconActive = True
        m_CurIcon = mb_Warning
    Else
        m_CurIcon = mb_None
    End If
    
    'If an icon is active, load and display it now
    If iconActive Then
        
        dpiIconSize = Interface.FixDPI(48)
        
        Dim iconFound As Boolean
        
        If (m_CurIcon = mb_Error) Then
            iconFound = IconsAndCursors.LoadResourceToDIB("generic_cancel", m_iconDIB, dpiIconSize, dpiIconSize, 0)
        ElseIf (m_CurIcon = mb_Warning) Then
            iconFound = IconsAndCursors.LoadResourceToDIB("generic_warning", m_iconDIB, dpiIconSize, dpiIconSize, 0)
        ElseIf (m_CurIcon = mb_Information) Then
            iconFound = IconsAndCursors.LoadResourceToDIB("generic_info", m_iconDIB, dpiIconSize, dpiIconSize, 0)
        End If
        
        If iconFound Then
            picIcon.SetSize dpiIconSize, dpiIconSize
            picIcon.Visible = True
        Else
            picIcon.Visible = False
            m_CurIcon = mb_None
            Set m_iconDIB = Nothing
            iconActive = False
        End If
        
    End If
    
    'If an icon is active, position it accordingly on the underlying form.  (Note that we must deal with horizontal
    ' positioning later, after we solve where the damn string fits.)
    If iconActive Then
        picIcon.SetTop dpiIconSize
    Else
        picIcon.Visible = False
    End If
    
End Sub

'DO NOT raise this form directly.  Instead, you must *always* use this ShowDialog wrapper.  It manually handles the
' messy process of laying out the form prior to actually showing it.
'
'Because message boxes maybe raised for errors that occur at random times (e.g. early in the program load process),
' this function will return TRUE if it can actually raise the message.  If it *can't*, you need to fall back to a
' default system message box.
Public Function ShowDialog(ByVal pMessage As String, ByVal pButtons As VbMsgBoxStyle, ByVal pTitle As String) As Boolean
    
    ShowDialog = True
    
    'Perform a few theme-related tasks before continuing.  (Specifically, grab relevant colors for the command bar
    ' at the bottom of the screen; we can't use an actual command bar, as we have to change button arrangements
    ' on-the-fly.)
    InitColors
    
    'Our first task is to parse the incoming MsgBoxStyle flag(s), and figure out what buttons and icons we need.
    
    'Let's start with buttons.  Besides figuring out button captions and visibility, we also want to calculate the
    ' minimum width required by the dialog.  This is one of several criteria we'll use to size the dialog horizontally.
    '
    '(Note that these button flags are not actual flags; they're treated as constants that each define a specific set of
    ' matching buttons.)
    DetectButtonFlag pButtons
    
    Dim buttonLeft As Long, buttonRight As Long
    SetButtonVisibility buttonLeft, buttonRight
    
    'Button order is now sorted and the correct ones are visible.  It would be convenient to also position them now,
    ' but because the buttons are *right* aligned, we can't fully position them until we calculate a final message
    ' box size (which hasn't happened yet).
    
    'Now that buttons are sorted, let's figure out if an icon is required.  (We do this before handling the string,
    ' because the lack of an icon means we can push the string all the way to the left.)
    Dim iconActive As Boolean, dpiIconSize As Long
    SetIconVisibility pButtons, iconActive, dpiIconSize
    
    'Now the messy one: string width.  Before proceeding, figure out where the left of the string sits;
    ' this obviously depends on whether or not an icon is active.
    Dim stringLeft As Long
    stringLeft = Interface.FixDPI(16)
    If iconActive Then stringLeft = stringLeft + picIcon.GetLeft + picIcon.GetWidth
    
    'We also want to know some basic metrics of the dialog itself, specifically how large we are allowed to
    ' physically make it.
    Dim finalFormWidth As Long, curCanvasWidth As Long, curScreenWidth As Long, curScreenRect As RectL
    
    'At present, the largest allowable size for a message box is the smaller of:
    ' 1) the primary canvas size, or...
    curCanvasWidth = FormMain.MainCanvas(0).GetCanvasWidth
    
    ' 2) 40% the width of the primary monitor.
    g_Displays.PrimaryDisplay.GetWorkingRect curScreenRect
    curScreenWidth = (curScreenRect.Right - curScreenRect.Left) * 0.4
    finalFormWidth = PDMath.Min2Int(curCanvasWidth, curScreenWidth)
    
    'If the width we've calculated is insanely small, give it some room to breathe
    If (finalFormWidth < Interface.FixDPI(300)) Then finalFormWidth = Interface.FixDPI(300)
    
    'Get a pdFont object in the current UI font; we need this for all string-measuring purposes
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(lblMsg.FontSize, lblMsg.FontBold, lblMsg.FontItalic, False)
    
    'Next, we need to branch based on the input string.  If the string contains line break chars, we'll use a
    ' word-wrap algorithm to measure the string.  If it *doesn't* contain line break chars, we'll see if it's short
    ' enough to fit on a single line.
    Dim i As Long, maxLineWidth As Long, useWordWrap As Boolean
    
    'If this is a single-line message, calculate a length for the incoming message string.  If it exceeds the
    ' form width we calculated above, active word wrap.
    If (InStr(1, pMessage, vbCrLf, vbBinaryCompare) = 0) Then
        maxLineWidth = tmpFont.GetWidthOfString(pMessage)
        useWordWrap = (finalFormWidth < maxLineWidth)
        
    'If this is *not* a single-line message, break the message into individual lines, and find the longest one.
    ' If that line is small enough, we'll use it as the size of our form, instead of using the largest allowable size.
    Else
    
        Dim strLines() As String
        strLines = Split(pMessage, vbCrLf, , vbBinaryCompare)
        
        Dim thisLineWidth As Long
        For i = LBound(strLines) To UBound(strLines)
            thisLineWidth = tmpFont.GetWidthOfString(strLines(i))
            If (thisLineWidth > maxLineWidth) Then maxLineWidth = thisLineWidth
        Next i
        
        useWordWrap = True
    
    End If
    
    Dim formHeight As Long
    Dim stringTop As Long, stringWidth As Long, stringHeight As Long
    Dim strPadding As Long
    
    'If the message needs wrapping, we next need to figure out how tall the string will be
    If useWordWrap Then
        
        'If the longest line in this message is shorter than the maximum allowed size of the message box,
        ' use that shorter string display size as the message box's on-screen width (instead of the maximum
        ' allowable width)
        strPadding = Interface.FixDPI(16)
        
        If (finalFormWidth < maxLineWidth) Then
            stringWidth = finalFormWidth - (stringLeft + strPadding)
        Else
            stringWidth = maxLineWidth + strPadding
        End If
        
        stringHeight = tmpFont.GetHeightOfWordwrapString(pMessage, stringWidth)
        
        'If this line only wraps once (e.g. it's just a long-ish sentence), and an icon is being displayed,
        ' we still need to see if the icon height is larger than the wrapped line height.
        If iconActive Then formHeight = picIcon.GetTop * 2 + dpiIconSize Else formHeight = 0
        formHeight = PDMath.Max2Int(formHeight, stringHeight + strPadding * 4)
        
        'Use the calculated form height to determine the string's position
        stringTop = (formHeight - stringHeight) \ 2
        
        '...and finally, move the label into position!
        lblMsg.SetPositionAndSize stringLeft, stringTop, stringWidth + 1, stringHeight
        
    'If this is a single-line message, yay for us!  The string width is our main limiting factor; figure out a width
    ' that works, while accounting for things like the icon's position (if any)
    Else
        
        strPadding = Interface.FixDPI(2)
        stringHeight = tmpFont.GetHeightOfString(pMessage) + strPadding
        
        'If the icon is active, we'll use it to size the form (instead of the string itself)
        If iconActive Then
            formHeight = picIcon.GetTop * 2 + dpiIconSize
        Else
            formHeight = stringHeight + Interface.FixDPI(32) * 2
        End If
        
        'Similarly, if an icon is active, we want to center the message relative to the icon
        If iconActive Then
            stringTop = picIcon.GetTop + (dpiIconSize - stringHeight) \ 2
        Else
            stringTop = (formHeight - stringHeight) \ 2
        End If
        
        lblMsg.SetPositionAndSize stringLeft, stringTop, maxLineWidth + strPadding, stringHeight
        
    End If
    
    'Use the final label size to calculate a final form width/height
    Dim formWidth As Long
    formWidth = lblMsg.GetLeft + lblMsg.GetWidth + Interface.FixDPI(32)
    
    'Note that short messages may be shorter than our button arrangement!  (Especially with Yes/No/Cancel.)
    ' As such, we need to ensure we have room for all buttons.
    Dim netButtonWidth As Long
    netButtonWidth = (buttonRight - buttonLeft) + Interface.FixDPI(16)
    If (netButtonWidth > formWidth) Then formWidth = netButtonWidth
    
    'Add the size of the bottom button panel to the calculated form height
    formHeight = formHeight + pnlBase.GetHeight
    
    'Next, add the size of the current system-mandated titlebar to the form
    Dim formHeightNonClient As Long
    Dim clientRect As winRect, windowRect As winRect
    If (Not g_WindowManager Is Nothing) Then
        g_WindowManager.GetClientWinRect Me.hWnd, clientRect
        g_WindowManager.GetWindowRect_API Me.hWnd, windowRect
        formHeightNonClient = formHeight + ((windowRect.y2 - windowRect.y1) - (clientRect.y2 - clientRect.y1))
    Else
        formHeightNonClient = formHeight
        ShowDialog = False
    End If
    
    'Next, you might think we'll position the window (e.g. set left/top), but this is actually handled automatically
    ' by the ShowPDDialog function, below.  We only need to set window dimensions here.
    If (Not g_WindowManager Is Nothing) Then
        g_WindowManager.SetSizeByHWnd Me.hWnd, formWidth, formHeightNonClient, True
        g_WindowManager.GetClientWinRect Me.hWnd, clientRect
    Else
        ShowDialog = False
    End If
    
    'Correctly position the button panel at the bottom of the form
    pnlBase.SetPositionAndSize 0, formHeight - pnlBase.GetHeight, formWidth, pnlBase.GetHeight
    
    'Position all buttons along the panel, in their final position(s)
    cmdBase(0).SetLeft clientRect.x2 - (cmdBase(0).GetWidth + FixDPI(16))
    cmdBase(1).SetLeft cmdBase(0).GetLeft - (cmdBase(1).GetWidth + FixDPI(8))
    cmdBase(2).SetLeft cmdBase(1).GetLeft - (cmdBase(2).GetWidth + FixDPI(8))
    
    'After all that work, we can finally apply the message caption and title.  (Note that the passed caption
    ' is *not* actually translated yet - that happens when themes get applied.)
    lblMsg.Caption = pMessage
    If (Not g_WindowManager Is Nothing) Then
        g_WindowManager.SetWindowCaptionW Me.hWnd, pTitle
    Else
        Me.Caption = pTitle
    End If
    
    'Provide a default answer of "cancel" (in the event that the user closes the dialog by clicking the
    ' "x" button in the top-right)
    m_userAnswer = vbCancel
    
    'Apply any custom styles to the form
    Interface.ApplyThemeAndTranslations Me
    
    'Display the form
    If ShowDialog Then ShowPDDialog vbModal, Me, True
    
End Function

Private Sub cmdBase_Click(Index As Integer)

    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    Me.Visible = False
    
    'Convert the button index into an actual response
    Select Case m_CurButtons
    
        Case mb_OKOnly
            m_userAnswer = vbOK
            
        Case mb_OKCancel
            If (Index = 1) Then
                m_userAnswer = vbOK
            ElseIf (Index = 0) Then
                m_userAnswer = vbCancel
            End If
            
        Case mb_YesNo
            If (Index = 1) Then
                m_userAnswer = vbYes
            ElseIf (Index = 0) Then
                m_userAnswer = vbNo
            End If
            
        Case mb_YesNoCancel
            If (Index = 2) Then
                m_userAnswer = vbYes
            ElseIf (Index = 1) Then
                m_userAnswer = vbNo
            ElseIf (Index = 0) Then
                m_userAnswer = vbCancel
            End If
    
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Render the icon when requested
Private Sub picIcon_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    GDI.FillRectToDC targetDC, 0, 0, ctlWidth, ctlHeight, g_Themer.GetGenericUIColor(UI_Background)
    If (Not m_iconDIB Is Nothing) Then m_iconDIB.AlphaBlendToDC targetDC, , 0, 0
End Sub
