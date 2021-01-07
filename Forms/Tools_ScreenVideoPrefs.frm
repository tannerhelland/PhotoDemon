VERSION 5.00
Begin VB.Form FormRecordAPNGPrefs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Animated screen capture (APNG)"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11775
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
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   785
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdRadioButton rdoAfter 
      Height          =   375
      Index           =   0
      Left            =   6000
      TabIndex        =   7
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "open it in PhotoDemon (for further editing)"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   4215
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "recording settings"
      FontSize        =   12
   End
   Begin PhotoDemon.pdSlider sldFrameRate 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1296
      Caption         =   "target frame rate (fps)"
      FontSizeCaption =   10
      Min             =   0.1
      Max             =   30
      SigDigits       =   1
      Value           =   10
      NotchPosition   =   2
      NotchValueCustom=   10
   End
   Begin PhotoDemon.pdButtonStrip btsLoop 
      Height          =   975
      Left            =   6360
      TabIndex        =   1
      Top             =   2280
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1720
      Caption         =   "repeat final animation"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdSlider sldLoop 
      Height          =   735
      Left            =   6360
      TabIndex        =   2
      Top             =   3360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      Caption         =   "repeat count"
      FontSizeCaption =   10
      Min             =   1
      Max             =   65535
      ScaleStyle      =   2
      Value           =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdButtonStrip btsMouse 
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1720
      Caption         =   "record mouse actions"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdSlider sldCompression 
      Height          =   735
      Left            =   6360
      TabIndex        =   4
      Top             =   1440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      Caption         =   "compression level"
      FontSizeCaption =   10
      Max             =   12
      Value           =   9
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   9
   End
   Begin PhotoDemon.pdSlider sldCountdown 
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1296
      Caption         =   "countdown before starting (in seconds)"
      FontSizeCaption =   10
      Max             =   20
      Value           =   3
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   3
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   1
      Left            =   6000
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "when recording finishes"
      FontSize        =   12
   End
   Begin PhotoDemon.pdRadioButton rdoAfter 
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   8
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "save it directly to an animated PNG file"
      Value           =   -1  'True
   End
End
Attribute VB_Name = "FormRecordAPNGPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Animated screen capture dialog
'Copyright 2020-2021 by Tanner Helland
'Created: 01/July/20
'Last updated: 24/July/20
'Last update: add preference for a countdown before recording; set to 0 to record immediately
'
'This dialog is just a thin wrapper to FormRecordAPNG.  It exists to allow the user to set
' some recording preferences; once those are set, the real magic happens in FormRecordAPNG.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsLoop_Click(ByVal buttonIndex As Long)
    ReflowInterface
End Sub

Private Sub cmdBar_OKClick()
    
    'Before hiding this window, retrieve our current window position; the launched
    ' screen recording window will use this to position itself the first time it's invoked
    Dim myRect As winRect
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.GetWindowRect_API Me.hWnd, myRect
    
    'Because this dialog is modal, it needs to be hidden before we invoke a modeless dialog
    Me.Hide
    
    'Also hide the main PhotoDemon window
    FormMain.WindowState = vbMinimized
    
    'The loop setting is a little weird.
    ' 0 = loop infinitely, 1 = loop once, 2+ = loop that many times exactly
    Dim loopCount As Long
    If (btsLoop.ListIndex = 0) Then
        loopCount = 1
    ElseIf (btsLoop.ListIndex = 1) Then
        loopCount = 0
    Else
        loopCount = CLng(sldLoop.Value + 1)
    End If
    
    'Launch the capture form, then note that the command bar will handle unloading this form
    FormRecordAPNG.ShowDialog VarPtr(myRect), sldFrameRate.Value, loopCount, (btsMouse.ListIndex >= 1), (btsMouse.ListIndex >= 2), sldCompression.Value, sldCountdown.Value, rdoAfter(1).Value
    
End Sub

Private Sub cmdBar_ResetClick()
    rdoAfter(1).Value = True
End Sub

Private Sub Form_Load()
    
    'If this dialog was previously used this session, we want to make sure the capture window
    ' has also been freed (as we need to reinitialize it)
    Set FormRecordAPNG = Nothing
    
    'Prep any UI elements
    btsMouse.AddItem "no", 0
    btsMouse.AddItem "cursor only", 1
    btsMouse.AddItem "cursor and clicks", 2
    btsMouse.ListIndex = 1
    
    btsLoop.AddItem "none", 0
    btsLoop.AddItem "forever", 1
    btsLoop.AddItem "custom", 2
    btsLoop.ListIndex = 0
    
    'Apply custom themes
    Interface.ApplyThemeAndTranslations Me
    
    'With theming handled, reflow the interface one final time before displaying the window
    ReflowInterface
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Interface.ReleaseFormTheming Me
End Sub

Private Sub ReflowInterface()

    Dim yPadding As Long, yPaddingTitle As Long
    yPadding = Interface.FixDPI(8)
    yPaddingTitle = Interface.FixDPI(12)
    
    Dim yOffset As Long
    yOffset = btsLoop.GetTop + btsLoop.GetHeight + yPadding
    sldLoop.Visible = (btsLoop.ListIndex = 2)
    If sldLoop.Visible Then
        sldLoop.SetTop yOffset
        yOffset = yOffset + sldLoop.GetHeight + yPaddingTitle
    Else
        yOffset = yOffset - yPadding + yPaddingTitle
    End If
    
End Sub

Private Sub rdoAfter_Click(Index As Integer)
    SyncAfterOptions
End Sub

Private Sub SyncAfterOptions()
    sldCompression.Visible = rdoAfter(1).Value
    btsLoop.Visible = rdoAfter(1).Value
    sldLoop.Visible = rdoAfter(1).Value And (btsLoop.ListIndex = 2)
End Sub
