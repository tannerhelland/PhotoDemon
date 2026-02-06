VERSION 5.00
Begin VB.Form FormScreenVideoPrefs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Animated screen capture"
   ClientHeight    =   5985
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
   ScaleHeight     =   399
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
      Caption         =   "open in PhotoDemon (for further editing)"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   5250
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
      Top             =   3120
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
      Top             =   4200
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
      Top             =   2520
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
      Caption         =   "save directly to image file"
      Value           =   -1  'True
   End
   Begin PhotoDemon.pdSlider sldQuality 
      Height          =   735
      Left            =   6360
      TabIndex        =   9
      Top             =   2280
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      Caption         =   "quality"
      FontSizeCaption =   10
      Max             =   100
      Value           =   75
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   75
   End
   Begin PhotoDemon.pdButtonStrip btsFormat 
      Height          =   975
      Left            =   6360
      TabIndex        =   10
      Top             =   1440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1720
      Caption         =   "format"
      FontSizeCaption =   10
   End
End
Attribute VB_Name = "FormScreenVideoPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Animated screen capture dialog
'Copyright 2020-2026 by Tanner Helland
'Created: 01/July/20
'Last updated: 01/October/21
'Last update: new option for exporting to WebP (in addition to existing APNG support)
'
'This dialog is just a thin wrapper to FormScreenVideo.  It exists to allow the user to set
' some recording preferences; once those are set, the real magic happens in FormScreenVideo.
'
'As of 2021, animated PNG (APNG) and WebP are supported as recording targets.  GIF is not
' supported due to performance issues and poor quality (and frankly, better alternatives like
' LiceCAP).  If you *really* need an animated GIF, select the option to import the frames directly
' into PhotoDemon after recording; from there, you can save as an animated GIF (although this is not
' recommended since the other formats will produce better results with smaller file sizes).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsFormat_Click(ByVal buttonIndex As Long)
    ReflowInterface
End Sub

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
    
    'Store all relevant parameters inside a serializer
    Dim cSettings As pdSerialize
    Set cSettings = New pdSerialize
    With cSettings
        .AddParam "frame-rate", sldFrameRate.Value
        .AddParam "countdown", sldCountdown.Value
        .AddParam "show-cursor", (btsMouse.ListIndex >= 1)
        .AddParam "show-clicks", (btsMouse.ListIndex >= 2)
        .AddParam "loop-count", loopCount
        .AddParam "png-compression", sldCompression.Value
        .AddParam "webp-quality", sldQuality.Value
        .AddParam "save-to-disk", rdoAfter(1).Value
        .AddParam "file-format", IIf(btsFormat.ListIndex = 0, "png", "webp")
    End With
    
    'Launch the capture form, then note that the command bar will handle unloading this form
    FormScreenVideo.ShowDialog VarPtr(myRect), cSettings.GetParamString()
    
End Sub

Private Sub cmdBar_ResetClick()
    rdoAfter(1).Value = True
End Sub

Private Sub Form_Load()
    
    'If this dialog was previously used this session, we want to make sure the capture window
    ' has also been freed (as we need to reinitialize it)
    Set FormScreenVideo = Nothing
    
    'Prep any UI elements
    btsMouse.AddItem "no", 0
    btsMouse.AddItem "cursor only", 1
    btsMouse.AddItem "cursor and clicks", 2
    btsMouse.ListIndex = 1
    
    btsFormat.AddItem "PNG", 0
    btsFormat.AddItem "WebP", 1
    btsFormat.ListIndex = 1
    
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
    
    'Hide all format-specific options if "load into PD" is selected (instead of save to disk)
    btsFormat.Visible = rdoAfter(1).Value
    btsLoop.Visible = rdoAfter(1).Value
    
    If rdoAfter(0).Value Then
        sldCompression.Visible = False
        sldQuality.Visible = False
        sldLoop.Visible = False
        
    'Expose save-to-disk options
    Else
        
        Dim yPadding As Long, yPaddingTitle As Long
        yPadding = Interface.FixDPI(6)
        yPaddingTitle = Interface.FixDPI(12)
        
        Dim yOffset As Long
        yOffset = btsFormat.GetTop + btsFormat.GetHeight + yPadding
        
        '0 = PNG, 1 = WebP
        sldCompression.Visible = (btsFormat.ListIndex = 0)
        sldQuality.Visible = (btsFormat.ListIndex = 1)
        
        'PNG
        If (btsFormat.ListIndex = 0) Then
            sldCompression.SetTop yOffset
            yOffset = sldCompression.GetTop + sldCompression.GetHeight + yPadding
        
        'WebP
        Else
            sldQuality.SetTop yOffset
            yOffset = sldQuality.GetTop + sldQuality.GetHeight + yPadding
        End If
        
        btsLoop.SetTop yOffset
        yOffset = btsLoop.GetTop + btsLoop.GetHeight + yPadding
        
        sldLoop.Visible = (btsLoop.ListIndex = 2)
        If sldLoop.Visible Then
            sldLoop.SetTop yOffset
            yOffset = yOffset + sldLoop.GetHeight + yPaddingTitle
        End If
        
    End If
    
End Sub

Private Sub rdoAfter_Click(Index As Integer)
    ReflowInterface
End Sub
