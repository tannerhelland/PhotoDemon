VERSION 5.00
Begin VB.Form dialog_ColorSelector 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Change color"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11535
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
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   769
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   2
      Top             =   5295
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdButtonToolbox cmdCapture 
      Height          =   600
      Left            =   10320
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1058
      AutoToggle      =   -1  'True
   End
   Begin PhotoDemon.pdTextBox txtHex 
      Height          =   315
      Left            =   6480
      TabIndex        =   6
      Top             =   3735
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Text            =   "abcdef"
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   10680
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   7
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   10080
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   8
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   9480
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   9
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   8880
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   10
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   8280
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   11
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   7680
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   24
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   7080
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   25
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   6480
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   26
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picSampleHSV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   22
      Top             =   1320
      Width           =   3735
   End
   Begin VB.PictureBox picSampleHSV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   20
      Top             =   720
      Width           =   3735
   End
   Begin VB.PictureBox picSampleHSV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   18
      Top             =   120
      Width           =   3735
   End
   Begin VB.PictureBox picSampleRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   17
      Top             =   3120
      Width           =   3735
   End
   Begin VB.PictureBox picSampleRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   15
      Top             =   2520
      Width           =   3735
   End
   Begin VB.PictureBox picSampleRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   13
      Top             =   1920
      Width           =   3735
   End
   Begin VB.PictureBox picOriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   5
      Top             =   4560
      Width           =   3735
   End
   Begin VB.PictureBox picCurrent 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   4
      Top             =   4080
      Width           =   3735
   End
   Begin VB.PictureBox picHue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   4320
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin PhotoDemon.pdSpinner tudRGB 
      Height          =   345
      Index           =   0
      Left            =   10320
      TabIndex        =   12
      Top             =   1905
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   255
   End
   Begin PhotoDemon.pdSpinner tudRGB 
      Height          =   345
      Index           =   1
      Left            =   10320
      TabIndex        =   14
      Top             =   2505
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   255
   End
   Begin PhotoDemon.pdSpinner tudRGB 
      Height          =   345
      Index           =   2
      Left            =   10320
      TabIndex        =   16
      Top             =   3105
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   255
   End
   Begin PhotoDemon.pdSpinner tudHSV 
      Height          =   345
      Index           =   0
      Left            =   10320
      TabIndex        =   19
      Top             =   105
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   359
   End
   Begin PhotoDemon.pdSpinner tudHSV 
      Height          =   345
      Index           =   1
      Left            =   10320
      TabIndex        =   21
      Top             =   705
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   100
   End
   Begin PhotoDemon.pdSpinner tudHSV 
      Height          =   345
      Index           =   2
      Left            =   10320
      TabIndex        =   23
      Top             =   1305
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   100
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   600
      Index           =   9
      Left            =   5085
      Top             =   4680
      Width           =   1305
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   1
      Caption         =   "recent colors:"
      ForeColor       =   0
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   720
      Index           =   8
      Left            =   5070
      Top             =   3765
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   1270
      Alignment       =   1
      Caption         =   "HTML / CSS:"
      ForeColor       =   0
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   7
      Left            =   5130
      Top             =   3180
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "blue:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   6
      Left            =   5115
      Top             =   2580
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "green:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   5
      Left            =   5085
      Top             =   1980
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "red:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   4
      Left            =   5040
      Top             =   1380
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "value:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   3
      Left            =   5115
      Top             =   780
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "saturation:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   2
      Left            =   5055
      Top             =   180
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "hue:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   525
      Index           =   1
      Left            =   30
      Top             =   4650
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   926
      Alignment       =   1
      Caption         =   "original:"
      FontSize        =   11
      ForeColor       =   0
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   405
      Index           =   0
      Left            =   75
      Top             =   4170
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   714
      Alignment       =   1
      Caption         =   "current:"
      FontSize        =   11
      ForeColor       =   0
      Layout          =   1
   End
End
Attribute VB_Name = "dialog_ColorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Selection Dialog
'Copyright 2013-2016 by Tanner Helland
'Created: 11/November/13
'Last updated: 14/May/16
'Last update: improve real-time handling of hex input
'
'Basic color selection dialog.  At present, the dialog is heavily modeled after GIMP's color selection dialog.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'The original color when the dialog was first loaded
Private oldColor As Long

'The new color selected by the user, if any
Private newUserColor As Long

'pdDIB for the primary color box (luminance/saturation) on the left
Private primaryBox As pdDIB

'pdDIB for the hue box on the rihgt
Private hueBox As pdDIB

'Currently selected color, including RGB and HSL attributes
Private curColor As Long
Private curRed As Long, curGreen As Long, curBlue As Long
Private curHue As Double, curSaturation As Double, curValue As Double

'One DIB for each of the individual color sample boxes
Private sRed As pdDIB, sGreen As pdDIB, sBlue As pdDIB
Private sHue As pdDIB, sSaturation As pdDIB, sValue As pdDIB

'Left/right/up arrows for the hue and color boxes; these are 7x13 (or 13x7) and loaded from the resource at run-time
Private leftSideArrow As pdDIB, rightSideArrow As pdDIB, upArrow As pdDIB

'A temporary DIB for drawing any other elements
Private m_tmpDIB As pdDIB

'Changing the various text boxes resyncs the dialog, unless this parameter is set.  (We use it to prevent
' infinite resyncs.)
Private m_suspendTextResync As Boolean, m_suspendHexInput As Boolean

Private Enum colorCheckType
    ccRed = 0
    ccGreen = 1
    ccBlue = 2
    ccHue = 3
    ccSaturation = 4
    ccValue = 5
End Enum

#If False Then
    Private Const ccRed = 0, ccGreen = 1, ccBlue = 2, ccHue = 3, ccSaturation = 4, ccValue = 5
#End If

'Recently used colors are loaded/saved from a custom XML file
Private xmlEngine As pdXML

'The file where we'll store recent color data when the program is closed.  This file will be saved in the
' /Data/Presets/ folder.
Private XMLFilename As String

'Because we have to color manage everything on this screen, we can't simply set picture box backcolors to match the
' recent color list.  We have to create special DIBs of each color, then blt those onto the respective boxes.
Private recentColors() As Long

'If a user control spawned this dialog, it will pass itself as a reference.  We can then send color updates back
' to the control, allowing for real-time updates on the screen despite a modal dialog being raised!
Private parentColorControl As pdColorSelector

'***
'All declarations below this line are necessary for capturing a color from an arbitrary point on the screen.

'Many thanks to VB coder LaVolpe for his original implementation of this technique, available here:
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=52878&lngWId=1

'Is capture mode active?
Private screenCaptureActive As Boolean

'A subclasser is required to retrieve mouse events
Private cSubclass As cSelfSubHookCallback

'Constants for sending and receiving various mouse events
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Private Const MOUSEEVENTF_LEFTUP As Long = &H4

'Various API declarations for capturing mosue events and retrieving colors accordingly
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal srcDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

'pdInputMouse makes it easier to deal with a custom hand cursor for the many picture boxes on the form
Private WithEvents m_MouseEvents As pdInputMouse
Attribute m_MouseEvents.VB_VarHelpID = -1

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The newly selected color (if any) is returned via this property
Public Property Get NewlySelectedColor() As Long
    NewlySelectedColor = newUserColor
End Property

'During screen capture, this sub is used to update the current color under the mouse cursor.  Note that such rapid
' redraws will be a bit slow in the IDE - compile for better results.
Private Sub UpdateHoveredColor()

    'Retrieve the current mouse location
    Dim mouseLocation As POINTAPI
    GetCursorPos mouseLocation
    
    'Retrieve a handle to the screen
    Dim srcDC As Long
    srcDC = GetWindowDC(GetDesktopWindow())
    
    'Retrieve the color at the given pixel location
    Dim hColor As Long
    hColor = GetPixel(srcDC, mouseLocation.x, mouseLocation.y)
    
    'Extract RGB components from the Long-type color
    curRed = Colors.ExtractRed(hColor)
    curGreen = Colors.ExtractGreen(hColor)
    curBlue = Colors.ExtractBlue(hColor)
    
    'Calculate new HSV values to match
    RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
    
    'Resync the interface to match the new value!
    SyncInterfaceToCurrentColor
    
End Sub

'Activate or deactivate "capture a color from the screen" mode.  Optionally, the original color can be restored if the
' user decides they don't actually want to capture a color from the screen.
Private Sub ToggleCaptureMode(ByVal toActivate As Boolean)

    Static prevCursorHandle As Long
    
    'Activation is requested!
    If toActivate And (Not screenCaptureActive) And g_IsProgramRunning Then
        
        'Disable any current capture or cursor handlers
        'cmdCapture.OverrideMouseCapture True
        PrepSpecialMouseHandling False
        
        screenCaptureActive = True
        
        'Start capture mode, and use a fake mouse click to retain capture even though the user isn't holding
        ' down a mouse button.  (Thank you to VB coder LaVolpe for this handy trick!)
        ReleaseCapture
        SetCapture Me.hWnd
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        
        'Assign a color dropper cursor to the screen
        prevCursorHandle = SetCursor(RequestCustomCursor("C_PIPETTE", 0, 0))
        
        'Start subclassing relevant mouse event messages
        Set cSubclass = New cSelfSubHookCallback
        cSubclass.ssc_Subclass Me.hWnd, , , Me
        cSubclass.ssc_AddMsg Me.hWnd, MSG_BEFORE, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP ', WM_RBUTTONDOWN, WM_RBUTTONUP
    
    End If
    
    'Deactivation is requested!
    If (Not toActivate) And screenCaptureActive And g_IsProgramRunning Then
    
        screenCaptureActive = False
        
        'As part of fooling the mouse capture function (see above), we must now make sure to release the mouse
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        ReleaseCapture
        
        'Restore the original cursor
        SetCursor prevCursorHandle
        
        'Terminate and unload the subclasser
        cSubclass.ssc_UnSubclass Me.hWnd
        cSubclass.ssc_Terminate
        Set cSubclass = Nothing
        
        'Re-enable any current capture or cursor handlers
        'cmdCapture.OverrideMouseCapture False
        PrepSpecialMouseHandling True
        
    End If

End Sub

Private Sub cmdBarMini_CancelClick()
    
    userAnswer = vbCancel
    Me.Hide
    
End Sub

Private Sub cmdBarMini_OKClick()
    
    'Store the newUserColor value (which the dialog handler will use to return the selected color)
    newUserColor = RGB(curRed, curGreen, curBlue)
    
    'Save the current list of recently used colors
    SaveRecentColorList
    
    userAnswer = vbOK
    Me.Hide
    
End Sub

'Capture color button
Private Sub cmdCapture_Click()
    
    'If it isn't already active, start mouse capture mode
    If (Not screenCaptureActive) Then
        ToggleCaptureMode True
    End If
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByVal initialColor As Long, Optional ByRef callingControl As pdColorSelector = Nothing)

    'Store a reference to the calling control (if any)
    Set parentColorControl = callingControl

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Load the left/right side hue box arrow images from the resource file
    Set leftSideArrow = New pdDIB
    Set rightSideArrow = New pdDIB
    Set upArrow = New pdDIB
    
    LoadResourceToDIB "CLR_ARROW_L", leftSideArrow
    LoadResourceToDIB "CLR_ARROW_R", rightSideArrow
    LoadResourceToDIB "CLR_ARROW_U", upArrow
    
    'The passed color may be an OLE constant rather than an actual RGB triplet, so convert it now.
    initialColor = ConvertSystemColor(initialColor)
    
    'Cache the currentColor parameter so we can access it elsewhere
    oldColor = initialColor
    
    'Render the old color to the screen.  Note that we must use a temporary DIB for this; otherwise, the color will
    ' not be properly color managed.
    Dim tmpDIB As New pdDIB
    tmpDIB.CreateBlank picOriginal.ScaleWidth, picOriginal.ScaleHeight, 24, oldColor
    tmpDIB.RenderToPictureBox picOriginal
    
    'Sync all current color values to the initial color
    curColor = initialColor
    curRed = Colors.ExtractRed(initialColor)
    curGreen = Colors.ExtractGreen(initialColor)
    curBlue = Colors.ExtractBlue(initialColor)
    
    RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
    
    'Synchronize the interface to this new color
    SyncInterfaceToCurrentColor
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
        
    'Render the vertical hue box
    DrawHueBox
    
    'Apply extra tooltips to certain controls
    cmdCapture.AssignImage "CS_FROM_SCREEN"
    cmdCapture.AssignTooltip "Click this button to enable color capturing from anywhere on the screen."
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Manually assign a hand cursor to the various picture boxes.
    PrepSpecialMouseHandling True
    
    'Initialize an XML engine, which we will use to read/write recent color data to file
    Set xmlEngine = New pdXML
    
    'The XML file will be stored in the Preset path (/Data/Presets)
    XMLFilename = g_UserPreferences.GetPresetPath & "Color_Selector.xml"
    
    'If an XML file exists, load its contents now
    LoadRecentColorList
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

'Capture-from-screen mode requires special handling
Private Sub PrepSpecialMouseHandling(ByVal handleMode As Boolean)
    
    If g_IsProgramRunning And handleMode Then
    
        Set m_MouseEvents = New pdInputMouse
        
        With m_MouseEvents
        
            .AddInputTracker picColor.hWnd, True, False, False, True, True
            .AddInputTracker picHue.hWnd, True, False, False, True, True
            .AddInputTracker picOriginal.hWnd, True, False, False, True, True
            
            Dim i As Long
            For i = picRecColor.lBound To picRecColor.UBound
                .AddInputTracker picRecColor(i).hWnd, True, False, False, True, True
            Next i
            
            For i = picSampleHSV.lBound To picSampleHSV.UBound
                .AddInputTracker picSampleHSV(i).hWnd, True, False, False, True, True
            Next i
            
            For i = picSampleRGB.lBound To picSampleRGB.UBound
                .AddInputTracker picSampleRGB(i).hWnd, True, False, False, True, True
            Next i
            
            .SetSystemCursor IDC_HAND
            
        End With
        
    Else
        Set m_MouseEvents = Nothing
    End If
    
End Sub

'If the user has used the color selector before, their last-used colors will be stored to an XML file.  Use this function
' to load those colors.
Private Sub LoadRecentColorList()

    'Start by seeing if an XML file with previously saved color data exists
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(XMLFilename) Then
        
        'Attempt to load and validate the current file; if we can't, create a new, blank XML object
        If Not xmlEngine.LoadXMLFile(XMLFilename) Then
            Debug.Print "List of recent colors appears to be invalid.  A new recent color list has been created."
            ResetXMLData
        End If
        
    Else
        ResetXMLData
    End If
        
    'We are now ready to load the actual color data from file.
    
    'The XML engine will do most the heavy lifting for this task.  We pass it a String array, and it fills it with
    ' all values corresponding to the given tag name and attribute.  (We must do this dynamically, because we don't
    ' know how many recent colors are actually saved - it could be anywhere from 0 to picRecColor.Count.)
    Dim allRecentColors() As String
    Dim numColors As Long
    
    If xmlEngine.FindAllAttributeValues(allRecentColors, "colorEntry", "id") Then
        
        numColors = UBound(allRecentColors) + 1
        
        'Make sure the file does not contain more entries than are allowed (shouldn't theoretically be possible,
        ' but it doesn't hurt to check).
        If numColors > picRecColor.Count Then numColors = picRecColor.Count
        
    'No recent color entries were found.
    Else
        numColors = 0
    End If
    
    Dim i As Long
    
    'If one or more recent colors were found, load them now.
    If numColors > 0 Then
        
        ReDim recentColors(0 To numColors - 1) As Long
        
        'Load the actual colors from the XML file
        Dim tmpColorString As String
        
        For i = 0 To numColors - 1
        
            'Retrieve the color, in string format
            tmpColorString = xmlEngine.GetUniqueTag_String("color", , , "colorEntry", "id", allRecentColors(i))
            
            'Translate the color into a long, and update the corresponding picture box
            If Len(tmpColorString) <> 0 Then recentColors(i) = CLng(tmpColorString)
            
        Next i
    
    'No recent colors were found.  Populate the list with a few default values.
    Else
        
        ReDim recentColors(0 To picRecColor.Count - 1)
        recentColors(0) = RGB(0, 0, 255)
        recentColors(1) = RGB(0, 255, 0)
        recentColors(2) = RGB(255, 0, 0)
        recentColors(3) = RGB(255, 0, 255)
        recentColors(4) = RGB(0, 255, 255)
        recentColors(5) = RGB(255, 255, 0)
        recentColors(6) = 0
        recentColors(7) = RGB(255, 255, 255)
    End If
    
    'For color management reasons, we must use DIBs to copy colors onto the recent color picture boxes
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Render the recent color list to their respective picture boxes
    For i = 0 To picRecColor.Count - 1
    
        If i <= UBound(recentColors) Then
            tmpDIB.CreateBlank picRecColor(i).ScaleWidth, picRecColor(i).ScaleHeight, 24, recentColors(i)
            tmpDIB.RenderToPictureBox picRecColor(i)
        End If
    
    Next i

End Sub

'Save the current list of last-used colors to an XML file, adding the color presently selected as the most-recent entry.
Private Sub SaveRecentColorList()
    
    'Reset whatever XML data we may have stored at present - we will be rewriting the full MRU file from scratch.
    ResetXMLData
    
    'We now need to update the colors array with the new color entry.  Start by seeing if this color is already in the
    ' array.  If it is, simply swap its order.
    Dim i As Long, j As Long
    
    Dim colorFound As Boolean
    colorFound = False
    
    For i = 0 To picRecColor.Count - 1
    
        'This color already exists in the list.  Move it to the top of the list, and shift everything else downward.
        If recentColors(i) = newUserColor Then
            
            colorFound = True
            
            For j = i To 1 Step -1
                recentColors(j) = recentColors(j - 1)
            Next j
            
            recentColors(0) = newUserColor
            Exit For
            
        End If
        
    Next i
    
    'If this color is not already in the list, add it now.
    If Not colorFound Then
        
        For i = picRecColor.Count - 1 To 1 Step -1
            recentColors(i) = recentColors(i - 1)
        Next i
        
        recentColors(0) = newUserColor
    
    End If
    
    'Add all color entries to the XML engine
    For i = 0 To UBound(recentColors)
        xmlEngine.WriteTagWithAttribute "colorEntry", "id", Str(i), "", True
        xmlEngine.WriteTag "color", recentColors(i)
        xmlEngine.CloseTag "colorEntry"
        xmlEngine.WriteBlankLine
    Next i
    
    'With the XML file now complete, write it out to file
    xmlEngine.WriteXMLToFile XMLFilename
    
End Sub

'When creating a new recent coclors file, or overwriting a corrupt one, use this to initialize the new XML file.
Private Sub ResetXMLData()
    xmlEngine.PrepareNewXML "Recent colors"
    xmlEngine.WriteBlankLine
    xmlEngine.WriteComment "Everything past this point is recent color data.  Entries are sorted in reverse chronological order."
    xmlEngine.WriteBlankLine
End Sub

'Refresh the various color box cursors when the mouse enters
Private Sub m_MouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseEvents.SetSystemCursor IDC_HAND, m_MouseEvents.GetLastHwnd
End Sub

Private Sub m_MouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseEvents.SetSystemCursor IDC_HAND, m_MouseEvents.GetLastHwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'As a failsafe, check for ongoing subclassing and release as necessary
    If screenCaptureActive Then ToggleCaptureMode False
    ReleaseFormTheming Me
    
End Sub

'The hue box only needs to be drawn once, when the dialog is first created
Private Sub DrawHueBox()
    
    Dim hVal As Double
    Dim r As Long, g As Long, b As Long
    
    'Because we want the hue box to be color-managed, we must create it as a DIB, then render it to the screen later
    Set hueBox = New pdDIB
    hueBox.CreateBlank picHue.ScaleWidth, picHue.ScaleHeight
    
    'Simple gradient-ish code implementation of drawing hue
    Dim y As Long
    For y = 0 To hueBox.GetDIBHeight
    
        'Based on our x-position, gradient a value between -1 and 5
        hVal = y / hueBox.GetDIBHeight
        
        'Generate a full-saturation hue for this position
        HSVtoRGB hVal, 1, 1, r, g, b
        
        'Draw the color
        GDI.DrawLineToDC hueBox.GetDIBDC, 0, y, picHue.ScaleWidth, y, RGB(r, g, b)
        
    Next y
    
    'Render the hue to the picture box, which will also activate color management
    hueBox.RenderToPictureBox picHue
    
End Sub

'When *all* current color values are updated and valid, use this function to synchronize the interface to match
' their appearance.
Private Sub SyncInterfaceToCurrentColor()
    
    Me.Picture = LoadPicture("")
    
    'Start by drawing the primary box (luminance/saturation) using the current values
    If (primaryBox Is Nothing) Then Set primaryBox = New pdDIB
    If (primaryBox.GetDIBWidth <> picColor.ScaleWidth) Or (primaryBox.GetDIBHeight <> picColor.ScaleHeight) Then
        primaryBox.CreateBlank picColor.ScaleWidth, picColor.ScaleHeight
    End If
    
    Dim pImageData() As Byte
    Dim psa As SAFEARRAY2D
    PrepSafeArray psa, primaryBox
    CopyMemory ByVal VarPtrArray(pImageData()), VarPtr(psa), 4
    
    Dim x As Long, y As Long, quickX As Long
    
    Dim tmpR As Long, tmpG As Long, tmpB As Long
    
    Dim loopWidth As Long, loopHeight As Long
    loopWidth = (primaryBox.GetDIBWidth - 1) * 3
    loopHeight = primaryBox.GetDIBHeight - 1
    
    Dim lineSaturation As Double
    
    'To improve performance, pre-calculate all value variants
    Dim valuePresets() As Double
    ReDim valuePresets(0 To loopWidth) As Double
    For x = 0 To loopWidth Step 3
        valuePresets(x) = x / loopWidth
    Next x
    
    For y = 0 To loopHeight
        
        'Saturation is consistent for each y-position
        lineSaturation = (loopHeight - y) / loopHeight
        
    For x = 0 To loopWidth Step 3
        
        'The x-axis position determines value (0 -> 1)
        'The y-axis position determines saturation (1 -> 0)
        HSVtoRGB curHue, lineSaturation, valuePresets(x), tmpR, tmpG, tmpB
        
        pImageData(x, y) = tmpB
        pImageData(x + 1, y) = tmpG
        pImageData(x + 2, y) = tmpR
        
    Next x
    Next y
    
    'With our work complete, point the ImageData() array away from the DIBs and deallocate it
    CopyMemory ByVal VarPtrArray(pImageData), 0&, 4
    
    'We now want to draw a circle around the point where the user's current color resides
    GDIPlusDrawCanvasCircle primaryBox.GetDIBDC, curValue * (loopWidth / 3), (1 - curSaturation) * loopHeight, FixDPI(7)
        
    'Render the primary color box
    primaryBox.RenderToPictureBox picColor
    
    'Render the current color box.  Note that we must use a temporary DIB for this; otherwise, the color will
    ' not be properly color managed.
    
    If (m_tmpDIB Is Nothing) Then Set m_tmpDIB = New pdDIB
    If (m_tmpDIB.GetDIBWidth <> picCurrent.ScaleWidth) Or (m_tmpDIB.GetDIBHeight <> picCurrent.ScaleHeight) Then
        m_tmpDIB.CreateBlank picCurrent.ScaleWidth, picCurrent.ScaleHeight, 24, RGB(curRed, curGreen, curBlue)
    Else
        GDI_Plus.GDIPlusFillDIBRect m_tmpDIB, 0, 0, m_tmpDIB.GetDIBWidth, m_tmpDIB.GetDIBHeight, RGB(curRed, curGreen, curBlue)
    End If
    
    m_tmpDIB.RenderToPictureBox picCurrent
    
    'Synchronize all text boxes to their current values
    RedrawAllTextBoxes
    
    'Position the arrows along the hue box properly according to the current hue
    Dim hueY As Long
    hueY = picHue.Top + 1 + (curHue * picHue.ScaleHeight)
    
    leftSideArrow.AlphaBlendToDC Me.hDC, , picHue.Left - leftSideArrow.GetDIBWidth, hueY - (leftSideArrow.GetDIBHeight \ 2)
    rightSideArrow.AlphaBlendToDC Me.hDC, , picHue.Left + picHue.Width, hueY - (rightSideArrow.GetDIBHeight \ 2)
    Me.Picture = Me.Image
    Me.Refresh
    
    'If we have a reference to a parent color selection user control, notify that control that the user's color
    ' has changed.
    If Not (parentColorControl Is Nothing) Then
        parentColorControl.NotifyOfLiveColorChange RGB(curRed, curGreen, curBlue)
    End If
    
End Sub

'Use this sub to resync all text boxes to the current RGB/HSV values
Private Sub RedrawAllTextBoxes()

    'We don't want the _Change events for the text boxes firing while we resync them, so we disable any resyncing in advance
    m_suspendTextResync = True
    
    'Start by matching up the text values themselves
    tudRGB(0) = curRed
    tudRGB(1) = curGreen
    tudRGB(2) = curBlue
    
    tudHSV(0) = curHue * 359
    tudHSV(1) = curSaturation * 100
    tudHSV(2) = curValue * 100
    
    'Next, prepare some universal values for the arrow image offsets
    Dim arrowOffset As Long
    arrowOffset = (upArrow.GetDIBWidth \ 2) - 1
    
    Dim leftOffset As Long
    leftOffset = picSampleRGB(0).Left
    
    Dim widthCheck As Long
    widthCheck = picSampleRGB(0).ScaleWidth - 1
    
    'Next, redraw all marker arrows
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + ((curRed / 255) * widthCheck) - arrowOffset, picSampleRGB(0).Top + picSampleRGB(0).Height
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + ((curGreen / 255) * widthCheck) - arrowOffset, picSampleRGB(1).Top + picSampleRGB(1).Height
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + ((curBlue / 255) * widthCheck) - arrowOffset, picSampleRGB(2).Top + picSampleRGB(2).Height
    
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + (curHue * widthCheck) - arrowOffset, picSampleHSV(0).Top + picSampleHSV(0).Height
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + (curSaturation * widthCheck) - arrowOffset, picSampleHSV(1).Top + picSampleHSV(1).Height
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + (curValue * widthCheck) - arrowOffset, picSampleHSV(2).Top + picSampleHSV(2).Height
    
    'Next, we need to prep all our color bar DIBs
    RenderSampleDIB sRed, ccRed
    RenderSampleDIB sGreen, ccGreen
    RenderSampleDIB sBlue, ccBlue
    
    RenderSampleDIB sHue, ccHue
    RenderSampleDIB sSaturation, ccSaturation
    RenderSampleDIB sValue, ccValue
    
    'Now we can render the bars to screen
    sRed.RenderToPictureBox picSampleRGB(0)
    sGreen.RenderToPictureBox picSampleRGB(1)
    sBlue.RenderToPictureBox picSampleRGB(2)
    
    sHue.RenderToPictureBox picSampleHSV(0)
    sSaturation.RenderToPictureBox picSampleHSV(1)
    sValue.RenderToPictureBox picSampleHSV(2)
    
    'Update the hex representation box
    If (Not m_suspendHexInput) Then txtHex = Colors.GetHexStringFromRGB(RGB(curRed, curGreen, curBlue))
    
    'Re-enable syncing
    m_suspendTextResync = False
    
End Sub

'When the user clicks the hue box (or moves with the mouse button down), this function is called.  It uses the y-value
' of the click to determine new image colors, then refreshes the interface.
Private Sub HueBoxClicked(ByVal clickY As Long)

    'Restrict mouse clicks to the picture box area
    If (clickY < 0) Then clickY = 0
    If (clickY > picHue.ScaleHeight) Then clickY = picHue.ScaleHeight

    'Calculate a new hue using the mouse's y-position as our guide
    curHue = clickY / picHue.ScaleHeight
    TrimHSV curHue
    
    'Rebuild our RGB variables to match
    HSVtoRGB curHue, curSaturation, curValue, curRed, curGreen, curBlue
    
    'Redraw any necessary interface elements
    SyncInterfaceToCurrentColor

End Sub

'When the user clicks the primary box (or moves with the mouse button down), this function is called.  It uses the coordinates
' of the click to determine new image colors, then refreshes the interface.
Private Sub PrimaryBoxClicked(ByVal clickX As Long, ByVal clickY As Long)

    'Calculate a new value using the mouse's x-position as our guide
    curValue = clickX / picColor.ScaleWidth
    TrimHSV curValue
    
    'Calculate a new saturation using the mouse's y-position as our guide
    curSaturation = clickY / picColor.ScaleHeight
    TrimHSV curSaturation
    curSaturation = 1 - curSaturation
    
    'Rebuild our RGB variables to match
    HSVtoRGB curHue, curSaturation, curValue, curRed, curGreen, curBlue
    
    'Redraw any necessary interface elements
    SyncInterfaceToCurrentColor

End Sub

Private Sub picColor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    PrimaryBoxClicked x, y
End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then PrimaryBoxClicked x, y
End Sub

Private Sub picHue_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HueBoxClicked y
End Sub

Private Sub picHue_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then HueBoxClicked y
End Sub

Private Sub TrimHSV(ByRef hsvValue As Double)
    If hsvValue > 1 Then hsvValue = 1
    If hsvValue < 0 Then hsvValue = 0
End Sub

'This sub handles the preparation of the individual color sample boxes (one each for R/G/B/H/S/V)
' (Because we want these boxes to be color-managed, we must create them as DIBs.)
Private Sub RenderSampleDIB(ByRef dstDIB As pdDIB, ByVal dibColorType As colorCheckType)

    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    If (dstDIB.GetDIBWidth <> picSampleRGB(0).ScaleWidth) Or (dstDIB.GetDIBHeight <> picSampleRGB(0).ScaleHeight) Then
        dstDIB.CreateBlank picSampleRGB(0).ScaleWidth, picSampleRGB(0).ScaleHeight
    End If
    
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, v As Double
    
    'Initialize each component to its default type; only one parameter will be changed per dibColorType
    r = curRed
    g = curGreen
    b = curBlue
    h = curHue
    s = curSaturation
    v = curValue
    
    Dim gradientValue As Double, gradientMax As Double
    gradientMax = dstDIB.GetDIBWidth
    
    'Simple gradient-ish code implementation of drawing any individual color component
    Dim x As Long
    For x = 0 To dstDIB.GetDIBWidth
    
        gradientValue = x / gradientMax
    
        'We handle RGB separately from HSV
        If dibColorType <= ccBlue Then
            
            Select Case dibColorType
            
                Case ccRed
                    r = gradientValue * 255
                    
                Case ccGreen
                    g = gradientValue * 255
                    
                Case Else
                    b = gradientValue * 255
                    
            End Select
            
        Else
        
            Select Case dibColorType
            
                Case ccHue
                    h = gradientValue
                
                Case ccSaturation
                    s = gradientValue
                
                Case ccValue
                    v = gradientValue
            
            End Select
            
            HSVtoRGB h, s, v, r, g, b
        
        End If
        
        'Draw the color
        GDI.DrawLineToDC dstDIB.GetDIBDC, x, 0, x, dstDIB.GetDIBHeight, RGB(r, g, b)
        
    Next x
    
End Sub

Private Sub picOriginal_Click()

    'Update the current color values with the color of this box
    curRed = Colors.ExtractRed(oldColor)
    curGreen = Colors.ExtractGreen(oldColor)
    curBlue = Colors.ExtractBlue(oldColor)
    
    'Calculate new HSV values to match
    RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
    
    'Resync the interface to match the new value!
    SyncInterfaceToCurrentColor

End Sub

'When a recent color is clicked, update the screen with the new color
Private Sub picRecColor_Click(Index As Integer)

    'Update the current color values with the color of this box
    curRed = Colors.ExtractRed(recentColors(Index))
    curGreen = Colors.ExtractGreen(recentColors(Index))
    curBlue = Colors.ExtractBlue(recentColors(Index))
    
    'Calculate new HSV values to match
    RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
    
    'Resync the interface to match the new value!
    SyncInterfaceToCurrentColor

End Sub

Private Sub picSampleHSV_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    HSVBoxClicked Index, x
End Sub

Private Sub picSampleHSV_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then HSVBoxClicked Index, x
End Sub

'Whenever one of the HSV sample boxes is clicked, this function is called; it calculates new RGB/HSV values, then redraws the interface
Private Sub HSVBoxClicked(ByVal boxIndex As Long, ByVal xPos As Long)

    Dim boxWidth As Long
    boxWidth = picSampleRGB(0).ScaleWidth
    
    'Restrict mouse clicks to the picture box area
    If xPos < 0 Then xPos = 0
    If xPos > boxWidth Then xPos = boxWidth

    Select Case (boxIndex + 3)
    
        Case ccHue
            curHue = (xPos / boxWidth)
        
        Case ccSaturation
            curSaturation = (xPos / boxWidth)
        
        Case ccValue
            curValue = (xPos / boxWidth)
    
    End Select
    
    'Recalculate RGB based on the new HSV values
    HSVtoRGB curHue, curSaturation, curValue, curRed, curGreen, curBlue
    
    'Redraw the interface
    SyncInterfaceToCurrentColor

End Sub

Private Sub picSampleRGB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    RGBBoxClicked Index, x
End Sub

Private Sub picSampleRGB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then RGBBoxClicked Index, x
End Sub

'Whenever one of the RGB sample boxes is clicked, this function is called; it calculates new RGB/HSV values, then redraws the interface
Private Sub RGBBoxClicked(ByVal boxIndex As Long, ByVal xPos As Long)

    Dim boxWidth As Long
    boxWidth = picSampleRGB(0).ScaleWidth
    
    'Restrict mouse clicks to the picture box area
    If xPos < 0 Then xPos = 0
    If xPos > boxWidth Then xPos = boxWidth

    Select Case boxIndex
    
        Case ccRed
            curRed = (xPos / boxWidth) * 255
        
        Case ccGreen
            curGreen = (xPos / boxWidth) * 255
        
        Case ccBlue
            curBlue = (xPos / boxWidth) * 255
    
    End Select
    
    'Recalculate HSV based on the new RGB values
    RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
    
    'Redraw the interface
    SyncInterfaceToCurrentColor

End Sub

'Whenever a text box value is changed, sync only the relevant value, then redraw the interface
Private Sub tudHSV_Change(Index As Integer)

    If Not m_suspendTextResync Then

        Select Case (Index + 3)
        
            Case ccHue
                If tudHSV(Index).IsValid Then curHue = tudHSV(Index) / 359
            
            Case ccSaturation
                If tudHSV(Index).IsValid Then curSaturation = tudHSV(Index) / 100
            
            Case ccValue
                If tudHSV(Index).IsValid Then curValue = tudHSV(Index) / 100
        
        End Select
        
        'Recalculate RGB based on the new HSV values
        HSVtoRGB curHue, curSaturation, curValue, curRed, curGreen, curBlue
        
        'Redraw the interface
        SyncInterfaceToCurrentColor
        
    End If

End Sub

'Whenever a text box value is changed, sync only the relevant value, then redraw the interface
Private Sub tudRGB_Change(Index As Integer)

    If Not m_suspendTextResync Then

        Select Case Index
        
            Case ccRed
                If tudRGB(Index).IsValid Then curRed = tudRGB(Index)
            
            Case ccGreen
                If tudRGB(Index).IsValid Then curGreen = tudRGB(Index)
        
            Case ccBlue
                If tudRGB(Index).IsValid Then curBlue = tudRGB(Index)
        
        End Select
        
        'Recalculate HSV values based on the new RGB values
        RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
        
        'Redraw the interface
        SyncInterfaceToCurrentColor
        
    End If

End Sub

'Full validation of hex input happens in its LostFocus event, but we also do a quick-and-dirty sync during change events
Private Sub txtHex_Change()
    
    If m_suspendHexInput Or m_suspendTextResync Then Exit Sub
    
    m_suspendHexInput = True
    
    Dim newText As String
    newText = txtHex.Text
    
    'If the hex input looks valid, update the colors to match; otherwise, ignore the text completely
    If DoesHexLookValid(newText) Then
        
        'Parse the string to calculate actual numeric values; we can use VB's Val() function for this!
        curRed = Val("&H" & Left$(newText, 2))
        curGreen = Val("&H" & Mid$(newText, 3, 2))
        curBlue = Val("&H" & Right$(newText, 2))
        
        'Calculate new HSV values to match
        RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
        
        'Resync the interface to match the new value!
        SyncInterfaceToCurrentColor
        
    End If
    
    m_suspendHexInput = False

End Sub

Private Sub txtHex_LostFocusAPI()
    
    m_suspendHexInput = True
    
    Dim newText As String
    newText = txtHex.Text
    
    'If the hex input looks valid, update the colors to match; otherwise, ignore the text completely
    If DoesHexLookValid(newText) Then
        
        'Change the text box to match our properly formatted string
        txtHex.Text = newText
        
        'Parse the string to calculate actual numeric values; we can use VB's Val() function for this!
        curRed = Val("&H" & Left$(newText, 2))
        curGreen = Val("&H" & Mid$(newText, 3, 2))
        curBlue = Val("&H" & Right$(newText, 2))
        
        'Calculate new HSV values to match
        RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
        
        'Resync the interface to match the new value!
        SyncInterfaceToCurrentColor
        
    Else
        txtHex.Text = Colors.GetHexStringFromRGB(RGB(curRed, curGreen, curBlue))
    End If
    
    m_suspendHexInput = False
    
End Sub

'This function *may modify the incoming string* so please review the comments thoroughly
Private Function DoesHexLookValid(ByRef hexStringToCheck As String) As Boolean

    'Before doing anything else, remove all invalid characters from the text box
    Dim validChars As String
    validChars = "0123456789abcdef"
    
    Dim curText As String
    curText = Trim$(hexStringToCheck)
    
    Dim newText As String, curChar As String
    
    Dim i As Long
    For i = 1 To Len(curText)
        curChar = Mid$(curText, i, 1)
        If InStr(1, validChars, curChar, vbTextCompare) > 0 Then newText = newText & curChar
    Next i
        
    newText = LCase(newText)
    
    'Make sure the length is 1, 3, or 6.  Each case is handled specially; other lengths are not valid
    Select Case Len(newText)
    
        'One character is treated as a shade of gray; extend it to six characters.  (I don't know if this is actually
        ' valid CSS, but it doesn't hurt to support it... right?)
        Case 1
            newText = String$(6, newText)
            DoesHexLookValid = True
        
        'Three characters is standard shorthand hex; expand each character as a pair
        Case 3
            newText = Left$(newText, 1) & Left$(newText, 1) & Mid$(newText, 2, 1) & Mid$(newText, 2, 1) & Right$(newText, 1) & Right$(newText, 1)
            DoesHexLookValid = True
            
        'Six characters is already valid, so no need to screw with it further.
        Case 6
            DoesHexLookValid = True
        
        Case Else
            'We can't handle this character string, so reset it
            newText = Colors.GetHexStringFromRGB(RGB(curRed, curGreen, curBlue))
            DoesHexLookValid = False
    
    End Select
    
    If DoesHexLookValid Then hexStringToCheck = newText

End Function

'All events subclassed by this window are processed here.
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************

    Select Case uMsg
    
        'While the mouse is moving, update the color on-screen so the user can see it.
        Case WM_MOUSEMOVE
            If screenCaptureActive Then
                UpdateHoveredColor
                bHandled = True
            End If
        
        'If the user clicks with the left button, select that as the new color.
        Case WM_LBUTTONDOWN
            If screenCaptureActive Then
                UpdateHoveredColor
                ToggleCaptureMode False
            End If
        
        'As a failsafe, if we somehow detect a mouse_up event, turn off toggle capture mode.
        Case WM_LBUTTONUP
            If screenCaptureActive Then
                UpdateHoveredColor
                ToggleCaptureMode False
            End If
            
    End Select
    
    'bHandled = True
    
    
' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub









