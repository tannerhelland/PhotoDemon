VERSION 5.00
Begin VB.Form dialog_MultiImage 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Multiple Images Found"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5595
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
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.smartCheckBox chkRepeat 
      Height          =   480
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   847
      Caption         =   "always apply this action to multi-image files"
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
   Begin PhotoDemon.jcbutton cmdAnswer 
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   1296
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Load each page as its own image"
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormMultiImage.frx":0000
      PictureAlign    =   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      ToolTip         =   "This will open all images in this file."
      TooltipType     =   1
      TooltipTitle    =   "Load All Images"
   End
   Begin PhotoDemon.jcbutton cmdAnswer 
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2100
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   1296
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Load only the first page"
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormMultiImage.frx":1052
      PictureAlign    =   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Load One Image Only"
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "%1 contains multiple pages (%2 in total).  How would you like to proceed?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   765
      Left            =   960
      TabIndex        =   0
      Top             =   270
      Width           =   4335
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "dialog_MultiImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Multi-Image Load Dialog
'Copyright ©2011-2013 by Tanner Helland
'Created: 01/December/12
'Last updated: 27/December/12
'Last update: add support for icon files, which may contain many embedded icons.
'
'Custom dialog box for asking the user how they want to treat a multi-image file (at present, an
' animated GIF, multipage TIFF, or ICO).
'
'This form is tied into the settable user preference for handling multipage images.  Checking the
' "remember this decision and don't ask me again" option will set that preference for the user.
' Note that this setting can also be changed from the Edit -> Preferences menu.
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'We want to temporarily suspend an hourglass cursor if necessary
Private restoreCursor As Boolean

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The ShowDialog routine presents the user with the form.  FormID MUST BE SET in advance of calling this.
Public Sub ShowDialog(ByVal srcFilename As String, ByVal numOfImages As Long)

    If Screen.MousePointer = vbHourglass Then
        restoreCursor = True
        Screen.MousePointer = vbNormal
    Else
        restoreCursor = False
    End If

    'Automatically draw a warning icon using the system icon set
    Dim iconY As Long
    iconY = 18
    If g_UseFancyFonts Then iconY = iconY + 2
    DrawSystemIcon IDI_QUESTION, Me.hDC, 22, iconY
    
    'Provide a default answer of "first image only" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbNo

    'Adjust the prompt to match this file's name and page count
    Dim FileExtension As String
    FileExtension = GetExtension(srcFilename)
    If UCase(FileExtension) = "GIF" Then
        lblWarning.Caption = g_Language.TranslateMessage("%1 is an animated GIF file (%2 frames total).  How would you like to proceed?", getFilename(srcFilename), numOfImages)
        cmdAnswer(0).Caption = g_Language.TranslateMessage("Load each frame as a separate image")
        cmdAnswer(0).ToolTip = g_Language.TranslateMessage("This option will load every frame in the animated GIF file as an individual image.")
        cmdAnswer(0).TooltipTitle = g_Language.TranslateMessage("Load all frames")
        cmdAnswer(1).Caption = g_Language.TranslateMessage("Load only the first frame")
        cmdAnswer(1).ToolTip = g_Language.TranslateMessage("This option will only load a single frame from the animated GIF file," & vbCrLf & "effectively treating at as a non-animated GIF file.")
        cmdAnswer(1).TooltipTitle = g_Language.TranslateMessage("Load one frame only")
    ElseIf UCase(FileExtension) = "ICO" Then
        lblWarning.Caption = g_Language.TranslateMessage("%1 contains multiple icons (%2 in total).  How would you like to proceed?", getFilename(srcFilename), numOfImages)
        cmdAnswer(0).Caption = g_Language.TranslateMessage("Load each icon as a separate image")
        cmdAnswer(0).ToolTip = g_Language.TranslateMessage("This option will load every icon in the ICO file as an individual image.")
        cmdAnswer(0).TooltipTitle = g_Language.TranslateMessage("Load all icons")
        cmdAnswer(1).Caption = g_Language.TranslateMessage("Load only the first icon")
        cmdAnswer(1).ToolTip = g_Language.TranslateMessage("This option will only load a single icon from the ICO file.")
        cmdAnswer(1).TooltipTitle = g_Language.TranslateMessage("Load one icon only")
    Else
        lblWarning.Caption = g_Language.TranslateMessage("%1 contains multiple pages (%2 in total).  How would you like to proceed?", getFilename(srcFilename), numOfImages)
        cmdAnswer(0).Caption = g_Language.TranslateMessage("Load each page as a separate image")
        cmdAnswer(0).ToolTip = g_Language.TranslateMessage("This option will load every page in the TIFF file as an individual image.")
        cmdAnswer(0).TooltipTitle = g_Language.TranslateMessage("Load all pages")
        cmdAnswer(1).Caption = g_Language.TranslateMessage("Load only the first page")
        cmdAnswer(1).ToolTip = g_Language.TranslateMessage("This option will only load a single page from the TIFF file.")
        cmdAnswer(1).TooltipTitle = g_Language.TranslateMessage("Load one page only")
    End If

    'Apply any custom styles to the form
    makeFormPretty Me

    'Display the form
    Me.Show vbModal, FormMain

End Sub

'Update the dialog's return value based on the pressed command button
Private Sub cmdAnswer_Click(Index As Integer)

    Select Case Index
    
        Case 0
            userAnswer = vbYes
            If CBool(chkRepeat.Value) Then g_UserPreferences.SetPreference_Long "General Preferences", "MultipageImagePrompt", 2
            
        Case 1
            userAnswer = vbNo
            If CBool(chkRepeat.Value) Then g_UserPreferences.SetPreference_Long "General Preferences", "MultipageImagePrompt", 1
            
    End Select
        
    If restoreCursor Then Screen.MousePointer = vbHourglass
        
    Me.Hide
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
