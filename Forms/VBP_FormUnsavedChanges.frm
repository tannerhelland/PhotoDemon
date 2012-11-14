VERSION 5.00
Begin VB.Form FormUnsavedChanges 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Unsaved Changes"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   231
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin PhotoDemon.jcbutton cmdSave 
      Height          =   735
      Left            =   4080
      TabIndex        =   2
      Top             =   1080
      Width           =   5220
      _ExtentX        =   9208
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
      Caption         =   "Save the image before closing it"
      ForeColor       =   2105376
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormUnsavedChanges.frx":0000
      PictureAlign    =   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      ToolTip         =   "Open a previously saved convolution filter."
      TooltipType     =   1
      TooltipTitle    =   "Save"
   End
   Begin PhotoDemon.jcbutton cmdDontSave 
      Height          =   735
      Left            =   4080
      TabIndex        =   3
      Top             =   1920
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   1296
      ButtonStyle     =   13
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Do not save the image (discard all changes)"
      ForeColor       =   2105376
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormUnsavedChanges.frx":1052
      PictureAlign    =   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Do Not Save"
   End
   Begin PhotoDemon.jcbutton cmdCancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      Top             =   2760
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   1296
      ButtonStyle     =   13
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Cancel, and return to editing"
      ForeColor       =   2105376
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormUnsavedChanges.frx":20A4
      PictureAlign    =   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Cancel"
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "This image (filename.jpg) has unsaved changes.  What would you like to do?"
      ForeColor       =   &H00202020&
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   210
      Width           =   5175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormUnsavedChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The ID number of the image being closed
Private imageBeingClosed As Long

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

Public Property Let formID(formID As Long)
    imageBeingClosed = formID
End Property

Private Sub cmdCancel_Click()
    userAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdDontSave_Click()
    userAnswer = vbNo
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    userAnswer = vbYes
    Me.Hide
End Sub

'The ShowDialog routine presents the user with the form.  FormID MUST BE SET in advance of calling this.
Public Sub ShowDialog()
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Draw the image being closed to the preview box
    pdImages(imageBeingClosed).mainLayer.renderToPictureBox picPreview
    
    'Adjust the save message to match this image's name
    lblWarning.Caption = """" & pdImages(imageBeingClosed).OriginalFileNameAndExtension & """ has unsaved changes." & vbCrLf & "What would you like to do?"

    'If the image has been saved before, update the tooltip text on the "Save" button accordingly
    If pdImages(imageBeingClosed).LocationOnDisk <> "" Then
        cmdSave.ToolTip = vbCrLf & "NOTE: if you click 'Save', PhotoDemon will save this image using its current file name." & vbCrLf & vbCrLf & "If you want to save it with a different file name, please select 'Cancel', then use the" & vbCrLf & " File -> Save As menu item."
    Else
        cmdSave.ToolTip = vbCrLf & "Because this image has not been saved before, you will be presented with a full Save As dialog."
    End If
    
    'Update the other tooltip buttons as well
    cmdDontSave.ToolTip = vbCrLf & "If you do not save this image, any changes you have made will be permanently lost."
    cmdCancel.ToolTip = vbCrLf & "Canceling will return you to the main PhotoDemon window."

    Me.Show vbModal, FormMain

End Sub

