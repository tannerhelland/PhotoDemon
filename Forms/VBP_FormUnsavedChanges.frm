VERSION 5.00
Begin VB.Form FormUnsavedChanges 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Unsaved Changes"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9360
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
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRepeat 
      Appearance      =   0  'Flat
      Caption         =   " Repeat this action for all unsaved images (X in total)"
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   3960
      TabIndex        =   5
      Top             =   4020
      Width           =   5175
   End
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1260
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
      BackColor       =   15790320
      Caption         =   "Save the image before closing it"
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
      Left            =   3960
      TabIndex        =   3
      Top             =   2040
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
      BackColor       =   15790320
      Caption         =   "Do not save the image (discard all changes)"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   2820
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
      BackColor       =   15790320
      Caption         =   "Cancel, and return to editing"
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormUnsavedChanges.frx":20A4
      PictureAlign    =   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Cancel"
   End
   Begin VB.Line lineBottom 
      BorderColor     =   &H8000000D&
      X1              =   8
      X2              =   616
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "This image (filename.jpg) has unsaved changes.  What would you like to do?"
      ForeColor       =   &H00202020&
      Height          =   765
      Left            =   4830
      TabIndex        =   1
      Top             =   360
      Width           =   4215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormUnsavedChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Unsaved Changes Dialog
'Copyright ©2011-2012 by Tanner Helland
'Created: 13/November/12
'Last updated: 14/November/12
'Last update: added a system "warning" icon to the dialog box.  This is drawn automatically.
'
'Custom dialog box for warning the user that they are about to close an image with unsaved changes.
'
'This form was built after much usability testing.  There are many bad ways to design a save prompt,
' and only a few good ones.  I felt that descriptive icons were necessary to help the user quickly
' determine what choice to make.  A preview of the image in question is also displayed, to make it
' absolutely certain that the user is not confused about which image they're dealing with.  (This is
' important for photos from a digital camera, which often have names like "1004701.jpg". Very
' descriptive tooltip text has also been added, and I genuinely believe that this is one of the best
' unsaved changes dialogs available.
'
'Finally, note that this prompt can be turned off completely from the Edit -> Preferences menu.
'
'***************************************************************************


Option Explicit

'The following Enum and two API declarations are used to draw the system information icon
Enum SystemIconConstants
    IDI_APPLICATION = 32512
    IDI_HAND = 32513
    IDI_QUESTION = 32514
    IDI_EXCLAMATION = 32515
    IDI_ASTERISK = 32516
    IDI_WINDOWS = 32517
End Enum

Private Declare Function LoadIconByID Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

'The ID number of the image being closed
Private imageBeingClosed As Long

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'Draw a system icon on the specified device context; this code is adopted from an example by Francesco Balena at http://www.devx.com/vb2themax/Tip/19108
Private Sub DrawSystemIcon(ByVal icon As SystemIconConstants, ByVal hDC As Long, ByVal x As Long, ByVal y As Long)
    Dim hIcon As Long
    hIcon = LoadIconByID(0, icon)
    DrawIcon hDC, x, y, hIcon
End Sub

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

Public Property Let formID(formID As Long)
    imageBeingClosed = formID
End Property

'The three choices available to the user correspond to message box responses of "Yes", "No", and "Cancel"
Private Sub CmdCancel_Click()
    userAnswer = vbCancel
    updateRepeatToAllUnsavedImages userAnswer
    Me.Hide
End Sub

Private Sub cmdDontSave_Click()
    userAnswer = vbNo
    updateRepeatToAllUnsavedImages userAnswer
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    userAnswer = vbYes
    updateRepeatToAllUnsavedImages userAnswer
    Me.Hide
End Sub

'The ShowDialog routine presents the user with the form.  FormID MUST BE SET in advance of calling this.
Public Sub ShowDialog()
    
    'Automatically draw a warning icon using the system icon set
    Dim iconY As Long
    iconY = 24
    If useFancyFonts Then iconY = iconY + 2
    DrawSystemIcon IDI_EXCLAMATION, Me.hDC, 277, iconY
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Draw the image being closed to the preview box
    pdImages(imageBeingClosed).mainLayer.renderToPictureBox picPreview
    
    'Adjust the save message to match this image's name
    lblWarning.Caption = pdImages(imageBeingClosed).OriginalFileNameAndExtension & " has unsaved changes.  What would you like to do?"

    'If the image has been saved before, update the tooltip text on the "Save" button accordingly
    If pdImages(imageBeingClosed).LocationOnDisk <> "" Then
        cmdSave.ToolTip = vbCrLf & "NOTE: if you click 'Save', PhotoDemon will save this image using its current file name." & vbCrLf & vbCrLf & "If you want to save it with a different file name, please select 'Cancel', then use the" & vbCrLf & " File -> Save As menu item."
    Else
        cmdSave.ToolTip = vbCrLf & "Because this image has not been saved before, you will be presented with a full Save As dialog."
    End If
    
    'Update the other tooltip buttons as well
    cmdDontSave.ToolTip = vbCrLf & "If you do not save this image, any changes you have made will be permanently lost."
    cmdCancel.ToolTip = vbCrLf & "Canceling will return you to the main PhotoDemon window."

    'Make some measurements of the form size.  We need these if we choose to display the check box at the bottom of the form
    Dim vDifference As Long
    Me.ScaleMode = vbTwips
    vDifference = Me.Height - Me.ScaleHeight
    
    'If there are multiple unsaved images, give the user a prompt to apply this action to all of them.
    ' (If there are not multiple unsaved images, hide that section from view.)
    If numOfUnsavedImages < 2 Then
        lineBottom.Visible = False
        chkRepeat.Visible = False
        Me.Height = vDifference + picPreview.Height + (picPreview.Top * 2)
    Else
        lineBottom.Visible = True
        chkRepeat.Visible = True
        
        'Change the text of the "repeat for all unsaved images" check box depending on how many unsaved images are present.
        If numOfUnsavedImages = 2 Then
            chkRepeat.Caption = " Repeat this action for both unsaved images"
        Else
            chkRepeat.Caption = " Repeat this action for all unsaved images (" & numOfUnsavedImages & " in total)"
        End If
        
        Me.Height = vDifference + (chkRepeat.Top + chkRepeat.Height) + picPreview.Top
    End If

    Me.ScaleMode = vbPixels

    'Apply any custom styles to the form
    makeFormPretty Me

    'Display the form
    Me.Show vbModal, FormMain

End Sub

'Before this dialog closes, this routine is called to update the user's preference for applying this action to all unsaved images
Private Sub updateRepeatToAllUnsavedImages(ByVal actionToApply As VbMsgBoxResult)
    
    If chkRepeat.Value = vbChecked Then
        dealWithAllUnsavedImages = True
        howToDealWithAllUnsavedImages = actionToApply
    Else
        dealWithAllUnsavedImages = False
    End If
    
End Sub

