VERSION 5.00
Begin VB.Form dialog_UnsavedChanges 
   Appearance      =   0  'Flat
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
      BackColor       =   &H00FFFFFF&
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
   Begin PhotoDemon.jcbutton cmdAnswer 
      Default         =   -1  'True
      Height          =   735
      Index           =   0
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
   Begin PhotoDemon.jcbutton cmdAnswer 
      Height          =   735
      Index           =   1
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
      Caption         =   "Do not save the image (discard all changes)"
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormUnsavedChanges.frx":1052
      PictureAlign    =   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Do Not Save"
   End
   Begin PhotoDemon.jcbutton cmdAnswer 
      Cancel          =   -1  'True
      Height          =   735
      Index           =   2
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
Attribute VB_Name = "dialog_UnsavedChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Unsaved Changes Dialog
'Copyright ©2011-2013 by Tanner Helland
'Created: 13/November/12
'Last updated: 01/December/12
'Last update: removed the DrawSystemIcon sub; now it can be found in the "Drawing" module
'
'Custom dialog box for warning the user that they are about to close an image with unsaved changes.
'
'This form was built after much usability testing.  There are many bad ways to design a save prompt,
' and only a few good ones.  I felt that descriptive icons were necessary to help the user quickly
' determine what choice to make.  A preview of the image in question is also displayed, to make it
' absolutely certain that the user is not confused about which image they're dealing with.  (This is
' important for photos from a digital camera, which often have names like "1004701.jpg".) Very
' descriptive tooltip text has also been added, and I genuinely believe that this is one of the best
' unsaved changes dialogs available.
'
'Finally, note that this prompt can be turned off completely from the Edit -> Preferences menu.
'
'***************************************************************************


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
    If pdImages(imageBeingClosed).mainLayer.getLayerColorDepth = 24 Then
        pdImages(imageBeingClosed).mainLayer.renderToPictureBox picPreview
    Else
        Dim tmpLayer As pdLayer
        Set tmpLayer = New pdLayer
        Dim nWidth As Long, nHeight As Long
        convertAspectRatio pdImages(imageBeingClosed).Width, pdImages(imageBeingClosed).Height, picPreview.ScaleWidth, picPreview.ScaleHeight, nWidth, nHeight
        tmpLayer.createFromExistingLayer pdImages(imageBeingClosed).mainLayer, nWidth, nHeight, True
        tmpLayer.compositeBackgroundColor
        tmpLayer.renderToPictureBox picPreview
        tmpLayer.eraseLayer
        Set tmpLayer = Nothing
    End If
    
    'Adjust the save message to match this image's name
    lblWarning.Caption = pdImages(imageBeingClosed).OriginalFileNameAndExtension & " has unsaved changes.  What would you like to do?"

    'If the image has been saved before, update the tooltip text on the "Save" button accordingly
    If pdImages(imageBeingClosed).LocationOnDisk <> "" Then
        cmdAnswer(0).ToolTip = vbCrLf & "NOTE: if you click 'Save', PhotoDemon will save this image using its current file name." & vbCrLf & vbCrLf & "If you want to save it with a different file name, please select 'Cancel', then use the" & vbCrLf & " File -> Save As menu item."
    Else
        cmdAnswer(0).ToolTip = vbCrLf & "Because this image has not been saved before, you will be presented with a full Save As dialog."
    End If
    
    'Update the other tooltip buttons as well
    cmdAnswer(1).ToolTip = vbCrLf & "If you do not save this image, any changes you have made will be permanently lost."
    cmdAnswer(2).ToolTip = vbCrLf & "Canceling will return you to the main PhotoDemon window."

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

'The three choices available to the user correspond to message box responses of "Yes", "No", and "Cancel"
Private Sub cmdAnswer_Click(Index As Integer)

    Select Case Index
    
        Case 0
            userAnswer = vbYes
        
        Case 1
            userAnswer = vbNo
            
        Case 2
            userAnswer = vbCancel
        
    End Select
    
    updateRepeatToAllUnsavedImages userAnswer
    Me.Hide

End Sub
