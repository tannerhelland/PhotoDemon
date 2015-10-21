VERSION 5.00
Begin VB.Form dialog_UnsavedChanges 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Unsaved Changes"
   ClientHeight    =   4500
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
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   0
      Left            =   3960
      TabIndex        =   0
      Top             =   1260
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   1296
      Caption         =   "Save the image before closing it"
   End
   Begin PhotoDemon.smartCheckBox chkRepeat 
      Height          =   330
      Left            =   3960
      TabIndex        =   3
      Top             =   4005
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   582
      Caption         =   "Repeat this action for all unsaved images (X in total)"
      Value           =   0
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   1
      Left            =   3960
      TabIndex        =   1
      Top             =   2070
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   1296
      Caption         =   "Do not save the image (discard all changes)"
   End
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   2880
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   1296
      Caption         =   "Cancel, and return to editing"
   End
   Begin VB.Line lineBottom 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   624
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "%1 has unsaved changes.  What would you like to do?"
      ForeColor       =   &H00202020&
      Height          =   765
      Left            =   4830
      TabIndex        =   4
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
'Copyright 2011-2015 by Tanner Helland
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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
Public Sub showDialog(ByRef ownerForm As Form)
    
    Dim i As Long
    
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    cmdAnswer(0).AssignImage "LRGSAVE"
    cmdAnswer(1).AssignImage "LRGDONTSAVE"
    cmdAnswer(2).AssignImage "LRGUNDO"
        
    'If the image has been saved before, update the tooltip text on the "Save" button accordingly
    If Len(pdImages(imageBeingClosed).locationOnDisk) <> 0 Then
        cmdAnswer(0).AssignTooltip "NOTE: if you click 'Save', PhotoDemon will save this image using its current file name." & vbCrLf & vbCrLf & "If you want to save it with a different file name, please select 'Cancel', then use the File -> Save As menu item."
    Else
        cmdAnswer(0).AssignTooltip "Because this image has not been saved before, you will be prompted to provide a file name for it."
    End If
    
    cmdAnswer(1).AssignTooltip "If you do not save this image, any changes you have made will be permanently lost."
    cmdAnswer(2).AssignTooltip "Canceling will return you to the main PhotoDemon window."
    
    'Automatically draw a warning icon using the system icon set
    Dim iconY As Long
    iconY = FixDPI(24)
    If g_UseFancyFonts Then iconY = iconY + FixDPI(2)
    DrawSystemIcon IDI_EXCLAMATION, Me.hDC, FixDPI(277), iconY
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
        
    'Adjust the save message to match this image's name
    lblWarning.Caption = g_Language.TranslateMessage("%1 has unsaved changes.  What would you like to do?", pdImages(imageBeingClosed).originalFileNameAndExtension)
    
    'Make some measurements of the form size.  We need these if we choose to display the check box at the bottom of the form
    Dim vDifference As Long
    Me.ScaleMode = vbTwips
    vDifference = Me.Height - Me.ScaleHeight
    
    'If there are multiple unsaved images, give the user a prompt to apply this action to all of them.
    ' (If there are not multiple unsaved images, hide that section from view.)
    If g_NumOfUnsavedImages < 2 Then
        lineBottom.Visible = False
        chkRepeat.Visible = False
        Me.Height = vDifference + picPreview.Height + (picPreview.Top * 2)
    Else
        lineBottom.Visible = True
        chkRepeat.Visible = True
        
        'Change the text of the "repeat for all unsaved images" check box depending on how many unsaved images are present.
        If g_NumOfUnsavedImages = 2 Then
            chkRepeat.Caption = g_Language.TranslateMessage(" Repeat this action for both unsaved images")
        Else
            chkRepeat.Caption = g_Language.TranslateMessage(" Repeat this action for all unsaved images (%1 in total)", g_NumOfUnsavedImages)
        End If
        
        Me.Height = vDifference + (chkRepeat.Top + chkRepeat.Height) + picPreview.Top
    End If

    Me.ScaleMode = vbPixels
    
    'When translations are active, some lengthy language may push the check box caption completely off-screen.
    ' To prevent this, give the check box a large buffer space if translations are active.
    If g_Language.translationActive Then
        chkRepeat.Left = FixDPI(8)
        chkRepeat.Width = Me.ScaleWidth - FixDPI(16)
    End If
    
    'Apply any custom styles to the form
    MakeFormPretty Me
        
    'Display the form
    ShowPDDialog vbModal, Me, True

End Sub

'Before this dialog closes, this routine is called to update the user's preference for applying this action to all unsaved images
Private Sub updateRepeatToAllUnsavedImages(ByVal actionToApply As VbMsgBoxResult)
    
    If chkRepeat.Visible And chkRepeat.Value = vbChecked Then
        g_DealWithAllUnsavedImages = True
        g_HowToDealWithAllUnsavedImages = actionToApply
    Else
        g_DealWithAllUnsavedImages = False
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

Private Sub Form_Activate()

    'Draw the image being closed to the preview box
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    If (Not pdImages(imageBeingClosed) Is Nothing) Then
        pdImages(imageBeingClosed).requestThumbnail tmpDIB, IIf(picPreview.ScaleWidth > picPreview.ScaleHeight, picPreview.ScaleHeight, picPreview.ScaleWidth)
    End If
    
    If (Not pdImages(imageBeingClosed) Is Nothing) And (Not tmpDIB Is Nothing) Then
        tmpDIB.renderToPictureBox picPreview
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
