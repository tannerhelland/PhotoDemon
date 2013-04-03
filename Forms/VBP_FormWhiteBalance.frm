VERSION 5.00
Begin VB.Form FormWhiteBalance 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " White Balance"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12120
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
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9180
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10650
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin VB.HScrollBar hsIgnore 
      Height          =   255
      Left            =   6120
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   2880
      Value           =   5
      Width           =   4815
   End
   Begin VB.TextBox txtIgnore 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   11040
      TabIndex        =   3
      Text            =   "0.05"
      Top             =   2835
      Width           =   735
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -120
      TabIndex        =   5
      Top             =   5760
      Width           =   12495
   End
   Begin VB.Label lblAmount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "strength:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   6000
      TabIndex        =   4
      Top             =   2520
      Width           =   960
   End
End
Attribute VB_Name = "FormWhiteBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'White Balance Handler
'Copyright ©2000-2013 by Tanner Helland
'Created: 03/July/12
'Last updated: 03/July/12
'Last update: first build
'
'White balance handler.  Unlike other programs, which shove this under the Levels dialog as an "auto levels"
' function, I consider it worthy of its own interface.  The reason is - white balance is an important function.
' It's arguably more useful than the Levels dialog, especially to a casual user, because it automatically
' calculates levels according to a reliable, often-accurate algorithm.  Rather than forcing the user through the
' Levels dialog (because really, how many people know that Auto Levels is actually White Balance in photography
' parlance?), PhotoDemon provides a full implementation of custom white balance handling.
' The value box on the form is the percentage of pixels ignored at the top and bottom of the histogram.
' 0.05 is the recommended default.  I've specified 1.5 as the maximum, but there's no reason it couldn't be set
' higher... just be forewarned that higher values (obviously) blow out the picture with increasing strength.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    'The scroll bar max and min values are used to check the gamma input for validity
    If EntryValid(txtIgnore, hsIgnore.Min / 100, hsIgnore.Max / 100) Then
        Me.Visible = False
        Process WhiteBalance, CSng(Val(txtIgnore))
        Unload Me
    Else
        AutoSelectText txtIgnore
    End If
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Render a preview
    AutoWhiteBalance CSng(Val(txtIgnore)), True, fxPreview
    
End Sub

'Correct white balance by stretching the histogram and ignoring pixels above or below the 0.05% threshold
Public Sub AutoWhiteBalance(Optional ByVal percentIgnore As Double = 0.05, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Adjusting image white balance..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    WhiteBalanceLayer percentIgnore, workingLayer, toPreview
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When the horizontal scroll bar is moved, change the text box to match
Private Sub hsIgnore_Change()
    copyToTextBoxF CSng(hsIgnore) / 100, txtIgnore
    AutoWhiteBalance CSng(Val(txtIgnore)), True, fxPreview
End Sub

Private Sub hsIgnore_Scroll()
    copyToTextBoxF CSng(hsIgnore) / 100, txtIgnore
    AutoWhiteBalance CSng(Val(txtIgnore)), True, fxPreview
End Sub

Private Sub txtIgnore_GotFocus()
    AutoSelectText txtIgnore
End Sub

'If the user changes the gamma value by hand, check it for numerical correctness, then change the horizontal scroll bar to match
Private Sub txtIgnore_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtIgnore, , True
    If EntryValid(txtIgnore, hsIgnore.Min / 100, hsIgnore.Max / 100, False, False) Then hsIgnore.Value = Val(txtIgnore) * 100
End Sub

