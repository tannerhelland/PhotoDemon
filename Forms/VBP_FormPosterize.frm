VERSION 5.00
Begin VB.Form FormPosterize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Posterize"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
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
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsBits 
      Height          =   255
      Left            =   240
      Max             =   7
      Min             =   1
      TabIndex        =   1
      Top             =   3360
      Value           =   7
      Width           =   4575
   End
   Begin VB.TextBox txtBits 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Text            =   "7"
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox PicPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   240
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox PicEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   2640
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   4200
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  Before                                           After"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2310
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "# of bits:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   1680
      TabIndex        =   4
      Top             =   2910
      Width           =   765
   End
End
Attribute VB_Name = "FormPosterize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Posterizing Effect Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 6/August/06
'Last update: previewing, optimization, comments, variable type changes
'
'Updated posterizing interface; it has been optimized for speed and
'  ease-of-implementation.  If only VB had bit-shift operators....
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    If EntryValid(txtBits, hsBits.Min, hsBits.Max) Then
        Me.Visible = False
        Process Posterize, hsBits.Value
        Unload Me
    Else
        AutoSelectText txtBits
    End If
End Sub

'Subroutine for reducing the representative bits in an image
Public Sub PosterizeImage(ByVal NumOfBits As Byte)
    
    'pStep is the distance between values that X number of bits allows
    Dim pStep As Double
    pStep = 255 / (2 ^ CLng(NumOfBits) - 1)
    
    'Look-up tables make this far more efficient
    Dim LookUp(0 To 255) As Byte
    For x = 0 To 255
        'Add 0.5 so that values are rounded, not truncated (slightly better results)
        LookUp(x) = CByte(Int(Int(CDbl(x) / pStep + 0.5) * pStep))
    Next x
    
    'Now it's easy - loop, and change pixels using the look-up table
    Message "Generating poster data..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        ImageData(QuickVal + 2, y) = LookUp(ImageData(QuickVal + 2, y))
        ImageData(QuickVal + 1, y) = LookUp(ImageData(QuickVal + 1, y))
        ImageData(QuickVal, y) = LookUp(ImageData(QuickVal, y))
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Create the previews
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    PreviewPosterize hsBits.Value
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'Same as above, but designed exclusively for previewing
Private Sub PreviewPosterize(ByVal NumOfBits As Byte)
    GetPreviewData PicPreview
    Dim pStep As Double
    pStep = 255 / (2 ^ CLng(NumOfBits) - 1)
    Dim LookUp(0 To 255) As Byte
    For x = 0 To 255
        LookUp(x) = CByte(Int(Int(CDbl(x) / pStep + 0.5) * pStep))
    Next x
    Dim QuickVal As Long
    For x = PreviewX To PreviewX + PreviewWidth
        QuickVal = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        ImageData(QuickVal + 2, y) = LookUp(ImageData(QuickVal + 2, y))
        ImageData(QuickVal + 1, y) = LookUp(ImageData(QuickVal + 1, y))
        ImageData(QuickVal, y) = LookUp(ImageData(QuickVal, y))
    Next y
    Next x
    SetPreviewData PicEffect
End Sub

'The following routines are for keeping the text box and scroll bar values in lock-step
Private Sub hsBits_Change()
    txtBits.Text = hsBits.Value
    PreviewPosterize hsBits.Value
End Sub

Private Sub hsBits_Scroll()
    txtBits.Text = hsBits.Value
    PreviewPosterize hsBits.Value
End Sub

Private Sub txtBits_Change()
    If EntryValid(txtBits, hsBits.Min, hsBits.Max, False, False) Then hsBits.Value = val(txtBits)
End Sub

Private Sub txtBits_GotFocus()
    AutoSelectText txtBits
End Sub
