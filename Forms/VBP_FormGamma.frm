VERSION 5.00
Begin VB.Form FormGamma 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Gamma Correction"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsGamma 
      Height          =   255
      Left            =   240
      Max             =   200
      Min             =   1
      TabIndex        =   2
      Top             =   4080
      Value           =   100
      Width           =   4575
   End
   Begin VB.TextBox txtGamma 
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
      TabIndex        =   1
      Text            =   "1.00"
      Top             =   3570
      Width           =   495
   End
   Begin VB.PictureBox PicPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   240
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox PicEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   2640
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   4680
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   4680
      Width           =   1125
   End
   Begin VB.ComboBox CboChannel 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
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
      TabIndex        =   9
      Top             =   2310
      Width           =   4575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
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
      TabIndex        =   6
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Channel:"
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
      TabIndex        =   5
      Top             =   2940
      Width           =   705
   End
End
Attribute VB_Name = "FormGamma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gamma Correction Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 12/May/01
'Last updated: 3/March/07
'Last update: preview!
'
'Updated version of the gamma handler; fully optimized, it uses a look-up
'table and can correct any color channel.
'
'***************************************************************************

Option Explicit

'Update the preview when the user changes the channel combo box
Private Sub CboChannel_Click()
    PreviewGamma CSng(val(txtGamma)), CByte(CboChannel.ListIndex)
End Sub

Private Sub CboChannel_KeyDown(KeyCode As Integer, Shift As Integer)
    PreviewGamma CSng(val(txtGamma)), CByte(CboChannel.ListIndex)
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    'The scroll bar max and min values are used to check the gamma input for validity
    If EntryValid(txtGamma, hsGamma.Min / 100, hsGamma.Max / 100) Then
        Me.Visible = False
        Process GammaCorrection, CSng(val(txtGamma)), CByte(CboChannel.ListIndex)
        Unload Me
    End If
End Sub

'Initialize the preview boxes and the gamma combo box
Private Sub Form_Load()
    
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect

    CboChannel.AddItem "RGB", 0
    CboChannel.AddItem "Red", 1
    CboChannel.AddItem "Green", 2
    CboChannel.AddItem "Blue", 3
    CboChannel.ListIndex = 0
    DoEvents
    
    PreviewGamma CSng(val(txtGamma)), CByte(CboChannel.ListIndex)
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'Basic gamma correction.  It's a simple function - use an exponent to adjust R/G/B values.
Public Sub GammaCorrect(ByVal Gamma As Single, ByVal Method As Byte)
    
    Dim r As Long, g As Long, b As Long
    Dim LookUp(0 To 255) As Integer
    
    Dim TempVal As Single
    
    GetImageData
    
    Message "Generating gamma table..."
    
    For x = 0 To 255
        TempVal = x / 255
        TempVal = TempVal ^ (1 / Gamma)
        TempVal = TempVal * 255
        
        If TempVal > 255 Then TempVal = 255
        If TempVal < 0 Then TempVal = 0
        
        LookUp(x) = TempVal
    Next x
    
    Message "Correcting gamma values..."
    SetProgBarMax PicWidthL
    
    Dim QuickX As Long
    For x = 0 To PicWidthL
        QuickX = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickX + 2, y)
        g = ImageData(QuickX + 1, y)
        b = ImageData(QuickX, y)
        If Method = 0 Then
            r = LookUp(r)
            g = LookUp(g)
            b = LookUp(b)
        ElseIf Method = 1 Then
            r = LookUp(r)
        ElseIf Method = 2 Then
            g = LookUp(g)
        ElseIf Method = 3 Then
            b = LookUp(b)
        End If
        ImageData(QuickX + 2, y) = r
        ImageData(QuickX + 1, y) = g
        ImageData(QuickX, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetImageData
    
End Sub

'When the horizontal scroll bar is moved, change the text box to match
Private Sub hsGamma_Change()
    txtGamma.Text = Format(CSng(hsGamma.Value) / 100, "0.00")
    PreviewGamma CSng(val(txtGamma)), CByte(CboChannel.ListIndex)
End Sub

Private Sub hsGamma_Scroll()
    txtGamma.Text = Format(CSng(hsGamma.Value) / 100, "0.00")
    PreviewGamma CSng(val(txtGamma)), CByte(CboChannel.ListIndex)
End Sub

'Render a gamma preview to the preview picture box
Private Sub PreviewGamma(ByVal Gamma As Single, ByVal Method As Byte)
    Dim r As Long, g As Long, b As Long
    Dim LookUp(0 To 255) As Integer
    Dim TempVal As Single
    Dim TempLong As Long
    GetPreviewData PicPreview
    
    For x = 0 To 255
        TempVal = x / 255
        TempVal = TempVal ^ (1 / Gamma)
        TempVal = TempVal * 255
        If TempVal > 255 Then TempVal = 255
        If TempVal < 0 Then TempVal = 0
        LookUp(x) = TempVal
    Next x
    
    For x = PreviewX To PreviewX + PreviewWidth
        TempLong = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        r = ImageData(TempLong + 2, y)
        g = ImageData(TempLong + 1, y)
        b = ImageData(TempLong, y)
        If Method = 0 Then
            r = LookUp(r)
            g = LookUp(g)
            b = LookUp(b)
        ElseIf Method = 1 Then
            r = LookUp(r)
        ElseIf Method = 2 Then
            g = LookUp(g)
        ElseIf Method = 3 Then
            b = LookUp(b)
        End If
        ImageData(TempLong + 2, y) = r
        ImageData(TempLong + 1, y) = g
        ImageData(TempLong, y) = b
    Next y
    Next x
    
    SetPreviewData PicEffect
    
End Sub

Private Sub txtGamma_GotFocus()
    AutoSelectText txtGamma
End Sub

'If the user changes the gamma value by hand, check it for numerical correctness, then change the horizontal scroll bar to match
Private Sub txtGamma_Change()
    If EntryValid(txtGamma, hsGamma.Min / 100, hsGamma.Max / 100, False, False) Then hsGamma.Value = val(txtGamma) * 100
End Sub
