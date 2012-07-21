VERSION 5.00
Begin VB.Form FormNoise 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Add Noise"
   ClientHeight    =   4890
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
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsNoise 
      Height          =   255
      Left            =   240
      Max             =   500
      Min             =   1
      MouseIcon       =   "VBP_FormNoise.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3240
      Value           =   1
      Width           =   4575
   End
   Begin VB.TextBox txtNoise 
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
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Text            =   "1"
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox PicEffect 
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
      Height          =   2175
      Left            =   2640
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox PicPreview 
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
      Height          =   2175
      Left            =   240
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.CheckBox ChkM 
      Appearance      =   0  'Flat
      Caption         =   "Monochromatic"
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
      Height          =   255
      Left            =   1680
      MouseIcon       =   "VBP_FormNoise.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      MouseIcon       =   "VBP_FormNoise.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4320
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      MouseIcon       =   "VBP_FormNoise.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4320
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
      TabIndex        =   8
      Top             =   2310
      Width           =   4575
   End
   Begin VB.Label Label1 
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
      TabIndex        =   5
      Top             =   2790
      Width           =   720
   End
End
Attribute VB_Name = "FormNoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Noise Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 3/15/01
'Last updated: 6/August/06
'Last update: previewing, optimization, comments, variable type changes
'
'Form for adding noise to an image.
'
'***************************************************************************

Option Explicit

Private Sub ChkM_Click()
    PreviewNoise hsNoise.Value, ChkM.Value
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    If EntryValid(txtNoise, hsNoise.Min, hsNoise.Max) Then
        FormNoise.Visible = False
        Process Noise, hsNoise.Value, ChkM.Value
        Unload Me
    Else
        AutoSelectText txtNoise
    End If
End Sub

'Subroutine for adding noise to an image
'Inputs: Amount of noise, monochromatic or not
Public Sub AddNoise(ByVal Noise As Long, ByVal MC As Boolean)

    'Although it's slow, we're stuck using random numbers for noise addition
    Randomize Timer
    
    Message "Adding noise..."
    
    SetProgBarMax PicWidthL
    
    Dim Ncolor As Long
    Dim dNoise As Long
    
    'Double the amount of noise we plan on using (so we can add noise above or below 0)
    dNoise = Noise * 2
    
    Dim r As Long, g As Long, b As Long
    Dim QuickX As Long
    
    For x = 0 To PicWidthL
        QuickX = x * 3
        
        For y = 0 To PicHeightL
            
            r = ImageData(QuickX + 2, y)
            g = ImageData(QuickX + 1, y)
            b = ImageData(QuickX, y)
            
            If MC = True Then
                'Monochromatic noise - same amount for each color
                Ncolor = (dNoise * Rnd) - Noise
                r = r + Ncolor
                g = g + Ncolor
                b = b + Ncolor
            Else
                'Colored noise - each color generated randomly
                r = r + (dNoise * Rnd) - Noise
                g = g + (dNoise * Rnd) - Noise
                b = b + (dNoise * Rnd) - Noise
            End If
            
            'Trim values
            ByteMeL r
            ByteMeL g
            ByteMeL b
            
            'Replace pixel data
            ImageData(QuickX + 2, y) = r
            ImageData(QuickX + 1, y) = g
            ImageData(QuickX, y) = b
            
        Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetImageData
    
End Sub

Private Sub Form_Load()
'Create the previews
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    PreviewNoise hsNoise.Value, ChkM.Value
End Sub


'Same as above, but performs the effect only on the preview boxes
Private Sub PreviewNoise(ByVal Noise As Long, ByVal MC As Boolean)
    Randomize Timer
    GetPreviewData PicPreview
    Dim Ncolor As Long
    Dim dNoise As Long
    dNoise = Noise * 2
    Dim r As Long, g As Long, b As Long
    Dim QuickX As Long
    For x = PreviewX To PreviewX + PreviewWidth
        QuickX = x * 3
        For y = PreviewY To PreviewY + PreviewHeight
            r = ImageData(QuickX + 2, y)
            g = ImageData(QuickX + 1, y)
            b = ImageData(QuickX, y)
            If MC = True Then
                Ncolor = (dNoise * Rnd) - Noise
                r = r + Ncolor
                g = g + Ncolor
                b = b + Ncolor
            Else
                r = r + (dNoise * Rnd) - Noise
                g = g + (dNoise * Rnd) - Noise
                b = b + (dNoise * Rnd) - Noise
            End If
            ByteMeL r
            ByteMeL g
            ByteMeL b
            ImageData(QuickX + 2, y) = r
            ImageData(QuickX + 1, y) = g
            ImageData(QuickX, y) = b
        Next y
    Next x
    SetPreviewData PicEffect
End Sub

'The following four routines keep the value of the textbox and scroll bar in lock-step
Private Sub hsNoise_Change()
    txtNoise.Text = hsNoise.Value
    PreviewNoise hsNoise.Value, ChkM.Value
End Sub

Private Sub hsNoise_Scroll()
    txtNoise.Text = hsNoise.Value
    PreviewNoise hsNoise.Value, ChkM.Value
End Sub

Private Sub txtNoise_Change()
    If EntryValid(txtNoise, hsNoise.Min, hsNoise.Max, False, False) Then hsNoise.Value = val(txtNoise)
End Sub

Private Sub txtNoise_GotFocus()
    AutoSelectText txtNoise
End Sub
