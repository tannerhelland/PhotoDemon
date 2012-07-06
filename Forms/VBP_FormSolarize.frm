VERSION 5.00
Begin VB.Form FormSolarize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Solarize"
   ClientHeight    =   4530
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
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsThreshold 
      Height          =   255
      Left            =   240
      Max             =   254
      Min             =   1
      MouseIcon       =   "VBP_FormSolarize.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3240
      Value           =   127
      Width           =   4575
   End
   Begin VB.TextBox txtThreshold 
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
      Text            =   "127"
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      MouseIcon       =   "VBP_FormSolarize.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3960
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      MouseIcon       =   "VBP_FormSolarize.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3960
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
      Caption         =   "Threshold:"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   2790
      Width           =   870
   End
End
Attribute VB_Name = "FormSolarize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Solarizing Effect Handler
'©2000-2012 Tanner Helland
'Created: 4/14/01
'Last updated: 05/July/12
'Last update: optimized for speed
'
'Updated solarizing interface; it has been optimized for speed and
'  ease-of-implementation.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max) Then
        Me.Visible = False
        Process Solarize, hsThreshold.Value
        Unload Me
    Else
        AutoSelectText txtThreshold
    End If
End Sub

'Subroutine for "solarizing" an image
Public Sub SolarizeImage(ByVal Threshold As Byte)
    
    Message "Solarizing image..."
    
    Dim r As Byte, g As Byte, b As Byte
    
    GetImageData
    SetProgBarMax PicWidthL
    
    Dim QuickVal As Long
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Solarizing is simple - invert every value above the threshold
        If r > Threshold Then r = 255 - r
        If g > Threshold Then g = 255 - g
        If b > Threshold Then b = 255 - b
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetProgBarVal PicWidthL
    SetImageData
    
End Sub

Private Sub Form_Load()
'Create the previews
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    PreviewSolarize hsThreshold.Value
End Sub

'Same as above, but exclusively for previewing
Private Sub PreviewSolarize(ByVal Threshold As Byte)

    Dim r As Byte, g As Byte, b As Byte
    
    GetPreviewData PicPreview
    
    Dim QuickVal As Long
    
    For x = PreviewX To PreviewX + PreviewWidth
        QuickVal = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        If r > Threshold Then r = 255 - r
        If g > Threshold Then g = 255 - g
        If b > Threshold Then b = 255 - b
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
    Next x
    
    SetPreviewData PicEffect
    
End Sub

'When the horizontal scroll bar is moved, update the preview and text box to match
Private Sub hsThreshold_Change()
    txtThreshold.Text = hsThreshold.Value
    PreviewSolarize hsThreshold.Value
End Sub

Private Sub hsThreshold_Scroll()
    txtThreshold.Text = hsThreshold.Value
    PreviewSolarize hsThreshold.Value
End Sub

'When the text box is changed, update the preview and text box to match (assuming the text box value is valid)
Private Sub txtThreshold_Change()
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max, False, False) Then hsThreshold.Value = val(txtThreshold)
End Sub

Private Sub txtThreshold_GotFocus()
    AutoSelectText txtThreshold
End Sub
