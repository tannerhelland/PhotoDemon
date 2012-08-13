VERSION 5.00
Begin VB.Form FormFade 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Fade Image"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
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
   ScaleHeight     =   308
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsPercent 
      Height          =   255
      Left            =   240
      Max             =   100
      Min             =   1
      TabIndex        =   1
      Top             =   3360
      Value           =   50
      Width           =   4575
   End
   Begin VB.TextBox txtPercent 
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
      Text            =   "50"
      Top             =   2850
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
      Negotiate       =   -1  'True
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   5
      TabStop         =   0   'False
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   4080
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Percent:"
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
      TabIndex        =   7
      Top             =   2880
      Width           =   705
   End
   Begin VB.Label Label2 
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
      TabIndex        =   6
      Top             =   2310
      Width           =   4575
   End
End
Attribute VB_Name = "FormFade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Fade Filter Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 11/7/01
'Last updated: 19/June/12
'Last update: condensed all fade routines into a single, percentage-based one.  The speed increase provided by
'             individual routines for various values was not proportionate to the extra code required.
'
'Module for handling the fade-style filter.  Basically, it's a quasi-contrast
'mixer between the grayscale and color images.
'
'***************************************************************************

Option Explicit

'OK button
Private Sub CmdOK_Click()
    'Error checking
    If EntryValid(txtPercent, hsPercent.Min, hsPercent.Max) Then
        Me.Visible = False
        Process Fade, hsPercent.Value
        Unload Me
    Else
        AutoSelectText txtPercent
    End If
End Sub

'Subroutine for fading an image to grayscale
Public Sub FadeImage(ByVal PercentFade As Long)
    
    Message "Fading image to gray..."
    
    GetImageData
    
    Dim r As Long, g As Long, b As Long
    Dim tGray As Long
    
    SetProgBarMax PicWidthL
    
    Dim QuickX As Long
    
    For x = 0 To PicWidthL
        QuickX = x * 3
    For y = 0 To PicHeightL
        
        r = ImageData(QuickX + 2, y)
        g = ImageData(QuickX + 1, y)
        b = ImageData(QuickX, y)
        
        'Calculate grayscale equivalent
        tGray = Int((222 * r + 707 * g + 71 * b) \ 1000)
        
        'Alphablend the colors with grayscale based on the percent fade function
        r = MixColors(r, tGray, PercentFade)
        g = MixColors(g, tGray, PercentFade)
        b = MixColors(b, tGray, PercentFade)
        
        'Set the new colors
        ImageData(QuickX + 2, y) = r
        ImageData(QuickX + 1, y) = g
        ImageData(QuickX, y) = b
        
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetImageData
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'Unfade is literally a reverse fade - rather than pushing values toward gray, we push them away from it
Public Sub UnfadeImage()
    
    Dim r As Long, g As Long, b As Long
    Dim tGray As Long
    
    Message "Unfading image..."
    
    SetProgBarMax PicWidthL
    
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Grayscale calculation
        tGray = Int((222 * r + 707 * g + 71 * b) \ 1000)
        
        'Use the contrast formula to move the colors AWAY from gray, rather
        'than closer to it (as the above formulas do)
        r = Abs(r - (tGray \ 2)) * 2
        g = Abs(g - (tGray \ 2)) * 2
        b = Abs(b - (tGray \ 2)) * 2
        
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
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
    PreviewFadeImage hsPercent.Value
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'Same routine as above, but meant for previewing only
Private Sub PreviewFadeImage(ByVal PercentFade As Long)
    
    GetPreviewData PicPreview
    
    Dim r As Long, g As Long, b As Long
    Dim tGray As Long
    
    Dim QuickX As Long
    For x = PreviewX To PreviewX + PreviewWidth
        QuickX = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        r = ImageData(QuickX + 2, y)
        g = ImageData(QuickX + 1, y)
        b = ImageData(QuickX, y)
        tGray = Int((222 * r + 707 * g + 71 * b) \ 1000)
        r = MixColors(r, tGray, PercentFade)
        g = MixColors(g, tGray, PercentFade)
        b = MixColors(b, tGray, PercentFade)
        ImageData(QuickX + 2, y) = r
        ImageData(QuickX + 1, y) = g
        ImageData(QuickX, y) = b
    Next y
    Next x
    SetPreviewData PicEffect
    
End Sub

Private Sub hsPercent_Change()
    txtPercent.Text = hsPercent.Value
    PreviewFadeImage hsPercent.Value
End Sub

Private Sub hsPercent_Scroll()
    txtPercent.Text = hsPercent.Value
    PreviewFadeImage hsPercent.Value
End Sub

Private Sub txtPercent_Change()
    If EntryValid(txtPercent, hsPercent.Min, hsPercent.Max, False, False) Then
        hsPercent.Value = val(txtPercent)
    End If
End Sub

Private Sub txtPercent_GotFocus()
    AutoSelectText txtPercent
End Sub

