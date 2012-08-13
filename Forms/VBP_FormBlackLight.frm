VERSION 5.00
Begin VB.Form FormBlackLight 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Blacklight Options"
   ClientHeight    =   4335
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
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsIntensity 
      Height          =   255
      Left            =   240
      Max             =   10
      Min             =   1
      TabIndex        =   1
      Top             =   3240
      Value           =   2
      Width           =   4575
   End
   Begin VB.TextBox txtIntensity 
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
      Text            =   "2"
      Top             =   2850
      Width           =   495
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
      TabIndex        =   6
      Top             =   120
      Width           =   2175
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
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3840
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3840
      Width           =   1125
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
      TabIndex        =   7
      Top             =   2340
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Intensity:"
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
      Top             =   2880
      Width           =   780
   End
End
Attribute VB_Name = "FormBlackLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Blacklight Form
'Copyright ©2001-2012 by Tanner Helland
'Created: some time 2001
'Last updated: 05/July/12
'Last update: code clean-up and optimization
'
'I found this effect on accident, and it has turned out to be one of
'my favorite effects.  On some images it looks very visually stunning.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()

    If EntryValid(txtIntensity, hsIntensity.Min, hsIntensity.Max) Then
        Me.Visible = False
        Process BlackLight, hsIntensity.Value
        Unload Me
    End If
    
End Sub

'Perform a blacklight filter
'Input: strength of the filter (min 1, no real max - but above 7 it becomes increasingly blown-out)
Public Sub fxBlackLight(Optional ByVal Weight As Integer = 2)

    Message "Running black light filter..."
    
    Dim r As Long, g As Long, b As Long
    Dim tGray As Long
        
    GetImageData
    
    SetProgBarMax PicWidthL
    
    Dim QuickVal As Long
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate an accurate grayscale value
        tGray = Int((222 * r + 707 * g + 71 * b) \ 1000)
        
        'Perform the blacklight conversion
        r = Abs(r - tGray) * Weight
        g = Abs(g - tGray) * Weight
        b = Abs(b - tGray) * Weight
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Private Sub Form_Load()

    'Create a copy of the image on the preview window
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    
    'Actually do the effect
    DrawPreview
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'Same as above, but operates on the preview window
Private Sub DrawPreview()
    Dim r As Long, g As Long, b As Long
    Dim tR As Long
    GetPreviewData PicPreview
    Dim QuickVal As Long
    Dim Weight As Long
    Weight = hsIntensity.Value
    
    For x = PreviewX To PreviewX + PreviewWidth
        QuickVal = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        tR = Int((222 * r + 707 * g + 71 * b) \ 1000)
        r = Abs(r - tR) * Weight
        g = Abs(g - tR) * Weight
        b = Abs(b - tR) * Weight
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
    Next x
    SetPreviewData PicEffect
End Sub

'The next three routines keep the scroll bar and text box values in sync
Private Sub hsIntensity_Change()
    DrawPreview
    txtIntensity.Text = hsIntensity.Value
End Sub

Private Sub hsIntensity_Scroll()
    DrawPreview
    txtIntensity.Text = hsIntensity.Value
End Sub

Private Sub txtIntensity_Change()
    If EntryValid(txtIntensity, hsIntensity.Min, hsIntensity.Max, False, False) Then
        hsIntensity.Value = val(txtIntensity)
    End If
End Sub

Private Sub txtIntensity_GotFocus()
    AutoSelectText txtIntensity
End Sub
