VERSION 5.00
Begin VB.Form FormWhiteBalance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " White Balance"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsIgnore 
      Height          =   255
      Left            =   240
      Max             =   100
      Min             =   1
      TabIndex        =   1
      Top             =   3360
      Value           =   5
      Width           =   4575
   End
   Begin VB.TextBox txtIgnore 
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
      Text            =   "0.05"
      Top             =   2850
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      Left            =   3600
      TabIndex        =   3
      Top             =   4080
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
      Left            =   2400
      TabIndex        =   2
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label lblPreview 
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
   Begin VB.Label lblAmount 
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
      TabIndex        =   4
      Top             =   2880
      Width           =   720
   End
End
Attribute VB_Name = "FormWhiteBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'White Balance Handler
'Copyright ©2000-2012 by Tanner Helland
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
Private Sub CmdOK_Click()
    'The scroll bar max and min values are used to check the gamma input for validity
    If EntryValid(txtIgnore, hsIgnore.Min / 100, hsIgnore.Max / 100) Then
        Me.Visible = False
        Process WhiteBalance, CSng(val(txtIgnore))
        Unload Me
    End If
End Sub

'Initialize the preview boxes and the gamma combo box
Private Sub Form_Load()
    
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    
    PreviewWhiteBalance CSng(val(txtIgnore))
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub


'Correct white balance by stretching the histogram and ignoring pixels above or below the 0.05% threshold
Public Sub AutoWhiteBalance(Optional ByVal percentIgnore As Single = 0.05)

    Dim r As Long, g As Long, b As Long
    
    Dim RMax As Byte, GMax As Byte, BMax As Byte
    Dim RMin As Byte, GMin As Byte, BMin As Byte
    RMax = 0: GMax = 0: BMax = 0
    RMin = 255: GMin = 255: BMin = 255
    
    Message "Preparing histogram data..."
    
    GetImageData
    
    percentIgnore = percentIgnore / 100
    
    'Prepare histogram arrays
    Dim rCount(0 To 255) As Long, gCount(0 To 255) As Long, bCount(0 To 255) As Long
    For x = 0 To 255
        rCount(x) = 0
        gCount(x) = 0
        bCount(x) = 0
    Next x
    
    'Build the image histogram
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        rCount(r) = rCount(r) + 1
        gCount(g) = gCount(g) + 1
        bCount(b) = bCount(b) + 1
    Next y
    Next x
    
    'With the histogram complete, we can now figure out how to stretch the RGB channels. We do this by calculating a min/max
    ' ratio where the top and bottom 0.05% (or user-specified value) of pixels are ignored.
    
    Dim foundYet As Boolean
    foundYet = False
    
    Dim NumOfPixels As Long
    NumOfPixels = PicWidthL * PicHeightL
    
    Dim wbThreshold As Long
    wbThreshold = NumOfPixels * percentIgnore
    
    r = 0: g = 0: b = 0
    
    Dim rTally As Long, gTally As Long, bTally As Long
    rTally = 0: gTally = 0: bTally = 0
    
    'Find minimum values of red, green, and blue
    Do
        If rCount(r) + rTally < wbThreshold Then
            r = r + 1
            rTally = rTally + rCount(r)
        Else
            RMin = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If gCount(g) + gTally < wbThreshold Then
            g = g + 1
            gTally = gTally + gCount(g)
        Else
            GMin = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If bCount(b) + bTally < wbThreshold Then
            b = b + 1
            bTally = bTally + bCount(b)
        Else
            BMin = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    'Now, find maximum values of red, green, and blue
    foundYet = False
    
    r = 255: g = 255: b = 255
    rTally = 0: gTally = 0: bTally = 0
    
    Do
        If rCount(r) + rTally < wbThreshold Then
            r = r - 1
            rTally = rTally + rCount(r)
        Else
            RMax = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If gCount(g) + gTally < wbThreshold Then
            g = g - 1
            gTally = gTally + gCount(g)
        Else
            GMax = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If bCount(b) + bTally < wbThreshold Then
            b = b - 1
            bTally = bTally + bCount(b)
        Else
            BMax = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    Message "Automatically adjusting white balance..."
    SetProgBarMax PicWidthL
    Dim Rdif As Integer, Gdif As Integer, Bdif As Integer
    Rdif = CInt(RMax) - CInt(RMin)
    Gdif = CInt(GMax) - CInt(GMin)
    Bdif = CInt(BMax) - CInt(BMin)
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        If Rdif <> 0 Then r = 255 * ((r - RMin) / Rdif)
        If Gdif <> 0 Then g = 255 * ((g - GMin) / Gdif)
        If Bdif <> 0 Then b = 255 * ((b - BMin) / Bdif)
        ByteMeL r
        ByteMeL g
        ByteMeL b
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
    Message "Finished."
End Sub

'When the horizontal scroll bar is moved, change the text box to match
Private Sub hsIgnore_Change()
    txtIgnore.Text = Format(CSng(hsIgnore.Value) / 100, "0.00")
    PreviewWhiteBalance CSng(val(txtIgnore))
End Sub

Private Sub hsIgnore_Scroll()
    txtIgnore.Text = Format(CSng(hsIgnore.Value) / 100, "0.00")
    PreviewWhiteBalance CSng(val(txtIgnore))
End Sub

'Render a gamma preview to the preview picture box
Private Sub PreviewWhiteBalance(ByVal percentIgnore As Single)
   
    Dim r As Long, g As Long, b As Long
    
    Dim RMax As Byte, GMax As Byte, BMax As Byte
    Dim RMin As Byte, GMin As Byte, BMin As Byte
    RMax = 0: GMax = 0: BMax = 0
    RMin = 255: GMin = 255: BMin = 255
        
    GetPreviewData PicPreview
    
    percentIgnore = percentIgnore / 100
    
    'Prepare histogram arrays
    Dim rCount(0 To 255) As Long, gCount(0 To 255) As Long, bCount(0 To 255) As Long
    For x = 0 To 255
        rCount(x) = 0
        gCount(x) = 0
        bCount(x) = 0
    Next x
    
    'Build the image histogram
    Dim QuickVal As Long
    For x = PreviewX To PreviewX + PreviewWidth
        QuickVal = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        rCount(r) = rCount(r) + 1
        gCount(g) = gCount(g) + 1
        bCount(b) = bCount(b) + 1
    Next y
    Next x
    
    'With the histogram complete, we can now figure out how to stretch the RGB channels. We do this by calculating a min/max
    ' ratio where the top and bottom 0.05% (or user-specified value) of pixels are ignored.
    
    Dim foundYet As Boolean
    foundYet = False
    
    Dim NumOfPixels As Long
    NumOfPixels = (PreviewWidth - PreviewX) * (PreviewHeight - PreviewY)
    
    Dim wbThreshold As Long
    wbThreshold = NumOfPixels * percentIgnore
    
    r = 0: g = 0: b = 0
    
    Dim rTally As Long, gTally As Long, bTally As Long
    rTally = 0: gTally = 0: bTally = 0
    
    'Find minimum values of red, green, and blue
    Do
        If rCount(r) + rTally < wbThreshold Then
            r = r + 1
            rTally = rTally + rCount(r)
        Else
            RMin = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If gCount(g) + gTally < wbThreshold Then
            g = g + 1
            gTally = gTally + gCount(g)
        Else
            GMin = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If bCount(b) + bTally < wbThreshold Then
            b = b + 1
            bTally = bTally + bCount(b)
        Else
            BMin = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    'Now, find maximum values of red, green, and blue
    foundYet = False
    
    r = 255: g = 255: b = 255
    rTally = 0: gTally = 0: bTally = 0
    
    Do
        If rCount(r) + rTally < wbThreshold Then
            r = r - 1
            rTally = rTally + rCount(r)
        Else
            RMax = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If gCount(g) + gTally < wbThreshold Then
            g = g - 1
            gTally = gTally + gCount(g)
        Else
            GMax = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If bCount(b) + bTally < wbThreshold Then
            b = b - 1
            bTally = bTally + bCount(b)
        Else
            BMax = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    Dim Rdif As Integer, Gdif As Integer, Bdif As Integer
    Rdif = CInt(RMax) - CInt(RMin)
    Gdif = CInt(GMax) - CInt(GMin)
    Bdif = CInt(BMax) - CInt(BMin)
    For x = PreviewX To PreviewX + PreviewWidth
        QuickVal = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        If Rdif <> 0 Then r = 255 * ((r - RMin) / Rdif)
        If Gdif <> 0 Then g = 255 * ((g - GMin) / Gdif)
        If Bdif <> 0 Then b = 255 * ((b - BMin) / Bdif)
        ByteMeL r
        ByteMeL g
        ByteMeL b
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
    Next x
    SetPreviewData PicEffect
    
End Sub

Private Sub txtIgnore_GotFocus()
    AutoSelectText txtIgnore
End Sub

'If the user changes the gamma value by hand, check it for numerical correctness, then change the horizontal scroll bar to match
Private Sub txtIgnore_Change()
    If EntryValid(txtIgnore, hsIgnore.Min / 100, hsIgnore.Max / 100, False, False) Then hsIgnore.Value = val(txtIgnore) * 100
End Sub
