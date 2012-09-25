VERSION 5.00
Begin VB.Form FormWhiteBalance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " White Balance"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   361
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsIgnore 
      Height          =   255
      Left            =   360
      Max             =   100
      Min             =   1
      TabIndex        =   1
      Top             =   3840
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
      Left            =   5280
      TabIndex        =   0
      Text            =   "0.05"
      Top             =   3795
      Width           =   735
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   4920
      TabIndex        =   3
      Top             =   4920
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   4920
      Width           =   1125
   End
   Begin VB.Label lblAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "after"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblBefore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "before"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   480
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
      Left            =   240
      TabIndex        =   4
      Top             =   3480
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
    Else
        AutoSelectText txtIgnore
    End If
End Sub

Private Sub Form_Activate()
    
    DrawPreviewImage picPreview
    AutoWhiteBalance CSng(val(txtIgnore)), True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'Correct white balance by stretching the histogram and ignoring pixels above or below the 0.05% threshold
Public Sub AutoWhiteBalance(Optional ByVal percentIgnore As Single = 0.05, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Preparing histogram data..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    Dim iWidth As Long, iHeight As Long
    iWidth = curLayerValues.Width
    iHeight = curLayerValues.Height
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color values
    Dim r As Long, g As Long, b As Long
    
    'Maximum and minimum values, which will be detected by our initial histogram run
    Dim RMax As Byte, GMax As Byte, BMax As Byte
    Dim RMin As Byte, GMin As Byte, BMin As Byte
    RMax = 0: GMax = 0: BMax = 0
    RMin = 255: GMin = 255: BMin = 255
    
    'Shrink the percentIgnore value down to 1% of the value we are passed (you'll see why in a moment)
    percentIgnore = percentIgnore / 100
    
    'Prepare histogram arrays
    Dim rCount(0 To 255) As Long, gCount(0 To 255) As Long, bCount(0 To 255) As Long
    For x = 0 To 255
        rCount(x) = 0
        gCount(x) = 0
        bCount(x) = 0
    Next x
    
    'Build the image histogram
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
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
    NumOfPixels = iWidth * iHeight
    
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
    
    'Finally, calculate the difference between max and min for each color
    Dim Rdif As Long, Gdif As Long, Bdif As Long
    Rdif = CLng(RMax) - CLng(RMin)
    Gdif = CLng(GMax) - CLng(GMin)
    Bdif = CLng(BMax) - CLng(BMin)
    
    'We can now build a final set of look-up tables that contain the results of every possible color transformation
    Dim rFinal(0 To 255) As Byte, gFinal(0 To 255) As Byte, bFinal(0 To 255) As Byte
    
    For x = 0 To 255
        If Rdif <> 0 Then r = 255 * ((x - RMin) / Rdif) Else r = x
        If Gdif <> 0 Then g = 255 * ((x - GMin) / Gdif) Else g = x
        If Bdif <> 0 Then b = 255 * ((x - BMin) / Bdif) Else b = x
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        rFinal(x) = r
        gFinal(x) = g
        bFinal(x) = b
    Next x
    
    If toPreview = False Then Message "Adjusting image white balance..."
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
            
        'Adjust white balance in a single pass (thanks to the magic of look-up tables)
        ImageData(QuickVal + 2, y) = rFinal(ImageData(QuickVal + 2, y))
        ImageData(QuickVal + 1, y) = gFinal(ImageData(QuickVal + 1, y))
        ImageData(QuickVal, y) = bFinal(ImageData(QuickVal, y))
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'When the horizontal scroll bar is moved, change the text box to match
Private Sub hsIgnore_Change()
    txtIgnore.Text = Format(CSng(hsIgnore.Value) / 100, "0.00")
    txtIgnore.Refresh
    AutoWhiteBalance CSng(val(txtIgnore)), True, picEffect
End Sub

Private Sub hsIgnore_Scroll()
    txtIgnore.Text = Format(CSng(hsIgnore.Value) / 100, "0.00")
    txtIgnore.Refresh
    AutoWhiteBalance CSng(val(txtIgnore)), True, picEffect
End Sub

Private Sub txtIgnore_GotFocus()
    AutoSelectText txtIgnore
End Sub

'If the user changes the gamma value by hand, check it for numerical correctness, then change the horizontal scroll bar to match
Private Sub txtIgnore_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtIgnore, , True
    If EntryValid(txtIgnore, hsIgnore.Min / 100, hsIgnore.Max / 100, False, False) Then hsIgnore.Value = val(txtIgnore) * 100
End Sub

