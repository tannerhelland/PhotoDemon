VERSION 5.00
Begin VB.Form FormGamma 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Gamma Correction"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12060
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   804
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtGamma 
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
      Index           =   2
      Left            =   11160
      MaxLength       =   4
      TabIndex        =   14
      Text            =   "1.00"
      Top             =   4620
      Width           =   615
   End
   Begin VB.HScrollBar hsGamma 
      Height          =   255
      Index           =   2
      Left            =   6120
      Max             =   300
      Min             =   1
      TabIndex        =   13
      Top             =   4680
      Value           =   100
      Width           =   4935
   End
   Begin VB.TextBox txtGamma 
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
      Index           =   1
      Left            =   11160
      MaxLength       =   4
      TabIndex        =   11
      Text            =   "1.00"
      Top             =   3780
      Width           =   615
   End
   Begin VB.HScrollBar hsGamma 
      Height          =   255
      Index           =   1
      Left            =   6120
      Max             =   300
      Min             =   1
      TabIndex        =   10
      Top             =   3840
      Value           =   100
      Width           =   4935
   End
   Begin VB.CheckBox chkUnison 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " keep all colors in sync"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   5160
      Value           =   1  'Checked
      Width           =   4935
   End
   Begin VB.PictureBox picChart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   8280
      ScaleHeight     =   159
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   7
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10590
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin VB.HScrollBar hsGamma 
      Height          =   255
      Index           =   0
      Left            =   6120
      Max             =   300
      Min             =   1
      TabIndex        =   2
      Top             =   3000
      Value           =   100
      Width           =   4935
   End
   Begin VB.TextBox txtGamma 
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
      Index           =   0
      Left            =   11160
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "1.00"
      Top             =   2940
      Width           =   615
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "blue:"
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
      Index           =   3
      Left            =   6000
      TabIndex        =   15
      Top             =   4320
      Width           =   540
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "green:"
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
      Index           =   0
      Left            =   6000
      TabIndex        =   12
      Top             =   3480
      Width           =   690
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "new gamma curve:"
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
      Index           =   2
      Left            =   6000
      TabIndex        =   8
      Top             =   1170
      Width           =   2040
   End
   Begin VB.Label lblBackground 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "red:"
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
      Index           =   1
      Left            =   6000
      TabIndex        =   4
      Top             =   2640
      Width           =   435
   End
End
Attribute VB_Name = "FormGamma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gamma Correction Handler
'Copyright ©2000-2013 by Tanner Helland
'Created: 12/May/01
'Last updated: 19/January/13
'Last update: added a gamma chart to the form
'
'Updated version of the gamma handler; fully optimized, it uses a look-up
' table and can correct any color channel.
'
'***************************************************************************

Option Explicit

Dim userChange As Boolean

Private Sub chkUnison_Click()
    
    If CBool(chkUnison) Then
        Dim newGamma As Double
        newGamma = CSng(hsGamma(0) + hsGamma(1) + hsGamma(2)) / 3
    
        userChange = False
        hsGamma(0) = newGamma
        hsGamma(1) = newGamma
        hsGamma(2) = newGamma
        userChange = True
    End If
    
    updatePreview
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    'The scroll bar max and min values are used to check the gamma input for validity
    Dim i As Long
    
    For i = 0 To 2
        If Not EntryValid(txtGamma(i), hsGamma(i).Min / 100, hsGamma(i).Max / 100, True, True) Then
            AutoSelectText txtGamma(i)
            Exit Sub
        End If
    Next i
        
    Me.Visible = False
    Process GammaCorrection, CSng(Val(txtGamma(0))), CSng(Val(txtGamma(1))), CSng(Val(txtGamma(2)))
    Unload Me
End Sub

Private Sub Form_Activate()
    
    userChange = True
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Finally, render a preview
    updatePreview
    
End Sub

'Basic gamma correction.  It's a simple function - use an exponent to adjust R/G/B values.
' Inputs: new gamma level, which channels to adjust (r/g/b/all), and optional preview information
Public Sub GammaCorrect(ByVal rGamma As Double, ByVal gGamma As Double, ByVal bGamma As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
     
    If toPreview = False Then Message "Adjusting gamma values..."
    
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
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Gamma can be easily applied using a look-up table
    Dim gLookup(0 To 2, 0 To 255) As Byte
    Dim tmpVal As Double
    
    For y = 0 To 2
    For x = 0 To 255
        tmpVal = x / 255
        Select Case y
            Case 0
                tmpVal = tmpVal ^ (1 / rGamma)
            Case 1
                tmpVal = tmpVal ^ (1 / gGamma)
            Case 2
                tmpVal = tmpVal ^ (1 / bGamma)
        End Select
        tmpVal = tmpVal * 255
        
        If tmpVal > 255 Then tmpVal = 255
        If tmpVal < 0 Then tmpVal = 0
        
        gLookup(y, x) = tmpVal
    Next x
    Next y
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
                
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, y) = gLookup(0, r)
        ImageData(QuickVal + 1, y) = gLookup(1, g)
        ImageData(QuickVal, y) = gLookup(2, b)
        
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

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When the horizontal scroll bar is moved, change the text box to match
Private Sub hsGamma_Change(Index As Integer)
    
    copyToTextBoxF CSng(hsGamma(Index).Value) / 100, txtGamma(Index)
        
    If userChange Then
        userChange = False
        
        If CBool(chkUnison) Then
            Select Case Index
                Case 0
                    hsGamma(1).Value = hsGamma(0).Value
                    hsGamma(2).Value = hsGamma(0).Value
                Case 1
                    hsGamma(0).Value = hsGamma(1).Value
                    hsGamma(2).Value = hsGamma(1).Value
                Case 2
                    hsGamma(0).Value = hsGamma(2).Value
                    hsGamma(1).Value = hsGamma(2).Value
            End Select
        End If
        
        userChange = True
        
        updatePreview
    End If
        
End Sub

Private Sub hsGamma_Scroll(Index As Integer)
    
    copyToTextBoxF CSng(hsGamma(Index).Value) / 100, txtGamma(Index)
        
    If userChange Then
        userChange = False
        
        If CBool(chkUnison) Then
            Select Case Index
                Case 0
                    hsGamma(1).Value = hsGamma(0).Value
                    hsGamma(2).Value = hsGamma(0).Value
                Case 1
                    hsGamma(0).Value = hsGamma(1).Value
                    hsGamma(2).Value = hsGamma(1).Value
                Case 2
                    hsGamma(0).Value = hsGamma(2).Value
                    hsGamma(1).Value = hsGamma(2).Value
            End Select
        End If
        
        userChange = True
        
        updatePreview
    End If
End Sub

Private Sub txtGamma_GotFocus(Index As Integer)
    AutoSelectText txtGamma
End Sub

'If the user changes the gamma value by hand, check it for numerical correctness, then change the horizontal scroll bar to match
Private Sub txtGamma_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    textValidate txtGamma(Index), , True
    If EntryValid(txtGamma(Index), hsGamma(Index).Min / 100, hsGamma(Index).Max / 100, False, False) And userChange Then hsGamma(Index).Value = Val(txtGamma(Index)) * 100
End Sub

'Redraw the preview effect and the gamma chart
Private Sub updatePreview()
    
    Dim prevX As Double, prevY As Double
    Dim curX As Double, curY As Double
    Dim x As Long, y As Long
    
    Dim xWidth As Long, yHeight As Long
    xWidth = picChart.ScaleWidth
    yHeight = picChart.ScaleHeight
        
    'Clear out the old chart and draw a gray line across the diagonal for reference
    picChart.Picture = LoadPicture("")
    picChart.ForeColor = RGB(127, 127, 127)
    DrawLineWuAA picChart.hDC, 0, yHeight, xWidth, 0, RGB(127, 127, 127)
    
    Dim gamVal As Double, tmpVal As Double
    
    'Draw each of the current gamma curves for the user's reference
    For y = 0 To 2
        
        'If all channels are in sync, draw only blue; otherwise, color each channel individually
        gamVal = Val(txtGamma(y))
        If (txtGamma(0) = txtGamma(1)) And (txtGamma(1) = txtGamma(2)) Then
            picChart.ForeColor = RGB(0, 0, 255)
        Else
        
            Select Case y
                Case 0
                    picChart.ForeColor = RGB(255, 0, 0)
                Case 1
                    picChart.ForeColor = RGB(0, 192, 0)
                Case 2
                    picChart.ForeColor = RGB(0, 0, 255)
            End Select
            
        End If
        
        prevX = 0
        prevY = yHeight
        curX = 0
        curY = yHeight
    
        'Draw the next channel (with antialiasing!)
        For x = 0 To xWidth
            tmpVal = x / xWidth
            tmpVal = tmpVal ^ (1 / gamVal)
            tmpVal = yHeight - (tmpVal * yHeight)
            curY = tmpVal
            curX = x
            DrawLineWuAA picChart.hDC, prevX, prevY, curX, curY, picChart.ForeColor
            prevX = curX
            prevY = curY
        Next x
        
    Next y
    
    picChart.Picture = picChart.Image
    picChart.Refresh

    'Once the chart is done, redraw the gamma preview as well
    GammaCorrect CSng(Val(txtGamma(0))), CSng(Val(txtGamma(1))), CSng(Val(txtGamma(2))), True, fxPreview
    
End Sub
