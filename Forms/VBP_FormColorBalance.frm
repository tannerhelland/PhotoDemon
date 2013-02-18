VERSION 5.00
Begin VB.Form FormColorBalance 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Adjust Color Balance"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12090
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
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
   Begin VB.TextBox txtBlue 
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
      Left            =   11160
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "0"
      Top             =   3435
      Width           =   615
   End
   Begin VB.HScrollBar hsBlue 
      Height          =   255
      Left            =   6120
      Max             =   100
      Min             =   -100
      TabIndex        =   8
      Top             =   3480
      Width           =   4935
   End
   Begin VB.TextBox txtGreen 
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
      Left            =   11160
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "0"
      Top             =   2475
      Width           =   615
   End
   Begin VB.HScrollBar hsGreen 
      Height          =   255
      Left            =   6120
      Max             =   100
      Min             =   -100
      TabIndex        =   5
      Top             =   2520
      Width           =   4935
   End
   Begin VB.TextBox txtRed 
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
      Left            =   11160
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "0"
      Top             =   1515
      Width           =   615
   End
   Begin VB.HScrollBar hsRed 
      Height          =   255
      Left            =   6120
      Max             =   100
      Min             =   -100
      TabIndex        =   2
      Top             =   1560
      Width           =   4935
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "yellow"
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
      Left            =   6360
      TabIndex        =   15
      Top             =   3840
      Width           =   675
   End
   Begin VB.Label lblMagenta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "magenta"
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
      Left            =   6360
      TabIndex        =   14
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label lblCyan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cyan"
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
      Left            =   6360
      TabIndex        =   13
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label lblBlue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "blue"
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
      Left            =   10350
      TabIndex        =   9
      Top             =   3840
      Width           =   450
   End
   Begin VB.Label lblGreen 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "green"
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
      Left            =   10230
      TabIndex        =   6
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblRed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "red"
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
      Left            =   10455
      TabIndex        =   3
      Top             =   1920
      Width           =   345
   End
End
Attribute VB_Name = "FormColorBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Balance Adjustment Form
'Copyright ©2012-2013 by Tanner Helland
'Created: 31/January/13
'Last updated: 17/February/13
'Last update: remove "preserve luminance" slider.  The image always looks better with luminance preserved.
'
'Fairly simple and standard color adjustment form.  Layout and feature set derived from comparable tools
' in GIMP and Photoshop.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    
    'Validate all textbox entries
    If Not EntryValid(txtRed, hsRed.Min, hsRed.Max, True, True) Then
        AutoSelectText txtRed
        Exit Sub
    End If
    
    'Validate all textbox entries
    If Not EntryValid(txtGreen, hsGreen.Min, hsGreen.Max, True, True) Then
        AutoSelectText txtGreen
        Exit Sub
    End If
    
    'Validate all textbox entries
    If Not EntryValid(txtBlue, hsBlue.Min, hsBlue.Max, True, True) Then
        AutoSelectText txtBlue
        Exit Sub
    End If
    
    Me.Visible = False
    Process AdjustColorBalance, CLng(hsRed), CLng(hsGreen), CLng(hsBlue), True
    Unload Me
    
End Sub

'Apply a new color balance to the image
' Input: offset for each of red, green, and blue
Public Sub ApplyColorBalance(ByVal rVal As Long, ByVal gVal As Long, ByVal bVal As Long, ByVal preserveLuminance As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Adjusting color balance..."
    
    Dim rModifier As Long, gModifier As Long, bModifier As Long
    rModifier = 0
    gModifier = 0
    bModifier = 0
    
    'Now, Build actual RGB modifiers based off the values provided
    If rVal < 0 Then
        gModifier = gModifier + -rVal
        bModifier = bModifier + -rVal
    Else
        rModifier = rModifier + rVal
    End If
   
    If gVal < 0 Then
        rModifier = rModifier + -gVal
        bModifier = bModifier + -gVal
    Else
        gModifier = gModifier + gVal
    End If
    
    If bVal < 0 Then
        rModifier = rModifier + -bVal
        gModifier = gModifier + -bVal
    Else
        bModifier = bModifier + bVal
    End If
    
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
    Dim h As Double, s As Double, l As Double
    
    Dim origLuminance As Double
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Get the original luminance
        origLuminance = getLuminance(r, g, b) / 255
        
        'Apply the modifiers
        r = r + rModifier
        g = g + gModifier
        b = b + bModifier
        
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        
        'If the user doesn't want us to maintain luminance, our work is done - assign the new values.
        'If they do want us to maintain luminance, things are a bit trickier.  We need to convert our values to
        ' HSL, then substitute the original luminance and convert back to RGB.
        If preserveLuminance Then
        
            'Convert the new values to HSL
            tRGBToHSL r, g, b, h, s, l
            
            'Now, convert back, using the original luminance
            tHSLToRGB h, s, origLuminance, r, g, b
        
        End If
        
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
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

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When the hue scroll bar is changed, redraw the preview
Private Sub hsRed_Change()
    copyToTextBoxI txtRed, hsRed.Value
    updatePreview
End Sub

Private Sub hsRed_Scroll()
    copyToTextBoxI txtRed, hsRed.Value
    updatePreview
End Sub

Private Sub hsBlue_Change()
    copyToTextBoxI txtBlue, hsBlue.Value
    updatePreview
End Sub

Private Sub hsBlue_Scroll()
    copyToTextBoxI txtBlue, hsBlue.Value
    updatePreview
End Sub

Private Sub hsGreen_Change()
    copyToTextBoxI txtGreen, hsGreen.Value
    updatePreview
End Sub

Private Sub hsGreen_Scroll()
    copyToTextBoxI txtGreen, hsGreen.Value
    updatePreview
End Sub

Private Sub txtRed_GotFocus()
    AutoSelectText txtRed
End Sub

Private Sub txtRed_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtRed, True
    If EntryValid(txtRed, hsRed.Min, hsRed.Max, False, False) Then hsRed.Value = Val(txtRed)
End Sub

Private Sub txtBlue_GotFocus()
    AutoSelectText txtBlue
End Sub

Private Sub txtBlue_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtBlue, True
    If EntryValid(txtBlue, hsBlue.Min, hsBlue.Max, False, False) Then hsBlue.Value = Val(txtBlue)
End Sub

Private Sub txtGreen_GotFocus()
    AutoSelectText txtGreen
End Sub

Private Sub txtGreen_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtGreen, True
    If EntryValid(txtGreen, hsGreen.Min, hsGreen.Max, False, False) Then hsGreen.Value = Val(txtGreen)
End Sub

Private Sub updatePreview()
    ApplyColorBalance hsRed, hsGreen, hsBlue, True, True, fxPreview
End Sub
