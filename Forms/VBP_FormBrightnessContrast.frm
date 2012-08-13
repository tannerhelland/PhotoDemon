VERSION 5.00
Begin VB.Form FormBrightnessContrast 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Brightness/Contrast"
   ClientHeight    =   5910
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
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsContrast 
      Height          =   255
      Left            =   240
      Max             =   100
      Min             =   -100
      TabIndex        =   3
      Top             =   4440
      Width           =   4575
   End
   Begin VB.HScrollBar hsBright 
      Height          =   255
      Left            =   240
      Max             =   255
      Min             =   -255
      TabIndex        =   1
      Top             =   3360
      Width           =   4575
   End
   Begin VB.TextBox txtContrast 
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
      TabIndex        =   2
      Text            =   "0"
      Top             =   3930
      Width           =   495
   End
   Begin VB.TextBox txtBrightness 
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
      Text            =   "0"
      Top             =   2850
      Width           =   495
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
      TabIndex        =   10
      Top             =   120
      Width           =   2175
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
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
   Begin VB.CheckBox chkSample 
      Appearance      =   0  'Flat
      Caption         =   "Sampled Contrast"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   5400
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   5400
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
      TabIndex        =   11
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contrast:"
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
      TabIndex        =   8
      Top             =   3960
      Width           =   750
   End
   Begin VB.Label LblBrightness 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brightness:"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   2880
      Width           =   900
   End
End
Attribute VB_Name = "FormBrightnessContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Brightness and Contrast Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 2/6/01
'Last updated: 11/June/12
'Last update: removed all image-streaming code as part of a system-wide purge.
'
'The wonderful brightness/contrast handler.  Everything is done via look-up
' tables, so it's especially fast.  It's all linear (not logarithmic; sorry).
' Maybe someday I'll change that, maybe not... honestly, I probably won't, since
' brightness and contrast are such stupid functions anyway.  People should be
' using levels or curves or white balancing instead!
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()

    'Check the text box values against the limits of their corresponding scroll bars - that'll catch
    ' any out-of-range errors
    If EntryValid(txtBrightness, hsBright.Min, hsBright.Max) And EntryValid(txtContrast, hsContrast.Min, hsContrast.Max) Then
        
        Me.Visible = False
        
        'Re-route the effect through the software processor, so it can be tracked
        Process BrightnessAndContrast, hsBright.Value, hsContrast.Value, chkSample.Value
        
        Unload Me
    End If
    
End Sub

'Single routine for modifying both brightness and contrast.  Brightness is in the range (-255,255) while
' contrast is (-100,100).  Optionally, the image can be sampled to obtain a true midpoint for the contrast function.
Public Sub BrightnessContrast(ByVal Bright As Integer, ByVal Contrast As Single, Optional ByVal TrueContrast As Boolean = True)
    
    Dim BrightTable(0 To 255) As Byte
    Dim BTCalc As Long
    
    GetImageData
    
    'If the brightness value is anything but 0, process it
    If (Bright <> 0) Then
        
        Message "Building brightness look-up table..."
        
        For x = 0 To 255
            BTCalc = x + Bright
            If BTCalc > 255 Then BTCalc = 255
            If BTCalc < 0 Then BTCalc = 0
            BrightTable(x) = CByte(BTCalc)
        Next x
        
        Message "Adjusting image brightness..."

        Dim QuickX As Long
        
        'Because contrast and brightness are handled together, set the progress bar maximum value
        ' contingent on whether we're handling just brightness, or both
        If Contrast = 0 Then SetProgBarMax PicWidthL Else SetProgBarMax PicWidthL * 2
        
        For x = 0 To PicWidthL
            QuickX = x * 3
        For y = 0 To PicHeightL
            ImageData(QuickX + 2, y) = BrightTable(ImageData(QuickX + 2, y))
            ImageData(QuickX + 1, y) = BrightTable(ImageData(QuickX + 1, y))
            ImageData(QuickX, y) = BrightTable(ImageData(QuickX, y))
        Next y
            If x Mod 20 = 0 Then SetProgBarVal x
        Next x
            
        If Contrast = 0 Then SetImageData
            
    End If
    
    'If the contrast value is anything but 0, process it
    If (Contrast <> 0) Then
    
        'Sampled contrast is my invention; traditionally contrast pushes colors toward or away from gray.
        ' I like the option to push the colors toward or away from the image's actual midpoint, which
        ' may not be gray.  For most well-framed photos the difference is minimal, but for images with
        ' non-traditional white balance, sampled contrast offers better results.
        If (TrueContrast = True) Then
            
            Dim RTotal As Long, GTotal As Long, BTotal As Long
            
            Message "Sampling image data..."
            
            Dim Mean As Long
            For x = 0 To PicWidthL
                QuickX = x * 3
            For y = 0 To PicHeightL
                RTotal = RTotal + ImageData(QuickX + 2, y)
                GTotal = GTotal + ImageData(QuickX + 1, y)
                BTotal = BTotal + ImageData(QuickX, y)
            Next y
            Next x
            RTotal = RTotal \ (PicWidthL * PicHeightL)
            GTotal = GTotal \ (PicWidthL * PicHeightL)
            BTotal = BTotal \ (PicWidthL * PicHeightL)
            Mean = (RTotal + GTotal + BTotal) \ 3
        Else
            Mean = 128
        End If
        
        Dim ContrastTable(0 To 255) As Byte, CTCalc As Long
        
        Message "Building contrast look-up table..."
        
        For x = 0 To 255
            CTCalc = x + (((x - Mean) * Contrast) \ 100)
            If CTCalc > 255 Then CTCalc = 255
            If CTCalc < 0 Then CTCalc = 0
            ContrastTable(x) = CByte(CTCalc)
        Next x
        
        Message "Adjusting image contrast..."
        
        For x = 0 To PicWidthL
            QuickX = x * 3
        For y = 0 To PicHeightL
            ImageData(QuickX + 2, y) = ContrastTable(ImageData(QuickX + 2, y))
            ImageData(QuickX + 1, y) = ContrastTable(ImageData(QuickX + 1, y))
            ImageData(QuickX, y) = ContrastTable(ImageData(QuickX, y))
        Next y
            If x Mod 20 = 0 Then SetProgBarVal PicWidthL + x
        Next x
        
        SetImageData
        
    End If
    
    SetProgBarVal 0
    Message "Finished."

End Sub

Private Sub Form_Load()
    
    'Initialize the preview windows
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    
   'Create the preview
    DrawBCPreview hsBright.Value, hsContrast.Value
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me

End Sub

'Same deal as above, but sampling is removed.  (It's not particularly useful for a preview, anyway.)
Private Sub DrawBCPreview(ByVal Bright As Integer, ByVal Contrast As Single)

    Dim BrightTable(0 To 255) As Byte, BTCalc As Long
    Dim QuickX As Long
    
    GetPreviewData PicPreview
    
    For x = 0 To 255
        BTCalc = x + Bright
        If BTCalc > 255 Then BTCalc = 255
        If BTCalc < 0 Then BTCalc = 0
        BrightTable(x) = CByte(BTCalc)
    Next x
    
    For x = PreviewX To PreviewX + PreviewWidth
        QuickX = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        ImageData(QuickX + 2, y) = BrightTable(ImageData(QuickX + 2, y))
        ImageData(QuickX + 1, y) = BrightTable(ImageData(QuickX + 1, y))
        ImageData(QuickX, y) = BrightTable(ImageData(QuickX, y))
    Next y
    Next x
        
    Dim ContrastTable(0 To 255) As Byte, CTCalc As Long
    For x = 0 To 255
        CTCalc = x + (((x - 128) * Contrast) \ 100)
        If CTCalc > 255 Then CTCalc = 255
        If CTCalc < 0 Then CTCalc = 0
        ContrastTable(x) = CByte(CTCalc)
    Next x

    For x = PreviewX To PreviewX + PreviewWidth
        QuickX = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        ImageData(QuickX + 2, y) = ContrastTable(ImageData(QuickX + 2, y))
        ImageData(QuickX + 1, y) = ContrastTable(ImageData(QuickX + 1, y))
        ImageData(QuickX, y) = ContrastTable(ImageData(QuickX, y))
    Next y
    Next x
    
    SetPreviewData PicEffect
    
End Sub

'Everything below this line is related to updating the text boxes and scroll bars when one or the
' other is modified by the user.  When that happens, the preview window also gets updated.
Private Sub hsBright_Change()
    DrawBCPreview hsBright.Value, hsContrast.Value
    txtBrightness.Text = hsBright.Value
End Sub

Private Sub hsBright_Scroll()
    DrawBCPreview hsBright.Value, hsContrast.Value
    txtBrightness.Text = hsBright.Value
End Sub

Private Sub hsContrast_Change()
    DrawBCPreview hsBright.Value, hsContrast.Value
    txtContrast.Text = hsContrast.Value
End Sub

Private Sub hsContrast_Scroll()
    DrawBCPreview hsBright.Value, hsContrast.Value
    txtContrast.Text = hsContrast.Value
End Sub

Private Sub txtBrightness_Change()
    If EntryValid(txtBrightness, hsBright.Min, hsBright.Max, False, False) Then
        hsBright.Value = val(txtBrightness)
    End If
End Sub

Private Sub txtBrightness_GotFocus()
    AutoSelectText txtBrightness
End Sub

Private Sub txtContrast_Change()
    If EntryValid(txtContrast, hsContrast.Min, hsContrast.Max, False, False) Then
        hsContrast.Value = val(txtContrast)
    End If
End Sub

Private Sub txtContrast_GotFocus()
    AutoSelectText txtContrast
End Sub
