VERSION 5.00
Begin VB.Form FormBrightnessContrast 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Brightness/Contrast"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
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
   ScaleHeight     =   430
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
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
      Height          =   2730
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   10
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
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
      Height          =   2730
      Left            =   3240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsContrast 
      Height          =   255
      Left            =   360
      Max             =   100
      Min             =   -100
      TabIndex        =   4
      Top             =   4680
      Width           =   5055
   End
   Begin VB.HScrollBar hsBright 
      Height          =   255
      Left            =   360
      Max             =   255
      Min             =   -255
      TabIndex        =   2
      Top             =   3840
      Width           =   5055
   End
   Begin VB.TextBox txtContrast 
      Alignment       =   2  'Center
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
      Left            =   5520
      TabIndex        =   5
      Text            =   "0"
      Top             =   4650
      Width           =   495
   End
   Begin VB.TextBox txtBrightness 
      Alignment       =   2  'Center
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
      Left            =   5520
      TabIndex        =   3
      Text            =   "0"
      Top             =   3810
      Width           =   495
   End
   Begin VB.CheckBox chkSample 
      Appearance      =   0  'Flat
      Caption         =   "Sample image for true contrast (slower but more accurate)"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   5280
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   6000
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   6000
      Width           =   1125
   End
   Begin VB.Label lblBeforeandAfter 
      BackStyle       =   0  'Transparent
      Caption         =   "  Before                                                           After"
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
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   3975
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
      Left            =   240
      TabIndex        =   8
      Top             =   4335
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
      Left            =   240
      TabIndex        =   7
      Top             =   3480
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
'Last updated: 09/September/12
'Last update: better optimized sampled contrast while previewing.  The sampled mean is now calculated only once, then
'              stored.  This way it doesn't have to be resampled every time the scroll bar is moved.
'
'The wonderful brightness/contrast handler.  Everything is done via look-up
' tables, so it's extremely fast.  It's all linear (not logarithmic; sorry).
' Maybe someday I'll change that, maybe not... honestly, I probably won't, since
' brightness and contrast are such stupid functions anyway.  People should be
' using levels or curves or white balance instead!
'
'***************************************************************************

Option Explicit

'While previewing, we don't need to repeatedly sample contrast.  Just do it once and store the value.
Dim previewHasSampled As Boolean
Dim previewSampledContrast As Long

'Update the preview when the "sample contrast" checkbox value is changed
Private Sub chkSample_Click()
    If chkSample.Value = vbChecked Then BrightnessContrast hsBright.Value, hsContrast.Value, True, True, picEffect Else BrightnessContrast hsBright.Value, hsContrast.Value, False, True, picEffect
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()

    'Check the text box values against the limits of their corresponding scroll bars - that'll catch
    ' any out-of-range errors
    If EntryValid(txtBrightness, hsBright.Min, hsBright.Max) Then
        
        If EntryValid(txtContrast, hsContrast.Min, hsContrast.Max) Then
        
            Me.Visible = False
            
            'Re-route the effect through the software processor, so it can be tracked
            If chkSample.Value = vbChecked Then
                Process BrightnessAndContrast, hsBright.Value, hsContrast.Value, True
            Else
                Process BrightnessAndContrast, hsBright.Value, hsContrast.Value, False
            End If
            
            Unload Me
            
        Else
            AutoSelectText txtContrast
        End If
    
    Else
        AutoSelectText txtBrightness
    End If
    
End Sub

'Single routine for modifying both brightness and contrast.  Brightness is in the range (-255,255) while
' contrast is (-100,100).  Optionally, the image can be sampled to obtain a true midpoint for the contrast function.
Public Sub BrightnessContrast(ByVal Bright As Long, ByVal Contrast As Single, Optional ByVal TrueContrast As Boolean = True, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Adjusting image brightness..."
    
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
    
    'If the brightness value is anything but 0, process it
    If (Bright <> 0) Then
        
        If toPreview = False Then
        
            Message "Adjusting image brightness..."
        
            'Because contrast and brightness are handled together, set the progress bar maximum value
            ' contingent on whether we're handling just brightness, or both brightness AND contrast.
            If Contrast <> 0 Then
                SetProgBarMax finalX * 2
                progBarCheck = findBestProgBarValue()
            End If
            
        End If
        
        'Look-up tables work brilliantly for brightness
        Dim BrightTable(0 To 255) As Byte
        Dim BTCalc As Long
        
        For x = 0 To 255
            BTCalc = x + Bright
            If BTCalc > 255 Then BTCalc = 255
            If BTCalc < 0 Then BTCalc = 0
            BrightTable(x) = CByte(BTCalc)
        Next x
        
        'Loop through each pixel in the image, converting values as we go
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
            
            'Use the look-up table to perform an ultra-quick brightness adjustment
            ImageData(QuickVal, y) = BrightTable(ImageData(QuickVal, y))
            ImageData(QuickVal + 1, y) = BrightTable(ImageData(QuickVal + 1, y))
            ImageData(QuickVal + 2, y) = BrightTable(ImageData(QuickVal + 2, y))
            
        Next y
            If toPreview = False Then
                If (x And progBarCheck) = 0 Then SetProgBarVal x
            End If
        Next x
        
    End If
    
    'If the contrast value is anything but 0, process it
    If (Contrast <> 0) Then
    
        'Contrast requires an average value to operate correctly; it works by pushing luminance values away from that average.
        Dim Mean As Long
    
        'Sampled contrast is my invention; traditionally contrast pushes colors toward or away from gray.
        ' I like the option to push the colors toward or away from the image's actual midpoint, which
        ' may not be gray.  For most white-balanced photos the difference is minimal, but for images with
        ' non-traditional white balance, sampled contrast offers better results.
        If TrueContrast Then
        
            If toPreview And previewHasSampled Then
            
                Mean = previewSampledContrast
            
            Else
            
                If toPreview = False Then Message "Sampling image data to determine true contrast..."
                
                Dim RTotal As Long, GTotal As Long, BTotal As Long
                RTotal = 0
                GTotal = 0
                BTotal = 0
                
                Dim NumOfPixels As Long
                NumOfPixels = 0
                
                For x = initX To finalX
                    QuickVal = x * qvDepth
                For y = initY To finalY
                    RTotal = RTotal + ImageData(QuickVal + 2, y)
                    GTotal = GTotal + ImageData(QuickVal + 1, y)
                    BTotal = BTotal + ImageData(QuickVal, y)
                    NumOfPixels = NumOfPixels + 1
                Next y
                Next x
                
                RTotal = RTotal \ NumOfPixels
                GTotal = GTotal \ NumOfPixels
                BTotal = BTotal \ NumOfPixels
                
                Mean = (RTotal + GTotal + BTotal) \ 3
                
                If toPreview Then
                    previewSampledContrast = Mean
                    previewHasSampled = True
                End If
            
            End If
                
        'If we're not using true contrast, set the mean to the traditional 127
        Else
            Mean = 127
        End If
            
        
        If toPreview = False Then Message "Adjusting image contrast..."
        
        'Like brightness, contrast works beautifully with look-up tables
        Dim ContrastTable(0 To 255) As Byte, CTCalc As Long
                
        For x = 0 To 255
            CTCalc = x + (((x - Mean) * Contrast) \ 100)
            If CTCalc > 255 Then CTCalc = 255
            If CTCalc < 0 Then CTCalc = 0
            ContrastTable(x) = CByte(CTCalc)
        Next x
        
        'Loop through each pixel in the image, converting values as we go
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
            
            'Use the look-up table to perform an ultra-quick brightness adjustment
            ImageData(QuickVal, y) = ContrastTable(ImageData(QuickVal, y))
            ImageData(QuickVal + 1, y) = ContrastTable(ImageData(QuickVal + 1, y))
            ImageData(QuickVal + 2, y) = ContrastTable(ImageData(QuickVal + 2, y))
            
        Next y
            If toPreview = False Then
                If (x And progBarCheck) = 0 Then
                    If Bright <> 0 Then SetProgBarVal x + finalX Else SetProgBarVal x
                End If
            End If
        Next x
        
    End If
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic

End Sub

'LOAD form
Private Sub Form_Load()
    
    'Initialize the preview windows
    DrawPreviewImage picPreview
    
    previewHasSampled = 0
    previewSampledContrast = 0
    
   'Create the preview
    BrightnessContrast hsBright.Value, hsContrast.Value, False, True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me

End Sub

'Everything below this line is related to updating the text boxes and scroll bars when one or the
' other is modified by the user.  When that happens, the preview window also gets updated.
Private Sub hsBright_Change()
    If chkSample.Value = vbChecked Then BrightnessContrast hsBright.Value, hsContrast.Value, True, True, picEffect Else BrightnessContrast hsBright.Value, hsContrast.Value, False, True, picEffect
    txtBrightness.Text = hsBright.Value
End Sub

Private Sub hsBright_Scroll()
    If chkSample.Value = vbChecked Then BrightnessContrast hsBright.Value, hsContrast.Value, True, True, picEffect Else BrightnessContrast hsBright.Value, hsContrast.Value, False, True, picEffect
    txtBrightness.Text = hsBright.Value
End Sub

Private Sub hsContrast_Change()
    If chkSample.Value = vbChecked Then BrightnessContrast hsBright.Value, hsContrast.Value, True, True, picEffect Else BrightnessContrast hsBright.Value, hsContrast.Value, False, True, picEffect
    txtContrast.Text = hsContrast.Value
End Sub

Private Sub hsContrast_Scroll()
    If chkSample.Value = vbChecked Then BrightnessContrast hsBright.Value, hsContrast.Value, True, True, picEffect Else BrightnessContrast hsBright.Value, hsContrast.Value, False, True, picEffect
    txtContrast.Text = hsContrast.Value
End Sub

Private Sub txtBrightness_Change()
    If EntryValid(txtBrightness, hsBright.Min, hsBright.Max, False, False) Then hsBright.Value = val(txtBrightness)
End Sub

Private Sub txtBrightness_GotFocus()
    AutoSelectText txtBrightness
End Sub

Private Sub txtContrast_Change()
    If EntryValid(txtContrast, hsContrast.Min, hsContrast.Max, False, False) Then hsContrast.Value = val(txtContrast)
End Sub

Private Sub txtContrast_GotFocus()
    AutoSelectText txtContrast
End Sub
