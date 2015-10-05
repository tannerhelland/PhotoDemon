VERSION 5.00
Begin VB.Form FormBrightnessContrast 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Brightness/Contrast"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12075
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
   ScaleWidth      =   805
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.smartCheckBox chkSample 
      Height          =   330
      Left            =   6120
      TabIndex        =   3
      Top             =   3600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   582
      Caption         =   "sample image for true contrast (slower but more accurate)"
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltBright 
      Height          =   720
      Left            =   6000
      TabIndex        =   1
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "brightness"
      Min             =   -255
      Max             =   255
      Value           =   -10
   End
   Begin PhotoDemon.sliderTextCombo sltContrast 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   2760
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "contrast"
      Min             =   -100
      Max             =   100
      Value           =   10
   End
End
Attribute VB_Name = "FormBrightnessContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Brightness and Contrast Handler
'Copyright 2001-2015 by Tanner Helland
'Created: 2/6/01
'Last updated: 16/August/13
'Last update: this dialog is my testbed for the new command bar user control, so it received a number of changes
'              relating to proper command bar implementation.
'
'The central brightness/contrast handler.  Everything is done via look-up tables, so it's extremely fast.
' It's all linear (not logarithmic; sorry). Maybe someday I'll change that, maybe not... honestly, I probably
' won't, since brightness and contrast are such stupid functions anyway.  People should be using levels or
' curves or white balance instead!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'While previewing, we don't need to repeatedly sample contrast.  Just do it once and store the value.
Private previewHasSampled As Boolean
Private previewSampledContrast As Long

'Update the preview when the "sample contrast" checkbox value is changed
Private Sub chkSample_Click()
    updatePreview
End Sub

'Single routine for modifying both brightness and contrast.  Brightness is in the range (-255,255) while
' contrast is (-100,100).  Optionally, the image can be sampled to obtain a true midpoint for the contrast function.
Public Sub BrightnessContrast(ByVal Bright As Long, ByVal Contrast As Double, Optional ByVal TrueContrast As Boolean = True, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Adjusting image brightness..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'If the brightness value is anything but 0, process it
    If (Bright <> 0) Then
        
        If Not toPreview Then
        
            Message "Adjusting image brightness..."
        
            'Because contrast and brightness are handled together, set the progress bar maximum value
            ' contingent on whether we're handling just brightness, or both brightness AND contrast.
            If (Contrast <> 0) Then
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
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal x
                End If
            End If
        Next x
        
    End If
    
    'If the contrast value is anything but 0, process it
    If (Contrast <> 0) And (Not cancelCurrentAction) Then
    
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
                
                Dim rTotal As Long, gTotal As Long, bTotal As Long
                rTotal = 0
                gTotal = 0
                bTotal = 0
                
                Dim NumOfPixels As Long
                NumOfPixels = 0
                
                For x = initX To finalX
                    QuickVal = x * qvDepth
                For y = initY To finalY
                    rTotal = rTotal + ImageData(QuickVal + 2, y)
                    gTotal = gTotal + ImageData(QuickVal + 1, y)
                    bTotal = bTotal + ImageData(QuickVal, y)
                    NumOfPixels = NumOfPixels + 1
                Next y
                Next x
                
                rTotal = rTotal \ NumOfPixels
                gTotal = gTotal \ NumOfPixels
                bTotal = bTotal \ NumOfPixels
                
                Mean = (rTotal + gTotal + bTotal) \ 3
                
                If toPreview Then
                    previewSampledContrast = Mean
                    previewHasSampled = True
                End If
            
            End If
                
        'If we're not using true contrast, set the mean to the traditional 127
        Else
            Mean = 127
        End If
            
        
        If Not toPreview Then Message "Adjusting image contrast..."
        
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
                    If userPressedESC() Then Exit For
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

'OK button.  Note that the command bar class handles validation, form hiding, and form unload for us.
Private Sub cmdBar_OKClick()
    Process "Brightness and contrast", , buildParams(sltBright, sltContrast, CBool(chkSample.Value)), UNDO_LAYER
End Sub

'Sometimes the command bar will perform actions (like loading a preset) that require an updated preview.  This function
' is fired by the control when it's ready for such an update.
Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

'RESET button.  All control default values will be reset according to the rules specified in the commandBar user control
' source.  If we want a different default value applied, we can specify that here.  The important thing to note is
' that THE VALUES VISIBLE IN THE IDE DESIGNER DO NOT MATTER.
Private Sub cmdBar_ResetClick()
    
End Sub

Private Sub Form_Activate()
    
    previewHasSampled = 0
    previewSampledContrast = 0
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltBright_Change()
    updatePreview
End Sub

Private Sub sltContrast_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then BrightnessContrast sltBright, sltContrast, CBool(chkSample.Value), True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


