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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.smartCheckBox chkUnison 
      Height          =   330
      Left            =   6120
      TabIndex        =   4
      Top             =   5280
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   582
      Caption         =   "keep all colors in sync"
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
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltGamma 
      Height          =   720
      Index           =   0
      Left            =   6000
      TabIndex        =   5
      Top             =   2640
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "red"
      Min             =   0.01
      Max             =   3
      SigDigits       =   2
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
   Begin PhotoDemon.sliderTextCombo sltGamma 
      Height          =   720
      Index           =   1
      Left            =   6000
      TabIndex        =   6
      Top             =   3540
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "green"
      Min             =   0.01
      Max             =   3
      SigDigits       =   2
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
   Begin PhotoDemon.sliderTextCombo sltGamma 
      Height          =   720
      Index           =   2
      Left            =   6000
      TabIndex        =   7
      Top             =   4440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "blue"
      Min             =   0.01
      Max             =   3
      SigDigits       =   2
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
   Begin VB.Label lblTitle 
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
      Height          =   1005
      Index           =   2
      Left            =   6000
      TabIndex        =   3
      Top             =   1170
      Width           =   2040
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormGamma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gamma Correction Handler
'Copyright 2000-2015 by Tanner Helland
'Created: 12/May/01
'Last updated: 23/April/13
'Last update: replaced all scroll bars and text boxes with my new combo text/scroll control.  Floating-point entry is
'              now much easier to deal with.  Also, added divide-by-zero checks to the main function, just in case.
'
'Updated version of the gamma handler; fully optimized, it uses a look-up
' table and can correct any color channel.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Dim userChange As Boolean

Private Sub chkUnison_Click()
    
    If CBool(chkUnison) Then
        Dim newGamma As Double
        newGamma = CDblCustom(sltGamma(0) + sltGamma(1) + sltGamma(2)) / 3
    
        userChange = False
        sltGamma(0) = newGamma
        sltGamma(1) = newGamma
        sltGamma(2) = newGamma
        userChange = True
    End If
    
    updatePreview
    
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Gamma", , buildParams(sltGamma(0), sltGamma(1), sltGamma(2)), UNDO_LAYER
End Sub

'When randomizing, do not check the "unison" box
Private Sub cmdBar_RandomizeClick()
    chkUnison.Value = vbUnchecked
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltGamma(0).Value = 1
    sltGamma(1).Value = 1
    sltGamma(2).Value = 1
End Sub

Private Sub Form_Activate()
    
    userChange = True
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Finally, render a preview
    updatePreview
    
End Sub

'Basic gamma correction.  It's a simple function - use an exponent to adjust R/G/B values.
' Inputs: new gamma level, which channels to adjust (r/g/b/all), and optional preview information
Public Sub GammaCorrect(ByVal rGamma As Double, ByVal gGamma As Double, ByVal bGamma As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
     
    If Not toPreview Then Message "Adjusting gamma values..."
    
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
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Make certain that the gamma adjustment values we were passed are not zero
    If rGamma = 0 Then rGamma = 0.01
    If gGamma = 0 Then gGamma = 0.01
    If bGamma = 0 Then bGamma = 0.01
    
    'Gamma can be easily applied using a look-up table
    Dim gLookUp(0 To 2, 0 To 255) As Byte
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
        
        gLookUp(y, x) = tmpVal
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
        ImageData(QuickVal + 2, y) = gLookUp(0, r)
        ImageData(QuickVal + 1, y) = gLookUp(1, g)
        ImageData(QuickVal, y) = gLookUp(2, b)
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
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

'Redraw the preview effect and the gamma chart
Private Sub updatePreview()

    If cmdBar.previewsAllowed Then
    
        Dim prevX As Double, prevY As Double
        Dim curX As Double, curY As Double
        Dim x As Long, y As Long
        
        Dim xWidth As Long, yHeight As Long
        xWidth = picChart.ScaleWidth
        yHeight = picChart.ScaleHeight
            
        'Clear out the old chart and draw a gray line across the diagonal for reference
        picChart.Picture = LoadPicture("")
        picChart.ForeColor = RGB(127, 127, 127)
        GDIPlusDrawLineToDC picChart.hDC, 0, yHeight, xWidth, 0, RGB(127, 127, 127)
        
        Dim gamVal As Double, tmpVal As Double
        
        'Draw each of the current gamma curves for the user's reference
        For y = 0 To 2
            
            'If all channels are in sync, draw only blue; otherwise, color each channel individually
            gamVal = sltGamma(y)
            If (sltGamma(0) = sltGamma(1)) And (sltGamma(1) = sltGamma(2)) Then
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
                GDIPlusDrawLineToDC picChart.hDC, prevX, prevY, curX, curY, picChart.ForeColor
                prevX = curX
                prevY = curY
            Next x
            
        Next y
        
        picChart.Picture = picChart.Image
        picChart.Refresh
    
        'Once the chart is done, redraw the gamma preview as well
        GammaCorrect sltGamma(0), sltGamma(1), sltGamma(2), True, fxPreview
        
    End If
    
End Sub

Private Sub sltGamma_Change(Index As Integer)

    If userChange And cmdBar.previewsAllowed Then
        userChange = False
        
        If CBool(chkUnison) Then
            Select Case Index
                Case 0
                    sltGamma(1).Value = sltGamma(0).Value
                    sltGamma(2).Value = sltGamma(0).Value
                Case 1
                    sltGamma(0).Value = sltGamma(1).Value
                    sltGamma(2).Value = sltGamma(1).Value
                Case 2
                    sltGamma(0).Value = sltGamma(2).Value
                    sltGamma(1).Value = sltGamma(2).Value
            End Select
        End If
        
        userChange = True
        
        updatePreview
    End If

End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


