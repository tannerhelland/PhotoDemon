VERSION 5.00
Begin VB.Form FormSplitTone 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Split Toning"
   ClientHeight    =   6480
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
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   11
      Top             =   5730
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picSaturation 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   6375
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   15
      Top             =   4470
      Width           =   4335
   End
   Begin VB.PictureBox picSaturation 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   6375
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   14
      Top             =   2100
      Width           =   4335
   End
   Begin VB.PictureBox picHue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   6375
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   13
      Top             =   3285
      Width           =   4335
   End
   Begin VB.PictureBox picHue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   6375
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   12
      Top             =   915
      Width           =   4335
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
   Begin PhotoDemon.sliderTextCombo sltBalance 
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   5175
      Width           =   5895
      _ExtentX        =   10610
      _ExtentY        =   873
      Max             =   100
      Value           =   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltHue 
      Height          =   495
      Index           =   0
      Left            =   6000
      TabIndex        =   3
      Top             =   435
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   360
      Value           =   180
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltSaturation 
      Height          =   495
      Index           =   0
      Left            =   6000
      TabIndex        =   4
      Top             =   1620
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   100
      Value           =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltHue 
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   8
      Top             =   2805
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   360
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltSaturation 
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   9
      Top             =   3990
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   100
      Value           =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "shadow saturation:"
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
      Left            =   5880
      TabIndex        =   10
      Top             =   3645
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "shadow hue:"
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
      Left            =   5880
      TabIndex        =   7
      Top             =   2460
      Width           =   1365
   End
   Begin VB.Label lblHue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "highlight hue:"
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
      Left            =   5880
      TabIndex        =   6
      Top             =   90
      Width           =   1485
   End
   Begin VB.Label lblSaturation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "highlight saturation:"
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
      Left            =   5880
      TabIndex        =   5
      Top             =   1275
      Width           =   2145
   End
   Begin VB.Label lblStrength 
      AutoSize        =   -1  'True
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
      Left            =   5880
      TabIndex        =   0
      Top             =   4830
      Width           =   960
   End
End
Attribute VB_Name = "FormSplitTone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Split Toning Form
'Copyright ©2014 by Audioglider and Tanner Helland
'Created: 07/May/14
'Last updated: 08/May/14
'Last update: add preview boxes for highlight/shadow hue and saturation
'
'This technique applies a different tones to shadows and highlights in the image.  For a comprehensive explanation
' of split-toning (and its historical relevance), see
' http://www.alternativephotography.com/wp/toning/split-toning-history
'
'Many thanks to expert coder Audioglider for contributing this tool to PhotoDemon.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

' Input: hue and saturation tones for Shadows and Highlights; Balance defines whether shadows tone or
'        highlights tone gets larger share of tonal range to tint.
Public Sub SplitTone(ByVal highlightHue As Double, ByVal highlightSaturation As Double, ByVal shadowHue As Double, ByVal shadowSaturation As Double, ByVal Balance As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Split toning image..."
    
    'Convert the modifiers to be on the same scale as the HSV translation routine
    highlightHue = highlightHue / 360
    highlightSaturation = highlightSaturation / 100
    shadowHue = shadowHue / 360
    shadowSaturation = shadowSaturation / 100
    
    'Convert balance mix value to [0,1]; it will be used to blend split-toned colors with their original RGB values
    Balance = Math_Functions.convertRange(0, 100, 0, 1, Balance)
    
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
    Dim newR As Long, newG As Long, newB As Long
    Dim v As Double
    
    Dim rHighlight As Double, gHighlight As Double, bHighlight As Double
    Dim rShadow As Double, gShadow As Double, bShadow As Double
    
    'To accelerate the main loop, use lookup tables for grayscale conversion
    Dim averageLookup(0 To 765) As Long, averageLookupFloat(0 To 255) As Single
    For x = 0 To 765
        averageLookup(x) = x \ 3
    Next x
    
    For x = 0 To 255
        averageLookupFloat(x) = x / 255
    Next x
    
    Dim average As Double, averageFloat As Double, invAverageFloat As Double
    Dim redToned As Double, greenToned As Double, blueToned As Double
    Dim rDiff As Double, gDiff As Double, bDiff As Double
    
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate luminance for this pixel
        v = CDbl(getLuminance(r, g, b)) / 255
        
        'Retrieve RGB conversions for the supplied highlight and shadow values
        fHSVtoRGB highlightHue, highlightSaturation, v, rHighlight, gHighlight, bHighlight
        fHSVtoRGB shadowHue, shadowSaturation, v, rShadow, gShadow, bShadow
        
        'Highlight and shadow values are returned in the range [0, 1]; convert them to [0, 255] before continuing
        rHighlight = rHighlight * 255
        rShadow = rShadow * 255
        gHighlight = gHighlight * 255
        gShadow = gShadow * 255
        bHighlight = bHighlight * 255
        bShadow = bShadow * 255
        
        'Calculate average values
        average = averageLookup(r + g + b)
        averageFloat = averageLookupFloat(average)
        invAverageFloat = 1 - averageFloat
        
        'Tone the specified colors
        redToned = (rHighlight * averageFloat) + (rShadow * (invAverageFloat))
        greenToned = (gHighlight * averageFloat) + (gShadow * (invAverageFloat))
        blueToned = (bHighlight * averageFloat) + (bShadow * (invAverageFloat))
        
        rDiff = redToned - average
        gDiff = greenToned - average
        bDiff = blueToned - average
        
        'Calculate final RGB values by splitting the difference beetween the toned values
        ' and this pixel's average tone
        newR = (average + rDiff - (gDiff / 2) - (bDiff / 2))
        newG = (average + gDiff - (rDiff / 2) - (bDiff / 2))
        newB = (average + bDiff - (rDiff / 2) - (gDiff / 2))
        
        If newR > 255 Then newR = 255
        If newR < 0 Then newR = 0
        If newG > 255 Then newG = 255
        If newG < 0 Then newG = 0
        If newB > 255 Then newB = 255
        If newB < 0 Then newB = 0
        
        'As a final step, apply the Balance parameter to the colors; effectively, this just mixes them
        ' with their original RGB values at the strength requested by the user.
        ImageData(QuickVal + 2, y) = BlendColors(r, newR, Balance)
        ImageData(QuickVal + 1, y) = BlendColors(g, newG, Balance)
        ImageData(QuickVal, y) = BlendColors(b, newB, Balance)
        
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

Private Sub cmdBar_OKClick()
    Process "Split toning", , buildParams(sltHue(0).Value, sltSaturation(0).Value, sltHue(1).Value, sltSaturation(1).Value, sltBalance.Value)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

'To help orient the user, slightly different reset values are used for this tool.
Private Sub cmdBar_ResetClick()
    sltHue(0).Value = 180
    sltSaturation(0).Value = 100
    sltSaturation(1).Value = 100
    sltBalance.Value = 50
End Sub

Private Sub Form_Activate()
    
    'Draw the hue and saturation preview boxes
    Dim i As Long
    For i = 0 To 1
        Drawing.drawHueBox_HSV picHue(i)
        Drawing.drawSaturationBox_HSV picSaturation(i), sltHue(i).Value / 360
    Next i
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltBalance_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then SplitTone sltHue(0), sltSaturation(0), sltHue(1), sltSaturation(1), sltBalance, True, fxPreview
End Sub

Private Sub sltHue_Change(Index As Integer)
    Drawing.drawSaturationBox_HSV picSaturation(Index), sltHue(Index).Value / 360
    updatePreview
End Sub

Private Sub sltSaturation_Change(Index As Integer)
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

