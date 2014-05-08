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
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10610
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
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
      Top             =   480
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
      Index           =   0
      Left            =   6000
      TabIndex        =   4
      Top             =   1320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   100
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
      Top             =   3840
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
      Top             =   4680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   100
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "shadows saturation:"
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
      Left            =   6000
      TabIndex        =   10
      Top             =   4320
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "shadows hue:"
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
      Left            =   6000
      TabIndex        =   7
      Top             =   3480
      Width           =   1470
   End
   Begin VB.Label lblHue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "highlights hue:"
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
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   1590
   End
   Begin VB.Label lblSaturation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " highlights saturation:"
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
      Top             =   960
      Width           =   2325
   End
   Begin VB.Label lblBalance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "balance:"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   2160
      Width           =   885
   End
End
Attribute VB_Name = "FormSplitTone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Split Toning Form
'Copyright ©2014 by Audioglider
'Created: 07/May/14
'Last updated: 07/May/14
'Last update: Initial build.
'
'This technique applies a different colored tone for shadows and
' highlights in the image.
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
    
    If toPreview = False Then Message "Split toning image..."
    
    'Convert the modifiers to be on the same scale as the HSV translation routine
    highlightHue = highlightHue / 360
    highlightSaturation = highlightSaturation / 100
    shadowHue = shadowHue / 360
    shadowSaturation = shadowSaturation / 100
    
    'Convert balance mix value to [0,1]
    Balance = convertRange(-100, 100, 0, 1, Balance)
    
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
    
    Dim average As Double
    Dim redToned As Double, greenToned As Double, blueToned As Double
    Dim rDiff As Double, gDiff As Double, bDiff As Double
    
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        v = CDbl(getLuminance(r, g, b)) / 255
        
        fHSVtoRGB highlightHue, highlightSaturation, v, rHighlight, gHighlight, bHighlight
        fHSVtoRGB shadowHue, shadowSaturation, v, rShadow, gShadow, bShadow
            
        average = (r + g + b) / 3
        redToned = (rHighlight * average) + (rShadow * (1 - average))
        greenToned = (gHighlight * average) + (gShadow * (1 - average))
        blueToned = (bHighlight * average) + (bShadow * (1 - average))
            
        rDiff = redToned - average
        gDiff = greenToned - average
        bDiff = blueToned - average
        
        newR = (average + rDiff - (gDiff / 2) - (bDiff / 2))
        newG = (average + gDiff - (rDiff / 2) - (bDiff / 2))
        newB = (average + bDiff - (rDiff / 2) - (gDiff / 2))
        
        If newR > 255 Then newR = 255
        If newR < 0 Then newR = 0
        If newG > 255 Then newG = 255
        If newG < 0 Then newG = 0
        If newB > 255 Then newB = 255
        If newB < 0 Then newB = 0
        
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
    Process "Split tone", , buildParams(sltHue(0).Value, sltSaturation(0).Value, sltHue(1).Value, sltSaturation(1).Value, sltBalance.Value)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
        
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
    updatePreview
End Sub

Private Sub sltSaturation_Change(Index As Integer)
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Function convertRange(ByVal originalStart As Double, ByVal originalEnd As Double, ByVal newStart As Double, ByVal newEnd As Double, ByVal Value As Double) As Double
    Dim dScale As Double
    
    dScale = (newEnd - newStart) / (originalEnd - originalStart)
    convertRange = (newStart + ((Value - originalStart) * dScale))
End Function

