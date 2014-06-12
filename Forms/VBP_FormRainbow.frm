VERSION 5.00
Begin VB.Form FormRainbow 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Rainbow"
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
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
   Begin PhotoDemon.sliderTextCombo sltOffset 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   1170
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   359
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
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   2160
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
   Begin PhotoDemon.sliderTextCombo sltStrength 
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   3240
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
   Begin PhotoDemon.sliderTextCombo sltSaturation 
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   4320
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "saturation boost:"
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
      TabIndex        =   9
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Label lblTitle 
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
      Index           =   2
      Left            =   6000
      TabIndex        =   3
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "angle:"
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
      TabIndex        =   2
      Top             =   1800
      Width           =   660
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "offset:"
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
      TabIndex        =   1
      Top             =   840
      Width           =   675
   End
End
Attribute VB_Name = "FormRainbow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Rainbow Effect dialog
'Copyright ©2003-2014 by Tanner Helland
'Created: sometime 2003
'Last updated: 11/June/14
'Last update: moved the function to its own dialog
'
'Fun Rainbow effect for an image.  Options should be self-explanatory.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Apply a rainbow overlay to an image
Public Sub ApplyRainbowEffect(ByVal hueOffset As Double, ByVal rainbowAngle As Double, ByVal rainbowStrength As Double, ByVal saturationBoost As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Sprinkling image with shimmering rainbows..."
    
    'Convert the hue modifier to the [0, 6] range
    hueOffset = hueOffset / 360
    
    'Convert strength from [0, 100] to [0, 1]
    rainbowStrength = rainbowStrength / 100
    
    'Convert saturation boosting from [0, 100] to [0, 1]
    saturationBoost = saturationBoost / 100
    
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
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * 0.5
    midX = midX + initX
    midY = CDbl(finalY - initY) * 0.5
    midY = midY + initY
    
    'Convert the rotation angle to radians
    rainbowAngle = rainbowAngle * (PI / 180)
    
    'Find the cos and sin of this angle and store the values
    Dim cosTheta As Double, sinTheta As Double
    cosTheta = Cos(rainbowAngle)
    sinTheta = Sin(rainbowAngle)
    
    'Using those values, build 4 lookup tables, one each for x/y times sin/cos
    Dim xSin() As Double, xCos() As Double
    ReDim xSin(initX To finalX) As Double
    ReDim xCos(initX To finalX) As Double
    
    For x = initX To finalX
        xSin(x) = (x - midX) * sinTheta + midY
        xCos(x) = (x - midX) * cosTheta + midX
    Next
    
    Dim ySin() As Double, yCos() As Double
    ReDim ySin(initY To finalY) As Double
    ReDim yCos(initY To finalY) As Double
    For y = initY To finalY
        ySin(y) = (y - midY) * sinTheta
        yCos(y) = (y - midY) * cosTheta
    Next y
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which are used to solve for the hue of a given point
    Dim srcX As Double, srcY As Double
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim rFloat As Double, gFloat As Double, bFloat As Double
    Dim h As Double, s As Double, l As Double
    Dim hVal As Double, xDistance As Double
    xDistance = finalX - initX
    
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Get red, green, and blue values from the array
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
                
        'Convert the RGB values the HSV space
        fRGBtoHSV r / 255, g / 255, b / 255, h, s, l
        
        'Solve for the original (x) position of this pixel in the image, accounting for rotation
        srcX = xCos(x) - ySin(y)
        
        'Based on the x-coordinate of a pixel, apply a predetermined hue gradient (stretching between -1 and 5)
        hVal = srcX / xDistance
        
        'Apply the hue offset
        hVal = hVal + hueOffset
        If hVal > 1 Then hVal = hVal - 1
        
        'Apply saturation boosting, if any
        If saturationBoost > 0 Then s = 1 * saturationBoost + (s * (1 - saturationBoost))
        
        'Now convert those HSL values back to RGB, but substitute in our artificial hue (and possibly
        ' saturation) value(s)
        fHSVtoRGB hVal, s, l, rFloat, gFloat, bFloat
        
        'Blend the original and new RGB values according to the requested strength
        r = BlendColors(r, rFloat * 255, rainbowStrength)
        g = BlendColors(g, gFloat * 255, rainbowStrength)
        b = BlendColors(b, bFloat * 255, rainbowStrength)
        
        'Assign the new RGB values back into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If Not toPreview Then
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
    Process "Rainbow", , buildParams(sltOffset.Value, sltAngle.Value, sltStrength.Value, sltSaturation.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltStrength.Value = 100
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

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyRainbowEffect sltOffset.Value, sltAngle.Value, sltStrength.Value, sltSaturation.Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

Private Sub sltOffset_Change()
    updatePreview
End Sub

Private Sub sltSaturation_Change()
    updatePreview
End Sub

Private Sub sltStrength_Change()
    updatePreview
End Sub
