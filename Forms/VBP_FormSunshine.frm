VERSION 5.00
Begin VB.Form FormSunshine 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Sunshine"
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
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   2010
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   200
      Value           =   72
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
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.sliderTextCombo sltRayCount 
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   360
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
   Begin PhotoDemon.sliderTextCombo sltXCenter 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
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
   Begin PhotoDemon.sliderTextCombo sltYCenter 
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
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
      TabIndex        =   9
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
   Begin PhotoDemon.colorSelector cpShine 
      Height          =   615
      Left            =   6120
      TabIndex        =   10
      Top             =   4080
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1085
      curColor        =   65535
   End
   Begin VB.Label lblShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "shine color:"
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
      TabIndex        =   11
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label lblExplanation 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: you can also set a center position by clicking the preview window."
      ForeColor       =   &H00404040&
      Height          =   435
      Index           =   0
      Left            =   6120
      TabIndex        =   8
      Top             =   1050
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "center position (x, y)"
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
      Index           =   4
      Left            =   6000
      TabIndex        =   7
      Top             =   120
      Width           =   2205
   End
   Begin VB.Label lblSaturation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "number of rays:"
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
      TabIndex        =   1
      Top             =   2640
      Width           =   1710
   End
   Begin VB.Label lblHue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "radius:"
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
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "FormSunshine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Sunshine Effect Form
'Copyright ©2014 by Audioglider
'Created: 30/May/14
'Last updated: 30/May/14
'Last update: Initial build
'
'This filter simulates the sun by generating a starburst effect. The X, Y
' coordinates sets the center of the burst, the Radius adjusts the size of
' of the center and the # of rays changes the the amount of rays of light
' that emanate from the center. All pretty self-explanatory :P
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Returns the largest integer not greater than x.
Private Function Floor(ByVal x As Double) As Long
    Floor = (-Int(x) * (-1))
End Function

Private Function GetGauss() As Double
    
    Dim sum As Double
    Dim i As Long
    
    Randomize Timer
    
    sum = 0
    For i = 0 To 5
        sum = sum + Rnd()
    Next i
    GetGauss = sum / 6
    
End Function

Public Sub SunShine(ByVal lRadius As Long, ByVal lSpokeCount As Long, ByVal lSpokeColor As Long, Optional ByVal centerX As Double = 0.1, Optional ByVal centerY As Double = 0.1, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Generating rays of happiness..."
    
    Dim i As Long
    Dim m_Radius As Double
    Dim m_Count As Long
    Dim m_Spoke() As Double
    Dim m_SpokeColor() As RGBQUAD
    
    m_Radius = lRadius
    m_Count = lSpokeCount
    
    ReDim m_Spoke(0 To m_Count - 1)
    ReDim m_SpokeColor(0 To m_Count - 1)
    For i = 0 To m_Count - 1
        m_Spoke(i) = GetGauss
        m_SpokeColor(i).Red = ExtractR(lSpokeColor)
        m_SpokeColor(i).Green = ExtractG(lSpokeColor)
        m_SpokeColor(i).Blue = ExtractB(lSpokeColor)
    Next i
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'If this is a preview, we need to adjust the radius values to match the size of the preview box
    If toPreview Then
        m_Radius = m_Radius * curDIBValues.previewModifier
    End If
    
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
    midX = CDbl(finalX - initX) * centerX
    midX = midX + initX
    midY = CDbl(finalY - initY) * centerY
    midY = midY + initY
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim newR As Double, newG As Double, newB As Double
    
    Dim u As Double, v As Double, t As Double
    Dim w As Double, w1 As Double, ws As Double, fRatio As Double
    Dim spokeCol As RGBQUAD
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
    
        u = (x - midX + 0.001) / m_Radius
        v = (y - midY + 0.001) / m_Radius
        
        t = (Atan2(u, v) / (2 * PI) + 0.51) * m_Count
        i = Floor(t)
        t = t - i
        i = i Mod m_Count
        
        w1 = m_Spoke(i) * (1 - t) + m_Spoke((i + 1) Mod m_Count) * t
        w1 = w1 * w1
        
        w = 1# / Sqr(u * u + v * v) * 0.9
        fRatio = fClamp(w, 0, 1)
        
        ws = fClamp(w1 * w, 0, 1)
        
        spokeCol.Red = m_SpokeColor(i).Red / 255 * (1 - t) + m_SpokeColor((i + 1) Mod m_Count).Red / 255 * t
        spokeCol.Green = m_SpokeColor(i).Green / 255 * (1 - t) + m_SpokeColor((i + 1) Mod m_Count).Green / 255 * t
        spokeCol.Blue = m_SpokeColor(i).Blue / 255 * (1 - t) + m_SpokeColor((i + 1) Mod m_Count).Blue / 255 * t
        
        If w > 1 Then
            newR = fClamp(spokeCol.Red * w, 0, 1)
            newG = fClamp(spokeCol.Green * w, 0, 1)
            newB = fClamp(spokeCol.Blue * w, 0, 1)
        Else
            newR = r / 255 * (1 - fRatio) + spokeCol.Red * fRatio
            newG = g / 255 * (1 - fRatio) + spokeCol.Green * fRatio
            newB = b / 255 * (1 - fRatio) + spokeCol.Blue * fRatio
        End If
            
        newR = (newR + ws) * 255
        newG = (newG + ws) * 255
        newB = (newB + ws) * 255

        If newR > 255 Then newR = 255
        If newG > 255 Then newG = 255
        If newB > 255 Then newB = 255
            
        'Assign the new values to each color channel
        dstImageData(QuickVal + 2, y) = newR
        dstImageData(QuickVal + 1, y) = newG
        dstImageData(QuickVal, y) = newB
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
       'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Sunshine", , buildParams(sltRadius.value, sltRayCount.value, cpShine.Color, sltXCenter.value, sltYCenter.value), UNDO_LAYER
End Sub
Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltXCenter.value = 0.1
    sltYCenter.value = 0.1
    sltRadius.value = 72
    sltRayCount.value = 100
    cpShine.Color = RGB(255, 255, 60)
End Sub

Private Sub cpShine_ColorChanged()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Display the previewed effect in the neighboring window
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()

    'Disable previewing until the form has been fully initialized
    cmdBar.markPreviewStatus False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub fxPreview_PointSelected(xRatio As Double, yRatio As Double)
    cmdBar.markPreviewStatus False
    sltXCenter.value = xRatio
    sltYCenter.value = yRatio
    cmdBar.markPreviewStatus True
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

Private Sub sltRayCount_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then SunShine sltRadius.value, sltRayCount.value, cpShine.Color, sltXCenter.value, sltYCenter.value, True, fxPreview
End Sub
'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Function fClamp(ByVal t As Double, ByVal dLow As Double, ByVal dHigh As Double) As Double
    If t < dHigh Then
        If t > dLow Then fClamp = t Else fClamp = dLow
        Exit Function
    End If
    fClamp = dHigh
End Function
