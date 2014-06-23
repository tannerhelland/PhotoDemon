VERSION 5.00
Begin VB.Form FormBilateral 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Bilateral Smoothing"
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
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   930
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   3
      Max             =   25
      Value           =   9
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
   Begin PhotoDemon.sliderTextCombo sltSpatialFactor 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   1770
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   255
      SigDigits       =   2
      Value           =   10
   End
   Begin PhotoDemon.sliderTextCombo sltSpatialPower 
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   2640
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   255
      SigDigits       =   2
      Value           =   2
   End
   Begin PhotoDemon.sliderTextCombo sltColorFactor 
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   3600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   255
      SigDigits       =   2
      Value           =   50
   End
   Begin PhotoDemon.sliderTextCombo sltColorPower 
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   4560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   255
      SigDigits       =   2
      Value           =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "color power:"
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
      Top             =   4200
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "color factor:"
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
      TabIndex        =   9
      Top             =   3240
      Width           =   1290
   End
   Begin VB.Label lblLuminance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "spatial power:"
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
      TabIndex        =   3
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label lblSaturation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "spatial factor:"
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
      TabIndex        =   2
      Top             =   1440
      Width           =   1440
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
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "FormBilateral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Bilateral Smoothing Form
'Copyright ©2014 by Audioglider
'Created: 19/June/14
'Last updated: 19/June/14
'Last update: Initial build
'
'This filter performs "selective" gaussian smoothing of areas of same color
' (domains) which removes noise and contrast artifacts while perserving
' sharp edges.
'
'The two major parameters "spatial factor" and "color factor" define the
' results of the filter. By changing them you can achieve either only noise
' reduction with little change to the image or achieve a silky effect
' to the entire image.
'
'More details on the algorithm can be found at:
' http://www.cs.duke.edu/~tomasi/papers/tomasi/tomasiIccv98.pdf
'***************************************************************************

Option Explicit

Private Const maxKernelSize As Long = 256
Private Const colorsCount As Long = 256

Dim spatialFunc() As Double
Dim colorFunc() As Double

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub initSpatialFunc(ByVal kernelSize As Long, ByVal spatialFactor As Double, ByVal spatialPower As Double)
    Dim c As Long, i As Long, k As Long
    
    c = kernelSize / 2
    
    ReDim spatialFunc(0 To kernelSize, 0 To kernelSize)
    For i = 0 To kernelSize - 1
        For k = 0 To kernelSize - 1
            spatialFunc(i, k) = Exp(-0.5 * ((Sqr(((i - c) * (i - c) + (k - c) * (k - c)) / spatialFactor) ^ spatialPower)))
        Next k
    Next i
End Sub

Private Sub initColorFunc(ByVal colorFactor As Double, ByVal colorPower As Double)
    Dim i As Long, k As Long
    
    ReDim colorFunc(0 To colorsCount - 1, 0 To colorsCount - 1)
    For i = 0 To colorsCount - 1
        For k = 0 To colorsCount - 1
            colorFunc(i, k) = Exp(-0.5 * ((Abs(i - k) / colorFactor) ^ colorPower))
        Next k
    Next i
End Sub

'Parameters: * kernelRadius [size of square for limiting surrounding pixels that take part in calculation.
' NOTE: Small values < 9 on high-res images do not provide significant results.]
' * spatialFactor [determines smoothing power within a color domain (neighborhood pixels of similar color]
' * spatialPower [exponent power, used in spatial function calculation]
' * colorFactor [determines the variance of color for a color domain]
' * colorPower [exponent power, used in color function calculation]
Public Sub BilateralSmoothing(ByVal kernelRadius As Long, ByVal spatialFactor As Double, ByVal spatialPower As Double, ByVal colorFactor As Double, ByVal colorPower As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
        
    If Not toPreview Then Message "Applying bilateral smoothing..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array. This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already-processed pixels from affecting the results of later pixels.)
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
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary. This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'If this is a preview, we need to adjust the kernal
    If toPreview Then kernelRadius = kernelRadius * curDIBValues.previewModifier
    
    'Color variables
    Dim srcR As Long, srcG As Long, srcB As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim srcR0 As Long, srcG0 As Long, srcB0 As Long
    
    Dim sCoefR As Double, sCoefG As Double, sCoefB As Double
    Dim sMembR As Double, sMembG As Double, sMembB As Double
    Dim coefR As Double, coefG As Double, coefB As Double
    Dim xOffset As Long, yOffset As Long, xMax As Long, yMax As Long
    Dim i As Long, k As Long
    
    'For performance improvements, color and spatial functions are precalculated prior to starting filter.
    initSpatialFunc kernelRadius * 2 + 1, spatialFactor, spatialPower
    initColorFunc colorFactor, colorPower
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        sCoefR = 0
        sCoefG = 0
        sCoefB = 0
        sMembR = 0
        sMembG = 0
        sMembB = 0
        
        srcR0 = srcImageData(QuickVal + 2, y)
        srcG0 = srcImageData(QuickVal + 1, y)
        srcB0 = srcImageData(QuickVal, y)
        
        xMax = x + kernelRadius
        yMax = y + kernelRadius
        For xOffset = x - kernelRadius To xMax - 1
            For yOffset = y - kernelRadius To yMax - 1
                            
                ' bounds check
                If (xOffset >= initX) And (xOffset < finalX) And (yOffset >= initY) And (yOffset < finalY) Then
                        
                    srcR = srcImageData(xOffset * qvDepth + 2, yOffset)
                    srcG = srcImageData(xOffset * qvDepth + 1, yOffset)
                    srcB = srcImageData(xOffset * qvDepth, yOffset)
                
                    coefR = spatialFunc(x - xOffset + kernelRadius, y - yOffset + kernelRadius) * colorFunc(srcR, srcR0)
                    coefG = spatialFunc(x - xOffset + kernelRadius, y - yOffset + kernelRadius) * colorFunc(srcG, srcG0)
                    coefB = spatialFunc(x - xOffset + kernelRadius, y - yOffset + kernelRadius) * colorFunc(srcB, srcB0)
                
                    sCoefR = sCoefR + coefR
                    sCoefG = sCoefG + coefG
                    sCoefB = sCoefB + coefB
                
                    sMembR = sMembR + coefR * srcR
                    sMembG = sMembG + coefG * srcG
                    sMembB = sMembB + coefB * srcB
                        
                End If
                        
            Next yOffset
        Next xOffset
              
        newR = sMembR / sCoefR
        newG = sMembG / sCoefG
        newB = sMembB / sCoefB
                
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
    Process "Bilateral smoothing", , buildParams(sltRadius.Value, sltSpatialFactor.Value, sltSpatialPower.Value, sltColorFactor.Value, sltColorPower.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 9
    sltSpatialFactor.Value = 10
    sltColorFactor.Value = 50
    sltSpatialPower.Value = 2
    sltColorPower.Value = 2
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

Private Sub sltColorFactor_Change()
    updatePreview
End Sub

Private Sub sltColorPower_Change()
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

Private Sub sltSpatialPower_Change()
    updatePreview
End Sub

Private Sub sltSpatialFactor_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then BilateralSmoothing sltRadius.Value, sltSpatialFactor.Value, sltSpatialPower.Value, sltColorFactor.Value, sltColorPower.Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub
