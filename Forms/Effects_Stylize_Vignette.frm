VERSION 5.00
Begin VB.Form FormVignette 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Apply Vignetting"
   ClientHeight    =   6540
   ClientLeft      =   -15
   ClientTop       =   225
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
   Begin PhotoDemon.sliderTextCombo sltXCenter 
      Height          =   405
      Left            =   6000
      TabIndex        =   10
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.smartOptionButton optShape 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   3
      Top             =   5340
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   582
      Caption         =   "fit to image"
      Value           =   -1  'True
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton optShape 
      Height          =   360
      Index           =   1
      Left            =   8880
      TabIndex        =   4
      Top             =   5340
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   582
      Caption         =   "circular"
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   720
      Left            =   6000
      TabIndex        =   5
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   1
      Max             =   100
      Value           =   60
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.sliderTextCombo sltFeathering 
      Height          =   720
      Left            =   6000
      TabIndex        =   6
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "softness"
      Min             =   1
      Max             =   100
      Value           =   30
   End
   Begin PhotoDemon.sliderTextCombo sltTransparency 
      Height          =   720
      Left            =   6000
      TabIndex        =   7
      Top             =   3120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "strength"
      Min             =   1
      Max             =   100
      Value           =   80
   End
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   930
      Left            =   6000
      TabIndex        =   8
      Top             =   3900
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1640
      Caption         =   "color"
      curColor        =   0
   End
   Begin PhotoDemon.sliderTextCombo sltYCenter 
      Height          =   405
      Left            =   9000
      TabIndex        =   11
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin VB.Label lblExplanation 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: you can also set a center position by clicking the preview window."
      ForeColor       =   &H00404040&
      Height          =   435
      Index           =   0
      Left            =   6120
      TabIndex        =   12
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
      Index           =   0
      Left            =   6000
      TabIndex        =   9
      Top             =   120
      Width           =   2205
   End
   Begin VB.Label lblShape 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "shape"
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
      Top             =   4980
      Width           =   615
   End
End
Attribute VB_Name = "FormVignette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Vignette tool
'Copyright 2013-2015 by Tanner Helland
'Created: 31/January/13
'Last updated: 09/January/14
'Last update: added center-point selection capabilities
'
'This tool allows the user to apply vignetting to an image.  Many options are available, and all should be
' self-explanatory!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply vignetting to an image
Public Sub ApplyVignette(ByVal maxRadius As Double, ByVal vFeathering As Double, ByVal vTransparency As Double, ByVal vMode As Boolean, ByVal newColor As Long, Optional ByVal centerPosX As Double = 0.5, Optional ByVal centerPosY As Double = 0.5, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying vignetting..."
        
    'Extract the RGB values of the vignetting color
    Dim newR As Byte, newG As Byte, newB As Byte
    newR = ExtractR(newColor)
    newG = ExtractG(newColor)
    newB = ExtractB(newColor)
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
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
    midX = CDbl(finalX - initX) * centerPosX
    midX = midX + initX
    midY = CDbl(finalY - initY) * centerPosY
    midY = midY + initY
            
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    Dim nX2 As Double, nY2 As Double
            
    'Radius is based off the smaller of the two dimensions - width or height
    Dim tWidth As Long, tHeight As Long
    tWidth = curDIBValues.Width
    tHeight = curDIBValues.Height
    Dim sRadiusW As Double, sRadiusH As Double
    Dim sRadiusW2 As Double, sRadiusH2 As Double
    
    sRadiusW = tWidth * (maxRadius / 100)
    sRadiusW2 = sRadiusW * sRadiusW
    sRadiusH = tHeight * (maxRadius / 100)
    sRadiusH2 = sRadiusH * sRadiusH
    
    'Adjust the vignetting to be a proportion of the image's maximum radius.  This ensures accurate correlations
    ' between the preview and the final result.
    Dim vFeathering2 As Double
    
    If vMode Then
        vFeathering2 = (vFeathering / 100) * (sRadiusW * sRadiusH)
    Else
        If sRadiusW < sRadiusH Then
            vFeathering2 = (vFeathering / 100) * (sRadiusW * sRadiusW)
        Else
            vFeathering2 = (vFeathering / 100) * (sRadiusH * sRadiusH)
        End If
    End If
    
    'Modify the transparency to be on a scale of [0, 1]
    vTransparency = 1 - (vTransparency / 100)
    
    Dim sRadiusCircular As Double, sRadiusMax As Double, sRadiusMin As Double
    If sRadiusW < sRadiusH Then
        sRadiusCircular = sRadiusW2
    Else
        sRadiusCircular = sRadiusH2
    End If
    sRadiusMin = sRadiusCircular - vFeathering2
    
    Dim blendVal As Double
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Remap the coordinates around a center point of (0, 0)
        nX = x - midX
        nY = y - midY
        nX2 = nX * nX
        nY2 = nY * nY
                
        'Fit to image (elliptical)
        If vMode Then
                
            'If the values are going to be out-of-bounds, force them to black
            sRadiusMax = sRadiusH2 - ((sRadiusH2 * nX2) / sRadiusW2)
            
            If nY2 > sRadiusMax Then
                
                dstImageData(QuickVal + 2, y) = BlendColors(newR, dstImageData(QuickVal + 2, y), vTransparency)
                dstImageData(QuickVal + 1, y) = BlendColors(newG, dstImageData(QuickVal + 1, y), vTransparency)
                dstImageData(QuickVal, y) = BlendColors(newB, dstImageData(QuickVal, y), vTransparency)
                
            'Otherwise, check for feathering
            Else
                sRadiusMin = sRadiusMax - vFeathering2
                
                If nY2 >= sRadiusMin Then
                    blendVal = (nY2 - sRadiusMin) / vFeathering2
                    blendVal = blendVal * (1 - vTransparency)
                    
                    dstImageData(QuickVal + 2, y) = BlendColors(dstImageData(QuickVal + 2, y), newR, blendVal)
                    dstImageData(QuickVal + 1, y) = BlendColors(dstImageData(QuickVal + 1, y), newG, blendVal)
                    dstImageData(QuickVal, y) = BlendColors(dstImageData(QuickVal, y), newB, blendVal)
                End If
                    
            End If
                
        'Circular
        Else
        
            'If the values are going to be out-of-bounds, force them to black
            If (nX2 + nY2) > sRadiusCircular Then
                dstImageData(QuickVal + 2, y) = BlendColors(newR, dstImageData(QuickVal + 2, y), vTransparency)
                dstImageData(QuickVal + 1, y) = BlendColors(newG, dstImageData(QuickVal + 1, y), vTransparency)
                dstImageData(QuickVal, y) = BlendColors(newB, dstImageData(QuickVal, y), vTransparency)
                
            'Otherwise, check for feathering
            Else
                
                If (nX2 + nY2) >= sRadiusMin Then
                    blendVal = (nX2 + nY2 - sRadiusMin) / vFeathering2
                    blendVal = blendVal * (1 - vTransparency)
                    
                    dstImageData(QuickVal + 2, y) = BlendColors(dstImageData(QuickVal + 2, y), newR, blendVal)
                    dstImageData(QuickVal + 1, y) = BlendColors(dstImageData(QuickVal + 1, y), newG, blendVal)
                    dstImageData(QuickVal, y) = BlendColors(dstImageData(QuickVal, y), newB, blendVal)
                End If
                
            End If
                
        End If
                        
    Next y
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
End Sub

Private Sub cmdBar_OKClick()
    Process "Vignetting", , buildParams(sltRadius.Value, sltFeathering.Value, sltTransparency.Value, optShape(0).Value, colorPicker.Color, sltXCenter.Value, sltYCenter.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltXCenter.Value = 0.5
    sltYCenter.Value = 0.5
    sltRadius.Value = 60
    sltFeathering.Value = 30
    sltTransparency.Value = 80
    colorPicker.Color = RGB(0, 0, 0)
End Sub

Private Sub colorPicker_ColorChanged()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Draw a preview of the effect
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub fxPreview_ColorSelected()
    colorPicker.Color = fxPreview.SelectedColor
    updatePreview
End Sub

'The user can right-click the preview area to select a new center point
Private Sub fxPreview_PointSelected(xRatio As Double, yRatio As Double)
    
    cmdBar.markPreviewStatus False
    sltXCenter.Value = xRatio
    sltYCenter.Value = yRatio
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub optShape_Click(Index As Integer)
    updatePreview
End Sub

Private Sub sltFeathering_Change()
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

Private Sub sltTransparency_Change()
    updatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyVignette sltRadius.Value, sltFeathering.Value, sltTransparency.Value, optShape(0).Value, colorPicker.Color, sltXCenter.Value, sltYCenter.Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltXCenter_Change()
    updatePreview
End Sub

Private Sub sltYCenter_Change()
    updatePreview
End Sub
