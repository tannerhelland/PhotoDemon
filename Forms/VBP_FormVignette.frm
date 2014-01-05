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
   Begin PhotoDemon.smartOptionButton optShape 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   7
      Top             =   4440
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   635
      Caption         =   "fit to image"
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      ColorSelection  =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton optShape 
      Height          =   360
      Index           =   1
      Left            =   8880
      TabIndex        =   8
      Top             =   4440
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   635
      Caption         =   "circular"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      TabIndex        =   9
      Top             =   810
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   100
      Value           =   60
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
   Begin PhotoDemon.sliderTextCombo sltFeathering 
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   1650
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   100
      Value           =   30
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
   Begin PhotoDemon.sliderTextCombo sltTransparency 
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   2490
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   100
      Value           =   80
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
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   3480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   873
      curColor        =   0
   End
   Begin VB.Label lblShape 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "shape:"
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
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vignetting color (click box to change):"
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
      TabIndex        =   5
      Top             =   3000
      Width           =   4020
   End
   Begin VB.Label lblFeathering 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "softness:"
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
      TabIndex        =   4
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label lblTransparency 
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
      Left            =   6000
      TabIndex        =   3
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label lblRadius 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "radius (percentage):"
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
      Top             =   480
      Width           =   2145
   End
End
Attribute VB_Name = "FormVignette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Vignette tool
'Copyright ©2013-2014 by Tanner Helland
'Created: 31/January/13
'Last updated: 24/August/13
'Last update: added command bar
'
'This tool allows the user to apply vignetting to an image.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Apply vignetting to an image
Public Sub ApplyVignette(ByVal maxRadius As Double, ByVal vFeathering As Double, ByVal vTransparency As Double, ByVal vMode As Boolean, ByVal newColor As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Applying vignetting..."
        
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
        
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
        
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    Dim nX2 As Double, nY2 As Double
            
    'Radius is based off the smaller of the two dimensions - width or height
    Dim tWidth As Long, tHeight As Long
    tWidth = curLayerValues.Width
    tHeight = curLayerValues.Height
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
    Process "Vignetting", , buildParams(sltRadius.Value, sltFeathering.Value, sltTransparency.Value, optShape(0).Value, colorPicker.Color)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 60
    sltFeathering.Value = 30
    sltTransparency.Value = 80
    colorPicker.Color = RGB(0, 0, 0)
End Sub

Private Sub colorPicker_ColorChanged()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
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
    If cmdBar.previewsAllowed Then ApplyVignette sltRadius.Value, sltFeathering.Value, sltTransparency.Value, optShape(0).Value, colorPicker.Color, True, fxPreview
End Sub
