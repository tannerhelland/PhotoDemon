VERSION 5.00
Begin VB.Form FormKaleidoscope 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Kaleidoscope"
   ClientHeight    =   6675
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   12135
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
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   809
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5925
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   1323
      BackColor       =   14802140
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
   Begin PhotoDemon.buttonStrip btsOptions 
      Height          =   600
      Left            =   6120
      TabIndex        =   3
      Top             =   4620
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   1058
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   0
      Left            =   5880
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   4
      Top             =   360
      Width           =   6135
      Begin PhotoDemon.sliderTextCombo sltMirrors 
         Height          =   720
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "number of mirrors"
         Min             =   1
         Max             =   16
         Value           =   3
         NotchPosition   =   2
         NotchValueCustom=   8
      End
      Begin PhotoDemon.sliderTextCombo sltAngle 
         Height          =   720
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "primary angle"
         Max             =   360
         SigDigits       =   1
      End
      Begin PhotoDemon.sliderTextCombo sltXCenter 
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         Max             =   1
         SigDigits       =   2
         Value           =   0.5
         NotchPosition   =   2
         NotchValueCustom=   0.5
      End
      Begin PhotoDemon.sliderTextCombo sltYCenter 
         Height          =   405
         Left            =   3120
         TabIndex        =   9
         Top             =   600
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
         Left            =   240
         TabIndex        =   11
         Top             =   1170
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
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   1
      Left            =   5880
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   6135
      Begin PhotoDemon.sliderTextCombo sltAngle2 
         Height          =   720
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "secondary angle"
         Max             =   360
         SigDigits       =   1
      End
      Begin PhotoDemon.sliderTextCombo sltRadius 
         Height          =   720
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "radius (percentage)"
         Min             =   1
         Max             =   100
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.buttonStrip btsQuality 
         Height          =   600
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   1058
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "render emphasis"
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
         Left            =   120
         TabIndex        =   14
         Top             =   2265
         Width           =   1755
      End
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "options"
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
      Index           =   6
      Left            =   6000
      TabIndex        =   2
      Top             =   4200
      Width           =   780
   End
End
Attribute VB_Name = "FormKaleidoscope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image "Kaleiodoscope" Distortion
'Copyright 2013-2015 by Tanner Helland
'Created: 14/January/13
'Last updated: 25/September/14
'Last update: interface improvements
'
'This tool allows the user to apply a simulated kaleidoscope distort to the image.  A number of variables can be
' set as part of the transformation; simply playing with the sliders should give a good indication of how they
' all work.
'
'As of January '14, the user can now select any center point for the effect.
'
'Finally, the transformation used by this tool is a modified version of a transformation originally written by
' Jerry Huxtable of JH Labs.  Jerry's original code is licensed under an Apache 2.0 license.  You may download his
' original version at the following link (good as of 14 January '13): http://www.jhlabs.com/ip/filters/index.html
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a "kaleidoscope" effect to an image
Public Sub KaleidoscopeImage(ByVal numMirrors As Long, ByVal primaryAngle As Double, ByVal secondaryAngle As Double, ByVal effectRadius As Double, ByVal useBilinear As Boolean, Optional ByVal centerX As Double = 0.5, Optional ByVal centerY As Double = 0.5, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Peering at image through imaginary kaleidoscope..."
    
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
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, EDGE_CLAMP, useBilinear, curDIBValues.maxX, curDIBValues.MaxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
          
    'Kaleidoscoping requires some specialized variables
    
    'Convert the input angles to radians
    primaryAngle = primaryAngle * (PI / 180)
    secondaryAngle = secondaryAngle * (PI / 180)
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerX
    midX = midX + initX
    midY = CDbl(finalY - initY) * centerY
    midY = midY + initY
    
    'Additional kaleidoscope variables
    Dim theta As Double, sRadius As Double, tRadius As Double, sDistance As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
            
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curDIBValues.Width
    tHeight = curDIBValues.Height
    sRadius = Sqr(tWidth * tWidth + tHeight * tHeight) / 2
              
    sRadius = sRadius * (effectRadius / 100)
                  
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Remap the coordinates around a center point of (0, 0)
        nX = x - midX
        nY = y - midY
        
        'Calculate distance
        sDistance = Sqr((nX * nX) + (nY * nY))
                
        'Calculate theta
        theta = Atan2(nY, nX) - primaryAngle - secondaryAngle
        theta = convertTriangle((theta / PI) * numMirrors * 0.5)
                
        'Calculate remapped x and y values
        If (sRadius > 0) Then
            
            tRadius = sRadius / Cos(theta)
            sDistance = tRadius * convertTriangle(sDistance / tRadius)

        Else
            tRadius = sDistance
        End If
        
        theta = theta + primaryAngle
        
        srcX = midX + sDistance * Cos(theta)
        srcY = midY + sDistance * Sin(theta)
        
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                
    Next y
        If Not toPreview Then
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

'Change the active options panel
Private Sub btsOptions_Click(ByVal buttonIndex As Long)
    picContainer(buttonIndex).Visible = True
    picContainer(Abs(1 - buttonIndex)).Visible = False
End Sub

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    updatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Kaleidoscope", , buildParams(sltMirrors, sltAngle, sltAngle2, sltRadius, (btsQuality.ListIndex = 0), sltXCenter.Value, sltYCenter.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltXCenter.Value = 0.5
    sltYCenter.Value = 0.5
    sltMirrors.Value = 3
    sltRadius.Value = 100
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
        
    'Create the preview
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Populate the options selector
    btsOptions.AddItem "basic", 0
    btsOptions.AddItem "advanced", 1
    btsOptions.ListIndex = 0
    
    'Populate the quality selector
    btsQuality.AddItem "quality", 0
    btsQuality.AddItem "speed", 1
    btsQuality.ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

Private Sub sltAngle2_Change()
    updatePreview
End Sub

Private Sub sltMirrors_Change()
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then KaleidoscopeImage sltMirrors, sltAngle, sltAngle2, sltRadius, (btsQuality.ListIndex = 0), sltXCenter.Value, sltYCenter.Value, True, fxPreview
End Sub

'Return a repeating triangle shape in the range [0, 1] with wavelength 1
Private Function convertTriangle(ByVal trInput As Double) As Double

    Dim tmpCalc As Double
    tmpCalc = Modulo(trInput, 1)
    convertTriangle = IIf(tmpCalc < 0.5, tmpCalc, 1 - tmpCalc)
    
End Function

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
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

Private Sub sltXCenter_Change()
    updatePreview
End Sub

Private Sub sltYCenter_Change()
    updatePreview
End Sub

