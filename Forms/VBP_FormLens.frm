VERSION 5.00
Begin VB.Form FormLens 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Lens Correction and Distortion"
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
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10590
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin VB.TextBox txtRadius 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "50"
      Top             =   2940
      Width           =   735
   End
   Begin VB.HScrollBar hsRadius 
      Height          =   255
      Left            =   6120
      Max             =   100
      Min             =   1
      TabIndex        =   8
      Top             =   3000
      Value           =   50
      Width           =   4815
   End
   Begin VB.OptionButton OptInterpolate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " quality"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   7
      Top             =   3840
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton OptInterpolate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " speed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Index           =   1
      Left            =   7560
      TabIndex        =   6
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtIndex 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   11040
      MaxLength       =   6
      TabIndex        =   4
      Text            =   "1.20"
      Top             =   2160
      Width           =   735
   End
   Begin VB.HScrollBar hsIndex 
      Height          =   255
      LargeChange     =   10
      Left            =   6120
      Max             =   500
      Min             =   90
      TabIndex        =   3
      Top             =   2220
      Value           =   120
      Width           =   4815
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label lblHeight 
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
      TabIndex        =   10
      Top             =   2640
      Width           =   2145
   End
   Begin VB.Label lblInterpolation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "render emphasis:"
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
      Top             =   3480
      Width           =   1845
   End
   Begin VB.Label lblAmount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lens strength (refractive index):"
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
      Top             =   1800
      Width           =   3330
   End
End
Attribute VB_Name = "FormLens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Lens Correction and Distortion
'Copyright ©2000-2013 by Tanner Helland
'Created: 05/January/13
'Last updated: 06/January/12
'Last update: applied additional optimizations and tweaks
'
'This tool allows the user to correct (or apply) a lens distortion to an image.  Bilinear interpolation
' (via reverse-mapping) is available for high-quality lensing.
'
'At present, the tool assumes that you want to refract the image around its centerpoint.  The code is already set up to handle
' alternative center points - there simply needs to be a good user interface technique for establishing the center.
'
'Finally, the transformation used by this tool is a modified version of a transformation originally written by
' Jerry Huxtable of JH Labs.  Jerry's original code is licensed under an Apache 2.0 license.  You may download his
' original version at the following link (good as of 07 January '13): http://www.jhlabs.com/ip/filters/index.html
'
'***************************************************************************

Option Explicit

'Use this to prevent the text box and scroll bar from updating each other in an endless loop
Dim userChange As Boolean

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'Before rendering anything, check to make sure the text boxes have valid input
    If Not EntryValid(txtIndex * 100, hsIndex.Min, hsIndex.Max, True, True) Then
        AutoSelectText txtIndex
        Exit Sub
    End If

    If Not EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, True, True) Then
        AutoSelectText txtRadius
        Exit Sub
    End If

    Me.Visible = False
    
    'Based on the user's selection, submit the proper processor request
    If OptInterpolate(0) Then
        Process DistortLens, CDbl(hsIndex / 100), hsRadius.Value, True
    Else
        Process DistortLens, CDbl(hsIndex / 100), hsRadius.Value, False
    End If
    
    Unload Me
    
End Sub

'Apply lens distortion (or correction) to an image
Public Sub ApplyLensDistortion(ByVal refractiveIndex As Double, ByVal lensRadius As Double, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    refractiveIndex = 1 / refractiveIndex

    If toPreview = False Then Message "Projecting image through simulated lens..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'Because interpolation may be used, it's necessary to keep pixel values within special ranges
    Dim xLimit As Long, yLimit As Long
    If useBilinear Then
        xLimit = finalX - 1
        yLimit = finalY - 1
    Else
        xLimit = finalX
        yLimit = finalY
    End If
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickVal2 As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
          
    'Lensing requires some specialized variables
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
    
    'Calculation values
    Dim sRadius As Double, sRadius2 As Double, sDistance As Double
    Dim theta As Double, theta2 As Double
    Dim xAngle As Double, yAngle As Double
    Dim firstAngle As Double, secondAngle As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    Dim nX2 As Double, nY2 As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
    
    Dim i As Long
    
    'Radius is based off the smaller of the two dimensions - width or height
    Dim tWidth As Long, tHeight As Long
    tWidth = curLayerValues.Width
    tHeight = curLayerValues.Height
    Dim sRadiusW As Double, sRadiusH As Double
    Dim sRadiusW2 As Double, sRadiusH2 As Double
    
    sRadiusW = tWidth * (lensRadius / 100)
    sRadiusW2 = sRadiusW * sRadiusW
    sRadiusH = tHeight * (lensRadius / 100)
    sRadiusH2 = sRadiusH * sRadiusH
    
    Dim sRadiusMult As Double
    sRadiusMult = sRadiusW * sRadiusH
              
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Remap the coordinates around a center point of (0, 0)
        nX = x - midX
        nY = y - midY
        nX2 = nX * nX
        nY2 = nY * nY
                
        'If the values are going to be out-of-bounds, simply maintain the current x and y values
        If nY2 >= (sRadiusH2 - ((sRadiusH2 * nX2) / sRadiusW2)) Then
            srcX = x
            srcY = y
        
        'Otherwise, reverse-map x and y back onto the original image using a reversed lens refraction calculation
        Else
        
            'Calculate theta
            theta = Sqr((1 - (nX2 / sRadiusW2) - (nY2 / sRadiusH2)) * sRadiusMult)
            theta2 = theta * theta
            
            'Calculate the angle for x
            xAngle = Acos(nX / Sqr(nX2 + theta2))
            firstAngle = PI_HALF - xAngle
            secondAngle = Asin(Sin(firstAngle) * refractiveIndex)
            secondAngle = PI_HALF - xAngle - secondAngle
            srcX = x - Tan(secondAngle) * theta
            
            'Now do the same thing for y
            yAngle = Acos(nY / Sqr(nY2 + theta2))
            firstAngle = PI_HALF - yAngle
            secondAngle = Asin(Sin(firstAngle) * refractiveIndex)
            secondAngle = PI_HALF - yAngle - secondAngle
            srcY = y - Tan(secondAngle) * theta
            
        End If
        
        'Make sure the source coordinates are in-bounds
        If srcX < 0 Then srcX = 0
        If srcY < 0 Then srcY = 0
        If srcX > xLimit Then srcX = xLimit
        If srcY > yLimit Then srcY = yLimit
        
        'Interpolate the result if desired, otherwise use nearest-neighbor
        If useBilinear Then
        
            For i = 0 To qvDepth - 1
                dstImageData(QuickVal + i, y) = getInterpolatedVal(srcX, srcY, srcImageData, i, qvDepth)
            Next i
        
        Else
        
            QuickVal2 = Int(srcX) * qvDepth
        
            For i = 0 To qvDepth - 1
                dstImageData(QuickVal + i, y) = srcImageData(QuickVal2 + i, Int(srcY))
            Next i
                
        End If
                
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
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

Private Sub Form_Activate()
    
    'Draw a preview of the effect
    updatePreview
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Mark scroll bar changes as coming from the user
    userChange = True
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Keep the scroll bar and the text box values in sync
Private Sub hsIndex_Change()
    If userChange Then
        txtIndex.Text = Format(CDbl(hsIndex.Value) / 100, "0.00")
        txtIndex.Refresh
    End If
    updatePreview
End Sub

Private Sub hsIndex_Scroll()
    txtIndex.Text = Format(CDbl(hsIndex.Value) / 100, "0.00")
    txtIndex.Refresh
    updatePreview
End Sub

Private Sub hsRadius_Change()
    copyToTextBoxI txtRadius, hsRadius.Value
    updatePreview
End Sub

Private Sub hsRadius_Scroll()
    copyToTextBoxI txtRadius, hsRadius.Value
    updatePreview
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub txtIndex_GotFocus()
    AutoSelectText txtIndex
End Sub

Private Sub txtIndex_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtIndex, True, True
    If EntryValid(txtIndex, hsIndex.Min / 100, hsIndex.Max / 100, False, False) Then
        userChange = False
        hsIndex.Value = Val(txtIndex) * 100
        userChange = True
    End If
End Sub

Private Sub txtRadius_GotFocus()
    AutoSelectText txtRadius
End Sub

Private Sub txtRadius_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtRadius
    If EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, False, False) Then hsRadius.Value = Val(txtRadius)
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()

    If OptInterpolate(0) Then
        ApplyLensDistortion CDbl(hsIndex / 100), hsRadius.Value, True, True, fxPreview
    Else
        ApplyLensDistortion CDbl(hsIndex / 100), hsRadius.Value, False, True, fxPreview
    End If

End Sub
