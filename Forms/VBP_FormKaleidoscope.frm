VERSION 5.00
Begin VB.Form FormKaleidoscope 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Kaleidoscope"
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
   Begin VB.HScrollBar hsAngle2 
      Height          =   255
      LargeChange     =   10
      Left            =   6120
      Max             =   3600
      TabIndex        =   17
      Top             =   2970
      Width           =   4815
   End
   Begin VB.TextBox txtAngle2 
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
      TabIndex        =   16
      Text            =   "0.0"
      Top             =   2910
      Width           =   735
   End
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
   Begin VB.TextBox txtMirrors 
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
      MaxLength       =   4
      TabIndex        =   12
      Text            =   "3"
      Top             =   1140
      Width           =   735
   End
   Begin VB.HScrollBar hsMirrors 
      Height          =   255
      Left            =   6120
      Max             =   16
      Min             =   1
      TabIndex        =   11
      Top             =   1200
      Value           =   3
      Width           =   4815
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
      Text            =   "100"
      Top             =   3690
      Width           =   735
   End
   Begin VB.HScrollBar hsRadius 
      Height          =   255
      Left            =   6120
      Max             =   100
      Min             =   1
      TabIndex        =   8
      Top             =   3750
      Value           =   100
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
      Top             =   4530
      Value           =   -1  'True
      Width           =   1695
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
      Left            =   7920
      TabIndex        =   6
      Top             =   4530
      Width           =   2535
   End
   Begin VB.TextBox txtAngle 
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
      Text            =   "0.0"
      Top             =   2040
      Width           =   735
   End
   Begin VB.HScrollBar hsAngle 
      Height          =   255
      LargeChange     =   10
      Left            =   6120
      Max             =   3600
      TabIndex        =   3
      Top             =   2100
      Width           =   4815
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "secondary angle:"
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
      TabIndex        =   18
      Top             =   2535
      Width           =   1800
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "number of mirrors:"
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
      TabIndex        =   13
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblTitle 
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
      Index           =   1
      Left            =   6000
      TabIndex        =   10
      Top             =   3390
      Width           =   2145
   End
   Begin VB.Label lblTitle 
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
      Index           =   2
      Left            =   6000
      TabIndex        =   5
      Top             =   4170
      Width           =   1845
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "primary angle:"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   1560
   End
End
Attribute VB_Name = "FormKaleidoscope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image "Kaleiodoscope" Distortion
'Copyright ©2000-2013 by Tanner Helland
'Created: 14/January/13
'Last updated: 15/January/13
'Last update: added use of pdFilterSupport class for better interpolation and edge handling
'
'This tool allows the user to apply a simulated kaleidoscope distort to the image.  A number of variables can be
' set as part of the transformation; simply playing with the sliders should give a good indication of how they
' all work.
'
'At present, the tool assumes that you want to apply the kaleidoscope effect around the center of the image.
' The code is already set up to handle alternative center points - there simply needs to be a good user interface
' technique for establishing the center.
'
'Finally, the transformation used by this tool is a modified version of a transformation originally written by
' Jerry Huxtable of JH Labs.  Jerry's original code is licensed under an Apache 2.0 license.  You may download his
' original version at the following link (good as of 14 January '13): http://www.jhlabs.com/ip/filters/index.html
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
    If Not EntryValid(txtAngle, hsAngle.Min / 10, hsAngle.Max / 10, True, True) Then
        AutoSelectText txtAngle
        Exit Sub
    End If

    If Not EntryValid(txtAngle2, hsAngle2.Min / 10, hsAngle2.Max / 10, True, True) Then
        AutoSelectText txtAngle2
        Exit Sub
    End If

    If Not EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, True, True) Then
        AutoSelectText txtRadius
        Exit Sub
    End If
    
    If Not EntryValid(txtMirrors, hsMirrors.Min, hsMirrors.Max, True, True) Then
        AutoSelectText txtMirrors
        Exit Sub
    End If

    Me.Visible = False
    
    'Based on the user's selection, submit the proper processor request
    Process DistortKaleidoscope, CDbl(hsMirrors), CDbl(hsAngle / 10), CDbl(hsAngle2 / 10), hsRadius.Value, OptInterpolate(0)
        
    Unload Me
    
End Sub

'Apply a "kaleidoscope" effect to an image
Public Sub KaleidoscopeImage(ByVal numMirrors As Double, ByVal primaryAngle As Double, ByVal secondaryAngle As Double, ByVal effectRadius As Double, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If toPreview = False Then Message "Peering at image through imaginary kaleidoscope..."
    
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
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, EDGE_CLAMP, useBilinear, curLayerValues.MaxX, curLayerValues.MaxY
    
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
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
    
    'Additional kaleidoscope variables
    Dim theta As Double, sRadius As Double, tRadius As Double, sDistance As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
            
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curLayerValues.Width
    tHeight = curLayerValues.Height
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
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Mark scroll bar changes as coming from the user
    userChange = True
    
    'Create the preview
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub hsMirrors_Change()
    copyToTextBoxI txtMirrors, hsMirrors.Value
    updatePreview
End Sub

Private Sub hsMirrors_Scroll()
    copyToTextBoxI txtMirrors, hsMirrors.Value
    updatePreview
End Sub

'Keep the scroll bar and the text box values in sync
Private Sub hsAngle_Change()
    If userChange Then
        txtAngle.Text = Format(CDbl(hsAngle.Value) / 10, "##0.0")
        txtAngle.Refresh
    End If
    updatePreview
End Sub

Private Sub hsAngle_Scroll()
    txtAngle.Text = Format(CDbl(hsAngle.Value) / 10, "##0.0")
    txtAngle.Refresh
    updatePreview
End Sub

Private Sub hsAngle2_Change()
    If userChange Then
        txtAngle2.Text = Format(CDbl(hsAngle2.Value) / 10, "##0.0")
        txtAngle2.Refresh
    End If
    updatePreview
End Sub

Private Sub hsAngle2_Scroll()
    txtAngle2.Text = Format(CDbl(hsAngle2.Value) / 10, "##0.0")
    txtAngle2.Refresh
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

Private Sub txtMirrors_GotFocus()
    AutoSelectText txtMirrors
End Sub

Private Sub txtMirrors_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtMirrors, True
    If EntryValid(txtMirrors, hsMirrors.Min, hsMirrors.Max, False, False) Then hsMirrors.Value = Val(txtMirrors)
End Sub

Private Sub txtAngle_GotFocus()
    AutoSelectText txtAngle
End Sub

Private Sub txtAngle_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtAngle, True, True
    If EntryValid(txtAngle, hsAngle.Min / 10, hsAngle.Max / 10, False, False) Then
        userChange = False
        hsAngle.Value = Val(txtAngle) * 10
        userChange = True
    End If
End Sub

Private Sub txtAngle2_GotFocus()
    AutoSelectText txtAngle2
End Sub

Private Sub txtAngle2_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtAngle2, True, True
    If EntryValid(txtAngle2, hsAngle2.Min / 10, hsAngle2.Max / 10, False, False) Then
        userChange = False
        hsAngle2.Value = Val(txtAngle2) * 10
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

    KaleidoscopeImage CDbl(hsMirrors), CDbl(hsAngle / 10), CDbl(hsAngle2 / 10), hsRadius.Value, OptInterpolate(0), True, fxPreview
    
End Sub

'Return a repeating triangle shape in the range [0, 1] with wavelength 1
Private Function convertTriangle(ByVal trInput As Double) As Double

    Dim tmpCalc As Double
    tmpCalc = Modulo(trInput, 1)
    convertTriangle = IIf(tmpCalc < 0.5, tmpCalc, 1 - tmpCalc)
    
End Function

'This is a modified module function; it handles negative values specially to ensure they work with our kaleidoscope function
Private Function Modulo(ByVal Quotient As Double, ByVal Divisor As Double) As Double
    Modulo = Quotient - Fix(Quotient / Divisor) * Divisor
    If Modulo < 0 Then Modulo = Modulo + Divisor
End Function
