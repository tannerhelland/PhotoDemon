VERSION 5.00
Begin VB.Form FormLensCorrect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Correct or Fix Lens Distortion"
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
   Begin VB.ComboBox cmbEdges 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3735
      Width           =   4860
   End
   Begin VB.HScrollBar hsZoom 
      Height          =   255
      LargeChange     =   10
      Left            =   6120
      Max             =   300
      Min             =   100
      TabIndex        =   12
      Top             =   2100
      Value           =   150
      Width           =   4815
   End
   Begin VB.TextBox txtZoom 
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
      TabIndex        =   11
      Text            =   "1.50"
      Top             =   2040
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
      TabIndex        =   7
      Text            =   "100"
      Top             =   2820
      Width           =   735
   End
   Begin VB.HScrollBar hsRadius 
      Height          =   255
      Left            =   6120
      Max             =   100
      Min             =   1
      TabIndex        =   6
      Top             =   2880
      Value           =   100
      Width           =   4815
   End
   Begin VB.TextBox txtStrength 
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
      Text            =   "3.00"
      Top             =   1200
      Width           =   735
   End
   Begin VB.HScrollBar hsStrength 
      Height          =   255
      LargeChange     =   10
      Left            =   6120
      Max             =   1000
      Min             =   100
      TabIndex        =   3
      Top             =   1260
      Value           =   300
      Width           =   4815
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   0
      Left            =   6120
      TabIndex        =   16
      Top             =   4680
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      Caption         =   "quality"
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
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   1
      Left            =   7920
      TabIndex        =   17
      Top             =   4680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      Caption         =   "speed"
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "if pixels lie outside the corrected area..."
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
      Left            =   6000
      TabIndex        =   15
      Top             =   3360
      Width           =   4170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "correction zoom:"
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
      TabIndex        =   13
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   2520
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
      Top             =   4290
      Width           =   1845
   End
   Begin VB.Label lblStrength 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "correction strength:"
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
      Top             =   840
      Width           =   2085
   End
End
Attribute VB_Name = "FormLensCorrect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Fix Lens Distort Tool
'Copyright ©2011-2013 by Tanner Helland
'Created: 22/January/13
'Last updated: 22/January/13
'Last update: initial build
'
'This tool allows the user to correct an existing lens distortion on an image.  Bilinear interpolation
' (via reverse-mapping) is available for a higher quality correction.
'
'A zoom parameter is also provided to help the user determine how much of the image they are willing
' to sacrifice as part of the correction.  If the distort is quite high, there is no real way to
' correct the image without cutting off parts of it (see sample images at http://photo.net/learn/fisheye/).
'
'For optimal quality, I suggest zooming out a ways, applying the correction, then cropping the resultant
' image to the desired shape.
'
'***************************************************************************

Option Explicit

'Use this to prevent the text box and scroll bar from updating each other in an endless loop
Dim userChange As Boolean

Private Sub cmbEdges_Click()
    updatePreview
End Sub

Private Sub cmbEdges_Scroll()
    updatePreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'Before rendering anything, check to make sure the text boxes have valid input
    If Not EntryValid(txtStrength * 100, hsStrength.Min, hsStrength.Max, True, True) Then
        AutoSelectText txtStrength
        Exit Sub
    End If

    If Not EntryValid(txtZoom * 100, hsZoom.Min, hsZoom.Max, True, True) Then
        AutoSelectText txtZoom
        Exit Sub
    End If

    If Not EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, True, True) Then
        AutoSelectText txtRadius
        Exit Sub
    End If

    Me.Visible = False
    
    'Based on the user's selection, submit the proper processor request
    Process DistortLensFix, CDbl(hsStrength / 100), CDbl(hsZoom / 100), hsRadius.Value, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value
    
    Unload Me
    
End Sub

'Correct lens distortion in an image
Public Sub ApplyLensCorrection(ByVal fixStrength As Double, ByVal fixZoom As Double, ByVal lensRadius As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Correcting image distortion..."
    
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
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
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
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, curLayerValues.MaxX, curLayerValues.MaxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
          
    'Lens distort correction requires a number of specialized variables
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
    
    'Rotation values
    Dim theta As Double, sRadius As Double, sRadius2 As Double, sDistance As Double
    Dim r As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curLayerValues.Width
    tHeight = curLayerValues.Height
    sRadius = Sqr(tWidth * tWidth + tHeight * tHeight) / 2
              
    Dim refDistance As Double
    refDistance = sRadius * 2 / fixStrength
              
    sRadius = sRadius * (lensRadius / 100)
    sRadius2 = sRadius * sRadius
              
    'Loop through each pixel in the image, converting values as we go
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
                            
        'Remap the coordinates around a center point of (0, 0)
        nX = X - midX
        nY = Y - midY
        
        'Calculate distance automatically
        sDistance = (nX * nX) + (nY * nY)
        
        If sDistance <= sRadius2 Then
            
            sDistance = Sqr(sDistance)
            r = sDistance / refDistance
            
            If r = 0 Then theta = 1 Else theta = Atn(r) / r
            srcX = midX + theta * nX * fixZoom
            srcY = midY + theta * nY * fixZoom
            
        Else
        
            srcX = X
            srcY = Y
            
        End If
        
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        fSupport.setPixels X, Y, srcX, srcY, srcImageData, dstImageData
                
    Next Y
        If Not toPreview Then
            If (X And progBarCheck) = 0 Then SetProgBarVal X
        End If
    Next X
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
End Sub

Private Sub Form_Activate()
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    popDistortEdgeBox cmbEdges, EDGE_ERASE
    
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
Private Sub hsStrength_Change()
    If userChange Then
        txtStrength.Text = Format(CDbl(hsStrength.Value) / 100, "0.00")
        txtStrength.Refresh
    End If
    updatePreview
End Sub

Private Sub hsStrength_Scroll()
    txtStrength.Text = Format(CDbl(hsStrength.Value) / 100, "0.00")
    txtStrength.Refresh
    updatePreview
End Sub

'Keep the scroll bar and the text box values in sync
Private Sub hsZoom_Change()
    If userChange Then
        txtZoom.Text = Format(CDbl(hsZoom.Value) / 100, "0.00")
        txtZoom.Refresh
    End If
    updatePreview
End Sub

Private Sub hsZoom_Scroll()
    txtZoom.Text = Format(CDbl(hsZoom.Value) / 100, "0.00")
    txtZoom.Refresh
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

Private Sub txtStrength_GotFocus()
    AutoSelectText txtStrength
End Sub

Private Sub txtStrength_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtStrength, True, True
    If EntryValid(txtStrength, hsStrength.Min / 100, hsStrength.Max / 100, False, False) Then
        userChange = False
        hsStrength.Value = Val(txtStrength) * 100
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

Private Sub txtZoom_GotFocus()
    AutoSelectText txtZoom
End Sub

Private Sub txtZoom_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtZoom
    If EntryValid(txtZoom, hsZoom.Min, hsZoom.Max, False, False) Then hsZoom.Value = Val(txtZoom)
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()

    ApplyLensCorrection CDbl(hsStrength / 100), CDbl(hsZoom / 100), hsRadius.Value, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value, True, fxPreview
    
End Sub
