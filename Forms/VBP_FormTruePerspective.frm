VERSION 5.00
Begin VB.Form FormTruePerspective 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " True Perspective Test"
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
   Begin VB.HScrollBar hsRatioX 
      Height          =   255
      Index           =   3
      LargeChange     =   10
      Left            =   6120
      Max             =   100
      Min             =   -100
      TabIndex        =   19
      Top             =   1800
      Width           =   2535
   End
   Begin VB.HScrollBar hsRatioY 
      Height          =   255
      Index           =   3
      LargeChange     =   10
      Left            =   6120
      Max             =   100
      Min             =   -100
      TabIndex        =   18
      Top             =   2160
      Width           =   2535
   End
   Begin VB.HScrollBar hsRatioX 
      Height          =   255
      Index           =   2
      LargeChange     =   10
      Left            =   9240
      Max             =   100
      Min             =   -100
      TabIndex        =   16
      Top             =   1800
      Width           =   2535
   End
   Begin VB.HScrollBar hsRatioY 
      Height          =   255
      Index           =   2
      LargeChange     =   10
      Left            =   9240
      Max             =   100
      Min             =   -100
      TabIndex        =   15
      Top             =   2160
      Width           =   2535
   End
   Begin VB.HScrollBar hsRatioX 
      Height          =   255
      Index           =   1
      LargeChange     =   10
      Left            =   9240
      Max             =   100
      Min             =   -100
      TabIndex        =   13
      Top             =   480
      Width           =   2535
   End
   Begin VB.HScrollBar hsRatioY 
      Height          =   255
      Index           =   1
      LargeChange     =   10
      Left            =   9240
      Max             =   100
      Min             =   -100
      TabIndex        =   12
      Top             =   840
      Width           =   2535
   End
   Begin VB.HScrollBar hsRatioY 
      Height          =   255
      Index           =   0
      LargeChange     =   10
      Left            =   6120
      Max             =   100
      Min             =   -100
      TabIndex        =   11
      Top             =   840
      Width           =   2535
   End
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
      TabIndex        =   7
      Top             =   3495
      Width           =   5700
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
   Begin VB.HScrollBar hsRatioX 
      Height          =   255
      Index           =   0
      LargeChange     =   10
      Left            =   6120
      Max             =   100
      Min             =   -100
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   0
      Left            =   6120
      TabIndex        =   9
      Top             =   4440
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
      TabIndex        =   10
      Top             =   4440
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
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "bottom left:"
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
      TabIndex        =   20
      Top             =   1440
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "bottom right:"
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
      Left            =   9120
      TabIndex        =   17
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "top right:"
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
      Left            =   9120
      TabIndex        =   14
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "if pixels lie outside the image..."
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
      TabIndex        =   8
      Top             =   3120
      Width           =   3315
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   12135
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
      TabIndex        =   4
      Top             =   4050
      Width           =   1845
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "top left:"
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
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "FormTruePerspective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Perspective Distortion
'Copyright ©2012-2013 by Tanner Helland
'Created: 04/April/13
'Last updated: 04/April/13
'Last update: initial build
'
'This tool allows the user to apply forced perspective to an image.  The code is similar (in theory) to the
' shearing algorithm used in FormShear.  Reverse-mapping is used to allow for high-quality antialiasing.
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

    Me.Visible = False
    
    Dim paramString As String
    paramString = ""
    Dim i As Long
    
    For i = 0 To 3
        paramString = paramString & CStr(CDbl(hsRatioX(i) / 100))
        paramString = paramString & "|" & CStr(CDbl(hsRatioY(i) / 100))
        If i < 3 Then paramString = paramString & "|"
    Next i
    
    'Based on the user's selection, submit the proper processor request
    Process FreePerspective, paramString, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value
    
    Unload Me
    
End Sub

'Apply horizontal and/or vertical perspective to an image by shrinking it in one or more directions
' Input: xRatio, a value from -100 to 100 that specifies the horizontal perspective
'        yRatio, same as xRatio but for vertical perspective
Public Sub TruePerspectiveImage(ByVal listOfModifiers As String, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If toPreview = False Then Message "Applying new perspective..."
    
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
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, curLayerValues.MaxX, curLayerValues.MaxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Parse the incoming parameter string into individual (x, y) pairs
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(listOfModifiers) > 0 Then cParams.setParamString listOfModifiers
    
    'In the final version of this function, the user will be able to specify all of these points.  For now,
    ' we generate them manually
    Dim x0 As Double, x1 As Double, x2 As Double, x3 As Double
    Dim y0 As Double, y1 As Double, y2 As Double, y3 As Double
        
    x0 = 0 - cParams.GetDouble(1)
    y0 = 0 + cParams.GetDouble(2)
    x1 = 1 - cParams.GetDouble(3)
    y1 = 0 + cParams.GetDouble(4)
    x2 = 1 - cParams.GetDouble(5)
    y2 = 1 + cParams.GetDouble(6)
    x3 = 0 - cParams.GetDouble(7)
    y3 = 1 + cParams.GetDouble(8)
    
    'First things first: we need to map the original image (in terms of the unit square)
    ' to the arbitrary quadrilateral defined by the user's parameters
    Dim dx1 As Double, dy1 As Double, dx2 As Double, dy2 As Double, dx3 As Double, dy3 As Double
    dx1 = x1 - x2
    dy1 = y1 - y2
    dx2 = x3 - x2
    dy2 = y3 - y2
    dx3 = x0 - x1 + x2 - x3
    dy3 = y0 - y1 + y2 - y3
        
    Dim a11 As Double, a21 As Double, a31 As Double
    Dim a12 As Double, a22 As Double, a32 As Double
    Dim a13 As Double, a23 As Double, a33 As Double
        
    a13 = (dx3 * dy2 - dx2 * dy3) / (dx1 * dy2 - dy1 * dx2)
    a23 = (dx1 * dy3 - dy1 * dx3) / (dx1 * dy2 - dy1 * dx2)
    a11 = x1 - x0 + a13 * x1
    a21 = x3 - x0 + a23 * x3
    a31 = x0
    a12 = y1 - y0 + a13 * y1
    a22 = y3 - y0 + a23 * y3
    a32 = y0
    
    a33 = 1
    
    'Next, we need to generate an inverse transformation (so we can reverse-map and use interpolation for better results)
    Dim ta11 As Double, ta21 As Double, ta31 As Double
    Dim ta12 As Double, ta22 As Double, ta32 As Double
    Dim ta13 As Double, ta23 As Double, ta33 As Double
    
    ta11 = a22 * a33 - a32 * a23
    ta21 = a32 * a13 - a12 * a33
    ta31 = a12 * a23 - a22 * a13
    ta12 = a31 * a23 - a21 * a33
    ta22 = a11 * a33 - a31 * a13
    ta32 = a21 * a13 - a11 * a23
    ta13 = a21 * a32 - a31 * a22
    ta23 = a31 * a12 - a11 * a32
    ta33 = a11 * a22 - a21 * a12
        
    Dim tmpF As Double
    tmpF = 1 / ta33

    a11 = ta11 * tmpF
    a21 = ta12 * tmpF
    a31 = ta13 * tmpF
    a12 = ta21 * tmpF
    a22 = ta22 * tmpF
    a32 = ta23 * tmpF
    a13 = ta31 * tmpF
    a23 = ta32 * tmpF
    a33 = 1
    
    'Next, we need to calculate the key set of transformation parameters, using the reverse-map data we just generated.
    Dim vA As Double, VB As Double, vC As Double, vD As Double, vE As Double, vF As Double, VG As Double, vH As Double, vI As Double
    
    vA = a22 * a33 - a32 * a23
    VB = a31 * a23 - a21 * a33
    vC = a21 * a32 - a31 * a22
    vD = a32 * a13 - a12 * a33
    vE = a11 * a33 - a31 * a13
    vF = a31 * a12 - a11 * a32
    VG = a12 * a23 - a22 * a13
    vH = a21 * a13 - a11 * a23
    vI = a11 * a22 - a21 * a12
    
    'Store region width and height as floating-point
    Dim imgWidth As Double, imgHeight As Double
    imgWidth = finalX - initX
    imgHeight = finalY - initY
    
    'Scale quad coordinates to the size of the image
    Dim invWidth As Double, invHeight As Double
    invWidth = 1 / imgWidth
    invHeight = 1 / imgHeight
    
    vA = vA * invWidth
    vD = vD * invWidth
    VG = VG * invWidth
    
    VB = VB * invHeight
    vE = vE * invHeight
    vH = vH * invHeight
    
    'Certain values can lead to divide-by-zero problems - check those in advance and convert 0 to something like 0.000001
    Dim chkDenom As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
                
        'Reverse-map the coordinates back onto the original image (to allow for resampling)
        chkDenom = (VG * x + vH * y + vI)
        If chkDenom = 0 Then chkDenom = 0.000000001
        srcX = imgWidth * (vA * x + VB * y + vC) / chkDenom
        srcY = imgHeight * (vD * x + vE * y + vF) / chkDenom
                
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
        
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    popDistortEdgeBox cmbEdges, EDGE_ERASE
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
        
    'Create the preview
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub hsRatioX_Change(Index As Integer)
    updatePreview
End Sub

Private Sub hsRatioX_Scroll(Index As Integer)
    updatePreview
End Sub

Private Sub hsRatioY_Change(Index As Integer)
    updatePreview
End Sub

Private Sub hsRatioY_Scroll(Index As Integer)
    updatePreview
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()

    Dim paramString As String
    paramString = ""
    Dim i As Long
    
    For i = 0 To 3
        paramString = paramString & CStr(CDbl(hsRatioX(i) / 100))
        paramString = paramString & "|" & CStr(CDbl(hsRatioY(i) / 100))
        If i < 3 Then paramString = paramString & "|"
    Next i

    TruePerspectiveImage paramString, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value, True, fxPreview
    
End Sub
