VERSION 5.00
Begin VB.Form FormPerspective 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Perspective"
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
   Begin VB.HScrollBar hsRatioY 
      Height          =   255
      LargeChange     =   10
      Left            =   6120
      Max             =   1000
      Min             =   -1000
      TabIndex        =   13
      Top             =   2700
      Width           =   4815
   End
   Begin VB.TextBox txtRatioY 
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
      TabIndex        =   12
      Text            =   "0.0"
      Top             =   2640
      Width           =   735
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
      TabIndex        =   8
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
   Begin VB.TextBox txtRatioX 
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
      Top             =   1800
      Width           =   735
   End
   Begin VB.HScrollBar hsRatioX 
      Height          =   255
      LargeChange     =   10
      Left            =   6120
      Max             =   1000
      Min             =   -1000
      TabIndex        =   3
      Top             =   1860
      Width           =   4815
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   0
      Left            =   6120
      TabIndex        =   10
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
      TabIndex        =   11
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
      Caption         =   "vertical perspective:"
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
      TabIndex        =   14
      Top             =   2280
      Width           =   2100
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
      TabIndex        =   9
      Top             =   3120
      Width           =   3315
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   4050
      Width           =   1845
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "horizontal perspective:"
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
      Top             =   1440
      Width           =   2400
   End
End
Attribute VB_Name = "FormPerspective"
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

    'Before rendering anything, check to make sure the text boxes have valid input
    If Not EntryValid(txtRatioX, hsRatioX.Min / 10, hsRatioX.Max / 10, True, True) Then
        AutoSelectText txtRatioX
        Exit Sub
    End If

    If Not EntryValid(txtRatioY, hsRatioY.Min / 10, hsRatioY.Max / 10, True, True) Then
        AutoSelectText txtRatioY
        Exit Sub
    End If

    Me.Visible = False
    
    'Based on the user's selection, submit the proper processor request
    Process DistortPerspective, CDbl(hsRatioX / 10), CDbl(hsRatioY / 10), CLng(cmbEdges.ListIndex), OptInterpolate(0).Value
    
    Unload Me
    
End Sub

'Apply horizontal and/or vertical perspective to an image by shrinking it in one or more directions
' Input: xRatio, a value from -100 to 100 that specifies the horizontal perspective
'        yRatio, same as xRatio but for vertical perspective
Public Sub PerspectiveImage(ByVal xRatio As Double, ByVal yRatio As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

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
        
    'Midpoints of the image in the x and y direction
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
    
    'Convert xRatio and yRatio to [0, 1] range
    xRatio = xRatio / 100
    yRatio = yRatio / 100
    
    'Store region width and height as floating-point
    Dim ImgWidth As Double, ImgHeight As Double
    ImgWidth = finalX - initX
    ImgHeight = finalY - initY
    
    'Build a look-up table for horizontal line size and offset
    Dim leftX() As Double, lineWidth() As Double
    ReDim leftX(initY To finalY) As Double
    ReDim lineWidth(initY To finalY) As Double
    
    For y = initY To finalY
            leftX(y) = ((finalY - y) / finalY) * midX * xRatio
            lineWidth(y) = ImgWidth - (leftX(y) * 2)
            If lineWidth(y) = 0 Then lineWidth(y) = 0.000000001
    Next y
    
    'Do the same for vertical line size and offset
    Dim topY() As Double, lineHeight() As Double
    ReDim topY(initX To finalX) As Double
    ReDim lineHeight(initX To finalX) As Double
    
    For x = initX To finalX
            topY(x) = ((finalX - x) / finalX) * midY * yRatio
            lineHeight(x) = ImgHeight - (topY(x) * 2)
            If lineHeight(x) = 0 Then lineHeight(x) = 0.000000001
    Next x
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
                
        'Reverse-map the coordinates back onto the original image (to allow for AA)
        srcX = ((x - leftX(y)) / lineWidth(y)) * ImgWidth
        srcY = ((y - topY(x)) / lineHeight(x)) * ImgHeight
        
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
    popDistortEdgeBox cmbEdges, EDGE_WRAP
        
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

'Keep the scroll bar and the text box values in sync
Private Sub hsRatioX_Change()
    If userChange Then
        txtRatioX.Text = Format(CDbl(hsRatioX.Value) / 10, "##0.0")
        txtRatioX.Refresh
    End If
    updatePreview
End Sub

Private Sub hsRatioX_Scroll()
    txtRatioX.Text = Format(CDbl(hsRatioX.Value) / 10, "##0.0")
    txtRatioX.Refresh
    updatePreview
End Sub

Private Sub hsRatioY_Change()
    If userChange Then
        txtRatioY.Text = Format(CDbl(hsRatioY.Value) / 10, "##0.0")
        txtRatioY.Refresh
    End If
    updatePreview
End Sub

Private Sub hsRatioY_Scroll()
    txtRatioY.Text = Format(CDbl(hsRatioY.Value) / 10, "##0.0")
    txtRatioY.Refresh
    updatePreview
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub txtRatioX_GotFocus()
    AutoSelectText txtRatioX
End Sub

Private Sub txtRatioX_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtRatioX, True, True
    If EntryValid(txtRatioX, hsRatioX.Min / 10, hsRatioX.Max / 10, False, False) Then
        userChange = False
        hsRatioX.Value = Val(txtRatioX) * 10
        userChange = True
    End If
End Sub

Private Sub txtRatioY_GotFocus()
    AutoSelectText txtRatioY
End Sub

Private Sub txtRatioY_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtRatioY, True, True
    If EntryValid(txtRatioY, hsRatioY.Min / 10, hsRatioY.Max / 10, False, False) Then
        userChange = False
        hsRatioY.Value = Val(txtRatioY) * 10
        userChange = True
    End If
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()

    PerspectiveImage CDbl(hsRatioX / 10), CDbl(hsRatioY / 10), CLng(cmbEdges.ListIndex), OptInterpolate(0).Value, True, fxPreview
    
End Sub
