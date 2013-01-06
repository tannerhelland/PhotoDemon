VERSION 5.00
Begin VB.Form FormSwirl 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Swirl"
   ClientHeight    =   8730
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   6255
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
   ScaleHeight     =   582
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "100"
      Top             =   6780
      Width           =   735
   End
   Begin VB.HScrollBar hsRadius 
      Height          =   255
      Left            =   360
      Max             =   100
      Min             =   1
      TabIndex        =   9
      Top             =   6840
      Value           =   100
      Width           =   4815
   End
   Begin VB.OptionButton OptInterpolate 
      Appearance      =   0  'Flat
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
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   7590
      Width           =   1095
   End
   Begin VB.OptionButton OptInterpolate 
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   7590
      Value           =   -1  'True
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
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   5
      Text            =   "0.0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.HScrollBar hsAngle 
      Height          =   255
      LargeChange     =   10
      Left            =   360
      Max             =   1800
      Min             =   -1800
      TabIndex        =   4
      Top             =   6060
      Width           =   4815
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5100
      Left            =   240
      ScaleHeight     =   338
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   382
      TabIndex        =   3
      Top             =   240
      Width           =   5760
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   8040
      Width           =   1245
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   8040
      Width           =   1245
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
      Left            =   240
      TabIndex        =   11
      Top             =   6480
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
      Left            =   240
      TabIndex        =   6
      Top             =   7260
      Width           =   1845
   End
   Begin VB.Label lblAmount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "swirl angle (degrees):"
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
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   2310
   End
End
Attribute VB_Name = "FormSwirl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image "Swirl" Distortion
'Copyright ©2000-2013 by Tanner Helland
'Created: 05/January/13
'Last updated: 05/January/12
'Last update: initial build
'
'This tool allows the user to "swirl" an image at an arbitrary angle in 1/10 degree increments.  Bilinear interpolation
' (via reverse-mapping) is available for a high-quality swirl.
'
'At present, the tool assumes that you want to swirl the image around its center.
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
    If Not EntryValid(txtAngle, -180, 180, True, True) Then
        AutoSelectText txtAngle
        Exit Sub
    End If

    If Not EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, True, True) Then
        AutoSelectText txtRadius
        Exit Sub
    End If

    Me.Visible = False
    
    'Based on the user's selection, submit the proper processor request
    If OptInterpolate(0) Then
        Process DistortSwirl, CDbl(hsAngle / 10), hsRadius.Value, False
    Else
        Process DistortSwirl, CDbl(hsAngle / 10), hsRadius.Value, True
    End If
    
    Unload Me
    
End Sub

Public Sub SwirlImage(ByVal swirlAngle As Double, ByVal swirlRadius As Double, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    'Reverse the rotationAngle value so that POSITIVE values indicate CLOCKWISE rotation.
    swirlAngle = -(swirlAngle / 10)

    If toPreview = False Then Message "Swirling image round and round..."
    
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
          
    'Swirling requires some specialized variables
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
    
    'Rotation values
    Dim theta As Double, sRadius As Double, sRadius2 As Double, sDistance As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
    
    Dim i As Long
    
    'Radius is based off the smaller of the two dimensions - width or height
    Dim tWidth As Long, tHeight As Long
    tWidth = curLayerValues.Width
    tHeight = curLayerValues.Height
    sRadius = Sqr(tWidth * tWidth + tHeight * tHeight) / 2
              
    sRadius = sRadius * (swirlRadius / 100)
    sRadius2 = sRadius * sRadius
              
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Remap the coordinates around a center point of (0, 0)
        nX = x - midX
        nY = y - midY
        
        'Calculate distance automatically
        sDistance = (nX * nX) + (nY * nY)
                
        'Calculate remapped x and y values
        If sDistance > sRadius2 Then
            srcX = x
            srcY = y
        Else
        
            sDistance = Sqr(sDistance)
            
            'Calculate theta
            theta = Atan2(nY, nX) + swirlAngle * ((sRadius - sDistance) / sRadius)
        
            srcX = midX + (sDistance * Cos(theta))
            srcY = midY + (sDistance * Sin(theta))
            
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
    
    'Create the preview
    DrawPreviewImage picPreview
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Mark scroll bar changes as coming from the user
    userChange = True
    
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

'Redraw the on-screen preview of the rotated image
Private Sub updatePreview()

    If OptInterpolate(0) Then
        SwirlImage CDbl(hsAngle / 10), hsRadius.Value, False, True, picPreview
    Else
        SwirlImage CDbl(hsAngle / 10), hsRadius.Value, True, True, picPreview
    End If

End Sub

Private Sub txtRadius_GotFocus()
    AutoSelectText txtRadius
End Sub

Private Sub txtRadius_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtRadius
    If EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, False, False) Then hsRadius.Value = Val(txtRadius)
End Sub

'Return the arctangent of two values (rise / run)
Public Function Atan2(ByVal y As Double, ByVal x As Double) As Double
 
    If y > 0 Then
      If x >= y Then
        Atan2 = Atn(y / x)
      ElseIf x <= -y Then
        Atan2 = Atn(y / x) + PI
      Else
        Atan2 = PI / 2 - Atn(x / y)
      End If
    Else
      If x >= -y Then
        Atan2 = Atn(y / x)
      ElseIf x <= y Then
        Atan2 = Atn(y / x) - PI
      Else
        Atan2 = -Atn(x / y) - PI / 2
      End If
    End If
 
  End Function
