VERSION 5.00
Begin VB.Form FormDiffuse 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Custom Diffuse"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   285
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
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   10
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsY 
      Height          =   255
      Left            =   240
      Max             =   10
      TabIndex        =   5
      Top             =   4680
      Value           =   5
      Width           =   5055
   End
   Begin VB.HScrollBar hsX 
      Height          =   255
      Left            =   240
      Max             =   10
      TabIndex        =   3
      Top             =   3840
      Value           =   5
      Width           =   5055
   End
   Begin VB.TextBox txtX 
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
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      Text            =   "0"
      Top             =   3810
      Width           =   615
   End
   Begin VB.TextBox txtY 
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
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      Text            =   "0"
      Top             =   4650
      Width           =   615
   End
   Begin VB.CheckBox ChkWrap 
      Appearance      =   0  'Flat
      Caption         =   "wrap edge values"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   5880
      Width           =   1245
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   5880
      Width           =   1245
   End
   Begin VB.Label lblAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "after"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3360
      TabIndex        =   12
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblBefore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "before"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vertical strength:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "horizontal strength:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   2085
   End
End
Attribute VB_Name = "FormDiffuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Diffuse Filter Handler
'Copyright ©2000-2013 by Tanner Helland
'Created: 8/14/01
'Last updated: 09/September/12
'Last update: rewrote all effects against new layer class
'
'Module for handling the diffusion-style filters.  Automates both saturated
'and wrapped diffusion.
'
'***************************************************************************

Option Explicit

'When previewing, we need to modify the strength to be representative of the final filter.  This means dividing by the
' original image width in order to establish the right ratio.
Dim iWidth As Long, iHeight As Long

Private Sub ChkWrap_Click()
    If ChkWrap.Value = vbChecked Then DiffuseCustom hsX.Value, hsY.Value, True, True, picEffect Else DiffuseCustom hsX.Value, hsY.Value, False, True, picEffect
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    
    'The max and min values of the scroll bars are used to validate the range of the text box
    If EntryValid(txtX, hsX.Min, hsX.Max) Then
        If EntryValid(txtY, hsY.Min, hsY.Max) Then
            
            FormDiffuse.Visible = False
            
            If ChkWrap.Value = vbChecked Then
                Process CustomDiffuse, hsX.Value, hsY.Value, True
            Else
                Process CustomDiffuse, hsX.Value, hsY.Value, False
            End If
            
            Unload Me
            
        Else
            AutoSelectText txtY
        End If
    Else
        AutoSelectText txtX
    End If
    
End Sub

Private Sub Form_Activate()
    
    'Note the current image's width and height, which will be needed to adjust the preview effect
    iWidth = pdImages(CurrentImage).Width
    iHeight = pdImages(CurrentImage).Height
    
    'Adjust the scroll bar dimensions to match the current image's width and height
    hsX.Max = pdImages(CurrentImage).Width - 1
    hsY.Max = pdImages(CurrentImage).Height - 1
    hsX.Value = hsX.Max \ 2
    hsY.Value = hsY.Max \ 2
    
    'Draw the left preview box (the original image)
    DrawPreviewImage picPreview
    
    'Draw the right preview box (the diffused effect)
    If ChkWrap.Value = vbChecked Then DiffuseCustom hsX.Value, hsY.Value, True, True, picEffect Else DiffuseCustom hsX.Value, hsY.Value, False, True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'Everything below this line relates to mirroring the input of the textboxes across the scrollbars (and vice versa)
Private Sub hsX_Change()
    copyToTextBoxI txtX, hsX.Value
    If ChkWrap.Value = vbChecked Then DiffuseCustom hsX.Value, hsY.Value, True, True, picEffect Else DiffuseCustom hsX.Value, hsY.Value, False, True, picEffect
End Sub

Private Sub hsX_Scroll()
    copyToTextBoxI txtX, hsX.Value
    If ChkWrap.Value = vbChecked Then DiffuseCustom hsX.Value, hsY.Value, True, True, picEffect Else DiffuseCustom hsX.Value, hsY.Value, False, True, picEffect
End Sub

Private Sub hsY_Change()
    copyToTextBoxI txtY, hsY.Value
    If ChkWrap.Value = vbChecked Then DiffuseCustom hsX.Value, hsY.Value, True, True, picEffect Else DiffuseCustom hsX.Value, hsY.Value, False, True, picEffect
End Sub

Private Sub hsY_Scroll()
    copyToTextBoxI txtY, hsY.Value
    If ChkWrap.Value = vbChecked Then DiffuseCustom hsX.Value, hsY.Value, True, True, picEffect Else DiffuseCustom hsX.Value, hsY.Value, False, True, picEffect
End Sub

Private Sub txtX_GotFocus()
    AutoSelectText txtX
End Sub

Private Sub txtX_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtX
    If EntryValid(txtX, hsX.Min, hsX.Max, False, False) Then hsX.Value = Val(txtX)
End Sub

Private Sub txtY_GotFocus()
    AutoSelectText txtY
End Sub

Private Sub txtY_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtY
    If EntryValid(txtY, hsY.Min, hsY.Max, False, False) Then hsY.Value = Val(txtY)
End Sub

'Custom diffuse effect
' Inputs: diameter in x direction, diameter in y direction, whether or not to wrap edge pixels, and optional preview settings
Public Sub DiffuseCustom(ByVal xDiffuse As Long, ByVal yDiffuse As Long, ByVal wrapPixels As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Simulating large image explosion..."
    
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
    
    'If this is a preview, we need to adjust the xDiffuse and yDiffuse values to match the size of the preview box
    If toPreview Then
        xDiffuse = (xDiffuse / iWidth) * curLayerValues.Width
        yDiffuse = (yDiffuse / iHeight) * curLayerValues.Height
    End If
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValDiffuseX As Long, QuickValDiffuseY As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    Dim MaxX As Long
    MaxX = finalX * qvDepth
        
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()

    'Seed the random number generator with a pseudo-random value (the number of milliseconds elapsed since midnight)
    Randomize Timer
    
    'hDX and hDY are the half-values (or radius) of the diffuse area.  Pre-calculating them is faster than recalculating
    ' them every time we need to access a radius value.
    Dim hDX As Single, hDY As Single
    hDX = xDiffuse / 2
    hDY = yDiffuse / 2
    
    'Finally, these two variables will be used to store the position of diffused pixels
    Dim DiffuseX As Long, DiffuseY As Long
    
    'Loop through each pixel in the image, diffusing as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        DiffuseX = Rnd * xDiffuse - hDX
        DiffuseY = Rnd * yDiffuse - hDY
        
        QuickValDiffuseX = (DiffuseX * qvDepth) + QuickVal
        QuickValDiffuseY = DiffuseY + y
            
        'Make sure the diffused pixel is within image boundaries, and if not adjust it according to the user's
        ' "wrapPixels" setting.
        If wrapPixels Then
            If QuickValDiffuseX < 0 Then QuickValDiffuseX = QuickValDiffuseX + MaxX
            If QuickValDiffuseY < 0 Then QuickValDiffuseY = QuickValDiffuseY + finalY
            
            If QuickValDiffuseX > MaxX Then QuickValDiffuseX = QuickValDiffuseX - MaxX
            If QuickValDiffuseY > finalY Then QuickValDiffuseY = QuickValDiffuseY - finalY
        Else
            If QuickValDiffuseX < 0 Then QuickValDiffuseX = 0
            If QuickValDiffuseY < 0 Then QuickValDiffuseY = 0
            
            If QuickValDiffuseX > MaxX Then QuickValDiffuseX = MaxX
            If QuickValDiffuseY > finalY Then QuickValDiffuseY = finalY
        End If
            
        dstImageData(QuickVal + 2, y) = srcImageData(QuickValDiffuseX + 2, QuickValDiffuseY)
        dstImageData(QuickVal + 1, y) = srcImageData(QuickValDiffuseX + 1, QuickValDiffuseY)
        dstImageData(QuickVal, y) = srcImageData(QuickValDiffuseX, QuickValDiffuseY)

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

'Diffuse an image with a small radius (2)
Public Sub Diffuse()

    Message "Simulating small image explosion..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA
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
    Dim QuickVal As Long, QuickValDiffuseX As Long, QuickValDiffuseY As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    Dim xCheck As Long, yCheck As Long
    xCheck = curLayerValues.MaxX * qvDepth
    yCheck = curLayerValues.MaxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()

    'Seed the random number generator with a pseudo-random value (the number of milliseconds elapsed since midnight)
    Randomize Timer
    
    'Finally, these two variables will be used to store the position of diffused pixels
    Dim DiffuseX As Long, DiffuseY As Long
    
    'Loop through each pixel in the image, diffusing as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        DiffuseX = Rnd * 4 - 2
        DiffuseY = Rnd * 4 - 2
        
        QuickValDiffuseX = (DiffuseX * qvDepth) + QuickVal
        QuickValDiffuseY = DiffuseY + y
            
        'Make sure the diffused pixel is within image boundaries
        If QuickValDiffuseX < 0 Then QuickValDiffuseX = 0
        If QuickValDiffuseY < 0 Then QuickValDiffuseY = 0
        
        If QuickValDiffuseX > xCheck Then QuickValDiffuseX = xCheck
        If QuickValDiffuseY > yCheck Then QuickValDiffuseY = yCheck
        
        dstImageData(QuickVal + 2, y) = srcImageData(QuickValDiffuseX + 2, QuickValDiffuseY)
        dstImageData(QuickVal + 1, y) = srcImageData(QuickValDiffuseX + 1, QuickValDiffuseY)
        dstImageData(QuickVal, y) = srcImageData(QuickValDiffuseX, QuickValDiffuseY)

    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'Diffuse an image with a larger radius (6)
Public Sub DiffuseMore()

    Message "Simulating large image explosion..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA
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
    Dim QuickVal As Long, QuickValDiffuseX As Long, QuickValDiffuseY As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    Dim MaxX As Long
    MaxX = finalX * qvDepth
        
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()

    'Seed the random number generator with a pseudo-random value (the number of milliseconds elapsed since midnight)
    Randomize Timer
    
    'Finally, these two variables will be used to store the position of diffused pixels
    Dim DiffuseX As Long, DiffuseY As Long
    
    'Loop through each pixel in the image, diffusing as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        DiffuseX = Rnd * 12 - 6
        DiffuseY = Rnd * 12 - 6
        
        QuickValDiffuseX = (DiffuseX * qvDepth) + QuickVal
        QuickValDiffuseY = DiffuseY + y
            
        'Make sure the diffused pixel is within image boundaries
        If QuickValDiffuseX < 0 Then QuickValDiffuseX = 0
        If QuickValDiffuseY < 0 Then QuickValDiffuseY = 0
        
        If QuickValDiffuseX > MaxX Then QuickValDiffuseX = MaxX
        If QuickValDiffuseY > finalY Then QuickValDiffuseY = finalY
        
        dstImageData(QuickVal + 2, y) = srcImageData(QuickValDiffuseX + 2, QuickValDiffuseY)
        dstImageData(QuickVal + 1, y) = srcImageData(QuickValDiffuseX + 1, QuickValDiffuseY)
        dstImageData(QuickVal, y) = srcImageData(QuickValDiffuseX, QuickValDiffuseY)

    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub
