VERSION 5.00
Begin VB.Form FormGaussianBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Gaussian Blur"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9030
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10500
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin VB.HScrollBar hsRadius 
      Height          =   255
      Left            =   6120
      Max             =   200
      Min             =   1
      TabIndex        =   2
      Top             =   2760
      Value           =   5
      Width           =   4935
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
      Left            =   11160
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "5"
      Top             =   2700
      Width           =   615
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5625
      _extentx        =   9922
      _extenty        =   9922
   End
   Begin VB.Label lblIDEWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"VBP_FormGaussianBlur.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   6000
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   5775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "radius:"
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
      Top             =   2400
      Width           =   735
   End
End
Attribute VB_Name = "FormGaussianBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gaussian Blur Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 01/July/10
'Last updated: 17/January/13
'Last update: rewrote as a full tool, instead of two 3x3 and 5x5 individual filters
'
'To my knowledge, this tool is the first of its kind in VB6 - a variable radius gaussian blur filter
' that utilizes a separable convolution kernel.

'The use of separable kernels makes this much, much faster than a standard Gaussian blur.  The exact speed
' gain for a P x Q kernel is PQ/(P + Q) - so for a radius of 4 (which is an actual kernel of 9x9) the
' processing time is 4.5x faster.  For a radius of 100, this is 100x faster than a traditional method.
'
'Despite this, it's still quite slow in the IDE.  I STRONGLY recommend compiling the project before
' applying any Gaussian blur of a large radius.
'
'***************************************************************************

Option Explicit

'When previewing, we need to modify the strength to be representative of the final filter.  This means dividing by the
' original image dimensions in order to establish the right ratio.
Dim iWidth As Long, iHeight As Long

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    If EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, True, True) Then
        Me.Visible = False
        Process GaussianBlur, hsRadius.Value
        Unload Me
    Else
        AutoSelectText txtRadius
    End If
    
End Sub

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub GaussianBlurFilter(ByVal gRadius As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Convolving image with separable gaussian kernel..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
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
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        If iWidth > iHeight Then
            gRadius = (gRadius / iWidth) * curLayerValues.Width
        Else
            gRadius = (gRadius / iHeight) * curLayerValues.Height
        End If
        If gRadius = 0 Then gRadius = 1
    End If
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValInner As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    SetProgBarMax finalX * 2
    progBarCheck = findBestProgBarValue()
    
    'Create a one-dimensional Gaussian kernel using the requested radius
    Dim gKernel() As Single
    ReDim gKernel(-gRadius To gRadius) As Single
    
    Dim numPixels As Long
    numPixels = (gRadius * 2) + 1
    
    'Calculate a standard deviation (sigma) using the GIMP formula:
    Dim stdDev As Single, stdDev2 As Single
    If gRadius > 1 Then
        stdDev = Sqr(-(gRadius * gRadius) / (2 * Log(1# / 255#)))
    'Note that this is my addition - for a radius of 1 the GIMP formula results in too small of a sigma value
    Else
        stdDev = 0.5
    End If
    stdDev2 = stdDev * stdDev
    
    'Populate the kernel using that sigma
    Dim i As Long
    Dim curVal As Single, sumVal As Single
    sumVal = 0
    
    For i = -gRadius To gRadius
        curVal = (1 / (Sqr(PI_DOUBLE) * stdDev)) * (EULER ^ (-1 * ((i * i) / (2 * stdDev2))))
        sumVal = sumVal + curVal
        gKernel(i) = curVal
    Next i
        
    'Normalize the kernel so that all values sum to 1
    For i = -gRadius To gRadius
        gKernel(i) = gKernel(i) / sumVal
        Message i & ":" & gKernel(i) & ":" & stdDev
    Next i
    
    'We now have a normalized 1-dimensional gaussian kernel available for convolution.
    
    'Color variables - in this case, sums for each color component
    Dim rSum As Single, gSum As Single, bSum As Single, aSum As Single
    
    'We now convolve the image twice - once in the horizontal direction, then again in the vertical direction.  This is
    ' referred to as "separable" convolution, and it's much faster than than traditional convolution, especially for
    ' large radii (the exact speed gain for a P x Q kernel is PQ/(P + Q) - so for a radius of 4 (which is an actual kernel
    ' of 9x9) the processing time is 4.5x faster).
    
    'First, perform a horizontal convolution.
        
    Dim chkX As Long
    Dim curFactor As Single
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        rSum = 0
        gSum = 0
        bSum = 0
        aSum = 0
    
        'Apply the convolution to the source array.  (This is a little confusing because we need to convolve the image
        ' twice - so first we modify the source, then we use that to modify the destination on the second pass.)
        For i = -gRadius To gRadius
        
            curFactor = gKernel(i)
            chkX = x + i
            
            'We need to give special treatment to pixels that lie off the image
            If chkX < initX Then chkX = initX
            If chkX > finalX Then chkX = finalX
            
            QuickValInner = chkX * qvDepth
            
            rSum = rSum + dstImageData(QuickValInner + 2, y) * curFactor
            gSum = gSum + dstImageData(QuickValInner + 1, y) * curFactor
            bSum = bSum + dstImageData(QuickValInner, y) * curFactor
            If qvDepth = 4 Then aSum = aSum + dstImageData(QuickValInner + 3, y) * curFactor
                    
        Next i
        
        'We now have sums for each of red, green, blue (and potentially alpha).  Apply those values to the source array.
        srcImageData(QuickVal + 2, y) = rSum '\ numPixels
        srcImageData(QuickVal + 1, y) = gSum '\ numPixels
        srcImageData(QuickVal, y) = bSum '\ numPixels
        If qvDepth = 4 Then srcImageData(QuickVal + 3, y) = aSum '\ numPixels
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'The source array now contains a horizontally convolved image.  We now need to convolve it vertically.
    Dim chkY As Long
    
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        rSum = 0
        gSum = 0
        bSum = 0
        aSum = 0
    
        'Apply the convolution to the source array.  (This is a little confusing because we need to convolve the image
        ' twice - so first we modify the source, then we use that to modify the destination on the second pass.)
        For i = -gRadius To gRadius
        
            curFactor = gKernel(i)
            chkY = y + i
            
            'We need to give special treatment to pixels that lie off the image
            If chkY < initY Then chkY = initY
            If chkY > finalY Then chkY = finalY
                        
            rSum = rSum + srcImageData(QuickVal + 2, chkY) * curFactor
            gSum = gSum + srcImageData(QuickVal + 1, chkY) * curFactor
            bSum = bSum + srcImageData(QuickVal, chkY) * curFactor
            If qvDepth = 4 Then aSum = aSum + srcImageData(QuickVal + 3, chkY) * curFactor
                    
        Next i
        
        'We now have sums for each of red, green, blue (and potentially alpha).  Apply those values to the source array.
        dstImageData(QuickVal + 2, y) = rSum '\ numPixels
        dstImageData(QuickVal + 1, y) = gSum '\ numPixels
        dstImageData(QuickVal, y) = bSum '\ numPixels
        If qvDepth = 4 Then srcImageData(QuickVal + 3, y) = aSum '\ numPixels
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal (x + finalX)
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

    'Note the current image's width and height, which will be needed to adjust the preview effect
    iWidth = pdImages(CurrentImage).Width
    iHeight = pdImages(CurrentImage).Height

    'Draw a preview of the effect
    updatePreview
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'If the program is not compiled, display a special warning for this tool
    If Not g_IsProgramCompiled Then lblIDEWarning.Visible = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The next three routines keep the scroll bar and text box values in sync
Private Sub hsRadius_Change()
    copyToTextBoxI txtRadius, hsRadius.Value
    updatePreview
End Sub

Private Sub hsRadius_Scroll()
    copyToTextBoxI txtRadius, hsRadius.Value
    updatePreview
End Sub

Private Sub txtRadius_GotFocus()
    AutoSelectText txtRadius
End Sub

Private Sub txtRadius_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtRadius
    If EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, False, False) Then
        hsRadius.Value = Val(txtRadius)
    End If
End Sub

Private Sub updatePreview()
    GaussianBlurFilter hsRadius.Value, True, fxPreview
End Sub
