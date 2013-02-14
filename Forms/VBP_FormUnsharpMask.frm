VERSION 5.00
Begin VB.Form FormUnsharpMask 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Unsharp Masking"
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
   Begin VB.TextBox txtAmount 
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
      MaxLength       =   4
      TabIndex        =   12
      Text            =   "1.0"
      Top             =   2715
      Width           =   615
   End
   Begin VB.HScrollBar hsAmount 
      Height          =   255
      Left            =   6120
      Max             =   110
      Min             =   11
      TabIndex        =   11
      Top             =   2760
      Value           =   21
      Width           =   4935
   End
   Begin VB.HScrollBar hsThreshold 
      Height          =   255
      Left            =   6120
      Max             =   255
      TabIndex        =   9
      Top             =   3720
      Width           =   4935
   End
   Begin VB.TextBox txtThreshold 
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
      TabIndex        =   8
      Text            =   "0"
      Top             =   3660
      Width           =   615
   End
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
      Top             =   1800
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
      Top             =   1740
      Width           =   615
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
   Begin VB.Label lblAmount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "amount:"
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
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "threshold:"
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
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Label lblIDEWarning 
      BackStyle       =   0  'Transparent
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
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "FormUnsharpMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Unsharp Masking Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 03/March/01
'Last updated: 17/January/13
'Last update: rewrote as a full tool, instead of a single hard-coded 5x5 implementation
'
'To my knowledge, this tool is the first of its kind in VB6 - a variable radius Unsharp Mask filter
' that utilizes all three traditional controls (radius, amount, and threshold) and is based on a
' true Gaussian kernel.

'The use of separable kernels makes this much, much faster than a standard unsharp mask function.  The
' exact speed gain for a P x Q kernel is PQ/(P + Q) - so for a radius of 4 (which is an actual kernel
' of 9x9) the processing time is 4.5x faster.  For a radius of 100, this is 100x faster than a
' traditional method.
'
'Despite this, it's still quite slow in the IDE.  I STRONGLY recommend compiling the project before
' applying any action at a large radius.
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

    'Validate all text box entries
    If Not EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, True, True) Then
        AutoSelectText txtRadius
        Exit Sub
    End If
    
    If Not EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max, True, True) Then
        AutoSelectText txtThreshold
        Exit Sub
    End If
    
    If Not EntryValid(txtAmount, CSng(hsAmount.Min - 10) / 10, CSng(hsAmount.Max - 10) / 10, True, True) Then
        AutoSelectText txtAmount
        Exit Sub
    End If

    Me.Visible = False
    Process Unsharp, hsRadius, hsAmount, hsThreshold
    Unload Me
    
End Sub

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub UnsharpMask(ByVal umRadius As Long, ByVal umAmount As Long, ByVal umThreshold As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
        
    If Not toPreview Then Message "Applying unsharp mask (step %1 of %2)...", 1, 2
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
            
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        If iWidth > iHeight Then
            umRadius = (umRadius / iWidth) * curLayerValues.Width
        Else
            umRadius = (umRadius / iHeight) * curLayerValues.Height
        End If
        If umRadius = 0 Then umRadius = 1
    End If
    
    CreateGaussianBlurLayer umRadius, workingLayer, srcLayer, toPreview, finalY * 2 + finalX
    
    'Now that we have a gaussian layer created in workingLayer, we can point arrays toward it and the source layer
    Dim dstImageData() As Byte
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
        
    If Not toPreview Then Message "Applying unsharp mask (step %1 of %2)...", 2, 2
        
    'ScaleFactor is used to apply the unsharp mask.  Maximum strength can be any value, but PhotoDemon locks it at 10.
    Dim scaleFactor As Double, invScaleFactor As Double
    scaleFactor = CDbl(umAmount) / 10
    invScaleFactor = 1 - scaleFactor

    Dim blendVal As Double
    
    'More color variables - in this case, sums for each color component
    Dim r As Long, g As Long, b As Long, a As Long
    Dim r2 As Long, g2 As Long, b2 As Long, a2 As Long
    Dim newR As Long, newG As Long, newB As Long, newA As Long
    Dim tLumDelta As Long
    
    umThreshold = umThreshold \ 5
    
    'The final step of the smart blur function is to find edges, and replace them with the blurred data as necessary
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Retrieve the original image's pixels
        r = dstImageData(QuickVal + 2, y)
        g = dstImageData(QuickVal + 1, y)
        b = dstImageData(QuickVal, y)
        
        'Now, retrieve the gaussian pixels
        r2 = srcImageData(QuickVal + 2, y)
        g2 = srcImageData(QuickVal + 1, y)
        b2 = srcImageData(QuickVal, y)
        
        tLumDelta = Abs(getLuminance(r, g, b) - getLuminance(r2, g2, b2))
                        
        'If the delta is below the specified threshold, sharpen it
        If tLumDelta > umThreshold Then
                        
            newR = (scaleFactor * r) + (invScaleFactor * r2)
            If newR > 255 Then newR = 255
            If newR < 0 Then newR = 0
                
            newG = (scaleFactor * g) + (invScaleFactor * g2)
            If newG > 255 Then newG = 255
            If newG < 0 Then newG = 0
                
            newB = (scaleFactor * b) + (invScaleFactor * b2)
            If newB > 255 Then newB = 255
            If newB < 0 Then newB = 0
            
            blendVal = tLumDelta / 255
            
            newR = BlendColors(newR, r, blendVal)
            newG = BlendColors(newG, g, blendVal)
            newB = BlendColors(newB, b, blendVal)
            
            dstImageData(QuickVal + 2, y) = newR
            dstImageData(QuickVal + 1, y) = newG
            dstImageData(QuickVal, y) = newB
            
            If qvDepth = 4 Then
                a2 = srcImageData(QuickVal + 3, y)
                a = dstImageData(QuickVal + 3, y)
                newA = (scaleFactor * a) + (invScaleFactor * a2)
                If newA > 255 Then newA = 255
                If newA < 0 Then newA = 0
                dstImageData(QuickVal + 3, y) = BlendColors(newA, a, blendVal)
            End If
            
        End If
                
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x + (finalY * 2)
        End If
    Next x
    
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    srcLayer.eraseLayer
    Set srcLayer = Nothing
    
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
    If Not g_IsProgramCompiled Then
        hsRadius.Max = 50
        lblIDEWarning.Caption = g_Language.TranslateMessage("WARNING!  This tool has been heavily optimized, but at high radius values it will still be quite slow inside the IDE.  Please compile before applying or previewing any radius larger than 20.")
        lblIDEWarning.Visible = True
    Else
        '32bpp images take quite a bit longer to process.  Limit the radius to 100 in this case.
        If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then hsRadius.Max = 100 Else hsRadius.Max = 200
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When the horizontal scroll bar is moved, change the text box to match
Private Sub hsAmount_Change()
    copyToTextBoxF CSng(hsAmount - 10) / 10, txtAmount, 1
    updatePreview
End Sub

Private Sub hsAmount_Scroll()
    copyToTextBoxF CSng(hsAmount - 10) / 10, txtAmount, 1
    updatePreview
End Sub

Private Sub txtAmount_GotFocus()
    AutoSelectText txtAmount
End Sub

Private Sub txtAmount_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtAmount, , True
    If EntryValid(txtAmount, CSng(hsAmount.Min - 10) / 10, CSng(hsAmount.Max - 10) / 10, False, False) Then
        hsAmount = Val(txtAmount) * 10
    End If
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

Private Sub hsThreshold_Change()
    copyToTextBoxI txtThreshold, hsThreshold.Value
    updatePreview
End Sub

Private Sub hsThreshold_Scroll()
    copyToTextBoxI txtThreshold, hsThreshold.Value
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

Private Sub txtThreshold_GotFocus()
    AutoSelectText txtThreshold
End Sub

Private Sub txtThreshold_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtThreshold
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max, False, False) Then
        hsThreshold.Value = Val(txtThreshold)
    End If
End Sub

Private Sub updatePreview()
    UnsharpMask hsRadius, hsAmount, hsThreshold, True, fxPreview
End Sub
