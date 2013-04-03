VERSION 5.00
Begin VB.Form FormSmartBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Smart Blur"
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
   Begin PhotoDemon.smartOptionButton optEdges 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   12
      Top             =   1800
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   635
      Caption         =   "smooth areas"
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
      TabIndex        =   9
      Text            =   "50"
      Top             =   3900
      Width           =   615
   End
   Begin VB.HScrollBar hsThreshold 
      Height          =   255
      Left            =   6120
      Max             =   255
      TabIndex        =   8
      Top             =   3960
      Value           =   50
      Width           =   4935
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
      Max             =   500
      Min             =   1
      TabIndex        =   2
      Top             =   3000
      Value           =   50
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
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "5.0"
      Top             =   2940
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
   Begin PhotoDemon.smartOptionButton optEdges 
      Height          =   360
      Index           =   1
      Left            =   6120
      TabIndex        =   13
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      Caption         =   "edges"
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
      Caption         =   "apply blur to:"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   1440
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
      Top             =   3600
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
      Height          =   1215
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
   Begin VB.Label lblTitle 
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
      Index           =   0
      Left            =   6000
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
End
Attribute VB_Name = "FormSmartBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Smart" Blur Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 17/January/13
'Last updated: 17/January/13
'Last update: initial build
'
'To my knowledge, this tool is the first of its kind in VB6 - an intelligent blur tool that selectively blurs
' edges differently from smooth areas of an image.  The user can specify the threshold to use, as well as whether
' to more strongly blur edges or smooth sections.
'
'The use of separable kernels helps this function remain swift, despite all the different things it's handling.
'
'Despite this, it's still quite slow in the IDE.  I STRONGLY recommend compiling the project before
' applying any actions at a large radius.
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
    If Not EntryValid(txtRadius, hsRadius.Min / 10, hsRadius.Max / 10, True, True) Then
        AutoSelectText txtRadius
        Exit Sub
    End If
    
    If Not EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max, True, True) Then
        AutoSelectText txtThreshold
        Exit Sub
    End If
    
    Me.Visible = False
    Process SmartBlur, CDbl(hsRadius.Value) / 10, hsThreshold.Value, optEdges(1)
    Unload Me
    
End Sub

'Convolve an image using a selective gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but processing speed obviously drops as the radius increases)
Public Sub SmartBlurFilter(ByVal gRadius As Double, ByVal gThreshold As Byte, ByVal smoothEdges As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Analyzing image in preparation for smart blur..."
            
    'More color variables - in this case, sums for each color component
    Dim r As Long, g As Long, b As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    Dim tDelta As Long
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    Dim gaussLayer As pdLayer
    Set gaussLayer = New pdLayer
    gaussLayer.createFromExistingLayer workingLayer
    
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
    End If
    
    CreateGaussianBlurLayer gRadius, srcLayer, gaussLayer, toPreview, finalY * 2 + finalX
        
    'Now that we have a gaussian layer created in gaussLayer, we can point arrays toward it and the source layer
    Dim dstImageData() As Byte
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    Dim GaussImageData() As Byte
    Dim gaussSA As SAFEARRAY2D
    prepSafeArray gaussSA, gaussLayer
    CopyMemory ByVal VarPtrArray(GaussImageData()), VarPtr(gaussSA), 4
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
        
    If toPreview = False Then Message "Applying smart blur..."
        
    Dim blendVal As Double
    
    'The final step of the smart blur function is to find edges, and replace them with the blurred data as necessary
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Retrieve the original image's pixels
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
        
        tDelta = (213 * r + 715 * g + 72 * b) \ 1000
        
        'Now, retrieve the gaussian pixels
        r2 = GaussImageData(QuickVal + 2, y)
        g2 = GaussImageData(QuickVal + 1, y)
        b2 = GaussImageData(QuickVal, y)
        
        'Calculate a delta between the two
        tDelta = tDelta - ((213 * r2 + 715 * g2 + 72 * b2) \ 1000)
        If tDelta < 0 Then tDelta = -tDelta
                
        'If the delta is below the specified threshold, replace it with the blurred data.
        If smoothEdges Then
        
            If tDelta > gThreshold Then
                If tDelta <> 0 Then blendVal = 1 - (gThreshold / tDelta) Else blendVal = 0
                dstImageData(QuickVal + 2, y) = BlendColors(srcImageData(QuickVal + 2, y), GaussImageData(QuickVal + 2, y), blendVal)
                dstImageData(QuickVal + 1, y) = BlendColors(srcImageData(QuickVal + 1, y), GaussImageData(QuickVal + 1, y), blendVal)
                dstImageData(QuickVal, y) = BlendColors(srcImageData(QuickVal, y), GaussImageData(QuickVal, y), blendVal)
                If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = BlendColors(srcImageData(QuickVal + 3, y), GaussImageData(QuickVal + 3, y), blendVal)
            End If
        
        Else
        
            If tDelta <= gThreshold Then
                If gThreshold <> 0 Then blendVal = 1 - (tDelta / gThreshold) Else blendVal = 1
                dstImageData(QuickVal + 2, y) = BlendColors(srcImageData(QuickVal + 2, y), GaussImageData(QuickVal + 2, y), blendVal)
                dstImageData(QuickVal + 1, y) = BlendColors(srcImageData(QuickVal + 1, y), GaussImageData(QuickVal + 1, y), blendVal)
                dstImageData(QuickVal, y) = BlendColors(srcImageData(QuickVal, y), GaussImageData(QuickVal, y), blendVal)
                If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = BlendColors(srcImageData(QuickVal + 3, y), GaussImageData(QuickVal + 3, y), blendVal)
            End If
        
        End If
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x + (finalY * 2)
        End If
    Next x
        
    'With our work complete, release all arrays
    CopyMemory ByVal VarPtrArray(GaussImageData), 0&, 4
    Erase GaussImageData
    
    gaussLayer.eraseLayer
    Set gaussLayer = Nothing
    
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
    If Not g_IsProgramCompiled Then
        lblIDEWarning.Caption = g_Language.TranslateMessage("WARNING!  This tool has been heavily optimized, but at high radius values it will still be quite slow inside the IDE.  Please compile before applying or previewing any radius larger than 20.")
        lblIDEWarning.Visible = True
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The next three routines keep the scroll bar and text box values in sync
Private Sub hsRadius_Change()
    copyToTextBoxF hsRadius.Value / 10, txtRadius, 1
    updatePreview
End Sub

Private Sub hsRadius_Scroll()
    copyToTextBoxF hsRadius.Value / 10, txtRadius, 1
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

Private Sub OptEdges_Click(Index As Integer)
    updatePreview
End Sub

Private Sub txtRadius_GotFocus()
    AutoSelectText txtRadius
End Sub

Private Sub txtRadius_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtRadius
    If EntryValid(txtRadius, hsRadius.Min / 10, hsRadius.Max / 10, False, False) Then
        hsRadius.Value = Val(txtRadius) * 10
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

'Render a new effect preview
Private Sub updatePreview()
    SmartBlurFilter CDbl(hsRadius.Value) / 10, hsThreshold.Value, optEdges(1), True, fxPreview
End Sub
