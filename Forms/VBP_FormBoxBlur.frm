VERSION 5.00
Begin VB.Form FormBoxBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Box Blur"
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
   Begin VB.TextBox txtWidth 
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
      TabIndex        =   8
      Text            =   "2"
      Top             =   2220
      Width           =   615
   End
   Begin VB.TextBox txtHeight 
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
      TabIndex        =   7
      Text            =   "2"
      Top             =   3180
      Width           =   615
   End
   Begin VB.HScrollBar hsWidth 
      Height          =   255
      Left            =   6120
      Max             =   500
      Min             =   1
      TabIndex        =   6
      Top             =   2280
      Value           =   2
      Width           =   4935
   End
   Begin VB.HScrollBar hsHeight 
      Height          =   255
      Left            =   6120
      Max             =   500
      Min             =   1
      TabIndex        =   5
      Top             =   3240
      Value           =   2
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
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartCheckBox chkUnison 
      Height          =   480
      Left            =   6120
      TabIndex        =   11
      Top             =   3840
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   847
      Caption         =   "keep both dimensions in sync"
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "box height:"
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
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "box width:"
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
      TabIndex        =   9
      Top             =   1920
      Width           =   1140
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
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   5775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   12135
   End
End
Attribute VB_Name = "FormBoxBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Box Blur Tool
'Copyright ©2000-2013 by Tanner Helland
'Created: some time 2000
'Last updated: 17/January/13
'Last update: rewrote as a full tool, instead of two 3x3 and 5x5 individual filters
'
'This is a heavily optimized box blur.  An "accumulation" technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before
' applying any box blur of a large radius.
'
'***************************************************************************

Option Explicit

'When previewing, we need to modify the strength to be representative of the final filter.  This means dividing by the
' original image dimensions in order to establish the right ratio.
Dim iWidth As Long, iHeight As Long

Dim userChange As Boolean

Private Sub chkUnison_Click()
    userChange = False
    If CBool(chkUnison) Then hsHeight.Value = hsWidth.Value
    userChange = True
    updatePreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'Validate text box entries
    If Not EntryValid(txtWidth, hsWidth.Min, hsWidth.Max, True, True) Then
        AutoSelectText txtWidth
        Exit Sub
    End If
    
    If Not EntryValid(txtHeight, hsHeight.Min, hsHeight.Max, True, True) Then
        AutoSelectText txtHeight
        Exit Sub
    End If
    
    Me.Visible = False
    Process BoxBlur, hsWidth.Value, hsHeight.Value
    Unload Me
    
End Sub

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub BoxBlurFilter(ByVal hRadius As Long, ByVal vRadius As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Applying box blur to image..."
        
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
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
        
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        hRadius = (hRadius / iWidth) * curLayerValues.Width
        vRadius = (vRadius / iHeight) * curLayerValues.Height
        If hRadius = 0 Then hRadius = 1
        If vRadius = 0 Then vRadius = 1
    End If
    
    Dim xRadius As Long, yRadius As Long
    xRadius = hRadius
    yRadius = vRadius
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If xRadius > (finalX - initX) Then xRadius = finalX - initX
    If yRadius > (finalY - initY) Then yRadius = finalY - initY
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValInner As Long, QuickY As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'The number of pixels in the current blur box are tracked dynamically.
    Dim NumOfPixels As Long
    NumOfPixels = 0
            
    'Blurring takes a lot of variables
    Dim rTotal As Long, gTotal As Long, bTotal As Long, aTotal As Long
    Dim lbX As Long, lbY As Long, ubX As Long, ubY As Long
    Dim obuX As Boolean, obuY As Boolean, oblY As Boolean
    Dim i As Long, j As Long
    
    Dim atBottom As Boolean
    atBottom = True
    
    Dim startY As Long, stopY As Long, yStep As Long
    
    rTotal = 0: gTotal = 0: bTotal = 0: aTotal = 0
    NumOfPixels = 0
    
    'Generate an initial array of blur data for the first pixel
    For X = initX To initX + xRadius - 1
        QuickVal = X * qvDepth
    For Y = initY To initY + yRadius '- 1
    
        rTotal = rTotal + srcImageData(QuickVal + 2, Y)
        gTotal = gTotal + srcImageData(QuickVal + 1, Y)
        bTotal = bTotal + srcImageData(QuickVal, Y)
        If qvDepth = 4 Then aTotal = aTotal + srcImageData(QuickVal + 3, Y)
        
        'Increase the pixel tally
        NumOfPixels = NumOfPixels + 1
        
    Next Y
    Next X
                
    'Loop through each pixel in the image, tallying blur values as we go
    For X = initX To finalX
            
        QuickVal = X * qvDepth
        
        'Determine the bounds of the current blur box in the X direction
        lbX = X - xRadius
        If lbX < 0 Then lbX = 0
        ubX = X + xRadius
        
        If ubX > finalX Then
            obuX = True
            ubX = finalX
        Else
            obuX = False
        End If
                
        'As part of my accumulation algorithm, I swap the inner loop's direction with each iteration.
        ' Set y-related loop variables depending on the direction of the next cycle.
        If atBottom Then
            lbY = 0
            ubY = yRadius
        Else
            lbY = finalY - yRadius
            ubY = finalY
        End If
        
        'Remove trailing values from the blur box if they lie outside the processing radius
        If lbX > 0 Then
        
            QuickValInner = (lbX - 1) * qvDepth
        
            For j = lbY To ubY
                rTotal = rTotal - srcImageData(QuickValInner + 2, j)
                gTotal = gTotal - srcImageData(QuickValInner + 1, j)
                bTotal = bTotal - srcImageData(QuickValInner, j)
                If qvDepth = 4 Then aTotal = aTotal - srcImageData(QuickValInner + 3, j)
                NumOfPixels = NumOfPixels - 1
            Next j
        
        End If
        
        'Add leading values to the blur box if they lie inside the processing radius
        If Not obuX Then
        
            QuickValInner = ubX * qvDepth
            
            For j = lbY To ubY
                rTotal = rTotal + srcImageData(QuickValInner + 2, j)
                gTotal = gTotal + srcImageData(QuickValInner + 1, j)
                bTotal = bTotal + srcImageData(QuickValInner, j)
                If qvDepth = 4 Then aTotal = aTotal + srcImageData(QuickValInner + 3, j)
                NumOfPixels = NumOfPixels + 1
            Next j
            
        End If
        
        'Depending on the direction we are moving, remove a line of pixels from the blur box
        ' (because the interior loop will add it back in).
        If atBottom Then
                
            For i = lbX To ubX
                QuickValInner = i * qvDepth
                rTotal = rTotal - srcImageData(QuickValInner + 2, yRadius)
                gTotal = gTotal - srcImageData(QuickValInner + 1, yRadius)
                bTotal = bTotal - srcImageData(QuickValInner, yRadius)
                If qvDepth = 4 Then aTotal = aTotal - srcImageData(QuickValInner + 3, yRadius)
                NumOfPixels = NumOfPixels - 1
            Next i
        
        Else
        
            QuickY = finalY - yRadius
        
            For i = lbX To ubX
                QuickValInner = i * qvDepth
                rTotal = rTotal - srcImageData(QuickValInner + 2, QuickY)
                gTotal = gTotal - srcImageData(QuickValInner + 1, QuickY)
                bTotal = bTotal - srcImageData(QuickValInner, QuickY)
                If qvDepth = 4 Then aTotal = aTotal - srcImageData(QuickValInner + 3, QuickY)
                NumOfPixels = NumOfPixels - 1
            Next i
        
        End If
        
        'Based on the direction we're traveling, reverse the interior loop boundaries as necessary.
        If atBottom Then
            startY = 0
            stopY = finalY
            yStep = 1
        Else
            startY = finalY
            stopY = 0
            yStep = -1
        End If
            
    'Process the next column.  This step is pretty much identical to the row steps above (but in a vertical direction, obviously)
    For Y = startY To stopY Step yStep
            
        'If we are at the bottom and moving up, we will REMOVE rows from the bottom and ADD them at the top.
        'If we are at the top and moving down, we will REMOVE rows from the top and ADD them at the bottom.
        'As such, there are two copies of this function, one per possible direction.
        If atBottom Then
        
            'Calculate bounds
            lbY = Y - yRadius
            If lbY < 0 Then lbY = 0
            
            ubY = Y + yRadius
            If ubY > finalY Then
                obuY = True
                ubY = finalY
            Else
                obuY = False
            End If
                                
            'Remove trailing values from the box
            If lbY > 0 Then
            
                QuickY = lbY - 1
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    rTotal = rTotal - srcImageData(QuickValInner + 2, QuickY)
                    gTotal = gTotal - srcImageData(QuickValInner + 1, QuickY)
                    bTotal = bTotal - srcImageData(QuickValInner, QuickY)
                    If qvDepth = 4 Then aTotal = aTotal - srcImageData(QuickValInner + 3, QuickY)
                    NumOfPixels = NumOfPixels - 1
                Next i
                        
            End If
                    
            'Add leading values
            If Not obuY Then
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    rTotal = rTotal + srcImageData(QuickValInner + 2, ubY)
                    gTotal = gTotal + srcImageData(QuickValInner + 1, ubY)
                    bTotal = bTotal + srcImageData(QuickValInner, ubY)
                    If qvDepth = 4 Then aTotal = aTotal + srcImageData(QuickValInner + 3, ubY)
                    NumOfPixels = NumOfPixels + 1
                Next i
            
            End If
            
        'The exact same code as above, but in the opposite direction
        Else
        
            lbY = Y - yRadius
            If lbY < 0 Then
                oblY = True
                lbY = 0
            Else
                oblY = False
            End If
            
            ubY = Y + yRadius
            If ubY > finalY Then ubY = finalY
                                
            If ubY < finalY Then
            
                QuickY = ubY + 1
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    rTotal = rTotal - srcImageData(QuickValInner + 2, QuickY)
                    gTotal = gTotal - srcImageData(QuickValInner + 1, QuickY)
                    bTotal = bTotal - srcImageData(QuickValInner, QuickY)
                    If qvDepth = 4 Then aTotal = aTotal - srcImageData(QuickValInner + 3, QuickY)
                    NumOfPixels = NumOfPixels - 1
                Next i
                        
            End If
                    
            If Not oblY Then
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    rTotal = rTotal + srcImageData(QuickValInner + 2, lbY)
                    gTotal = gTotal + srcImageData(QuickValInner + 1, lbY)
                    bTotal = bTotal + srcImageData(QuickValInner, lbY)
                    If qvDepth = 4 Then aTotal = aTotal + srcImageData(QuickValInner + 3, lbY)
                    NumOfPixels = NumOfPixels + 1
                Next i
            
            End If
        
        End If
                
        'With the blur box successfully calculated, we can finally apply the results to the image.
        dstImageData(QuickVal + 2, Y) = rTotal \ NumOfPixels
        dstImageData(QuickVal + 1, Y) = gTotal \ NumOfPixels
        dstImageData(QuickVal, Y) = bTotal \ NumOfPixels
        If qvDepth = 4 Then dstImageData(QuickVal + 3, Y) = aTotal \ NumOfPixels
    
    Next Y
        atBottom = Not atBottom
        If toPreview = False Then
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

    userChange = True

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

'These routines keep the scroll bar and text box values in sync

Private Sub hsHeight_Change()
    userChange = False
    copyToTextBoxI txtHeight, hsHeight.Value
    If CBool(chkUnison) Then syncScrollBars False
    userChange = True
    updatePreview
End Sub

Private Sub hsWidth_Change()
    userChange = False
    copyToTextBoxI txtWidth, hsWidth.Value
    If CBool(chkUnison) Then syncScrollBars True
    userChange = True
    updatePreview
End Sub

Private Sub hsHeight_Scroll()
    userChange = False
    copyToTextBoxI txtHeight, hsHeight.Value
    If CBool(chkUnison) Then syncScrollBars False
    userChange = True
    updatePreview
End Sub

Private Sub hsWidth_Scroll()
    userChange = False
    copyToTextBoxI txtWidth, hsWidth.Value
    If CBool(chkUnison) Then syncScrollBars True
    userChange = True
    updatePreview
End Sub

Private Sub txtHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    userChange = False
    textValidate txtHeight
    If EntryValid(txtHeight, hsHeight.Min, hsHeight.Max, False, False) Then hsHeight.Value = Val(txtHeight)
    userChange = True
    updatePreview
End Sub

Private Sub txtHeight_GotFocus()
    AutoSelectText txtHeight
End Sub

Private Sub txtWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    userChange = False
    textValidate txtWidth
    If EntryValid(txtWidth, hsWidth.Min, hsWidth.Max, False, False) Then hsWidth.Value = Val(txtWidth)
    userChange = True
    updatePreview
End Sub

Private Sub txtWidth_GotFocus()
    AutoSelectText txtWidth
End Sub

'Keep the two scroll bars in sync.  Some extra work has to be done to makes sure scrollbar max values aren't exceeded.
Private Sub syncScrollBars(ByVal srcHorizontal As Boolean)
    
    If hsWidth.Value = hsHeight.Value Then Exit Sub
    
    Dim tmpVal As Long
    
    If srcHorizontal Then
        tmpVal = hsWidth.Value
        If tmpVal < hsHeight.Max Then hsHeight.Value = hsWidth.Value Else hsHeight.Value = hsHeight.Max
    Else
        tmpVal = hsHeight.Value
        If tmpVal < hsWidth.Max Then hsWidth.Value = hsHeight.Value Else hsWidth.Value = hsWidth.Max
    End If
    
End Sub
Private Sub updatePreview()
    BoxBlurFilter hsWidth.Value, hsHeight.Value, True, fxPreview
End Sub

