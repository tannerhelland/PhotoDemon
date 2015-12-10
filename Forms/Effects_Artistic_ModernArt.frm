VERSION 5.00
Begin VB.Form FormModernArt 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Modern art"
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   705
      Index           =   0
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "horizontal strength"
      Min             =   1
      Max             =   200
      Value           =   5
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   705
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "vertical strength"
      Min             =   1
      Max             =   200
      Value           =   5
   End
   Begin PhotoDemon.buttonStrip btsKernelShape 
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   3600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Left            =   6000
      Top             =   3240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "kernel shape"
      FontSize        =   12
   End
End
Attribute VB_Name = "FormModernArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Modern Art Tool
'Copyright 2013-2015 by Tanner Helland
'Created: 09/Feb/13
'Last updated: 23/November/15
'Last update: convert to XML parameter list
'
'This is a heavily optimized "extreme rank" function.  An accumulation technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before applying any
' filter of a large radius (> 20).
'
'Extreme rank is a function of my own creation.  Basically, it performs both a minimum and a maxmimum rank calculation,
' and then it sets the pixel to whichever value is further from the current one.  This leads to an odd cut-out or stencil
' look unlike any other filter I've seen.  I'm not sure how much utility such a function provides, but it's fun so I
' include it.  :)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a "modern art" filter to the current master image (basically a max/min rank algorithm, with some tweaks)
'Input: radius of the median (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub ApplyModernArt(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.setParamString parameterList
    
    Dim hRadius As Double, vRadius As Double, kernelShape As PD_PIXEL_REGION_SHAPE
    hRadius = cParams.GetDouble("hRadius", 1#)
    vRadius = cParams.GetDouble("vRadius", hRadius)
    kernelShape = cParams.GetLong("kernelShape", PDPRS_Rectangle)
    
    If Not toPreview Then Message "Applying modern art techniques..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent medianred pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
            
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
        
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        hRadius = hRadius * curDIBValues.previewModifier
        vRadius = vRadius * curDIBValues.previewModifier
    End If
    
    'Range-check the radius.  (During previews, the line of code above may cause the radius to drop to zero.)
    If (hRadius = 0) Then hRadius = 1
    If (vRadius = 0) Then vRadius = 1
    
    'Split the radius into integer-only components, and make sure each isn't larger than the image itself
    ' in that dimension.
    Dim xRadius As Long, yRadius As Long
    xRadius = hRadius: yRadius = vRadius
    If xRadius > (finalX - initX) Then xRadius = finalX - initX
    If yRadius > (finalY - initY) Then yRadius = finalY - initY
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'The x-dimension of the image has a stride of (width * 4) for 32-bit images; precalculate this, to spare us some
    ' processing time in the inner loop.
    initX = initX * qvDepth
    finalX = finalX * qvDepth
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not toPreview Then
        SetProgBarMax finalX
        progBarCheck = findBestProgBarValue()
    End If
    
    'The number of pixels in the current median box are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'Median filtering takes a lot of variables
    Dim rValues() As Long, gValues() As Long, bValues() As Long, aValues() As Long
    ReDim rValues(0 To 255) As Long
    ReDim gValues(0 To 255) As Long
    ReDim bValues(0 To 255) As Long
    ReDim aValues(0 To 255) As Long
    
    Dim cutoffTotal As Long
    Dim r As Long, g As Long, b As Long
    Dim lowR As Long, lowG As Long, lowB As Long
    Dim highR As Long, highG As Long, highB As Long
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    Dim directionDown As Boolean
    directionDown = True
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator(srcDIB, hRadius, vRadius, kernelShape) Then
    
        numOfPixels = cPixelIterator.LockTargetHistograms(rValues, gValues, bValues, aValues, False)
        
        'Loop through each pixel in the image, applying the filter as we go
        For x = initX To finalX Step qvDepth
            
            'Based on the direction we're traveling, reverse the interior loop boundaries as necessary.
            If directionDown Then
                startY = initY
                stopY = finalY
                yStep = 1
            Else
                startY = finalY
                stopY = initY
                yStep = -1
            End If
            
            'Process the next column.  This step is pretty much identical to the row steps above (but in a vertical direction, obviously)
            For y = startY To stopY Step yStep
                
                'With the median box successfully calculated, we can now find the actual median for this pixel.
                
                'Loop through each color component histogram, until we've passed the desired percentile of pixels
                lowR = 0
                lowG = 0
                lowB = 0
                cutoffTotal = 0.01 * numOfPixels
                If cutoffTotal = 0 Then cutoffTotal = 1
                
                i = 0
                Do
                    If rValues(i) <> 0 Then lowR = lowR + rValues(i)
                    i = i + 1
                Loop While (lowR < cutoffTotal)
                lowR = i - 1
                
                i = 0
                Do
                    If gValues(i) <> 0 Then lowG = lowG + gValues(i)
                    i = i + 1
                Loop While (lowG < cutoffTotal)
                lowG = i - 1
        
                i = 0
                Do
                    If bValues(i) <> 0 Then lowB = lowB + bValues(i)
                    i = i + 1
                Loop While (lowB < cutoffTotal)
                lowB = i - 1
                
                'Now do the same thing at the top of the histogram
                highR = 0
                highG = 0
                highB = 0
                cutoffTotal = 0.01 * numOfPixels
                If cutoffTotal = 0 Then cutoffTotal = 1
                
                i = 255
                Do
                    If rValues(i) <> 0 Then highR = highR + rValues(i)
                    i = i - 1
                Loop While (highR < cutoffTotal)
                highR = i + 1
                
                i = 255
                Do
                    If gValues(i) <> 0 Then highG = highG + gValues(i)
                    i = i - 1
                Loop While (highG < cutoffTotal)
                highG = i + 1
                
                i = 255
                Do
                    If bValues(i) <> 0 Then highB = highB + bValues(i)
                    i = i - 1
                Loop While (highB < cutoffTotal)
                highB = i + 1
                
                'Retrieve the original pixel data, and replace it with the processed result
                b = dstImageData(x, y)
                If Abs(lowB - b) > (highB - b) Then
                    dstImageData(x, y) = lowB
                Else
                    dstImageData(x, y) = highB
                End If
                
                g = dstImageData(x + 1, y)
                If Abs(lowG - g) > (highG - g) Then
                    dstImageData(x + 1, y) = lowG
                Else
                    dstImageData(x + 1, y) = highG
                End If
                
                r = dstImageData(x + 2, y)
                If Abs(lowR - r) > (highR - r) Then
                    dstImageData(x + 2, y) = lowR
                Else
                    dstImageData(x + 2, y) = highR
                End If
                
                'Move the iterator in the correct direction
                If directionDown Then
                    If y < finalY Then numOfPixels = cPixelIterator.MoveYDown
                Else
                    If y > initY Then numOfPixels = cPixelIterator.MoveYUp
                End If
                
            Next y
            
            'Reverse y-directionality on each pass
            directionDown = Not directionDown
            If x < finalX Then numOfPixels = cPixelIterator.MoveXRight
            
            'Update the progress bar every (progBarCheck) lines
            If Not toPreview Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal x
                End If
            End If
                
        Next x
        
        'Release the pixel iterator
        cPixelIterator.ReleaseTargetHistograms rValues, gValues, bValues, aValues
        
        'Release our local array that points to the target DIB
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        
        'Erase our temporary DIB
        srcDIB.eraseDIB
        Set srcDIB = Nothing
    
        'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
        finalizeImageData toPreview, dstPic
        
    End If

End Sub

Private Sub btsKernelShape_Click(ByVal buttonIndex As Long)
    updatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Modern art", , GetLocalParamString(), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
        
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()

    'Disable previews while we initialize everything
    cmdBar.markPreviewStatus False
    
    'Populate the kernel shape box with whatever shapes PD currently supports
    Interface.PopKernelShapeButtonStrip btsKernelShape, PDPRS_Rectangle
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyModernArt GetLocalParamString(), True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltRadius_Change(Index As Integer)
    updatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = buildParamList("hRadius", sltRadius(0).Value, "vRadius", sltRadius(1).Value, "kernelShape", btsKernelShape.ListIndex)
End Function
