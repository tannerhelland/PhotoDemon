VERSION 5.00
Begin VB.Form FormMeanShift 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Mean shift"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   1
      Max             =   50
      Value           =   5
   End
   Begin PhotoDemon.sliderTextCombo sltThreshold 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "threshold"
      Min             =   1
      Max             =   50
      Value           =   15
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
Attribute VB_Name = "FormMeanShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Mean Shift Effect Tool
'Copyright 2015-2015 by Tanner Helland
'Created: 02/October/15
'Last updated: 08/December/15
'Last update: convert to the new pdPixelIterator class
'
'Mean shift filter, heavily optimized.  Wiki has a nice summary of this technique:
' https://en.wikipedia.org/wiki/Mean_shift
'
'Note that an "accumulation" technique is used instead of the standard sliding window mechanism.  The pdPixelIterator
' class handles this portion of the code, so all this filter has to do is process the resulting histograms.
'
'As with all area-based filters, this function will be slow inside the IDE.  I STRONGLY recommend compiling before using.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Sub ApplyMeanShiftFilter(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.setParamString parameterList
    
    Dim mRadius As Long, mThreshold As Long, kernelShape As PD_PIXEL_REGION_SHAPE
    mRadius = cParams.GetLong("radius", 1&)
    mThreshold = cParams.GetLong("threshold", 0&)
    kernelShape = cParams.GetLong("kernelShape", PDPRS_Rectangle)
    
    If Not toPreview Then Message "Applying mean shift filter..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second copy of the target DIB.
    ' (This is necessary to prevent processed pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        mRadius = mRadius * curDIBValues.previewModifier
        If mRadius < 1 Then mRadius = 1
    End If
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If (finalY - initY) < (finalX - initX) Then
        If mRadius > (finalY - initY) Then mRadius = finalY - initY
    Else
        If mRadius > (finalX - initX) Then mRadius = finalX - initX
    End If
        
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
    
    'We use an optimized histogram technique for calculating means, which means a lot of intermediate values are required
    Dim rValues() As Long, gValues() As Long, bValues() As Long, aValues() As Long
    ReDim rValues(0 To 255) As Long
    ReDim gValues(0 To 255) As Long
    ReDim bValues(0 To 255) As Long
    ReDim aValues(0 To 255) As Long
    
    Dim r As Long, g As Long, b As Long
    Dim lColor As Long, hColor As Long, cCount As Long
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    Dim directionDown As Boolean
    directionDown = True
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator(srcDIB, mRadius, mRadius, kernelShape) Then
    
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
            
                'With a local histogram successfully built for the area surrounding this pixel, we now need to find the
                ' shifted mean value.  We do this by averaging only the pixels whose color is within the caller-specified
                ' threshold of this pixel.  This limits the average to similar pixels only.
        
                'Blue
                
                'Start by finding low/high bounds
                lColor = dstImageData(x, y)
                hColor = lColor
                
                lColor = lColor - mThreshold
                If lColor < 0 Then lColor = 0
                hColor = hColor + mThreshold
                If hColor > 255 Then hColor = 255
                
                'Search from low to high, tallying colors as we go
                b = 0: cCount = 0
                For i = lColor To hColor
                    b = b + i * bValues(i)
                    cCount = cCount + bValues(i)
                Next i
                
                'Take the mean of this range of values
                If cCount > 0 Then b = b / cCount Else b = 255
                
                'Repeat for green
                lColor = dstImageData(x + 1, y)
                hColor = lColor
                
                lColor = lColor - mThreshold
                If lColor < 0 Then lColor = 0
                hColor = hColor + mThreshold
                If hColor > 255 Then hColor = 255
                
                g = 0: cCount = 0
                For i = lColor To hColor
                    g = g + i * gValues(i)
                    cCount = cCount + gValues(i)
                Next i
                
                If cCount > 0 Then g = g / cCount Else g = 255
                
                'Repeat for red
                lColor = dstImageData(x + 2, y)
                hColor = lColor
                
                lColor = lColor - mThreshold
                If lColor < 0 Then lColor = 0
                hColor = hColor + mThreshold
                If hColor > 255 Then hColor = 255
                
                r = 0: cCount = 0
                For i = lColor To hColor
                    r = r + i * rValues(i)
                    cCount = cCount + rValues(i)
                Next i
                
                If cCount > 0 Then r = r / cCount Else r = 255
                
                'Finally, apply the results to the image.
                dstImageData(x, y) = b
                dstImageData(x + 1, y) = g
                dstImageData(x + 2, y) = r
                
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
    Process "Mean shift", , GetLocalParamString(), UNDO_LAYER
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
    
    'Disable previews while we initialize the dialog
    cmdBar.markPreviewStatus False
    
    'Populate the kernel shape box with whatever shapes PD currently supports
    Interface.PopKernelShapeButtonStrip btsKernelShape, PDPRS_Rectangle
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltThreshold_Change()
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then Me.ApplyMeanShiftFilter GetLocalParamString(), True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = buildParamList("radius", sltRadius.Value, "threshold", sltThreshold.Value, "kernelShape", btsKernelShape.ListIndex)
End Function
