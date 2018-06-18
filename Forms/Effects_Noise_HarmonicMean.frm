VERSION 5.00
Begin VB.Form FormHarmonicMean 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Harmonic Mean"
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
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sltRadius 
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
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdSlider sltRadius 
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
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdButtonStrip btsKernelShape 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   3240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
      Caption         =   "kernel shape"
   End
End
Attribute VB_Name = "FormHarmonicMean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Harmonic mean Tool
'Copyright 2013-2018 by Tanner Helland
'Created: 27/July/17
'Last updated: 27/July/17
'Last update: initial build
'
'This is a heavily optimized "harmonic mean" function.  An accumulation technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform quite well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before applying any
' filter of a large radius (> 20).
'
'Harmonic mean is an edge-preserving noise removal filter.  It calculate the harmonic mean
' (https://en.wikipedia.org/wiki/Harmonic_mean) for a region around each pixel, and sets pixel values accordingly.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub ApplyHarmonicMean(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString parameterList
    
    Dim hRadius As Double, vRadius As Double, kernelShape As PD_PIXEL_REGION_SHAPE
    hRadius = cParams.GetDouble("radius-x", 1#)
    vRadius = cParams.GetDouble("radius-y", hRadius)
    kernelShape = cParams.GetLong("kernelshape", PDPRS_Rectangle)
    
    If (Not toPreview) Then Message "Applying harmonic mean filter..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent medianred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
            
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
    If (Not toPreview) Then
        SetProgBarMax finalX
        progBarCheck = ProgressBars.FindBestProgBarValue()
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
    
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    Dim directionDown As Boolean
    directionDown = True
    
    Dim pxSum As Double, pxCount As Long
    Dim finalR As Double, finalG As Double, finalB As Double
    
    'Prebuild a lookup table for all possible (1 / i) values.  To allow us to process the case of i=0
    ' (e.g. black pixels), we increment all values by 1.0, then subtract 1.0 in the inner loop, after the
    ' mean has been calculated.
    Dim oneDiv() As Double
    ReDim oneDiv(0 To 255) As Double
    For i = 0 To 255
        oneDiv(i) = 1# / (i + 1)
    Next i
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator(srcDIB, hRadius, vRadius, kernelShape) Then
    
        numOfPixels = cPixelIterator.LockTargetHistograms_RGBA(rValues, gValues, bValues, aValues, False)
        
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
                
                'With histograms successfully calculated, we can now find the harmonic mean for this pixel.
                
                'Loop through each color component histogram, and average all non-zero pixels found
                pxSum = 0#
                pxCount = 0
                
                For i = 0 To 255
                    pxSum = pxSum + rValues(i) * oneDiv(i)
                    pxCount = pxCount + rValues(i)
                Next i
                If (pxSum > 0#) Then finalR = pxCount / pxSum Else finalR = 1#
                
                'Repeat for green and blue
                pxSum = 0#
                pxCount = 0
                
                For i = 0 To 255
                    pxSum = pxSum + gValues(i) * oneDiv(i)
                    pxCount = pxCount + gValues(i)
                Next i
                If (pxSum > 0#) Then finalG = pxCount / pxSum Else finalG = 1#
                
                pxSum = 0#
                pxCount = 0
                
                For i = 0 To 255
                    pxSum = pxSum + bValues(i) * oneDiv(i)
                    pxCount = pxCount + bValues(i)
                Next i
                
                If (pxSum > 0#) Then finalB = pxCount / pxSum Else finalB = 1#
                
                'Subtract one from the calculated average (which is how we compensate for black pixels),
                ' then perform a failsafe upper-bound check.  (Lower bound is guaranteed safe.)
                finalR = finalR - 1
                finalG = finalG - 1
                finalB = finalB - 1
                If (finalR > 255#) Then finalR = 255#
                If (finalR < 0#) Then finalR = 0#
                If (finalG > 255#) Then finalG = 255#
                If (finalG < 0#) Then finalG = 0#
                If (finalB > 255#) Then finalB = 255#
                If (finalB < 0#) Then finalB = 0#
                
                'Update the pixel data in the destination image with our final result(s)
                dstImageData(x, y) = finalB
                dstImageData(x + 1, y) = finalG
                dstImageData(x + 2, y) = finalR
                
                'Move the iterator in the correct direction
                If directionDown Then
                    If (y < finalY) Then numOfPixels = cPixelIterator.MoveYDown
                Else
                    If (y > initY) Then numOfPixels = cPixelIterator.MoveYUp
                End If
                
            Next y
            
            'Reverse y-directionality on each pass
            directionDown = Not directionDown
            If (x < finalX) Then numOfPixels = cPixelIterator.MoveXRight
            
            'Update the progress bar every (progBarCheck) lines
            If (Not toPreview) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal x
                End If
            End If
                
        Next x
        
        'Release the pixel iterator
        cPixelIterator.ReleaseTargetHistograms_RGBA rValues, gValues, bValues, aValues
        
        'Release our local array that points to the target DIB
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        
        'Erase our temporary DIB
        Set srcDIB = Nothing
    
        'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
        EffectPrep.FinalizeImageData toPreview, dstPic
        
    End If

End Sub

Private Sub btsKernelShape_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Harmonic mean", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews while we initialize everything
    cmdBar.MarkPreviewStatus False
    
    'Populate the kernel shape box with whatever shapes PD currently supports
    Interface.PopKernelShapeButtonStrip btsKernelShape, PDPRS_Rectangle
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyHarmonicMean GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltRadius_Change(Index As Integer)
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("radius-x", sltRadius(0).Value, "radius-y", sltRadius(1).Value, "kernelshape", btsKernelShape.ListIndex)
End Function
