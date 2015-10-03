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
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   1
      Max             =   50
      Value           =   5
   End
   Begin PhotoDemon.sliderTextCombo sltThreshold 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "threshold"
      Min             =   1
      Max             =   50
      Value           =   15
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
'Last updated: 02/October/15
'Last update: initial build
'
'Mean shift filter, heavily optimized.  Wiki has a nice summary of this technique:
' https://en.wikipedia.org/wiki/Mean_shift
'
'Note that an "accumulation" technique is used instead of the standard sliding window mechanism.
' (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform pretty well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before using.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a mean-shift filter to the active image.
'Inputs: - radius of the search area (min 1, no real max - but the scroll bar is maxed at 200 presently)
'        - threshold of valid mean contributors; for example, a threshold of "10" means that a color must be
'           within +/- 10 in order to contribute to the mean for a given pixel
Public Sub ApplyMeanShiftFilter(ByVal mRadius As Long, ByVal mThreshold As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying mean shift filter..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent processed pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent medianred pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
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
    Dim QuickVal As Long, QuickValInner As Long, QuickY As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
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
    Dim rValues(0 To 255) As Long, gValues(0 To 255) As Long, bValues(0 To 255) As Long
    Dim lbX As Long, lbY As Long, ubX As Long, ubY As Long
    Dim obuX As Boolean, obuY As Boolean, oblY As Boolean
    Dim startY As Long, stopY As Long, yStep As Long
    Dim i As Long, j As Long
    Dim r As Long, g As Long, b As Long
    Dim lColor As Long, hColor As Long, cCount As Long
    
    Dim atBottom As Boolean
    atBottom = True
    
    'Generate an initial array of median data for the first pixel
    For x = initX To initX + mRadius - 1
        QuickVal = x * qvDepth
    For y = initY To initY + mRadius
    
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
        
        rValues(r) = rValues(r) + 1
        gValues(g) = gValues(g) + 1
        bValues(b) = bValues(b) + 1
        
    Next y
    Next x
                
    'Loop through each pixel in the image, tallying median values as we go
    For x = initX To finalX
            
        QuickVal = x * qvDepth
        
        'Determine the bounds of the current median box in the X direction
        lbX = x - mRadius
        If lbX < 0 Then lbX = 0
        
        ubX = x + mRadius
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
            ubY = mRadius
        Else
            lbY = finalY - mRadius
            ubY = finalY
        End If
        
        'Remove trailing values from the median box if they lie outside the processing radius
        If lbX > 0 Then
        
            QuickValInner = (lbX - 1) * qvDepth
        
            For j = lbY To ubY
                r = srcImageData(QuickValInner + 2, j)
                g = srcImageData(QuickValInner + 1, j)
                b = srcImageData(QuickValInner, j)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
            Next j
        
        End If
        
        'Add leading values to the median box if they lie inside the processing radius
        If Not obuX Then
        
            QuickValInner = ubX * qvDepth
            
            For j = lbY To ubY
                r = srcImageData(QuickValInner + 2, j)
                g = srcImageData(QuickValInner + 1, j)
                b = srcImageData(QuickValInner, j)
                rValues(r) = rValues(r) + 1
                gValues(g) = gValues(g) + 1
                bValues(b) = bValues(b) + 1
            Next j
            
        End If
        
        'Depending on the direction we are moving, remove a line of pixels from the median box
        ' (because the interior loop will add it back in).
        If atBottom Then
        
            For i = lbX To ubX
                QuickValInner = i * qvDepth
                r = srcImageData(QuickValInner + 2, mRadius)
                g = srcImageData(QuickValInner + 1, mRadius)
                b = srcImageData(QuickValInner, mRadius)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
            Next i
       
        Else
       
            QuickY = finalY - mRadius
       
            For i = lbX To ubX
                QuickValInner = i * qvDepth
                r = srcImageData(QuickValInner + 2, QuickY)
                g = srcImageData(QuickValInner + 1, QuickY)
                b = srcImageData(QuickValInner, QuickY)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
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
    For y = startY To stopY Step yStep
            
        'If we are at the bottom and moving up, we will REMOVE rows from the bottom and ADD them at the top.
        'If we are at the top and moving down, we will REMOVE rows from the top and ADD them at the bottom.
        'As such, there are two copies of this function, one per possible direction.
        If atBottom Then
        
            'Calculate bounds
            lbY = y - mRadius
            If lbY < 0 Then lbY = 0
            
            ubY = y + mRadius
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
                    r = srcImageData(QuickValInner + 2, QuickY)
                    g = srcImageData(QuickValInner + 1, QuickY)
                    b = srcImageData(QuickValInner, QuickY)
                    rValues(r) = rValues(r) - 1
                    gValues(g) = gValues(g) - 1
                    bValues(b) = bValues(b) - 1
                Next i
                        
            End If
                    
            'Add leading values
            If Not obuY Then
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, ubY)
                    g = srcImageData(QuickValInner + 1, ubY)
                    b = srcImageData(QuickValInner, ubY)
                    rValues(r) = rValues(r) + 1
                    gValues(g) = gValues(g) + 1
                    bValues(b) = bValues(b) + 1
                Next i
            
            End If
            
        'The exact same code as above, but in the opposite direction
        Else
        
            lbY = y - mRadius
            If lbY < 0 Then
                oblY = True
                lbY = 0
            Else
                oblY = False
            End If
            
            ubY = y + mRadius
            If ubY > finalY Then ubY = finalY
                                
            If ubY < finalY Then
            
                QuickY = ubY + 1
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, QuickY)
                    g = srcImageData(QuickValInner + 1, QuickY)
                    b = srcImageData(QuickValInner, QuickY)
                    rValues(r) = rValues(r) - 1
                    gValues(g) = gValues(g) - 1
                    bValues(b) = bValues(b) - 1
                Next i
                        
            End If
                    
            If Not oblY Then
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, lbY)
                    g = srcImageData(QuickValInner + 1, lbY)
                    b = srcImageData(QuickValInner, lbY)
                    rValues(r) = rValues(r) + 1
                    gValues(g) = gValues(g) + 1
                    bValues(b) = bValues(b) + 1
                Next i
            
            End If
        
        End If
        
        'With a local histogram successfully built for the area surrounding this pixel, we now need to find the
        ' shifted mean value.  We do this by averaging only the pixels whose color is within the caller-specified
        ' threshold of this pixel.  This limits the average to similar pixels only.
        
        'Red
        
        'Start by finding low/high bounds
        lColor = srcImageData(QuickVal + 2, y)
        hColor = lColor
        
        lColor = lColor - mThreshold
        If lColor < 0 Then lColor = 0
        hColor = hColor + mThreshold
        If hColor > 255 Then hColor = 255
        
        'Search from low to high, tallying colors as we go
        r = 0: cCount = 0
        For i = lColor To hColor
            r = r + i * rValues(i)
            cCount = cCount + rValues(i)
        Next i
                
        'Take the mean of this range of values
        r = r / cCount
        
        'Repeat for green
        lColor = srcImageData(QuickVal + 1, y)
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
        
        g = g / cCount
        
        'Repeat for blue
        lColor = srcImageData(QuickVal, y)
        hColor = lColor
        
        lColor = lColor - mThreshold
        If lColor < 0 Then lColor = 0
        hColor = hColor + mThreshold
        If hColor > 255 Then hColor = 255
        
        b = 0: cCount = 0
        For i = lColor To hColor
            b = b + i * bValues(i)
            cCount = cCount + bValues(i)
        Next i
        
        b = b / cCount
        
        'Finally, apply the results to the image.
        dstImageData(QuickVal + 2, y) = r
        dstImageData(QuickVal + 1, y) = g
        dstImageData(QuickVal, y) = b
        
    Next y
        atBottom = Not atBottom
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
        
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    srcDIB.eraseDIB
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic

End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Mean shift", , buildParams(sltRadius.Value, sltThreshold.Value), UNDO_LAYER
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
    If cmdBar.previewsAllowed Then Me.ApplyMeanShiftFilter sltRadius.Value, sltThreshold.Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

