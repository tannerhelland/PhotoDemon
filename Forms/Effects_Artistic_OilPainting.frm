VERSION 5.00
Begin VB.Form FormOilPainting 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Oil painting"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11415
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   761
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11415
      _ExtentX        =   20135
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
      Left            =   6000
      TabIndex        =   2
      Top             =   2040
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1244
      Caption         =   "brush size"
      Min             =   1
      Max             =   200
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdSlider sltPercent 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2880
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1244
      Caption         =   "detail"
      Min             =   1
      Max             =   50
      Value           =   15
      DefaultValue    =   15
   End
End
Attribute VB_Name = "FormOilPainting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Oil Painting Effect Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 09/August/13
'Last updated: 26/July/17
'Last update: performance improvements, migrate to XML params
'
'Oil painting image effect, heavily optimized.  An "accumulation" technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before applying this
' effect at a large radius (> 10).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply an "oil painting" effect to the image (heavily optimized accumulation implementation!)
'Inputs: radius of the effect (min 1, no real max - but the scroll bar is maxed at 200 presently)
'        smoothness of the effect; smaller values indicate less smoothness (e.g. less bins are used to calculate luminance)
Public Sub ApplyOilPaintingEffect(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim mRadius As Long, mLevels As Double, kernelShape As PD_PixelRegionShape
    mRadius = cParams.GetLong("radius", 1&)
    mLevels = cParams.GetDouble("levels", 50#)
    kernelShape = cParams.GetLong("kernelshape", PDPRS_Rectangle)
    
    If (Not toPreview) Then Message "Repainting image with oils..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    workingDIB.WrapArrayAroundDIB dstImageData, dstSA
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent processed pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'Create a second local array.  This will contain the a copy of the current image,
    ' and we will use it as our source reference.
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        mRadius = mRadius * curDIBValues.previewModifier
        If (mRadius < 1) Then mRadius = 1
    End If
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If (finalY - initY) < (finalX - initX) Then
        If mRadius > (finalY - initY) Then mRadius = finalY - initY
    Else
        If mRadius > (finalX - initX) Then mRadius = finalX - initX
    End If
    
    Dim xStride As Long, xStrideInner As Long, quickY As Long
    
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
            
    'Oil painting takes a lot of variables
    Dim rValues(0 To 255) As Long, gValues(0 To 255) As Long, bValues(0 To 255) As Long, lValues(0 To 255) As Long
    Dim lbX As Long, lbY As Long, ubX As Long, ubY As Long
    Dim obuX As Boolean, obuY As Boolean, oblY As Boolean
    Dim i As Long, j As Long
    Dim r As Long, g As Long, b As Long, l As Long
    
    Dim atBottom As Boolean
    atBottom = True
    
    Dim startY As Long, stopY As Long, yStep As Long
    
    'We'll use the requested number of levels to create a gradient output.  Thus divide mLevels by 255, and build a
    ' look-up table for it.
    mLevels = mLevels / 255
    Dim maxPossibleLum As Long
    maxPossibleLum = 255 * mLevels
    
    Dim lLookup() As Byte
    ReDim lLookup(0 To 765) As Byte
    For i = 0 To 765
        lLookup(i) = CLng(CDbl(i / 3) * mLevels)
    Next i
    
    'Later in the function, we must find the most populated luminance bin; these values help us track it
    Dim maxBinCount As Long, maxBinIndex As Long
    
    'Generate an initial array of median data for the first pixel
    For x = initX To initX + mRadius - 1
        xStride = x * 4
    For y = initY To initY + mRadius
    
        r = srcImageData(xStride + 2, y)
        g = srcImageData(xStride + 1, y)
        b = srcImageData(xStride, y)
        l = lLookup(r + g + b)
        rValues(l) = rValues(l) + r
        gValues(l) = gValues(l) + g
        bValues(l) = bValues(l) + b
        lValues(l) = lValues(l) + 1
        
    Next y
    Next x
                
    'Loop through each pixel in the image, tallying median values as we go
    For x = initX To finalX
            
        xStride = x * 4
        
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
        
            xStrideInner = (lbX - 1) * 4
        
            For j = lbY To ubY
                r = srcImageData(xStrideInner + 2, j)
                g = srcImageData(xStrideInner + 1, j)
                b = srcImageData(xStrideInner, j)
                l = lLookup(r + g + b)
                rValues(l) = rValues(l) - r
                gValues(l) = gValues(l) - g
                bValues(l) = bValues(l) - b
                lValues(l) = lValues(l) - 1
            Next j
        
        End If
        
        'Add leading values to the median box if they lie inside the processing radius
        If Not obuX Then
        
            xStrideInner = ubX * 4
            
            For j = lbY To ubY
                r = srcImageData(xStrideInner + 2, j)
                g = srcImageData(xStrideInner + 1, j)
                b = srcImageData(xStrideInner, j)
                l = lLookup(r + g + b)
                rValues(l) = rValues(l) + r
                gValues(l) = gValues(l) + g
                bValues(l) = bValues(l) + b
                lValues(l) = lValues(l) + 1
            Next j
            
        End If
        
        'Depending on the direction we are moving, remove a line of pixels from the median box
        ' (because the interior loop will add it back in).
        If atBottom Then
        
            For i = lbX To ubX
                xStrideInner = i * 4
                r = srcImageData(xStrideInner + 2, mRadius)
                g = srcImageData(xStrideInner + 1, mRadius)
                b = srcImageData(xStrideInner, mRadius)
                l = lLookup(r + g + b)
                rValues(l) = rValues(l) - r
                gValues(l) = gValues(l) - g
                bValues(l) = bValues(l) - b
                lValues(l) = lValues(l) - 1
            Next i
       
        Else
       
            quickY = finalY - mRadius
       
            For i = lbX To ubX
                xStrideInner = i * 4
                r = srcImageData(xStrideInner + 2, quickY)
                g = srcImageData(xStrideInner + 1, quickY)
                b = srcImageData(xStrideInner, quickY)
                l = lLookup(r + g + b)
                rValues(l) = rValues(l) - r
                gValues(l) = gValues(l) - g
                bValues(l) = bValues(l) - b
                lValues(l) = lValues(l) - 1
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
            
                quickY = lbY - 1
            
                For i = lbX To ubX
                    xStrideInner = i * 4
                    r = srcImageData(xStrideInner + 2, quickY)
                    g = srcImageData(xStrideInner + 1, quickY)
                    b = srcImageData(xStrideInner, quickY)
                    l = lLookup(r + g + b)
                    rValues(l) = rValues(l) - r
                    gValues(l) = gValues(l) - g
                    bValues(l) = bValues(l) - b
                    lValues(l) = lValues(l) - 1
                Next i
                        
            End If
                    
            'Add leading values
            If Not obuY Then
            
                For i = lbX To ubX
                    xStrideInner = i * 4
                    r = srcImageData(xStrideInner + 2, ubY)
                    g = srcImageData(xStrideInner + 1, ubY)
                    b = srcImageData(xStrideInner, ubY)
                    l = lLookup(r + g + b)
                    rValues(l) = rValues(l) + r
                    gValues(l) = gValues(l) + g
                    bValues(l) = bValues(l) + b
                    lValues(l) = lValues(l) + 1
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
            
                quickY = ubY + 1
            
                For i = lbX To ubX
                    xStrideInner = i * 4
                    r = srcImageData(xStrideInner + 2, quickY)
                    g = srcImageData(xStrideInner + 1, quickY)
                    b = srcImageData(xStrideInner, quickY)
                    l = lLookup(r + g + b)
                    rValues(l) = rValues(l) - r
                    gValues(l) = gValues(l) - g
                    bValues(l) = bValues(l) - b
                    lValues(l) = lValues(l) - 1
                Next i
                        
            End If
                    
            If Not oblY Then
            
                For i = lbX To ubX
                    xStrideInner = i * 4
                    r = srcImageData(xStrideInner + 2, lbY)
                    g = srcImageData(xStrideInner + 1, lbY)
                    b = srcImageData(xStrideInner, lbY)
                    l = lLookup(r + g + b)
                    rValues(l) = rValues(l) + r
                    gValues(l) = gValues(l) + g
                    bValues(l) = bValues(l) + b
                    lValues(l) = lValues(l) + 1
                Next i
            
            End If
        
        End If
        
        'With a local histogram successfully built for the area surrounding this pixel, we now need to find the
        ' maximum luminance value.
        maxBinCount = 0
        For i = 0 To maxPossibleLum
            If lValues(i) > maxBinCount Then
                maxBinCount = lValues(i)
                maxBinIndex = i
            End If
        Next i
                
        r = rValues(maxBinIndex) / maxBinCount
        If r > 255 Then r = 255
        g = gValues(maxBinIndex) / maxBinCount
        If g > 255 Then g = 255
        b = bValues(maxBinIndex) / maxBinCount
        If b > 255 Then b = 255
                
        'Finally, apply the results to the image.
        dstImageData(xStride + 2, y) = r
        dstImageData(xStride + 1, y) = g
        dstImageData(xStride, y) = b
        
    Next y
        atBottom = Not atBottom
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
        
    'Safely deallocate all image arrays
    srcDIB.UnwrapArrayFromDIB srcImageData
    workingDIB.UnwrapArrayFromDIB dstImageData
    
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic

End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Oil painting", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    'Disable previews while we initialize the dialog
    cmdBar.SetPreviewStatus False
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltPercent_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyOilPaintingEffect GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("radius", sltRadius.Value, "levels", sltPercent.Value)
End Function
