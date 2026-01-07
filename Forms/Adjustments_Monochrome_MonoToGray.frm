VERSION 5.00
Begin VB.Form FormMonoToColor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Monochrome to gray"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   Begin PhotoDemon.pdCommandBar cmdBar 
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
      Left            =   6000
      TabIndex        =   2
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   1
      Max             =   8
      Value           =   3
      DefaultValue    =   3
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Left            =   6000
      Top             =   3240
      Width           =   5850
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "advice from the experts"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   1770
      Left            =   6120
      Top             =   3840
      Width           =   5535
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Explanation"
      FontSize        =   9
      ForeColor       =   4210752
      Layout          =   1
   End
End
Attribute VB_Name = "FormMonoToColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Monochrome to Color (technically grayscale) Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 13/Feb/13
'Last updated: 23/August/13
'Last update: added command bar
'
'This is a heavily optimized monochrome-to-grayscale function.  An "accumulation" technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before using it.
'
'This technique is one of my own creation, though I doubt I'm the first one to think of it.  Basically, search a box of some size,
' and within that box, count the number of black (<128) and white (>=128) pixels.  Average the total found to arrive at a grayscale
' value, and assign that to the pixel.
'
'The blue channel alone is used for calculations (for speed purposes - three times faster than checking all three channels!) so results will
' be wonky when used on a color image.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Given a monochrome image, convert it to grayscale
'Input: radius of the search area (min 1, no real max - but there are diminishing returns above 50)
Public Sub ConvertMonoToColor(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Converting monochrome image to grayscale..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim mRadius As Long
    mRadius = cParams.GetLong("radius", sltRadius.Value)
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    workingDIB.WrapArrayAroundDIB dstImageData, dstSA
    
    'Create a second local array.  This will contain the a copy of the current image,
    ' and we will use it as our source reference.
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    srcDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
        
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        mRadius = mRadius * curDIBValues.previewModifier
        If (mRadius < 1) Then mRadius = 1
    End If
    
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
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'The number of pixels in the current median box are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'Median filtering takes a lot of variables
    Dim highValues As Long
    Dim lbX As Long, lbY As Long, ubX As Long, ubY As Long
    Dim obuX As Boolean, obuY As Boolean, oblY As Boolean
    Dim i As Long, j As Long
    Dim fGray As Long
    
    Dim atBottom As Boolean
    atBottom = True
    
    Dim startY As Long, stopY As Long, yStep As Long
    
    numOfPixels = 0
    
    'Generate an initial array of median data for the first pixel
    For x = initX To initX + mRadius - 1
        xStride = x * 4
    For y = initY To initY + mRadius
    
        If (srcImageData(xStride, y) > 127) Then highValues = highValues + 1
        
        'Increase the pixel tally
        numOfPixels = numOfPixels + 1
        
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
        
        'Remove trailing values from the box if they lie outside the processing radius
        If lbX > 0 Then
        
            xStrideInner = (lbX - 1) * 4
        
            For j = lbY To ubY
                If srcImageData(xStrideInner, j) > 127 Then highValues = highValues - 1
                numOfPixels = numOfPixels - 1
            Next j
        
        End If
        
        'Add leading values to the box if they lie inside the processing radius
        If Not obuX Then
        
            xStrideInner = ubX * 4
            
            For j = lbY To ubY
                If srcImageData(xStrideInner, j) > 127 Then highValues = highValues + 1
                numOfPixels = numOfPixels + 1
            Next j
            
        End If
        
        'Depending on the direction we are moving, remove a line of pixels from the median box
        ' (because the interior loop will add it back in).
        If atBottom Then
                
            For i = lbX To ubX
                xStrideInner = i * 4
                If srcImageData(xStrideInner, mRadius) > 127 Then highValues = highValues - 1
                numOfPixels = numOfPixels - 1
            Next i
        
        Else
        
            quickY = finalY - mRadius
        
            For i = lbX To ubX
                xStrideInner = i * 4
                If srcImageData(xStrideInner, quickY) > 127 Then highValues = highValues - 1
                numOfPixels = numOfPixels - 1
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
                    If srcImageData(xStrideInner, quickY) > 127 Then highValues = highValues - 1
                    numOfPixels = numOfPixels - 1
                Next i
                        
            End If
                    
            'Add leading values
            If Not obuY Then
            
                For i = lbX To ubX
                    xStrideInner = i * 4
                    If srcImageData(xStrideInner, ubY) > 127 Then highValues = highValues + 1
                    numOfPixels = numOfPixels + 1
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
                    If srcImageData(xStrideInner, quickY) > 127 Then highValues = highValues - 1
                    numOfPixels = numOfPixels - 1
                Next i
                        
            End If
                    
            If Not oblY Then
            
                For i = lbX To ubX
                    xStrideInner = i * 4
                    If srcImageData(xStrideInner, lbY) > 127 Then highValues = highValues + 1
                    numOfPixels = numOfPixels + 1
                Next i
            
            End If
        
        End If
                
        'With the box successfully calculated, we can now estimate a grayscale value for this pixel
        fGray = (highValues / numOfPixels) * 255
        
        'Finally, apply the results to the image.
        dstImageData(xStride + 2, y) = fGray
        dstImageData(xStride + 1, y) = fGray
        dstImageData(xStride, y) = fGray
        
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
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic

End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Monochrome to gray", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    'Disable previews while we initialize the dialog
    cmdBar.SetPreviewStatus False
    
    'Provide a small explanation about how this process works
    lblExplanation.Caption = g_Language.TranslateMessage("Like all monochrome-to-grayscale tools, this tool will produce a blurry image.  You can use the Effects -> Sharpen -> Unsharp Masking tool to fix this.  (For best results, use an Unsharp Mask radius at least as large as this radius.)")
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ConvertMonoToColor GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "radius", sltRadius.Value
    GetLocalParamString = cParams.GetParamString()
    
End Function
