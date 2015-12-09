VERSION 5.00
Begin VB.Form FormRangeFilter 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Range filter"
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
   Begin PhotoDemon.smartCheckBox chkSynchronize 
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   3120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      Caption         =   "synchronize search radius"
   End
   Begin PhotoDemon.buttonStrip btsKernelShape 
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   4080
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Left            =   6000
      Top             =   3720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "kernel shape"
      FontSize        =   12
   End
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
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "horizontal radius"
      Min             =   1
      Max             =   50
      Value           =   5
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   705
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "vertical radius"
      Min             =   1
      Max             =   50
      Value           =   5
   End
End
Attribute VB_Name = "FormRangeFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Range filter edge detection tool
'Copyright 2015-2015 by Tanner Helland
'Created: 23/November/15
'Last updated: 23/November/15
'Last update: initial build
'
'This is a heavily optimized "range filter" function.  An accumulation technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before applying any
' filter of a large radius (> 20).
'
'Range filtering is an edge-detection technique.  It searches some variable range around each pixel, looking for the
' maximum difference between any two pixels in the current search area.  Gain can optionally be applied to boost the
' output of the function.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a "range filter" filter to the current master image (basically a max/min rank algorithm, with some tweaks)
'Input: radius of the search (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub ApplyRangeFilter(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.setParamString parameterList
    
    Dim hRadius As Double, vRadius As Double, kernelShape As PD_PIXEL_REGION_SHAPE
    hRadius = cParams.GetDouble("hRadius", 1#)
    vRadius = cParams.GetDouble("vRadius", hRadius)
    kernelShape = cParams.GetLong("kernelShape", PDPRS_Circle)
    
    If Not toPreview Then Message "Searching each pixel range for edges..."
        
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
    If (hRadius < 1) Then hRadius = 1
    If (vRadius < 1) Then vRadius = 1
    
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
    
    'The number of pixels in the current box are tracked dynamically.
    Dim NumOfPixels As Long
    NumOfPixels = 0
            
    'Accumulation filters like this take a lot of variables
    'We use an optimized histogram technique for calculating means, which means a lot of intermediate values are required
    Dim rValues() As Long, gValues() As Long, bValues() As Long, aValues() As Long
    ReDim rValues(0 To 255) As Long
    ReDim gValues(0 To 255) As Long
    ReDim bValues(0 To 255) As Long
    ReDim aValues(0 To 255) As Long
    
    Dim r As Long, g As Long, b As Long
    Dim lowR As Long, lowG As Long, lowB As Long
    Dim highR As Long, highG As Long, highB As Long
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    Dim directionDown As Boolean
    directionDown = True
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator(srcDIB, xRadius, yRadius, kernelShape) Then
        
        NumOfPixels = cPixelIterator.LockTargetHistograms(rValues, gValues, bValues, aValues, False)
        
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
                
                'With the accumulation box successfully calculated, we can now find max/min values for this region.
                
                'Loop through each color component of the histogram, looking for minimum values
                lowR = 0: lowG = 0: lowB = 0
                
                i = 0
                Do While (rValues(i) = 0)
                    i = i + 1
                    If i > 255 Then Exit Do
                Loop
                lowR = i
                
                i = 0
                Do While (gValues(i) = 0)
                    i = i + 1
                    If i > 255 Then Exit Do
                Loop
                lowG = i
                
                i = 0
                Do While (bValues(i) = 0)
                    i = i + 1
                    If i > 255 Then Exit Do
                Loop
                lowB = i
                
                'Now do the same thing at the top of the histogram
                highR = 255
                highG = 255
                highB = 255
                
                i = 255
                Do While (rValues(i) = 0)
                    i = i - 1
                    If i < 0 Then Exit Do
                Loop
                highR = i
                
                i = 255
                Do While (gValues(i) = 0)
                    i = i - 1
                    If i < 0 Then Exit Do
                Loop
                highG = i
                
                i = 255
                Do While (bValues(i) = 0)
                    i = i - 1
                    If i < 0 Then Exit Do
                Loop
                highB = i
                
                'Failsafe check for empty histograms
                If highB < lowB Then highB = lowB
                If highG < lowG Then highG = lowG
                If highR < lowR Then highR = lowR
                
                'Set each channel to the difference between their max/min values.
                dstImageData(x, y) = highB - lowB
                dstImageData(x + 1, y) = highG - lowG
                dstImageData(x + 2, y) = highR - lowR
                
                'Move the iterator in the correct direction
                If directionDown Then
                    If y < finalY Then NumOfPixels = cPixelIterator.MoveYDown
                Else
                    If y > initY Then NumOfPixels = cPixelIterator.MoveYUp
                End If
                
            Next y
            
            'Reverse y-directionality on each pass
            directionDown = Not directionDown
            If x < finalX Then NumOfPixels = cPixelIterator.MoveXRight
            
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

Private Sub chkSynchronize_Click()
    If CBool(chkSynchronize.Value) Then sltRadius(1).Value = sltRadius(0).Value
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Range filter", , GetLocalParamString(), UNDO_LAYER
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
    Interface.PopKernelShapeButtonStrip btsKernelShape, PDPRS_Circle
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyRangeFilter GetLocalParamString(), True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltRadius_Change(Index As Integer)
    
    If CBool(chkSynchronize.Value) Then
        If sltRadius(Abs(Index - 1)).Value <> sltRadius(Index).Value Then sltRadius(Abs(Index - 1)).Value = sltRadius(Index).Value
    End If
    
    updatePreview
    
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = buildParamList("hRadius", sltRadius(0).Value, "vRadius", sltRadius(1).Value, "kernelShape", btsKernelShape.ListIndex)
End Function
