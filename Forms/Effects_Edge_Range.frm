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
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "horizontal strength"
      Min             =   1
      Max             =   50
      Value           =   5
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   705
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   2880
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "vertical strength"
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
    
    Dim hRadius As Double, vRadius As Double
    hRadius = cParams.GetDouble("hRadius", 1#)
    vRadius = cParams.GetDouble("vRadius", hRadius)
    
    If Not toPreview Then Message "Searching each pixel range for edges..."
        
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
    
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
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
    Dim QuickVal As Long, QuickValInner As Long, QuickY As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'The number of pixels in the current box are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'Accumulation filters like this take a lot of variables
    Dim rValues(0 To 255) As Long, gValues(0 To 255) As Long, bValues(0 To 255) As Long
    Dim lbX As Long, lbY As Long, ubX As Long, ubY As Long
    Dim obuX As Boolean, obuY As Boolean, oblY As Boolean
    Dim i As Long, j As Long
    Dim cutoffTotal As Long
    Dim r As Long, g As Long, b As Long
    Dim lowR As Long, lowG As Long, lowB As Long
    Dim highR As Long, highG As Long, highB As Long
    
    Dim atBottom As Boolean
    atBottom = True
    
    Dim startY As Long, stopY As Long, yStep As Long
    
    numOfPixels = 0
    
    'Generate an initial array of median data for the first pixel
    For x = initX To initX + xRadius - 1
        QuickVal = x * qvDepth
    For y = initY To initY + yRadius
    
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
        rValues(r) = rValues(r) + 1
        gValues(g) = gValues(g) + 1
        bValues(b) = bValues(b) + 1
        
        'Increase the pixel tally
        numOfPixels = numOfPixels + 1
        
    Next y
    Next x
                
    'Loop through each pixel in the image, tallying median values as we go
    For x = initX To finalX
            
        QuickVal = x * qvDepth
        
        'Determine the bounds of the current median box in the X direction
        lbX = x - xRadius
        If lbX < 0 Then lbX = 0
        ubX = x + xRadius
        
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
                numOfPixels = numOfPixels - 1
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
                numOfPixels = numOfPixels + 1
            Next j
            
        End If
        
        'Depending on the direction we are moving, remove a line of pixels from the median box
        ' (because the interior loop will add it back in).
        If atBottom Then
                
            For i = lbX To ubX
                QuickValInner = i * qvDepth
                r = srcImageData(QuickValInner + 2, yRadius)
                g = srcImageData(QuickValInner + 1, yRadius)
                b = srcImageData(QuickValInner, yRadius)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
                numOfPixels = numOfPixels - 1
            Next i
        
        Else
        
            QuickY = finalY - yRadius
        
            For i = lbX To ubX
                QuickValInner = i * qvDepth
                r = srcImageData(QuickValInner + 2, QuickY)
                g = srcImageData(QuickValInner + 1, QuickY)
                b = srcImageData(QuickValInner, QuickY)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
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
            lbY = y - yRadius
            If lbY < 0 Then lbY = 0
            
            ubY = y + yRadius
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
                    numOfPixels = numOfPixels - 1
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
                    numOfPixels = numOfPixels + 1
                Next i
            
            End If
            
        'The exact same code as above, but in the opposite direction
        Else
        
            lbY = y - yRadius
            If lbY < 0 Then
                oblY = True
                lbY = 0
            Else
                oblY = False
            End If
            
            ubY = y + yRadius
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
                    numOfPixels = numOfPixels - 1
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
                    numOfPixels = numOfPixels + 1
                Next i
            
            End If
        
        End If
                
        'With the accumulation box successfully calculated, we can now find max/min values for this region.
        
        'Loop through each color component of the histogram, looking for minimum values
        lowR = 0
        lowG = 0
        lowB = 0
        
        i = 0
        Do While (rValues(i) = 0)
            i = i + 1
        Loop
        lowR = i
        
        i = 0
        Do While (gValues(i) = 0)
            i = i + 1
        Loop
        lowG = i
        
        i = 0
        Do While (bValues(i) = 0)
            i = i + 1
        Loop
        lowB = i
        
        'Now do the same thing at the top of the histogram
        highR = 255
        highG = 255
        highB = 255
        
        i = 255
        Do While (rValues(i) = 0)
            i = i - 1
        Loop
        highR = i
        
        i = 255
        Do While (gValues(i) = 0)
            i = i - 1
        Loop
        highG = i
        
        i = 255
        Do While (bValues(i) = 0)
            i = i - 1
        Loop
        highB = i
        
        'Set each channel to the difference between their max/min values.
        dstImageData(QuickVal, y) = highB - lowB
        dstImageData(QuickVal + 1, y) = highG - lowG
        dstImageData(QuickVal + 2, y) = highR - lowR
        
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
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic

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
    updatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = buildParamList("hRadius", sltRadius(0).Value, "vRadius", sltRadius(1).Value)
End Function
