Attribute VB_Name = "Filters_Miscellaneous"
'***************************************************************************
'Filter Module
'Copyright ©2000-2014 by Tanner Helland
'Created: 13/October/00
'Last updated: 23/July/13
'Last update: added a public function for filling histogram arrays with data.  This should allow me to trim unnecessary
'             code from a number of other places.
'
'The general image filter module; contains unorganized routines at present.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Fill the supplied arrays with histogram data for the current image
' In order, the arrays that need to be supplied are:
' 1) array for histogram data (dimensioned [0,3][0,255] - the first ordinal specifies channel)
' 2) array for logarithmic histogram data (dimensioned same as hData)
' 3) Array for max channel values (dimensioned [0,3])
' 4) Array for max log channel values
' 5) Array of where the maximum channel values occur (histogram index)
Public Sub fillHistogramArrays(ByRef hData() As Double, ByRef hDataLog() As Double, ByRef channelMax() As Double, ByRef channelMaxLog() As Double, ByRef channelMaxPosition() As Byte)
    
    'Redimension the various arrays
    ReDim hData(0 To 3, 0 To 255) As Double
    ReDim hDataLog(0 To 3, 0 To 255) As Double
    ReDim channelMax(0 To 3) As Double
    ReDim channelMaxLog(0 To 3) As Double
    ReDim channelMaxPosition(0 To 3) As Byte
    
    'Create a local array and point it at the pixel data we want to scan
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, , , , True
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'These variables will hold temporary histogram values
    Dim r As Long, g As Long, b As Long, l As Long
    
    'If the histogram has already been used, we need to clear out all the
    'maximum values and histogram values
    Dim hMax As Double, hMaxLog As Double
    hMax = 0:    hMaxLog = 0
    
    For x = 0 To 3
        channelMax(x) = 0
        channelMaxLog(x) = 0
        For y = 0 To 255
            hData(x, y) = 0
        Next y
    Next x
    
    'Build a look-up table for luminance conversion; 765 = 255 * 3
    Dim lumLookup(0 To 765) As Byte
    
    For x = 0 To 765
        lumLookup(x) = x \ 3
    Next x
    
    'Run a quick loop through the image, gathering what we need to calculate our histogram
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'We have to gather the red, green, and blue in order to calculate luminance
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Rather than generate authentic luminance (which requires a costly HSL conversion routine), we use a simpler average value.
        l = lumLookup(r + g + b)
        
        'Increment each value in the array, depending on its present value; this will let us see how many pixels of
        ' each color value (and luminance value) there are in the image
        
        'Red
        hData(0, r) = hData(0, r) + 1
        'Green
        hData(1, g) = hData(1, g) + 1
        'Blue
        hData(2, b) = hData(2, b) + 1
        'Luminance
        hData(3, l) = hData(3, l) + 1
        
    Next y
    Next x
    
    'With our dataset successfully collected, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Run a quick loop through the completed array to find maximum values
    For x = 0 To 3
        For y = 0 To 255
            If hData(x, y) > channelMax(x) Then
                channelMax(x) = hData(x, y)
                channelMaxPosition(x) = y
            End If
        Next y
    Next x
    
    'Now calculate the logarithmic version of the histogram
    For x = 0 To 3
        If channelMax(x) <> 0 Then channelMaxLog(x) = Log(channelMax(x)) Else channelMaxLog(x) = 0
    Next x
    
    For x = 0 To 3
        For y = 0 To 255
            If hData(x, y) <> 0 Then
                hDataLog(x, y) = Log(hData(x, y))
            Else
                hDataLog(x, y) = 0
            End If
        Next y
    Next x
    
End Sub

'Convert the image's color depth to a new value.  (Currently, only 24bpp and 32bpp is allowed.)
Public Sub ConvertImageColorDepth(ByVal newColorDepth As Long, Optional ByVal newBackColor As Long = vbWhite)

    Message "Converting image mode..."
    
    'TODO: handle this with layers.  I'm not sure how best to do this, just yet - I may move all transparency
    '      concepts to the Layer menu.
    
    If newColorDepth = 24 Then
    
        'Ask the current DIB to convert itself to 24bpp mode
        pdImages(g_CurrentImage).getActiveDIB.convertTo24bpp newBackColor
        
        'Because PD now uses an "always 32bpp" approach to layers, convert the image back to 32bpp mode.
        ' (All its alpha values will be 255.)
        pdImages(g_CurrentImage).getActiveDIB.convertTo32bpp 255
        
    Else
    
        'Ask the current DIB to convert itself to 32bpp mode
        pdImages(g_CurrentImage).getActiveDIB.convertTo32bpp 255
    
    End If
    
    Message "Finished."
    
    'Redraw the main window
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Load the last Undo file and alpha-blend it with the current image
Public Sub MenuFadeLastEffect()

    Message "Fading last effect..."
    
    'Create a temporary DIB and use it to load the last Undo file's pixel data
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createFromFile pdImages(g_CurrentImage).undoManager.getLastUndoFile()
    
    'Create a local array and point it at the pixel data of that undo file
    Dim uImageData() As Byte
    Dim uSA As SAFEARRAY2D
    prepSafeArray uSA, tmpDIB
    CopyMemory ByVal VarPtrArray(uImageData()), VarPtr(uSA), 4
        
    'Create another array, but point it at the pixel data of the current image
    Dim cImageData() As Byte
    Dim cSA As SAFEARRAY2D
    prepSafeArray cSA, pdImages(g_CurrentImage).getActiveDIB()
    CopyMemory ByVal VarPtrArray(cImageData()), VarPtr(cSA), 4
    
    'Because the undo file and current image may be different sizes (if the last action was a resize, for example), we need to
    ' find the minimum width and height to make sure there are no out-of-bound errors.
    Dim minWidth As Long, minHeight As Long
    If tmpDIB.getDIBWidth < pdImages(g_CurrentImage).Width Then minWidth = tmpDIB.getDIBWidth Else minWidth = pdImages(g_CurrentImage).Width
    If tmpDIB.getDIBHeight < pdImages(g_CurrentImage).Height Then minHeight = tmpDIB.getDIBHeight Else minHeight = pdImages(g_CurrentImage).Height
        
    'Set the progress bar maximum value to that minimum width value
    SetProgBarMax minWidth
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValUndo As Long, qvDepth As Long, qvDepthUndo As Long
    qvDepth = pdImages(g_CurrentImage).getActiveDIB().getDIBColorDepth \ 8
    qvDepthUndo = tmpDIB.getDIBColorDepth \ 8
        
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Local loop variables can be more efficiently cached by VB's compiler
    Dim x As Long, y As Long
    
    'Finally, prepare a look-up table for the alpha-blend
    Dim aLookUp(0 To 255, 0 To 255) As Byte
    Dim tmpCalc As Long
    
    For x = 0 To 255
    For y = 0 To 255
        tmpCalc = (x + y) \ 2
        aLookUp(x, y) = CByte(tmpCalc)
    Next y
    Next x
        
    'Loop through both images, alpha-blending pixels as we go
    For x = 0 To minWidth - 1
        QuickVal = x * qvDepth
        QuickValUndo = x * qvDepthUndo
    For y = 0 To minHeight - 1
        cImageData(QuickVal, y) = aLookUp(cImageData(QuickVal, y), uImageData(QuickValUndo, y))
        cImageData(QuickVal + 1, y) = aLookUp(cImageData(QuickVal + 1, y), uImageData(QuickValUndo + 1, y))
        cImageData(QuickVal + 2, y) = aLookUp(cImageData(QuickVal + 2, y), uImageData(QuickValUndo + 2, y))
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
        
    'With our work complete, point both ImageData() arrays away from their respective DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(uImageData), 0&, 4
    Erase uImageData
    
    CopyMemory ByVal VarPtrArray(cImageData), 0&, 4
    Erase cImageData
        
    'Erase our temporary DIB as well
    tmpDIB.eraseDIB
    Set tmpDIB = Nothing
    
    'Render the final image to the screen
    SetProgBarVal 0
    releaseProgressBar
    Message "Finished."
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Render an image using faux thermography; basically, map luminance values as if they were heat, and use a modified hue spectrum for representation.
' (I have manually tweaked the values at certain ranges to better approximate actual thermography.)
Public Sub MenuHeatMap()

    Message "Performing thermographic analysis..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
    Dim hVal As Double, sVal As Double, lVal As Double
    Dim h As Double, s As Double, l As Double
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        grayVal = gLookup(r + g + b)
        
        'Based on the luminance of this pixel, apply a predetermined hue gradient (stretching between -1 and 5)
        hVal = (CSng(grayVal) / 255) * 360
        
        'If the hue is "below" blue, gradually darken the corresponding luminance value
        If hVal < 120 Then
            lVal = (0.35 * (hVal / 120)) + 0.15
        Else
            lVal = 0.5
        End If
        
        'Invert the hue
        hVal = 360 - hVal
                
        'Place hue in the range of -1 to 5, per the requirements of our HSL conversion algorithm
        hVal = (hVal - 60) / 60
        
        'Use nearly full saturation (for dramatic effect)
        sVal = 0.8
        
        'Use RGB to calculate hue, saturation, and luminance
        tRGBToHSL r, g, b, h, s, l
        
        'Now convert those HSL values back to RGB, but substitute in our artificial hue value (calculated above)
        tHSLToRGB hVal, sVal, lVal, r, g, b
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'A very neat comic-book filter that actually blends together a number of other filters into one!
Public Sub MenuComicBook()
    
    Dim gRadius As Long
    gRadius = 20
    
    Dim gThreshold As Long
    gThreshold = 8
    
    Message "Animating image (stage %1 of %2)...", 1, 3
                
    'More color variables - in this case, sums for each color component
    Dim r As Long, g As Long, b As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    Dim tDelta As Long
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    Dim gaussDIB As pdDIB
    Set gaussDIB = New pdDIB
    gaussDIB.createFromExistingDIB workingDIB
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    CreateGaussianBlurDIB gRadius, srcDIB, gaussDIB, False, finalY + finalY + finalX + finalX
    
    If cancelCurrentAction Then
        srcDIB.eraseDIB
        gaussDIB.eraseDIB
        finalizeImageData
        Exit Sub
    End If
        
    'Now that we have a gaussian DIB created in gaussDIB, we can point arrays toward it and the source DIB
    Dim dstImageData() As Byte
    prepImageData dstSA
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    Dim GaussImageData() As Byte
    Dim gaussSA As SAFEARRAY2D
    prepSafeArray gaussSA, gaussDIB
    CopyMemory ByVal VarPtrArray(GaussImageData()), VarPtr(gaussSA), 4
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
        
    Message "Animating image (stage %1 of %2)...", 2, 3
        
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
        If tDelta > gThreshold Then
            If tDelta <> 0 Then blendVal = 1 - (gThreshold / tDelta) Else blendVal = 0
            dstImageData(QuickVal + 2, y) = BlendColors(srcImageData(QuickVal + 2, y), GaussImageData(QuickVal + 2, y), blendVal)
            dstImageData(QuickVal + 1, y) = BlendColors(srcImageData(QuickVal + 1, y), GaussImageData(QuickVal + 1, y), blendVal)
            dstImageData(QuickVal, y) = BlendColors(srcImageData(QuickVal, y), GaussImageData(QuickVal, y), blendVal)
            If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = BlendColors(srcImageData(QuickVal + 3, y), GaussImageData(QuickVal + 3, y), blendVal)
        End If
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x + (finalY * 2)
        End If
    Next x
        
    'With our work complete, release all arrays
    CopyMemory ByVal VarPtrArray(GaussImageData), 0&, 4
    Erase GaussImageData
    
    gaussDIB.eraseDIB
    Set gaussDIB = Nothing
    
    'Because this function occurs in multiple passes, it requires specialized cancel behavior.  All array references must be dropped
    ' or the program will experience a hard-freeze.
    If cancelCurrentAction Then
        CopyMemory ByVal VarPtrArray(dstImageData()), 0&, 4
        CopyMemory ByVal VarPtrArray(srcImageData()), 0&, 4
        finalizeImageData
        Exit Sub
    End If
    
    'The last thing we need to do is sketch in the edges of the image.
    
    Message "Animating image (stage %1 of %2)...", 3, 3
    
    'We can't do this at the borders of the image, so shrink the functional area by one in each dimension.
    initX = initX + 1
    initY = initY + 1
    finalX = finalX - 1
    finalY = finalY - 1
    
    Dim QuickValRight As Long, QuickValLeft As Long, tmpColor As Long, tMin As Long
    Dim z As Long
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
        QuickValRight = (x + 1) * qvDepth
        QuickValLeft = (x - 1) * qvDepth
    For y = initY To finalY
        For z = 0 To 2
    
            tMin = 255
            tmpColor = srcImageData(QuickValRight + z, y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValRight + z, y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValRight + z, y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            
            If tMin > 255 Then tMin = 255
            If tMin < 0 Then tMin = 0
            
            Select Case z
            
                Case 0
                    b = 255 - (srcImageData(QuickVal, y) - tMin)
            
                Case 1
                    g = 255 - (srcImageData(QuickVal + 1, y) - tMin)
                    
                Case 2
                    r = 255 - (srcImageData(QuickVal + 2, y) - tMin)
            
            End Select
                    
        Next z
        
        r2 = dstImageData(QuickVal + 2, y)
        g2 = dstImageData(QuickVal + 1, y)
        b2 = dstImageData(QuickVal, y)
        
        r = ((CSng(r) / 255) * (CSng(r2) / 255)) * 255
        g = ((CSng(g) / 255) * (CSng(g2) / 255)) * 255
        b = ((CSng(b) / 255) * (CSng(b2) / 255)) * 255
        
        dstImageData(QuickVal + 2, y) = r
        dstImageData(QuickVal + 1, y) = g
        dstImageData(QuickVal, y) = b
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x + finalX + (finalY * 2)
        End If
    Next x
    
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'Wacky filter discovered by trial-and-error.  I named it "synthesize".
Public Sub MenuSynthesize()

    Message "Synthesizing new image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        grayVal = gLookup(r + g + b)
        
        r = g + b - grayVal
        g = r + b - grayVal
        b = r + g - grayVal
        
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'Another random filter discovered by trial-and-error.  "Alien" effect.
Public Sub MenuAlien()

    Message "Abducting image and probing it for weaknesses..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        newR = b + g - r
        newG = r + b - g
        newB = r + g - b
        
        If newR > 255 Then newR = 255
        If newR < 0 Then newR = 0
        If newG > 255 Then newG = 255
        If newG < 0 Then newG = 0
        If newB > 255 Then newB = 255
        If newB < 0 Then newB = 0
        
        ImageData(QuickVal + 2, y) = newR
        ImageData(QuickVal + 1, y) = newG
        ImageData(QuickVal, y) = newB
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
  
End Sub

'Very improved version of "sepia".  This is more involved than a typical "change to brown" effect - the white balance and
' shading is also adjusted to give the image a more "antique" look.
Public Sub MenuAntique()
    
    Message "Accelerating to 88mph in order to antique-ify this image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'We're going to need grayscale values as part of the effect; grayscale is easily optimized via a look-up table
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
    
    'We're going to use gamma conversion as part of the effect; gamma is easily optimized via a look-up table
    Dim gammaLookUp(0 To 255) As Byte
    Dim tmpVal As Double
    For x = 0 To 255
        tmpVal = x / 255
        tmpVal = tmpVal ^ (1 / 1.6)
        tmpVal = tmpVal * 255
        If tmpVal > 255 Then tmpVal = 255
        If tmpVal < 0 Then tmpVal = 0
        gammaLookUp(x) = CByte(tmpVal)
    Next x
    
    'Finally, we also need to adjust brightness.  A look-up table is once again invaluable
    Dim bLookup(0 To 255) As Byte
    For x = 0 To 255
        tmpVal = x * 1.75
        If tmpVal > 255 Then tmpVal = 255
        bLookup(x) = CByte(tmpVal)
    Next x
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim gray As Long
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        gray = gLookup(r + g + b)
        
        r = (r + gray) \ 2
        g = (g + gray) \ 2
        b = (b + gray) \ 2
        
        r = (g * b) \ 256
        g = (b * r) \ 256
        b = (r * g) \ 256
        
        newR = bLookup(r)
        newG = bLookup(g)
        newB = bLookup(b)
        
        newR = gammaLookUp(newR)
        newG = gammaLookUp(newG)
        newB = gammaLookUp(newB)
        
        ImageData(QuickVal + 2, y) = newR
        ImageData(QuickVal + 1, y) = newG
        ImageData(QuickVal, y) = newB
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'Dull but standard "sepia" transformation.  Values derived from the w3c standard at:
' https://dvcs.w3.org/hg/FXTF/raw-file/tip/filters/index.html#sepiaEquivalent
Public Sub MenuSepia()
    
    Message "Engaging hipsters to perform sepia conversion..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Double, newG As Double, newB As Double
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
                
        newR = CSng(r) * 0.393 + CSng(g) * 0.769 + CSng(b) * 0.189
        newG = CSng(r) * 0.349 + CSng(g) * 0.686 + CSng(b) * 0.168
        newB = CSng(r) * 0.272 + CSng(g) * 0.534 + CSng(b) * 0.131
        
        r = newR
        g = newG
        b = newB
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'Makes the picture appear like it has been shaken
Public Sub MenuVibrate()

    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("vibrate") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & "0|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "1|0|0|0|-1|"
    tmpString = tmpString & "0|-1|0|1|0|"
    tmpString = tmpString & "0|0|1|0|0|"
    tmpString = tmpString & "0|1|0|-1|0|"
    tmpString = tmpString & "-1|0|0|0|1"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString

End Sub

'Another filter found by trial-and-error.  "Dream" effect.
Public Sub MenuDream()

    Message "Putting image to sleep, then measuring its REM cycles..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim grayVal As Long
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        newR = ImageData(QuickVal + 2, y)
        newG = ImageData(QuickVal + 1, y)
        newB = ImageData(QuickVal, y)
        
        grayVal = gLookup(newR + newG + newB)
        
        r = Abs(newR - grayVal) + Abs(newR - newG) + Abs(newR - newB) + (newR \ 2)
        g = Abs(newG - grayVal) + Abs(newG - newB) + Abs(newG - newR) + (newG \ 2)
        b = Abs(newB - grayVal) + Abs(newB - newR) + Abs(newB - newG) + (newB \ 2)
        
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'A bright-green filter I've aptly named "radioactive".
Public Sub MenuRadioactive()

    Message "Injecting image with non-ionizing radiation..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        If r = 0 Then r = 1
        If g = 0 Then g = 1
        If b = 0 Then b = 1
        
        newR = (g * b) \ r
        newG = (r * b) \ g
        newB = (r * g) \ b
        
        If newR > 255 Then newR = 255
        If newG > 255 Then newG = 255
        If newB > 255 Then newB = 255
        
        newG = 255 - newG
        
        ImageData(QuickVal + 2, y) = newR
        ImageData(QuickVal + 1, y) = newG
        ImageData(QuickVal, y) = newB
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'Stretch out the contrast and convert the image to dramatic black and white.  Originally called the "comic book" filter, since renamed to Film Noir.
Public Sub MenuFilmNoir()

    Message "Embuing image with the essence of F. Miller..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
    
    'Same goes for contrast
    Dim cLookup(0 To 255) As Byte, cCalc As Long
                
    For x = 0 To 255
        cCalc = x + (((x - 127) * 30) \ 100)
        If cCalc > 255 Then cCalc = 255
        If cCalc < 0 Then cCalc = 0
        cLookup(x) = CByte(cCalc)
    Next x
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        grayVal = gLookup(r + g + b)
        grayVal = cLookup(grayVal)
        
        ImageData(QuickVal + 2, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal, y) = grayVal
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'Subroutine for counting the number of unique colors in an image
Public Sub MenuCountColors()
    
    Message "Counting the number of unique colors in this image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'This array will track whether or not a given color has been detected in the image
    Dim UniqueColors() As Boolean
    ReDim UniqueColors(0 To 16777216) As Boolean
    
    'Total number of unique colors counted so far
    Dim totalCount As Long
    totalCount = 0
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim chkValue As Long
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        chkValue = RGB(r, g, b)
        If UniqueColors(chkValue) = False Then
            totalCount = totalCount + 1
            UniqueColors(chkValue) = True
        End If
        
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Also, erase the counting array
    Erase UniqueColors
    
    'Reset the progress bar
    SetProgBarVal 0
    releaseProgressBar
    
    'Show the user our final tally
    Message "Total number of unique colors: %1", totalCount
    pdMsgBox "This image contains %1 unique colors.", vbOKOnly + vbApplicationModal + vbInformation, "Count Image Colors", totalCount
    
End Sub

'You can use this section of code to test out your own filters.  I've left some sample code below.
Public Sub MenuTest()
    
    pdMsgBox "This menu item only appears in the Visual Basic IDE." & vbCrLf & vbCrLf & "You can use the MenuTest() sub in the Filters_Miscellaneous module to test your own filters.  I typically do this first, then once the filter is working properly, I give it a subroutine of its own.", vbInformation + vbOKOnly + vbApplicationModal, " PhotoDemon Pro Tip"
    
    Exit Sub
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long, grayVal As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim hVal As Double, sVal As Double, lVal As Double
    Dim h As Double, s As Double, l As Double
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        grayVal = gLookup(r + g + b)
        
        'Put interesting color transformations here.  As an example, here's one possible sepia formula.
        newR = grayVal + 40
        newG = grayVal + 20
        newB = grayVal - 30
                                
        If newR < 0 Then newR = 0
        If newG < 0 Then newG = 0
        If newB < 0 Then newB = 0
        
        If newR > 255 Then newR = 255
        If newG > 255 Then newG = 255
        If newB > 255 Then newB = 255
                
        ImageData(QuickVal + 2, y) = newR
        ImageData(QuickVal + 1, y) = newG
        ImageData(QuickVal, y) = newB
                
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

    
End Sub
