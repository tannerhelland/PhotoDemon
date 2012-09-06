Attribute VB_Name = "Filters_Miscellaneous"
'***************************************************************************
'Filter Module
'Copyright ©2000-2012 by Tanner Helland
'Created: 13/October/00
'Last updated: 05/September/12
'Last update: rewrote all code against the new pdLayer class.
'
'The general image filter module; contains unorganized routines at present.
'
'***************************************************************************

Option Explicit

'Load the last Undo file and alpha-blend it with the current image
Public Sub MenuFadeLastEffect()

    Message "Fading last effect..."
    
    'Create a temporary layer and use it to load the last Undo file's pixel data
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromFile GetLastUndoFile()
    
    'Create a local array and point it at the pixel data of that undo file
    Dim uImageData() As Byte
    Dim uSA As SAFEARRAY2D
    prepSafeArray uSA, tmpLayer
    CopyMemory ByVal VarPtrArray(uImageData()), VarPtr(uSA), 4
        
    'Create another array, but point it at the pixel data of the current image
    Dim cImageData() As Byte
    Dim cSA As SAFEARRAY2D
    prepImageData cSA
    CopyMemory ByVal VarPtrArray(cImageData()), VarPtr(cSA), 4
    
    'Because the undo file and current image may be different sizes (if the last action was a resize, for example), we need to
    ' find the minimum width and height to make sure there are no out-of-bound errors.
    Dim minWidth As Long, minHeight As Long
    If tmpLayer.getLayerWidth < pdImages(CurrentImage).Width Then minWidth = tmpLayer.getLayerWidth Else minWidth = pdImages(CurrentImage).Width
    If tmpLayer.getLayerHeight < pdImages(CurrentImage).Height Then minHeight = tmpLayer.getLayerHeight Else minHeight = pdImages(CurrentImage).Height
        
    'Set the progress bar maximum value to that minimum width value
    SetProgBarMax minWidth
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValUndo As Long, qvDepth As Long, qvDepthUndo As Long
    qvDepth = pdImages(CurrentImage).mainLayer.getLayerColorDepth \ 8
    qvDepthUndo = tmpLayer.getLayerColorDepth \ 8
        
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
    CopyMemory ByVal VarPtrArray(cImageData), 0&, 4
    Erase cImageData
    
    CopyMemory ByVal VarPtrArray(uImageData), 0&, 4
    Erase uImageData
    
    'Erase our temporary layer as well
    tmpLayer.eraseLayer
    Set tmpLayer = Nothing
    
    'Finally, pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'Right now this is a work in progress; it's somewhat based off... <description forthcoming>
Public Sub MenuAnimate()

    MsgBox "Sorry, but this filter is still under heavy development.  It's disabled right now due to some instability in the code.  Stay tuned for updates!", vbInformation + vbOKOnly + vbApplicationModal, "Animate filter disabled... for now"
    
    Message "Animate filter canceled"
    
    Exit Sub

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
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
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
        If (x And progBarCheck) = 0 Then SetProgBarVal x
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
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
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
        If (x And progBarCheck) = 0 Then SetProgBarVal x
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
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
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
    Dim tmpVal As Single
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
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'Makes the picture appear like it has been shaken
Public Sub MenuVibrate()
    FilterSize = 5
    ReDim FM(-2 To 2, -2 To 2) As Long
    FM(-2, -2) = 1
    FM(-1, -1) = -1
    FM(0, 0) = 1
    FM(1, 1) = -1
    FM(2, 2) = 1
    FM(-1, 1) = 1
    FM(-2, 2) = -1
    FM(1, -1) = 1
    FM(2, -2) = -1
    FilterWeight = 1
    FilterBias = 0
    DoFilter "Vibrate"
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
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
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
        If (x And progBarCheck) = 0 Then SetProgBarVal x
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
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
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
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'Stretch out the contrast and convert the image to dramatic black and white.  "Comic book" filter.
Public Sub MenuComicBook()

    Message "Embuing image with the essence of F. Miller..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
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
        
        r = Abs(r * (g - b + g + r)) / 255
        g = Abs(r * (b - g + b + r)) / 255
        b = Abs(g * (b - g + b + r)) / 255
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        grayVal = gLookup(r + g + b)
        
        ImageData(QuickVal + 2, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal, y) = grayVal
        
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
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
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
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
    
    'Show the user our finaly tally
    Message "Total number of unique colors: " & totalCount
    MsgBox "This image contains " & totalCount & " unique colors.", vbOKOnly + vbApplicationModal + vbInformation, "Count Image Colors"
    
End Sub

'You can use this section of code to test out your own filters.  I've left some sample code below.
Public Sub MenuTest()
    
    MsgBox "This menu item only appears in the Visual Basic IDE." & vbCrLf & vbCrLf & "You can use the MenuTest() sub in the Filters_Miscellaneous module to test out your own filters.  I typically do this first, then once the filter is working properly, I give it a subroutine of its own.", vbInformation + vbOKOnly + vbApplicationModal, PROGRAMNAME & " Pro Tip"
    
    'Currently testing plugin downloading
    'zLibEnabled = False
    'ScanEnabled = False
    'FreeImageEnabled = False
    'FormPluginDownloader.Show 1, FormMain
    
    Exit Sub
    
    'Supply an arbitrary process value to the Undo file generator
    CreateUndoFile 9999
    
    GetImageData
    Dim r As Long, g As Long, b As Long
    Dim TC As Long, tR As Long, tB As Long, tG As Long
    Dim bR As Byte, bG As Byte, bB As Byte, bC As Byte
    Dim HH As Single, SS As Single, LL As Single
    Dim tH As Single, tS As Single, tL As Single
    Dim xCalc As Long, yCalc As Long
    Dim totalval As Long
    totalval = 0
    
    SetProgBarMax PicWidthL
    
    For x = 0 To PicWidthL
    For y = 0 To PicHeightL
        tR = ImageData(x * 3 + 2, y)
        tG = ImageData(x * 3 + 1, y)
        tB = ImageData(x * 3, y)
        TC = (tR + tG + tB) \ 3
        
        'If TC = 0 Then TC = 1
        'bR = TR
        'bG = TG
        'bB = TB
        'bC = TC
        
        'bR = bR Xor bC
        'bG = bG Xor bC
        'bB = bB Xor bC
        tR = TC * 1.2
        tG = TC
        tB = TC * 0.8
        
        ByteMeL tR
        ByteMeL tG
        ByteMeL tB
        
        ImageData(x * 3 + 2, y) = tR
        ImageData(x * 3 + 1, y) = tG
        ImageData(x * 3, y) = tB
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetProgBarVal 0
    
    setImageData
    
    
    
    
    

    
    
    Exit Sub
    
    
    
    
    
    
    
    
    'xCalc = 10
    'yCalc = 5
    'Randomize
    Message "Running test filter..."
    SetProgBarMax PicWidthL
    xCalc = PicWidthL + PicHeightL
    For x = 0 To PicWidthL
    For y = 0 To PicHeightL
        tR = ImageData(x * 3 + 2, y)
        tG = ImageData(x * 3 + 1, y)
        tB = ImageData(x * 3, y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        'Get the hue and saturation
        'tRGBToHSL TR, TG, TB, HH, SS, LL

        'Working on night-vision
        r = TC \ 2
        b = r
        g = tG + ((tR + tB) \ 4)
        
        'Convert back to RGB using our artificial luminance values
        'tHSLToRGB HH, SS, LL, r, g, b
        
        ByteMeL r
        ByteMeL g
        ByteMeL b
        ImageData(x * 3 + 2, y) = r
        ImageData(x * 3 + 1, y) = g
        ImageData(x * 3, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    setImageData
End Sub

'Here's all the filters I haven't yet found a home for
Public Sub TempHolderForUnplacedFilters()
        Dim r As Long, g As Long, b As Long
        Dim TC As Long, tR As Long, tG As Long, tB As Long
        Dim xCalc As Single, yCalc As Single
        
        'Purpleize
        r = tR
        g = tG + Abs(r - TC)
        b = tB + Abs(g - TC)
        r = tR + Abs(b - TC)
        
        'Trigonometric noise
        xCalc = 6
        yCalc = xCalc \ 2
        'horizontal
        r = tR * Cos((x Mod xCalc - yCalc) / yCalc)
        g = tG * Cos((x Mod xCalc - yCalc) / yCalc)
        b = tB * Cos((x Mod xCalc - yCalc) / yCalc)
        'vertical
        r = tR * Cos(((x + y) Mod xCalc - yCalc) / yCalc)
        g = tG * Cos(((x + y) Mod xCalc - yCalc) / yCalc)
        b = tB * Cos(((x + y) Mod xCalc - yCalc) / yCalc)
        
        'Two-tiered invert
        If tR < 128 Then r = Abs(tR - 128) Else r = Abs(tR - 255)
        If tG < 128 Then g = Abs(tG - 128) Else g = Abs(tG - 255)
        If tB < 128 Then b = Abs(tB - 128) Else b = Abs(tB - 255)
        
        'Difference between colors (can cycle between colors)
        r = Abs(tR - tB)
        g = Abs(tG - tR)
        b = Abs(tB - tG)
        
        'Psycho versions...
        If TC = 0 Then TC = 1
        r = tR Mod TC
        g = tG Mod TC
        b = tB Mod TC
        'and all the bitwise operators
        
        'AutoEnhance clone...?
        r = (tR - TC) + tR
        g = (tG - TC) + tG
        b = (tB - TC) + tB
        'another AutoEnhance clone?
        If TC = 0 Then TC = 1
        r = tR / TC * tR
        g = tG / TC * TC
        b = tB / TC * tB
        'Yet another AutoEnhance version...
        r = Abs(tR - TC) - TC + 2 * tR
        g = Abs(tG - TC) - TC + 2 * tG
        b = Abs(tB - TC) - TC + 2 * tB
        
        'Strange infrared effect
        r = Abs(tR - 64)
        g = Abs(r - 64)
        b = Abs(g - 64)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = TC + 70
        r = r + (((r - 128) * 100) \ 100)
        g = Abs(TC - 6) + 70
        g = g + (((g - 128) * 100) \ 100)
        b = (TC + 4) + 70
        b = b + (((b - 128) * 100) \ 100)
        r = (r - TC) * 5
        g = (g - TC) * 5
        b = (b - TC) * 5
End Sub

'HSL <-> RGB conversion routines
Public Sub tRGBToHSL(r As Long, g As Long, b As Long, h As Single, s As Single, l As Single)
    Dim Max As Single
    Dim Min As Single
    Dim delta As Single
    Dim rR As Single, rG As Single, rB As Single
    
    rR = r / 255
    rG = g / 255
    rB = b / 255

    'Note: HSL are calculated in the following ranges:
    ' Hue: [-1,5]
    ' Saturation: [0,1] (Note that if saturation = 0, hue is technically undefined)
    ' Lightness: [0,1]

    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
        
    'Calculate luminance
    l = (Max + Min) / 2
        
    'If the maximum and minimum are identical, this image is gray, meaning it has no saturation and an undefined hue.
    If Max = Min Then
        s = 0
        h = 0
    Else
        
        'Calculate saturation
        If l <= 0.5 Then
            s = (Max - Min) / (Max + Min)
        Else
            s = (Max - Min) / (2 - Max - Min)
        End If
        
        'Calculate hue
        delta = Max - Min

        If rR = Max Then
            h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
        End If
        
        'If you prefer hue in the [0,360] range instead of [-1, 5] you can use this code
        'h = h * 60
        'If h < 0 Then h = h + 360

    End If

    'Tanner's final note: if byte values are preferred to floating-point, this code will return hue on [0,240],
    ' saturation on [0,255], and luminance on [0,255]
    'H = Int(H * 40 + 40)
    'S = Int(S * 255)
    'L = Int(L * 255)
    
End Sub

'Convert HSL values to RGB values
Public Sub tHSLToRGB(h As Single, s As Single, l As Single, r As Long, g As Long, b As Long)

    Dim rR As Single, rG As Single, rB As Single
    Dim Min As Single, Max As Single

    'Unsaturated pixels do not technically have hue - they only have luminance
    If s = 0 Then
        rR = l: rG = l: rB = l
    Else
        If l <= 0.5 Then
             Min = l * (1 - s)
        Else
            Min = l - s * (1 - l)
        End If
      
        Max = 2 * l - Min
      
        If (h < 1) Then
            
            rR = Max
            
            If (h < 0) Then
                rG = Min
                rB = rG - h * (Max - Min)
            Else
                rB = Min
                rG = h * (Max - Min) + rB
            End If
        
        ElseIf (h < 3) Then
            
            rG = Max
         
            If (h < 2) Then
                rB = Min
                rR = rB - (h - 2) * (Max - Min)
            Else
                rR = Min
                rB = (h - 2) * (Max - Min) + rR
            End If
        
        Else
        
            rB = Max
            
            If (h < 4) Then
                rR = Min
                rG = rR - (h - 4) * (Max - Min)
            Else
                rG = Min
                rR = (h - 4) * (Max - Min) + rG
            End If
         
        End If
            
   End If
   
   r = rR * 255
   g = rG * 255
   b = rB * 255
   
   'Failsafe added 29 August '12
   'This should never return RGB values > 255, but it doesn't hurt to make sure.
   If r > 255 Then r = 255
   If g > 255 Then g = 255
   If b > 255 Then b = 255
   
End Sub

'Return the maximum of three variables
Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
   If (rR > rG) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function

'Return the minimum of three variables
Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
   If (rR < rG) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function
