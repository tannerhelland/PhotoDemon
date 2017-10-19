Attribute VB_Name = "Filters_Miscellaneous"
'***************************************************************************
'Filter Module
'Copyright 2000-2017 by Tanner Helland
'Created: 13/October/00
'Last updated: 07/September/15
'Last update: continued work on moving crap out of this module
'
'The general image filter module; contains unorganized routines at present.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Dull but standard "sepia" transformation.  Values derived from the w3c standard at:
' https://dvcs.w3.org/hg/FXTF/raw-file/tip/filters/index.html#sepiaEquivalent
Public Sub MenuSepia()
    
    Message "Engaging hipsters to perform sepia conversion..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Double, newG As Double, newB As Double
        
    'Apply the filter
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
                
        newR = CSng(r) * 0.393 + CSng(g) * 0.769 + CSng(b) * 0.189
        newG = CSng(r) * 0.349 + CSng(g) * 0.686 + CSng(b) * 0.168
        newB = CSng(r) * 0.272 + CSng(g) * 0.534 + CSng(b) * 0.131
        
        r = newR
        g = newG
        b = newB
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        imageData(quickVal + 2, y) = r
        imageData(quickVal + 1, y) = g
        imageData(quickVal, y) = b
        
    Next y
        If (x And progBarCheck) = 0 Then
            If Interface.UserPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub

'Subroutine for counting the number of unique colors in an image
Public Sub MenuCountColors()
    
    Message "Counting the number of unique colors in this image..."
    
    'Grab a composited copy of the image
    Dim tmpImageComposite As pdDIB
    pdImages(g_CurrentImage).GetCompositedImage tmpImageComposite, True
    If (tmpImageComposite Is Nothing) Then Exit Sub
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepSafeArray tmpSA, tmpImageComposite
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim qvDepth As Long
    qvDepth = tmpImageComposite.GetDIBColorDepth \ 8
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = tmpImageComposite.GetDIBStride - 1
    finalY = tmpImageComposite.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'This array will track whether or not a given color has been detected in the image
    Dim uniqueColors() As Byte
    ReDim uniqueColors(0 To 16777216) As Byte
    
    'Total number of unique colors counted so far
    Dim totalCount As Long
    totalCount = 0
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim chkValue As Long
        
    'Apply the filter
    For y = initY To finalY
    For x = initX To finalX Step qvDepth
        b = imageData(x, y)
        g = imageData(x + 1, y)
        r = imageData(x + 2, y)
        
        chkValue = RGB(r, g, b)
        If uniqueColors(chkValue) = 0 Then
            totalCount = totalCount + 1
            uniqueColors(chkValue) = 1
        End If
    Next x
        If (y And progBarCheck) = 0 Then SetProgBarVal y
    Next y
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    Erase uniqueColors
    Set tmpImageComposite = Nothing
    
    'Reset the progress bar
    SetProgBarVal 0
    ReleaseProgressBar
    
    'Show the user our final tally
    Message "Total unique colors: %1", totalCount
    PDMsgBox "This image contains %1 unique colors.", vbOKOnly Or vbInformation, "Count image colors", totalCount
    
End Sub

'Placeholder function I'm using to remind myself how to best use the new palette generator functions.
Public Sub MenuApplyTestPalette()

    'Create a local array and point it at the pixel data we want to operate on
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Make a smaller, localized copy of the DIB.  (50k pixels is more than enough for accurate
    ' palette generation, and using a fixed size guarantees roughly O(1) time for palette generation.)
    Dim megapixelSize As Long
    megapixelSize = 50000
    Dim smallDIB As pdDIB
    If DIBs.ResizeDIBByPixelCount(workingDIB, smallDIB, megapixelSize) Then
        
        'Construct an optimized palette based on the small image
        Dim srcPalette() As RGBQuad
        If Palettes.GetOptimizedPalette(smallDIB, srcPalette, 256) Then
        
            'Apply the optimized palette to the full-sized DIB.
            'Palettes.ApplyPaletteToImage workingDIB, srcPalette
            'Palettes.ApplyPaletteToImage_SysAPI workingDIB, srcPalette
            'Palettes.ApplyPaletteToImage_LossyHashTable workingDIB, srcPalette
            Palettes.ApplyPaletteToImage_Octree workingDIB, srcPalette
            
        End If
        
    End If
    
    'Debug.Print "Finished: " & VBHacks.GetTimerDifferenceNow(startTime) * 1000
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub
