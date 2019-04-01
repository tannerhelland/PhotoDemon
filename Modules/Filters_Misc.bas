Attribute VB_Name = "Filters_Miscellaneous"
'***************************************************************************
'Filter Module
'Copyright 2000-2019 by Tanner Helland
'Created: 13/October/00
'Last updated: 07/September/15
'Last update: continued work on moving crap out of this module
'
'The general image filter module; contains unorganized routines at present.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
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
    PDImages.GetActiveImage.GetCompositedImage tmpImageComposite, True
    If (tmpImageComposite Is Nothing) Then Exit Sub
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
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
    
    Dim r As Long, g As Long, b As Long, a As Long
    
    'A special color counting class is used to count unique RGB and RGBA values
    Dim cTree As pdColorCount
    Set cTree = New pdColorCount
    cTree.SetAlphaTracking DIBs.IsDIBTransparent(tmpImageComposite)
    
    PDDebug.LogAction "Initial memory", PDM_Mem_Report
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Iterate through all pixels, counting unique values as we go.
    For y = initY To finalY
        tmpImageComposite.WrapArrayAroundScanline imageData, tmpSA, y
    For x = initX To finalX Step qvDepth
    
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        a = imageData(x + 3)
        cTree.AddColor r, g, b, a
        
    Next x
        If (y And progBarCheck) = 0 Then SetProgBarVal y
    Next y
    
    PDDebug.LogAction "Color count time: " & VBHacks.GetTimeDiffNowAsString(startTime)
    PDDebug.LogAction "Final memory", PDM_Mem_Report
    
    'Safely deallocate imageData()
    tmpImageComposite.UnwrapArrayFromDIB imageData
    Set tmpImageComposite = Nothing
    
    'Reset the progress bar
    SetProgBarVal 0
    ReleaseProgressBar
    
    'Show the user our final tally
    Dim numColorsRGB As Long, numColorsRGBA As Long
    numColorsRGB = cTree.GetUniqueRGBCount
    numColorsRGBA = cTree.GetUniqueRGBACount
    Set cTree = Nothing
    
    Message "Total colors: %1", numColorsRGB
    PDMsgBox "This image contains %1 unique colors (RGB)." & vbCrLf & vbCrLf & "This image contains %2 unique color + opacity values (RGBA).", vbOKOnly Or vbInformation, "Count image colors", numColorsRGB, numColorsRGBA
    
End Sub
