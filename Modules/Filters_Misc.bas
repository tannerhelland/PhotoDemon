Attribute VB_Name = "Filters_Miscellaneous"
'***************************************************************************
'Filter Module
'Copyright 2000-2026 by Tanner Helland
'Created: 13/October/00
'Last updated: 07/September/15
'Last update: continued work on moving crap out of this module
'
'The general image filter module; contains unorganized routines at present.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Subroutine for counting the number of unique colors in an image
Public Sub MenuCountColors()
    
    Message "Counting the number of unique colors in this image..."
    
    'Grab a composited copy of the image
    Dim tmpImageComposite As pdDIB
    PDImages.GetActiveImage.GetCompositedImage tmpImageComposite, True
    If (tmpImageComposite Is Nothing) Then Exit Sub
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
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
    For x = initX To finalX Step 4
    
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
    PDMsgBox "This image contains %1 unique colors (RGB)." & vbCrLf & vbCrLf & "This image contains %2 unique color + opacity values (RGBA).", vbOKOnly Or vbInformation, "Count unique colors", numColorsRGB, numColorsRGBA
    
End Sub
