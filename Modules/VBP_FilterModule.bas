Attribute VB_Name = "Filters_Area"
'***************************************************************************
'Filter (Area) Interface
'Copyright ©2000-2013 by Tanner Helland
'Created: 12/June/01
'Last updated: 08/September/12
'Last update: rewrote all filters against layers
'Still needs: interpolation for isometric conversion
'
'Holder for generalized area filters, including most of the project's convolution filters.
' Also contains the DoFilter routine, which is central to running custom filters
' (as well as many of the intrinsic PhotoDemon ones, like blur/sharpen/etc).
'
'***************************************************************************

Option Explicit

'These constants are related to saving/loading custom filters to/from a file
Public Const CUSTOM_FILTER_ID As String * 4 = "DScf"
Public Const CUSTOM_FILTER_VERSION_2003 = &H80000000
Public Const CUSTOM_FILTER_VERSION_2012 = &H80000001

'The omnipotent DoFilter routine - it takes whatever is in g_FM() - the "filter matrix" and applies it to the image
Public Sub DoFilter(Optional ByVal FilterType As String = "custom", Optional ByVal InvertResult As Boolean = False, Optional ByVal srcFilterFile As String = "", Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    'If requested, load the custom filter data from a file
    If srcFilterFile <> "" Then
        If toPreview = False Then Message "Loading custom filter information..."
        Dim FilterReturn As Boolean
        FilterReturn = LoadCustomFilterData(srcFilterFile)
        If FilterReturn = False Then
            Err.Raise 1024, PROGRAMNAME, "Invalid custom filter file"
            Exit Sub
        End If
    End If
    
    'Note that the only purpose of the FilterType string is to display this message
    If toPreview = False Then Message "Applying " & FilterType & " filter..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, x2 As Long, y2 As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    Dim checkXMin As Long, checkXMax As Long, checkYMin As Long, checkYMax As Long
    checkXMin = curLayerValues.MinX
    checkXMax = curLayerValues.MaxX
    checkYMin = curLayerValues.MinY
    checkYMax = curLayerValues.MaxY
            
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
    
    'CalcVar determines the size of each sub-loop (so that we don't waste time running a 5x5 matrix on 3x3 filters)
    Dim CalcVar As Long
    CalcVar = (g_FilterSize \ 2)
        
    'iFM() will hold the contents of g_FM() - the filter matrix; we don't use FM directly in case other events want to access it
    Dim iFM() As Long
    
    'Resize iFM according to the size of the filter matrix, then copy over the contents of g_FM()
    If g_FilterSize = 3 Then ReDim iFM(-1 To 1, -1 To 1) As Long Else ReDim iFM(-2 To 2, -2 To 2) As Long
    iFM = g_FM
    
    'FilterWeightA and FilterBiasA are copies of the global g_FilterWeight and g_FilterBias variables; again, we don't use the originals in case other events
    ' want to access them
    Dim FilterWeightA As Long, FilterBiasA As Long
    FilterWeightA = g_FilterWeight
    FilterBiasA = g_FilterBias
    
    'FilterWeightTemp will be reset for every pixel, and decremented appropriately when attempting to calculate the value for pixels
    ' outside the image perimeter
    Dim FilterWeightTemp As Long
    
    'Temporary calculation variables
    Dim CalcX As Long, CalcY As Long
    
    'Create a temporary layer and resize it to the same size as the current image
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer workingLayer
    
    'Create a local array and point it at the pixel data of our temporary layer.  This will be used to access the current pixel data
    ' without modifications, while the actual image data will be modified by the filter as it's processed.
    Dim tmpData() As Byte
    Dim tSA As SAFEARRAY2D
    prepSafeArray tSA, tmpLayer
    CopyMemory ByVal VarPtrArray(tmpData()), VarPtr(tSA), 4
    
    'QuickValInner is like QuickVal below, but for sub-loops
    Dim QuickValInner As Long
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Reset our values upon beginning analysis on a new pixel
        r = 0
        g = 0
        b = 0
        FilterWeightTemp = FilterWeightA
        
        'Run a sub-loop around the current pixel
        For x2 = x - CalcVar To x + CalcVar
            QuickValInner = x2 * qvDepth
        For y2 = y - CalcVar To y + CalcVar
        
            CalcX = x2 - x
            CalcY = y2 - y
            
            'If no filter value is being applied to this pixel, ignore it (GoTo's aren't generally a part of good programming,
            ' but because VB does not provide a "continue next" type mechanism, GoTo's are all we've got.)
            If iFM(CalcX, CalcY) = 0 Then GoTo NextCustomFilterPixel
            
            'If this pixel lies outside the image perimeter, ignore it and adjust g_FilterWeight accordingly
            If x2 < checkXMin Or y2 < checkYMin Or x2 > checkXMax Or y2 > checkYMax Then
                FilterWeightTemp = FilterWeightTemp - iFM(CalcX, CalcY)
                GoTo NextCustomFilterPixel
            End If
            
            'Adjust red, green, and blue according to the values in the filter matrix (FM)
            r = r + (tmpData(QuickValInner + 2, y2) * iFM(CalcX, CalcY))
            g = g + (tmpData(QuickValInner + 1, y2) * iFM(CalcX, CalcY))
            b = b + (tmpData(QuickValInner, y2) * iFM(CalcX, CalcY))

NextCustomFilterPixel:  Next y2
        Next x2
        
        'If a weight has been set, apply it now
        If (FilterWeightTemp <> 1) Then
            If (FilterWeightTemp <> 0) Then
                r = r \ FilterWeightTemp
                g = g \ FilterWeightTemp
                b = b \ FilterWeightTemp
            Else
                r = 0
                g = 0
                b = 0
            End If
        End If
        
        'If a bias has been specified, apply it now
        If FilterBiasA <> 0 Then
            r = r + FilterBiasA
            g = g + FilterBiasA
            b = b + FilterBiasA
        End If
        
        'Make sure all values are between 0 and 255
        If r < 0 Then
            r = 0
        ElseIf r > 255 Then
            r = 255
        End If
        
        If g < 0 Then
            g = 0
        ElseIf g > 255 Then
            g = 255
        End If
        
        If b < 0 Then
            b = 0
        ElseIf b > 255 Then
            b = 255
        End If
        
        'If inversion is specified, apply it now
        If InvertResult Then
            r = 255 - r
            g = 255 - g
            b = 255 - b
        End If
        
        'Finally, remember the new value in our tData array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() and tmpData() away from their respective DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    CopyMemory ByVal VarPtrArray(tmpData), 0&, 4
    Erase tmpData
    
    'Erase our temporary layer
    tmpLayer.eraseLayer
    Set tmpLayer = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'This subroutine will load the data from a custom filter file straight into the g_FM() array
Public Function LoadCustomFilterData(ByRef srcFilterPath As String) As Boolean
    
    'These are used to load values from the filter file; previously, they were integers, but in
    ' 2012 I changed them to Longs.  PhotoDemon loads both types.
    Dim tmpVal As Integer
    Dim tmpValLong As Long
    
    Dim x As Long, y As Long
    
    'Open the specified path
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open srcFilterPath For Binary As #fileNum
        
        'Verify that the filter is actually a valid filter file
        Dim VerifyID As String * 4
        Get #fileNum, 1, VerifyID
        If (VerifyID <> CUSTOM_FILTER_ID) Then
            Close #fileNum
            LoadCustomFilterData = False
            Exit Function
        End If
        'End verification
       
        'Next get the version number (gotta have this for backwards compatibility)
        Dim VersionNumber As Long
        Get #fileNum, , VersionNumber
        If (VersionNumber <> CUSTOM_FILTER_VERSION_2003) And (VersionNumber <> CUSTOM_FILTER_VERSION_2012) Then
            Message "Unsupported custom filter version."
            Close #fileNum
            LoadCustomFilterData = False
        End If
        'End version check
        
        If VersionNumber = CUSTOM_FILTER_VERSION_2003 Then
            Get #fileNum, , tmpVal
            g_FilterWeight = tmpVal
            Get #fileNum, , tmpVal
            g_FilterBias = tmpVal
        ElseIf VersionNumber = CUSTOM_FILTER_VERSION_2012 Then
            Get #fileNum, , tmpValLong
            g_FilterWeight = tmpValLong
            Get #fileNum, , tmpValLong
            g_FilterBias = tmpValLong
        End If
        
        'Resize the filter array to fit the default filter size
        g_FilterSize = 5
        ReDim g_FM(-2 To 2, -2 To 2) As Long
        'Dim a temporary array from which to load the array data
        Dim tFilterArray(0 To 24) As Long
        
        If VersionNumber = CUSTOM_FILTER_VERSION_2003 Then
            For x = 0 To 24
                Get #fileNum, , tmpVal
                tFilterArray(x) = tmpVal
            Next x
        ElseIf VersionNumber = CUSTOM_FILTER_VERSION_2012 Then
            For x = 0 To 24
                Get #fileNum, , tmpValLong
                tFilterArray(x) = tmpValLong
            Next x
        End If
        
        'Now dump the temporary array into the filter array
        For x = -2 To 2
        For y = -2 To 2
            g_FM(x, y) = tFilterArray((x + 2) + (y + 2) * 5)
        Next y
        Next x
    'Close the file up
    Close #fileNum
    LoadCustomFilterData = True
End Function

'A very, very gentle softening effect
Public Sub FilterAntialias()
    g_FilterSize = 3
    ReDim g_FM(-1 To 1, -1 To 1) As Long
    g_FM(-1, 0) = 1
    g_FM(1, 0) = 1
    g_FM(0, -1) = 1
    g_FM(0, 1) = 1
    g_FM(0, 0) = 6
    g_FilterWeight = 10
    g_FilterBias = 0
    DoFilter "Antialias"
End Sub

'"Soften an image" (aka, apply a gentle 3x3 blur)
Public Sub FilterSoften()
    
    g_FilterSize = 3
    ReDim g_FM(-1 To 1, -1 To 1) As Long
    
    g_FM(-1, -1) = 1
    g_FM(-1, 0) = 1
    g_FM(-1, 1) = 1
    
    g_FM(0, -1) = 1
    g_FM(0, 0) = 8
    g_FM(0, 1) = 1
    
    g_FM(1, -1) = 1
    g_FM(1, 0) = 1
    g_FM(1, 1) = 1
    
    g_FilterWeight = 16
    g_FilterBias = 0
    
    DoFilter "Soften"
    
End Sub

'"Soften an image more" (aka, apply a gentle 5x5 blur)
Public Sub FilterSoftenMore()
    
    g_FilterSize = 5
    ReDim g_FM(-2 To 2, -2 To 2) As Long
    
    g_FM(-2, -2) = 1
    g_FM(-2, -1) = 1
    g_FM(-2, 0) = 1
    g_FM(-2, 1) = 1
    g_FM(-2, 2) = 1
    
    g_FM(-1, -2) = 1
    g_FM(-1, -1) = 1
    g_FM(-1, 0) = 1
    g_FM(-1, 1) = 1
    g_FM(-1, 2) = 1
    
    g_FM(0, -2) = 1
    g_FM(0, -1) = 1
    g_FM(0, 0) = 24
    g_FM(0, 1) = 1
    g_FM(0, 2) = 1
    
    g_FM(1, -2) = 1
    g_FM(1, -1) = 1
    g_FM(1, 0) = 1
    g_FM(1, 1) = 1
    g_FM(1, 2) = 1
    
    g_FM(2, -2) = 1
    g_FM(2, -1) = 1
    g_FM(2, 0) = 1
    g_FM(2, 1) = 1
    g_FM(2, 2) = 1
    
    g_FilterWeight = 48
    g_FilterBias = 0
    
    DoFilter "Strong Soften"
    
End Sub

'Blur an image using a 3x3 convolution matrix
Public Sub FilterBlur()
        
    g_FilterSize = 3
    ReDim g_FM(-1 To 1, -1 To 1) As Long
    
    g_FM(-1, -1) = 1
    g_FM(-1, 0) = 1
    g_FM(-1, 1) = 1
    
    g_FM(0, -1) = 1
    g_FM(0, 0) = 1
    g_FM(0, 1) = 1
    
    g_FM(1, -1) = 1
    g_FM(1, 0) = 1
    g_FM(1, 1) = 1
    
    g_FilterWeight = 9
    g_FilterBias = 0
    
    DoFilter "Blur"
    
End Sub

'Blur an image using a 5x5 convolution matrix
Public Sub FilterBlurMore()
    
    g_FilterSize = 5
    ReDim g_FM(-2 To 2, -2 To 2) As Long
    
    g_FM(-2, -2) = 1
    g_FM(-2, -1) = 1
    g_FM(-2, 0) = 1
    g_FM(-2, 1) = 1
    g_FM(-2, 2) = 1
    
    g_FM(-1, -2) = 1
    g_FM(-1, -1) = 1
    g_FM(-1, 0) = 1
    g_FM(-1, 1) = 1
    g_FM(-1, 2) = 1
    
    g_FM(0, -2) = 1
    g_FM(0, -1) = 1
    g_FM(0, 0) = 1
    g_FM(0, 1) = 1
    g_FM(0, 2) = 1
    
    g_FM(1, -2) = 1
    g_FM(1, -1) = 1
    g_FM(1, 0) = 1
    g_FM(1, 1) = 1
    g_FM(1, 2) = 1
    
    g_FM(2, -2) = 1
    g_FM(2, -1) = 1
    g_FM(2, 0) = 1
    g_FM(2, 1) = 1
    g_FM(2, 2) = 1
    
    g_FilterWeight = 25
    g_FilterBias = 0
    
    DoFilter "Strong Blur"
    
End Sub

'3x3 Gaussian blur
Public Sub FilterGaussianBlur()

    g_FilterSize = 3
    ReDim g_FM(-1 To 1, -1 To 1) As Long
    
    g_FM(-1, -1) = 1
    g_FM(0, -1) = 2
    g_FM(1, -1) = 1
    
    g_FM(-1, 0) = 2
    g_FM(0, 0) = 4
    g_FM(1, 0) = 2
    
    g_FM(-1, 1) = 1
    g_FM(0, 1) = 2
    g_FM(1, 1) = 1
    
    g_FilterWeight = 16
    g_FilterBias = 0
    
    DoFilter "Gaussian Blur"
    
End Sub

'5x5 Gaussian blur
Public Sub FilterGaussianBlurMore()

    g_FilterSize = 5
    ReDim g_FM(-2 To 2, -2 To 2) As Long
    
    g_FM(-2, -2) = 1
    g_FM(-1, -2) = 4
    g_FM(0, -2) = 7
    g_FM(1, -2) = 4
    g_FM(2, -2) = 1
    
    g_FM(-2, -1) = 4
    g_FM(-1, -1) = 16
    g_FM(0, -1) = 26
    g_FM(1, -1) = 16
    g_FM(2, -1) = 4
    
    g_FM(-2, 0) = 7
    g_FM(-1, 0) = 26
    g_FM(0, 0) = 41
    g_FM(1, 0) = 26
    g_FM(2, 0) = 7
    
    g_FM(-2, 1) = 4
    g_FM(-1, 1) = 16
    g_FM(0, 1) = 26
    g_FM(1, 1) = 16
    g_FM(2, 1) = 4
    
    g_FM(-2, 2) = 1
    g_FM(-1, 2) = 4
    g_FM(0, 2) = 7
    g_FM(1, 2) = 4
    g_FM(2, 2) = 1
    
    g_FilterWeight = 273
    g_FilterBias = 0
    
    DoFilter "Strong Gaussian Blur"
    
End Sub

'Sharpen an image via convolution filter
Public Sub FilterSharpen()
    
    g_FilterSize = 3
    ReDim g_FM(-1 To 1, -1 To 1) As Long
    
    g_FM(-1, -1) = -1
    g_FM(0, -1) = -1
    g_FM(1, -1) = -1
    
    g_FM(-1, 0) = -1
    g_FM(0, 0) = 15
    g_FM(1, 0) = -1
    
    g_FM(-1, 1) = -1
    g_FM(0, 1) = -1
    g_FM(1, 1) = -1
    
    g_FilterWeight = 7
    g_FilterBias = 0
    
    DoFilter "Sharpen"
  
End Sub

'Strongly sharpen an image via convolution filter
Public Sub FilterSharpenMore()

    g_FilterSize = 3
    ReDim g_FM(-1 To 1, -1 To 1) As Long
    
    g_FM(-1, -1) = 0
    g_FM(0, -1) = -1
    g_FM(1, -1) = 0
    
    g_FM(-1, 0) = -1
    g_FM(0, 0) = 5
    g_FM(1, 0) = -1
    
    g_FM(-1, 1) = 0
    g_FM(0, 1) = -1
    g_FM(1, 1) = 0
    
    g_FilterWeight = 1
    g_FilterBias = 0
    
    DoFilter "Strong Sharpen"
  
End Sub

'"Unsharp" an image - it's a stupid name, but that's the industry standard.  Basically, blur the image, then subtract that from the original image.
Public Sub FilterUnsharp()

    g_FilterSize = 3
    ReDim g_FM(-1 To 1, -1 To 1) As Long
    
    g_FM(-1, -1) = -1
    g_FM(0, -1) = -2
    g_FM(1, -1) = -1
    
    g_FM(-1, 0) = -2
    g_FM(0, 0) = 24
    g_FM(1, 0) = -2
    
    g_FM(-1, 1) = -1
    g_FM(0, 1) = -2
    g_FM(1, 1) = -1
    
    g_FilterWeight = 12
    g_FilterBias = 0
    
    DoFilter "Unsharp"
  
  End Sub

'Apply a grid blur to an image; basically, blur every vertical line, then every horizontal line, then average the results
Public Sub FilterGridBlur()

    Message "Generating grids..."

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
    
    Dim iWidth As Long, iHeight As Long
    iWidth = curLayerValues.Width
    iHeight = curLayerValues.Height
            
    Dim NumOfPixels As Long
    NumOfPixels = iWidth + iHeight
            
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
    Dim rax() As Long, gax() As Long, bax() As Long
    Dim ray() As Long, gay() As Long, bay() As Long
    ReDim rax(0 To iWidth) As Long, gax(0 To iWidth) As Long, bax(0 To iWidth) As Long
    ReDim ray(0 To iHeight) As Long, gay(0 To iHeight), bay(0 To iHeight)
    
    'Generate the averages for vertical lines
    For x = initX To finalX
        r = 0
        g = 0
        b = 0
        QuickVal = x * qvDepth
        For y = initY To finalY
            r = r + ImageData(QuickVal + 2, y)
            g = g + ImageData(QuickVal + 1, y)
            b = b + ImageData(QuickVal, y)
        Next y
        rax(x) = r
        gax(x) = g
        bax(x) = b
    Next x
    
    'Generate the averages for horizontal lines
    For y = initY To finalY
        r = 0
        g = 0
        b = 0
        For x = initX To finalX
            QuickVal = x * qvDepth
            r = r + ImageData(QuickVal + 2, y)
            g = g + ImageData(QuickVal + 1, y)
            b = b + ImageData(QuickVal, y)
        Next x
        ray(y) = r
        gay(y) = g
        bay(y) = b
    Next y
    
    Message "Applying grid blur..."
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Average the horizontal and vertical values for each color component
        r = (rax(x) + ray(y)) \ NumOfPixels
        g = (gax(x) + gay(y)) \ NumOfPixels
        b = (bax(x) + bay(y)) \ NumOfPixels
        
        'The colors shouldn't exceed 255, but it doesn't hurt to double-check
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        'Assign the new RGB values back into the array
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

'Convert an image to its isometric equivalent.  This can be very useful for developers of isometric games.
Public Sub FilterIsometric()

    'If a selection is active, remove it.  (This is not the most elegant solution, but we can fix it at a later date.)
    If pdImages(CurrentImage).selectionActive Then
        pdImages(CurrentImage).selectionActive = False
        tInit tSelection, False
    End If

    Message "Preparing conversion tables..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepImageData srcSA
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Make note of the current image's width and height
    Dim hWidth As Single
    Dim oWidth As Long, oHeight As Long
    oWidth = curLayerValues.Width - 1
    oHeight = curLayerValues.Height - 1
    hWidth = oWidth / 2
    
    Dim nWidth As Long, nHeight As Long
    nWidth = oWidth + oHeight + 1
    nHeight = nWidth \ 2
    
    'Create a second local array.  This will contain the pixel data of the new isometric image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    
    Dim dstLayer As pdLayer
    Set dstLayer = New pdLayer
    dstLayer.createBlank nWidth + 1, nHeight + 1, pdImages(CurrentImage).mainLayer.getLayerColorDepth
    
    prepSafeArray dstSA, dstLayer
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    Dim srcX As Single, srcY As Single
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim dstQuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
        
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    SetProgBarMax nWidth
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
        
    'Interpolated loop calculation
    Dim lOffset As Long
        
    Message "Converting image to isometric view..."
        
    'Run through the destination image pixels, converting to isometric as we go
    For x = 0 To nWidth
        dstQuickVal = x * qvDepth
    For y = 0 To nHeight
        
        srcX = getIsometricX(x, y, hWidth)
        srcY = getIsometricY(x, y, hWidth)
                
        'If the pixel is inside the image, reverse-map it using bilinear interpolation.
        ' (Note: this will also reverse-map alpha values if they are present in the image.)
        If (srcX >= 0 And srcX < oWidth And srcY >= 0 And srcY < oHeight) Then
            
            For lOffset = 0 To qvDepth - 1
                dstImageData(dstQuickVal + lOffset, y) = getInterpolatedVal(srcX, srcY, srcImageData, lOffset, qvDepth)
            Next lOffset
                    
        'Out-of-bound pixels don't need interpolation - just set them manually
        Else
            'If the image is 32bpp, set outlying pixels as fully transparent
            If qvDepth = 4 Then dstImageData(dstQuickVal + 3, y) = 0
            dstImageData(dstQuickVal + 2, y) = 255
            dstImageData(dstQuickVal + 1, y) = 255
            dstImageData(dstQuickVal, y) = 255
        End If
        
    
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
    
    'With our work complete, point both ImageData() arrays away from their respective DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'dstImageData now contains the isometric image.  We need to transfer that back into the current image.
    pdImages(CurrentImage).mainLayer.eraseLayer
    pdImages(CurrentImage).mainLayer.createFromExistingLayer dstLayer
    
    'With that transfer complete, we can erase our temporary layer
    dstLayer.eraseLayer
    Set dstLayer = Nothing
    
    'Update the current image size
    pdImages(CurrentImage).updateSize
    DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
    
    Message "Finished. "
    
    'Redraw the image
    FitOnScreen
    
    'Reset the progress bar to zero
    SetProgBarVal 0

End Sub

'These two functions translate a normal (x,y) coordinate to an isometric plane
Private Function getIsometricX(ByVal xc As Long, ByVal yc As Long, ByVal tWidth As Long) As Single
    getIsometricX = (xc / 2) - yc + tWidth
End Function

Private Function getIsometricY(ByVal xc As Long, ByVal yc As Long, ByVal tWidth As Long) As Single
    getIsometricY = (xc / 2) + yc - tWidth
End Function

'This function takes an x and y value - as floating-point - and uses their position to calculate an interpolated value
' for an imaginary pixel in that location.  Offset (r/g/b/alpha) and image color depth are also required.
Public Function getInterpolatedVal(ByVal x1 As Double, ByVal y1 As Double, ByRef iData() As Byte, ByRef iOffset As Long, ByRef iDepth As Long) As Byte
        
    'Retrieve the four surrounding pixel values
    Dim topLeft As Double, topRight As Double, bottomLeft As Double, bottomRight As Double
    topLeft = iData(Int(x1) * iDepth + iOffset, Int(y1))
    topRight = iData(Int(x1 + 1) * iDepth + iOffset, Int(y1))
    bottomLeft = iData(Int(x1) * iDepth + iOffset, Int(y1 + 1))
    bottomRight = iData(Int(x1 + 1) * iDepth + iOffset, Int(y1 + 1))
    
    'Calculate blend ratios
    Dim yBlend As Double
    Dim xBlend As Double, xBlendInv As Double
    yBlend = y1 - Int(y1)
    xBlend = x1 - Int(x1)
    xBlendInv = 1 - xBlend
    
    'Blend in the x-direction
    Dim topRowColor As Double, bottomRowColor As Double
    topRowColor = topRight * xBlend + topLeft * xBlendInv
    bottomRowColor = bottomRight * xBlend + bottomLeft * xBlendInv
    
    'Blend in the y-direction
    getInterpolatedVal = bottomRowColor * yBlend + topRowColor * (1 - yBlend)

End Function

'This function takes an x and y value - as floating-point - and uses their position to calculate an interpolated value
' for an imaginary pixel in that location.  Offset (r/g/b/alpha) and image color depth are also required.
Public Function getInterpolatedValWrap(ByVal x1 As Double, ByVal y1 As Double, ByVal xMax As Long, yMax As Long, ByRef iData() As Byte, ByRef iOffset As Long, ByRef iDepth As Long) As Byte
        
    'Retrieve the four surrounding pixel values
    Dim topLeft As Double, topRight As Double, bottomLeft As Double, bottomRight As Double
    topLeft = iData(Int(x1) * iDepth + iOffset, Int(y1))
    If Int(x1) = xMax Then
        topRight = iData(0 + iOffset, Int(y1))
    Else
        topRight = iData(Int(x1 + 1) * iDepth + iOffset, Int(y1))
    End If
    If Int(y1) = yMax Then
        bottomLeft = iData(Int(x1) * iDepth + iOffset, 0)
    Else
        bottomLeft = iData(Int(x1) * iDepth + iOffset, Int(y1 + 1))
    End If
    
    If Int(x1) = xMax Then
        If Int(y1) = yMax Then
            bottomRight = iData(0 + iOffset, 0)
        Else
            bottomRight = iData(0 + iOffset, Int(y1 + 1))
        End If
    Else
        If Int(y1) = yMax Then
            bottomRight = iData(Int(x1 + 1) * iDepth + iOffset, 0)
        Else
            bottomRight = iData(Int(x1 + 1) * iDepth + iOffset, Int(y1 + 1))
        End If
    End If
    
    'Calculate blend ratios
    Dim yBlend As Double
    Dim xBlend As Double, xBlendInv As Double
    yBlend = y1 - Int(y1)
    xBlend = x1 - Int(x1)
    xBlendInv = 1 - xBlend
    
    'Blend in the x-direction
    Dim topRowColor As Double, bottomRowColor As Double
    topRowColor = topRight * xBlend + topLeft * xBlendInv
    bottomRowColor = bottomRight * xBlend + bottomLeft * xBlendInv
    
    'Blend in the y-direction
    getInterpolatedValWrap = bottomRowColor * yBlend + topRowColor * (1 - yBlend)

End Function

