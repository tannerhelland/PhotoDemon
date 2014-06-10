Attribute VB_Name = "Filters_Area"
'***************************************************************************
'Filter (Area) Interface
'Copyright ©2001-2014 by Tanner Helland
'Created: 12/June/01
'Last updated: 10/June/14
'Last update: rewrite central convolution function to accept source/destination layers; this will allow us to use it from
'              any arbitrary internal function.
'
'Holder module for generalized area filters, including most of the project's convolution filters.
'
'The most interesting function is ConvolveDIB, which applies any arbitrary 5x5 convolution filter to any arbitrary DIB.  This function
' is used internally for nearly all edge-detection functions, and other generic convolution effects.  (Note that some convolution
' filters, like Gaussian Blur, have their own specialized, optimized implementations.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'These constants are related to saving/loading custom filters to/from a file
Public Const CUSTOM_FILTER_ID As String * 4 = "DScf"
Public Const CUSTOM_FILTER_VERSION_2003 = &H80000000
Public Const CUSTOM_FILTER_VERSION_2012 = &H80000001
Public Const CUSTOM_FILTER_VERSION_2014 As String = "8.2014"

'The omnipotent ApplyConvolutionFilter routine, which applies the supplied convolution filter to the current image.
' Note that as of June '13, ApplyConvolutionFilter uses a full param string for supplying convolution details.  The relevant
' ParamString format is as follows:
'    Name: String (can't be blank, but can be a single space)
'    Invert: Boolean
'    Divisor: Double
'    Offset: Long
'    25 Double values, which correspond to entries in a 5x5 convolution matrix, in left-to-right, top-to-bottom order.
Public Sub ApplyConvolutionFilter(ByVal fullParamString As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
        
    'Prepare a param parser
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString fullParamString
        
    'Note that the only purpose of the FilterType string is to display this message
    If Not toPreview Then Message "Applying %1 filter...", cParams.GetString(1)
    
    'Create a local array and point it at the pixel data of the current image.  Note that the current layer is referred to as the
    ' DESTINATION image for the convolution; we will make a separate temp copy of the image to use as the SOURCE.
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent processed pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    
    'Use the central ConvolveDIB function to apply the convolution
    ConvolveDIB fullParamString, srcDIB, workingDIB, toPreview
    
    
    'Free our temporary DIB
    srcDIB.eraseDIB
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic
        
End Sub

'Apply any convolution filter to a pdDIB object.  This is primarily used by the ApplyConvolutionFilter function, above, but can also be linked
' internally to apply multiple convolutions in succession, or to create standalone convolved images that can then be blended together.
Public Function ConvolveDIB(ByVal fullParamString As String, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Prepare a param parser; this is necessary for parsing out the individual convolution parameters from the param string
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString fullParamString
    
    'Create a local array and point it at the destination pixel data
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent processed pixel values from corrupting subsequent calculations.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, x2 As Long, y2 As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
    
    Dim checkXMin As Long, checkXMax As Long, checkYMin As Long, checkYMax As Long
    checkXMin = initX
    checkXMax = finalX
    checkYMin = initY
    checkYMax = finalY
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
        
    'We can now parse out the relevant filter values from the param string
    Dim invertResult As Boolean
    invertResult = cParams.GetBool(2)
    
    Dim FilterWeightA As Double, FilterBiasA As Double
    FilterWeightA = cParams.GetDouble(3)
    FilterBiasA = cParams.GetDouble(4)
    
    Dim iFM(-2 To 2, -2 To 2) As Double
    For x = -2 To 2
    For y = -2 To 2
        iFM(x, y) = cParams.GetDouble((x + 2) + (y + 2) * 5 + 5)
    Next y
    Next x
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    
    'FilterWeightTemp will be reset for every pixel, and decremented appropriately when attempting to calculate the value for pixels
    ' outside the image perimeter
    Dim FilterWeightTemp As Double
    
    'Temporary calculation variables
    Dim CalcX As Long, CalcY As Long
    
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
        For x2 = x - 2 To x + 2
            QuickValInner = x2 * qvDepth
        For y2 = y - 2 To y + 2
        
            CalcX = x2 - x
            CalcY = y2 - y
            
            'If no filter value is being applied to this pixel, ignore it (GoTo's aren't generally a part of good programming,
            ' but because VB does not provide a "continue next" type mechanism, GoTo's are all we've got.)
            If iFM(CalcX, CalcY) <> 0 Then
            
                'If this pixel lies outside the image perimeter, ignore it and adjust g_FilterWeight accordingly
                If (x2 < checkXMin) Or (y2 < checkYMin) Or (x2 > checkXMax) Or (y2 > checkYMax) Then
                    
                    FilterWeightTemp = FilterWeightTemp - iFM(CalcX, CalcY)
                
                Else
                
                    'Adjust red, green, and blue according to the values in the filter matrix (FM)
                    r = r + (srcImageData(QuickValInner + 2, y2) * iFM(CalcX, CalcY))
                    g = g + (srcImageData(QuickValInner + 1, y2) * iFM(CalcX, CalcY))
                    b = b + (srcImageData(QuickValInner, y2) * iFM(CalcX, CalcY))
                    
                End If
                
            End If
    
        Next y2
        Next x2
        
        'If a weight has been set, apply it now
        If (FilterWeightTemp <> 1) Then
        
            'Catch potential divide-by-zero errors
            If (FilterWeightTemp <> 0) Then
                r = r / FilterWeightTemp
                g = g / FilterWeightTemp
                b = b / FilterWeightTemp
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
        If invertResult Then
            r = 255 - r
            g = 255 - g
            b = 255 - b
        End If
        
        'Copy the calculated value into the destination array
        dstImageData(QuickVal + 2, y) = r
        dstImageData(QuickVal + 1, y) = g
        dstImageData(QuickVal, y) = b
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() and srcImageData() away from their respective DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
        
    'Return success/failure
    If cancelCurrentAction Then ConvolveDIB = 0 Else ConvolveDIB = 1

End Function

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
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim iWidth As Long, iHeight As Long
    iWidth = curDIBValues.Width
    iHeight = curDIBValues.Height
            
    Dim NumOfPixels As Long
    NumOfPixels = iWidth + iHeight
            
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
