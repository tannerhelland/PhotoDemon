VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdSeamCarving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Image seam carving (smart resize/"content-aware scale"/"liquid rescale") Engine
'Copyright 2014-2026 by Tanner Helland
'Created: 06/January/14
'Last updated: 02/April/20
'Last update: fixed some remaining 32-bpp issues, minor perf and memory usage improvements
'
'Content-aware scaling is a very exciting addition to PhotoDemon 6.2.  (As a comparison,
' PhotoShop didn't gain this feature until CS4, so it's relatively modern stuff!)
'
'Normal scaling algorithms work by shrinking or enlarging all image pixels equally.
' Such algorithms make no distinction between visually important pixels and visually
' unimportant ones.  Unfortunately, when the aspect ratio of an image is changed using
' such an algorithm, noticeable distortion results, and the end result looks unpleasant.
'
'Content-aware scaling tries to circumvent this by selectively removing the least visually
' important parts of an image (as determined by some type of per-pixel "energy" calculation).
' By preferentially removing uninteresting pixels over interesting ones, important parts
' of an image can be preserved while uninteresting parts are removed.  The result is often
' a more aesthetically pleasing image, even under severe aspect ratio changes.
'
'For reference, the original 2007 paper that first proposed this technique - called
' "seam carving" - is available here: http://www.win.tue.nl/~wstahw/edu/2IV05/seamcarving.pdf
'
'I have written PhotoDemon's implementation from scratch, using the original paper as my
' primary resource.  Unfortunately, my current implementation is somewhat slow (though still
' faster than many other open-source implementations!) on account of all the seam finding tasks
' we have to perform.  Ideas/patches for improved performance are welcome.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Source image, which will be carved into smaller and smaller bits by the function
Private m_SourceImage As pdDIB

'Energy array, which describes the energy of each pixel in the source image.
' MAKE CERTAIN ITS DIMENSIONS MATCH THE SOURCE IMAGE; otherwise, you will get errors.
' Generally speaking, this class does not care *how* the energy image is generated,
' just that the data it represents is valid and reasonably accurate.
Private m_Energy() As Byte

'Final image, which will only exist after a seam carve operation has been carried out
Private m_DestinationImage As pdDIB

'Shrinking and enlarging are handled separately
Private Enum PD_SeamCarveScaling
    scs_Shrink = 0
    scs_None = 1
    scs_Grow = 2
End Enum

#If False Then
    Private Const scs_Shrink = 0, scs_None = 1, scs_Grow = 2
#End If

'For performance reasons, the seam carver class maintains its own copy of the source image.
Friend Sub SetSourceImage(ByRef srcDIB As pdDIB)
    
    If (m_SourceImage Is Nothing) Then Set m_SourceImage = New pdDIB
    m_SourceImage.CreateFromExistingDIB srcDIB
    
    'Generate the initial energy map
    Filters_Scientific.GetImageGrad_MagOnly m_SourceImage, m_Energy
    
End Sub

'Retrieve the carved image into a DIB of the calling function's choosing
Friend Function GetCarvedImage() As pdDIB
    Set GetCarvedImage = m_DestinationImage
End Function

'Generate a vertical seam map for the current energy image.  Because those images may have been shrunk in one (or more)
' directions by previous seam map operations, a final X and Y value are also explicitly supplied, so the array can be
' re-scanned while ignoring irrelevant pixels.
Friend Sub GetVerticalSeamMap(ByRef seamMap() As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef maxVal As Long, ByRef minVal As Long, ByRef minRow As Long, Optional ByVal scanDirectionNormal As Boolean = True)
    
    Dim x As Long, y As Long, xCheck As Long
    
    'Populate the first row of the energy array by simply copying over the relevant energy values
    For y = 0 To finalY
        seamMap(0, y) = m_Energy(0, y)
    Next y
    
    'Now we can start traversing the energy array.
    Dim leftUp As Long, leftMiddle As Long, leftDown As Long, curEnergy As Long
    
    'At each point, generate a new energy for the pixel using the smallest energy value above said pixel
    For x = 1 To finalX
        xCheck = (x - 1)
    For y = 0 To finalY
        
        leftMiddle = seamMap(xCheck, y)
        curEnergy = m_Energy(x, y)
        
        'Note that we must check edge pixels differently; hence the nested IF statements here
        If (y > 0) Then
        
            leftUp = seamMap(xCheck, y - 1)
            If (y < finalY) Then
            
                leftDown = seamMap(xCheck, y + 1)
            
                'This is not a left or right edge pixel.  Check all three pixels above for a minimum value.
                If (leftUp < leftMiddle) Then
                    If (leftUp < leftDown) Then seamMap(x, y) = curEnergy + leftUp Else seamMap(x, y) = curEnergy + leftDown
                Else
                    If (leftMiddle < leftDown) Then seamMap(x, y) = curEnergy + leftMiddle Else seamMap(x, y) = curEnergy + leftDown
                End If
            
            Else
            
                'This is a right edge pixel.  Check only two pixels above.
                If (leftUp < leftMiddle) Then seamMap(x, y) = curEnergy + leftUp Else seamMap(x, y) = curEnergy + leftMiddle
                
            End If
        
        'This is a left edge pixel.  Check only two pixels above.
        Else
        
            leftDown = seamMap(xCheck, y + 1)
            If (leftMiddle < leftDown) Then seamMap(x, y) = curEnergy + leftMiddle Else seamMap(x, y) = curEnergy + leftDown
            
        End If
        
    Next y
    Next x
    
    'The seamMap array now contains a cumulative energy map for the image, which we can reverse-track to
    ' find out which seams should be removed!
    
    'As a convenience to subsequent functions, this function also returns the maximum value of the seam map.
    ' Processed pixels can be set to this value, which prevents them from being re-selected on subsequent runs.
    maxVal = 0
    minVal = LONG_MAX
    
    Dim scanStart As Long, scanEnd As Long, scanStep As Long
    
    If scanDirectionNormal Then
        scanStart = 0
        scanEnd = finalY
        scanStep = 1
    Else
        scanStart = finalY
        scanEnd = 0
        scanStep = -1
    End If
    
    For y = scanStart To scanEnd Step scanStep
    
        curEnergy = seamMap(finalX, y)
        If (curEnergy > maxVal) Then maxVal = curEnergy
        
        If (curEnergy < minVal) Then
            minVal = curEnergy
            minRow = y
        End If
        
    Next y
        
End Sub

'Given a valid seam map, remove one horizontal seam (which will result in a 1px smaller vertical image)
Private Sub ShrinkVertically(ByRef seamMap() As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef maxVal As Long, ByRef minVal As Long, ByRef rowIndex As Long)

    'Start by finding the smallest energy value in the final column of the seam map
    Dim x As Long, y As Long
    
    Dim xCheck As Long
    Dim leftUp As Long, leftMiddle As Long, leftDown As Long
    
    'rowIndex now contains the y coordinate of the minimum energy seam terminus.  Starting there, traverse the
    ' image leftward, removing lowest-energy values as we go (and shifting all data past that pixel upward).
    For x = finalX To 0 Step -1
    
        'Remove the minimum value from the energy map and shift all corresponding data up.
        If (rowIndex < finalY) Then
            
            'Both the energy image and source image must have their data shifted.
            
            'Source image is easy - just use BitBlt
            GDI.BitBltWrapper m_SourceImage.GetDIBDC, x, rowIndex, 1, finalY - rowIndex, m_SourceImage.GetDIBDC, x, rowIndex + 1, vbSrcCopy
            
            'Energy array must be done manually
            For y = rowIndex + 1 To finalY
                m_Energy(x, y - 1) = m_Energy(x, y)
            Next y
            
        End If
        
        'Find the minimum value of the next row left.
        If (x > 0) Then
        
            xCheck = x - 1
            leftMiddle = seamMap(xCheck, rowIndex)
        
            'Note that we must check edge pixels differently; hence the nested IF statements here
            If (rowIndex > 0) Then
        
                leftUp = seamMap(xCheck, rowIndex - 1)
                If (rowIndex < finalY) Then
                
                    leftDown = seamMap(xCheck, rowIndex + 1)
                
                    'This is not a left or right edge pixel.  Check all three pixels above for a minimum value.
                    If (leftUp < leftMiddle) Then
                        If (leftUp < leftDown) Then rowIndex = rowIndex - 1 Else rowIndex = rowIndex + 1
                    Else
                        If (leftMiddle > leftDown) Then rowIndex = rowIndex + 1
                    End If
                
                Else
                
                    'This is a right edge pixel.  Check only two pixels above.
                    If (leftUp < leftMiddle) Then rowIndex = rowIndex - 1
                    
                End If
            
            'This is a left edge pixel.  Check only two pixels above.
            Else
            
                leftDown = seamMap(xCheck, rowIndex + 1)
                If (leftMiddle > leftDown) Then rowIndex = rowIndex + 1
                
            End If
            
        End If
    
    Next x

End Sub

'Given a valid seam map, add n horizontal seams (which will result in an n-px larger vertical image)
Private Sub GrowVertically(ByRef seamMap() As Long, ByVal numOfAddedSeams As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef maxVal As Long, ByRef minVal As Long, ByRef rowIndex As Long, ByRef progBarOffset As Long)

    'Due to limitations of the seam carving algorithm, this function can only add a max of
    ' [imageheight \ 2] seams at a time.  If the user requires more than that, call this
    ' function again.
    If (numOfAddedSeams > m_SourceImage.GetDIBHeight \ 2) Then numOfAddedSeams = m_SourceImage.GetDIBHeight \ 2
    
    Dim x As Long, y As Long, i As Long
    
    'Growing seams is a bit of a different approach from removing seams.  Because added seams will always have minimal energy,
    ' we must vary the order in which seams are added.  To do this, we generate a "visibility map".  This map (a 2D array the size
    ' of the image) will store how many times each source pixel is duplicated in the destination image.  Eventually it would be nice
    ' to add some interpolation capabilities to this function, but for now, we only do blind pixel insertion.
    Dim visibilityMap() As Long
    ReDim visibilityMap(0 To finalX, 0 To finalY) As Long
    
    Dim xCheck As Long
    Dim leftUp As Long, leftMiddle As Long, leftDown As Long, curEnergy As Long
    
    Dim lastProcessedRow As Long
    
    'Repeat the tracing process for each seam
    For i = 1 To numOfAddedSeams
        
        lastProcessedRow = rowIndex
        
        'rowIndex initially contains the y coordinate of the minimum energy seam terminus.
        ' Starting there, traverse the image leftward, marking lowest-energy values as we go
        ' in the visibility map.
        For x = finalX To 0 Step -1
        
            visibilityMap(x, rowIndex) = visibilityMap(x, rowIndex) + 1
            
            'Artificially increase the energy of this position, to reduce the chances of it being selected again in the future
            seamMap(x, rowIndex) = seamMap(x, rowIndex) + 64
            
            'Find the minimum value of the next row left.
            If (x > 0) Then
            
                xCheck = x - 1
                leftMiddle = seamMap(xCheck, rowIndex)
            
                'Note that we must check edge pixels differently; hence the nested IF statements here
                If (rowIndex > 0) Then
            
                    leftUp = seamMap(xCheck, rowIndex - 1)
                    If (rowIndex < finalY) Then
                    
                        leftDown = seamMap(xCheck, rowIndex + 1)
                    
                        'This is not a top or bottom edge pixel.  Check all three pixels leftward for a minimum value.
                        If (leftUp < leftMiddle) Then
                            If (leftUp < leftDown) Then rowIndex = rowIndex - 1 Else rowIndex = rowIndex + 1
                        Else
                            If (leftMiddle > leftDown) Then rowIndex = rowIndex + 1
                        End If
                    
                    Else
                    
                        'This is a bottom edge pixel.  Check only two pixels left.
                        If (leftUp < leftMiddle) Then rowIndex = rowIndex - 1
                        
                    End If
                
                'This is a top edge pixel.  Check only two pixels leftward.
                Else
                
                    leftDown = seamMap(xCheck, rowIndex + 1)
                    If (leftMiddle > leftDown) Then rowIndex = rowIndex + 1
                    
                End If
                
            End If
        
        Next x
        
        'Mark the initial pixel for this seam as having max energy, which will prevent it from being re-selected by future seam analyses
        seamMap(finalX, lastProcessedRow) = maxVal + 1
        
        'Find the next-smallest column, which will be the source of our next pixel insertion
        minVal = LONG_MAX
        For y = 0 To finalY
            curEnergy = seamMap(finalX, y)
            If (curEnergy < minVal) Then
                minVal = curEnergy
                rowIndex = y
            End If
        Next y
        
        If (i And 8) = 0 Then SetProgBarVal progBarOffset + i
    
    Next i
    
    'We now have a completed visibility map for this image.  Next, we will use that visibility map to construct a new destination image.
    
    'Create a blank destination image
    Dim tmpImage As pdDIB
    Set tmpImage = New pdDIB
    tmpImage.CreateBlank finalX + 1, finalY + 1 + numOfAddedSeams, m_SourceImage.GetDIBColorDepth
    
    'Obtain a pointer to the raw DIB bits of both the source and destination images
    Dim srcImageData() As Long, srcSA As SafeArray2D
    m_SourceImage.WrapLongArrayAroundDIB srcImageData, srcSA
    
    Dim dstImageData() As Long, dstSA As SafeArray2D
    tmpImage.WrapLongArrayAroundDIB dstImageData, dstSA
    
    'Note that if the image is going to be resized again in the future, we also need to update the energy map
    Dim newEnergyMap() As Byte
    ReDim newEnergyMap(0 To finalX, 0 To finalY + numOfAddedSeams) As Byte
        
    Dim quickY As Long, yOffset As Long
    
    For x = 0 To finalX
    
        'Reset the offset for each line
        yOffset = 0
    
        'Loop through each pixel in the SOURCE image
        For y = 0 To finalY
        
            quickY = y + yOffset
            
            'Copy pixels from the source image to the destination image, using the visibility map
            ' to duplicate pixels as necessary.  (Also update the energy map as we go.)
            For i = 0 To visibilityMap(x, y)
                dstImageData(x, quickY + i) = srcImageData(x, y)
                newEnergyMap(x, quickY + i) = m_Energy(x, y)
            Next i
            
            'Permanently increase the destination offset by the visibility map value at this position
            yOffset = yOffset + visibilityMap(x, y)
        
        Next y
    
    Next x
    
    'Release our image pointers
    m_SourceImage.UnwrapLongArrayFromDIB srcImageData
    tmpImage.UnwrapLongArrayFromDIB dstImageData
    
    'Replace the source image with the newly generated destination image
    m_SourceImage.CreateFromExistingDIB tmpImage
    Set tmpImage = Nothing
    
    'Replace the original energy array with the resized copy
    ReDim m_Energy(0 To finalX, 0 To finalY + numOfAddedSeams) As Byte
    CopyMemoryStrict VarPtr(m_Energy(0, 0)), VarPtr(newEnergyMap(0, 0)), (finalX + 1) * (finalY + 1 + numOfAddedSeams)
    
End Sub

'Generate a horizontal seam map for the current energy image.  Because the image may have already been
' shrunk in one (or more) directions by previous seam map operations, a final X and Y value are also
' explicitly supplied; we use these values instead of the bounds of the seamMap() array (which generally
' *won't* be resized since large allocations are expensive).
'
'This function returns several things via ByRef parameters:
' - maxVal:
Private Sub GetHorizontalSeamMap(ByRef seamMap() As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef maxVal As Long, ByRef minVal As Long, ByRef minColumn As Long, Optional ByVal scanDirectionNormal As Boolean = True)
    
    Dim x As Long, y As Long, yCheck As Long
    
    'Populate the first row of the energy array by simply copying over the relevant energy values
    For x = 0 To finalX
        seamMap(x, 0) = m_Energy(x, 0)
    Next x
    
    Dim topLeft As Long, topMiddle As Long, topRight As Long, curEnergy As Long
    
    'Now we can start traversing the energy array.  At each point, generate a new energy for the pixel using
    ' the smallest energy value above said pixel
    For y = 1 To finalY
        yCheck = y - 1
    For x = 0 To finalX
        
        topMiddle = seamMap(x, yCheck)
        curEnergy = m_Energy(x, y)
        
        'Note that we must check edge pixels differently; hence the nested IF statements here
        If (x > 0) Then
        
            topLeft = seamMap(x - 1, yCheck)
            If (x < finalX) Then
            
                topRight = seamMap(x + 1, yCheck)
            
                'This is not a left or right edge pixel.  Check all three pixels above for a minimum value.
                If (topLeft < topMiddle) Then
                    If (topLeft < topRight) Then seamMap(x, y) = curEnergy + topLeft Else seamMap(x, y) = curEnergy + topRight
                Else
                    If (topMiddle < topRight) Then seamMap(x, y) = curEnergy + topMiddle Else seamMap(x, y) = curEnergy + topRight
                End If
            
            Else
            
                'This is a right edge pixel.  Check only two pixels above.
                If (topLeft < topMiddle) Then seamMap(x, y) = curEnergy + topLeft Else seamMap(x, y) = curEnergy + topMiddle
                
            End If
        
        'This is a left edge pixel.  Check only two pixels above.
        Else
            topRight = seamMap(x + 1, yCheck)
            If (topMiddle < topRight) Then seamMap(x, y) = curEnergy + topMiddle Else seamMap(x, y) = curEnergy + topRight
        End If
        
    Next x
    Next y
    
    'The seamMap array now contains a cumulative energy map for the image, which we can reverse-track to
    ' find out which seams should be removed!
    
    'As a convenience to subsequent functions, this function also returns the maximum value of the seam map.
    ' Processed pixels can be set to this value, which prevents them from being re-selected on subsequent runs.
    maxVal = 0
    minVal = LONG_MAX
    
    Dim scanStart As Long, scanEnd As Long, scanStep As Long
    
    If scanDirectionNormal Then
        scanStart = 0
        scanEnd = finalX
        scanStep = 1
    Else
        scanStart = finalX
        scanEnd = 0
        scanStep = -1
    End If
    
    For x = scanStart To scanEnd Step scanStep
    
        curEnergy = seamMap(x, finalY)
        If (curEnergy > maxVal) Then maxVal = curEnergy
        
        If (curEnergy < minVal) Then
            minVal = curEnergy
            minColumn = x
        End If
        
    Next x
    
End Sub

'Given a valid seam map, remove one vertical seam (which will result in a 1px smaller horizontal image)
Private Sub ShrinkHorizontally(ByRef seamMap() As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef maxVal As Long, ByRef minVal As Long, ByRef columnIndex As Long)

    'Start by finding the smallest energy value in the final row of the seam map
    Dim y As Long
    
    Dim yCheck As Long
    Dim topLeft As Long, topMiddle As Long, topRight As Long
    
    'As part of shrinking the image, we will need to shift pixels right-to-left.  RtlMoveMemory performance
    ' falls off a cliff when performing overlapping writes, so even though it's counterintuitive,
    ' we actually get a nice boost by copying memory to a dedicated temp buffer, then copying it from
    ' there into its overlapped position.
    
    '(Note that we do this for both the energy map, and actual pixels.)
    Dim overlapFix() As Byte, pxOverlapFix() As Byte, copySize As Long
    ReDim overlapFix(0 To finalX) As Byte
    ReDim pxOverlapFix(0 To finalX * 4 + 3) As Byte
    
    Dim srcPixels() As Long, srcSA As SafeArray2D
    m_SourceImage.WrapLongArrayAroundDIB srcPixels, srcSA
    
    'ColumnIndex now contains the x coordinate of the minimum energy seam terminus.  Starting there, traverse the
    ' image upward, removing lowest-energy values as we go (and shifting all data past that pixel to the left).
    For y = finalY To 0 Step -1
    
        'Remove the minimum value from the energy map and shift all corresponding data left.
        If (columnIndex < finalX) Then
            
            'Both the energy image and source image must have their data shifted.
            ' (Note that we use dedicated temporary buffers to reduce perf penalties
            '  caused by overlapping memcpy)
            copySize = (finalX - columnIndex) * 4
            CopyMemoryStrict VarPtr(pxOverlapFix(0)), VarPtr(srcPixels(columnIndex + 1, y)), copySize
            CopyMemoryStrict VarPtr(srcPixels(columnIndex, y)), VarPtr(pxOverlapFix(0)), copySize
            'GDI.BitBltWrapper m_SourceImage.GetDIBDC, columnIndex, y, finalX - columnIndex, 1, m_SourceImage.GetDIBDC, columnIndex + 1, y, vbSrcCopy
            
            copySize = finalX - columnIndex
            CopyMemoryStrict VarPtr(overlapFix(0)), VarPtr(m_Energy(columnIndex + 1, y)), copySize
            CopyMemoryStrict VarPtr(m_Energy(columnIndex, y)), VarPtr(overlapFix(0)), copySize
            
        End If
        
        'Find the minimum value of the next row up.
        If (y > 0) Then
        
            yCheck = y - 1
            topMiddle = seamMap(columnIndex, yCheck)
        
            'Note that we must check edge pixels differently; hence the nested IF statements here
            If (columnIndex > 0) Then
        
                topLeft = seamMap(columnIndex - 1, yCheck)
                If (columnIndex < finalX) Then
                
                    topRight = seamMap(columnIndex + 1, yCheck)
                
                    'This is not a left or right edge pixel.  Check all three pixels above for a minimum value.
                    If (topLeft < topMiddle) Then
                        If (topLeft < topRight) Then columnIndex = columnIndex - 1 Else columnIndex = columnIndex + 1
                    Else
                        If (topMiddle > topRight) Then columnIndex = columnIndex + 1
                    End If
                
                Else
                
                    'This is a right edge pixel.  Check only two pixels above.
                    If (topLeft < topMiddle) Then columnIndex = columnIndex - 1
                    
                End If
            
            'This is a left edge pixel.  Check only two pixels above.
            Else
                topRight = seamMap(columnIndex + 1, yCheck)
                If (topMiddle > topRight) Then columnIndex = columnIndex + 1
            End If
            
        End If
    
    Next y
    
    m_SourceImage.UnwrapLongArrayFromDIB srcPixels

End Sub

'Given a valid seam map, add n vertical seams (which will result in an n-px larger horizontal image)
Private Sub GrowHorizontally(ByRef seamMap() As Long, ByVal numOfAddedSeams As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef maxVal As Long, ByRef minVal As Long, ByRef columnIndex As Long)

    'Due to limitations of the seam carving algorithm, this function can only add
    ' [imagewidth \ 2] seams at a time.  If the user requires more than that,
    ' call this function again.
    If (numOfAddedSeams > m_SourceImage.GetDIBWidth \ 2) Then numOfAddedSeams = m_SourceImage.GetDIBWidth \ 2
    
    Dim x As Long, y As Long, i As Long
    
    'Growing seams is a bit of a different approach from removing seams.  Because added seams will always have minimal energy,
    ' we must vary the order in which seams are added.  To do this, we generate a "visibility map".  This map (a 2D array the size
    ' of the image) will store how many times each source pixel is duplicated in the destination image.  Eventually it would be nice
    ' to add some interpolation capabilities to this function, but for now, we only do blind pixel insertion.
    Dim visibilityMap() As Long
    ReDim visibilityMap(0 To finalX, 0 To finalY) As Long
    
    Dim yCheck As Long
    Dim topLeft As Long, topMiddle As Long, topRight As Long, curEnergy As Long
    
    Dim lastProcessedColumn As Long
    
    'Repeat the tracing process for each seam
    For i = 1 To numOfAddedSeams
        
        lastProcessedColumn = columnIndex
        
        'columnIndex initially contains the x coordinate of the minimum energy seam terminus.  Starting there, traverse the
        ' image upward, marking lowest-energy values as we go in the visibility map.
        For y = finalY To 0 Step -1
        
            visibilityMap(columnIndex, y) = visibilityMap(columnIndex, y) + 1
            
            'Artificially increase the energy of this position, to reduce the chances of it being selected again in the future
            seamMap(columnIndex, y) = seamMap(columnIndex, y) + 64
            
            'Find the minimum value of the next row up.
            If (y > 0) Then
            
                yCheck = y - 1
                topMiddle = seamMap(columnIndex, yCheck)
            
                'Note that we must check edge pixels differently; hence the nested IF statements here
                If (columnIndex > 0) Then
            
                    topLeft = seamMap(columnIndex - 1, yCheck)
                    If (columnIndex < finalX) Then
                    
                        topRight = seamMap(columnIndex + 1, yCheck)
                    
                        'This is not a left or right edge pixel.  Check all three pixels above for a minimum value.
                        If (topLeft < topMiddle) Then
                            If (topLeft < topRight) Then columnIndex = columnIndex - 1 Else columnIndex = columnIndex + 1
                        Else
                            If (topMiddle > topRight) Then columnIndex = columnIndex + 1
                        End If
                    
                    'This is a right edge pixel.  Check only two pixels above.
                    Else
                        If (topLeft < topMiddle) Then columnIndex = columnIndex - 1
                    End If
                
                'This is a left edge pixel.  Check only two pixels above.
                Else
                    topRight = seamMap(columnIndex + 1, yCheck)
                    If (topMiddle > topRight) Then columnIndex = columnIndex + 1
                End If
                
            End If
        
        Next y
        
        'Mark the initial pixel for this seam as having max energy, which will prevent it from being re-selected by future seam analyses
        seamMap(lastProcessedColumn, finalY) = maxVal + 1
        
        'Find the next-smallest row, which will be the source of our next pixel insertion
        minVal = LONG_MAX
        For x = 0 To finalX
            curEnergy = seamMap(x, finalY)
            If (curEnergy < minVal) Then
                minVal = curEnergy
                columnIndex = x
            End If
        Next x
        
        If (i And 8) = 0 Then SetProgBarVal i
    
    Next i
    
    'We now have a completed visibility map for this image.
    
    'Next, we will use the visibility map to construct a new destination image.
    
    'Create a blank destination image
    Dim tmpImage As pdDIB
    Set tmpImage = New pdDIB
    tmpImage.CreateBlank finalX + 1 + numOfAddedSeams, finalY + 1, m_SourceImage.GetDIBColorDepth
    
    'Obtain a pointer to the raw DIB bits of both the source and destination images
    Dim srcImageData() As Long, srcSA As SafeArray2D
    m_SourceImage.WrapLongArrayAroundDIB srcImageData, srcSA
    
    Dim dstImageData() As Long, dstSA As SafeArray2D
    tmpImage.WrapLongArrayAroundDIB dstImageData, dstSA
    
    'Note that if the image is going to be resized again in the future, we also need to update the energy map
    Dim newEnergyMap() As Byte
    ReDim newEnergyMap(0 To finalX + numOfAddedSeams, 0 To finalY) As Byte
    
    Dim quickSrcX As Long, quickDstX As Long, xOffset As Long, innerXOffset As Long
    
    For y = 0 To finalY
    
        'Reset the offset for each line
        xOffset = 0
    
        'Loop through each pixel in the SOURCE image
        For x = 0 To finalX
        
            quickSrcX = x
            quickDstX = (x + xOffset)
            
            'Copy pixels from the source image to the destination image, using the visibility map to duplicate pixels as necessary
            For i = 0 To visibilityMap(x, y)
            
                innerXOffset = quickDstX + i
                
                'Remap RGB data to the destination image
                dstImageData(innerXOffset, y) = srcImageData(quickSrcX, y)
                
                'Copy energy data too
                newEnergyMap(x + xOffset, y) = m_Energy(x, y)
                
            Next i
            
            'Permanently increase the destination offset by the visibility map value at this position
            xOffset = xOffset + visibilityMap(x, y)
        
        Next x
    
    Next y
    
    'Release our image pointers
    m_SourceImage.UnwrapLongArrayFromDIB srcImageData
    tmpImage.UnwrapLongArrayFromDIB dstImageData
    
    'Replace the source image with the newly generated destination image
    m_SourceImage.CreateFromExistingDIB tmpImage
    Set tmpImage = Nothing
    
    'Replace the original energy array with the resized copy
    ReDim m_Energy(0 To finalX + numOfAddedSeams, 0 To finalY) As Byte
    CopyMemoryStrict VarPtr(m_Energy(0, 0)), VarPtr(newEnergyMap(0, 0)), (finalX + 1 + numOfAddedSeams) * (finalY + 1)

End Sub

'Though they share some similar elements, shrinking and enlarging an image in either direction are handled separately.
Friend Function StartSeamCarve(ByVal newWidth As Long, ByVal newHeight As Long) As Boolean

    'We will be referencing image size frequently in this function, so cache these values in advance
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = m_SourceImage.GetDIBWidth
    srcHeight = m_SourceImage.GetDIBHeight
    
    'Start by determining if we are widening or shrinking the image in the horizontal and/or vertical directions
    ' We'll use a simple system: -1 means shrinking, 0 means no change, and 1 means enlarging.
    Dim horizontalStatus As PD_SeamCarveScaling, verticalStatus As PD_SeamCarveScaling
    If (newWidth < srcWidth) Then
        horizontalStatus = scs_Shrink
    ElseIf (newWidth = srcWidth) Then
        horizontalStatus = scs_None
    Else
        horizontalStatus = scs_Grow
    End If
    
    If (newHeight < srcHeight) Then
        verticalStatus = scs_Shrink
    ElseIf (newHeight = srcHeight) Then
        verticalStatus = scs_None
    Else
        verticalStatus = scs_Grow
    End If
    
    'We must generate at least one (and possibly more) seam constructs using the image energy as our guide
    Dim seamData() As Long, seamMax As Long, seamMin As Long, seamIndex As Long
    ReDim seamData(0 To PDMath.Max2Int(newWidth, srcWidth), 0 To PDMath.Max2Int(newHeight, srcHeight)) As Long
    
    Dim i As Long
    
    'Because we need to transfer image data back and forth between DIBs, a temporary DIB comes in handy
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateFromExistingDIB m_SourceImage
    
    'Determine a progress bar maximum value.  We'll refresh the progress bar approximately once for each
    ' seam added or removed from the image.
    Dim hDiff As Long, vDiff As Long
    hDiff = Abs(newWidth - srcWidth)
    vDiff = Abs(newHeight - srcHeight)
    SetProgBarMax vDiff + hDiff
    
    'Start with the horizontal direction
    Select Case horizontalStatus
    
        'Horizontal size isn't changing - do nothing!
        Case scs_None
        
        'Horizontal size is shrinking
        Case scs_Shrink
            
            'We now have a "seam map" (e.g. energy vector representation) of the initial image.  Use that move to remove
            ' however many pixels are required to reach the newly specified size.
            For i = 1 To srcWidth - newWidth
            
                'Retrieve an updated horizontal seam map of the image
                GetHorizontalSeamMap seamData, srcWidth - i, srcHeight - 1, seamMax, seamMin, seamIndex, (i And 1) = 0
                
                'Shrink the image by a single pixel in the horizontal direction
                ShrinkHorizontally seamData, srcWidth - i, srcHeight - 1, seamMax, seamMin, seamIndex
                
                If (i And 8) = 0 Then
                    SetProgBarVal i
                    If Interface.UserPressedESC() Then Exit For
                End If
                
            Next i
            
            'Check for user cancellation (by pressing ESC)
            If g_cancelCurrentAction Then Exit Function
            
            'All necessary horizontal seams have been removed from the image.  Overwrite the original source image with
            ' the new image data.
            tmpDIB.CreateBlank newWidth, srcHeight, m_SourceImage.GetDIBColorDepth
            GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, newWidth, srcHeight, m_SourceImage.GetDIBDC, 0, 0, vbSrcCopy
            m_SourceImage.CreateFromExistingDIB tmpDIB
            
            'Note the new width now, in case the user is also changing the vertical size of the image
            srcWidth = m_SourceImage.GetDIBWidth
    
        'Horizontal size is growing
        Case scs_Grow
        
            Do
            
                'Retrieve an updated horizontal seam map of the image
                GetHorizontalSeamMap seamData, srcWidth - 1, srcHeight - 1, seamMax, seamMin, seamIndex, (i And 1) = 0
                
                'Grow the image horizontally, using the newly generated seam map as our guide
                GrowHorizontally seamData, newWidth - srcWidth, srcWidth - 1, srcHeight - 1, seamMax, seamMin, seamIndex
                
                'Update the source width, in case this function must be called again
                srcWidth = m_SourceImage.GetDIBWidth
                
                If Interface.UserPressedESC() Then Exit Do
                
            Loop While m_SourceImage.GetDIBWidth < newWidth
            
            'Update the temp image (which may be used by subsequent vertical resize functions)
            tmpDIB.CreateFromExistingDIB m_SourceImage
    
    End Select
    
    'Next, process the vertical direction
    Select Case verticalStatus
    
        'Vertical size isn't changing - do nothing!
        Case scs_None
        
        'Vertical size is shrinking
        Case scs_Shrink
            
            'We now have a "seam map" (e.g. energy vector representation) of the initial image.  Use that move to remove
            ' however many pixels are required to reach the newly specified size.
            For i = 1 To srcHeight - newHeight
            
                'Retrieve an updated vertical seam map of the image
                GetVerticalSeamMap seamData, srcWidth - 1, srcHeight - i, seamMax, seamMin, seamIndex, (i And 1) = 0
                
                'Shrink the image by a single pixel in the vertical direction
                ShrinkVertically seamData, srcWidth - 1, srcHeight - i, seamMax, seamMin, seamIndex
                
                If (i And 8) = 0 Then
                    SetProgBarVal hDiff + i
                    If Interface.UserPressedESC() Then Exit For
                End If
                
            Next i
            
            'Check for user cancellation (by pressing ESC)
            If g_cancelCurrentAction Then Exit Function
            
            'All necessary vertical seams have been removed from the image.  Overwrite the original source image with
            ' the new image data.
            tmpDIB.CreateBlank newWidth, newHeight, m_SourceImage.GetDIBColorDepth
            GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, newWidth, newHeight, m_SourceImage.GetDIBDC, 0, 0, vbSrcCopy
            m_SourceImage.CreateFromExistingDIB tmpDIB
            
    
        'Vertical size is growing
        Case scs_Grow
        
            Do
            
                'Retrieve an updated vertical seam map of the image
                GetVerticalSeamMap seamData, srcWidth - 1, srcHeight - 1, seamMax, seamMin, seamIndex, (i And 1) = 0
                
                'Grow the image vertically, using the newly generated seam map as our guide
                GrowVertically seamData, newHeight - srcHeight, srcWidth - 1, srcHeight - 1, seamMax, seamMin, seamIndex, hDiff
                
                'Update the source height, in case this function must be called again
                srcHeight = m_SourceImage.GetDIBHeight
                
                If Interface.UserPressedESC() Then Exit Do
            
            Loop While m_SourceImage.GetDIBHeight < newHeight
            
            'Update the temp image (which will be used to overwrite the original image)
            tmpDIB.CreateFromExistingDIB m_SourceImage
    
    
    End Select
    
    'Now that all seam carves have been carried out, generate a final destination image
    ' (and make sure it is marked with correct premultiplication status)
    Set m_DestinationImage = New pdDIB
    m_DestinationImage.CreateFromExistingDIB tmpDIB
    m_DestinationImage.SetInitialAlphaPremultiplicationState True
    
    'Reset the progress bar
    SetProgBarVal 0
    ReleaseProgressBar
    
End Function
