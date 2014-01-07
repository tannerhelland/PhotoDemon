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
'Copyright ©2013-2014 by Tanner Helland
'Created: 06/January/14
'Last updated: 06/January/14
'Last update: initial build
'
'Content-aware scaling is a very exciting addition to PhotoDemon 6.2.  (As a comparison, PhotoShop didn't gain this
' feature until CS4, so it's pretty cutting-edge stuff!)
'
'Normal scaling algorithms work by shrinking or enlarging all image pixels equally.  Such algorithms make no distinction
' between visually important pixels and visually unimportant ones.  Unfortunately, when the aspect ratio of an image is
' changed using such an algorithm, noticeable distortion will result, and the end result will typically be unpleasant.
'
'Content-aware scaling circumvents this by selectively removing the least visually important parts of an image
' (as determined by some type of per-pixel "energy" calculation).  By preferentially removing uninteresting pixels
' before interesting ones, important parts of an image can be preserved while uninteresting parts are removed.  The
' result is often a much more aesthetically pleasing image, even under severe aspect ratio changes.
'
'For reference, the original 2007 paper that first proposed this technique - called "seam carving" is available here:
' http://www.win.tue.nl/~wstahw/edu/2IV05/seamcarving.pdf
'
'I have written PhotoDemon's implementation from scratch, using the original paper as my primary resource.  Unfortunately,
' my current implementation is quite slow (though still faster than many other implementations!) on account of all the
' seam finding operations that must be performed.  I am investigating ways to further improve the algorithm's performance,
' but I remain worried that this task may prove a bit much for VB6.  We'll see.
'
'Image *enlarging* remains to be implemented!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Source image, which will be carved into smaller and smaller bits by the function
Private m_SourceImage As pdLayer

'Energy image, which can be generated however the calling function wants
Private m_EnergyImage As pdLayer

'Final image, which will only exist after a seam carve operation has been carried out
Private m_DestinationImage As pdLayer

'Shrinking and enlarging are handled separately
Private Enum ScaleChange
    ScaleSmaller = 0
    ScaleNone = 1
    ScaleLarger = 2
End Enum

#If False Then
    Private Const ScaleSmaller = 0, ScaleNone = 1, ScaleLarger = 2
#End If

'For performance reasons, the seam carver class maintains its own copy of the source image.
Public Sub setSourceImage(ByRef srcLayer As pdLayer)
    Set m_SourceImage = New pdLayer
    m_SourceImage.createFromExistingLayer srcLayer
End Sub

'Retrieve the carved image into a layer of the calling function's choosing
Public Function getCarvedImage() As pdLayer
    Set getCarvedImage = m_DestinationImage
End Function

Public Sub setEnergyImage(ByRef srcLayer As pdLayer)
    
    'Create a copy of the energy image
    Set m_EnergyImage = New pdLayer
    m_EnergyImage.createFromExistingLayer srcLayer
    
    'Transparency data is not relevant for computing pixel energy, so remove it if present
    If m_EnergyImage.getLayerColorDepth = 32 Then m_EnergyImage.convertTo24bpp
    
    'To speed up processing, we are going to store all energy data in the blue channel of the image.
    Dim energyData() As Byte
    Dim energySA As SAFEARRAY2D
    prepSafeArray energySA, m_EnergyImage
    CopyMemory ByVal VarPtrArray(energyData()), VarPtr(energySA), 4
    
    Dim x As Long, y As Long, QuickVal As Long
    Dim r As Long, g As Long, b As Long
    
    Dim grayLookup(0 To 765) As Byte
    
    For x = 0 To 765
        grayLookup(x) = x \ 3
    Next x
    
    For x = 0 To srcLayer.getLayerWidth - 1
        QuickVal = x * 3
    For y = 0 To srcLayer.getLayerHeight - 1
        r = energyData(QuickVal + 2, y)
        g = energyData(QuickVal + 1, y)
        b = energyData(QuickVal, y)
        energyData(QuickVal, y) = grayLookup(r + g + b)
    Next y
    Next x
    
    CopyMemory ByVal VarPtrArray(energyData), 0&, 4
    
End Sub

'Generate a vertical seam map for the current energy image.  Because those images may have been shrunk in one (or more)
' directions by previous seam map operations, a final X and Y value are also explicitly supplied, so the array can be
' re-scanned while ignoring irrelevant pixels.
Public Sub getVerticalSeamMap(ByRef seamMap() As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef maxVal As Long)

    'Resize the seam map to the requested dimensions
    ReDim seamMap(0 To finalX, 0 To finalY) As Long
    
    'Obtain a pointer to the energy image's DIB data
    Dim energyData() As Byte
    Dim energySA As SAFEARRAY2D
    prepSafeArray energySA, m_EnergyImage
    CopyMemory ByVal VarPtrArray(energyData()), VarPtr(energySA), 4
    
    Dim x As Long, y As Long, xCheck As Long
    Dim r As Long, g As Long, b As Long
    
    'Populate the first row of the energy array by simply copying over the relevant energy values
    For y = 0 To finalY
        seamMap(0, y) = energyData(0, y)
    Next y
    
    Dim leftUp As Long, leftMiddle As Long, leftDown As Long, curEnergy As Long, QuickX As Long
    
    'Now we can start traversing the energy array.  At each point, generate a new energy for the pixel using
    ' the smallest energy value above said pixel
    For x = 1 To finalX
        xCheck = (x - 1)
        QuickX = x * 3
    For y = 0 To finalY
        
        leftMiddle = seamMap(xCheck, y)
        curEnergy = energyData(QuickX, y)
        
        'Note that we must check edge pixels differently; hence the nested IF statements here
        If y > 0 Then
        
            leftUp = seamMap(xCheck, y - 1)
        
            If y < finalY Then
            
                leftDown = seamMap(xCheck, y + 1)
            
                'This is not a left or right edge pixel.  Check all three pixels above for a minimum value.
                If leftUp < leftMiddle Then
                    If leftUp < leftDown Then
                        seamMap(x, y) = curEnergy + leftUp
                    Else
                        seamMap(x, y) = curEnergy + leftDown
                    End If
                Else
                    If leftMiddle < leftDown Then
                        seamMap(x, y) = curEnergy + leftMiddle
                    Else
                        seamMap(x, y) = curEnergy + leftDown
                    End If
                End If
            
            Else
            
                'This is a right edge pixel.  Check only two pixels above.
                If leftUp < leftMiddle Then
                    seamMap(x, y) = curEnergy + leftUp
                Else
                    seamMap(x, y) = curEnergy + leftMiddle
                End If
            
            End If
        
        'This is a left edge pixel.  Check only two pixels above.
        Else
        
            leftDown = seamMap(xCheck, y + 1)
        
            If leftMiddle < leftDown Then
                seamMap(x, y) = curEnergy + leftMiddle
            Else
                seamMap(x, y) = curEnergy + leftDown
            End If
        
        End If
        
    Next y
    Next x
    
    'The seamMap array now contains a cumulative energy map for the image, which we can reverse-track to
    ' find out which seams should be removed!
    
    'As a convenience to subsequent functions, this function also returns the maximum value of the seam map.
    ' Processed pixels can be set to this value, which prevents them from being re-selected on subsequent runs.
    maxVal = 0
    For y = 0 To finalY
        If seamMap(finalX, y) > maxVal Then maxVal = seamMap(finalX, y)
    Next y
    
    'Release our energy data pointer
    CopyMemory ByVal VarPtrArray(energyData), 0&, 4
    
End Sub

'Given a valid seam map, remove one horizontal seam (which will result in a 1px smaller vertical image)
Private Function shrinkVertically(ByRef seamMap() As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef maxVal As Long)

    'Start by finding the smallest energy value in the final column of the seam map
    Dim x As Long, y As Long
    Dim rowIndex As Long, minVal As Long
    rowIndex = 0
    minVal = maxVal
    
    For y = 0 To finalY
        If seamMap(finalX, y) < minVal Then
            minVal = seamMap(finalX, y)
            rowIndex = y
        End If
    Next y
    
    Dim xCheck As Long
    Dim leftUp As Long, leftMiddle As Long, leftDown As Long
    
    'rowIndex now contains the y coordinate of the minimum energy seam terminus.  Starting there, traverse the
    ' image leftward, removing lowest-energy values as we go (and shifting all data past that pixel upward).
    For x = finalX To 0 Step -1
    
        'Remove the minimum value from the energy map and shift all corresponding data up.
        If rowIndex < finalY Then
            
            'Both the energy image and source image must have their data shifted
            BitBlt m_EnergyImage.getLayerDC, x, rowIndex, 1, finalY - rowIndex, m_EnergyImage.getLayerDC, x, rowIndex + 1, vbSrcCopy
            BitBlt m_SourceImage.getLayerDC, x, rowIndex, 1, finalY - rowIndex, m_SourceImage.getLayerDC, x, rowIndex + 1, vbSrcCopy
            
        End If
        
        'Find the minimum value of the next row left.
        If x > 0 Then
        
            xCheck = x - 1
            leftMiddle = seamMap(xCheck, rowIndex)
        
            'Note that we must check edge pixels differently; hence the nested IF statements here
            If rowIndex > 0 Then
        
                leftUp = seamMap(xCheck, rowIndex - 1)
        
                If rowIndex < finalY Then
                
                    leftDown = seamMap(xCheck, rowIndex + 1)
                
                    'This is not a left or right edge pixel.  Check all three pixels above for a minimum value.
                    If leftUp < leftMiddle Then
                        If leftUp < leftDown Then
                            rowIndex = rowIndex - 1
                        Else
                            rowIndex = rowIndex + 1
                        End If
                    Else
                        If leftMiddle > leftDown Then
                            rowIndex = rowIndex + 1
                        'Else
                            'rowIndex = rowIndex
                        End If
                    End If
                
                Else
                
                    'This is a right edge pixel.  Check only two pixels above.
                    If leftUp < leftMiddle Then
                        rowIndex = rowIndex - 1
                    'Else
                        'rowIndex = rowIndex
                    End If
                
                End If
            
            'This is a left edge pixel.  Check only two pixels above.
            Else
            
                leftDown = seamMap(xCheck, rowIndex + 1)
            
                If leftMiddle > leftDown Then
                    rowIndex = rowIndex + 1
                'Else
                    'rowIndex = rowIndex
                End If
            
            End If
            
        End If
    
    Next x

End Function

'Generate a horizontal seam map for the current energy image.  Because those images may have been shrunk in one (or more)
' directions by previous seam map operations, a final X and Y value are also explicitly supplied, so the array can be
' re-scanned while ignoring irrelevant pixels.
Public Sub getHorizontalSeamMap(ByRef seamMap() As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef maxVal As Long)

    'Resize the seam map to the requested dimensions
    ReDim seamMap(0 To finalX, 0 To finalY) As Long
    
    'Obtain a pointer to the energy image's DIB data
    Dim energyData() As Byte
    Dim energySA As SAFEARRAY2D
    prepSafeArray energySA, m_EnergyImage
    CopyMemory ByVal VarPtrArray(energyData()), VarPtr(energySA), 4
    
    Dim x As Long, y As Long, yCheck As Long
    Dim r As Long, g As Long, b As Long
    
    'Populate the first row of the energy array by simply copying over the relevant energy values
    For x = 0 To finalX
        seamMap(x, 0) = energyData(x * 3, 0)
    Next x
    
    Dim topLeft As Long, topMiddle As Long, topRight As Long, curEnergy As Long
    
    'Now we can start traversing the energy array.  At each point, generate a new energy for the pixel using
    ' the smallest energy value above said pixel
    For y = 1 To finalY
        yCheck = y - 1
    For x = 0 To finalX
        
        topMiddle = seamMap(x, yCheck)
        curEnergy = energyData(x * 3, y)
        
        'Note that we must check edge pixels differently; hence the nested IF statements here
        If x > 0 Then
        
            topLeft = seamMap(x - 1, yCheck)
        
            If x < finalX Then
            
                topRight = seamMap(x + 1, yCheck)
            
                'This is not a left or right edge pixel.  Check all three pixels above for a minimum value.
                If topLeft < topMiddle Then
                    If topLeft < topRight Then
                        seamMap(x, y) = curEnergy + topLeft
                    Else
                        seamMap(x, y) = curEnergy + topRight
                    End If
                Else
                    If topMiddle < topRight Then
                        seamMap(x, y) = curEnergy + topMiddle
                    Else
                        seamMap(x, y) = curEnergy + topRight
                    End If
                End If
            
            Else
            
                'This is a right edge pixel.  Check only two pixels above.
                If topLeft < topMiddle Then
                    seamMap(x, y) = curEnergy + topLeft
                Else
                    seamMap(x, y) = curEnergy + topMiddle
                End If
            
            End If
        
        'This is a left edge pixel.  Check only two pixels above.
        Else
        
            topRight = seamMap(x + 1, yCheck)
        
            If topMiddle < topRight Then
                seamMap(x, y) = curEnergy + topMiddle
            Else
                seamMap(x, y) = curEnergy + topRight
            End If
        
        End If
        
    Next x
    Next y
    
    'The seamMap array now contains a cumulative energy map for the image, which we can reverse-track to
    ' find out which seams should be removed!
    
    'As a convenience to subsequent functions, this function also returns the maximum value of the seam map.
    ' Processed pixels can be set to this value, which prevents them from being re-selected on subsequent runs.
    maxVal = 0
    For x = 0 To finalX
        If seamMap(x, finalY) > maxVal Then maxVal = seamMap(x, finalY)
    Next x
    
    'Release our energy data pointer
    CopyMemory ByVal VarPtrArray(energyData), 0&, 4
    
End Sub

'Given a valid seam map, remove one vertical seam (which will result in a 1px smaller horizontal image)
Private Function shrinkHorizontally(ByRef seamMap() As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef maxVal As Long)

    'Start by finding the smallest energy value in the final row of the seam map
    Dim x As Long, y As Long
    Dim columnIndex As Long, minVal As Long
    columnIndex = 0
    minVal = maxVal
    
    For x = 0 To finalX
        If seamMap(x, finalY) < minVal Then
            minVal = seamMap(x, finalY)
            columnIndex = x
        End If
    Next x
    
    Dim yCheck As Long
    Dim topLeft As Long, topMiddle As Long, topRight As Long
    
    'ColumnIndex now contains the x coordinate of the minimum energy seam terminus.  Starting there, traverse the
    ' image upward, removing lowest-energy values as we go (and shifting all data past that pixel to the left).
    For y = finalY To 0 Step -1
    
        'Remove the minimum value from the energy map and shift all corresponding data left.
        If columnIndex < finalX Then
            
            'Both the energy image and source image must have their data shifted
            BitBlt m_EnergyImage.getLayerDC, columnIndex, y, finalX - columnIndex, 1, m_EnergyImage.getLayerDC, columnIndex + 1, y, vbSrcCopy
            BitBlt m_SourceImage.getLayerDC, columnIndex, y, finalX - columnIndex, 1, m_SourceImage.getLayerDC, columnIndex + 1, y, vbSrcCopy
            
        End If
        
        'Find the minimum value of the next row up.
        If y > 0 Then
        
            yCheck = y - 1
            topMiddle = seamMap(columnIndex, yCheck)
        
            'Note that we must check edge pixels differently; hence the nested IF statements here
            If columnIndex > 0 Then
        
                topLeft = seamMap(columnIndex - 1, yCheck)
        
                If columnIndex < finalX Then
                
                    topRight = seamMap(columnIndex + 1, yCheck)
                
                    'This is not a left or right edge pixel.  Check all three pixels above for a minimum value.
                    If topLeft < topMiddle Then
                        If topLeft < topRight Then
                            columnIndex = columnIndex - 1
                        Else
                            columnIndex = columnIndex + 1
                        End If
                    Else
                        If topMiddle > topRight Then
                            columnIndex = columnIndex + 1
                        'Else
                            'columnIndex = columnIndex
                        End If
                    End If
                
                Else
                
                    'This is a right edge pixel.  Check only two pixels above.
                    If topLeft < topMiddle Then
                        columnIndex = columnIndex - 1
                    'Else
                        'columnIndex = columnIndex
                    End If
                
                End If
            
            'This is a left edge pixel.  Check only two pixels above.
            Else
            
                topRight = seamMap(columnIndex + 1, yCheck)
            
                If topMiddle > topRight Then
                    columnIndex = columnIndex + 1
                'Else
                    'columnIndex = columnIndex
                End If
            
            End If
            
        End If
    
    Next y

End Function

'Though they share some similar elements, shrinking and enlarging an image in either direction are handled separately.
Public Function startSeamCarve(ByVal newWidth As Long, ByVal newHeight As Long) As Boolean

    'We will be referencing image size frequently in this function, so cache these values in advance
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = m_SourceImage.getLayerWidth
    srcHeight = m_SourceImage.getLayerHeight
    
    'Start by determining if we are widening or shrinking the image in the horizontal and/or vertical directions
    ' We'll use a simple system: -1 means shrinking, 0 means no change, and 1 means enlarging.
    Dim horizontalStatus As ScaleChange, verticalStatus As ScaleChange
    If newWidth < srcWidth Then
        horizontalStatus = ScaleSmaller
    ElseIf newWidth = srcWidth Then
        horizontalStatus = ScaleNone
    Else
        horizontalStatus = ScaleLarger
    End If
    
    If newHeight < srcHeight Then
        verticalStatus = ScaleSmaller
    ElseIf newHeight = srcHeight Then
        verticalStatus = ScaleNone
    Else
        verticalStatus = ScaleLarger
    End If
    
    'We must generate at least one (and possibly more) seam constructs using the image energy as our guide
    Dim seamData() As Long, seamMax As Long
    
    Dim i As Long
    
    'Because we need to transfer image data back and forth between DIBs, a temporary layer comes in handy
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer m_SourceImage
    
    'Determine a progress bar maximum value.  We'll refresh the progress bar approximately once for each
    ' seam added or removed from the image.
    Dim hDiff As Long, vDiff As Long
    hDiff = Abs(newWidth - srcWidth)
    vDiff = Abs(newHeight - srcHeight)
    SetProgBarMax vDiff + hDiff
    
    'Start with the horizontal direction
    Select Case horizontalStatus
    
        'Horizontal size isn't changing - do nothing!
        Case ScaleNone
        
        'Horizontal size is shrinking
        Case ScaleSmaller
            
            'We now have a "seam map" (e.g. energy vector representation) of the initial image.  Use that move to remove
            ' however many pixels are required to reach the newly specified size.
            For i = 1 To srcWidth - newWidth
            
                'Retrieve an updated horizontal seam map of the image
                getHorizontalSeamMap seamData, srcWidth - i, srcHeight - 1, seamMax
                
                'Shrink the image by a single pixel in the horizontal direction
                shrinkHorizontally seamData, srcWidth - i, srcHeight - 1, seamMax
                
                If (i And 8) = 0 Then SetProgBarVal i
                
            Next i
            
            'All necessary horizontal seams have been removed from the image.  Overwrite the original source image with
            ' the new image data.
            tmpLayer.createBlank newWidth, srcHeight, m_SourceImage.getLayerColorDepth
            BitBlt tmpLayer.getLayerDC, 0, 0, newWidth, srcHeight, m_SourceImage.getLayerDC, 0, 0, vbSrcCopy
            m_SourceImage.createFromExistingLayer tmpLayer
            
            'Note the new width now, in case the user is also changing the vertical size of the image
            srcWidth = m_SourceImage.getLayerWidth
    
        'Horizontal size is growing
        Case ScaleLarger
    
    End Select
    
    'Next, process the vertical direction
    Select Case verticalStatus
    
        'Vertical size isn't changing - do nothing!
        Case ScaleNone
        
        'Vertical size is shrinking
        Case ScaleSmaller
            
            'We now have a "seam map" (e.g. energy vector representation) of the initial image.  Use that move to remove
            ' however many pixels are required to reach the newly specified size.
            For i = 1 To srcHeight - newHeight
            
                'Retrieve an updated vertical seam map of the image
                getVerticalSeamMap seamData, srcWidth - 1, srcHeight - i, seamMax
                
                'Shrink the image by a single pixel in the vertical direction
                shrinkVertically seamData, srcWidth - 1, srcHeight - i, seamMax
                
                If (i And 8) = 0 Then SetProgBarVal hDiff + i
                
            Next i
            
            'All necessary vertical seams have been removed from the image.  Overwrite the original source image with
            ' the new image data.
            tmpLayer.createBlank newWidth, newHeight, m_SourceImage.getLayerColorDepth
            BitBlt tmpLayer.getLayerDC, 0, 0, newWidth, newHeight, m_SourceImage.getLayerDC, 0, 0, vbSrcCopy
            m_SourceImage.createFromExistingLayer tmpLayer
            
    
        'Vertical size is growing
        Case ScaleLarger
    
    End Select
    
    
    'Now that all seam carves have been carried out, generate a final destination image
    Set m_DestinationImage = New pdLayer
    m_DestinationImage.createFromExistingLayer tmpLayer
    
    'Reset the progress bar
    SetProgBarVal 0
    
    Exit Function
    
    'TESTING ONLY!!!  Fill the original image with a copy of the seam map!  We can use this to see if we're generating
    ' the map correctly.
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already embossed pixels from screwing up our results for later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, m_SourceImage
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, QuickVal As Long
    
    'Start by finding a maximum energy value
    Dim maxEnergy As Long, curEnergy As Long
    maxEnergy = 0
    For x = 0 To srcWidth - 1
    For y = 0 To srcHeight - 1
        If seamData(x, y) > maxEnergy Then maxEnergy = seamData(x, y)
    Next y
    Next x
    
    For x = 0 To srcWidth - 1
        QuickVal = x * 3
    For y = 0 To srcHeight - 1
        curEnergy = (seamData(x, y) / maxEnergy) * 255
        srcImageData(QuickVal, y) = curEnergy
        srcImageData(QuickVal + 1, y) = curEnergy
        srcImageData(QuickVal + 2, y) = curEnergy
    Next y
    Next x
    
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    
    Set m_DestinationImage = New pdLayer
    m_DestinationImage.createFromExistingLayer m_SourceImage

End Function
