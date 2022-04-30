Attribute VB_Name = "SelectionFilters"
'***************************************************************************
'Selection Tools: Filters
'Copyright 2013-2022 by Tanner Helland
'Created: 21/June/13
'Last updated: 18/February/22
'Last update: minor updates to fix compatibility with new "multiple selections" features
'
'This module should only contain selection filters (e.g. "grow", "border", etc).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Enum PD_SelectionDialog
    pdsd_Grow = 0
    pdsd_Shrink = 1
    pdsd_Border = 2
    pdsd_Feather = 3
    pdsd_Sharpen = 4
End Enum

#If False Then
    Private Const pdsd_Grow = 0, pdsd_Shrink = 1, pdsd_Border = 2, pdsd_Feather = 3, pdsd_Sharpen = 4
#End If

'Present a selection-related dialog box (grow, shrink, feather, etc).  This function will return a msgBoxResult value so
' the calling function knows how to proceed, and if the user successfully selected a value, it will be stored in the
' returnValue variable.
Public Function DisplaySelectionDialog(ByVal typeOfDialog As PD_SelectionDialog, ByRef ReturnValue As Double) As VbMsgBoxResult

    Load FormSelectionDialogs
    FormSelectionDialogs.ShowDialog typeOfDialog
    
    DisplaySelectionDialog = FormSelectionDialogs.DialogResult
    ReturnValue = FormSelectionDialogs.paramValue
    
    Unload FormSelectionDialogs
    Set FormSelectionDialogs = Nothing

End Function

'Invert the current selection.  Note that this will make a transformable selection non-transformable - to maintain transformability, use
' the "exterior"/"interior" options on the main form.
' TODO: swap exterior/interior automatically, if a valid option
Public Sub InvertCurrentSelection()

    'Unlock any existing selection, and condense any composite selections down to a single raster layer.
    PDImages.GetActiveImage.MainSelection.LockRelease
    PDImages.GetActiveImage.SetSelectionActive False
    PDImages.GetActiveImage.MainSelection.SquashCompositeToRaster
    
    Message "Inverting..."
    
    'Point a standard 2D byte array at the selection mask
    Dim x As Long, y As Long
    Dim selMaskData() As Long, selMaskSA As SafeArray1D
    
    Dim maskWidth As Long, maskHeight As Long
    maskWidth = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth - 1
    maskHeight = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    SetProgBarMax maskHeight
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'After all that work, the Invert code itself is very small and unexciting!
    For y = 0 To maskHeight
        PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.WrapLongArrayAroundScanline selMaskData, selMaskSA, y
    For x = 0 To maskWidth
        selMaskData(x) = Not selMaskData(x)
    Next x
        If (y And progBarCheck) = 0 Then SetProgBarVal y
    Next y
    
    PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.UnwrapLongArrayFromDIB selMaskData
    
    'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
    ' being non-transformable)
    PDImages.GetActiveImage.MainSelection.SetSelectionShape ss_Raster
    PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
    
    'Apply any final UI changes
    SetProgBarVal 0
    ReleaseProgressBar
    Message "Selection inversion complete."
        
    'Note that if no selections are found, we want to basically perform a "select none" operation.
    ' (This can occur if the user performs a Select > All followed by Select > Invert.)
    If PDImages.GetActiveImage.MainSelection.FindNewBoundsManually() Then
    
        'At least one valid selection pixel still exists.  Activate it as the "new" selection.
        
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
            
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'No selection pixels exist.  Unload any active selection data.
    Else
        Selections.RemoveCurrentSelection
    End If

End Sub

'Feather the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub FeatherCurrentSelection(ByVal displayDialog As Boolean, Optional ByVal featherRadius As Double = 0#)

    'If a dialog has been requested, display one to the user.  Otherwise, proceed with the feathering.
    If displayDialog Then
        
        Dim retRadius As Double
        If DisplaySelectionDialog(pdsd_Feather, retRadius) = vbOK Then
            Process "Feather selection", False, TextSupport.BuildParamList("filtervalue", retRadius), UNDO_Selection
        End If
        
    Else
    
        Message "Feathering selection..."
    
        'Unlock any existing selection, and condense any composite selections down to a single raster layer.
        PDImages.GetActiveImage.MainSelection.LockRelease
        PDImages.GetActiveImage.SetSelectionActive False
        PDImages.GetActiveImage.MainSelection.SquashCompositeToRaster
        
        'Retrieve just the alpha channel of the current selection
        Dim tmpArray() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, tmpArray
        
        'Blur that temporary array
        Dim arrWidth As Long, arrHeight As Long
        arrWidth = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth
        arrHeight = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBHeight
        Filters_ByteArray.HorizontalBlur_ByteArray tmpArray, arrWidth, arrHeight, featherRadius, featherRadius
        Filters_ByteArray.VerticalBlur_ByteArray tmpArray, arrWidth, arrHeight, featherRadius, featherRadius
        
        'Reconstruct the DIB from the transparency table
        DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, tmpArray
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection
        ' (such as being a non-vector raster selection now)
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        PDImages.GetActiveImage.MainSelection.FindNewBoundsManually     'This sets selection shape to RASTER
        
        'Lock in the completed selection and ensure it's still marked as ACTIVE
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
        
        'Finalize any remaining UI elements
        ProgressBars.SetProgBarVal 0
        ProgressBars.ReleaseProgressBar
        Message "Feathering complete."
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If

End Sub

'Sharpen (un-feather?) the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub SharpenCurrentSelection(ByVal displayDialog As Boolean, Optional ByVal sharpenRadius As Double = 0#)

    'If a dialog has been requested, display one to the user.  Otherwise, proceed with the feathering.
    If displayDialog Then
        
        Dim retRadius As Double
        If (DisplaySelectionDialog(pdsd_Sharpen, retRadius) = vbOK) Then
            Process "Sharpen selection", False, BuildParamList("filtervalue", retRadius), UNDO_Selection
        End If
        
    Else
    
        Message "Sharpening selection..."
    
        'Unlock any existing selection, and condense any composite selections down to a single raster layer.
        PDImages.GetActiveImage.MainSelection.LockRelease
        PDImages.GetActiveImage.SetSelectionActive False
        PDImages.GetActiveImage.MainSelection.SquashCompositeToRaster
                
        'Retrieve just the alpha channel of the current selection, and clone it so that we have two copies
        Dim tmpArray() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, tmpArray
        
        Dim tmpDstArray() As Byte
        ReDim tmpDstArray(0 To PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth - 1, PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBHeight - 1) As Byte
        CopyMemoryStrict VarPtr(tmpDstArray(0, 0)), VarPtr(tmpArray(0, 0)), PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth * PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBHeight
        
        'Blur the first temporary array
        Dim arrWidth As Long, arrHeight As Long
        arrWidth = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth
        arrHeight = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBHeight
        Filters_ByteArray.HorizontalBlur_ByteArray tmpArray, arrWidth, arrHeight, sharpenRadius, sharpenRadius
        Filters_ByteArray.VerticalBlur_ByteArray tmpArray, arrWidth, arrHeight, sharpenRadius, sharpenRadius
        
        'We're now going to perform an "unsharp mask" effect, but because we're using a single channel, it goes a bit faster
        Dim progBarCheck As Long
        SetProgBarMax PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBHeight
        progBarCheck = ProgressBars.FindBestProgBarValue()
        
        'ScaleFactor is used to apply the unsharp mask.  Maximum strength can be any value, but PhotoDemon locks it at 10
        ' for selections (which are predictably feathered, using exact gaussian techniques).
        Dim scaleFactor As Double, invScaleFactor As Double
        scaleFactor = sharpenRadius
        invScaleFactor = 1# - scaleFactor
        
        Dim iWidth As Long, iHeight As Long
        iWidth = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth - 1
        iHeight = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBHeight - 1
        
        Dim lOrig As Long, lBlur As Long, lDelta As Single, lFull As Single, lNew As Long
        Dim x As Long, y As Long
        
        Const ONE_DIV_255 As Double = 1# / 255#
        
        For y = 0 To iHeight
        For x = 0 To iWidth
            
            'Retrieve the original and blurred byte values
            lOrig = tmpDstArray(x, y)
            lBlur = tmpArray(x, y)
            
            'Calculate the delta between the two, which is then converted to a blend factor
            lDelta = Abs(lOrig - lBlur) * ONE_DIV_255
            
            'Calculate a "fully" sharpened value; we're going to manually feather between this value and the original,
            ' based on the delta between the two.
            lFull = (scaleFactor * lOrig) + (invScaleFactor * lBlur)
            
            'Feather to arrive at a final "unsharp" value
            lNew = (1# - lDelta) * lFull + (lDelta * lOrig)
            If (lNew < 0) Then
                lNew = 0
            ElseIf (lNew > 255) Then
                lNew = 255
            End If
            
            'Since we're doing a per-pixel loop, we can safely store the result back into the destination array
            tmpDstArray(x, y) = lNew
            
        Next x
            If (x And progBarCheck) = 0 Then SetProgBarVal y
        Next y
        
        'Reconstruct the DIB from the finished transparency table
        DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, tmpDstArray
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
        ' being non-transformable)
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        PDImages.GetActiveImage.MainSelection.FindNewBoundsManually
        
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
                
        SetProgBarVal 0
        ReleaseProgressBar
        
        Message "Feathering complete."
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If

End Sub

'Grow the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub GrowCurrentSelection(ByVal displayDialog As Boolean, Optional ByVal growSize As Double = 0#)

    'If a dialog has been requested, display one to the user.  Otherwise, proceed with the feathering.
    If displayDialog Then
        
        Dim retSize As Double
        If DisplaySelectionDialog(pdsd_Grow, retSize) = vbOK Then
            Process "Grow selection", False, BuildParamList("filtervalue", retSize), UNDO_Selection
        End If
        
    Else
    
        Message "Growing selection..."
    
        'Unlock any existing selection, and condense any composite selections down to a single raster layer.
        PDImages.GetActiveImage.MainSelection.LockRelease
        PDImages.GetActiveImage.SetSelectionActive False
        PDImages.GetActiveImage.MainSelection.SquashCompositeToRaster
        
        'Use PD's built-in Median function to dilate the selected area
        Dim arrWidth As Long, arrHeight As Long
        arrWidth = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth
        arrHeight = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBHeight
        
        Dim tmpArray() As Byte
        ReDim tmpArray(0 To arrWidth - 1, 0 To arrHeight - 1) As Byte
        
        Dim srcBytes() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, srcBytes
        
        If Filters_ByteArray.Dilate_ByteArray(growSize, PDPRS_Circle, srcBytes, tmpArray, arrWidth, arrHeight) Then
            DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, tmpArray
        End If
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
        ' being non-transformable)
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        PDImages.GetActiveImage.MainSelection.FindNewBoundsManually
        
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
                
        SetProgBarVal 0
        ReleaseProgressBar
        
        Message "Selection resize complete."
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If
    
End Sub

'Shrink the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub ShrinkCurrentSelection(ByVal displayDialog As Boolean, Optional ByVal shrinkSize As Double = 0#)

    'If a dialog has been requested, display one to the user.  Otherwise, proceed with the feathering.
    If displayDialog Then
        
        Dim retSize As Double
        If DisplaySelectionDialog(pdsd_Shrink, retSize) = vbOK Then
            Process "Shrink selection", False, BuildParamList("filtervalue", retSize), UNDO_Selection
        End If
        
    Else
    
        Message "Shrinking selection..."
    
        'Unlock any existing selection, and condense any composite selections down to a single raster layer.
        PDImages.GetActiveImage.MainSelection.LockRelease
        PDImages.GetActiveImage.SetSelectionActive False
        PDImages.GetActiveImage.MainSelection.SquashCompositeToRaster
        
        'Use PD's built-in Median function to dilate the selected area
        Dim arrWidth As Long, arrHeight As Long
        arrWidth = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth
        arrHeight = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBHeight
        
        Dim tmpArray() As Byte
        ReDim tmpArray(0 To arrWidth - 1, 0 To arrHeight - 1) As Byte
        
        Dim srcBytes() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, srcBytes
        
        Filters_ByteArray.Erode_ByteArray shrinkSize, PDPRS_Circle, srcBytes, tmpArray, arrWidth, arrHeight
        DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, tmpArray
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
        ' being non-transformable)
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        PDImages.GetActiveImage.MainSelection.FindNewBoundsManually
        
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
                
        SetProgBarVal 0
        ReleaseProgressBar
        
        Message "Selection resize complete."
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If
    
End Sub

'Convert the current selection to border-type.  Note that this will make a transformable selection non-transformable.
Public Sub BorderCurrentSelection(ByVal displayDialog As Boolean, Optional ByVal borderRadius As Double = 0#)

    'If a dialog has been requested, display one to the user.  Otherwise, proceed with the feathering.
    If displayDialog Then
        
        Dim retSize As Double
        If DisplaySelectionDialog(pdsd_Border, retSize) = vbOK Then
            Process "Border selection", False, BuildParamList("filtervalue", retSize), UNDO_Selection
        End If
        
    Else
    
        Message "Finding selection border..."
    
        'Unlock any existing selection, and condense any composite selections down to a single raster layer.
        PDImages.GetActiveImage.MainSelection.LockRelease
        PDImages.GetActiveImage.SetSelectionActive False
        PDImages.GetActiveImage.MainSelection.SquashCompositeToRaster
        
        'Bordering a selection requires two passes: a grow pass and a shrink pass.  The results of these two passes are then blended
        ' to create the final bordered selection.
        
        'First, extract selection data into a byte array so we can use optimized analysis functions
        Dim arrWidth As Long, arrHeight As Long
        arrWidth = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth
        arrHeight = PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBHeight
        
        Dim srcArray() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, srcArray
        
        'Next, generate a shrink (erode) pass
        Dim shrinkBytes() As Byte
        ReDim shrinkBytes(0 To arrWidth - 1, 0 To arrHeight - 1) As Byte
        Filters_ByteArray.Erode_ByteArray borderRadius, PDPRS_Circle, srcArray, shrinkBytes, arrWidth, arrHeight, False, PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth * 2
        
        'Generate a grow (dilate) pass
        Dim growBytes() As Byte
        ReDim growBytes(0 To arrWidth - 1, 0 To arrHeight - 1) As Byte
        Filters_ByteArray.Dilate_ByteArray borderRadius, PDPRS_Circle, srcArray, growBytes, arrWidth, arrHeight, False, PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth * 2, PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBWidth
        
        'Finally, XOR those results together: that's our border!
        Dim x As Long, y As Long
        For y = 0 To arrHeight - 1
        For x = 0 To arrWidth - 1
            srcArray(x, y) = shrinkBytes(x, y) Xor growBytes(x, y)
        Next x
        Next y
        
        'Reconstruct the target DIB from our final array
        DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, srcArray
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
        ' being non-transformable)
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        PDImages.GetActiveImage.MainSelection.FindNewBoundsManually
                
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
                
        SetProgBarVal 0
        ReleaseProgressBar
        
        Message "Selection resize complete."
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If
    
End Sub

'TODO: UI for settings
Public Sub Selection_ContentAwareFill(ByVal displayDialog As Boolean)
    
    'Ensure a selection exists
    If (Not Selections.SelectionsAllowed(False)) Then Exit Sub
    
    If displayDialog Then
        Processor.Process "Content-aware fill", False, vbNullString, UNDO_Layer
    Else
            
        'If this is a vector layer, rasterize it
        If PDImages.GetActiveImage.GetActiveLayer.IsLayerVector Then Layers.RasterizeLayer PDImages.GetActiveImage.GetActiveLayerIndex
        
        'Retrieve the current selection mask as a simple byte array.  Note that the array *will* likely be smaller
        ' than the current image/layer/whatever.
        '
        'TODO: clone the region of the selection mask that overlays the current layer, then work from there.
        ' (The current solution only works on single-layer images.)
        Dim srcMask() As Byte, srcMaskRect As RectF
        srcMaskRect = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect
        DIBs.GetSingleChannel_2D PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, srcMask, 0, VarPtr(srcMaskRect)
        
        'Execute the fill
        Dim newLayerDIB As pdDIB
        Set newLayerDIB = New pdDIB
        newLayerDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveDIB
        
        Dim cInpaint As pdInpaint
        Set cInpaint = New pdInpaint
        cInpaint.ContentAwareFill PDImages.GetActiveImage.GetActiveDIB, newLayerDIB, srcMask, srcMaskRect
        
        'Assign the new image and notify the parent image of the change
        PDImages.GetActiveImage.GetActiveLayer.SetLayerDIB newLayerDIB
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
        
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If
    
End Sub
