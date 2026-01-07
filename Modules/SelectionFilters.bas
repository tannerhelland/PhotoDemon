Attribute VB_Name = "SelectionFilters"
'***************************************************************************
'Selection Tools: Filters
'Copyright 2013-2026 by Tanner Helland
'Created: 21/June/13
'Last updated: 05/May/22
'Last update: new stroke, fill, and content-aware fill features
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
Public Sub Selection_Invert()

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
    
    'Ask the selection to find new boundaries.  This will also set all relevant parameters for the
    ' modified selection (such as being non-transformable)
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
        PDDebug.LogAction "No bounds found; removing selection."
        Selections.RemoveCurrentSelection
    End If

End Sub

'Feather the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub Selection_Blur(ByVal displayDialog As Boolean, Optional ByVal featherRadius As Double = 0#)

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
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the
        ' modified selection (such as being non-transformable)
        PDImages.GetActiveImage.MainSelection.SetSelectionShape ss_Raster
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        
        'Apply any final UI changes
        SetProgBarVal 0
        ReleaseProgressBar
        Message "Feathering complete."
        
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
            PDDebug.LogAction "No bounds found; removing selection."
            Selections.RemoveCurrentSelection
        End If

    End If

End Sub

'Sharpen (un-feather?) the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub Selection_Sharpen(ByVal displayDialog As Boolean, Optional ByVal sharpenRadius As Double = 0#)

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
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the
        ' modified selection (such as being non-transformable)
        PDImages.GetActiveImage.MainSelection.SetSelectionShape ss_Raster
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        
        'Apply any final UI changes
        SetProgBarVal 0
        ReleaseProgressBar
        Message "Feathering complete."
        
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
            PDDebug.LogAction "No bounds found; removing selection."
            Selections.RemoveCurrentSelection
        End If
    
    End If

End Sub

'Grow the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub Selection_Grow(ByVal displayDialog As Boolean, Optional ByVal growSize As Double = 0#)

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
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the
        ' modified selection (such as being non-transformable)
        PDImages.GetActiveImage.MainSelection.SetSelectionShape ss_Raster
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        
        'Apply any final UI changes
        SetProgBarVal 0
        ReleaseProgressBar
        Message "Selection resize complete."
        
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
            PDDebug.LogAction "No bounds found; removing selection."
            Selections.RemoveCurrentSelection
        End If

    End If
    
End Sub

'Shrink the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub Selection_Shrink(ByVal displayDialog As Boolean, Optional ByVal shrinkSize As Double = 0#)

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
        
        Dim tmpArray() As Byte, srcBytes() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, tmpArray
        Filters_ByteArray.PadByteArray_NoClamp tmpArray, arrWidth, arrHeight, srcBytes, 1, 1
        
        ReDim tmpArray(0 To arrWidth + 1, 0 To arrHeight + 1) As Byte
        Filters_ByteArray.Erode_ByteArray shrinkSize, PDPRS_Circle, srcBytes, tmpArray, arrWidth + 2, arrHeight + 2
        
        ReDim srcBytes(0 To arrWidth, 0 To arrHeight) As Byte
        Filters_ByteArray.UnPadByteArray srcBytes, arrWidth, arrHeight, tmpArray, 1, 1, True
        
        DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, srcBytes
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the
        ' modified selection (such as being non-transformable)
        PDImages.GetActiveImage.MainSelection.SetSelectionShape ss_Raster
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        
        'Apply any final UI changes
        SetProgBarVal 0
        ReleaseProgressBar
        Message "Selection resize complete."
        
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
            PDDebug.LogAction "No bounds found; removing selection."
            Selections.RemoveCurrentSelection
        End If

    End If
    
End Sub

'Convert the current selection to border-type.  Note that this will make a transformable selection non-transformable.
Public Sub Selection_ConvertToBorder(ByVal displayDialog As Boolean, Optional ByVal borderRadius As Double = 0#)

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
        
        'To ensure correct edge behavior (particularly on Select > All), null-pad the array with a blank
        ' row/column of pixels on all sides.
        Dim srcArrayPadded() As Byte
        Filters_ByteArray.PadByteArray_NoClamp srcArray, arrWidth, arrHeight, srcArrayPadded, 1, 1
        
        'Next, generate a shrink (erode) pass
        Dim shrinkBytes() As Byte
        ReDim shrinkBytes(0 To arrWidth + 1, 0 To arrHeight + 1) As Byte
        Filters_ByteArray.Erode_ByteArray borderRadius, PDPRS_Circle, srcArrayPadded, shrinkBytes, arrWidth + 2, arrHeight + 2, False, arrWidth * 2
        
        'Generate a grow (dilate) pass
        Dim growBytes() As Byte
        ReDim growBytes(0 To arrWidth + 1, 0 To arrHeight + 1) As Byte
        Filters_ByteArray.Dilate_ByteArray borderRadius, PDPRS_Circle, srcArrayPadded, growBytes, arrWidth + 2, arrHeight + 2, False, arrWidth * 2, arrWidth + 1
        
        'Finally, XOR those results together: that's our border!
        Dim x As Long, y As Long
        For y = 0 To arrHeight + 1
        For x = 0 To arrWidth + 1
            srcArrayPadded(x, y) = shrinkBytes(x, y) Xor growBytes(x, y)
        Next x
        Next y
        
        'Unpad the final array back to original dimensions (removing the blank row/column on all edges)
        ReDim srcArray(0 To arrWidth - 1, 0 To arrHeight - 1) As Byte
        Filters_ByteArray.UnPadByteArray srcArray, arrWidth, arrHeight, srcArrayPadded, 1, 1, True
        
        'Reconstruct the target DIB from the final array
        DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, srcArray
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the
        ' modified selection (such as being non-transformable)
        PDImages.GetActiveImage.MainSelection.SetSelectionShape ss_Raster
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        
        'Apply any final UI changes
        SetProgBarVal 0
        ReleaseProgressBar
        Message "Selection resize complete."
        
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
            PDDebug.LogAction "No bounds found; removing selection."
            Selections.RemoveCurrentSelection
        End If
        
    End If
    
End Sub

Public Sub Selection_Clear(ByVal displayDialog As Boolean)
    
    'Ignore the "display dialog" setting; there's no UI for this tool
    If displayDialog Then
        Process "Clear", False, vbNullString, UNDO_Layer
    Else
        
        'If a selection is active, use it as the basis for the clear
        If PDImages.GetActiveImage.IsSelectionActive() Then
            Selections.EraseSelectedArea PDImages.GetActiveImage.GetActiveLayerIndex
        
        'If a selection is *not* active, simply erase the current layer
        Else
            PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB.ResetDIB 0
            PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
            Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        End If
        
    End If
    
End Sub

'Content-aware fill (aka "Heal Selection") was added in v9.0
Public Sub Selection_ContentAwareFill(ByVal displayDialog As Boolean, Optional ByRef fxParams As String = vbNullString)
    
    'Ensure a selection exists
    If (Not Selections.SelectionsAllowed(False)) Then Exit Sub
    
    If displayDialog Then
        Interface.ShowPDDialog vbModal, FormFillContentAware
    Else
        
        'Prep a parameter retrieval class
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString fxParams
        
        'Prepare the source layer for in-painting.
        
        'IMPORTANT NOTE: in early drafts, this feature passed the full-size source layer to pdInpaint.
        ' This works fine, but it isn't necessary, and it consumes extra memory because the mask and a
        ' bunch of temporary structures inside pdInpaint must all be created at the same size as the
        ' source image.
        '
        'So instead, we now create a temporary copy of the source layer at the minimum size required by
        ' the inpainter.  This minimum size is easy to calculate - it's just the rectangle of the target
        ' area (the area "to be filled" as defined by the mask) expanded by the user's sampling radius.
        '
        'This greatly reduces memory requirements of the inpainter and it's much faster to prepare the
        ' cropped version here, rather than adding extra boundary checks to the inpainter (which has to
        ' analyze pixels millions of times on the average fill - and bounds-checking every one of those
        ' accesses becomes disproportionately expensive).
        '
        'Anyway, I mention this up-front because you can easily pass a full-size source image and mask
        ' to the inpainter and it will work great.  It will just consume a lot more memory.
        
        'If this is a vector layer, rasterize it
        If PDImages.GetActiveImage.GetActiveLayer.IsLayerVector Then Layers.RasterizeLayer PDImages.GetActiveImage.GetActiveLayerIndex
        
        'If the layer has any active affine transforms, make them permanent
        If PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True) Then PDImages.GetActiveImage.GetActiveLayer.MakeCanvasTransformsPermanent
        
        'These are the source and destination DIBs we're ultimately going to fill, as well as the
        ' mask byte array (and rectangle defining the mask boundary area), but how we fill these
        ' depends on the size and position of the active layer vs its parent image.
        Dim tmpSrcCopy As pdDIB, tmpDstCopy As pdDIB
        Dim srcMask() As Byte
        
        Set tmpSrcCopy = New pdDIB
        Set tmpDstCopy = New pdDIB
        
        'Retrieve the boundary rectangle of the "to-be-filled region".  (This comes directly from
        ' PhotoDemon's selection engine - it's just the boundary of the current selection, in image
        ' coordinate space.)
        Dim baseFillRect As RectF
        baseFillRect = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect
        
        'Some selection methods may produce boundary rects with floating-point values (due to the way
        ' mouse inputs are handled).  Because we'll be using these to address pixels, we want all values
        ' clamped to their nearest equivalent integer.
        baseFillRect.Width = Int(baseFillRect.Width + PDMath.Frac(baseFillRect.Left) + 0.5)
        baseFillRect.Height = Int(baseFillRect.Height + PDMath.Frac(baseFillRect.Top) + 0.5)
        baseFillRect.Left = Int(baseFillRect.Left)
        baseFillRect.Top = Int(baseFillRect.Top)
        If (baseFillRect.Width <= 0) Or (baseFillRect.Height <= 0) Then Exit Sub
        
        'Note the centroid of the selection
        Dim selectionCentroid As PointFloat
        selectionCentroid.x = (baseFillRect.Left + baseFillRect.Width / 2)
        selectionCentroid.y = (baseFillRect.Top + baseFillRect.Height / 2)
        
        'Determine whether we are allowed to sample in all directions
        Dim sampleUp As Boolean, sampleDown As Boolean, sampleLeft As Boolean, sampleRight As Boolean
        sampleUp = cParams.GetBool("sample-up", True, True)
        sampleLeft = cParams.GetBool("sample-left", True, True)
        sampleRight = cParams.GetBool("sample-right", True, True)
        sampleDown = cParams.GetBool("sample-down", True, True)
        
        'The source and destination images need to be the same size as this "to-be-filled region",
        ' but expanded by the user's [sampling radius] in all directions the user allows.
        Dim userSampleRadius As Long
        userSampleRadius = cParams.GetLong("search-radius", 200)
        If (userSampleRadius < 1) Then userSampleRadius = 1
        If (userSampleRadius > 500) Then userSampleRadius = 500
        
        Dim expandedFillRect As RectF
        expandedFillRect.Left = baseFillRect.Left
        If sampleLeft Then expandedFillRect.Left = expandedFillRect.Left - userSampleRadius
        expandedFillRect.Top = baseFillRect.Top
        If sampleUp Then expandedFillRect.Top = expandedFillRect.Top - userSampleRadius
        
        Dim widenAmount As Long
        If sampleLeft Then widenAmount = userSampleRadius Else widenAmount = 0
        If sampleRight Then widenAmount = widenAmount + userSampleRadius
        expandedFillRect.Width = baseFillRect.Width + widenAmount
        
        If sampleUp Then widenAmount = userSampleRadius Else widenAmount = 0
        If sampleDown Then widenAmount = widenAmount + userSampleRadius
        expandedFillRect.Height = baseFillRect.Height + widenAmount
        
        'The user is allowed to inpaint along image boundaries (this is actually a common use-case),
        ' so pre-trim the target rectangle to the boundaries of the parent image.
        If (expandedFillRect.Left < 0) Then expandedFillRect.Left = 0
        If (expandedFillRect.Top < 0) Then expandedFillRect.Top = 0
        If (expandedFillRect.Left + expandedFillRect.Width > PDImages.GetActiveImage.Width()) Then expandedFillRect.Width = (PDImages.GetActiveImage.Width() - expandedFillRect.Left)
        If (expandedFillRect.Top + expandedFillRect.Height > PDImages.GetActiveImage.Height()) Then expandedFillRect.Height = (PDImages.GetActiveImage.Height() - expandedFillRect.Top)
        
        'We now have a properly expanded target rectangle.  This rectangle is the size we want the
        ' source image, destination image, and mask to be.
        
        'Note, however, that this rectangle is in *IMAGE* coordinates - not *LAYER* coordinates.
        ' If the active layer and image are the same size, great - we can use this as-is.  If the
        ' current layer is a different size than the parent image, however, we're gonna need to
        ' perform additional math (and modify the above rectangle accordingly).
        
        'To determine whether that extra math is necessary, see if the current layer is the same
        ' size as the parent image (as in e.g. a normal single-layer JPEG), or whether it's a
        ' different size or has non-zero offsets (as in e.g. a PSD).
        Dim layerNotFullSize As Boolean
        layerNotFullSize = (PDImages.GetActiveImage.GetActiveLayer.GetLayerOffsetX <> 0)
        layerNotFullSize = layerNotFullSize Or (PDImages.GetActiveImage.GetActiveLayer.GetLayerOffsetY <> 0)
        layerNotFullSize = layerNotFullSize Or (PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(True) <> PDImages.GetActiveImage.Width)
        layerNotFullSize = layerNotFullSize Or (PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(True) <> PDImages.GetActiveImage.Height)
        
        If layerNotFullSize Then
            
            'This layer has non-zero offsets or a size that varies from its parent image.  We need to calculate
            ' an intersection between it and the target rectangle we calculated above.
            Dim origLayerRect As RectF
            PDImages.GetActiveImage.GetActiveLayer.GetLayerBoundaryRect origLayerRect
            
            Dim overlapRectImage As RectF
            If (Not GDI_Plus.IntersectRectF(overlapRectImage, origLayerRect, expandedFillRect)) Then Exit Sub
            
            'We're still here, so the current layer and the target selection overlap, and that region of
            ' overlap is stored in overlapRectImage (so-named because it's in IMAGE coordinates).
            ' Calculate the same rect, but in layer coordinates.
            Dim overlapRectLayer As RectF
            With PDImages.GetActiveImage.GetActiveLayer
                overlapRectLayer.Left = overlapRectImage.Left - .GetLayerOffsetX
                overlapRectLayer.Top = overlapRectImage.Top - .GetLayerOffsetY
                overlapRectLayer.Width = overlapRectImage.Width
                overlapRectLayer.Height = overlapRectImage.Height
            End With
            
            'We have everything we need to prepare the inpainter!  Pull the relevant rectangle from the
            ' source layer and mirror it into the destination DIB.
            tmpSrcCopy.CreateBlank Int(overlapRectLayer.Width), Int(overlapRectLayer.Height), 32, 0, 0
            GDI.BitBltWrapper tmpSrcCopy.GetDIBDC, 0, 0, Int(overlapRectLayer.Width), Int(overlapRectLayer.Height), PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB.GetDIBDC, Int(overlapRectLayer.Left), Int(overlapRectLayer.Top), vbSrcCopy
            
            'Pull the relevant rect out of the selection mask as well
            DIBs.GetSingleChannel_2D PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, srcMask, 0, VarPtr(overlapRectImage)
            
        'Layer is the size of the image (e.g. a single-layer JPEG), which makes this step much easier!
        Else
            
            'The boundary rectangle we have already calculated will work fine.  Crop out the relevant
            ' region from the source image and clone it to the destination image.
            tmpSrcCopy.CreateBlank Int(expandedFillRect.Width), Int(expandedFillRect.Height), 32, 0, 0
            GDI.BitBltWrapper tmpSrcCopy.GetDIBDC, 0, 0, Int(expandedFillRect.Width), Int(expandedFillRect.Height), PDImages.GetActiveImage.GetActiveDIB.GetDIBDC, Int(expandedFillRect.Left), Int(expandedFillRect.Top), vbSrcCopy
            
            'Retrieve the selection mask using the same rect.
            DIBs.GetSingleChannel_2D PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, srcMask, 0, VarPtr(expandedFillRect)
            
        End If
        
        'Clone the source DIB into a temporary "destination" DIB.  This DIB will contain the result of
        ' the content-aware fill operation.  (But if the user cancels the operation, we must leave the
        ' original layer DIB untouched.)
        tmpDstCopy.CreateFromExistingDIB tmpSrcCopy
        
        'Pass all user parameters to the inpainting class
        Dim cInpaint As pdInpaint
        Set cInpaint = New pdInpaint
        cInpaint.SetAllowOutliers cParams.GetDouble("allow-outliers", 0.15)
        cInpaint.SetMaxNumNeighbors cParams.GetLong("patch-size", 20)
        cInpaint.SetMaxRandomCandidates cParams.GetLong("random-candidates", 60)
        cInpaint.SetRefinement cParams.GetDouble("refinement", 0.5)
        cInpaint.SetSearchRadius cParams.GetLong("search-radius", 200)
        cInpaint.SetSamplingCenter (selectionCentroid.x - expandedFillRect.Left), (selectionCentroid.y - expandedFillRect.Top)
        
        'Execute the content-aware fill
        cInpaint.ContentAwareFill tmpDstCopy, srcMask, True
        
        'TODO: check success/fail after adding user-cancellation support
        
        'With the fill finished, we now need to blend the results against the original layer.  Note that
        ' non-masked pixels (0) and fully-masked pixels (255) have already been dealt with - it's values
        ' between 0 and 255 that we must handle here.  (The inpainter only fills whole pixels, so any
        ' feathering needs to be manually handled by the caller.)
        '
        'Note also that this version of the code works with both straight and premultiplied alpha.
        ' It doesn't actually perform alpha-blending - it only performs a weighted average of the
        ' before-and-after results.  This is deliberate to produce correct results when blending
        ' semi-transparent pixels that are only partially selected.
        Dim x As Long, y As Long, xOffset As Long, selValue As Byte
        Dim xMax As Long, yMax As Long
        xMax = tmpSrcCopy.GetDIBWidth - 1
        yMax = tmpSrcCopy.GetDIBHeight - 1
        
        Dim pxSrc() As Byte, saSrc As SafeArray1D, ptrSrc As Long, strideSrc As Long
        tmpSrcCopy.WrapArrayAroundScanline pxSrc, saSrc, 0
        ptrSrc = saSrc.pvData
        strideSrc = saSrc.cElements
        
        Dim pxDst() As Byte, saDst As SafeArray1D, ptrDst As Long, strideDst As Long
        tmpDstCopy.WrapArrayAroundScanline pxDst, saDst, 0
        ptrDst = saDst.pvData
        strideDst = saDst.cElements
        
        Dim newR As Long, newG As Long, newB As Long, newA As Long
        Dim oldR As Long, oldG As Long, oldB As Long, oldA As Long
        
        Dim blendAmount As Double
        Const ONE_DIV_255 As Double = 1# / 255#
        
        For y = 0 To yMax
            
            'Update array pointers to point at the current line in both the source and destination images
            saSrc.pvData = ptrSrc + strideSrc * y
            saDst.pvData = ptrDst + strideDst * y
            
        For x = 0 To xMax
            
            selValue = srcMask(x, y)
            If (selValue > 0) Then
                If (selValue < 255) Then
                    
                    xOffset = x * 4
                    
                    newB = pxDst(xOffset)
                    newG = pxDst(xOffset + 1)
                    newR = pxDst(xOffset + 2)
                    newA = pxDst(xOffset + 3)
                    
                    oldB = pxSrc(xOffset)
                    oldG = pxSrc(xOffset + 1)
                    oldR = pxSrc(xOffset + 2)
                    oldA = pxSrc(xOffset + 3)
                    
                    'Calculate a weighted blend of the old and new pixel values and replace the destination
                    ' value with the result.
                    blendAmount = selValue * ONE_DIV_255
                    pxDst(xOffset) = (blendAmount * newB) + oldB * (1# - blendAmount)
                    pxDst(xOffset + 1) = (blendAmount * newG) + oldG * (1# - blendAmount)
                    pxDst(xOffset + 2) = (blendAmount * newR) + oldR * (1# - blendAmount)
                    pxDst(xOffset + 3) = (blendAmount * newA) + oldA * (1# - blendAmount)
                    
                End If
            End If
            
        Next x
        Next y
        
        'Free unsafe array wrappers!
        tmpSrcCopy.UnwrapArrayFromDIB pxSrc
        tmpDstCopy.UnwrapArrayFromDIB pxDst
        
        'Copy the finished result *back* into the active source layer
        If layerNotFullSize Then
            GDI.BitBltWrapper PDImages.GetActiveImage.GetActiveDIB.GetDIBDC, Int(overlapRectLayer.Left), Int(overlapRectLayer.Top), Int(overlapRectLayer.Width), Int(overlapRectLayer.Height), tmpDstCopy.GetDIBDC, 0, 0, vbSrcCopy
        Else
            GDI.BitBltWrapper PDImages.GetActiveImage.GetActiveDIB.GetDIBDC, Int(expandedFillRect.Left), Int(expandedFillRect.Top), Int(expandedFillRect.Width), Int(expandedFillRect.Height), tmpDstCopy.GetDIBDC, 0, 0, vbSrcCopy
        End If
        
        'Notify the parent image of the change, then redraw the viewport before exiting
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If
    
End Sub

'Fill the selected area of the target layer.  (If a selection is *not* active, just fill the whole layer.)
Public Sub Selection_Fill(ByVal displayDialog As Boolean, Optional ByRef fxParams As String = vbNullString)
    If displayDialog Then
        Interface.ShowPDDialog vbModal, FormFill
    Else
        FormFill.ApplyFillEffect fxParams, False
    End If
End Sub

'Stroke the boundary of the target layer.  (If a selection is *not* active, stroke the boundary of the current layer.)
Public Sub Selection_Stroke(ByVal displayDialog As Boolean, Optional ByRef fxParams As String = vbNullString)
    If displayDialog Then
        Interface.ShowPDDialog vbModal, FormStroke
    Else
        FormStroke.ApplyStrokeEffect fxParams, False
    End If
End Sub
