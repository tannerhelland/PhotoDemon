Attribute VB_Name = "FastDrawing"
'***************************************************************************
'Fast API Graphics Routines Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 12/June/01
'Last updated: 31/August/12
'Last update: Completed work on prepImageData and finalizeImageData, which are the much-improved successors to
'             Get/SetImageData.  These routines rely on SafeArrays instead of Get/SetDIBits, so they are quite a bit
'             faster.  They are also image independent; passing a "preview" flag will result in the ability to paint the
'             results of a filter to any picture box, and it's all managed internally - meaning the filter/effect routine
'             doesn't have to worry about a thing.  A public "curLayerValues" variable contains everything a filter could
'             ever want to know about the data it's working on.  Most of the values it provides are unused at present, but
'             could be useful once selections/layers are implemented.
'
'This interface provides API support for the main image interaction routines. It assigns memory data
' into a useable array, and later transfers that array back into memory.  Very fast, very compact, can't
' live without it. These functions are arguably the most integral part of PhotoDemon.
'
'If you want to know more about how DIB sections work - and why they're so fast compared to VB's internal
' .PSet and .Point methods - please visit http://www.tannerhelland.com/42/vb-graphics-programming-3/
'
'***************************************************************************

Option Explicit

Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Long, lpSrc As Long, ByVal byteLength As Long)

Public workingLayer As pdLayer

'prepImageData will fill a variable of this type with everything a filter or effect could possibly want to know
' about the layer it's operating on.  Filters are free to ignore this data, but it is always made available.
Public Type FilterInfo
    Left As Long            'Coordinates of the top-left location the filter is supposed to operate on
    Top As Long
    Right As Long           'Right and Bottom could be inferred, but we do it here to minimize effort on the calling routine's part
    Bottom As Long
    Width As Long           'Dimensions of the area the filter is supposed to operate on
    Height As Long
    MinX As Long            'The lowest coordinate the filter is allowed to check.  This is almost always (0, 0)
    MinY As Long
    MaxX As Long            'The highest coordinate the filter is allowed to check.  This is almost always (width, height)
    MaxY As Long
    colorDepth As Long      'The colorDepth of the current layer; right now, this should always be 24 or 32
    BytesPerPixel As Long   'BPP is colorDepth / 8.  It is provided for convenience.
    LayerX As Long          'Filters shouldn't have to worry about where the layer is physically located, but when it comes
    LayerY As Long          ' time to set the layer back in place, these may be useful (as when previewing, for example)
End Type

Public curLayerValues As FilterInfo


'prepPixelData's job is to copy the relevant layer into a temporary object, which is what individual filters and effects
' will operate on.  prepPixelData() also populates the relevant SafeArray object and a host of other variables, which
' filters and effects can then copy locally to ensure the fastest possible runtime speed.
'
'If the filter will be rendering a preview only, it can specify the picture box that will receive the preview effect.
' This function will automatically adjust its parameters accordingly, and the filter routine will not have to make any
' modifications to its code.
'
'Finally, the calling routine can optionally specify a different progress bar maximum value.  By default, this is the current
' layer's width, but some routines run vertically and the progress bar needs to be changed accordingly.
Public Sub prepImageData(ByRef tmpSA As SAFEARRAY2D, Optional isPreview As Boolean = False, Optional previewPictureBox As PictureBox, Optional newProgBarMax As Long = -1)

    'Prepare our temporary layer
    Set workingLayer = New pdLayer
    
    'If this is not a preview, simply copy the current layer without modification
    If isPreview = False Then
    
        'Check for an active selection; if one is present, use that instead of the full layer
        If pdImages(CurrentImage).selectionActive Then
            workingLayer.createBlank pdImages(CurrentImage).mainSelection.selWidth, pdImages(CurrentImage).mainSelection.selHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
            BitBlt workingLayer.getLayerDC, 0, 0, pdImages(CurrentImage).mainSelection.selWidth, pdImages(CurrentImage).mainSelection.selHeight, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).mainSelection.selLeft, pdImages(CurrentImage).mainSelection.selTop, vbSrcCopy
        Else
            workingLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer
        End If
    
    'If this IS a preview, more work is involved.
    Else
    
        'Start by calculating the aspect ratio of both the current image and the previewing picture box
        Dim dstWidth As Single, dstHeight As Single
        dstWidth = previewPictureBox.ScaleWidth
        dstHeight = previewPictureBox.ScaleHeight
    
        Dim srcWidth As Single, srcHeight As Single
        
        'The source values need to be adjusted contingent on whether this is a selection or a full-image preview
        If pdImages(CurrentImage).selectionActive Then
            srcWidth = pdImages(CurrentImage).mainSelection.selWidth
            srcHeight = pdImages(CurrentImage).mainSelection.selHeight
        Else
            srcWidth = pdImages(CurrentImage).mainLayer.getLayerWidth
            srcHeight = pdImages(CurrentImage).mainLayer.getLayerHeight
        End If
    
        Dim srcAspect As Single, dstAspect As Single
        srcAspect = srcWidth / srcHeight
        dstAspect = dstWidth / dstHeight
        
        'Now, use that aspect ratio to determine a proper size for our temporary layer
        Dim newWidth As Long, newHeight As Long
    
        If srcAspect > dstAspect Then
            newWidth = dstWidth
            newHeight = CSng(srcHeight / srcWidth) * newWidth + 0.5
        Else
            newHeight = dstHeight
            newWidth = CSng(srcWidth / srcHeight) * newHeight + 0.5
        End If
                
        'And finally, create our workingLayer using these values
        If pdImages(CurrentImage).selectionActive Then
            Dim copyLayer As pdLayer
            Set copyLayer = New pdLayer
            copyLayer.createBlank pdImages(CurrentImage).mainSelection.selWidth, pdImages(CurrentImage).mainSelection.selHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
            BitBlt copyLayer.getLayerDC, 0, 0, pdImages(CurrentImage).mainSelection.selWidth, pdImages(CurrentImage).mainSelection.selHeight, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).mainSelection.selLeft, pdImages(CurrentImage).mainSelection.selTop, vbSrcCopy
            workingLayer.createFromExistingLayer copyLayer, newWidth, newHeight
            copyLayer.eraseLayer
            Set copyLayer = Nothing
        Else
            workingLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, newWidth, newHeight
        End If
        
    End If
    
    'With our temporary layer successfully created, populate the relevant SafeArray variable
    prepSafeArray tmpSA, workingLayer

    'Finally, populate the ubiquitous curLayerValues variable with everything a filter might want to know
    With curLayerValues
        .Left = 0
        .Top = 0
        .Right = workingLayer.getLayerWidth - 1
        .Bottom = workingLayer.getLayerHeight - 1
        .Width = workingLayer.getLayerWidth
        .Height = workingLayer.getLayerHeight
        .MinX = 0
        .MinY = 0
        .MaxX = workingLayer.getLayerWidth - 1
        .MaxY = workingLayer.getLayerHeight - 1
        .colorDepth = workingLayer.getLayerColorDepth
        .BytesPerPixel = (workingLayer.getLayerColorDepth \ 8)
        .LayerX = 0
        .LayerY = 0
    End With

    'Set up the progress bar (only if this is NOT a preview, mind you)
    If isPreview = False Then
        If newProgBarMax = -1 Then
            SetProgBarMax (curLayerValues.Left + curLayerValues.Width)
        Else
            SetProgBarMax newProgBarMax
        End If
    End If
    
    'MsgBox "prepImageData worked: " & workingLayer.getLayerHeight & ", " & workingLayer.getLayerWidth & " (" & workingLayer.getLayerArrayWidth & ")" & ", " & workingLayer.getLayerDIBits

End Sub

'This function can be used to populate a valid SAFEARRAY2D structure against any layer
Public Sub prepSafeArray(ByRef srcSA As SAFEARRAY2D, ByRef srcLayer As pdLayer)
    
    'With our temporary layer successfully created, populate the relevant SafeArray variable
    With srcSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcLayer.getLayerHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcLayer.getLayerArrayWidth
        .pvData = srcLayer.getLayerDIBits
    End With
    
End Sub

'The counterpart to prepImageData, finalizeImageData copies the working layer back into its source then renders it
' to the screen.  Like prepImageData(), a preview target can also be named.  In this case, finalizeImageData will rely on
' the values calculated by prepImageData(), as it's presumed that preImageData will ALWAYS be called before this routine.
Public Sub finalizeImageData(Optional isPreview As Boolean = False, Optional previewPictureBox As PictureBox)

    'If this is not a preview, our job is simple - get the newly processed DIB rendered to the screen.
    If isPreview = False Then
        
        Message "Rendering image to screen..."
        
        'Paint the working layer over the original layer
        If pdImages(CurrentImage).selectionActive Then
            BitBlt pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).mainSelection.selLeft, pdImages(CurrentImage).mainSelection.selTop, workingLayer.getLayerWidth, workingLayer.getLayerHeight, workingLayer.getLayerDC, 0, 0, vbSrcCopy
        Else
            BitBlt pdImages(CurrentImage).mainLayer.getLayerDC, curLayerValues.LayerX, curLayerValues.LayerY, curLayerValues.Width, curLayerValues.Height, workingLayer.getLayerDC, 0, 0, vbSrcCopy
        End If
                
        'workingLayer has served its purpose, so erase it from memory
        Set workingLayer = Nothing
                
        'If we're setting data to the screen, we can reasonably assume that the progress bar should be reset
        SetProgBarVal 0
        
        'Pass control to the viewport renderer, which will perform the actual rendering
        ScrollViewport FormMain.ActiveForm
        
        Message "Finished."
        
    Else
    
        'If the current layer is 32bpp, precomposite it against a checkerboard background before rendering
        If workingLayer.getLayerColorDepth = 32 Then workingLayer.compositeBackgroundColor
    
        'Allow workingLayer to paint itself to the target picture box
        workingLayer.renderToPictureBox previewPictureBox
        
        'workingLayer has served its purpose, so erase it from memory
        Set workingLayer = Nothing
        
    End If
    
End Sub
