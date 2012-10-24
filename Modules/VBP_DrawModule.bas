Attribute VB_Name = "Drawing"
'***************************************************************************
'PhotoDemon Drawing Routines
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/3/01
'Last updated: 03/October/12
'Last update: Rewrote DrawPreviewImage to respect selections
'
'Miscellaneous drawing routines that don't fit elsewhere.  At present, this includes rendering preview images,
' drawing the canvas background of image forms, and a gradient-rendering sub (used primarily on the histogram form).
'
'***************************************************************************

Option Explicit

'Used to draw the main image onto a preview picture box
Public Sub DrawPreviewImage(ByRef dstPicture As PictureBox, Optional ByVal useOtherPictureSrc As Boolean = False, Optional ByRef otherPictureSrc As pdLayer, Optional forceWhiteBackground As Boolean = False)
    
    Dim tmpLayer As pdLayer
    
    'Start by calculating the aspect ratio of both the current image and the previewing picture box
    Dim dstWidth As Single, dstHeight As Single
    dstWidth = dstPicture.ScaleWidth
    dstHeight = dstPicture.ScaleHeight
    
    Dim SrcWidth As Single, SrcHeight As Single
    
    'The source values need to be adjusted contingent on whether this is a selection or a full-image preview
    If pdImages(CurrentImage).selectionActive Then
        SrcWidth = pdImages(CurrentImage).mainSelection.selWidth
        SrcHeight = pdImages(CurrentImage).mainSelection.selHeight
    Else
        SrcWidth = pdImages(CurrentImage).mainLayer.getLayerWidth
        SrcHeight = pdImages(CurrentImage).mainLayer.getLayerHeight
    End If
    
    Dim srcAspect As Single, dstAspect As Single
    srcAspect = SrcWidth / SrcHeight
    dstAspect = dstWidth / dstHeight
        
    'Now, use that aspect ratio to determine a proper size for our temporary layer
    Dim newWidth As Long, newHeight As Long
    
    If srcAspect > dstAspect Then
        newWidth = dstWidth
        newHeight = CSng(SrcHeight / SrcWidth) * newWidth + 0.5
    Else
        newHeight = dstHeight
        newWidth = CSng(SrcWidth / SrcHeight) * newHeight + 0.5
    End If
    
    'Normally this will draw a preview of FormMain.ActiveForm's relevant image.  However, another picture source can be specified.
    If useOtherPictureSrc = False Then
        
        'Check to see if a selection is active; if it isn't, simply render the full form
        If pdImages(CurrentImage).selectionActive = False Then
        
            If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then
                Set tmpLayer = New pdLayer
                tmpLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, newWidth, newHeight, True
                If forceWhiteBackground Then tmpLayer.compositeBackgroundColor 255, 255, 255 Else tmpLayer.compositeBackgroundColor
                tmpLayer.renderToPictureBox dstPicture
            Else
                pdImages(CurrentImage).mainLayer.renderToPictureBox dstPicture
            End If
        
        Else
        
            'Copy the current selection into a temporary layer
            Set tmpLayer = New pdLayer
            tmpLayer.createBlank pdImages(CurrentImage).mainSelection.selWidth, pdImages(CurrentImage).mainSelection.selHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
            BitBlt tmpLayer.getLayerDC, 0, 0, pdImages(CurrentImage).mainSelection.selWidth, pdImages(CurrentImage).mainSelection.selHeight, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).mainSelection.selLeft, pdImages(CurrentImage).mainSelection.selTop, vbSrcCopy
        
            'If the image is transparent, composite it; otherwise, render the preview using the temporary object
            If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then
                If forceWhiteBackground Then tmpLayer.compositeBackgroundColor 255, 255, 255 Else tmpLayer.compositeBackgroundColor
            End If
            
            tmpLayer.renderToPictureBox dstPicture
            
        End If
        
    Else
    
        If otherPictureSrc.getLayerColorDepth = 32 Then
            Set tmpLayer = New pdLayer
            tmpLayer.createFromExistingLayer otherPictureSrc, newWidth, newHeight, True
            If forceWhiteBackground Then tmpLayer.compositeBackgroundColor 255, 255, 255 Else tmpLayer.compositeBackgroundColor
            tmpLayer.renderToPictureBox dstPicture
        Else
            otherPictureSrc.renderToPictureBox dstPicture
        End If
        
    End If
    
End Sub

'Draw a gradient from Color1 to Color 2 (RGB longs) on a specified picture box
Public Sub DrawGradient(ByVal DstPicBox As Object, ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal drawHorizontal As Boolean = False)

    'Calculation variables (used to interpolate between the gradient colors)
    Dim VR As Single, VG As Single, VB As Single
    Dim x As Long, y As Long
    
    'Red, green, and blue variables for each gradient color
    Dim r As Long, g As Long, b As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    
    'Extract the red, green, and blue values from the gradient colors (which were passed as Longs)
    r = ExtractR(Color1)
    g = ExtractG(Color1)
    b = ExtractB(Color1)
    r2 = ExtractR(Color2)
    g2 = ExtractG(Color2)
    b2 = ExtractB(Color2)
    
    'Width and height variables are faster than repeated access of .ScaleWidth/Height properties
    Dim tmpHeight As Long
    Dim tmpWidth As Long
    tmpWidth = DstPicBox.ScaleWidth
    tmpHeight = DstPicBox.ScaleHeight

    'Create a calculation variable, which will be used to determine the interpolation step between
    ' each gradient color
    If drawHorizontal Then
        VR = Abs(r - r2) / tmpWidth
        VG = Abs(g - g2) / tmpWidth
        VB = Abs(b - b2) / tmpWidth
    Else
        VR = Abs(r - r2) / tmpHeight
        VG = Abs(g - g2) / tmpHeight
        VB = Abs(b - b2) / tmpHeight
    End If
    
    'If the second color is less than the first value, make the step negative
    If r2 < r Then VR = -VR
    If g2 < g Then VG = -VG
    If b2 < b Then VB = -VB
    
    'Run a loop across the picture box, changing the gradient color according to the step calculated earlier
    If drawHorizontal Then
        For x = 0 To tmpWidth
            r2 = r + VR * x
            g2 = g + VG * x
            b2 = b + VB * x
            DstPicBox.Line (x, 0)-(x, tmpHeight), RGB(r2, g2, b2)
        Next x
    Else
        For y = 0 To tmpHeight
            r2 = r + VR * y
            g2 = g + VG * y
            b2 = b + VB * y
            DstPicBox.Line (0, y)-(tmpWidth, y), RGB(r2, g2, b2)
        Next y
    End If
    
End Sub
