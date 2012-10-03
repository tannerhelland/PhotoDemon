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
    SrcWidth = pdImages(CurrentImage).mainLayer.getLayerWidth
    SrcHeight = pdImages(CurrentImage).mainLayer.getLayerHeight
    
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
        
        If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then
            Set tmpLayer = New pdLayer
            tmpLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, newWidth, newHeight, True
            If forceWhiteBackground Then tmpLayer.compositeBackgroundColor 255, 255, 255 Else tmpLayer.compositeBackgroundColor
            tmpLayer.renderToPictureBox dstPicture
        Else
            pdImages(CurrentImage).mainLayer.renderToPictureBox dstPicture
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

'A simple routine to draw the canvas background; the public CanvasBackground variable is used to determine
' draw mode: -1 is a checkerboard effect, any other value is treated as an RGB long
Public Sub DrawSpecificCanvas(ByRef dstForm As Form)

    '-1 indicates the user wants a checkerboard background pattern
    If CanvasBackground = -1 Then

        Dim stepIntervalX As Long, stepIntervalY As Long
        stepIntervalX = dstForm.PicCH.ScaleWidth
        stepIntervalY = dstForm.PicCH.ScaleHeight
            
        Dim x As Long, y As Long
        Dim srchDC As Long
        srchDC = dstForm.PicCH.hDC
            
        For x = 0 To pdImages(dstForm.Tag).backBuffer.getLayerWidth Step stepIntervalX
        For y = 0 To pdImages(dstForm.Tag).backBuffer.getLayerHeight Step stepIntervalY
            BitBlt pdImages(dstForm.Tag).backBuffer.getLayerDC, x, y, stepIntervalX, stepIntervalY, srchDC, 0, 0, vbSrcCopy
        Next y
        Next x
    
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
