Attribute VB_Name = "Drawing"
'***************************************************************************
'PhotoDemon Drawing Routines
'Copyright 2001-2015 by Tanner Helland
'Created: 4/3/01
'Last updated: 01/December/12
'Last update: Added DrawSystemIcon function (previously used for only the "unsaved changes" dialog
'
'Miscellaneous drawing routines that don't fit elsewhere.  At present, this includes rendering preview images,
' drawing the canvas background of image forms, and a gradient-rendering sub (used primarily on the histogram form).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The following Enum and two API declarations are used to draw the system information icon
Public Enum SystemIconConstants
    IDI_APPLICATION = 32512
    IDI_HAND = 32513
    IDI_QUESTION = 32514
    IDI_EXCLAMATION = 32515
    IDI_ASTERISK = 32516
    IDI_WINDOWS = 32517
End Enum

Private Declare Function LoadIconByID Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

'GDI drawing functions
Private Const PS_SOLID = 0
Private Const PS_DASH = 1
Private Const PS_DOT = 2
Private Const PS_DASHDOT = 3
Private Const PS_DASHDOTDOT = 4

Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5

Private Const HS_DIAGCROSS = 5

Private Const NULL_BRUSH = 5

Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (ByVal dibPointer As Long, ByVal iUsage As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal pointerToRectOfOldCoords As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal targetDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal targetDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, ByVal refToPeviousPoint As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long

'DC API functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'API for converting between hWnd-specific coordinate spaces.  Note that the function technically accepts an
' array of POINTAPI points; the address passed to lpPoints should be the address of the first point in the array
' (e.g. ByRef PointArray(0)), while the cPoints parameter is the number of points in the array.  If two points are
' passed, a special Rect transform may occur on RtL systems; see http://msdn.microsoft.com/en-us/library/dd145046%28v=vs.85%29.aspx
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lpPoints As POINTAPI, ByVal cPoints As Long) As Long

'Given a target picture box, draw a hue preview across the horizontal axis.  This is helpful for tools that provide
' a hue slider, so that the user can easily find a color of their choosing.  Optionally, saturation and luminance
' can be provided, though it's generally assumed that those values will both be 1.0.
Public Sub drawHueBox_HSV(ByRef dstPic As PictureBox, Optional ByVal dstSaturation As Double = 1, Optional ByVal dstLuminance As Double = 1)

    'Retrieve the picture box's dimensions
    Dim picWidth As Long, picHeight As Long
    picWidth = dstPic.ScaleWidth
    picHeight = dstPic.ScaleHeight
    
    'Use a DIB to hold the hue box before we render it on-screen.  Why?  So we can color-manage it, of course!
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createBlank picWidth, picHeight, 24, 0
    
    Dim tmpR As Double, tmpG As Double, tmpB As Double
    
    'From left-to-right, draw a full hue range onto the DIB
    Dim x As Long
    For x = 0 To tmpDIB.getDIBWidth - 1
        fHSVtoRGB x / tmpDIB.getDIBWidth, dstSaturation, dstLuminance, tmpR, tmpG, tmpB
        drawLineToDC tmpDIB.getDIBDC, x, 0, x, picHeight, RGB(tmpR * 255, tmpG * 255, tmpB * 255)
    Next x
    
    'With the hue box complete, render it onto the destination picture box, with color management applied
    tmpDIB.renderToPictureBox dstPic

End Sub

'Given a target picture box, draw a saturation preview across the horizontal axis.  This is helpful for tools that provide
' a saturation slider, so that the user can easily find a color of their choosing.  Optionally, hue and luminance can be
' provided - hue is STRONGLY recommended, but luminance can safely be assumed to be 1.0 (in most cases).
Public Sub drawSaturationBox_HSV(ByRef dstPic As PictureBox, Optional ByVal dstHue As Double = 1, Optional ByVal dstLuminance As Double = 1)

    'Retrieve the picture box's dimensions
    Dim picWidth As Long, picHeight As Long
    picWidth = dstPic.ScaleWidth
    picHeight = dstPic.ScaleHeight
    
    'Use a DIB to hold the hue box before we render it on-screen.  Why?  So we can color-manage it, of course!
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createBlank picWidth, picHeight, 24, 0
    
    Dim tmpR As Double, tmpG As Double, tmpB As Double
    
    'From left-to-right, draw a full hue range onto the DIB
    Dim x As Long
    For x = 0 To tmpDIB.getDIBWidth - 1
        fHSVtoRGB dstHue, x / tmpDIB.getDIBWidth, dstLuminance, tmpR, tmpG, tmpB
        drawLineToDC tmpDIB.getDIBDC, x, 0, x, picHeight, RGB(tmpR * 255, tmpG * 255, tmpB * 255)
    Next x
    
    'With the hue box complete, render it onto the destination picture box, with color management applied
    tmpDIB.renderToPictureBox dstPic

End Sub

'Basic wrapper to line-drawing via GDI
Public Sub drawLineToDC(ByVal targetDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal crColor As Long)

    'Create a pen with the specified color
    Dim newPen As Long
    newPen = CreatePen(PS_SOLID, 1, crColor)
    
    'Select the pen into the target DC
    Dim oldObject As Long
    oldObject = SelectObject(targetDC, newPen)
    
    'Render the line
    MoveToEx targetDC, x1, y1, 0&
    LineTo targetDC, x2, y2
    
    'Remove the pen and delete it
    SelectObject targetDC, oldObject
    DeleteObject newPen

End Sub

'Basic wrappers for rect-filling and rect-tracing via GDI
Public Sub fillRectToDC(ByVal targetDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal crColor As Long)

    'Create a brush with the specified color
    Dim tmpBrush As Long
    tmpBrush = CreateSolidBrush(crColor)
    
    'Select the brush into the target DC
    Dim oldObject As Long
    oldObject = SelectObject(targetDC, tmpBrush)
    
    'Fill the rect
    Rectangle targetDC, x1, y1, x2, y2
    
    'Remove the brush and delete it
    SelectObject targetDC, oldObject
    DeleteObject tmpBrush

End Sub

Public Sub outlineRectToDC(ByVal targetDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal crColor As Long)
    drawLineToDC targetDC, x1, y1, x2, y1, crColor
    drawLineToDC targetDC, x2, y1, x2, y2, crColor
    drawLineToDC targetDC, x2, y2, x1, y2, crColor
    drawLineToDC targetDC, x1, y2, x1, y1, crColor
End Sub

'Draw a system icon on the specified device context; this code is adopted from an example by Francesco Balena at http://www.devx.com/vb2themax/Tip/19108
Public Sub DrawSystemIcon(ByVal icon As SystemIconConstants, ByVal hDC As Long, ByVal x As Long, ByVal y As Long)
    Dim hIcon As Long
    hIcon = LoadIconByID(0, icon)
    DrawIcon hDC, x, y, hIcon
End Sub

'Used to draw the main image onto a preview picture box
Public Sub DrawPreviewImage(ByRef dstPicture As PictureBox, Optional ByVal useOtherPictureSrc As Boolean = False, Optional ByRef otherPictureSrc As pdDIB, Optional forceWhiteBackground As Boolean = False)
    
    Dim tmpDIB As pdDIB
    
    'Start by calculating the aspect ratio of both the current image and the previewing picture box
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = dstPicture.ScaleWidth
    dstHeight = dstPicture.ScaleHeight
    
    Dim srcWidth As Double, srcHeight As Double
    
    'The source values need to be adjusted contingent on whether this is a selection or a full-image preview
    If useOtherPictureSrc Then
        srcWidth = otherPictureSrc.getDIBWidth
        srcHeight = otherPictureSrc.getDIBHeight
    Else
        If pdImages(g_CurrentImage).selectionActive Then
            srcWidth = pdImages(g_CurrentImage).mainSelection.boundWidth
            srcHeight = pdImages(g_CurrentImage).mainSelection.boundHeight
        Else
            srcWidth = pdImages(g_CurrentImage).getActiveDIB().getDIBWidth
            srcHeight = pdImages(g_CurrentImage).getActiveDIB().getDIBHeight
        End If
    End If
            
    'Now, use that aspect ratio to determine a proper size for our temporary DIB
    Dim newWidth As Long, newHeight As Long
    
    convertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
    
    'Normally this will draw a preview of pdImages(g_CurrentImage).containingForm's relevant image.  However, another picture source can be specified.
    If Not useOtherPictureSrc Then
        
        'Check to see if a selection is active; if it isn't, simply render the full form
        If Not pdImages(g_CurrentImage).selectionActive Then
        
            If pdImages(g_CurrentImage).getActiveDIB().getDIBColorDepth = 32 Then
                Set tmpDIB = New pdDIB
                tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).getActiveDIB(), newWidth, newHeight, True
                If forceWhiteBackground Then tmpDIB.compositeBackgroundColor 255, 255, 255
                tmpDIB.renderToPictureBox dstPicture
            Else
                pdImages(g_CurrentImage).getActiveDIB().renderToPictureBox dstPicture
            End If
        
        Else
        
            'Copy the current selection into a temporary DIB
            Set tmpDIB = New pdDIB
            tmpDIB.createBlank pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).getActiveDIB().getDIBColorDepth
            BitBlt tmpDIB.getDIBDC, 0, 0, pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).getActiveDIB().getDIBDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, vbSrcCopy
        
            'If the image is transparent, composite it; otherwise, render the preview using the temporary object
            If pdImages(g_CurrentImage).getActiveDIB().getDIBColorDepth = 32 Then
                If forceWhiteBackground Then tmpDIB.compositeBackgroundColor 255, 255, 255
            End If
            
            tmpDIB.renderToPictureBox dstPicture
            
        End If
        
    Else
    
        If otherPictureSrc.getDIBColorDepth = 32 Then
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB otherPictureSrc, newWidth, newHeight, True
            If forceWhiteBackground Then tmpDIB.compositeBackgroundColor 255, 255, 255
            tmpDIB.renderToPictureBox dstPicture
        Else
            otherPictureSrc.renderToPictureBox dstPicture
        End If
        
    End If
    
End Sub

'Draw a gradient from Color1 to Color 2 (RGB longs) on a specified picture box
Public Sub DrawGradient(ByVal DstPicBox As Object, ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal drawHorizontal As Boolean = False)

    'Calculation variables (used to interpolate between the gradient colors)
    Dim vR As Double, vG As Double, vB As Double
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
        vR = Abs(r - r2) / tmpWidth
        vG = Abs(g - g2) / tmpWidth
        vB = Abs(b - b2) / tmpWidth
    Else
        vR = Abs(r - r2) / tmpHeight
        vG = Abs(g - g2) / tmpHeight
        vB = Abs(b - b2) / tmpHeight
    End If
    
    'If a component of the right color is less than the matching component of the left color, make the step negative
    If r2 < r Then vR = -vR
    If g2 < g Then vG = -vG
    If b2 < b Then vB = -vB
    
    'Run a loop across the picture box, changing the gradient color according to the step calculated earlier
    If drawHorizontal Then
        For x = 0 To tmpWidth
            r2 = r + vR * x
            g2 = g + vG * x
            b2 = b + vB * x
            DstPicBox.Line (x, 0)-(x, tmpHeight), RGB(r2, g2, b2)
        Next x
    Else
        For y = 0 To tmpHeight
            r2 = r + vR * y
            g2 = g + vG * y
            b2 = b + vB * y
            DstPicBox.Line (0, y)-(tmpWidth, y), RGB(r2, g2, b2)
        Next y
    End If
    
End Sub

'Draw a horizontal gradient to a specified DIB from x-position xLeft to xRigth,
' using ColorLeft and ColorRight (RGB longs) as the gradient endpoints.
Public Sub DrawHorizontalGradientToDIB(ByVal dstDIB As pdDIB, ByVal xLeft As Long, ByVal xRight As Long, ByVal colorLeft As Long, ByVal colorRight As Long)
    
    Dim x As Long
    
    'Red, green, and blue variables for each gradient color
    Dim rLeft As Long, gLeft As Long, bLeft As Long
    Dim rRight As Long, gRight As Long, Bright As Long
    
    'Extract the red, green, and blue values from the gradient colors (which were passed as Longs)
    rLeft = ExtractR(colorLeft)
    gLeft = ExtractG(colorLeft)
    bLeft = ExtractB(colorLeft)
    rRight = ExtractR(colorRight)
    gRight = ExtractG(colorRight)
    Bright = ExtractB(colorRight)
    
    'Calculate a width for the gradient area
    Dim gradWidth As Long
    gradWidth = xRight - xLeft
    
    Dim blendRatio As Double
    Dim newR As Byte, newG As Byte, newB As Byte
    
    '32bpp DIBs need to use GDI+ instead of GDI, to make sure the alpha channel is supported
    Dim alphaMatters As Boolean
    If dstDIB.getDIBColorDepth = 32 Then alphaMatters = True Else alphaMatters = False
    
    'If alpha is relevant, cache a GDI+ image handle and pen in advance
    Dim hGdipImage As Long, hGdipPen As Long
    If alphaMatters Then hGdipImage = GDI_Plus.getGDIPlusGraphicsFromDC(dstDIB.getDIBDC, False)
    
    'Run a loop across the DIB, changing the gradient color according to the step calculated earlier
    For x = xLeft To xRight
        
        'Calculate a blend ratio for this position
        blendRatio = (x - xLeft) / gradWidth
        
        'Calculate blendd RGB values for this position
        newR = BlendColors(rLeft, rRight, blendRatio)
        newG = BlendColors(gLeft, gRight, blendRatio)
        newB = BlendColors(bLeft, Bright, blendRatio)
        
        'Draw a vertical line at this position, using the calculated color
        If alphaMatters Then
        
            hGdipPen = GDI_Plus.getGDIPlusPenHandle(RGB(newR, newG, newB), 255, 1, LineCapFlat)
            GDI_Plus.GDIPlusDrawLine_Fast hGdipImage, hGdipPen, x, 0, x, dstDIB.getDIBHeight
            GDI_Plus.releaseGDIPlusPen hGdipPen
        
        Else
            drawLineToDC dstDIB.getDIBDC, x, 0, x, dstDIB.getDIBHeight, RGB(newR, newG, newB)
        End If
        
    Next x
    
    'Release our GDI+ handle, if any
    If alphaMatters Then GDI_Plus.releaseGDIPlusGraphics hGdipImage
    
End Sub

'Given a source DIB, fill it with a 2x2 alpha checkerboard pattern matching the user's current preferences.
' (The resulting DIB size is contingent on the user's checkerboard pattern size preference, FYI.)
Public Sub createAlphaCheckerboardDIB(ByRef srcDIB As pdDIB)

    'Retrieve the user's preferred alpha checkerboard colors, and convert the longs into individual RGB components
    Dim chkColorOne As Long, chkColorTwo As Long
    chkColorOne = g_UserPreferences.GetPref_Long("Transparency", "AlphaCheckOne", RGB(255, 255, 255))
    chkColorTwo = g_UserPreferences.GetPref_Long("Transparency", "AlphaCheckTwo", RGB(204, 204, 204))
    
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = ExtractR(chkColorOne)
    r2 = ExtractR(chkColorTwo)
    g1 = ExtractG(chkColorOne)
    g2 = ExtractG(chkColorTwo)
    b1 = ExtractB(chkColorOne)
    b2 = ExtractB(chkColorTwo)
    
    'Determine a checkerboard block size based on the current user preference
    Dim chkSize As Long
    chkSize = g_UserPreferences.GetPref_Long("Transparency", "AlphaCheckSize", 1)
    
    Select Case chkSize
    
        'Small (4x4 checks)
        Case 0
            chkSize = 4
            
        'Medium (8x8 checks)
        Case 1
            chkSize = 8
        
        'Large (16x16 checks)
        Case Else
            chkSize = 16
        
    End Select
    
    'Resize the source DIB to fit a 2x2 block pattern of the requested checkerboard pattern
    srcDIB.createBlank chkSize * 2, chkSize * 2
    
    'Point a temporary array directly at the source DIB's bitmap bits.
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Fill the source DIB with the checkerboard pattern
    Dim x As Long, y As Long, QuickX As Long
    For x = 0 To srcDIB.getDIBWidth - 1
        QuickX = x * 3
    For y = 0 To srcDIB.getDIBHeight - 1
         
        If (((x \ chkSize) + (y \ chkSize)) And 1) = 0 Then
            srcImageData(QuickX + 2, y) = r1
            srcImageData(QuickX + 1, y) = g1
            srcImageData(QuickX, y) = b1
        Else
            srcImageData(QuickX + 2, y) = r2
            srcImageData(QuickX + 1, y) = g2
            srcImageData(QuickX, y) = b2
        End If
        
    Next y
    Next x
    
    'Release our temporary array and exit
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData

End Sub

'Given a source DIB, fill it with the alpha checkerboard pattern.  32bpp images can then be alpha blended onto it.
Public Sub fillDIBWithAlphaCheckerboard(ByRef srcDIB As pdDIB, ByVal x1 As Long, ByVal y1 As Long, ByVal bltWidth As Long, ByVal bltHeight As Long)

    'Create a pattern brush from the public checkerboard image
    Dim hCheckerboard As Long
    hCheckerboard = CreatePatternBrush(g_CheckerboardPattern.getDIBHandle)
    
    'Select the brush into the target DIB's DC
    Dim hOldBrush As Long
    hOldBrush = SelectObject(srcDIB.getDIBDC, hCheckerboard)
    
    'Paint the DIB
    SetBrushOrgEx srcDIB.getDIBDC, x1, y1, 0&
    PatBlt srcDIB.getDIBDC, x1, y1, bltWidth, bltHeight, vbPatCopy
    
    'Remove and delete the brush
    SelectObject srcDIB.getDIBDC, hOldBrush
    DeleteObject hCheckerboard

End Sub

'Given an (x,y) pair on the current viewport, convert the value to coordinates on the image.
Public Function convertCanvasCoordsToImageCoords(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Double, ByVal canvasY As Double, ByRef imgX As Double, ByRef imgY As Double, Optional ByVal forceInBounds As Boolean = False) As Boolean

    If Not (srcImage.imgViewport Is Nothing) Then
    
        'Get the current zoom value from the source image
        Dim zoomVal As Double
        zoomVal = g_Zoom.getZoomValue(srcImage.currentZoomValue)
        
        'Get a copy of the translated image rect, in canvas coordinates.  If the canvas is a window, and the zoomed
        ' image is a poster sliding around behind it, the translate image rect contains the poster coordinates,
        ' relative to the window.  What's great about this rect is that it's already accounted for scroll bars,
        ' so we can ignore their value(s) here.
        Dim translatedImageRect As RECTF
        srcImage.imgViewport.getImageRectTranslated translatedImageRect
        
        'Translating the canvas coordinate pair back to the image is now easy.  Subtract the top/left offset,
        ' then divide by zoom - that's all there is to it!
        imgX = (canvasX - translatedImageRect.Left) / zoomVal
        imgY = (canvasY - translatedImageRect.Top) / zoomVal
        
        'If the caller wants the coordinates bound-checked, apply it now
        If forceInBounds Then
            If imgX < 0 Then imgX = 0
            If imgY < 0 Then imgY = 0
            If imgX >= srcImage.Width - 1 Then imgX = srcImage.Width - 1
            If imgY >= srcImage.Height - 1 Then imgY = srcImage.Height - 1
        End If
        
        convertCanvasCoordsToImageCoords = True
        
    Else
        convertCanvasCoordsToImageCoords = False
    End If
    
End Function

'Given an (x,y) pair on the current image, convert the value to coordinates on the current viewport canvas.
Public Sub convertImageCoordsToCanvasCoords(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal imgX As Double, ByVal imgY As Double, ByRef canvasX As Double, ByRef canvasY As Double, Optional ByVal forceInBounds As Boolean = False)

    If Not (srcImage.imgViewport Is Nothing) Then
    
        'Get the current zoom value from the source image
        Dim zoomVal As Double
        zoomVal = g_Zoom.getZoomValue(srcImage.currentZoomValue)
            
        'Get a copy of the translated image rect, in canvas coordinates.  If the canvas is a window, and the zoomed
        ' image is a poster sliding around behind it, the translate image rect contains the poster coordinates,
        ' relative to the window.  What's great about this rect is that it's already accounted for scroll bars,
        ' so we can ignore their value(s) here.
        Dim translatedImageRect As RECTF
        srcImage.imgViewport.getImageRectTranslated translatedImageRect
        
        'Translating the canvas coordinate pair back to the image is now easy.  Add the top/left offset,
        ' then multiply by zoom - that's all there is to it!
        canvasX = (imgX * zoomVal) + translatedImageRect.Left
        canvasY = (imgY * zoomVal) + translatedImageRect.Top
        
        'If the caller wants the coordinates bound-checked, apply it now
        If forceInBounds Then
        
            'Get a copy of the current viewport intersection rect, which determines bounds of this function
            Dim vIntersectRect As RECTF
            srcImage.imgViewport.getIntersectRectCanvas vIntersectRect
            
            If canvasX < vIntersectRect.Left Then canvasX = vIntersectRect.Left
            If canvasY < vIntersectRect.Top Then canvasY = vIntersectRect.Top
            If canvasX >= vIntersectRect.Left + vIntersectRect.Width Then canvasX = vIntersectRect.Left + vIntersectRect.Width - 1
            If canvasY >= vIntersectRect.Top + vIntersectRect.Height Then canvasY = vIntersectRect.Top + vIntersectRect.Height - 1
            
        End If
        
    End If
    
End Sub

'Given an (x,y) pair on the current image, convert the value to coordinates relative to the current layer.  This is especially relevant
' if the layer has one or more non-destructive affine transforms active.
Public Function convertImageCoordsToLayerCoords(ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByVal imgX As Single, ByVal imgY As Single, ByRef layerX As Single, ByRef layerY As Single) As Boolean

    If srcImage Is Nothing Then Exit Function
    If srcLayer Is Nothing Then Exit Function
    
    'If the layer has one or more active affine transforms, this step becomes complicated.
    If srcLayer.affineTransformsActive(False) Then
    
        'Create a copy of either the layer's transformation matrix, or a custom matrix if passed in
        Dim tmpMatrix As pdGraphicsMatrix
        srcLayer.getCopyOfLayerTransformationMatrix tmpMatrix
        
        'Invert the matrix
        If tmpMatrix.InvertMatrix() Then
            
            'We now need to convert the image coordinates against the layer transformation matrix
            tmpMatrix.applyMatrixToXYPair imgX, imgY
            
            'In order for the matrix conversion to work, it has to offset coordinates by the current layer offset.  (Rotation is
            ' particularly important in that regard, as the center-point is crucial.)  As such, we now need to undo that translation.
            ' In rare circumstances the caller can disable this behavior, for example while transforming a layer, because the original
            ' rotation matrix must be used.
            layerX = imgX + srcLayer.getLayerOffsetX
            layerY = imgY + srcLayer.getLayerOffsetY
            
            convertImageCoordsToLayerCoords = True
        
        'If we can't invert the matrix, we're in trouble.  Copy out the layer coordinates as a failsafe.
        Else
            
            layerX = imgX
            layerY = imgY
            
            Debug.Print "WARNING! Transformation matrix could not be generated."
            
            convertImageCoordsToLayerCoords = False
            
        End If
    
    'If the layer doesn't have affine transforms active, this step is easy.
    Else
    
        'Layer coordinates are identical to image coordinates
        layerX = imgX
        layerY = imgY
        
        convertImageCoordsToLayerCoords = True
    
    End If
    
End Function

'Given an array of (x,y) pairs set in the current image's coordinate space, convert each pair to the supplied viewport canvas space.
Public Sub convertListOfImageCoordsToCanvasCoords(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef listOfPoints() As POINTFLOAT, Optional ByVal forceInBounds As Boolean = False)

    If srcImage.imgViewport Is Nothing Then Exit Sub
    
    'Get the current zoom value from the source image
    Dim zoomVal As Double
    zoomVal = g_Zoom.getZoomValue(srcImage.currentZoomValue)
    
    'Get a copy of the translated image rect, in canvas coordinates.  If the canvas is a window, and the zoomed
    ' image is a poster sliding around behind it, the translate image rect contains the poster coordinates,
    ' relative to the window.  What's great about this rect is that it's already accounted for scroll bars,
    ' so we can ignore their value(s) here.
    Dim translatedImageRect As RECTF
    srcImage.imgViewport.getImageRectTranslated translatedImageRect
    
    'If the caller wants the coordinates bound-checked, we also need to grab a copy of the viewport
    ' intersection rect, which controls boundaries
    Dim vIntersectRect As RECTF
    If forceInBounds Then srcImage.imgViewport.getIntersectRectCanvas vIntersectRect
    
    Dim canvasX As Double, canvasY As Double
    
    'Iterate through each point in turn; note that bounds are automatically detected, and there is not currently a way to override
    ' this behavior.
    Dim i As Long
    For i = LBound(listOfPoints) To UBound(listOfPoints)
        
        'Translating the canvas coordinate pair back to the image is now easy.  Add the top/left offset,
        ' then multiply by zoom - that's all there is to it!
        canvasX = (listOfPoints(i).x * zoomVal) + translatedImageRect.Left
        canvasY = (listOfPoints(i).y * zoomVal) + translatedImageRect.Top
        
        'If the caller wants the coordinates bound-checked, apply it now
        If forceInBounds Then
            If canvasX < vIntersectRect.Left Then canvasX = vIntersectRect.Left
            If canvasY < vIntersectRect.Top Then canvasY = vIntersectRect.Top
            If canvasX >= vIntersectRect.Left + vIntersectRect.Width Then canvasX = vIntersectRect.Left + vIntersectRect.Width - 1
            If canvasY >= vIntersectRect.Top + vIntersectRect.Height Then canvasY = vIntersectRect.Top + vIntersectRect.Height - 1
        End If
        
        'Store the updated coordinate pair
        listOfPoints(i).x = canvasX
        listOfPoints(i).y = canvasY
    
    Next i
        
End Sub

'Given a source hWnd and a destination hWnd, translate a coordinate pair between their unique coordinate spaces.  Note that
' the screen coordinate space will be used as an intermediary in the conversion.
Public Sub convertCoordsBetweenHwnds(ByVal srcHwnd As Long, ByVal dstHwnd As Long, ByVal srcX As Long, ByVal srcY As Long, ByRef dstX As Long, ByRef dstY As Long)
    
    'The API we're using require POINTAPI structs
    Dim tmpPoint As POINTAPI
    
    With tmpPoint
        .x = srcX
        .y = srcY
    End With
    
    'Transform the coordinates
    MapWindowPoints srcHwnd, dstHwnd, tmpPoint, 1
    
    'Report the transformed points back to the user
    dstX = tmpPoint.x
    dstY = tmpPoint.y
    
End Sub

'Given a specific layer, return a RECT filled with that layer's corner coordinates -
' IN THE CANVAS COORDINATE SPACE (hence the function name).
Public Sub getCanvasRectForLayer(ByVal layerIndex As Long, ByRef dstRect As RECT, Optional ByVal useCanvasModifiers As Boolean = False)

    Dim tmpX As Double, tmpY As Double
    
    With pdImages(g_CurrentImage).getLayerByIndex(layerIndex)
        
        'Start with the top-left corner
        convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), .getLayerOffsetX, .getLayerOffsetY, tmpX, tmpY
        dstRect.Left = tmpX
        dstRect.Top = tmpY
        
        'End with the bottom-right corner
        convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), .getLayerOffsetX + .getLayerWidth(useCanvasModifiers), .getLayerOffsetY + .getLayerHeight(useCanvasModifiers), tmpX, tmpY
        
        'Because layers support sub-pixel positioning, but the canvas rect renderer *does not*, we must manually align the right/bottom coords
        dstRect.Right = Int(tmpX + 0.99)
        dstRect.Bottom = Int(tmpY + 0.99)
        
    End With

End Sub

'Same as above, but using floating-point values
Public Sub getCanvasRectForLayerF(ByVal layerIndex As Long, ByRef dstRect As RECTF, Optional ByVal useCanvasModifiers As Boolean = False)

    Dim tmpX As Double, tmpY As Double
    
    With pdImages(g_CurrentImage).getLayerByIndex(layerIndex)
        
        'Start with the top-left corner
        convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), .getLayerOffsetX, .getLayerOffsetY, tmpX, tmpY
        dstRect.Left = tmpX
        dstRect.Top = tmpY
        
        'End with the bottom-right corner
        convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), .getLayerOffsetX + .getLayerWidth(useCanvasModifiers), .getLayerOffsetY + .getLayerHeight(useCanvasModifiers), tmpX, tmpY
        dstRect.Width = tmpX - dstRect.Left
        dstRect.Height = tmpY - dstRect.Top
        
    End With

End Sub

'On the current viewport, render lines around the active layer
Public Sub drawLayerBoundaries(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer)

    'In the old days, we could get away with assuming layer boundaries form a rectangle, but as of PD 7.0, affine transforms
    ' mean this is no longer guaranteed.
    '
    'So instead of filling a rect, we must retrieve the four layer corner coordinates as floating-point pairs.
    Dim layerCorners() As POINTFLOAT
    ReDim layerCorners(0 To 3) As POINTFLOAT
    
    srcLayer.getLayerCornerCoordinates layerCorners, True, False
    
    'Next, convert each corner from image coordinate space to the active viewport coordinate space
    Drawing.convertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, layerCorners, False
    
    'Pass the list of coordinates to a pdGraphicsPath object; it will handle the actual UI rendering
    Dim tmpPath As pdGraphicsPath
    Set tmpPath = New pdGraphicsPath
    
    'Note that we must add the layer boundary lines manually - otherwise, the top-right and bottom-left corners will connect
    ' due to the way srcLayer.getLayerCornerCoordinates returns the points!
    tmpPath.addLine layerCorners(0).x, layerCorners(0).y, layerCorners(1).x, layerCorners(1).y
    tmpPath.addLine layerCorners(1).x, layerCorners(1).y, layerCorners(3).x, layerCorners(3).y
    tmpPath.addLine layerCorners(3).x, layerCorners(3).y, layerCorners(2).x, layerCorners(2).y
    tmpPath.addLine layerCorners(2).x, layerCorners(2).y, layerCorners(0).x, layerCorners(0).y
    
    'Render the final UI
    tmpPath.strokePathToDIB_UIStyle Nothing, dstCanvas.hDC
    
End Sub

'On the current viewport, render standard PD transformation nodes (layer corners, currently) atop the active layer.
Public Sub drawLayerCornerNodes(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, Optional ByVal curPOI As Long = -1)

    'In the old days, we could get away with assuming layer boundaries form a rectangle, but as of PD 7.0, affine transforms
    ' mean this is no longer guaranteed.
    '
    'So instead of filling a rect, we must retrieve the four layer corner coordinates as floating-point pairs.
    Dim layerCorners() As POINTFLOAT
    ReDim layerCorners(0 To 3) As POINTFLOAT
    
    srcLayer.getLayerCornerCoordinates layerCorners, True, False
    
    'Next, convert each corner from image coordinate space to the active viewport coordinate space
    Drawing.convertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, layerCorners, False
    
    Dim circRadius As Long, circAlpha As Long
    circRadius = 7
    circAlpha = 190
    
    Dim dstDC As Long
    dstDC = dstCanvas.hDC
    
    'Use GDI+ to render four corner circles
    Dim i As Long
    For i = 0 To 3
        GDI_Plus.GDIPlusDrawCanvasSquare dstDC, layerCorners(i).x, layerCorners(i).y, circRadius, circAlpha, CBool(i = curPOI)
    Next i
    
End Sub

'As of PD 7.0, on-canvas rotation is now supported.  Use this function to render the current rotation node.
Public Sub drawLayerRotateNode(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, Optional ByVal curPOI As Long = -1)
    
    'Retrieve the layer rotate node position from the specified layer, and convert it into the canvas coordinate space
    Dim layerRotateNodes() As POINTFLOAT
    ReDim layerRotateNodes(0 To 4) As POINTFLOAT
    
    Dim rotateUIRect As RECTF
    srcLayer.getLayerRotationNodeCoordinates layerRotateNodes, True
    Drawing.convertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, layerRotateNodes, False
    
    'Render the circles
    Dim circRadius As Long, circAlpha As Long
    circRadius = 7
    circAlpha = 190
    
    Dim dstDC As Long
    dstDC = dstCanvas.hDC
    
    Dim i As Long
    For i = 1 To 4
        GDIPlusDrawCanvasCircle dstDC, layerRotateNodes(i).x, layerRotateNodes(i).y, circRadius, circAlpha, CBool(curPOI = i + 3)
    Next i
    
    'As a convenience to the user, we also draw some additional UI features if a rotation node is actively hovered by the mouse.
    If (curPOI >= 4) And (curPOI <= 7) Then
        
        Dim relevantPoint As Long
        relevantPoint = curPOI - 3
        
        'First, draw a line from the center of the layer to the rotation node, to provide visual feedback on where the rotation
        ' will actually occur.
        Dim tmpPath As pdGraphicsPath
        Set tmpPath = New pdGraphicsPath
        tmpPath.addLine layerRotateNodes(0).x, layerRotateNodes(0).y, layerRotateNodes(relevantPoint).x, layerRotateNodes(relevantPoint).y
        tmpPath.strokePathToDIB_UIStyle Nothing, dstDC
        
        'Next, we are going to draw an arc with arrows on the end, to display where the actual rotation will occur.
        ' (At present, we skip this step if shearing is active, as I haven't figured out how to correctly skew the arc into the
        '  proper on-screen coordinate space.)
        If (srcLayer.getLayerShearX = 0) And (srcLayer.getLayerShearY = 0) Then
            
            tmpPath.resetPath
        
            'Start by finding the distance of the rotation line.
            Dim rRadius As Double
            rRadius = Math_Functions.distanceTwoPoints(layerRotateNodes(0).x, layerRotateNodes(0).y, layerRotateNodes(relevantPoint).x, layerRotateNodes(relevantPoint).y)
            
            'From there, bounds are easy-peasy
            Dim rotateBoundRect As RECTF
            With rotateBoundRect
                .Left = layerRotateNodes(0).x - rRadius
                .Top = layerRotateNodes(0).y - rRadius
                .Width = rRadius * 2
                .Height = rRadius * 2
            End With
            
            'Arc sweep and arc length are inter-related.  What we ultimately want is a (roughly) equal arc size regardless of zoom or
            ' the underlying image size.  This is difficult to predict as larger images and/or higher zoom will result in larger arc widths
            ' for an identical radius.  As such, we hard-code an approximate arc length, then generate an arc sweep from it.
            '
            'In my testing, 80-ish pixels is a reasonably good size across many image dimensions.  Note that we *do* correct for DPI here.
            Dim arcLength As Double
            arcLength = FixDPIFloat(70)
            
            'Switching between arc length and sweep is easy; see https://en.wikipedia.org/wiki/Arc_%28geometry%29#Length_of_an_arc_of_a_circle
            Dim arcSweep As Double
            arcSweep = (arcLength * 180) / (PI * rRadius)
            
            'Make sure the arc fits within a valid range (e.g. no complete circles or nearly-straight lines)
            If arcSweep > 90 Then arcSweep = 90
            If arcSweep < 30 Then arcSweep = 30
            
            'We need to modify the default layer angle depending on the current POI
            Dim relevantAngle As Double
            relevantAngle = srcLayer.getLayerAngle + ((relevantPoint - 1) * 90)
            
            tmpPath.addArc rotateBoundRect, relevantAngle - arcSweep / 2, arcSweep
            tmpPath.strokePathToDIB_UIStyle Nothing, dstDC, False, True, , LineCapArrowAnchor, LineCapArrowAnchor
            
        End If
        
    End If
    
End Sub

'Need a quick and dirty DC for something?  Call this.  (Just remember to free the DC when you're done!)
Public Function GetMemoryDC() As Long
    
    GetMemoryDC = CreateCompatibleDC(0&)
    
    'In debug mode, track how many DCs the program requests
    #If DEBUGMODE = 1 Then
        If GetMemoryDC <> 0 Then
            g_DCsCreated = g_DCsCreated + 1
        Else
            pdDebug.LogAction "WARNING!  Drawing.GetMemoryDC() failed to create a new memory DC!"
        End If
    #End If
    
End Function

Public Sub FreeMemoryDC(ByVal srcDC As Long)
    
    If srcDC <> 0 Then
        
        Dim delConfirm As Long
        delConfirm = DeleteDC(srcDC)
    
        'In debug mode, track how many DCs the program frees
        #If DEBUGMODE = 1 Then
            If delConfirm <> 0 Then
                g_DCsDestroyed = g_DCsDestroyed + 1
            Else
                pdDebug.LogAction "WARNING!  Drawing.FreeMemoryDC() failed to release DC #" & srcDC & "."
            End If
        #End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  Drawing.FreeMemoryDC() was passed a null DC.  Fix this!"
        #End If
    End If
    
End Sub
