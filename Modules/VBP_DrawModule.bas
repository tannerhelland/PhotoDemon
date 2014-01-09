Attribute VB_Name = "Drawing"
'***************************************************************************
'PhotoDemon Drawing Routines
'Copyright ©2001-2014 by Tanner Helland
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
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

'GDI drawing functions
Private Const PS_SOLID As Long = &H0
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal pointerToRectOfOldCoords As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (ByVal dibPointer As Long, ByVal iUsage As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal targetDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal targetDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, ByVal refToPeviousPoint As Long) As Long

'If a layer is smaller than a picture box set to receive it, use this function
Public Function getDrawCoordinatesForSmallLayer(ByRef srcLayer As pdLayer, ByRef dstObject As Object, ByRef drawX As Long, ByRef drawY As Long) As Boolean

End Function

'Basic wrapper to line-drawing via the API
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

'Simplified pure-VB function for rendering text to an object.
Public Sub drawTextOnObject(ByRef dstObject As Object, ByVal sText As String, ByVal xPos As Long, ByVal yPos As Long, Optional ByVal newFontSize As Long = 12, Optional ByVal newFontColor As Long = 0, Optional ByVal makeFontBold As Boolean = False, Optional ByVal makeFontItalic As Boolean = False)

    dstObject.CurrentX = xPos
    dstObject.CurrentY = yPos
    dstObject.FontSize = newFontSize
    dstObject.ForeColor = newFontColor
    dstObject.FontBold = makeFontBold
    dstObject.FontItalic = makeFontItalic
    dstObject.Print sText

End Sub

'Draw a system icon on the specified device context; this code is adopted from an example by Francesco Balena at http://www.devx.com/vb2themax/Tip/19108
Public Sub DrawSystemIcon(ByVal icon As SystemIconConstants, ByVal hDC As Long, ByVal X As Long, ByVal Y As Long)
    Dim hIcon As Long
    hIcon = LoadIconByID(0, icon)
    DrawIcon hDC, X, Y, hIcon
End Sub

'Used to draw the main image onto a preview picture box
Public Sub DrawPreviewImage(ByRef dstPicture As PictureBox, Optional ByVal useOtherPictureSrc As Boolean = False, Optional ByRef otherPictureSrc As pdLayer, Optional forceWhiteBackground As Boolean = False)
    
    Dim tmpLayer As pdLayer
    
    'Start by calculating the aspect ratio of both the current image and the previewing picture box
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = dstPicture.ScaleWidth
    dstHeight = dstPicture.ScaleHeight
    
    Dim srcWidth As Double, srcHeight As Double
    
    'The source values need to be adjusted contingent on whether this is a selection or a full-image preview
    If useOtherPictureSrc Then
        srcWidth = otherPictureSrc.getLayerWidth
        srcHeight = otherPictureSrc.getLayerHeight
    Else
        If pdImages(g_CurrentImage).selectionActive Then
            srcWidth = pdImages(g_CurrentImage).mainSelection.boundWidth
            srcHeight = pdImages(g_CurrentImage).mainSelection.boundHeight
        Else
            srcWidth = pdImages(g_CurrentImage).getActiveLayer().getLayerWidth
            srcHeight = pdImages(g_CurrentImage).getActiveLayer().getLayerHeight
        End If
    End If
            
    'Now, use that aspect ratio to determine a proper size for our temporary layer
    Dim newWidth As Long, newHeight As Long
    
    convertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
    
    'Normally this will draw a preview of pdImages(g_CurrentImage).containingForm's relevant image.  However, another picture source can be specified.
    If Not useOtherPictureSrc Then
        
        'Check to see if a selection is active; if it isn't, simply render the full form
        If Not pdImages(g_CurrentImage).selectionActive Then
        
            If pdImages(g_CurrentImage).getActiveLayer().getLayerColorDepth = 32 Then
                Set tmpLayer = New pdLayer
                tmpLayer.createFromExistingLayer pdImages(g_CurrentImage).getActiveLayer(), newWidth, newHeight, True
                If forceWhiteBackground Then tmpLayer.compositeBackgroundColor 255, 255, 255
                tmpLayer.renderToPictureBox dstPicture
            Else
                pdImages(g_CurrentImage).getActiveLayer().renderToPictureBox dstPicture
            End If
        
        Else
        
            'Copy the current selection into a temporary layer
            Set tmpLayer = New pdLayer
            tmpLayer.createBlank pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).getActiveLayer().getLayerColorDepth
            BitBlt tmpLayer.getLayerDC, 0, 0, pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).getActiveLayer().getLayerDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, vbSrcCopy
        
            'If the image is transparent, composite it; otherwise, render the preview using the temporary object
            If pdImages(g_CurrentImage).getActiveLayer().getLayerColorDepth = 32 Then
                If forceWhiteBackground Then tmpLayer.compositeBackgroundColor 255, 255, 255
            End If
            
            tmpLayer.renderToPictureBox dstPicture
            
        End If
        
    Else
    
        If otherPictureSrc.getLayerColorDepth = 32 Then
            Set tmpLayer = New pdLayer
            tmpLayer.createFromExistingLayer otherPictureSrc, newWidth, newHeight, True
            If forceWhiteBackground Then tmpLayer.compositeBackgroundColor 255, 255, 255
            tmpLayer.renderToPictureBox dstPicture
        Else
            otherPictureSrc.renderToPictureBox dstPicture
        End If
        
    End If
    
End Sub

'Draw a gradient from Color1 to Color 2 (RGB longs) on a specified picture box
Public Sub DrawGradient(ByVal DstPicBox As Object, ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal drawHorizontal As Boolean = False)

    'Calculation variables (used to interpolate between the gradient colors)
    Dim VR As Double, VG As Double, VB As Double
    Dim X As Long, Y As Long
    
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
        For X = 0 To tmpWidth
            r2 = r + VR * X
            g2 = g + VG * X
            b2 = b + VB * X
            DstPicBox.Line (X, 0)-(X, tmpHeight), RGB(r2, g2, b2)
        Next X
    Else
        For Y = 0 To tmpHeight
            r2 = r + VR * Y
            g2 = g + VG * Y
            b2 = b + VB * Y
            DstPicBox.Line (0, Y)-(tmpWidth, Y), RGB(r2, g2, b2)
        Next Y
    End If
    
End Sub

'Given a source layer, fill it with a 2x2 alpha checkerboard pattern matching the user's current preferences.
' (The resulting layer size is contingent on the user's checkerboard pattern size preference, FYI.)
Public Sub createAlphaCheckerboardLayer(ByRef srcLayer As pdLayer)

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
    
    'Resize the source layer to fit a 2x2 block pattern of the requested checkerboard pattern
    srcLayer.createBlank chkSize * 2, chkSize * 2
    
    'Point a temporary array directly at the source layer's bitmap bits.
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Fill the source layer with the checkerboard pattern
    Dim X As Long, Y As Long, QuickX As Long
    For X = 0 To srcLayer.getLayerWidth - 1
        QuickX = X * 3
    For Y = 0 To srcLayer.getLayerHeight - 1
         
        If (((X \ chkSize) + (Y \ chkSize)) And 1) = 0 Then
            srcImageData(QuickX + 2, Y) = r1
            srcImageData(QuickX + 1, Y) = g1
            srcImageData(QuickX, Y) = b1
        Else
            srcImageData(QuickX + 2, Y) = r2
            srcImageData(QuickX + 1, Y) = g2
            srcImageData(QuickX, Y) = b2
        End If
        
    Next Y
    Next X
    
    'Release our temporary array and exit
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData

End Sub

'Given a source layer, fill it with the alpha checkerboard pattern.  32bpp images can then be alpha blended onto it.
Public Sub fillLayerWithAlphaCheckerboard(ByRef srcLayer As pdLayer, ByVal x1 As Long, ByVal y1 As Long, ByVal bltWidth As Long, ByVal bltHeight As Long)

    'Create a pattern brush from the public checkerboard image
    Dim hCheckerboard As Long
    hCheckerboard = CreatePatternBrush(g_CheckerboardPattern.getLayerDIB)
    
    'Select the brush into the target layer's DC
    Dim hOldBrush As Long
    hOldBrush = SelectObject(srcLayer.getLayerDC, hCheckerboard)
    
    'Paint the layer
    SetBrushOrgEx srcLayer.getLayerDC, x1, y1, 0&
    PatBlt srcLayer.getLayerDC, x1, y1, bltWidth, bltHeight, vbPatCopy
    
    'Remove and delete the brush
    SelectObject srcLayer.getLayerDC, hOldBrush
    DeleteObject hCheckerboard

End Sub
