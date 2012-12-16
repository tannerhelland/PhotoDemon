Attribute VB_Name = "Misc_Uncategorized"
'***************************************************************************
'Miscellaneous Operations Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 03/October/12
'Last update: Reorganized this massive module into a bunch of smaller ones for improved organization.
'
'If a function doesn't have a home in a more appropriate module, it gets stuck here.  Over time, I'm
' hoping to clear out most of this module in favor of a more organized approach.
'
'***************************************************************************

Option Explicit

'Distance value for mouse_over events and selections; a literal "radius" below which the mouse cursor is considered "over" a point
Private Const mouseSelAccuracy As Single = 8

'Convert a system color (such as "button face" or "inactive window") to a literal RGB value
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal oColor As OLE_COLOR, ByVal HPALETTE As Long, ByRef cColorRef As Long) As Long

'Convert a width and height pair to a new max width and height, while preserving aspect ratio
Public Sub convertAspectRatio(ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef newWidth As Long, ByRef newHeight As Long)
    
    Dim srcAspect As Single, dstAspect As Single
    srcAspect = srcWidth / srcHeight
    dstAspect = dstWidth / dstHeight
    
    If srcAspect > dstAspect Then
        newWidth = dstWidth
        newHeight = CSng(srcHeight / srcWidth) * newWidth + 0.5
    Else
        newHeight = dstHeight
        newWidth = CSng(srcWidth / srcHeight) * newHeight + 0.5
    End If

End Sub

'Given the number of colors in an image (as supplied by getQuickColorCount, below), return the highest color depth
' that includes all those colors and is supported by PhotoDemon (1/4/8/24/32)
Public Function getColorDepthFromColorCount(ByVal srcColors As Long, ByRef refLayer As pdLayer) As Long
    
    If srcColors <= 256 Then
        If srcColors > 16 Then
            getColorDepthFromColorCount = 8
        Else
            
            'FreeImage only supports the writing of 4bpp and 1bpp images if they are grayscale.  Thus, only
            ' mark images as 4bpp or 1bpp if they are gray/b&w - otherwise, consider them 8bpp indexed color.
            If (srcColors > 2) Then
                                
                If g_IsImageGray Then
                    getColorDepthFromColorCount = 4
                Else
                    getColorDepthFromColorCount = 8
                End If
                
            Else
                If g_IsImageGray Then
                    getColorDepthFromColorCount = 1
                Else
                    getColorDepthFromColorCount = 8
                End If
            End If
            
        End If
    Else
        If refLayer.getLayerColorDepth = 24 Then
            getColorDepthFromColorCount = 24
        Else
            getColorDepthFromColorCount = 32
        End If
    End If

End Function

'When images are loaded, this function is used to quickly determine the image's color count.  It stops once 257 is reached,
' as at that point the program will automatically treat the image as 24 or 32bpp (contingent on presence of an alpha channel).
Public Function getQuickColorCount(ByVal srcImage As pdImage) As Long
    
    Message "Verifying image color count..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcImage.mainLayer
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = srcImage.Width - 1
    finalY = srcImage.Height - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = (srcImage.mainLayer.getLayerColorDepth) \ 8
    
    'This array will track whether or not a given color has been detected in the image.  (I don't know if powers of two
    ' are allocated more efficiently, but it doesn't hurt to stick to that rule.)
    Dim UniqueColors() As Long
    ReDim UniqueColors(0 To 511) As Long
    
    Dim i As Long
    For i = 0 To 255
        UniqueColors(i) = -1
    Next i
    
    'Total number of unique colors counted so far
    Dim totalCount As Long
    totalCount = 0
    
    'Finally, a bunch of variables used in color calculation
    Dim R As Long, g As Long, b As Long
    Dim chkValue As Long
    Dim colorFound As Boolean
        
    'Apply the filter
    For x = 0 To finalX
        QuickVal = x * qvDepth
    For y = 0 To finalY
        
        R = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        chkValue = RGB(R, g, b)
        colorFound = False
        
        'Now, loop through the colors we've accumulated thus far and compare this entry against each of them.
        For i = 0 To totalCount
            If UniqueColors(i) = chkValue Then
                colorFound = True
                Exit For
            End If
        Next i
        
        'If colorFound is still false, store this value in the array and increment our color counter
        If Not colorFound Then
            UniqueColors(totalCount) = chkValue
            totalCount = totalCount + 1
            'ReDim Preserve UniqueColors(0 To totalCount) As Long
        End If
        
        'If the image has more than 256 colors, treat it as 24/32 bpp
        If totalCount > 256 Then Exit For
        
    Next y
        If totalCount > 256 Then Exit For
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Also, erase the counting array
    'Erase UniqueColors
    
    'If we've made it this far, the color count has a maximum value of 257.
    ' If it is less than 257, analyze it to see if it contains all gray values.
    If totalCount <= 256 Then
    
        g_IsImageGray = True
    
        'Loop through all available colors
        For i = 0 To totalCount - 1
        
            R = ExtractR(UniqueColors(i))
            g = ExtractG(UniqueColors(i))
            b = ExtractB(UniqueColors(i))
            
            'If any of the components do not match, this is not a grayscale image
            If (R <> g) Or (g <> b) Or (R <> b) Then
                g_IsImageGray = False
                Exit For
            End If
            
        Next i
    
    'If the image contains more than 256 colors, it is not grayscale
    Else
        g_IsImageGray = False
    End If
    
    getQuickColorCount = totalCount
    
End Function

'Given an OLE color, return an RGB
Public Function ConvertSystemColor(ByVal colorRef As OLE_COLOR) As Long
    
    'OleTranslateColor returns -1 if it fails; if that happens, default to white
    If OleTranslateColor(colorRef, 0, ConvertSystemColor) Then
        ConvertSystemColor = RGB(255, 255, 255)
    End If
    
End Function

'Populate a text box with a given integer value.  This is done constantly across the program, so I use a sub to handle it, as
' there may be additional validations that need to be performed, and it's nice to be able to adjust those from a single location.
Public Sub copyToTextBoxI(ByRef dstTextBox As TextBox, ByVal srcValue As Long)

    'Remember the current cursor position
    Dim cursorPos As Long
    cursorPos = dstTextBox.SelStart

    'Overwrite the current text box value with the new value
    dstTextBox = CStr(srcValue)
    dstTextBox.Refresh
    
    'Restore the cursor to its original position
    If cursorPos >= Len(dstTextBox) Then cursorPos = Len(dstTextBox)
    dstTextBox.SelStart = cursorPos

End Sub

'Populate a text box with a given floating-point value.  This is done constantly across the program, so I use a sub to handle it, as
' there may be additional validations that need to be performed, and it's nice to be able to adjust those from a single location.
Public Sub copyToTextBoxF(ByVal srcValue As Double, ByRef dstTextBox As TextBox)

    'Remember the current cursor position
    Dim cursorPos As Long
    cursorPos = dstTextBox.SelStart

    'PhotoDemon never allows more than two significant digits for floating-point text boxes
    dstTextBox = Format(CStr(srcValue), "#0.00")
    dstTextBox.Refresh
    
    'Restore the cursor to its original position
    If cursorPos >= Len(dstTextBox) Then cursorPos = Len(dstTextBox)
    dstTextBox.SelStart = cursorPos

End Sub

'Let a form know whether the mouse pointer is over its image or just the viewport
Public Function isMouseOverImage(ByVal x1 As Long, ByVal y1 As Long, ByRef srcForm As Form) As Boolean

    If (x1 >= pdImages(srcForm.Tag).targetLeft) And (x1 <= pdImages(srcForm.Tag).targetLeft + pdImages(srcForm.Tag).targetWidth) Then
        If (y1 >= pdImages(srcForm.Tag).targetTop) And (y1 <= pdImages(srcForm.Tag).targetTop + pdImages(srcForm.Tag).targetHeight) Then
            isMouseOverImage = True
            Exit Function
        Else
            isMouseOverImage = False
        End If
        isMouseOverImage = False
    End If

End Function

'Calculate and display the current mouse position.
' INPUTS: x and y coordinates of the mouse cursor, current form, and optionally two long-type variables to receive the relative
'          coordinates (e.g. location on the image) of the current mouse position.
Public Sub displayImageCoordinates(ByVal x1 As Single, ByVal y1 As Single, ByRef srcForm As Form, Optional ByRef copyX As Single, Optional ByRef copyY As Single)

    If isMouseOverImage(x1, y1, srcForm) Then
            
        'Grab the current zoom value
        Static ZoomVal As Single
        ZoomVal = Zoom.ZoomArray(pdImages(srcForm.Tag).CurrentZoomValue)
            
        'Calculate x and y positions, while taking into account zoom and scroll values
        x1 = srcForm.HScroll.Value + Int((x1 - pdImages(srcForm.Tag).targetLeft) / ZoomVal)
        y1 = srcForm.VScroll.Value + Int((y1 - pdImages(srcForm.Tag).targetTop) / ZoomVal)
            
        'When zoomed very far out, the values might be calculated incorrectly.  Force them to the image dimensions if necessary.
        If x1 < 0 Then x1 = 0
        If y1 < 0 Then y1 = 0
        If x1 > pdImages(srcForm.Tag).Width Then x1 = pdImages(srcForm.Tag).Width
        If y1 > pdImages(srcForm.Tag).Height Then y1 = pdImages(srcForm.Tag).Height
        
        'If the user has requested copies of these coordinates, assign them now
        If copyX Then copyX = x1
        If copyY Then copyY = y1
        
        FormMain.lblCoordinates.Caption = "(" & x1 & "," & y1 & ")"
        FormMain.lblCoordinates.Refresh
        'DoEvents
        
    End If
    
End Sub

'If an x or y location is NOT in the image, find the nearest coordinate that IS in the image
Public Sub findNearestImageCoordinates(ByRef x1 As Single, ByRef y1 As Single, ByRef srcForm As Form)

    'Grab the current zoom value
    Static ZoomVal As Single
    ZoomVal = Zoom.ZoomArray(pdImages(srcForm.Tag).CurrentZoomValue)

    'Calculate x and y positions, while taking into account zoom and scroll values
    x1 = srcForm.HScroll.Value + Int((x1 - pdImages(srcForm.Tag).targetLeft) / ZoomVal)
    y1 = srcForm.VScroll.Value + Int((y1 - pdImages(srcForm.Tag).targetTop) / ZoomVal)
    
    'Force any invalid values to their nearest matching point in the image
    If x1 < 0 Then x1 = 0
    If y1 < 0 Then y1 = 0
    If x1 >= pdImages(srcForm.Tag).Width Then x1 = pdImages(srcForm.Tag).Width - 1
    If y1 >= pdImages(srcForm.Tag).Height Then y1 = pdImages(srcForm.Tag).Height - 1

End Sub

'This sub will return a constant correlating to the nearest selection point.  Its return values are:
' 0 - Cursor is not near a selection point
' 1 - NW corner
' 2 - NE corner
' 3 - SE corner
' 4 - SW corner
' 5 - N edge
' 6 - E edge
' 7 - S edge
' 8 - W edge
' 9 - interior of selection, not near a corner or edge
Public Function findNearestSelectionCoordinates(ByRef x1 As Single, ByRef y1 As Single, ByRef srcForm As Form) As Long

    'Grab the current zoom value
    Static ZoomVal As Single
    ZoomVal = Zoom.ZoomArray(pdImages(srcForm.Tag).CurrentZoomValue)

    'Calculate x and y positions, while taking into account zoom and scroll values
    x1 = srcForm.HScroll.Value + Int((x1 - pdImages(srcForm.Tag).targetLeft) / ZoomVal)
    y1 = srcForm.VScroll.Value + Int((y1 - pdImages(srcForm.Tag).targetTop) / ZoomVal)
    
    'Force any invalid values to their nearest matching point in the image
    If x1 < 0 Then x1 = 0
    If y1 < 0 Then y1 = 0
    If x1 > pdImages(srcForm.Tag).Width Then x1 = pdImages(srcForm.Tag).Width
    If y1 > pdImages(srcForm.Tag).Height Then y1 = pdImages(srcForm.Tag).Height

    'With x1 and y1 now representative of a location within the image, it's time to start calculating distances.
    Static tLeft As Single, tTop As Single, tRight As Single, tBottom As Single
    tLeft = pdImages(srcForm.Tag).mainSelection.selLeft
    tTop = pdImages(srcForm.Tag).mainSelection.selTop
    tRight = pdImages(srcForm.Tag).mainSelection.selLeft + pdImages(srcForm.Tag).mainSelection.selWidth
    tBottom = pdImages(srcForm.Tag).mainSelection.selTop + pdImages(srcForm.Tag).mainSelection.selHeight
    
    'Adjust the mouseAccuracy value based on the current zoom value
    Static mouseAccuracy As Single
    mouseAccuracy = mouseSelAccuracy * (1 / ZoomVal)
    
    'Before doing anything else, make sure the pointer is actually worth checking - e.g. make sure it's near the selection
    If (x1 < tLeft - mouseAccuracy) Or (x1 > tRight + mouseAccuracy) Or (y1 < tTop - mouseAccuracy) Or (y1 > tBottom + mouseAccuracy) Then
        findNearestSelectionCoordinates = 0
        Exit Function
    End If
    
    'If we made it here, this mouse location is worth evaluating.  Corners get preference, so check them first.
    Static nwDist As Single, neDist As Single, seDist As Single, swDist As Single
    
    nwDist = distanceTwoPoints(x1, y1, tLeft, tTop)
    neDist = distanceTwoPoints(x1, y1, tRight, tTop)
    swDist = distanceTwoPoints(x1, y1, tLeft, tBottom)
    seDist = distanceTwoPoints(x1, y1, tRight, tBottom)
    
    'Find the smallest distance for this mouse position
    Static minDistance As Single
    Static closestPoint As Long
    minDistance = mouseAccuracy
    closestPoint = -1
    
    If nwDist <= minDistance Then
        minDistance = nwDist
        closestPoint = 1
    End If
    
    If neDist <= minDistance Then
        minDistance = neDist
        closestPoint = 2
    End If
    
    If seDist <= minDistance Then
        minDistance = seDist
        closestPoint = 3
    End If
    
    If swDist <= minDistance Then
        minDistance = swDist
        closestPoint = 4
    End If
    
    'Was a close point found?  If yes, then return that value
    If closestPoint <> -1 Then
        findNearestSelectionCoordinates = closestPoint
        Exit Function
    End If

    'If we're at this line of code, a closest corner was not found.  So check edges next.
    Static nDist As Single, eDist As Single, sDist As Single, wDist As Single
    
    nDist = distanceOneDimension(y1, tTop)
    eDist = distanceOneDimension(x1, tRight)
    sDist = distanceOneDimension(y1, tBottom)
    wDist = distanceOneDimension(x1, tLeft)
    
    If (nDist <= minDistance) Then
        minDistance = nDist
        closestPoint = 5
    End If
    
    If (eDist <= minDistance) Then
        minDistance = eDist
        closestPoint = 6
    End If
    
    If (sDist <= minDistance) Then
        minDistance = sDist
        closestPoint = 7
    End If
    
    If (wDist <= minDistance) Then
        minDistance = wDist
        closestPoint = 8
    End If
    
    'Was a close point found?  If yes, then return that value.
    If closestPoint <> -1 Then
        findNearestSelectionCoordinates = closestPoint
        Exit Function
    End If

    'If we're at this line of code, a closest edge was not found.  Perform one final check to ensure that the mouse is within the
    ' image's boundaries, and if it is, return the "move selection" ID, then exit.
    If (x1 > tLeft) And (x1 < tRight) And (y1 > tTop) And (y1 < tBottom) Then
        findNearestSelectionCoordinates = 9
    Else
        findNearestSelectionCoordinates = 0
    End If

End Function

'Return the distance between two values on the same line
Public Function distanceOneDimension(ByVal x1 As Single, ByVal x2 As Single) As Single
    distanceOneDimension = Sqr((x1 - x2) ^ 2)
End Function

'Return the distance between two points
Public Function distanceTwoPoints(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Single
    distanceTwoPoints = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function

'Extract the red, green, or blue value from an RGB() Long
Public Function ExtractR(ByVal CurrentColor As Long) As Integer
    ExtractR = CurrentColor Mod 256
End Function

Public Function ExtractG(ByVal CurrentColor As Long) As Integer
    ExtractG = (CurrentColor \ 256) And 255
End Function

Public Function ExtractB(ByVal CurrentColor As Long) As Integer
    ExtractB = (CurrentColor \ 65536) And 255
End Function

'Blend byte1 w/ byte2 based on mixRatio.  mixRatio is expected to be a value between 0 and 1.
Public Function BlendColors(ByVal Color1 As Byte, ByVal Color2 As Byte, ByRef mixRatio As Single) As Byte
    BlendColors = ((1 - mixRatio) * Color1) + (mixRatio * Color2)
End Function
