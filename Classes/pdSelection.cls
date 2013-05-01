VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Selection class
'Copyright ©2011-2013 by Tanner Helland
'Created: 25/September/12
'Last updated: 03/October/12
'Last update: added a "Red" highlight draw option
'
'This class handles all selections in PhotoDemon. At present, selections are simply rectangular areas - but this will change
' in the future as new selection modes are added.
'
'All selections are treated as a subset of pixels within a rectangular area. The x1, y1, x2, y2 coordinates define opposite
' corners of the selection area. There is no guarantee that these coordinates will appear in a given corner - so they must
' be checked manually. When a selection is set, those values are translated into Left, Top, Width and Height values, which
' are in turn used for everything from cropping to applying filters.
'
'Note also that this class is treated as a subset of pdLayer(). Right now each layer object only contains one selection,
' but there's no reason it couldn't store multiples in the future. (Although I'm not sure what purpose that would serve.)
'
'***************************************************************************

Option Explicit

'SafeArray types for pointing VB arrays at arbitrary memory locations (in our case, bitmap data)
Private Type SAFEARRAYBOUND
    cElements As Long
    lBound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(1) As SAFEARRAYBOUND
End Type

'SafeArray API functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal byteLength As Long)
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

'DC API functions
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

'Object API functions
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Pen drawing
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

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

'Rectangle drawing
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

'Brush creation
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'Stock object retrieval
Private Const NULL_BRUSH = 5
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

'RasterOp selection
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long

'What type of selection is this?
Public Enum SelectionType
    sRectangle = 0
    sSquare = 1
    sCircle = 2
End Enum

Private sType As SelectionType

'The coordinates of the current selection (if rectangular)
Private x1 As Double, x2 As Double
Private y1 As Double, y2 As Double

'The actual coordinates passed from the user's mouse movements to this selection object.  These coordinates are not modified in any way.
' (e.g. these coordinates could lie outside the image, viewport, or screen, because no bounds checking is applied to them)
Private actualX1 As Double, actualX2 As Double
Private actualY1 As Double, actualY2 As Double

'When a selection is "locked in", the x and y values are converted to these values
Public selLeft As Long, selTop As Long
Public selWidth As Long, selHeight As Long

'Is this selection "locked in"?
Private lockedIn As Boolean

'When the selection is moved, it's necessary to know the difference between the current mouse point and the original left / top values
Private moveXDist As Double, moveYDist As Double

'Is this selection requesting text box updates? Because if it is, ignore external requests to update.
Public rejectRefreshRequests As Boolean

'What is the current transformation mode? (This is important for dragging to resize the selection.)
Private selTransformationType As Long

'Is transformation mode active?
Private transformModeActive As Boolean

'What image does this selection belong to?
Public containingPDImage As pdImage

'This DIB contains the selection mask for this selection object.  Black pixels in the selection mask represent unselected pixels in the image.
' White pixels represent selected ones.  Non-white and non-black pixels can be used for aliasing.
Public selMask As pdLayer

'After the mask has been generated, its status will be set as "ready".  Any function that changes the selected area will reset this to "false".
' The selection rendering code can check this value to see if the mask needs to be recreated before rendering it to screen.
Private isMaskReady As Boolean

'set a transformation type
Public Sub setTransformationType(ByVal transType As Long)
    selTransformationType = transType
End Sub

'Set a selection type
Public Sub setSelectionType(ByVal selType As SelectionType)
    sType = selType
End Sub

'Takes x and y coordinates (from a _MouseDown event, typically) and uses them in a manner specified by the current transform operation.
' Note that this this should only be called after a transformation type has been set.
Public Sub setInitialTransformCoordinates(ByVal x As Double, y As Double)

    'If new transform coordinates are being set, this selection must be "unlocked"
    lockedIn = False
    
    'The use of setInitialCoordinates means this IS a transformation
    transformModeActive = True
    
    'Based on the transform mode, set the initial points accordingly
    Select Case selTransformationType
    
        'Case 0 should never occur, but if it does - treat this like a normal initial coordinate call
        Case 0
            x1 = x
            y1 = y
            x2 = x
            y2 = y
        
        '1 - NW corner
        Case 1
            x1 = selLeft + selWidth
            y1 = selTop + selHeight
            x2 = selLeft
            y2 = selTop
        
        '2 - NE corner
        Case 2
            x1 = selLeft
            y1 = selTop + selHeight
            x2 = x
            y2 = y
        
        '3 - SE corner
        Case 3
            x1 = selLeft
            y1 = selTop
            x2 = x
            y2 = y
        
        '4 - SW corner
        Case 4
            x1 = selLeft + selWidth
            y1 = selTop
            x2 = x
            y2 = y
        
        '5 - N edge
        Case 5
            x1 = selLeft
            x2 = selLeft + selWidth
            y1 = selTop + selHeight
            y2 = y
        
        '6 - E edge
        Case 6
            x1 = selLeft
            x2 = x
            y1 = selTop
            y2 = selTop + selHeight
        
        '7 - S edge
        Case 7
            x1 = selLeft
            x2 = selLeft + selWidth
            y1 = selTop
            y2 = y
        
        '8 - W edge
        Case 8
            x1 = selLeft + selWidth
            x2 = x
            y1 = selTop
            y2 = selTop + selHeight
        
        '9 - interior of selection, not near a corner or edge
        Case 9
            moveXDist = x - selLeft
            moveYDist = y - selTop
    
    End Select
    
    updateInternalCoords
    
End Sub

'Takes x and y coordinates (from a _MouseDown event, typically) and stores them internally
Public Sub setInitialCoordinates(ByVal x As Double, y As Double)
    
    'If new initial coordinates are being set, this selection must be "unlocked"
    lockedIn = False
    
    'The use of setInitialCoordinates means this is not a transformation
    transformModeActive = False
    
    x1 = x
    y1 = y
    
    'Make a back-up copy of the coordinates.  These values will never be changed except by the user moving the mouse.
    actualX1 = x
    actualY1 = y
    
    If sType = sRectangle Then
        x2 = x
        y2 = y
    End If
    
    updateInternalCoords
    
End Sub

'Takes x and y coordinates (from a _MouseMove event, typically) and stores them internally
Public Sub setAdditionalCoordinates(ByVal x As Double, y As Double)
    
    'Check for an active transformation mode.  (A transformation is something like resizing or moving an existing selection,
    ' versus drawing a new one from scratch.)
    If transformModeActive Then
    
        'Based on the transform mode, set the initial points accordingly
        Select Case selTransformationType
        
            'Cases 0 should never occur, but if it does - treat this like a normal subsequent coordinate call
            'Cases 1-4 can similarly be treated like a normal subsequent coordinate call
            ' (See setInitialTransformCoordinates for details on what the different cases mean.  And yes, I should convert
            '  these to constants instead of numbers... :P)
            Case 0 To 4
                x2 = x
                y2 = y
    
                updateInternalCoords
            
            '5 - N edge
            Case 5
                y2 = y
                updateInternalCoords
            
            '6 - E edge
            Case 6
                x2 = x
                updateInternalCoords
            
            '7 - S edge
            Case 7
                y2 = y
                updateInternalCoords
            
            '8 - W edge
            Case 8
                x2 = x
                updateInternalCoords
            
            '9 - interior of selection, not near a corner or edge (e.g. move the selection, but don't resize it)
            Case 9
            
                'When moving the selection, force it to retain its current size.  This means the selection cannot be moved
                ' off the image - it must remain completely on it, at its current size.
                x1 = x - moveXDist
                If x1 < 0 Then x1 = 0
                If x1 + selWidth > containingPDImage.Width Then x1 = containingPDImage.Width - selWidth
                
                y1 = y - moveYDist
                If y1 < 0 Then y1 = 0
                If y1 + selHeight > containingPDImage.Height Then y1 = containingPDImage.Height - selHeight
                
                x2 = x1 + selWidth
                y2 = y1 + selHeight
                                
                updateInternalCoords
                'Message x1 & "," & y1 & " <--> " & x2 & "," & y2 & " | " & selLeft & "," & selTop & " <--> " & selWidth & "," & selHeight
        
        End Select
    
    'No transform mode makes this easy; just update the second set of points and redraw the selection
    Else
    
        x2 = x
        y2 = y
        
        'Make a back-up copy of the coordinates.  These values will never be changed except by the user moving the mouse.
        actualX2 = x
        actualY2 = y
    
        updateInternalCoords
        
    End If
    
End Sub

'Has this selection been locked in?
Public Function isLockedIn() As Boolean
    isLockedIn = lockedIn
End Function

'Whenever the x1, y1, x2, y2 values are changed, this sub needs to be called to update the left/right/width/height values accordingly
Private Sub updateInternalCoords()

    'Mark the selection mask as "not ready", as it will need to be redrawn after new coordinates are set
    isMaskReady = False

    'Calculate left, top width, height based off the current x1/x2/y1/y2 values
    If x1 < x2 Then
        If x1 < 0 Then x1 = 0
        If x2 > containingPDImage.Width Then x2 = containingPDImage.Width
        selLeft = x1
        selWidth = x2 - x1
    Else
        If x2 < 0 Then x2 = 0
        If x1 > containingPDImage.Width Then x1 = containingPDImage.Width
        selLeft = x2
        selWidth = x1 - x2
    End If
    
    If y1 < y2 Then
        If y1 < 0 Then y1 = 0
        If y2 > containingPDImage.Height Then y2 = containingPDImage.Height
        selTop = y1
        selHeight = y2 - y1
    Else
        If y2 < 0 Then y2 = 0
        If y1 > containingPDImage.Height Then y1 = containingPDImage.Height
        selTop = y2
        selHeight = y1 - y2
    End If
    
    'Perform a quick sanity check on the calculated variables
    If selLeft < 0 Then selLeft = 0
    If selTop < 0 Then selTop = 0
    If selLeft + selWidth > containingPDImage.Width Then selLeft = containingPDImage.Width - selWidth
    If selTop + selHeight > containingPDImage.Height Then selTop = containingPDImage.Height - selHeight
    
    'Adjust the main form's selection text boxes to match the newly calculated values
    refreshTextBoxes

End Sub

'Because selections can be created off the image, it is sometimes necessary to check if ALL selection coordinates lie off the
' image.  If this is the case, we don't want to finalize the current selection - we want to forget about it.
Public Function areAllCoordinatesInvalid() As Boolean

    'Transforms (resizing, moving, etc) do their own validation.  As such, this function is unnecessary because transforms
    ' can only result in a valid selection.
    If transformModeActive Then
        areAllCoordinatesInvalid = False
    
    'If this selection wasn't just adjusted by a transform, it may be invalid (for example, if all its coordinates lie off the image).
    ' Check for validity now.
    Else
        
        'If both coordinates lie outside the image, there is still a chance the selection itself can be valid (for example,
        ' if the selection extends the full width or height of the image).  Check for this special case now.
        If (actualX1 < 0) And (actualX2 < 0) Then areAllCoordinatesInvalid = True
        If (actualX1 > containingPDImage.Width) And (actualX2 > containingPDImage.Width) Then areAllCoordinatesInvalid = True
        If (actualY1 < 0) And (actualY2 < 0) Then areAllCoordinatesInvalid = True
        If (actualY1 > containingPDImage.Height) And (actualY2 > containingPDImage.Height) Then areAllCoordinatesInvalid = True
        
    End If

End Function

'Use this to populate the text boxes on the main form with the current selection values
Public Sub refreshTextBoxes()

    rejectRefreshRequests = True

    If selLeft >= 0 Then FormMain.tudSelLeft.Value = selLeft
    If selTop >= 0 Then FormMain.tudSelTop.Value = selTop
    If selWidth > 0 Then FormMain.tudSelWidth.Value = selWidth
    If selHeight > 0 Then FormMain.tudSelHeight.Value = selHeight

    rejectRefreshRequests = False

End Sub

'Update this selection using the values in the main form's selection text boxes
Public Sub updateViaTextBox()

    'Ignore text box update requests until the selection is locked in
    If isLockedIn = False Then Exit Sub

    rejectRefreshRequests = True

    'Mark the selection mask as "not ready", as it will need to be redrawn after new coordinates are set
    isMaskReady = False

    'Check all text box entries for validity, then update the corresponding selection values
    If FormMain.tudSelLeft.IsValid(False) Then
        If FormMain.tudSelLeft.Value > containingPDImage.Width - 1 Then FormMain.tudSelLeft = containingPDImage.Width - 1
        selLeft = FormMain.tudSelLeft
    End If
    If FormMain.tudSelTop.IsValid(False) Then
        If FormMain.tudSelTop > containingPDImage.Height - 1 Then FormMain.tudSelTop = containingPDImage.Height - 1
        selTop = FormMain.tudSelTop
    End If
    If FormMain.tudSelWidth.IsValid(False) Then
        If FormMain.tudSelWidth + FormMain.tudSelLeft > containingPDImage.Width Then
            'FormMain.txtSelWidth = pdImages(CurrentImage).Width - Val(FormMain.txtSelLeft)
            selLeft = containingPDImage.Width - FormMain.tudSelWidth
            FormMain.tudSelLeft = selLeft
        End If
        selWidth = FormMain.tudSelWidth
    End If
    If FormMain.tudSelHeight.IsValid(False) Then
        If FormMain.tudSelHeight + FormMain.tudSelTop > containingPDImage.Height Then
            'FormMain.txtSelHeight = pdImages(CurrentImage).Height - Val(FormMain.txtSelTop)
            selTop = containingPDImage.Height - FormMain.tudSelHeight
            FormMain.tudSelTop = selTop
        End If
        selHeight = FormMain.tudSelHeight
    End If
    
    'Perform one final sanity check on all selection values
    If selLeft < 0 Then
        selLeft = 0
        FormMain.tudSelLeft = selLeft
    End If
    If selTop < 0 Then
        selTop = 0
        FormMain.tudSelTop = selTop
    End If
    If selWidth > containingPDImage.Width Then
        selWidth = containingPDImage.Width
        FormMain.tudSelWidth = selWidth
    End If
    If selWidth = 0 Then
        selWidth = 1
        FormMain.tudSelWidth = selWidth
    End If
    If selHeight > containingPDImage.Height Then
        selHeight = containingPDImage.Height
        FormMain.tudSelHeight = selHeight
    End If
    If selHeight = 0 Then
        selHeight = 1
        FormMain.tudSelHeight = selHeight
    End If
    
    'Adjust the x1, y1, x2, y2 values to match these new ones
    x1 = selLeft
    y1 = selTop
    x2 = selLeft + selWidth
    y2 = selTop + selHeight
    
    rejectRefreshRequests = False
    
    'Redraw the selection to the active form
    RenderViewport FormMain.ActiveForm
    
End Sub

'"Lock-in" a selection. Typically this is prompted by a _MouseUp event
Public Sub lockIn(ByRef targetForm As Form)
    
    'Mark this selection as locked-in
    lockedIn = True
        
    'Update the internal coordinates one final time
    updateInternalCoords
    
    'The final thing we need to check for is the width and height, which may be still be zero at this point.  Do not allow selections of size 0.
    If selWidth < 1 Then selWidth = 1
    If selHeight < 1 Then selHeight = 1
    
    'Adjust the main form's selection text boxes to match the newly calculated values
    refreshTextBoxes
        
End Sub

'"Unlock" a selection
Public Sub lockRelease()

    lockedIn = False

End Sub

'Create a selection mask based on the current selection type.  A few items to note:
' 1) The selection mask is always the size of the full image.  This makes transforms (e.g. "grow selection") much easier to handle.
' 2) Black pixels in the mask represent unselected pixels in the image.  White represents selected.  Other shades can be used for aliasing
'     or partial selections.
' 3) The selection mask is stored as a pdLayer object, so any image filters can be applied to it.
' 4) For shape-based selections (rectangle, square, etc), the selection's dimensions need to be set BEFORE calling this function.  This function
'     relies on things like selLeft and selWidth to know where to render the mask.
Private Sub createSelectionMask()

    'Start by creating a blank mask (this will also erase any existing mask)
    selMask.createBlank containingPDImage.Width, containingPDImage.Height, 24, 0
    
    'The actual rendering of the selection will vary based on the current selection type (obviously).
    Select Case sType
    
        'Squares and rectangles are easy - use FillRect to draw a white rectangle at the mask's location
        Case sRectangle, sSquare
        
            Dim maskRect As RECT
            SetRect maskRect, selLeft, selTop, selLeft + selWidth, selTop + selHeight
            
            Dim hBrush As Long
            hBrush = CreateSolidBrush(RGB(255, 255, 255))
            FillRect selMask.getLayerDC, maskRect, hBrush
            DeleteObject hBrush
    
    End Select

End Sub

'Render a selection using one of several methods.
Public Sub renderCustom(ByRef dstLayer As pdLayer, ByRef dstForm As Form, ByVal imgLeft As Long, ByVal imgTop As Long, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal renderType As SelectionRender)

    'If an update has been requested, but we are already in the middle of one, deny subsequent requests
    If rejectRefreshRequests Then Exit Sub

    'Is the selection mask ready to be used?  If not, prepare it now.
    If Not isMaskReady Then createSelectionMask
    
    'Before drawing anything, we need to convert our "image" coordinates back to "viewport" coordinates.
    ' (e.g., the location of this selection is stored as image coordinates - but we need to have it in "viewport" coordinates so we can
    '        draw it to the screen, while accounting for things like zoom and scroll values)
    Dim ZoomVal As Double
    ZoomVal = g_Zoom.ZoomArray(pdImages(dstForm.Tag).CurrentZoomValue)
    
    Dim hScrollVal As Long, vScrollVal As Long
    hScrollVal = dstForm.HScroll.Value
    vScrollVal = dstForm.VScroll.Value
    
    Dim dstLeft As Double, dstRight As Double
    Dim dstTop As Double, dstBottom As Double
        
    dstLeft = ((selLeft - hScrollVal) * ZoomVal) + imgLeft
    dstRight = ((selLeft + selWidth - hScrollVal) * ZoomVal) + imgLeft + 0.5
    dstTop = ((selTop - vScrollVal) * ZoomVal) + imgTop
    dstBottom = ((selTop + selHeight - vScrollVal) * ZoomVal) + imgTop + 0.5
    
    'Several different selection rendering methods use the same variables.  I've declared the common ones here.
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    Dim x As Long, y As Long
    Dim QuickVal As Long
    
    Dim selMaskData() As Byte
    Dim selMaskSA As SAFEARRAY2D
    Dim selMaskQuickVal As Long
    
    Dim srcX As Long, srcY As Long
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = containingPDImage.Width - 1
    srcHeight = containingPDImage.Height - 1
    
    'Just to be safe, declare some bounding variables in advance as well
    Dim rightBound As Long, bottomBound As Long
    rightBound = imgLeft + imgWidth - 1
    bottomBound = imgTop + imgHeight - 1
            
    If rightBound > pdImages(dstForm.Tag).backBuffer.getLayerWidth - 1 Then rightBound = pdImages(dstForm.Tag).backBuffer.getLayerWidth - 1
    If bottomBound > pdImages(dstForm.Tag).backBuffer.getLayerHeight - 1 Then bottomBound = pdImages(dstForm.Tag).backBuffer.getLayerHeight - 1
        
    'Based on the selection type, draw a corresponding selection of some sort
    Select Case renderType
    
        'Lightbox style
        Case 0
        
            'Start by acquiring a pointer to the target layer's pixel bits
            prepLayerSafeArray dstLayer, tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
            
            'Next, acquire a pointer to the selection mask's pixel bits
            prepLayerSafeArray selMask, selMaskSA
            CopyMemory ByVal VarPtrArray(selMaskData()), VarPtr(selMaskSA), 4
            
            'Use a look-up table to dim the non-active parts of the image
            Dim hData(0 To 255) As Byte
            
            For x = 0 To 255
                hData(x) = x \ 3
            Next x
            
            dstRight = dstRight - 1
            dstBottom = dstBottom - 1
            
            'Dim the non-active parts of the image
            For x = imgLeft To rightBound
                QuickVal = x * 3
            For y = imgTop To bottomBound
            
                'Calculate source x and y positions, while taking into account zoom and scroll values
                srcX = hScrollVal + Int((x - imgLeft) / ZoomVal)
                srcY = vScrollVal + Int((y - imgTop) / ZoomVal)
                
                'Make sure srcX and srcY are not out of bounds
                If srcX < 0 Then srcX = 0
                If srcX > srcWidth Then srcX = srcWidth
                If srcY < 0 Then srcY = 0
                If srcY > srcHeight Then srcY = srcHeight
                
                selMaskQuickVal = srcX * 3
            
                'Remember that lightbox works a little weirdly - UNSELECTED pixels are marked, while selected ones are left untouched
                If selMaskData(selMaskQuickVal, srcY) = 0 Then
                    iData(QuickVal + 2, y) = hData(iData(QuickVal + 2, y))
                    iData(QuickVal + 1, y) = hData(iData(QuickVal + 1, y))
                    iData(QuickVal, y) = hData(iData(QuickVal, y))
                End If
                
            Next y
            Next x
            
            'With our work complete, point all arrays away from their respective DIBs, then deallocate them
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
            CopyMemory ByVal VarPtrArray(selMaskData), 0&, 4
            Erase iData
            Erase selMaskData
                    
            dstRight = dstRight + 1
            dstBottom = dstBottom + 1
        
        'Highlight style (Paint.NET)
        Case 1 To 2
        
            'Start by acquiring a pointer to the target layer's pixel bits
            prepLayerSafeArray dstLayer, tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
            
            'Next, acquire a pointer to the selection mask's pixel bits
            prepLayerSafeArray selMask, selMaskSA
            CopyMemory ByVal VarPtrArray(selMaskData()), VarPtr(selMaskSA), 4
            
            'Use look-up tables to highlight the active parts of the image
            Dim rData(0 To 255) As Byte, gData(0 To 255) As Byte, bData(0 To 255) As Byte
            
            Dim newR As Long, newG As Long, newB As Long
            
            If renderType = sHighlightBlue Then
                newR = 112
                newG = 183
                newB = 255
            Else
                newR = 255
                newG = 35
                newB = 49
            End If
            
            For x = 0 To 255
                rData(x) = (newR + x) \ 2
                gData(x) = (newG + x) \ 2
                bData(x) = (newB + x) \ 2
            Next x
            
            dstRight = dstRight - 1
            dstBottom = dstBottom - 1
            
            'Dim the non-active parts of the image
            For x = imgLeft To rightBound
                QuickVal = x * 3
            For y = imgTop To bottomBound
            
                'Calculate source x and y positions, while taking into account zoom and scroll values
                srcX = hScrollVal + Int((x - imgLeft) / ZoomVal)
                srcY = vScrollVal + Int((y - imgTop) / ZoomVal)
                
                'Make sure srcX and srcY are not out of bounds
                If srcX < 0 Then srcX = 0
                If srcX > srcWidth Then srcX = srcWidth
                If srcY < 0 Then srcY = 0
                If srcY > srcHeight Then srcY = srcHeight
                
                selMaskQuickVal = srcX * 3
            
                'If this pixels is over an active pixel in the selection mask, mark it with the selected color
                If selMaskData(selMaskQuickVal, srcY) > 0 Then
                    iData(QuickVal + 2, y) = rData(iData(QuickVal + 2, y))
                    iData(QuickVal + 1, y) = gData(iData(QuickVal + 1, y))
                    iData(QuickVal, y) = bData(iData(QuickVal, y))
                End If
                
            Next y
            Next x
            
            'With our work complete, point all arrays away from their respective DIBs, then deallocate them
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
            CopyMemory ByVal VarPtrArray(selMaskData), 0&, 4
            Erase iData
            Erase selMaskData
            
            dstRight = dstRight + 1
            dstBottom = dstBottom + 1
            
        'NOTE: focus rectangle is not used at present, so this technique may or may not work.  I simply keep it around for reference's sake.
        'Focus rectangle ("Simple" style)
        Case Else
        
            'Store the destination DC to a local variable
            Dim DstDC As Long
            DstDC = dstLayer.getLayerDC
            
            'Since we'll be using the API to draw our selection area, we need to initialize several brushes
            Dim hPen As Long, hOldPen As Long
            
            hPen = CreatePen(PS_DOT, 0, RGB(0, 0, 0))
            hOldPen = SelectObject(DstDC, hPen)
            
            'Get a transparent brush
            Dim hBrush As Long, hOldBrush As Long
            hBrush = GetStockObject(NULL_BRUSH)
            hOldBrush = SelectObject(DstDC, hBrush)
            
            'Change the rasterOp to XOR (this will invert the line)
            SetROP2 DstDC, vbSrcInvert
            
            'If dstLeft > imgLeft Then dstLeft = dstLeft - 1
            'If dstTop > imgTop Then dstTop = dstTop - 1
            'dstRight = dstRight + 1
            'dstBottom = dstBottom + 1
            
            'Draw the rectangle
            Rectangle DstDC, dstLeft, dstTop, dstRight, dstBottom
            
            'Restore the normal COPY rOp
            SetROP2 DstDC, vbSrcCopy
            
            'Remove the brush from the DC
            SelectObject DstDC, hOldBrush
            DeleteObject hBrush
            
            'Remove the pen from the DC
            SelectObject DstDC, hOldPen
            DeleteObject hPen
                    
    End Select
        
End Sub

Private Sub Class_Initialize()
    
    lockedIn = False
    
    'Initialize the selection mask for this object and mark it as "not ready" (because no mask has been drawn yet)
    Set selMask = New pdLayer
    isMaskReady = False
        
End Sub

'Sometimes a selection needs to directly access the bits of a target layer. This routine sets up the corresponding SafeArray.
Private Sub prepLayerSafeArray(ByRef srcLayer As pdLayer, ByRef tmpSA As SAFEARRAY2D)
    With tmpSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcLayer.getLayerHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcLayer.getLayerArrayWidth
        .pvData = srcLayer.getLayerDIBits
    End With
End Sub

Private Sub Class_Terminate()

    'Erase the selection mask
    If Not (selMask Is Nothing) Then
        selMask.eraseLayer
        Set selMask = Nothing
    End If

End Sub
