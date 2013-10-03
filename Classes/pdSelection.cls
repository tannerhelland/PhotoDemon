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
'Copyright ©2012-2013 by Tanner Helland
'Created: 25/September/12
'Last updated: 16/September/13
'Last update: fix bug where feathering radius was still used to calculate selection boundaries, if feathering had been
'              used previously but since disabled.
'
'This class handles all selections in PhotoDemon.  Rectangular and elliptical selections are now implemented, with more
' to follow in the coming months.
'
'All selections are treated as a subset of pixels within a rectangular area. The x1, y1, x2, y2 coordinates define opposite
' corners of the selection area. There is no guarantee that these coordinates will appear in a given corner - so they must
' be checked manually. When a selection is set, those values are translated into Left, Top, Width and Height values, which
' are in turn used for everything from cropping to applying filters.
'
'Note also that this class is treated as a subset of pdLayer(). Right now each layer object only contains one selection,
' but there's no reason it couldn't store multiples in the future. (Some software allows for selection blending modes, e.g.
' you can draw one selection, then draw another and the two will be merged into one - multiple selections are required
' for that.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Used when writing/reading selection data to/from a file
Private Const SELECTION_IDENTIFIER As String * 4 = "PDsd"
Private Const SELECTION_FILE_VERSION_2013 As Long = &H1002

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
'Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
'Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

'Region drawing
Private Declare Function CreateEllipticRgnIndirect Lib "gdi32" (ByRef lpRect As RECT) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

'Brush creation
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'Stock object retrieval
Private Const NULL_BRUSH = 5
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

'RasterOp selection
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long

'What shape does this selection have?
Public Enum SelectionShape
    sRectangle = 0
    sCircle = 1
    sLine = 2
    sPolygon = 3
End Enum

#If False Then
    Const sRectangle = 0
    Const sCircle = 1
    Const sLine = 2
    Const sPolygon = 3
#End If

Private sShape As SelectionShape

'What type of selection is this?
Public Enum SelectionType
    sInterior = 0
    sExterior = 1
    sBorder = 2
End Enum

#If False Then
    Const sInterior = 0
    Const sExterior = 1
    Const sBorder = 2
#End If

Private sType As SelectionType

'What kind of smoothing (feathering) will be applied to this selection?
Public Enum SelectionSmoothing
    sNone = 0
    sAntialiased = 1
    sFullyFeathered = 2
End Enum

#If False Then
    Const sNone = 0
    Const sAntialiased = 1
    Const sFullyFeathered = 2
#End If

Private sSmooth As SelectionSmoothing

'The coordinates of the current selection (if rectangular)
Public x1 As Double, x2 As Double
Public y1 As Double, y2 As Double

'The actual coordinates passed from the user's mouse movements to this selection object.  These coordinates are not modified in any way.
' (e.g. these coordinates could lie outside the image, viewport, or screen, because no bounds checking is applied to them)
Private actualX1 As Double, actualX2 As Double
Private actualY1 As Double, actualY2 As Double

'When a selection is "locked in", the x and y values of corner points are converted to these values
Public selLeft As Long, selTop As Long
Public selWidth As Long, selHeight As Long

'And finally, for nonstandard or inverted selections, an additional set of coordinates is required to track the bounding area of the
' selection as a whole.  External functions need these values to determine how to handle the selection's data.
Public boundLeft As Long, boundTop As Long
Public boundWidth As Long, boundHeight As Long

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

'Until a mask has been created, this value will remain "false".  It is used to optimize operations on blank selections.
Public maskHasBeenCreated As Boolean

'After the mask has been generated, its status will be set as "ready".  Any function that changes the selected area will reset this to "false".
' The selection rendering code can check this value to see if the mask needs to be recreated before rendering it to screen.
Private isMaskReady As Boolean

'Some types of transformations (basic shapes, like rectangles, etc), are transformable, meaning that after they are created, the user can
' click on them again to resize or move them.  Complex transformations (magic wand, lasso, etc), may not be transformable.  This boolean
' is read by the mouse tracker in the image form, and it uses it to determine if the user is allowed to transform the current selection.
Public isTransformable As Boolean

'If the user holds "Shift" while moving the selection, it will be forced to a 1:1 aspect ratio.
Private isSquare As Boolean

'Rectangular selections can have rounded corners.  If this value is > 0, round the corners by that many pixels.
Private roundedCorners As Long

'Selections can be border-only.  This value specifies the border radius, which cannot be larger than (minimum dimension \ 2)
Private borderSize As Long

'Selections can have any amount of feathering.  Feathering is implemented identically to GIMP (http://docs.gimp.org/en/gimp-painting.html#gimp-concepts-selection),
' specifically, feathering is implemented by applying a Gaussian blur to the selection mask with radius equal to the feathering size.
Private featheringRadius As Long

'Line selections have variable width.
Private selectionLineWidth As Long

'For Undo/Redo purposes, it is sometimes necessary to store a separate set of selection parameters and data.  Such parameters
' are stored in this string
Private backupParamString As String, backupSelectionState As Boolean, backupTransformable As Boolean
Private backupMask As pdLayer

'Undo functions involve writing past selection data to file.  That past data can be set by calling this function.
Public Sub setBackupParamString()
    backupSelectionState = lockedIn
    backupTransformable = isTransformable
    backupParamString = buildParams(sShape, sType, sSmooth, featheringRadius, borderSize, roundedCorners, selectionLineWidth, selLeft, selTop, selWidth, selHeight, x1, y1, x2, y2)
    backupParamString = translateParamStringToRelativeFormat(backupParamString, True)
    Set backupMask = New pdLayer
    backupMask.createFromExistingLayer selMask
End Sub

'To save code elsewhere, a selection can be initialized by using a param string generated by a pdParamString object.  This condenses a huge
' list of variables into a single string.  This function is especially helpful when writing full selection data to/from file.
Public Sub initFromParamString(ByVal paramString As String)

    'Before doing anything else, convert the paramString from its current "relative" coordinates to "absolute" coordinates based on
    ' the size of the parent image.
    paramString = translateParamStringToRelativeFormat(paramString, False)

    'First, mark the mask as "not ready"
    isMaskReady = False

    'Start by creating a parameter parser to handle the parameter string.  This class will parse out individual parameters
    ' as specific data types upon request.
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(paramString) > 0 Then cParams.setParamString paramString
    
    'The order of param string values is HARD-CODED per the following table:
    
    '01 - selection shape
    setSelectionShape cParams.GetLong(1)
    
    '02 - selection type
    setSelectionType cParams.GetLong(2)
    
    '03 - smoothing type
    setSmoothingType cParams.GetLong(3)
    
    '04 - feathering radius
    setFeatheringRadius cParams.GetLong(4)
    
    '05 - border size
    setBorderSize cParams.GetLong(5)
    
    '06 - rounded corner amount
    setRoundedCornerAmount cParams.GetLong(6)
    
    '07 - selection line width
    setSelectionLineWidth cParams.GetLong(7)
    
    '08 - left
    selLeft = cParams.GetLong(8)
    
    '09 - top
    selTop = cParams.GetLong(9)
    
    '10 - width
    selWidth = cParams.GetLong(10)
    
    '11 - height
    selHeight = cParams.GetLong(11)
    
    '12 - x1
    x1 = cParams.GetDouble(12)
    
    '13 - y1
    y1 = cParams.GetDouble(13)
    
    '14 - x2
    x2 = cParams.GetDouble(14)
    
    '15 - y2
    y2 = cParams.GetDouble(15)
    
End Sub

'Return all of this selection's important settings as a specifically formatted parameter string.  This string can be passed to a selection
' object at any future point to recreate this selection exactly.
' (See also: "initFromParamString" above)
Public Function getSelectionParamString() As String
    getSelectionParamString = buildParams(sShape, sType, sSmooth, featheringRadius, borderSize, roundedCorners, selectionLineWidth, selLeft, selTop, selWidth, selHeight, x1, y1, x2, y2)
    getSelectionParamString = translateParamStringToRelativeFormat(getSelectionParamString, True)
End Function

'Simple - use this to select the entire image attached to this selection object.  Note that this DOES NOT render the new selection on-screen,
' and it doesn't render a matching selection mask.  The calling function is responsible for that.
Public Sub selectAll()

    'Set basic information about this selection
    setSelectionType sInterior
    setSelectionShape sRectangle
    setFeatheringRadius 0
    setSmoothingType sAntialiased
    setRoundedCornerAmount 0
    x1 = 0
    y1 = 0
    x2 = containingPDImage.Width
    y2 = containingPDImage.Height
    selLeft = 0
    selTop = 0
    selWidth = containingPDImage.Width
    selHeight = containingPDImage.Height
    isTransformable = True
    isMaskReady = False
    
End Sub

'Request a 1:1 aspect ratio selection (squares, circles)
Public Sub requestSquare(ByVal requestChoice As Boolean, Optional ByVal forceSelectionRefresh As Boolean = False)
    
    isSquare = requestChoice
    If isTransformable Then isMaskReady = False
    
    'Redraw the selection if necessary
    If forceSelectionRefresh And (Not containingPDImage Is Nothing) Then
        updateInternalCoords
        createSelectionMask
    End If
    
End Sub

'Request a redraw of the selection mask.  We must do this when loading an Undo or Redo request after the image size has
' been changed; otherwise, OOB errors can occur (because the selection mask will be a different size than the image).
Public Sub requestNewMask()

    If (Not containingPDImage Is Nothing) Then
        updateInternalCoords
        If isTransformable Then
            createSelectionMask
        Else
        
        End If
    End If

End Sub

'Get/set feathering radius
Public Function getFeatheringRadius() As Long
    getFeatheringRadius = featheringRadius
End Function

Public Sub setFeatheringRadius(ByVal newRadius As Long)
    If isTransformable Then isMaskReady = False
    featheringRadius = newRadius
End Sub

'Get/set the border size
Public Function getBorderSize() As Long
    getBorderSize = borderSize
End Function

Public Sub setBorderSize(ByVal newBorderSize As Long)
    If isTransformable Then isMaskReady = False
    borderSize = newBorderSize
End Sub

'Get/set a transformation type
Public Function getTransformationType()
    If transformModeActive Then
        getTransformationType = selTransformationType
    Else
        getTransformationType = 0
    End If
End Function

Public Sub setTransformationType(ByVal transShape As Long)
    selTransformationType = transShape
End Sub

'Get/set the selection's current coordinates
Public Sub getSelectionCoordinates(ByVal numPoint As Long, ByRef xDst As Double, ByRef yDst As Double)
    If numPoint = 1 Then
        xDst = x1
        yDst = y1
    Else
        xDst = x2
        yDst = y2
    End If
End Sub

'Get/set a selection shape
Public Function getSelectionShape() As SelectionShape
    getSelectionShape = sShape
End Function

Public Sub setSelectionShape(ByVal selShape As SelectionShape)
    
    'Certain types of shapes are transformable.  Mark those now.
    Select Case selShape
        
        Case sRectangle, sCircle, sLine
            isTransformable = True
        
    End Select
    
    If isTransformable Then isMaskReady = False
    sShape = selShape
End Sub

'Get/set a selection type
Public Function getSelectionType() As SelectionType
    getSelectionType = sType
End Function

Public Sub setSelectionType(ByVal selType As SelectionType)
    sType = selType
End Sub

'Get/set a smoothing type
Public Function getSmoothingType() As SelectionSmoothing
    getSmoothingType = sSmooth
End Function

Public Sub setSmoothingType(ByVal selSmooth As SelectionSmoothing)
    If isTransformable Then isMaskReady = False
    sSmooth = selSmooth
End Sub

'Change the amount of corner rounding for rectangular selections
Public Function getRoundedCornerAmount() As Long
    getRoundedCornerAmount = roundedCorners
End Function

Public Sub setRoundedCornerAmount(ByVal cornerRounding As Long)
    If isTransformable Then isMaskReady = False
    roundedCorners = cornerRounding
End Sub

'Get/set the line width of line selections
Public Function getSelectionLineWidth() As Long
    getSelectionLineWidth = selectionLineWidth
End Function

Public Sub setSelectionLineWidth(ByVal newLineWidth As Long)
    If isTransformable Then isMaskReady = False
    selectionLineWidth = newLineWidth
End Sub

'Takes x and y coordinates (from a _MouseDown event, typically) and uses them in a manner specified by the current transform operation.
' Note that this this should only be called after a transformation type has been set.
Public Sub setInitialTransformCoordinates(ByVal X As Double, Y As Double)

    'If new transform coordinates are being set, this selection must be "unlocked"
    lockedIn = False
    
    'The use of setInitialCoordinates means this IS a transformation
    transformModeActive = True
    
    'Different selection types handle transformation differently.  For example, rectangular selections can be resized in multiple directions,
    ' but a line selection will only move its endpoints.  So we must sort input twice - first, sort by selection type, and second, by
    ' transformation type.
    
    Select Case sShape
    
        'Rectangular and elliptical selections are handled identically
        Case sRectangle, sCircle
        
            'Based on the transform mode, set the initial points accordingly
            Select Case selTransformationType
            
                'Case 0 should never occur, but if it does - treat this like a normal initial coordinate call
                Case 0
                    x1 = X
                    y1 = Y
                    x2 = X
                    y2 = Y
                
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
                    x2 = X
                    y2 = Y
                
                '3 - SE corner
                Case 3
                    x1 = selLeft
                    y1 = selTop
                    x2 = X
                    y2 = Y
                
                '4 - SW corner
                Case 4
                    x1 = selLeft + selWidth
                    y1 = selTop
                    x2 = X
                    y2 = Y
                
                '5 - N edge
                Case 5
                    x1 = selLeft
                    x2 = selLeft + selWidth
                    y1 = selTop + selHeight
                    y2 = Y
                
                '6 - E edge
                Case 6
                    x1 = selLeft
                    x2 = X
                    y1 = selTop
                    y2 = selTop + selHeight
                
                '7 - S edge
                Case 7
                    x1 = selLeft
                    x2 = selLeft + selWidth
                    y1 = selTop
                    y2 = Y
                
                '8 - W edge
                Case 8
                    x1 = selLeft + selWidth
                    x2 = X
                    y1 = selTop
                    y2 = selTop + selHeight
                
                '9 - interior of selection, not near a corner or edge
                Case 9
                    moveXDist = X - selLeft
                    moveYDist = Y - selTop
            
            End Select
        
        
        'Line selections
        Case sLine
        
            'Case 0 should never occur, but if it does - treat this like a normal initial coordinate call
            Case 0
                x1 = X
                y1 = Y
                x2 = X
                y2 = Y
            
            'First point is being moved
            Case 1
                x1 = X
                y1 = Y
            
            'Second point is being moved
            Case 2
                x2 = X
                y2 = Y
        
    End Select
    
    updateInternalCoords
    
End Sub

'Takes x and y coordinates (from a _MouseDown event, typically) and stores them internally
Public Sub setInitialCoordinates(ByVal X As Double, Y As Double)
    
    'If new initial coordinates are being set, this selection must be "unlocked"
    lockedIn = False
    
    'The use of setInitialCoordinates means this is not a transformation
    transformModeActive = False
    
    x1 = X
    y1 = Y
    
    'Make a back-up copy of the coordinates.  These values will never be changed except by the user moving the mouse.
    actualX1 = X
    actualY1 = Y
    
    'Set the second set of point to match the first set
    Select Case sShape
        
        Case sRectangle, sCircle, sLine
            x2 = X
            y2 = Y
            
            'Also, mark selections of any of these shapes (rectangle, circle, line) as transformable
            isTransformable = True
        
        Case Else
        
    End Select
    
    updateInternalCoords
    
End Sub

'Takes x and y coordinates (from a _MouseMove event, typically) and stores them internally
Public Sub setAdditionalCoordinates(ByVal X As Double, Y As Double)
    
    'Check for an active transformation mode.  (A transformation is something like resizing or moving an existing selection,
    ' versus drawing a new one from scratch.)
    If transformModeActive Then
        
        Select Case sShape
    
            Case sRectangle, sCircle
            
                'Based on the transform mode, set the initial points accordingly
                Select Case selTransformationType
                
                    'Cases 0 should never occur, but if it does - treat this like a normal subsequent coordinate call
                    'Cases 1-4 can similarly be treated like a normal subsequent coordinate call
                    ' (See setInitialTransformCoordinates for details on what the different cases mean.  And yes, I should convert
                    '  these to constants instead of numbers... :P)
                    Case 0 To 4
                        x2 = X
                        y2 = Y
                    
                    '5 - N edge
                    Case 5
                        y2 = Y
                        
                    '6 - E edge
                    Case 6
                        x2 = X
                        
                    '7 - S edge
                    Case 7
                        y2 = Y
                        
                    '8 - W edge
                    Case 8
                        x2 = X
                        
                    '9 - interior of selection, not near a corner or edge (e.g. move the selection, but don't resize it)
                    Case 9
                    
                        x1 = X - moveXDist
                        y1 = Y - moveYDist
                        x2 = x1 + selWidth
                        y2 = y1 + selHeight
                        'Message x1 & "," & y1 & " <--> " & x2 & "," & y2 & " | " & selLeft & "," & selTop & " <--> " & selWidth & "," & selHeight
                
                End Select
        
            Case sLine
            
                'Based on the transform mode, set the initial points accordingly
                Select Case selTransformationType
            
                    'Case 0 should never occur
                    Case 0
                    
                    'First point is being moved
                    Case 1
                        x1 = X
                        y1 = Y
                        
                    'Second point is being moved
                    Case 2
                        x2 = X
                        y2 = Y
                        
                End Select
                        
        End Select
        
        'If a transform mode is active, re-mark the selection as being transformable
        isTransformable = True
        
        updateInternalCoords
    
    'No transform mode makes this easy; just update the second set of points and redraw the selection
    Else
    
        x2 = X
        y2 = Y
        
        'Make a back-up copy of the coordinates.  These values will never be changed except by the user moving the mouse.
        actualX2 = X
        actualY2 = Y
    
        updateInternalCoords
        
    End If
    
End Sub

'Has this selection been locked in?
Public Function isLockedIn() As Boolean
    isLockedIn = lockedIn
End Function

'Whenever the x1, y1, x2, y2 values are changed, this sub needs to be called to update the left/right/width/height values accordingly
Private Sub updateInternalCoords()

    'This function only needs to be run if the selection is transformable.  If it is not, the internal coordinates will already be correctly set.
    If isTransformable Then
    
        'Mark the selection mask as "not ready", as it will need to be redrawn after new coordinates are set
        isMaskReady = False

        'If a square (1:1 aspect ratio) selection has been requested, calculate new coordinates now.
        
        ' (This set of if/then blocks looks complicated, but it's actually very simple - we simply have to account for every variation
        '  of quadrants, because the selection can be drawn up or down in both directions, giving eight possible variants of x1 </> x2
        '  and y1 </> y2.  By covering all those cases, square selections can be drawn in any direction.)
        If isSquare Then
            If x1 < x2 Then
                If y1 < y2 Then
                    If Abs(x1 - x2) > Abs(y1 - y2) Then
                        y2 = y1 + Abs(x1 - x2)
                    Else
                        x2 = x1 + Abs(y1 - y2)
                    End If
                Else
                    If Abs(x1 - x2) > Abs(y1 - y2) Then
                        y2 = y1 - Abs(x1 - x2)
                    Else
                        x2 = x1 + Abs(y1 - y2)
                    End If
                End If
            Else
                If y1 < y2 Then
                    If Abs(x1 - x2) > Abs(y1 - y2) Then
                        y2 = y1 + Abs(x1 - x2)
                    Else
                        x2 = x1 - Abs(y1 - y2)
                    End If
                Else
                    If Abs(x1 - x2) > Abs(y1 - y2) Then
                        y2 = y1 - Abs(x1 - x2)
                    Else
                        x2 = x1 - Abs(y1 - y2)
                    End If
                End If
            End If
        End If
        
        'Finally, calculate a left, top, width and height for this selection based off the current x1/x2/y1/y2 values
        If x1 < x2 Then
            selLeft = x1
            selWidth = x2 - x1
        Else
            selLeft = x2
            selWidth = x1 - x2
        End If
        
        If y1 < y2 Then
            selTop = y1
            selHeight = y2 - y1
        Else
            selTop = y2
            selHeight = y1 - y2
        End If
        
    End If

End Sub

'Because selections can be created beyond the parent image's borders, it is necessary to check if ALL selection coordinates lie off the
' image.  If this is the case, we don't want to finalize the current selection - we want to forget it.
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

'Update this selection using the values in the main form's selection text boxes
Public Sub updateViaTextBox()

    'Ignore text box update requests until the selection is locked in
    If Not isLockedIn Then Exit Sub

    rejectRefreshRequests = True

    'Mark the selection mask as "not ready", as it will need to be redrawn after new coordinates are set
    isMaskReady = False

    'Check all text box entries for validity, then update the corresponding selection values
    If toolbar_Selections.tudSel(0).IsValid(False) Then selLeft = toolbar_Selections.tudSel(0)
    If toolbar_Selections.tudSel(1).IsValid(False) Then selTop = toolbar_Selections.tudSel(1)
    If toolbar_Selections.tudSel(2).IsValid(False) Then selWidth = toolbar_Selections.tudSel(2)
    If toolbar_Selections.tudSel(3).IsValid(False) Then selHeight = toolbar_Selections.tudSel(3)
    
    'Adjust the x1, y1, x2, y2 values to match these new ones
    Select Case sShape
    
        Case sRectangle, sCircle
            x1 = selLeft
            y1 = selTop
            x2 = selLeft + selWidth
            y2 = selTop + selHeight
            
        Case sLine
            x1 = selLeft
            y1 = selTop
            x2 = selWidth
            y2 = selHeight
        
    End Select
    
    rejectRefreshRequests = False
    
End Sub

'"Lock-in" a selection. Typically this is prompted by a _MouseUp event
Public Sub lockIn()
    
    'Mark this selection as locked-in
    lockedIn = True
        
    'For transformable selections, update the internal coordinates one final time
    If isTransformable Then
        
        updateInternalCoords
    
        'The final thing we need to check for is the width and height, which may be still be zero at this point.  Do not allow selections of size 0.
        If selWidth < 1 Then selWidth = 1
        If selHeight < 1 Then selHeight = 1
        
    End If
        
End Sub

'"Unlock" a selection
Public Sub lockRelease()
    lockedIn = False
End Sub

'Create a selection mask based on the current selection type.  A few items to note:
' 1) The selection mask is always the size of the full image.  This makes transforms (e.g. "grow selection") much easier to handle.
' 2) Black pixels (0) in the mask represent unselected pixels in the image.  White (255) represents selected.  Other values can be
'     used to specify aliasing or partial selections.
' 3) The selection mask is stored as a pdLayer object, so any image filters can be applied to it.
' 4) For shape-based selections (rectangle, square, etc), the selection's dimensions need to be set BEFORE calling this function.  This function
'     relies on things like selLeft and selWidth to know where to render the mask.  For non-shape-based selections, this function will call a
'     separate function to find the bounding rect.
Private Sub createSelectionMask()

    maskHasBeenCreated = True

    Dim maskBackColor As Long, maskForeColor As Long
    
    'Interior and exterior selections are rendered using identical code; the only difference is the colors used
    If sType = sExterior Then
        maskBackColor = RGB(255, 255, 255)
        maskForeColor = RGB(0, 0, 0)
    Else
        maskBackColor = RGB(0, 0, 0)
        maskForeColor = RGB(255, 255, 255)
    End If

    'Some selection types (line selections) need to know max/min values, which are separate from left/top/width/height
    Dim minX As Long, maxX As Long, MinY As Long, MaxY As Long
    If x1 < x2 Then
        minX = x1
        maxX = x2
    Else
        minX = x2
        maxX = x1
    End If
    If y1 < y2 Then
        MinY = y1
        MaxY = y2
    Else
        MinY = y2
        MaxY = y1
    End If

    'At present, mask creation is only applicable for certain transformable shapes (rectangles, ellipses, lines).  Other functions,
    ' like "Invert selection", rely on an already-created mask - so attempting to create a mask again will have undesirable behavior.
    Debug.Print "Drawing new selection mask now"

    'Start by creating a blank mask (this will also erase any existing mask)
    selMask.createBlank containingPDImage.Width, containingPDImage.Height, 24, maskBackColor
    
    Dim hBrush As Long
    Dim maskRect As RECT
    Dim maskRegion As Long
    
    'Border-style selections require some additional checking; in particular, if the border value is too large, just ignore it
    ' and draw a regular shape instead.
    Dim rectModifier As Long
    rectModifier = borderSize
    
    Dim modifiedCornerVal As Long
    
    Dim minDimension As Long
    If selWidth < selHeight Then minDimension = selWidth Else minDimension = selHeight
    
    Dim ignoreBorderValue As Boolean
    ignoreBorderValue = False
    
    'Border selections require a bit of extra calculation, depending on the shape - basically, all bordered selections are created by
    ' rendering a smaller version of the selection within itself.  In the case of rectangles and ellipses, we don't want the smaller
    ' version to have negative values, so we check it in advance to make sure the border value is acceptable.
    If sType = sBorder Then
        
        Select Case sShape
            
            Case sRectangle, sCircle
                If rectModifier < 1 Then rectModifier = 1
                If rectModifier > (minDimension \ 2) Then ignoreBorderValue = True
                
            Case Else
            
        End Select
            
    End If
                
    'The actual rendering of the selection will vary based on the current selection type (obviously).
    Select Case sShape
    
        'Squares and rectangles are easy - use FillRect to draw a white rectangle at the mask's location.  If corner rounding
        ' has been specified, apply it using CreateRoundRectRgn.
        Case sRectangle
        
            'RECTANGLE SELECTION, NO ROUNDED CORNERS
            ' (Note: rectangular selections ignore antialiasing, as there's no point)
            If roundedCorners = 0 Then
            
                'Interior/exterior selections use identical rendering code.
                If (sType = sInterior) Or (sType = sExterior) Or ignoreBorderValue Then
                
                    SetRect maskRect, selLeft, selTop, selLeft + selWidth, selTop + selHeight
                    hBrush = CreateSolidBrush(maskForeColor)
                    FillRect selMask.getLayerDC, maskRect, hBrush
                    DeleteObject hBrush
                    
                'Border selections are a bit different; we must draw two rectangles - an exterior and an interior one.
                Else
                
                    'Draw the exterior rectangle
                    hBrush = CreateSolidBrush(maskForeColor)
                    SetRect maskRect, selLeft, selTop, selLeft + selWidth, selTop + selHeight
                    FillRect selMask.getLayerDC, maskRect, hBrush
                    DeleteObject hBrush
                    
                    'Next, draw the interior one
                    hBrush = CreateSolidBrush(maskBackColor)
                    SetRect maskRect, selLeft + rectModifier, selTop + rectModifier, selLeft + selWidth - rectModifier, selTop + selHeight - rectModifier
                    FillRect selMask.getLayerDC, maskRect, hBrush
                    DeleteObject hBrush
                    
                End If
            
            'RECTANGLE SELECTION *WITH* ROUNDED CORNERS
            Else
                
                'Interior/exterior selections use identical rendering code.
                If (sType = sInterior) Or (sType = sExterior) Or ignoreBorderValue Then
                    
                    'Use GDI+ for antialiasing, regular GDI for non-antialiased
                    If sSmooth > sNone Then
                        If roundedCorners > minDimension Then modifiedCornerVal = minDimension Else modifiedCornerVal = roundedCorners
                        GDIPlusDrawRoundRect selMask, selLeft, selTop, selWidth, selHeight, modifiedCornerVal, maskForeColor, True
                    Else
                        maskRegion = CreateRoundRectRgn(selLeft, selTop, selLeft + selWidth + 1, selTop + selHeight + 1, roundedCorners, roundedCorners)
                        hBrush = CreateSolidBrush(maskForeColor)
                        FillRgn selMask.getLayerDC, maskRegion, hBrush
                        DeleteObject hBrush
                        DeleteObject maskRegion
                    End If
                    
                'Border selections are a bit different; we must draw two rounded rectangles - an exterior and an interior one.
                Else
                    
                    If sSmooth > sNone Then
                    
                        'Draw the exterior rectangle
                        If roundedCorners > minDimension Then modifiedCornerVal = minDimension Else modifiedCornerVal = roundedCorners
                        GDIPlusDrawRoundRect selMask, selLeft, selTop, selWidth, selHeight, modifiedCornerVal, maskForeColor, True
                        
                        'Next, draw the interior one.  Note that specialized handling is required to prevent the corner curvature from rendering incorrectly.
                        If (selWidth - (rectModifier * 2) - 1) < (selHeight - (rectModifier * 2) - 1) Then minDimension = (selWidth - (rectModifier * 2) - 1) Else minDimension = (selHeight - (rectModifier * 2) - 1)
                        If roundedCorners > minDimension Then modifiedCornerVal = minDimension Else modifiedCornerVal = roundedCorners
                        GDIPlusDrawRoundRect selMask, selLeft + rectModifier, selTop + rectModifier, selWidth - rectModifier * 2, selHeight - rectModifier * 2, modifiedCornerVal, maskBackColor, True
                    
                    Else
                        
                        maskRegion = CreateRoundRectRgn(selLeft, selTop, selLeft + selWidth + 1, selTop + selHeight + 1, roundedCorners, roundedCorners)
                        hBrush = CreateSolidBrush(maskForeColor)
                        FillRgn selMask.getLayerDC, maskRegion, hBrush
                        DeleteObject hBrush
                        DeleteObject maskRegion
                        
                        maskRegion = CreateRoundRectRgn(selLeft + rectModifier, selTop + rectModifier, selLeft + selWidth + 1 - rectModifier, selTop + selHeight + 1 - rectModifier, roundedCorners, roundedCorners)
                        hBrush = CreateSolidBrush(maskBackColor)
                        FillRgn selMask.getLayerDC, maskRegion, hBrush
                        DeleteObject hBrush
                        DeleteObject maskRegion
                        
                    End If
                    
                End If
                
            End If
                    
        'CIRCLES / ELLIPSES
        Case sCircle
        
            'Interior/exterior selections use identical rendering code.
            If (sType = sInterior) Or (sType = sExterior) Or ignoreBorderValue Then
            
                If sSmooth > sNone Then
                    GDIPlusDrawEllipseToDC selMask.getLayerDC, selLeft, selTop, selWidth - 1, selHeight - 1, maskForeColor, True
                Else
                    SetRect maskRect, selLeft, selTop, selLeft + selWidth + 1, selTop + selHeight + 1
                    maskRegion = CreateEllipticRgnIndirect(maskRect)
                    hBrush = CreateSolidBrush(maskForeColor)
                    FillRgn selMask.getLayerDC, maskRegion, hBrush
                    DeleteObject hBrush
                    DeleteObject maskRegion
                End If
                
            'Border selections are a bit different; we must draw two ellipses - an exterior and an interior one.
            Else
            
                If sSmooth > sNone Then
                
                    'Draw the exterior oval
                    GDIPlusDrawEllipseToDC selMask.getLayerDC, selLeft, selTop, selWidth, selHeight, maskForeColor, True
                    
                    'Next, draw the interior one
                    GDIPlusDrawEllipseToDC selMask.getLayerDC, selLeft + rectModifier, selTop + rectModifier, selWidth - rectModifier * 2, selHeight - rectModifier * 2, maskBackColor, True
                        
                Else
                    
                    SetRect maskRect, selLeft, selTop, selLeft + selWidth + 1, selTop + selHeight + 1
                    maskRegion = CreateEllipticRgnIndirect(maskRect)
                    hBrush = CreateSolidBrush(maskForeColor)
                    FillRgn selMask.getLayerDC, maskRegion, hBrush
                    DeleteObject hBrush
                    DeleteObject maskRegion
                    
                    SetRect maskRect, selLeft + rectModifier, selTop + rectModifier, selLeft + selWidth + 1 - rectModifier, selTop + selHeight + 1 - rectModifier
                    maskRegion = CreateEllipticRgnIndirect(maskRect)
                    hBrush = CreateSolidBrush(maskBackColor)
                    FillRgn selMask.getLayerDC, maskRegion, hBrush
                    DeleteObject hBrush
                    DeleteObject maskRegion
                    
                End If
                        
            End If
            
        
        'Line selections
        Case sLine
        
            'Interior/exterior line selections use identical rendering code.
            If (sType = sInterior) Or (sType = sExterior) Or ignoreBorderValue Then
            
                If sSmooth > sNone Then
                    GDIPlusDrawLineToDC selMask.getLayerDC, x1, y1, x2, y2, maskForeColor, 255, selectionLineWidth, True, LineCapRound
                Else
                    GDIPlusDrawLineToDC selMask.getLayerDC, x1, y1, x2, y2, maskForeColor, 255, selectionLineWidth, False, LineCapRound
                End If
                
            'Border selections are a bit different; we must draw two lines - an interior and an exterior one
            Else
            
                Dim useLineAA As Boolean
                If sSmooth > sNone Then useLineAA = True Else useLineAA = False
                
                'Draw the exterior line
                GDIPlusDrawLineToDC selMask.getLayerDC, x1, y1, x2, y2, maskForeColor, 255, selectionLineWidth, useLineAA, LineCapRound
                
                'Next, draw the interior one
                Dim xDiff As Double, yDiff As Double
                xDiff = Abs(x1 - x2)
                yDiff = Abs(y1 - y2)
                Dim xModifier As Long, yModifier As Long
                If xDiff > yDiff Then
                    If xDiff = 0 Then xDiff = 0.00001
                    xModifier = borderSize
                    yModifier = borderSize * (yDiff / xDiff)
                Else
                    If yDiff = 0 Then yDiff = 0.00001
                    xModifier = borderSize * (xDiff / yDiff)
                    yModifier = borderSize
                End If
                
                Dim newLineWidth As Double
                newLineWidth = selectionLineWidth - borderSize
                
                If newLineWidth > 0 Then
                    If x1 = minX Then
                        If y1 = MinY Then
                            GDIPlusDrawLineToDC selMask.getLayerDC, x1 + xModifier, y1 + yModifier, x2 - xModifier, y2 - yModifier, maskBackColor, 255, newLineWidth, useLineAA, LineCapRound
                        Else
                            GDIPlusDrawLineToDC selMask.getLayerDC, x1 + xModifier, y1 - yModifier, x2 - xModifier, y2 + yModifier, maskBackColor, 255, newLineWidth, useLineAA, LineCapRound
                        End If
                    Else
                        If y1 = MinY Then
                            GDIPlusDrawLineToDC selMask.getLayerDC, x1 - xModifier, y1 + yModifier, x2 + xModifier, y2 - yModifier, maskBackColor, 255, newLineWidth, useLineAA, LineCapRound
                        Else
                            GDIPlusDrawLineToDC selMask.getLayerDC, x1 - xModifier, y1 - yModifier, x2 + xModifier, y2 + yModifier, maskBackColor, 255, newLineWidth, useLineAA, LineCapRound
                        End If
                    End If
                End If
                
            End If
            
        
        'Other selection types will be added in the future
        Case Else
    
    End Select
    
    'We now need to establish a bounding region for the selection.  For certain types of selections, we can do this with existing knowledge
    ' (e.g. the selLeft/Top/Width/Height values may reflect this).  For other types of selections, we need to find bounds via pixel searching.
    Select Case sShape
        
        'Rectangle and ellipse selections are easy; bounds have already been set by the width/height values
        Case sRectangle, sCircle
            
            If sType = sExterior Then
                boundLeft = 0
                boundTop = 0
                boundWidth = containingPDImage.Width
                boundHeight = containingPDImage.Height
            Else
                If sSmooth = sFullyFeathered Then
                    boundLeft = selLeft - featheringRadius
                    boundTop = selTop - featheringRadius
                    boundWidth = selWidth + featheringRadius * 2
                    boundHeight = selHeight + featheringRadius * 2
                Else
                    boundLeft = selLeft
                    boundTop = selTop
                    boundWidth = selWidth
                    boundHeight = selHeight
                End If
            End If
            
        'Lines are a bit weirder.  Because the two points are not sorted by distance (e.g. x1 may be greater than x2), we have to check that now.
        Case sLine

            If sType = sExterior Then
                boundLeft = 0
                boundTop = 0
                boundWidth = containingPDImage.Width
                boundHeight = containingPDImage.Height
            Else
                If sSmooth = sFullyFeathered Then
                    boundLeft = minX - (selectionLineWidth / 2) - featheringRadius
                    boundTop = MinY - (selectionLineWidth / 2) - featheringRadius
                    boundWidth = (maxX - minX) + selectionLineWidth + featheringRadius * 2
                    boundHeight = (MaxY - MinY) + selectionLineWidth + featheringRadius * 2
                Else
                    boundLeft = minX - (selectionLineWidth / 2)
                    boundTop = MinY - (selectionLineWidth / 2)
                    boundWidth = (maxX - minX) + selectionLineWidth
                    boundHeight = (MaxY - MinY) + selectionLineWidth
                End If
            End If
            
        Case Else
        
            'findBoundsManually()
        
    End Select
    
    'Do some basic bounds checking on the bound values to make sure they lie inside the image.
    If boundLeft < 0 Then
        boundWidth = boundWidth + boundLeft
        boundLeft = 0
    End If
    If boundTop < 0 Then
        boundHeight = boundHeight + boundTop
        boundTop = 0
    End If
    If boundLeft + boundWidth > containingPDImage.Width Then boundWidth = containingPDImage.Width - boundLeft
    If boundTop + boundHeight > containingPDImage.Height Then boundHeight = containingPDImage.Height - boundTop
    
    'Finally, if the selection is locked and feathering has been requested, apply it now
    If lockedIn And (sSmooth = sFullyFeathered) And (featheringRadius > 0) Then
    
        'Time to experiment with a new kind of selection feathering!  This time, we're going to take advantage of the fast drawing time
        ' for antialiased shapes to see if we can improve performance over the original Gaussian blur implementation.
        
        'If GDI+ 1.1 exists, use it for a faster blur operation.  If only v1.0 is found, fall back to PD's internal Gaussian blur function.
        If g_GDIPlusFXAvailable Then
        
            GDIPlusBlurLayer selMask, featheringRadius, boundLeft, boundTop, boundWidth, boundHeight
        
        Else
        
            'For exterior selections, the whole mask needs to be blurred.  No way around that.  But for interior and bordered selections,
            ' we can blur only the active area of the selection - saving precious time in the process.
            If (sType = sExterior) Or ((boundLeft = 0) And (boundTop = 0) And (boundWidth = containingPDImage.Width - 1) And (boundHeight = containingPDImage.Height - 1)) Then
                Dim tmpLayer As pdLayer
                Set tmpLayer = New pdLayer
                tmpLayer.createFromExistingLayer selMask
            
                Message "Feathering selection..."
                CreateGaussianBlurLayer featheringRadius, tmpLayer, selMask
                SetProgBarVal 0
                Message "Finished."
            Else
            
                'Create a layer the size of the active selection, and copy the selection mask into that
                Message "Feathering selection..."
                Dim srcLayer As pdLayer
                Set srcLayer = New pdLayer
                srcLayer.createBlank boundWidth, boundHeight, 24
                BitBlt srcLayer.getLayerDC, 0, 0, boundWidth, boundHeight, selMask.getLayerDC, boundLeft, boundTop, vbSrcCopy
                
                'Create a second layer (which will receive the blurred data)
                Dim dstLayer As pdLayer
                Set dstLayer = New pdLayer
                dstLayer.createBlank boundWidth, boundHeight, 24
                
                'Blur the selection
                CreateGaussianBlurLayer featheringRadius, srcLayer, dstLayer
                
                'Copy the blurred area back onto the selection mask
                BitBlt selMask.getLayerDC, boundLeft, boundTop, boundWidth, boundHeight, dstLayer.getLayerDC, 0, 0, vbSrcCopy
                
                'Erase both temporary layers
                srcLayer.eraseLayer
                Set srcLayer = Nothing
                dstLayer.eraseLayer
                Set dstLayer = Nothing
                
                SetProgBarVal 0
                Message "Finished."
            
            End If
        
        End If
    
    End If
    
    'Mark the mask as ready for us
    isMaskReady = True

End Sub

Public Sub renderTransformNodes(ByRef dstForm As Form, ByVal imgLeft As Long, ByVal imgTop As Long)
    
    'If an update has been requested, but we are already in the middle of one, deny subsequent requests
    If rejectRefreshRequests Then Exit Sub
    
    'Before drawing the nodes, we need to convert the selection's coordinates from "image coordinates" to "viewport coordinates".
    ' (Remember that the selection's location is stored as coordinates relative to the IMAGE ITSELF - but to render the transform nodes
    '   on the screen, we need to translate the coordinates to the current viewport - this makes them display accurately when scroll
    '   and zoom parameters are in effect.)
    Dim ZoomVal As Double
    ZoomVal = g_Zoom.ZoomArray(pdImages(dstForm.Tag).CurrentZoomValue)
    
    Dim hScrollVal As Long, vScrollVal As Long
    hScrollVal = dstForm.HScroll.Value
    vScrollVal = dstForm.VScroll.Value
    
    Dim dstLeft As Double, dstRight As Double
    Dim dstTop As Double, dstBottom As Double
    
    If (sShape = sRectangle) Or (sShape = sCircle) Then
        dstLeft = ((selLeft - hScrollVal) * ZoomVal) + imgLeft
        dstRight = ((selLeft + selWidth - hScrollVal) * ZoomVal) + imgLeft + 0.5
        dstTop = ((selTop - vScrollVal) * ZoomVal) + imgTop
        dstBottom = ((selTop + selHeight - vScrollVal) * ZoomVal) + imgTop + 0.5
    Else
        dstLeft = ((x1 - hScrollVal) * ZoomVal) + imgLeft
        dstRight = ((x2 - hScrollVal) * ZoomVal) + imgLeft + 0.5
        dstTop = ((y1 - vScrollVal) * ZoomVal) + imgTop
        dstBottom = ((y2 - vScrollVal) * ZoomVal) + imgTop + 0.5
    End If
    
    'If this selection is transformable, we can now draw transform circles at appropriate locations for the selection's shape.
    If isTransformable Then
        
        Dim circRadius As Long
        circRadius = 7
                
        Dim circAlpha As Long
        circAlpha = 190
        
        Select Case sShape
        
            Case sRectangle, sCircle
        
                'Draw corner circles first
                GDIPlusDrawCanvasCircle dstForm.hDC, dstLeft, dstTop, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstForm.hDC, dstRight, dstTop, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstForm.hDC, dstRight, dstBottom, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstForm.hDC, dstLeft, dstBottom, circRadius, circAlpha
                
                'Draw mid-line circles next
                Dim midHorizontal As Long, midVertical As Long
                midHorizontal = (dstLeft + dstRight) \ 2
                midVertical = (dstTop + dstBottom) \ 2
                
                GDIPlusDrawCanvasCircle dstForm.hDC, dstLeft, midVertical, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstForm.hDC, midHorizontal, dstTop, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstForm.hDC, dstRight, midVertical, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstForm.hDC, midHorizontal, dstBottom, circRadius, circAlpha
                
            Case sLine
            
                'Lines only need transform nodes at the end points of the line
                GDIPlusDrawCanvasCircle dstForm.hDC, dstLeft, dstTop, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstForm.hDC, dstRight, dstBottom, circRadius, circAlpha
                
        End Select
        
        dstForm.Picture = dstForm.Image
                
    End If
    
End Sub

'Render a selection using one of several methods.
Public Sub renderCustom(ByRef dstLayer As pdLayer, ByRef dstForm As Form, ByVal imgLeft As Long, ByVal imgTop As Long, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal renderType As SelectionRender)

    'If an update has been requested, but we are already in the middle of one, deny subsequent requests
    If rejectRefreshRequests Then Exit Sub

    'Is the selection mask ready to be used?  If not, prepare it now.
    If Not isMaskReady Then createSelectionMask
    
    'Before rendering the selection, we need to convert the its coordinates from "image coordinates" to "viewport coordinates".
    ' (Remember that the selection's location is stored as coordinates relative to the IMAGE ITSELF - but to render it on the screen,
    '   we need to translate the coordinates to the current viewport - this makes them display accurately when scroll and zoom are
    '   in effect.)
    Dim ZoomVal As Double
    ZoomVal = g_Zoom.ZoomArray(pdImages(dstForm.Tag).CurrentZoomValue)
    
    Dim hScrollVal As Long, vScrollVal As Long
    hScrollVal = dstForm.HScroll.Value
    vScrollVal = dstForm.VScroll.Value
    
    Dim dstLeft As Double, dstRight As Double
    Dim dstTop As Double, dstBottom As Double
        
    dstLeft = ((boundLeft - hScrollVal) * ZoomVal) + imgLeft
    dstRight = ((boundLeft + boundWidth - hScrollVal) * ZoomVal) + imgLeft + 0.5
    dstTop = ((boundTop - vScrollVal) * ZoomVal) + imgTop
    dstBottom = ((boundTop + boundHeight - vScrollVal) * ZoomVal) + imgTop + 0.5
    
    'Several different selection rendering methods use the same variables.  I've declared the common ones here.
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    Dim X As Long, Y As Long
    Dim QuickVal As Long
    
    Dim selMaskData() As Byte
    Dim selMaskSA As SAFEARRAY2D
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = containingPDImage.Width - 1
    srcHeight = containingPDImage.Height - 1
    
    'Just to be safe, declare some bounding variables in advance as well.  PhotoDemon now allows selections to lie off the image, so
    ' we must make sure we don't attempt to render selection pixels in any out-of-bound regions.
    Dim rightBound As Long, bottomBound As Long
    rightBound = imgLeft + imgWidth - 1
    bottomBound = imgTop + imgHeight - 1
            
    If rightBound > pdImages(dstForm.Tag).backBuffer.getLayerWidth - 1 Then rightBound = pdImages(dstForm.Tag).backBuffer.getLayerWidth - 1
    If bottomBound > pdImages(dstForm.Tag).backBuffer.getLayerHeight - 1 Then bottomBound = pdImages(dstForm.Tag).backBuffer.getLayerHeight - 1
    
    Dim xLookup() As Long, yLookup() As Long
    Dim tmpCalculation As Long
    
    Dim tmpSelHolder As Byte
    Dim tmpSelDivider As Double
    
    'Based on the selection type, draw a corresponding selection visualization.
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
            
            For X = 0 To 255
                hData(X) = X \ 3
            Next X
            
            dstRight = dstRight - 1
            dstBottom = dstBottom - 1
            
            'Build look-up tables for both the x and y direction.  This is faster than constantly calculating new values for
            ' every pixel inside the loop.
            ReDim xLookup(imgLeft To rightBound) As Long
            For X = imgLeft To rightBound
                tmpCalculation = hScrollVal + Int((X - imgLeft) / ZoomVal)
                If tmpCalculation < 0 Then tmpCalculation = 0
                If tmpCalculation > srcWidth Then tmpCalculation = srcWidth
                xLookup(X) = tmpCalculation * 3
            Next X
            
            ReDim yLookup(imgTop To bottomBound) As Long
            For Y = imgTop To bottomBound
                tmpCalculation = vScrollVal + Int((Y - imgTop) / ZoomVal)
                If tmpCalculation < 0 Then tmpCalculation = 0
                If tmpCalculation > srcHeight Then tmpCalculation = srcHeight
                yLookup(Y) = tmpCalculation
            Next Y
                        
            'Dim the non-active parts of the image
            For X = imgLeft To rightBound
                QuickVal = X * 3
            For Y = imgTop To bottomBound
            
                tmpSelHolder = selMaskData(xLookup(X), yLookup(Y))
                
                'Remember that lightbox works a little weirdly - UNSELECTED pixels are marked, while selected ones are left untouched
                If tmpSelHolder = 0 Then
                    iData(QuickVal + 2, Y) = hData(iData(QuickVal + 2, Y))
                    iData(QuickVal + 1, Y) = hData(iData(QuickVal + 1, Y))
                    iData(QuickVal, Y) = hData(iData(QuickVal, Y))
                ElseIf tmpSelHolder < 255 Then
                    tmpSelDivider = tmpSelHolder / 255
                    iData(QuickVal + 2, Y) = BlendColors(hData(iData(QuickVal + 2, Y)), iData(QuickVal + 2, Y), tmpSelDivider)
                    iData(QuickVal + 1, Y) = BlendColors(hData(iData(QuickVal + 1, Y)), iData(QuickVal + 1, Y), tmpSelDivider)
                    iData(QuickVal, Y) = BlendColors(hData(iData(QuickVal, Y)), iData(QuickVal, Y), tmpSelDivider)
                End If
                
            Next Y
            Next X
            
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
            
            For X = 0 To 255
                rData(X) = (newR + X) \ 2
                gData(X) = (newG + X) \ 2
                bData(X) = (newB + X) \ 2
            Next X
            
            dstRight = dstRight - 1
            dstBottom = dstBottom - 1
            
            'Build look-up tables for both the x and y direction.  This is faster than calculating
            ' such values for every pixel inside the loop.
            ReDim xLookup(imgLeft To rightBound) As Long
            For X = imgLeft To rightBound
                tmpCalculation = hScrollVal + Int((X - imgLeft) / ZoomVal)
                If tmpCalculation < 0 Then tmpCalculation = 0
                If tmpCalculation > srcWidth Then tmpCalculation = srcWidth
                xLookup(X) = tmpCalculation * 3
            Next X
            
            ReDim yLookup(imgTop To bottomBound) As Long
            For Y = imgTop To bottomBound
                tmpCalculation = vScrollVal + Int((Y - imgTop) / ZoomVal)
                If tmpCalculation < 0 Then tmpCalculation = 0
                If tmpCalculation > srcHeight Then tmpCalculation = srcHeight
                yLookup(Y) = tmpCalculation
            Next Y
            
            'Dim the non-active parts of the image
            For X = imgLeft To rightBound
                QuickVal = X * 3
            For Y = imgTop To bottomBound
            
                tmpSelHolder = selMaskData(xLookup(X), yLookup(Y))
                
                'If this pixels is over an active pixel in the selection mask, mark it with the selected color
                If tmpSelHolder = 255 Then
                    iData(QuickVal + 2, Y) = rData(iData(QuickVal + 2, Y))
                    iData(QuickVal + 1, Y) = gData(iData(QuickVal + 1, Y))
                    iData(QuickVal, Y) = bData(iData(QuickVal, Y))
                    
                'Partially selected pixels are rendered with antialiasing
                ElseIf tmpSelHolder > 0 Then
                    tmpSelDivider = tmpSelHolder / 255
                    iData(QuickVal + 2, Y) = BlendColors(iData(QuickVal + 2, Y), rData(iData(QuickVal + 2, Y)), tmpSelDivider)
                    iData(QuickVal + 1, Y) = BlendColors(iData(QuickVal + 1, Y), gData(iData(QuickVal + 1, Y)), tmpSelDivider)
                    iData(QuickVal, Y) = BlendColors(iData(QuickVal, Y), bData(iData(QuickVal, Y)), tmpSelDivider)
                End If
                
            Next Y
            Next X
            
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
            Dim dstDC As Long
            dstDC = dstLayer.getLayerDC
            
            'Since we'll be using the API to draw our selection area, we need to initialize several brushes
            Dim hPen As Long, hOldPen As Long
            
            hPen = CreatePen(PS_DOT, 0, RGB(0, 0, 0))
            hOldPen = SelectObject(dstDC, hPen)
            
            'Get a transparent brush
            Dim hBrush As Long, hOldBrush As Long
            hBrush = GetStockObject(NULL_BRUSH)
            hOldBrush = SelectObject(dstDC, hBrush)
            
            'Change the rasterOp to XOR (this will invert the line)
            SetROP2 dstDC, vbSrcInvert
                        
            'Draw the rectangle
            Rectangle dstDC, dstLeft, dstTop, dstRight, dstBottom
            
            'Restore the normal COPY rOp
            SetROP2 dstDC, vbSrcCopy
            
            'Remove the brush from the DC
            SelectObject dstDC, hOldBrush
            DeleteObject hBrush
            
            'Remove the pen from the DC
            SelectObject dstDC, hOldPen
            DeleteObject hPen
                    
    End Select
        
End Sub

Private Sub Class_Initialize()
    
    lockedIn = False
    
    'Initialize the selection mask for this object and mark it as "not ready" (because no mask has been drawn yet)
    Set selMask = New pdLayer
    isMaskReady = False
    
    'Note that a mask has never been created for this selection
    maskHasBeenCreated = False
    
    'Mark it as not transformable... yet
    isTransformable = False
    
    'By default, don't round off the corners of rectangular selections
    roundedCorners = 0
    
    'By default, selection lines have a width of 10
    selectionLineWidth = 1
    
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

'Read all relevant information about a selection from file, and populate this selection with the data.
' (NOTE: this function will not generate a selection mask or render the selection on-screen - it will only apply the stored values to
'        the relevant variables!  The calling function must explicitly request a redraw if desired.)
Public Function readSelectionFromFile(ByVal srcFilename As String, Optional ignoreLockStatus As Boolean = False) As Boolean

    If Not FileExist(srcFilename) Then
        Message "Requested selection file could not be located.  Selection load abandoned."
        readSelectionFromFile = False
        Exit Function
    End If

    Dim fileNum As Integer
    fileNum = FreeFile
    
    Dim layerLoadedSuccessfully As Boolean
    layerLoadedSuccessfully = False
    
    Open srcFilename For Binary As #fileNum
    
        'Check to make sure this is actually a PhotoDemon selection file
        Dim selectionIDCheck As String * 4
        Get #fileNum, 1, selectionIDCheck
        If (selectionIDCheck <> SELECTION_IDENTIFIER) Then
            Close #fileNum
            Debug.Print "Blank or missing selection mask found.  Abandoning selection load for performance reasons..."
            readSelectionFromFile = False
            Exit Function
        End If
        
        'Check to make sure this selection file verison is supported (not implemented right now, because there's only one selection version!)
        Dim selectionVersionCheck As Long
        Get #fileNum, , selectionVersionCheck
        
        'Size of the original image associated with this selection.  These values can be used to resize the selection mask as necessary if
        ' the selection is saved in raster format, and its original image was a different size from this one.
        Dim origWidth As Long, origHeight As Long
        Get #fileNum, , origWidth
        Get #fileNum, , origHeight
        
        'Retrieve the selection parameters into a string
        Dim paramString As String, lenParamString As Long
        Get #fileNum, , lenParamString
        paramString = String(lenParamString, " ")
        Get #fileNum, , paramString
        
        'Parse the parameter string into its individual components, and use those to initialize this selection
        initFromParamString paramString
        
        'Was the selection locked in?
        Get #fileNum, , lockedIn
        
        'Is the selection transformable?
        Get #fileNum, , isTransformable
        
        'If this selection was not transformable (meaning it could not be described in vector format), a full copy of its selection mask
        ' will be embedded in the file.  Load that data now.
        If Not isTransformable Then
        
            Dim tmpLayer As pdLayer
            Set tmpLayer = New pdLayer
            
            layerLoadedSuccessfully = tmpLayer.createFromFile("", True, fileNum)
            
            'Create a mask for this selection that is the size of the current image.  (Remember: the selection mask is always the full size
            ' of the image.)  Then, copy the contents of the embedded selection mask into this selection mask, with resizing as necessary.
            If layerLoadedSuccessfully Then selMask.createFromExistingLayer tmpLayer, containingPDImage.Width, containingPDImage.Height, True
            
            maskHasBeenCreated = True
            isMaskReady = True
            
        End If
        
    Close #fileNum
    
    'MsgBox "read from file: " & vbCrLf & srcFilename & vbCrLf & lockedIn & vbCrLf & paramString
    
    'If this selection wasn't transformable, we need to find its boundaries now
    If Not isTransformable And layerLoadedSuccessfully Then findNewBoundsManually
    
    'If this selection was locked in at the time of its save, lock this selection in as well
    If Not ignoreLockStatus Then
        If lockedIn Then
            lockIn
            containingPDImage.selectionActive = True
        Else
            lockRelease
            containingPDImage.selectionActive = False
        End If
    End If

    readSelectionFromFile = True
    
End Function

'Write all relevant information about this selection out to a file.
Public Function writeSelectionToFile(ByVal dstFilename As String) As Boolean

    'Delete any existing file (overwrite)
    If FileExist(dstFilename) Then Kill dstFilename
        
    Dim fileNum As Integer
    fileNum = FreeFile
    
    'Open the file and dump out only the essential information
    Open dstFilename For Binary As #fileNum
    
        'Identifiers
        Put #fileNum, 1, SELECTION_IDENTIFIER
        Put #fileNum, , SELECTION_FILE_VERSION_2013
        
        'Size of the attached image.  This is used to resize the selection mask as necessary if the selection is saved in raster format,
        ' and its original image was a different size from this one.
        Put #fileNum, , containingPDImage.Width
        Put #fileNum, , containingPDImage.Height
        
        'Sometimes we write the selection's current state to file.  Other times (Undo/Redo, specifically) we write out the selection's
        ' previous state.  Check for that and write the desired data as necessary.
        Dim writeSelectionState As Boolean, writeTransformable As Boolean
        Dim writeParamString As String
        Dim writeLayer As pdLayer
        Set writeLayer = New pdLayer
        
        If Len(backupParamString) > 0 Then
            writeSelectionState = backupSelectionState
            writeTransformable = backupTransformable
            writeParamString = backupParamString
            If Not writeTransformable Then writeLayer.createFromExistingLayer backupMask
        Else
            writeSelectionState = lockedIn
            writeTransformable = isTransformable
            writeParamString = getSelectionParamString
            If Not writeTransformable Then writeLayer.createFromExistingLayer selMask
        End If
        
        Put #fileNum, , CLng(Len(writeParamString))
        Put #fileNum, , writeParamString
        
        'Write out the selection's current stated (locked in or not), as well as its transformability.
        Put #fileNum, , writeSelectionState
        Put #fileNum, , writeTransformable
        
        'If this selection is not transformable, we must also embed the full selection mask.
        If Not writeTransformable And maskHasBeenCreated Then writeLayer.writeToFile "", True, True, fileNum
        
    Close #fileNum
    
    'Debug.Print "wrote to file: " & vbCrLf & dstFilename & vbCrLf & paramString
    
    'Clear the backup parameter string
    backupParamString = ""
    
    writeSelectionToFile = True
    
End Function

'By default, a pdSelection's parameter string contains absolute coordinates (e.g. "100, 100" for x1, y1).  When saving to
' file or recording as part of a macro, we want to use relative formats (e.g. "0.5, 0.5" for a coordinate at the center of
' an image).  This allows selections to be correctly applied to images of any size.  Thus, when writing to or reading from
' a selection file, or when creating a paramstring for macro purposes, this function will be called; it automatically handles
' the translation of all coordinate values between absolute and relative geometries.
Private Function translateParamStringToRelativeFormat(ByVal paramString As String, ByVal toRelativeFormat As Boolean) As String

    Dim tmpParser As pdParamString
    Set tmpParser = New pdParamString
    tmpParser.setParamString paramString
    
    'Convert absolute values to relative values (typically used when WRITING a selection file)
    If toRelativeFormat Then
    
        '08 - left
        tmpParser.setNewValueAtLocation 8, CStr(CDbl(tmpParser.GetLong(8)) / CDbl(containingPDImage.Width))
        
        '09 - top
        tmpParser.setNewValueAtLocation 9, CStr(CDbl(tmpParser.GetLong(9)) / CDbl(containingPDImage.Height))
        
        '10 - width
        tmpParser.setNewValueAtLocation 10, CStr(CDbl(tmpParser.GetLong(10)) / CDbl(containingPDImage.Width))
        
        '11 - height
        tmpParser.setNewValueAtLocation 11, CStr(CDbl(tmpParser.GetLong(11)) / CDbl(containingPDImage.Height))
        
        '12 - x1
        tmpParser.setNewValueAtLocation 12, CStr(CDbl(tmpParser.GetLong(12)) / CDbl(containingPDImage.Width))
        
        '13 - y1
        tmpParser.setNewValueAtLocation 13, CStr(CDbl(tmpParser.GetLong(13)) / CDbl(containingPDImage.Height))
        
        '14 - x2
        tmpParser.setNewValueAtLocation 14, CStr(CDbl(tmpParser.GetLong(14)) / CDbl(containingPDImage.Width))
        
        '15 - y2
        tmpParser.setNewValueAtLocation 15, CStr(CDbl(tmpParser.GetLong(15)) / CDbl(containingPDImage.Height))
    
    'Convert relative values to absolute values (typically used when READING a selection file)
    Else
    
        '08 - left
        tmpParser.setNewValueAtLocation 8, CStr(CLng(tmpParser.GetDouble(8) * CDbl(containingPDImage.Width)))
        
        '09 - top
        tmpParser.setNewValueAtLocation 9, CStr(CLng(tmpParser.GetDouble(9) * CDbl(containingPDImage.Height)))
        
        '10 - width
        tmpParser.setNewValueAtLocation 10, CStr(CLng(tmpParser.GetDouble(10) * CDbl(containingPDImage.Width)))
        
        '11 - height
        tmpParser.setNewValueAtLocation 11, CStr(CLng(tmpParser.GetDouble(11) * CDbl(containingPDImage.Height)))
        
        '12 - x1
        tmpParser.setNewValueAtLocation 12, CStr(CLng(tmpParser.GetDouble(12) * CDbl(containingPDImage.Width)))
        
        '13 - y1
        tmpParser.setNewValueAtLocation 13, CStr(CLng(tmpParser.GetDouble(13) * CDbl(containingPDImage.Height)))
        
        '14 - x2
        tmpParser.setNewValueAtLocation 14, CStr(CLng(tmpParser.GetDouble(14) * CDbl(containingPDImage.Width)))
        
        '15 - y2
        tmpParser.setNewValueAtLocation 15, CStr(CLng(tmpParser.GetDouble(15) * CDbl(containingPDImage.Height)))
    
    End If
    
    'Rebuild the parameter string
    tmpParser.reassembleModifiedParamString
    
    'Return the modified string
    translateParamStringToRelativeFormat = tmpParser.getParamString

End Function

'Invert the current selection.  Note that this will make a transformable selection non-transformable - to maintain transformability, use
' the "exterior"/"interior" options on the main form.
Public Sub invertSelection()

    Message "Inverting selection..."
    
    'Point a standard 2D byte array at the selection mask
    Dim X As Long, Y As Long
    Dim QuickVal As Long
    
    Dim selMaskData() As Byte
    Dim selMaskSA As SAFEARRAY2D
    prepLayerSafeArray selMask, selMaskSA
    CopyMemory ByVal VarPtrArray(selMaskData()), VarPtr(selMaskSA), 4
    
    Dim maskWidth As Long, maskHeight As Long
    maskWidth = selMask.getLayerWidth - 1
    maskHeight = selMask.getLayerHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    SetProgBarMax maskWidth
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'After all that work, the Invert code itself is very small and unexciting!
    For X = 0 To maskWidth
        QuickVal = X * 3
    For Y = 0 To maskHeight
        selMaskData(QuickVal, Y) = 255 - selMaskData(QuickVal, Y)
        selMaskData(QuickVal + 1, Y) = 255 - selMaskData(QuickVal + 1, Y)
        selMaskData(QuickVal + 2, Y) = 255 - selMaskData(QuickVal + 2, Y)
    Next Y
        If (X And progBarCheck) = 0 Then
            'If userPressedESC() Then Exit For
            SetProgBarVal X
        End If
    Next X
    
    'Release our temporary byte array
    CopyMemory ByVal VarPtrArray(selMaskData), 0&, 4
    Erase selMaskData
    
    'We now need to find selection boundaries.  Use a standalone function for that.
    findNewBoundsManually
    
    'Inverting a selection makes it non-transformable
    isTransformable = False
    
    'Mark the mask as ready (it should already be marked this way, but set it again just to be safe)
    isMaskReady = True
    
    SetProgBarVal 0
    Message "Selection inversion complete."

End Sub

'When working with raster selections (e.g. non-transformable ones), we still want to minimize selection processing time by processing the
' smallest possible rectangle that includes all selected pixels.  This function will scan the selection mask and populate the
' selLeft/Top/Width/Height and boundLeft/Top/Width/Height values automatically, based on the mask's contents.
Public Sub findNewBoundsManually()

    Message "Finding new selection boundaries..."

    'This function assumes two things: that a selection is not transformable (otherwise we'd know the boundaries already), and
    ' a mask has already been created.  If either of these two conditions is not met, this function may fail.
    isTransformable = False
    isMaskReady = True

    'Point a standard 2D byte array at the selection mask
    Dim X As Long, Y As Long
    Dim QuickVal As Long
    
    Dim selMaskData() As Byte
    Dim selMaskSA As SAFEARRAY2D
    prepLayerSafeArray selMask, selMaskSA
    CopyMemory ByVal VarPtrArray(selMaskData()), VarPtr(selMaskSA), 4
    
    Dim maskWidth As Long, maskHeight As Long
    maskWidth = selMask.getLayerWidth - 1
    maskHeight = selMask.getLayerHeight - 1
    
    Dim boundFound As Boolean
    
    'Find the top bound first.
    boundFound = False
    Y = 0
    Do
    
        For X = 0 To maskWidth
            QuickVal = X * 3
            If selMaskData(QuickVal, Y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(QuickVal + 1, Y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(QuickVal + 2, Y) <> 0 Then
                boundFound = True
            End If
            
            If boundFound Then
                selTop = Y
                boundTop = selTop
                Exit For
            End If
            
        Next X
        
        'Blank masks shouldn't be possible, but if they are, mark the boundary as found to prevent a program lock
        If (Not boundFound) And (Y = maskHeight) Then
            Debug.Print "No top boundary found - is mask blank?"
            boundFound = True
            selTop = 0
            boundTop = selTop
        End If
        
        Y = Y + 1
    
    Loop While Not boundFound
    
    'Next, find the bottom bound
    boundFound = False
    Y = maskHeight
    
    Do
    
        For X = 0 To maskWidth
            QuickVal = X * 3
            If selMaskData(QuickVal, Y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(QuickVal + 1, Y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(QuickVal + 2, Y) <> 0 Then
                boundFound = True
            End If
            
            If boundFound Then
                selHeight = Y - selTop + 1
                boundHeight = selHeight
                Exit For
            End If
            
        Next X
        
        'Blank masks shouldn't be possible, but if they are, mark the boundary as found to prevent a program lock
        If (Not boundFound) And (Y = 0) Then
            Debug.Print "No bottom boundary found - is mask blank?"
            boundFound = True
            selHeight = maskHeight - selTop + 1
            boundHeight = selHeight
        End If
        
        Y = Y - 1
    
    Loop While Not boundFound
    
    'Next, find the left bound
    boundFound = False
    X = 0
    Do
    
        For Y = 0 To maskHeight
            If selMaskData(X, Y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(X + 1, Y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(X + 2, Y) <> 0 Then
                boundFound = True
            End If
            
            If boundFound Then
                selLeft = X \ 3
                boundLeft = selLeft
                Exit For
            End If
            
        Next Y
        
        'Blank masks shouldn't be possible, but if they are, mark the boundary as found to prevent a program lock
        If (Not boundFound) And (X = maskWidth * 3) Then
            Debug.Print "No left boundary found - is mask blank?"
            boundFound = True
            selLeft = 0
            boundLeft = selLeft
        End If
        
        X = X + 3
    
    Loop While Not boundFound
    
    'Finally, find the right bound
    boundFound = False
    X = maskWidth * 3
    
    Do
    
        For Y = 0 To maskHeight
            If selMaskData(X, Y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(X + 1, Y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(X + 2, Y) <> 0 Then
                boundFound = True
            End If
            
            If boundFound Then
                selWidth = (X \ 3) - selLeft + 1
                boundWidth = selWidth
                Exit For
            End If
            
        Next Y
        
        'Blank masks shouldn't be possible, but if they are, mark the boundary as found to prevent a program lock
        If (Not boundFound) And (X = 0) Then
            Debug.Print "No right boundary found - is mask blank?"
            boundFound = True
            selWidth = maskWidth - selLeft + 1
            boundWidth = selWidth
        End If
        
        X = X - 3
        
    Loop While Not boundFound
    
    'All selection boundaries have now been located
    
    'Release our temporary byte array and exit
    CopyMemory ByVal VarPtrArray(selMaskData), 0&, 4
    Erase selMaskData
    
End Sub

'Sharpen (un-feather?) the current selection.  In the future, this code will be replaced with a much simpler "sharpen layer" function, but until
' that happens, a copy-and-paste implementation is provided here for ease of use.
Public Sub sharpenSelection(ByVal sharpenRadius As Long)
        
    'Point an array at the current selection mask
    Dim selMaskData() As Byte
    Dim selMaskSA As SAFEARRAY2D
    
    'Create a second local array.  This will contain the a copy of the selection mask, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer selMask
            
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long
    
    'Unsharp masking requires a gaussian blur layer to operate.  Create one now.
    If CreateGaussianBlurLayer(sharpenRadius, selMask, srcLayer, False, selMask.getLayerHeight * 2 + selMask.getLayerWidth) Then
    
        'Now that we have a gaussian layer created in workingLayer, we can point arrays toward it and the source layer
        prepLayerSafeArray selMask, selMaskSA
        CopyMemory ByVal VarPtrArray(selMaskData()), VarPtr(selMaskSA), 4
        
        Dim srcImageData() As Byte
        Dim srcSA As SAFEARRAY2D
        prepLayerSafeArray srcLayer, srcSA
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim QuickVal As Long, qvDepth As Long
        qvDepth = 3
        
        'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
        ' based on the size of the area to be processed.
        Dim progBarCheck As Long
        progBarCheck = findBestProgBarValue()
            
        'ScaleFactor is used to apply the unsharp mask.  Maximum strength can be any value, but PhotoDemon locks it at 10.
        Dim scaleFactor As Double, invScaleFactor As Double
        scaleFactor = sharpenRadius
        invScaleFactor = 1 - scaleFactor
    
        Dim blendVal As Double
        
        'More color variables - in this case, sums for each color component
        Dim r As Long, g As Long, b As Long
        Dim r2 As Long, g2 As Long, b2 As Long
        Dim newR As Long, newG As Long, newB As Long
        Dim tLumDelta As Long
        
        'The final step of the smart blur function is to find edges, and replace them with the blurred data as necessary
        For X = 0 To selMask.getLayerWidth - 1
            QuickVal = X * 3
        For Y = 0 To selMask.getLayerHeight - 1
            
            'Retrieve the original image's pixels
            r = selMaskData(QuickVal + 2, Y)
            g = selMaskData(QuickVal + 1, Y)
            b = selMaskData(QuickVal, Y)
            
            'Now, retrieve the gaussian pixels
            r2 = srcImageData(QuickVal + 2, Y)
            g2 = srcImageData(QuickVal + 1, Y)
            b2 = srcImageData(QuickVal, Y)
            
            tLumDelta = Abs(getLuminance(r, g, b) - getLuminance(r2, g2, b2))
                
            newR = (scaleFactor * r) + (invScaleFactor * r2)
            If newR > 255 Then newR = 255
            If newR < 0 Then newR = 0
                
            newG = (scaleFactor * g) + (invScaleFactor * g2)
            If newG > 255 Then newG = 255
            If newG < 0 Then newG = 0
                
            newB = (scaleFactor * b) + (invScaleFactor * b2)
            If newB > 255 Then newB = 255
            If newB < 0 Then newB = 0
            
            blendVal = tLumDelta / 255
            
            newR = BlendColors(newR, r, blendVal)
            newG = BlendColors(newG, g, blendVal)
            newB = BlendColors(newB, b, blendVal)
            
            selMaskData(QuickVal + 2, Y) = newR
            selMaskData(QuickVal + 1, Y) = newG
            selMaskData(QuickVal, Y) = newB
                    
        Next Y
            If (X And progBarCheck) = 0 Then
                SetProgBarVal X + (selMask.getLayerHeight * 2)
            End If
        Next X
        
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        Erase srcImageData
        
        srcLayer.eraseLayer
        Set srcLayer = Nothing
        
        CopyMemory ByVal VarPtrArray(selMaskData), 0&, 4
        Erase selMaskData
        
    End If
        
End Sub

'Convert the selection to border-type
Public Sub borderSelection(ByVal borderRadius As Long)

    'Bordering a selection requires two passes: a grow pass and a shrink pass.  The results of these two passes are then blended
    ' to create the final bordered selection.
    
    'Start by creating the grow and shrink layers using the median function.
    
    'Create a second local array.  This will contain the a copy of the selection mask, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim growLayer As pdLayer
    Set growLayer = New pdLayer
    growLayer.createFromExistingLayer selMask
    
    Dim shrinkLayer As pdLayer
    Set shrinkLayer = New pdLayer
    shrinkLayer.createFromExistingLayer selMask
    
    'Unsharp masking requires a gaussian blur layer to operate.  Create one now.
    CreateMedianLayer borderRadius, 1, selMask, shrinkLayer, False, selMask.getLayerWidth * 2
    CreateMedianLayer borderRadius, 100, selMask, growLayer, False, selMask.getLayerWidth * 2, selMask.getLayerWidth
    
    'Blend those two layers over the existing mask
    selMask.createFromExistingLayer growLayer
    growLayer.eraseLayer
    BitBlt selMask.getLayerDC, 0, 0, selMask.getLayerWidth, selMask.getLayerHeight, shrinkLayer.getLayerDC, 0, 0, vbSrcInvert
    
    'Erase the temporary layers
    shrinkLayer.eraseLayer
        
End Sub
