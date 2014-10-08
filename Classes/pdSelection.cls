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
'Copyright ©2012-2014 by Tanner Helland
'Created: 25/September/12
'Last updated: 05/October/14
'Last update: initial build of lasso selections
'
'This class handles all selections in PhotoDemon.  Rectangular and elliptical selections are now implemented, with more
' to follow in the coming months.
'
'All selections are treated as a subset of pixels within a rectangular area. The x1, y1, x2, y2 coordinates define opposite
' corners of the selection area. There is no guarantee that these coordinates will appear in a given corner - so they must
' be checked manually. When a selection is set, those values are translated into Left, Top, Width and Height values, which
' are in turn used for everything from cropping to applying filters.
'
'Note also that this class is treated as a subset of pdImage(). Right now each image object only contains one selection,
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
Private Const SELECTION_FILE_VERSION_2014 As Long = &H1002

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

'Rectangle drawing
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'What shape does this selection have?
Public Enum SelectionShape
    sRectangle = 0
    sCircle = 1
    sLine = 2
    sLasso = 3
    sRaster = 10
End Enum

#If False Then
    Const sRectangle = 0
    Const sCircle = 1
    Const sLine = 2
    Const sLasso = 3
    Const sRaster = 10
#End If

Private sShape As SelectionShape

'What type of selection is this?
Public Enum SelectionType
    sInterior = 0
    sExterior = 1
    sBorder = 2
End Enum

#If False Then
    Const sInterior = 0, sExterior = 1, sBorder = 2
#End If

Private sType As SelectionType

'What kind of smoothing (feathering) will be applied to this selection?
Public Enum SelectionSmoothing
    sNone = 0
    sAntialiased = 1
    sFullyFeathered = 2
End Enum

#If False Then
    Const sNone = 0, sAntialiased = 1, sFullyFeathered = 2
#End If

Private sSmooth As SelectionSmoothing

'Is a special drawing display mode in use?
Public Enum SelectionDrawingDisplay
    sNormalDisplay = 0
    sOutlineOnlyDisplay = 1
End Enum

#If False Then
    Const sNormalDisplay = 0, sOutlineOnlyDisplay = 1
#End If

Private sDrawingDisplay As SelectionDrawingDisplay

'The coordinates of the current selection (if rectangular)
Public x1 As Double, x2 As Double
Public y1 As Double, y2 As Double

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
Public selMask As pdDIB

'Until a mask has been created, this value will remain "false".  It is used to optimize operations on blank selections.
Public maskHasBeenCreated As Boolean

'After the mask has been generated, its status will be set as "ready".  Any function that changes the selected area will reset this to "false".
' The selection rendering code can check this value to see if the mask needs to be recreated before rendering it to screen.  If external
' functions modify the selection in some way (e.g. Select menu dialogs), they need to set this to TRUE to prevent the engine from attempting
' to recreate a mask from existing vector data.
Public isMaskReady As Boolean

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

'Lasso/Polygon selections have a variable number of points.  As such, we have to track their contents dynamically.
Private m_numOfLassoPoints As Long
Private m_LassoPoints() As POINTFLOAT

'When transforms are applied to a lasso selection, we must maintain a backup copy of the original lasso curve.  Any transformation
' values are then applied to this backup data, and the backup is erased when the mouse is released.
Private m_LassoPointsBackup() As POINTFLOAT

'Lasso selections support rendering via both lines and curves.  If sSmoothStroke is 0, the lasso is rendered as a closed polygon.
' Increasing smooth stroke values affect the tension of the cardinal spline used to enclose the region.
Private sSmoothStroke As Double

'To save code elsewhere, a selection can be initialized by using a param string generated by a pdParamString object.  This condenses a huge
' list of variables into a single string.  This function is especially helpful when writing full selection data to/from file.
Public Sub initFromParamString(ByVal paramString As String)

    Dim i As Long

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
    
    '9 - top
    selTop = cParams.GetLong(9)
    
    '10 - width
    selWidth = cParams.GetLong(10)
    
    '11 - height
    selHeight = cParams.GetLong(11)
    
    'Beyond this point, parameters vary by selection type
    Select Case sShape
    
        'Rectangles, ellipses, and lines only require bounding rect values (stored as two (x, y) pairs).  Retrieve these
        ' values from positions [12, 15] inclusive.
        Case sRectangle, sCircle, sLine
        
            x1 = cParams.GetDouble(12)
            y1 = cParams.GetDouble(13)
            x2 = cParams.GetDouble(14)
            y2 = cParams.GetDouble(15)
        
        'Lasso/freehand selections have a variable number of parameters, based on the number of points in the lasso.
        ' Position (13) stores the number of points.
        Case sLasso
            
            '12 - Draw mode display
            setDrawModeDisplay cParams.GetLong(12)
            
            '13 - Smooth stroke
            setSmoothStroke cParams.GetDouble(13)
            
            '14 - number of points in the lasso array
            m_numOfLassoPoints = cParams.GetLong(14)
            
            If m_numOfLassoPoints > 0 Then
            
                ReDim m_LassoPoints(0 To m_numOfLassoPoints - 1) As POINTFLOAT
                
                'Retrieve all remaining lasso points
                If m_numOfLassoPoints > 0 Then
                    For i = 0 To m_numOfLassoPoints - 1
                        m_LassoPoints(i).x = cParams.GetDouble(15 + i * 2)
                        m_LassoPoints(i).y = cParams.GetDouble(16 + i * 2)
                    Next i
                End If
                
            Else
                ReDim m_LassoPoints(0 To 127) As POINTFLOAT
            End If
        
        'Other types (e.g. raster selections) cannot be initiated this way, so we can ignore them.
        Case Else
    
    End Select
        
End Sub

'Return all of this selection's important settings as a specifically formatted parameter string.  This string can be passed to a selection
' object at any future point to recreate this selection exactly.
' (See also: "initFromParamString" above)
Public Function getSelectionParamString() As String
    
    Dim i As Long
    
    'All selection types store a preset list of standardized values
    getSelectionParamString = buildParams(sShape, sType, sSmooth, featheringRadius, borderSize, roundedCorners, selectionLineWidth, selLeft, selTop, selWidth, selHeight)
    
    'After the preset values, each selection type can write its own coords, relevant to its particular shape
    Select Case sShape
    
        'Rectangles, ellipses, and lines require only (x1, y1) and (x2, y2) coordinates
        Case sRectangle, sCircle, sLine
            getSelectionParamString = getSelectionParamString & "|" & buildParams(x1, y1, x2, y2)
            
        'Lasso selections are trickier; they have a dynamic length based on the number of drawn points.
        ' (TODO: use a string builder class for this, to improve performance)
        Case sLasso
            
            getSelectionParamString = getSelectionParamString & "|" & CStr(sDrawingDisplay) & "|" & Trim$(Str(sSmoothStroke)) & "|" & CStr(m_numOfLassoPoints)
            
            If m_numOfLassoPoints > 0 Then
                For i = 0 To m_numOfLassoPoints - 1
                    getSelectionParamString = getSelectionParamString & "|" & Trim$(Str(m_LassoPoints(i).x)) & "|" & Trim$(Str(m_LassoPoints(i).y))
                Next i
            End If
            
        'Raster selections cannot be initiated this way!
        Case sRaster
    
    End Select
    
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
    If (sShape <> sRaster) And (sShape <> sLasso) Then
    
        isMaskReady = False
    
        'Redraw the selection if necessary
        If forceSelectionRefresh And (Not containingPDImage Is Nothing) Then
            updateInternalCoords
            createSelectionMask
        End If
        
    End If
    
End Sub

'Request a redraw of the selection mask.  We must do this when loading an Undo or Redo request after the image size has
' been changed; otherwise, OOB errors can occur (because the selection mask will be a different size than the image).
Public Sub requestNewMask()

    If (Not containingPDImage Is Nothing) Then
        
        updateInternalCoords
        
        'Transformable selections are rendered using polygon geometry; as such, redrawing them requires invoking
        ' whatever polygon code we used previously.
        If (sShape <> sRaster) Then
            createSelectionMask
            
        'Raster selections are simply bitmaps.  We will pad them to match the new image size, but other
        ' than that, we do not change their existing composition.
        Else
        
        End If
        
    End If

End Sub

'Get/set feathering radius
Public Function getFeatheringRadius() As Long
    getFeatheringRadius = featheringRadius
End Function

Public Sub setFeatheringRadius(ByVal newRadius As Long)
    If (sShape <> sRaster) Then isMaskReady = False
    featheringRadius = newRadius
End Sub

'Get/set the border size
Public Function getBorderSize() As Long
    getBorderSize = borderSize
End Function

Public Sub setBorderSize(ByVal newBorderSize As Long)
    If (sShape <> sRaster) Then isMaskReady = False
    borderSize = newBorderSize
End Sub

'Get/set a transformation type
Public Function getTransformationType()
    If transformModeActive Then
        getTransformationType = selTransformationType
    Else
        getTransformationType = -1
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
            
        Case sLasso
            isTransformable = False
        
    End Select
    
    If (sShape <> sRaster) Then isMaskReady = False
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
    If (sShape <> sRaster) Then isMaskReady = False
    sSmooth = selSmooth
End Sub

'Change the amount of corner rounding for rectangular selections
Public Function getRoundedCornerAmount() As Long
    getRoundedCornerAmount = roundedCorners
End Function

Public Sub setRoundedCornerAmount(ByVal cornerRounding As Long)
    If (sShape <> sRaster) Then isMaskReady = False
    roundedCorners = cornerRounding
End Sub

'Get/set the line width of line selections
Public Function getSelectionLineWidth() As Long
    getSelectionLineWidth = selectionLineWidth
End Function

Public Sub setSelectionLineWidth(ByVal newLineWidth As Long)
    If (sShape <> sRaster) Then isMaskReady = False
    selectionLineWidth = newLineWidth
End Sub

'Get/set special drawing mode views.  Note that these are only relevant for certain selection types.
Public Function getDrawModeDisplay() As SelectionDrawingDisplay
    getDrawModeDisplay = sDrawingDisplay
End Function

Public Sub setDrawModeDisplay(ByVal newDisplayMode As SelectionDrawingDisplay)
    sDrawingDisplay = newDisplayMode
End Sub

'Get/set smooth stroke mode.  Note that this is only relevant for certain selection types.
Public Function getSmoothStroke() As Double
    getSmoothStroke = sSmoothStroke
End Function

Public Sub setSmoothStroke(ByVal newSmoothness As Double)
    If (sShape <> sRaster) Then isMaskReady = False
    sSmoothStroke = newSmoothness
End Sub

'For lasso and polygon selections, this function will return the current selecton region as a GDI+ region handle
Public Function getGdipRegionForSelection() As Long
    getGdipRegionForSelection = GDI_Plus.getGDIPlusRegionFromPoints(m_numOfLassoPoints, VarPtr(m_LassoPoints(0)), fillmodewinding)
End Function

'Takes x and y coordinates (from a _MouseDown event, typically) and uses them in a manner specified by the current transform operation.
' Note that this this should only be called after a transformation type has been set.
Public Sub setInitialTransformCoordinates(ByVal x As Double, y As Double)

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
            
                'Case -1 should never occur, because -1 means "no point of interest clicked"
                Case -1
                    Debug.Print "Selection transform initiated on a non-existent point - FIX THIS!"
                    
                '0 - NW corner
                Case 0
                    x1 = selLeft + selWidth
                    y1 = selTop + selHeight
                    x2 = selLeft
                    y2 = selTop
                
                '1 - NE corner
                Case 1
                    x1 = selLeft
                    y1 = selTop + selHeight
                    x2 = x
                    y2 = y
                
                '2 - SE corner
                Case 2
                    x1 = selLeft
                    y1 = selTop
                    x2 = x
                    y2 = y
                
                '3 - SW corner
                Case 3
                    x1 = selLeft + selWidth
                    y1 = selTop
                    x2 = x
                    y2 = y
                
                '4 - N edge
                Case 4
                    x1 = selLeft
                    x2 = selLeft + selWidth
                    y1 = selTop + selHeight
                    y2 = y
                
                '5 - E edge
                Case 5
                    x1 = selLeft
                    x2 = x
                    y1 = selTop
                    y2 = selTop + selHeight
                
                '6 - S edge
                Case 6
                    x1 = selLeft
                    x2 = selLeft + selWidth
                    y1 = selTop
                    y2 = y
                
                '7 - W edge
                Case 7
                    x1 = selLeft + selWidth
                    x2 = x
                    y1 = selTop
                    y2 = selTop + selHeight
                
                '8 - interior of selection, not near a corner or edge
                Case 8
                    moveXDist = x - selLeft
                    moveYDist = y - selTop
            
            End Select
        
        
        'Line selections
        Case sLine
        
            'Based on the transform mode, set the initial points accordingly
            Select Case selTransformationType
        
                'Case -1 should never occur, but if it does - treat this like a normal initial coordinate call
                Case -1
                    x1 = x
                    y1 = y
                    x2 = x
                    y2 = y
                
                'First point is being moved
                Case 0
                    x1 = x
                    y1 = y
                
                'Second point is being moved
                Case 1
                    x2 = x
                    y2 = y
                    
            End Select
                
        
        'Lasso selections can only be moved, so the only valid transform type is type 0
        Case sLasso
        
            'Based on the transform mode, set the initial points accordingly
            Select Case selTransformationType
        
                '0 - interior of the selection, which will trigger a move transformation
                Case 0
                    moveXDist = x
                    moveYDist = y
                    
                    'For the transform to work, we need to apply any transformed data to the *original lasso points*.
                    ' Make a backup copy of the original array now.
                    ReDim m_LassoPointsBackup(0 To m_numOfLassoPoints - 1) As POINTFLOAT
                    CopyMemory ByVal VarPtr(m_LassoPointsBackup(0)), ByVal VarPtr(m_LassoPoints(0)), m_numOfLassoPoints * 8
                    
            End Select
        
        'Some selection types cannot be transformed!
        Case sRaster
        
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
    
    'Set the second set of point to match the first set
    Select Case sShape
        
        Case sRectangle, sCircle, sLine
            x2 = x
            y2 = y
            
            'Also, mark selections of any of these shapes (rectangle, circle, line) as transformable
            isTransformable = True
        
        'Lasso selections are non-transformable, and they have a variable number of points
        Case sLasso
            isTransformable = False
            
            m_numOfLassoPoints = 1
            ReDim m_LassoPoints(0 To 127) As POINTFLOAT
            
            m_LassoPoints(0).x = x
            m_LassoPoints(0).y = y
            
            'Finding the bounding rect of a lasso selection is a little more complicated than other selections,
            ' because we have to search all points to find the full bounding rect.  To avoid the need for that,
            ' we'll keep running track as we go - so start out by setting our bounding rect to the initial coordinates.
            selLeft = x
            selTop = y
            selWidth = 1
            selHeight = 1
        
        'Other selection types (e.g. raster selections) cannot be created this way
        Case Else
        
    End Select
    
    updateInternalCoords
    
End Sub

'Takes x and y coordinates (from a _MouseMove event, typically) and stores them internally
Public Sub setAdditionalCoordinates(ByVal x As Double, y As Double)
    
    'Check for an active transformation mode.  (A transformation is something like resizing or moving an existing selection,
    ' versus drawing a new one from scratch.)
    If transformModeActive Then
        
        Dim transX As Double, transY As Double
        Dim i As Long
        
        Select Case sShape
    
            Case sRectangle, sCircle
            
                'Based on the transform mode, set the initial points accordingly
                Select Case selTransformationType
                
                    'Case -1 should never occur, but if it does - treat this like a normal subsequent coordinate call
                    'Cases 1-4 can similarly be treated like a normal subsequent coordinate call
                    ' (See setInitialTransformCoordinates for details on what the different cases mean.  And yes, I should convert
                    '  these to constants instead of numbers... :P)
                    Case 0 To 3
                        x2 = x
                        y2 = y
                    
                    '4 - N edge
                    Case 4
                        y2 = y
                        
                    '5 - E edge
                    Case 5
                        x2 = x
                        
                    '6 - S edge
                    Case 6
                        y2 = y
                        
                    '7 - W edge
                    Case 7
                        x2 = x
                        
                    '8 - interior of selection, not near a corner or edge (e.g. move the selection, but don't resize it)
                    Case 8
                    
                        x1 = x - moveXDist
                        y1 = y - moveYDist
                        x2 = x1 + selWidth
                        y2 = y1 + selHeight
                        
                End Select
                
                'If a transform mode is active, re-mark the selection as being transformable
                isTransformable = True
        
            Case sLine
            
                'Based on the transform mode, set the initial points accordingly
                Select Case selTransformationType
            
                    'Case -1 should never occur.  (-1 represents an invalid transformation request)
                    Case -1
                    
                    'First point is being moved
                    Case 0
                        x1 = x
                        y1 = y
                        
                    'Second point is being moved
                    Case 1
                        x2 = x
                        y2 = y
                        
                End Select
                
                'If a transform mode is active, re-mark the selection as being transformable
                isTransformable = True
            
            'Lasso transforms require us to transform the entire lasso array
            Case sLasso
            
                'Based on the transform mode, set the initial points accordingly
                Select Case selTransformationType
            
                    'Case -1 should never occur.  (-1 represents an invalid transformation request)
                    Case -1
                    
                    'Move transform
                    Case 0
                        
                        'Rebuild the main lasso array by copying all points from the backup array, and applying the current
                        ' x/y transformation distance to them.
                        For i = 0 To m_numOfLassoPoints - 1
                            m_LassoPoints(i).x = m_LassoPointsBackup(i).x + (x - moveXDist)
                            m_LassoPoints(i).y = m_LassoPointsBackup(i).y + (y - moveYDist)
                        Next i
                        
                End Select
                        
        End Select
                
    'This is not a transform, meaning the selection is being created for the first time.  For standard selection types,
    ' this simply means copying the passed (x, y) values.  Polygon/lasso selections are more complicated.
    Else
        
        Select Case sShape
        
            'Rectangle, ellipse, and line selections are easy - just copy the passed (x, y) values into (x2, y2)
            Case sRectangle, sCircle, sLine
                x2 = x
                y2 = y
                
            'Lasso selections will increment the current lasso array, and add the new points to their list
            Case sLasso
            
                'Make room in the point array
                m_numOfLassoPoints = m_numOfLassoPoints + 1
                If m_numOfLassoPoints > UBound(m_LassoPoints) Then ReDim Preserve m_LassoPoints(0 To m_numOfLassoPoints * 2 - 1) As POINTFLOAT
                
                'Store the new point
                m_LassoPoints(m_numOfLassoPoints - 1).x = x
                m_LassoPoints(m_numOfLassoPoints - 1).y = y
            
            Case Else
            
        End Select
        
    End If
    
    'Update the bounding rect for the selection as a whole, based on the new coordinates
    updateInternalCoords
    
End Sub

'Has this selection been locked in?
Public Function isLockedIn() As Boolean
    isLockedIn = lockedIn
End Function

'If the user is using the SHIFT key to request a square-shaped (or circle-shaped) selection, this function will be called.
Private Sub makeCoordinatesSquare()

    Select Case sShape
                
        Case sRectangle, sCircle, sLine
        
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
            
        'Other selection types do not currently support square modifiers
        Case Else
        
    End Select
    
End Sub

'Whenever internal vector or coordinate values are changed, this sub needs to be called to update the left/right/width/height
' values accordingly.  If desired, a full check can be initiated by setting that param to TRUE.  Normally this is not advised,
' as a full scan of all available coordinates may take some time for lasso selections.  Similarly, for lasso or polygon selections,
' the newX and newY parameters can be passed.  These will be compared against existing max/min values and updated accordingly,
' which is much faster than a full check.  (From this, you might infer that setting forceFullCheck to TRUE and also passing
' newX/newY values is pointless!)
Private Sub updateInternalCoords()

    'This function only needs to be run if the selection is stored in vector format.  If it is not, a bounding rect
    ' will already be correctly set.
    If (sShape <> sRaster) Then
    
        Dim i As Long
    
        'Mark the selection mask as "not ready", as it will need to be redrawn after new coordinates are set
        isMaskReady = False

        'If a square (1:1 aspect ratio) selection has been requested, calculate new coordinates now.
        
        ' (This set of if/then blocks looks complicated, but it's actually very simple - we simply have to account for every variation
        '  of quadrants, because the selection can be drawn up or down in both directions, giving eight possible variants of x1 </> x2
        '  and y1 </> y2.  By covering all those cases, square selections can be drawn in any direction.)
        If isSquare Then makeCoordinatesSquare
        
        'Finally, calculate a left, top, width and height for this selection based off the current individual coordinate values
        Select Case sShape
        
            'Rectangles, ellipses, and lines all use the same (x1, y1) - (x2, y2) coordinate system, so finding selection bounds is easy.
            Case sRectangle, sCircle, sLine
            
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
            
            'The way we find bounds for lasso selections varies according to the forceFullCheck parameter.  If FALSE, simply compare
            ' the passed points to the existing max/min values.  If TRUE, search all points to find max/min values.
            Case sLasso
            
                'Search the entire lasso array for new left/top/width/height values
                selLeft = LONG_MAX
                selTop = LONG_MAX
                
                Dim selMaxX As Long, selMaxY As Long
                
                For i = 0 To m_numOfLassoPoints - 1
                    
                    With m_LassoPoints(i)
                        If .x < selLeft Then selLeft = .x
                        If .y < selTop Then selTop = .y
                        If .x > selMaxX Then selMaxX = .x
                        If .y > selMaxY Then selMaxY = .y
                    End With
                    
                Next i
                
                selWidth = selMaxX - selLeft
                selHeight = selMaxY - selTop
                
                'If curvature is active, we must perform a manual scan as well
                If isLockedIn And (sSmoothStroke > 0) Then findNewBoundsManually True
                
            Case Else
            
        End Select
        
    End If

End Sub

'Because selections can be created beyond the parent image's borders, it is necessary to check if ALL selection coordinates lie off the
' image.  If this is the case, we don't want to finalize the current selection - we want to forget it.  Note that this function requires
' the selLeft/Top/Width/Height values to be correctly set in advance!
Public Function areAllCoordinatesInvalid() As Boolean

    'Selections must lie at least partially on the image.  If all coordinates lie outside the image, return TRUE, and the caller
    ' will know to remove the selection completely.
    Select Case sShape
    
        'Rectangles, ellipses, and lines are easy - check the bounding box, and if it lies completely outside the image,
        ' reject it.  Note that this occurs before a final bound rect has been calculated, so you have no choice but to
        ' rely on intermediate x/y coords instead of the boundLeft/boundTop etc values.
        Case sRectangle, sCircle, sLine, sLasso
            
            If (selLeft < 0) And (selLeft + selWidth < 0) Then areAllCoordinatesInvalid = True
            If (selLeft > containingPDImage.Width) And (selLeft + selWidth > containingPDImage.Width) Then areAllCoordinatesInvalid = True
            If (selTop < 0) And (selTop + selHeight < 0) Then areAllCoordinatesInvalid = True
            If (selTop > containingPDImage.Height) And (selTop + selHeight > containingPDImage.Height) Then areAllCoordinatesInvalid = True
            
        'In the future, additional selection types can be handled here.
        Case Else
            
    End Select
    
End Function

'Update this selection using the values in the main form's selection text boxes
Public Sub updateViaTextBox()

    'Ignore text box update requests until the selection is locked in
    If Not isLockedIn Then Exit Sub

    rejectRefreshRequests = True

    'Mark the selection mask as "not ready", as it will need to be redrawn after new coordinates are set
    isMaskReady = False

    'Check all text box entries for validity, then update the corresponding selection values.
    Select Case sShape
    
        'Rectangles, ellipses, and lines all support movement via text box
        Case sRectangle, sCircle, sLine
            If toolbar_Tools.tudSel(0).IsValid(False) Then selLeft = toolbar_Tools.tudSel(0)
            If toolbar_Tools.tudSel(1).IsValid(False) Then selTop = toolbar_Tools.tudSel(1)
            If toolbar_Tools.tudSel(2).IsValid(False) Then selWidth = toolbar_Tools.tudSel(2)
            If toolbar_Tools.tudSel(3).IsValid(False) Then selHeight = toolbar_Tools.tudSel(3)
            
        'I haven't decided if other selection types will support movement via text box...
        Case Else
        
    End Select
    
    'For some selection types, we need to update more than just the selLeft/Top/Width/Height values.
    
    Select Case sShape
    
        'Adjust the x1, y1, x2, y2 values to match the retrieved text box values
        Case sRectangle, sCircle
            x1 = selLeft
            y1 = selTop
            x2 = selLeft + selWidth
            y2 = selTop + selHeight
        
        'Adjust the x1, y1, x2, y2 values to match the retrieved text box values, but note that x2 and y2 do not correspond to
        ' width and height on line selections; they correspond to x2/y2 coordinates.
        Case sLine
            x1 = selLeft
            y1 = selTop
            x2 = selWidth
            y2 = selHeight
            
        Case Else
        
    End Select
    
    rejectRefreshRequests = False
    
End Sub

'"Lock-in" a selection. Typically this is prompted by a _MouseUp event
Public Sub lockIn()
    
    'Mark this selection as locked-in
    lockedIn = True
        
    'For vector selections, update the internal coordinates one final time
    If (sShape <> sRaster) Then
        
        updateInternalCoords
    
        'The final thing we need to check for is the width and height, which may be still be zero at this point.
        ' Due to the way outside filters and effects use selection bounding rects, we can't allow selections of size 0.
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
' 3) The selection mask is stored as a pdDIB object, so any image filters can be applied to it.
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
    ' As such, use caution when calling this function, as the existing mask will be completely erased.
    
    'Start by creating a blank mask (this will also erase any existing mask)
    selMask.createBlank containingPDImage.Width, containingPDImage.Height, 24, maskBackColor
        
    'For rounded corners, the smallest dimension (width or height) controls the maximum value of the curvature.
    ' Calculate that value in advance.
    Dim minDimension As Long
    If selWidth < selHeight Then minDimension = selWidth Else minDimension = selHeight
        
    Dim modifiedCornerVal As Long
    If roundedCorners > minDimension Then modifiedCornerVal = minDimension Else modifiedCornerVal = roundedCorners
        
    'Border selections require a bit of extra calculation, depending on the shape - basically, all bordered selections are created by
    ' rendering a smaller version of the selection within itself.  In the case of rectangles and ellipses, we don't want the smaller
    ' version to have negative values, so we check it in advance to make sure the border value is acceptable.
    Dim ignoreBorderValue As Boolean
    ignoreBorderValue = False
    
    Dim rectModifier As Long
    rectModifier = borderSize
    
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
                
                    GDI_Plus.GDIPlusFillDIBRect selMask, selLeft, selTop, selWidth, selHeight, maskForeColor
                    
                'Border selections are a bit different; we must draw two rectangles - an exterior and an interior one.
                Else
                
                    'Draw the exterior rectangle
                    GDI_Plus.GDIPlusFillDIBRect selMask, selLeft, selTop, selWidth, selHeight, maskForeColor
                    
                    'Next, draw the interior one
                    GDI_Plus.GDIPlusFillDIBRect selMask, selLeft + rectModifier, selTop + rectModifier, selWidth - rectModifier * 2, selHeight - rectModifier * 2, maskBackColor
                    
                End If
            
            'RECTANGLE SELECTION *WITH* ROUNDED CORNERS
            Else
                
                'Interior/exterior selections use identical rendering code.  The only difference is the render color (whether the
                ' rounded rect represents the selected or non-selected region, basically).
                If (sType = sInterior) Or (sType = sExterior) Or ignoreBorderValue Then
                    
                    'Use GDI+ for antialiasing, regular GDI for non-antialiased
                    GDIPlusDrawRoundRect selMask, selLeft, selTop, selWidth, selHeight, modifiedCornerVal, maskForeColor, (sSmooth > sNone) 'True
                    
                'Border selections are a bit different; we must draw two rounded rectangles - an exterior and an interior one.
                Else
                    
                    'Draw the exterior rectangle
                    GDIPlusDrawRoundRect selMask, selLeft, selTop, selWidth, selHeight, modifiedCornerVal, maskForeColor, (sSmooth > sNone)
                    
                    'Next, draw the interior one.  Note that specialized handling is required to prevent the corner curvature from rendering incorrectly.
                    If (selWidth - (rectModifier * 2) - 1) < (selHeight - (rectModifier * 2) - 1) Then minDimension = (selWidth - (rectModifier * 2) - 1) Else minDimension = (selHeight - (rectModifier * 2) - 1)
                    If roundedCorners > minDimension Then modifiedCornerVal = minDimension Else modifiedCornerVal = roundedCorners
                    GDIPlusDrawRoundRect selMask, selLeft + rectModifier, selTop + rectModifier, selWidth - rectModifier * 2, selHeight - rectModifier * 2, modifiedCornerVal, maskBackColor, (sSmooth > sNone)
                    
                End If
                
            End If
                    
        'CIRCLES / ELLIPSES
        Case sCircle
        
            'Interior/exterior selections use identical rendering code.
            If (sType = sInterior) Or (sType = sExterior) Or ignoreBorderValue Then
            
                GDIPlusDrawEllipseToDC selMask.getDIBDC, selLeft, selTop, selWidth - 1, selHeight - 1, maskForeColor, (sSmooth > sNone)
                
            'Border selections are a bit different; we must draw two ellipses - an exterior and an interior one.
            Else
            
                'Draw the exterior oval
                GDIPlusDrawEllipseToDC selMask.getDIBDC, selLeft, selTop, selWidth, selHeight, maskForeColor, (sSmooth > sNone)
                
                'Next, draw the interior one
                GDIPlusDrawEllipseToDC selMask.getDIBDC, selLeft + rectModifier, selTop + rectModifier, selWidth - rectModifier * 2, selHeight - rectModifier * 2, maskBackColor, (sSmooth > sNone)
                        
            End If
            
        
        'Line selections
        Case sLine
        
            'Interior/exterior line selections use identical rendering code.
            If (sType = sInterior) Or (sType = sExterior) Or ignoreBorderValue Then
            
                GDIPlusDrawLineToDC selMask.getDIBDC, x1, y1, x2, y2, maskForeColor, 255, selectionLineWidth, (sSmooth > sNone), LineCapRound
                
            'Border selections are a bit different; we must draw two lines - an interior and an exterior one
            Else
                
                'Draw the exterior line
                GDIPlusDrawLineToDC selMask.getDIBDC, x1, y1, x2, y2, maskForeColor, 255, selectionLineWidth, (sSmooth > sNone), LineCapRound
                
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
                            GDIPlusDrawLineToDC selMask.getDIBDC, x1 + xModifier, y1 + yModifier, x2 - xModifier, y2 - yModifier, maskBackColor, 255, newLineWidth, (sSmooth > sNone), LineCapRound
                        Else
                            GDIPlusDrawLineToDC selMask.getDIBDC, x1 + xModifier, y1 - yModifier, x2 - xModifier, y2 + yModifier, maskBackColor, 255, newLineWidth, (sSmooth > sNone), LineCapRound
                        End If
                    Else
                        If y1 = MinY Then
                            GDIPlusDrawLineToDC selMask.getDIBDC, x1 - xModifier, y1 + yModifier, x2 + xModifier, y2 - yModifier, maskBackColor, 255, newLineWidth, (sSmooth > sNone), LineCapRound
                        Else
                            GDIPlusDrawLineToDC selMask.getDIBDC, x1 - xModifier, y1 - yModifier, x2 + xModifier, y2 + yModifier, maskBackColor, 255, newLineWidth, (sSmooth > sNone), LineCapRound
                        End If
                    End If
                End If
                
            End If
            
        
        'Lasso selections do not support the "border" type while drawing.  It must be triggered by using the Select menu.
        Case sLasso
        
            'Smooth stroke affects whether curves or lines are used to draw the shape
            GDIPlusDrawFilledShapeToDC selMask.getDIBDC, m_numOfLassoPoints, VarPtr(m_LassoPoints(0)), maskForeColor, 255, (sSmooth > sNone), True, sSmoothStroke, fillmodewinding
            
        'Other selection types will be added in the future
        Case Else
    
    End Select
    
    'We now need to establish a bounding region for the selection.  For certain types of selections, we can do this with existing knowledge
    ' (e.g. the selLeft/Top/Width/Height values may reflect this).  For other types of selections, we need to find bounds via pixel searching.
    '
    'Why not just use selLeft/Top/Width/Height?  The bounding rect may differ from those values if feathering is in use.
    ' The viewport renderer uses the actual bounding rect to optimize its rendering of the selection effect, so it needs values that
    ' incorporate the full affected area, including any feathering or other modifications.
    
    'As a rule, exterior selections enclose the entire image boundary, so we sort by sType first
    Select Case sType
    
        'Exterior selections typically bound the entire image.  We could search for a smaller area, but at present the costs
        ' of this outweigh any potential benefits.
        Case sExterior
            boundLeft = 0
            boundTop = 0
            boundWidth = containingPDImage.Width
            boundHeight = containingPDImage.Height
        
        'Interior and bordered selections are handled more normally
        Case Else
        
            'Next, we sort bound calculations by selection shape
            Select Case sShape
            
                'Rectangle, ellipse, and lasso selections are easy; bounds have already been set by the width/height values
                Case sRectangle, sCircle
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
                    
                'Lines are a bit weirder.  Because the two points are not sorted by distance (e.g. x1 may be greater than x2), we have to check that now.
                Case sLine
                
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
                
                'Lasso selections support artificial smoothing via cardinal spline.  If active, we must manually calculate bounds,
                ' because the spline may extend beyond the drawn coordinates.
                Case sLasso
                
                    If sSmoothStroke = 0 Then
                    
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
                        
                    Else
                        If isLockedIn Then findNewBoundsManually True
                    End If
                
                'Other shapes currently rely on manual bounds-checking, using the rendered mask as the guide
                Case Else
                    findNewBoundsManually
                    
            End Select
                
    End Select
    
    'Do some basic bounds checking on the bound values to make sure they lie inside the image.  This is important because the
    ' selection mask (and any code that operates on it) assumes a match to image boundaries, despite the fact that selection
    ' points can actually lie anywhere on the canvas - even outside the image!
    fixBoundsToImageSize
    
    'Finally, if the selection is locked and feathering has been requested, apply it now
    If lockedIn And (sSmooth = sFullyFeathered) And (featheringRadius > 0) Then applyFeatheringToMask
    
    'Mark the mask as ready for use
    isMaskReady = True

End Sub

'When calculating selection boundaries, all bounds must ultimately lie on or inside image borders.  Use this function to verify that.
' (Because bounds have already been precisely calculated, this function's behavior does not need to differ by selection type.)
Private Sub fixBoundsToImageSize()
    
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
    
End Sub

'Apply feathering to the current selection mask.  If the user is on Win 7 or later, we do this via GDI+.  Otherwise we use
' our own internal blur functions (which are fast but lower quality).
Private Sub applyFeatheringToMask()
    
    'If GDI+ 1.1 exists, use it for a faster blur operation.
    If g_GDIPlusFXAvailable Then
    
        GDIPlusBlurDIB selMask, featheringRadius, boundLeft, boundTop, boundWidth, boundHeight
    
    'If GDI+ v1.0 is found, blurring is not supported, so fall back to PD's internal Gaussian blur function.
    Else
    
        'For exterior selections, the whole mask needs to be blurred.  No way around that.  But for interior and bordered selections,
        ' we can blur only the active area of the selection - saving precious time in the process.
        If (sType = sExterior) Or ((boundLeft = 0) And (boundTop = 0) And (boundWidth = containingPDImage.Width - 1) And (boundHeight = containingPDImage.Height - 1)) Then
            
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB selMask
        
            Message "Feathering selection..."
            CreateApproximateGaussianBlurDIB featheringRadius, tmpDIB, selMask, 1, True
            SetProgBarVal 0
            releaseProgressBar
            Message "Finished."
            
        Else
        
            'Create a DIB the size of the active selection, and copy the selection mask into that
            Message "Feathering selection..."
            Dim srcDIB As pdDIB
            Set srcDIB = New pdDIB
            srcDIB.createBlank boundWidth, boundHeight, 24
            BitBlt srcDIB.getDIBDC, 0, 0, boundWidth, boundHeight, selMask.getDIBDC, boundLeft, boundTop, vbSrcCopy
            
            'Create a second DIB (which will receive the blurred data)
            Dim dstDIB As pdDIB
            Set dstDIB = New pdDIB
            dstDIB.createBlank boundWidth, boundHeight, 24
            
            'Blur the selection
            CreateApproximateGaussianBlurDIB featheringRadius, srcDIB, dstDIB, 1, True
            
            'Copy the blurred area back onto the selection mask
            BitBlt selMask.getDIBDC, boundLeft, boundTop, boundWidth, boundHeight, dstDIB.getDIBDC, 0, 0, vbSrcCopy
            
            'Erase both temporary DIBs
            Set srcDIB = Nothing
            Set dstDIB = Nothing
            
            SetProgBarVal 0
            releaseProgressBar
            Message "Finished."
        
        End If
    
    End If

End Sub

'While a selection tool is active, we draw transform nodes around it.  The viewport renderer invokes this function as necessary.
Public Sub renderTransformNodes(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, ByVal imgLeft As Long, ByVal imgTop As Long)
    
    'If an update has been requested, but we are already in the middle of one, deny subsequent requests
    If rejectRefreshRequests Then Exit Sub
    
    'Before drawing the nodes, we need to convert the selection's coordinates from "image coordinates" to "viewport coordinates".
    ' (Remember that the selection's location is stored as coordinates relative to the IMAGE ITSELF - but to render the transform nodes
    '   on the screen, we need to translate the coordinates to the current viewport - this makes them display accurately when scroll
    '   and zoom parameters are in effect.)
    Dim dstLeft As Double, dstRight As Double
    Dim dstTop As Double, dstBottom As Double
    
    If (sShape = sRectangle) Or (sShape = sCircle) Then
        Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), selLeft, selTop, dstLeft, dstTop
        Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), selLeft + selWidth, selTop + selHeight, dstRight, dstBottom
    Else
        Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), x1, y1, dstLeft, dstTop
        Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), x2, y2, dstRight, dstBottom
    End If
    
    
    'If this selection is in vector format, we have the option of drawing transform circles at appropriate locations
    ' for the selection's shape.
    If (sShape <> sRaster) Then
        
        Dim circRadius As Long
        circRadius = 7
                
        Dim circAlpha As Long
        circAlpha = 190
        
        Select Case sShape
        
            Case sRectangle, sCircle
        
                'Draw corner circles first
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstLeft, dstTop, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstRight, dstTop, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstRight, dstBottom, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstLeft, dstBottom, circRadius, circAlpha
                
                'Draw mid-line circles next
                Dim midHorizontal As Long, midVertical As Long
                midHorizontal = (dstLeft + dstRight) \ 2
                midVertical = (dstTop + dstBottom) \ 2
                
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstLeft, midVertical, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstCanvas.hDC, midHorizontal, dstTop, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstRight, midVertical, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstCanvas.hDC, midHorizontal, dstBottom, circRadius, circAlpha
                
            Case sLine
            
                'Lines only need transform nodes at the end points of the line
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstLeft, dstTop, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstRight, dstBottom, circRadius, circAlpha
                
            'Other selection types will require more involved transform code
            Case Else
                
        End Select
        
        dstCanvas.requestBufferSync
                
    End If
    
End Sub

'Render the current selection mask using one of several methods.  Note that this code operates entirely independent of selection type.
' It only requires that valid selection boundaries be specified in advance, which it uses to accelerate the render process.
' (Also note that the renderColor default constant evaluates to RGB(112, 183, 255), but VB doesn't let us use RGB inside a function dec.)
Public Sub renderCustom(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, ByVal imgLeft As Long, ByVal imgTop As Long, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal renderType As SelectionRender, Optional ByVal renderColor As Long = 16758640)
    
    'If an update has been requested, but we are already in the middle of one, deny subsequent requests
    If rejectRefreshRequests Then Exit Sub

    'Is the selection mask ready to be used?  If not, prepare it now.
    If Not isMaskReady Then createSelectionMask
    
    'Is a special drawing mode active?  If it is, this function will be canceled in favor of a custom draw routine.
    If (sShape = sLasso) And (sDrawingDisplay = sOutlineOnlyDisplay) And (Not lockedIn) Then
        renderOutlineOnly dstDIB, srcImage, dstCanvas
        Exit Sub
    End If
    
    'Before rendering the selection, we need to convert the its coordinates from "image coordinates" to "viewport coordinates".
    ' (Remember that the selection's location is stored as coordinates relative to the IMAGE ITSELF - but to render it on the screen,
    '   we need to translate the coordinates to the current viewport - this makes them display accurately when scroll and zoom are
    '   in effect.)
    Dim dstLeft As Double, dstRight As Double
    Dim dstTop As Double, dstBottom As Double
    Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, boundLeft, boundTop, dstLeft, dstTop
    Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, boundLeft + boundWidth, boundTop + boundHeight, dstRight, dstBottom
    
    'Several different selection rendering methods use the same variables.  I've declared the common ones here.
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    Dim x As Long, y As Long
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
            
    If rightBound > srcImage.backBuffer.getDIBWidth - 1 Then rightBound = srcImage.backBuffer.getDIBWidth - 1
    If bottomBound > srcImage.backBuffer.getDIBHeight - 1 Then bottomBound = srcImage.backBuffer.getDIBHeight - 1
    
    'Build look-up tables for both the x and y direction.  This is faster than constantly calculating new values for
    ' every pixel inside the loop.  Note also that this code is identical among various selection render methods.
    Dim newX As Double, newY As Double
    Dim xLookup() As Long, yLookup() As Long
    
    ReDim xLookup(imgLeft To rightBound) As Long
    For x = imgLeft To rightBound
        Drawing.convertCanvasCoordsToImageCoords FormMain.mainCanvas(0), srcImage, x, imgTop, newX, newY, True
        xLookup(x) = newX * 3
    Next x
    
    ReDim yLookup(imgTop To bottomBound) As Long
    For y = imgTop To bottomBound
        Drawing.convertCanvasCoordsToImageCoords FormMain.mainCanvas(0), srcImage, imgLeft, y, newX, newY, True
        yLookup(y) = newY
    Next y
    
    Dim tmpSelHolder As Byte
    Dim tmpSelDivider As Double
    
    'Start by acquiring a pointer to the target DIB's pixel bits
    prepDIBSafeArray dstDIB, tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Next, acquire a pointer to the selection mask's pixel bits
    prepDIBSafeArray selMask, selMaskSA
    CopyMemory ByVal VarPtrArray(selMaskData()), VarPtr(selMaskSA), 4
    
    'Based on the selection type, draw a corresponding selection visualization.
    Select Case renderType
    
        'Lightbox style
        Case SELECTION_RENDER_LIGHTBOX
            
            'Use a look-up table to dim the non-active parts of the image
            Dim hData(0 To 255) As Byte
            
            For x = 0 To 255
                hData(x) = x \ 3
            Next x
                                    
            'Dim the non-active parts of the image
            For x = imgLeft To rightBound
                QuickVal = x * 3
            For y = imgTop To bottomBound
            
                tmpSelHolder = selMaskData(xLookup(x), yLookup(y))
                
                'Remember that lightbox works a little weirdly - UNSELECTED pixels are marked, while selected ones are left untouched
                If tmpSelHolder = 0 Then
                    iData(QuickVal + 2, y) = hData(iData(QuickVal + 2, y))
                    iData(QuickVal + 1, y) = hData(iData(QuickVal + 1, y))
                    iData(QuickVal, y) = hData(iData(QuickVal, y))
                ElseIf tmpSelHolder < 255 Then
                    tmpSelDivider = tmpSelHolder / 255
                    iData(QuickVal + 2, y) = BlendColors(hData(iData(QuickVal + 2, y)), iData(QuickVal + 2, y), tmpSelDivider)
                    iData(QuickVal + 1, y) = BlendColors(hData(iData(QuickVal + 1, y)), iData(QuickVal + 1, y), tmpSelDivider)
                    iData(QuickVal, y) = BlendColors(hData(iData(QuickVal, y)), iData(QuickVal, y), tmpSelDivider)
                End If
                
            Next y
            Next x
            
        'Highlight style (Paint.NET)
        Case SELECTION_RENDER_HIGHLIGHT
        
            'Extract RGB values from the passed color
            Dim newR As Long, newG As Long, newB As Long
            newR = ExtractR(renderColor)
            newG = ExtractG(renderColor)
            newB = ExtractB(renderColor)
            
            'Use look-up tables to highlight the active parts of the image
            Dim rData(0 To 255) As Byte, gData(0 To 255) As Byte, bData(0 To 255) As Byte
            
            For x = 0 To 255
                rData(x) = (newR + x) \ 2
                gData(x) = (newG + x) \ 2
                bData(x) = (newB + x) \ 2
            Next x
                        
            'Highlight the selected parts of the image
            For x = imgLeft To rightBound
                QuickVal = x * 3
            For y = imgTop To bottomBound
            
                tmpSelHolder = selMaskData(xLookup(x), yLookup(y))
                
                'If this pixels is over an active pixel in the selection mask, mark it with the selected color
                If tmpSelHolder = 255 Then
                    iData(QuickVal + 2, y) = rData(iData(QuickVal + 2, y))
                    iData(QuickVal + 1, y) = gData(iData(QuickVal + 1, y))
                    iData(QuickVal, y) = bData(iData(QuickVal, y))
                    
                'Partially selected pixels are rendered with antialiasing
                ElseIf tmpSelHolder > 0 Then
                    tmpSelDivider = tmpSelHolder / 255
                    iData(QuickVal + 2, y) = BlendColors(iData(QuickVal + 2, y), rData(iData(QuickVal + 2, y)), tmpSelDivider)
                    iData(QuickVal + 1, y) = BlendColors(iData(QuickVal + 1, y), gData(iData(QuickVal + 1, y)), tmpSelDivider)
                    iData(QuickVal, y) = BlendColors(iData(QuickVal, y), bData(iData(QuickVal, y)), tmpSelDivider)
                End If
                
            Next y
            Next x
            
    End Select
    
    'With our work complete, point all arrays away from their respective DIBs, then deallocate them
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    CopyMemory ByVal VarPtrArray(selMaskData), 0&, 4
    
End Sub

'Lasso and Polygon tools support the display of an outline (instead of normal highlight) during the draw phase.  It is assumed that the
' caller has verified whether this mode is what the user wants, PRIOR to calling this function.
Private Sub renderOutlineOnly(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas)

    'To render the outline correctly, we must first translate all lasso (or polygon) points from the image coordinate space
    ' to the viewport coordinate space.
    Dim tmpViewportSpace() As POINTFLOAT
    ReDim tmpViewportSpace(0 To m_numOfLassoPoints - 1) As POINTFLOAT
    
    Dim tmpX As Double, tmpY As Double
    
    Dim i As Long
    For i = 0 To m_numOfLassoPoints - 1
        Drawing.convertImageCoordsToCanvasCoords dstCanvas, srcImage, m_LassoPoints(i).x, m_LassoPoints(i).y, tmpX, tmpY
        tmpViewportSpace(i).x = tmpX
        tmpViewportSpace(i).y = tmpY
    Next i
    
    'Use GDI+ to stroke the temporary path
    GDI_Plus.GDIPlusStrokePathToDC dstDIB.getDIBDC, m_numOfLassoPoints, VarPtr(tmpViewportSpace(0)), False, RGB(0, 0, 0), 192, True, 2
    GDI_Plus.GDIPlusStrokePathToDC dstDIB.getDIBDC, m_numOfLassoPoints, VarPtr(tmpViewportSpace(0)), False, RGB(255, 255, 255), 192, True, 1
    
End Sub

Private Sub Class_Initialize()
    
    lockedIn = False
    
    'Initialize the selection mask for this object and mark it as "not ready" (because no mask has been drawn yet)
    Set selMask = New pdDIB
    isMaskReady = False
    
    'Note that a mask has never been created for this selection
    maskHasBeenCreated = False
    
    'Mark it as not transformable... yet
    isTransformable = False
    
    'By default, don't round off the corners of rectangular selections
    roundedCorners = 0
    
    'By default, selection lines have a width of 1
    selectionLineWidth = 1
    
    'By default, the normal display mode is used for lasso and polygon selections
    sDrawingDisplay = sNormalDisplay
    
    'No lasso points yet
    m_numOfLassoPoints = 0
    ReDim m_LassoPoints(0) As POINTFLOAT
    
End Sub

'Sometimes a selection needs to directly access the bits of a target DIB. This routine sets up the corresponding SafeArray.
Private Sub prepDIBSafeArray(ByRef srcDIB As pdDIB, ByRef tmpSA As SAFEARRAY2D)
    With tmpSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcDIB.getDIBHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcDIB.getDIBArrayWidth
        .pvData = srcDIB.getActualDIBBits
    End With
End Sub

Private Sub Class_Terminate()

    'Erase the selection mask
    If Not (selMask Is Nothing) Then
        selMask.eraseDIB
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
    
    Dim dibLoadedSuccessfully As Boolean
    dibLoadedSuccessfully = False
    
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
        
        'If this selection is of raster type (meaning it could not be described in vector format), a full copy of its selection mask
        ' will be embedded in the file.  Load that data now.
        If (sShape = sRaster) Then
        
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            
            dibLoadedSuccessfully = tmpDIB.createFromFile("", True, fileNum)
            
            'Create a mask for this selection that is the size of the current image.  (Remember: the selection mask is always the full size
            ' of the image.)  Then, copy the contents of the embedded selection mask into this selection mask, with resizing as necessary.
            If dibLoadedSuccessfully Then selMask.createFromExistingDIB tmpDIB, containingPDImage.Width, containingPDImage.Height, True
            
            maskHasBeenCreated = True
            isMaskReady = True
            
        End If
        
    Close #fileNum
    
    'MsgBox "read from file: " & vbCrLf & srcFilename & vbCrLf & lockedIn & vbCrLf & paramString
    
    'If this selection isn't a vector, we need to manually find its boundaries now
    If (sShape = sRaster) And dibLoadedSuccessfully Then findNewBoundsManually
    
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
        Put #fileNum, , SELECTION_FILE_VERSION_2014
        
        'Size of the attached image.  This is used to resize the selection mask as necessary if the selection is saved in raster format,
        ' and its original image was a different size from this one.
        Put #fileNum, , containingPDImage.Width
        Put #fileNum, , containingPDImage.Height
        
        'Sometimes we write the selection's current state to file.  Other times (Undo/Redo, specifically) we write out the selection's
        ' previous state.  Check for that and write the desired data as necessary.
        Dim writeSelectionState As Boolean, writeTransformable As Boolean
        Dim writeParamString As String
        Dim writeDIB As pdDIB
        Set writeDIB = New pdDIB
        
        writeSelectionState = lockedIn
        writeTransformable = isTransformable
        writeParamString = getSelectionParamString
        If (sShape = sRaster) Then writeDIB.createFromExistingDIB selMask
        
        Put #fileNum, , CLng(Len(writeParamString))
        Put #fileNum, , writeParamString
        
        'Write out the selection's current stated (locked in or not), as well as its transformability.
        Put #fileNum, , writeSelectionState
        Put #fileNum, , writeTransformable
        
        'If this selection is not transformable, we must also embed the full selection mask.
        If (sShape = sRaster) And maskHasBeenCreated Then writeDIB.writeToFile "", True, True, fileNum
        
    Close #fileNum
    
    'Debug.Print "wrote to file: " & vbCrLf & dstFilename & vbCrLf & paramString
    
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
    
    Dim i As Long
    
    'Convert absolute values to relative values (typically used when WRITING a selection file)
    If toRelativeFormat Then
    
        'All selection types have their left/top/width/height values normalized to [0, 1]. These values are stored in
        ' positions [8, 11] inclusive.
        tmpParser.setNewValueAtLocation 8, Trim$(Str(CDbl(tmpParser.GetLong(8)) / CDbl(containingPDImage.Width)))
        tmpParser.setNewValueAtLocation 9, Trim$(Str(CDbl(tmpParser.GetLong(9)) / CDbl(containingPDImage.Height)))
        tmpParser.setNewValueAtLocation 10, Trim$(Str(CDbl(tmpParser.GetLong(10)) / CDbl(containingPDImage.Width)))
        tmpParser.setNewValueAtLocation 11, Trim$(Str(CDbl(tmpParser.GetLong(11)) / CDbl(containingPDImage.Height)))
        
        'Other transformable parameters vary by selection type, which is stored as the first parameter
        Select Case tmpParser.GetLong(1)
        
            'Rectangles, ellipses, and lines all need their bounding rect transformed.  The bounding rect values are
            ' stored in positions [12, 15] inclusive.
            Case sRectangle, sCircle, sLine
                
                tmpParser.setNewValueAtLocation 12, Trim$(Str(CDbl(tmpParser.GetLong(12)) / CDbl(containingPDImage.Width)))
                tmpParser.setNewValueAtLocation 13, Trim$(Str(CDbl(tmpParser.GetLong(13)) / CDbl(containingPDImage.Height)))
                tmpParser.setNewValueAtLocation 14, Trim$(Str(CDbl(tmpParser.GetLong(14)) / CDbl(containingPDImage.Width)))
                tmpParser.setNewValueAtLocation 15, Trim$(Str(CDbl(tmpParser.GetLong(15)) / CDbl(containingPDImage.Height)))
                
            'Lasso selections are more complicated.  Position [14] contains the number of points in the lasso.  Every parameter
            ' past that point needs to be transformed in an x/y pattern.
            Case sLasso
            
                If tmpParser.GetLong(14) > 0 Then
                    For i = 0 To tmpParser.GetLong(14) - 1
                        tmpParser.setNewValueAtLocation 15 + i * 2, Trim$(Str(CDbl(tmpParser.GetDouble(15 + i * 2)) / CDbl(containingPDImage.Width)))
                        tmpParser.setNewValueAtLocation 16 + i * 2, Trim$(Str(CDbl(tmpParser.GetDouble(16 + i * 2)) / CDbl(containingPDImage.Height)))
                    Next i
                End If
                
            Case Else
            
        End Select
    
    'Convert relative values to absolute values (typically used when READING a selection file)
    Else
        
        'All selection types have their left/top/width/height values normalized from [0, 1]. These values are stored in
        ' positions [8, 11] inclusive.
        tmpParser.setNewValueAtLocation 8, Trim$(Str(CLng(tmpParser.GetDouble(8) * CDbl(containingPDImage.Width))))
        tmpParser.setNewValueAtLocation 9, Trim$(Str(CLng(tmpParser.GetDouble(9) * CDbl(containingPDImage.Height))))
        tmpParser.setNewValueAtLocation 10, Trim$(Str(CLng(tmpParser.GetDouble(10) * CDbl(containingPDImage.Width))))
        tmpParser.setNewValueAtLocation 11, Trim$(Str(CLng(tmpParser.GetDouble(11) * CDbl(containingPDImage.Height))))
        
        'Other transformable parameters vary by selection type, which is stored as the first parameter
        Select Case tmpParser.GetLong(1)
        
            'Rectangles, ellipses, and lines all need their bounding rect transformed.  The bounding rect values are
            ' stored in positions [12, 15] inclusive.
            Case sRectangle, sCircle, sLine
            
                tmpParser.setNewValueAtLocation 12, Trim$(Str(CLng(tmpParser.GetDouble(12) * CDbl(containingPDImage.Width))))
                tmpParser.setNewValueAtLocation 13, Trim$(Str(CLng(tmpParser.GetDouble(13) * CDbl(containingPDImage.Height))))
                tmpParser.setNewValueAtLocation 14, Trim$(Str(CLng(tmpParser.GetDouble(14) * CDbl(containingPDImage.Width))))
                tmpParser.setNewValueAtLocation 15, Trim$(Str(CLng(tmpParser.GetDouble(15) * CDbl(containingPDImage.Height))))
                
            'Lasso selections are more complicated.  Position [13] contains the number of points in the lasso.  Every parameter
            ' past that point needs to be transformed in an x/y pattern.
            Case sLasso
            
                If tmpParser.GetLong(14) > 0 Then
                    For i = 0 To tmpParser.GetLong(14) - 1
                        tmpParser.setNewValueAtLocation 15 + i * 2, Trim$(Str(CDbl(tmpParser.GetDouble(15 + i * 2)) * CDbl(containingPDImage.Width)))
                        tmpParser.setNewValueAtLocation 16 + i * 2, Trim$(Str(CDbl(tmpParser.GetDouble(16 + i * 2)) * CDbl(containingPDImage.Height)))
                    Next i
                End If
            
            Case Else
            
        End Select
            
    End If
    
    'Rebuild the parameter string
    tmpParser.reassembleModifiedParamString
    
    'Return the modified string
    translateParamStringToRelativeFormat = tmpParser.getParamString

End Function

'When working with raster selections (e.g. non-transformable ones), we still want to minimize selection processing time by processing the
' smallest possible rectangle that includes all selected pixels.  This function will scan the selection mask and populate the
' selLeft/Top/Width/Height and boundLeft/Top/Width/Height values automatically, based on the mask's contents.
'
'By default, this function will mark a selection as type "Raster", since it's assumed that a vector selection would already have the data
' necessary to determine its own bounds.  This behavior can be overridden, but do so with caution, because incorrectly calculated bounds
' will cause errors on tools that try to map the selection back to the active layer.
Public Sub findNewBoundsManually(Optional ByVal overrideRasterState As Boolean = False)

    Message "Finding new selection boundaries..."

    'This function assumes two things: that a selection is not transformable (otherwise we'd know the boundaries already), and
    ' a mask has already been created.  If either of these two conditions is not met, this function may fail.
    '
    'Because these states can be inferred, this function will automatically set certain selection parameters.  This behavior can
    ' be overriden, BUT DO NOT OVERRIDE IT without understanding the consequences - in particular, that a mismatch between
    ' vector selection data and the selection mask will cause errors in the mask-to-layer mapping code.
    If Not overrideRasterState Then
    
        isTransformable = False
        isMaskReady = True
        
        'Because the selection is being converted to pure raster data, we must also update its shape
        sShape = sRaster
        
    End If

    'Point a standard 2D byte array at the selection mask
    Dim x As Long, y As Long
    Dim QuickVal As Long
    
    Dim selMaskData() As Byte
    Dim selMaskSA As SAFEARRAY2D
    prepDIBSafeArray selMask, selMaskSA
    CopyMemory ByVal VarPtrArray(selMaskData()), VarPtr(selMaskSA), 4
    
    Dim maskWidth As Long, maskHeight As Long
    maskWidth = selMask.getDIBWidth - 1
    maskHeight = selMask.getDIBHeight - 1
    
    Dim boundFound As Boolean
    
    'Find the top bound first.
    boundFound = False
    y = 0
    Do
    
        For x = 0 To maskWidth
            QuickVal = x * 3
            If selMaskData(QuickVal, y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(QuickVal + 1, y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(QuickVal + 2, y) <> 0 Then
                boundFound = True
            End If
            
            If boundFound Then
                selTop = y
                boundTop = selTop
                Exit For
            End If
            
        Next x
        
        'Blank masks shouldn't be possible, but if they are, mark the boundary as found to prevent a program lock
        If (Not boundFound) And (y >= maskHeight) Then
            Debug.Print "No top boundary found - is mask blank?"
            boundFound = True
            selTop = 0
            boundTop = selTop
        End If
        
        y = y + 1
    
    Loop While (Not boundFound)
    
    'Next, find the bottom bound
    boundFound = False
    y = maskHeight
    
    Do
    
        For x = 0 To maskWidth
            QuickVal = x * 3
            If selMaskData(QuickVal, y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(QuickVal + 1, y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(QuickVal + 2, y) <> 0 Then
                boundFound = True
            End If
            
            If boundFound Then
                selHeight = y - selTop + 1
                boundHeight = selHeight
                Exit For
            End If
            
        Next x
        
        'Blank masks shouldn't be possible, but if they are, mark the boundary as found to prevent a program lock
        If (Not boundFound) And (y = 0) Then
            Debug.Print "No bottom boundary found - is mask blank?"
            boundFound = True
            selHeight = maskHeight - selTop + 1
            boundHeight = selHeight
        End If
        
        y = y - 1
    
    Loop While (Not boundFound)
    
    'Next, find the left bound
    boundFound = False
    x = 0
    Do
    
        For y = 0 To maskHeight
            If selMaskData(x, y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(x + 1, y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(x + 2, y) <> 0 Then
                boundFound = True
            End If
            
            If boundFound Then
                selLeft = x \ 3
                boundLeft = selLeft
                Exit For
            End If
            
        Next y
        
        'Blank masks shouldn't be possible, but if they are, mark the boundary as found to prevent a program lock
        If (Not boundFound) And (x = maskWidth * 3) Then
            Debug.Print "No left boundary found - is mask blank?"
            boundFound = True
            selLeft = 0
            boundLeft = selLeft
        End If
        
        x = x + 3
    
    Loop While (Not boundFound)
    
    'Finally, find the right bound
    boundFound = False
    x = maskWidth * 3
    
    Do
    
        For y = 0 To maskHeight
            If selMaskData(x, y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(x + 1, y) <> 0 Then
                boundFound = True
            ElseIf selMaskData(x + 2, y) <> 0 Then
                boundFound = True
            End If
            
            If boundFound Then
                selWidth = (x \ 3) - selLeft + 1
                boundWidth = selWidth
                Exit For
            End If
            
        Next y
        
        'Blank masks shouldn't be possible, but if they are, mark the boundary as found to prevent a program lock
        If (Not boundFound) And (x = 0) Then
            Debug.Print "No right boundary found - is mask blank?"
            boundFound = True
            selWidth = maskWidth - selLeft + 1
            boundWidth = selWidth
        End If
        
        x = x - 3
        
    Loop While (Not boundFound)
    
    'All selection boundaries have now been located
    
    'Release our temporary byte array and exit
    CopyMemory ByVal VarPtrArray(selMaskData), 0&, 4
    Erase selMaskData
    
End Sub

'External functions can use this function to request a thumbnail version of the selection mask.
Public Function requestThumbnail(ByRef dstThumbnailDIB As pdDIB, Optional ByVal thumbnailSize As Long = 64) As Boolean
    
    If selMask Is Nothing Then
        requestThumbnail = False
        Exit Function
    End If
    
    'Thumbnails have some interesting requirements.  We always want them to be square, with the image set in the middle
    ' of the thumbnail (with aspect ratio preserved) and any empty edges made transparent.
    
    'We also need to determine the thumbnail's actual width and height, and any x and y offset necessary to preserve the
    ' aspect ratio and center the image on the thumbnail.
    Dim tIcoWidth As Long, tIcoHeight As Long, tX As Double, tY As Double
    
    'Start by determining proper dimensions for the resized thumbnail image.
    convertAspectRatio selMask.getDIBWidth, selMask.getDIBHeight, thumbnailSize, thumbnailSize, tIcoWidth, tIcoHeight
    
    'If the form is wider than it is tall, center the thumbnail vertically
    If tIcoWidth > tIcoHeight Then
        tX = 0
        tY = (thumbnailSize - tIcoHeight) / 2
    
    '...otherwise, center it horizontally
    Else
        tY = 0
        tX = (thumbnailSize - tIcoWidth) / 2
    End If
    
    'Prepare the destination DIB
    dstThumbnailDIB.createBlank thumbnailSize, thumbnailSize, 32, 0
    
    'Note that the user's thumbnail performance setting affects the quality used here.
    If g_ThumbnailPerformance = PD_PERF_FASTEST Then
        If g_ImageFormats.FreeImageEnabled Then
            requestThumbnail = FreeImageResizeDIBFast(dstThumbnailDIB, tX, tY, tIcoWidth, tIcoHeight, selMask, FILTER_BILINEAR, True)
        Else
            requestThumbnail = GDIPlusResizeDIB(dstThumbnailDIB, tX, tY, tIcoWidth, tIcoHeight, selMask, 0, 0, selMask.getDIBWidth, selMask.getDIBHeight, InterpolationModeNearestNeighbor)
        End If
    Else
        requestThumbnail = GDIPlusResizeDIB(dstThumbnailDIB, tX, tY, tIcoWidth, tIcoHeight, selMask, 0, 0, selMask.getDIBWidth, selMask.getDIBHeight, InterpolationModeHighQualityBicubic)
    End If
    
End Function
