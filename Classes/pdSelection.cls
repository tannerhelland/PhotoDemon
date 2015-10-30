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
'Copyright 2012-2015 by Tanner Helland
'Created: 25/September/12
'Last updated: 18/October/14
'Last update: add outline mode for all tools (except magic wand); also toooons of refactoring
'
'This class handles all selections in PhotoDemon.  Rectangular, elliptical, line, polygon, and lasso selections are
' currently implemented, with magic wand set to be tackled in the coming weeks.
'
'All selections are treated as a subset of pixels within a bounding rectangle.  Each selection will calculate this bounding
' rectangle differently; simple shapes like rectangles and ellipses can calculate it using their x1, y1, x2, y2 coordinates,
' while more complex shapes (lasso, polygon) must search their point collection and manually construct a bounding rect that
' way.  Certain modifiable parameters - e.g. feathering, possibly border width - may also affect the bounding rect calculation.
'
'To that end, multiple coordinate sets are used within this class.  Individual shapes have their own coordinate collections;
' these vary from simple coordinate pairs, as with rectangle selections, to large variable-size arrays, as with a lasso
' selection.  From these, a uniform set of selLeft, selTop, selWidth, and selHeight values are calculated, which define the
' boundaries of the raw geometric shape constructed by the selection.  These four values are primarily used internally, or to
' show the user the size of their created shape (rectangle and elliptical selections allow the user to modify these values
' via text box), but generally they should not be used externally.
'
'At selection mask creation time, a final bounding rect (used by variables boundLeft, boundTop, boundWidth, and boundHeight)
' is created, factoring in any extra settings like feathering, border considerations, etc.  These bounding values are the
' ones that external functions should use, because they enclose the full image area affected by the selection.  In some rare
' cases, e.g. raster selections, these boundary coordinates may actually be calculated by a manual scan of the selection mask.
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
Private Const SELECTION_IDENTIFIER As Long = &H64734450     'PDsd
Private Const SELECTION_FILE_VERSION_2014 As Long = &H1002
Private Const SELECTION_FILE_VERSION_2014_2 As Long = &H1003

'What shape does this selection have?
Public Enum PD_SELECTION_SHAPE
    sRectangle = 1
    sCircle = 2
    sLine = 4
    sPolygon = 8
    sLasso = 16
    sWand = 32
    sRaster = 64
    sAllNonRasterShapes = 128
End Enum

#If False Then
    Const sRectangle = 1, sCircle = 2, sLine = 4, sPolygon = 8, sLasso = 16, sWand = 32, sRaster = 64, sAllNonRasterShapes = 128
#End If

Private sShape As PD_SELECTION_SHAPE

'What area does this selection encompass?  (Most selection shapes allow the user to change the selection state between interior, exterior,
' and bordered on the fly.
Public Enum PD_SELECTION_AREA
    sInterior = 0
    sExterior = 1
    sBorder = 2
End Enum

#If False Then
    Const sInterior = 0, sExterior = 1, sBorder = 2
#End If

'What kind of smoothing (feathering) will be applied to this selection?
Public Enum PD_SELECTION_SMOOTHING
    sNone = 0
    sAntialiased = 1
    sFullyFeathered = 2
End Enum

#If False Then
    Const sNone = 0, sAntialiased = 1, sFullyFeathered = 2
#End If

'Is a special drawing display mode in use?
Public Enum SelectionDrawingDisplay
    sNormalDisplay = 0
    sOutlineOnlyDisplay = 1
End Enum

#If False Then
    Const sNormalDisplay = 0, sOutlineOnlyDisplay = 1
#End If

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

'Lasso and polygon selections have a variable number of points.  As such, we have to track their contents dynamically.
Private m_numOfLassoPoints As Long
Private m_LassoPoints() As POINTFLOAT

Private m_numOfPolygonPoints As Long
Private m_PolygonPoints() As POINTFLOAT

'When move transforms are applied to lasso or polygon selections, we must maintain a backup copy of the original point array.
' Any transformation values are then applied to this backup data, and the backup is erased when the mouse is released.
Private m_LassoPointsBackup() As POINTFLOAT
Private m_PolygonPointsBackup() As POINTFLOAT

'When a polygon or lasso selection is closed, these values will be set to TRUE
Private m_PolygonClosed As Boolean, m_LassoClosed As Boolean

'Each different selection shape has a number of properties specific to that shape.  Instead of storing these in open variables,
' they are stored in a dictionary, and only created/accessed as necessary.
Private m_PropertyDict As pdDictionary

'When accessing properties, use the following enum.  The property dictionary automatically handles the interface between these
' values and the dictionary itself.
Public Enum PD_SELECTION_PROPERTY
    SP_AREA = 0
    SP_SMOOTHING = 1
    SP_BORDER_WIDTH = 2
    SP_FEATHERING_RADIUS = 3
    SP_ROUNDED_CORNER_RADIUS = 4
    SP_LINE_WIDTH = 5
    SP_SMOOTH_STROKE = 6
    SP_POLYGON_CURVATURE = 7
    SP_WAND_TOLERANCE = 8
    SP_WAND_SEARCH_MODE = 9
    SP_WAND_SAMPLE_MERGED = 10
    SP_WAND_COMPARE_METHOD = 11
    SP_ROUNDED_CORNER_MODIFIED = 12
    SP_BORDER_WIDTH_MODIFIED = 13
    SP_BORDER_WIDTH_IGNORE = 14
End Enum

#If False Then
    Const SP_AREA = 0, SP_SMOOTHING = 1, SP_BORDER_WIDTH = 2, SP_FEATHERING_RADIUS = 3, SP_ROUNDED_CORNER_RADIUS = 4, SP_LINE_WIDTH = 5
    Const SP_SMOOTH_STROKE = 6, SP_POLYGON_CURVATURE = 7, SP_WAND_TOLERANCE = 8, SP_WAND_SEARCH_MODE = 9, SP_WAND_SAMPLE_MERGED = 10
    Const SP_WAND_COMPARE_METHOD = 11, SP_ROUNDED_CORNER_MODIFIED = 12, SP_BORDER_WIDTH_MODIFIED = 13, SP_BORDER_WIDTH_IGNORE = 14
#End If

'Selection properties can be retrieved via these functions.  Note that the default function returns a STRING; this is because properties are
' frequently passed around via param string.  To retrieve a numeric value, use the _Long or _Double function variants.
Public Function getSelectionProperty(ByVal PropertyName As PD_SELECTION_PROPERTY) As String
    getSelectionProperty = m_PropertyDict.getEntry_String(PropertyName, "", True)
End Function

Public Function getSelectionProperty_Long(ByVal PropertyName As PD_SELECTION_PROPERTY, Optional ByVal defaultValueIfMissing As Long = 0) As Long
    getSelectionProperty_Long = m_PropertyDict.getEntry_Long(PropertyName, defaultValueIfMissing)
End Function

Public Function getSelectionProperty_Double(ByVal PropertyName As PD_SELECTION_PROPERTY, Optional ByVal defaultValueIfMissing As Double = 0#) As Double
    getSelectionProperty_Double = m_PropertyDict.getEntry_Double(PropertyName, defaultValueIfMissing)
End Function

Public Function getSelectionProperty_Boolean(ByVal PropertyName As PD_SELECTION_PROPERTY, Optional ByVal defaultValueIfMissing As Boolean = False) As Boolean
    getSelectionProperty_Boolean = m_PropertyDict.getEntry_Boolean(PropertyName, defaultValueIfMissing)
End Function

'Selection properties can be set via this function
Public Sub setSelectionProperty(ByVal PropertyName As PD_SELECTION_PROPERTY, ByVal propertyValue As Variant)
    
    'Check to see if this value already exists in the collection
    If m_PropertyDict.doesKeyExist(PropertyName) Then
    
        'If the values are the same, we don't need to actually add it
        If m_PropertyDict.getEntry_Variant(PropertyName) = propertyValue Then Exit Sub
        
    End If
    
    m_PropertyDict.AddEntry PropertyName, propertyValue
    
    'Some property changes require us to redraw the selection mask, because it will be different due to the new parameter.
    ' To minimize the impact of property changes, we try to refresh the mask only if absolutely necessary.
    Select Case PropertyName
    
        Case SP_AREA
            If (sShape <> sRaster) Then isMaskReady = False
        
        Case SP_SMOOTHING
            If (sShape <> sRaster) Then isMaskReady = False
        
        Case SP_BORDER_WIDTH
            If (sShape <> sRaster) And (sShape <> sWand) Then isMaskReady = False
        
        Case SP_FEATHERING_RADIUS
            If (sShape <> sRaster) Then isMaskReady = False
        
        Case SP_ROUNDED_CORNER_RADIUS
            If (sShape = sRectangle) Then isMaskReady = False
        
        Case SP_LINE_WIDTH
            If (sShape = sLine) Then isMaskReady = False
        
        Case SP_SMOOTH_STROKE
            If (sShape = sLasso) Then isMaskReady = False
        
        Case SP_POLYGON_CURVATURE
            If (sShape = sPolygon) Then isMaskReady = False
        
        Case SP_WAND_TOLERANCE
            If (sShape = sWand) Then isMaskReady = False
        
        Case SP_WAND_SEARCH_MODE
            If (sShape = sWand) Then isMaskReady = False
        
        Case SP_WAND_SAMPLE_MERGED
            If (sShape = sWand) Then isMaskReady = False
        
        Case SP_WAND_COMPARE_METHOD
            If (sShape = sWand) Then isMaskReady = False
        
    End Select
    
End Sub

'To save code elsewhere, a selection can be initialized by using a param string generated by a pdParamString object.  This condenses a huge
' list of variables into a single string.  This function is especially helpful when writing full selection data to/from file.
Public Sub initFromParamString(ByVal paramString As String)

    Dim i As Long
    
    'Before doing anything else, convert the paramString from its current "relative" coordinates to "absolute" coordinates based on
    ' the size of the parent image.
    paramString = translateParamStringToRelativeFormat(paramString, False)
    
    'While we're here, also make a backup of the current paramstring and mask state
    Dim origMaskState As Boolean
    origMaskState = isMaskReady
    
    Dim backupParamString As String
    backupParamString = translateParamStringToRelativeFormat(getSelectionParamString, False)
    
    'Start by creating a parameter parser to handle the parameter string.  This class will parse out individual parameters
    ' as specific data types upon request.
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(paramString) <> 0 Then cParams.setParamString paramString
    
    'The order of param string values is HARD-CODED per the following table:
    
    '01 - selection shape
    setSelectionShape cParams.GetLong(1)
    
    '02 - selection area (interior, exterior, border)
    setSelectionProperty SP_AREA, cParams.GetLong(2)
    
    '03 - smoothing type (none, AA, fully feathered)
    setSelectionProperty SP_SMOOTHING, cParams.GetLong(3)
    
    '04 - feathering radius
    setSelectionProperty SP_FEATHERING_RADIUS, cParams.GetLong(4)
    
    '05 - border size
    setSelectionProperty SP_BORDER_WIDTH, cParams.GetLong(5)
    
    '06 - rounded corner amount
    setSelectionProperty SP_ROUNDED_CORNER_RADIUS, cParams.GetLong(6)
    
    '07 - selection line width
    setSelectionProperty SP_LINE_WIDTH, cParams.GetLong(7)
    
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
        
        'Polygon selections have a variable number of parameters, based on the number of points in the polygon.
        ' Position (14) stores the number of points.
        Case sPolygon
            
            '12 - Polygon curvature
            setSelectionProperty SP_POLYGON_CURVATURE, cParams.GetDouble(12)
            
            '13 - number of points in the polygon array
            m_numOfPolygonPoints = cParams.GetLong(13)
            
            If m_numOfPolygonPoints > 0 Then
            
                ReDim m_PolygonPoints(0 To m_numOfPolygonPoints - 1) As POINTFLOAT
                
                'Retrieve all polygon points from the param string
                If m_numOfPolygonPoints > 0 Then
                    For i = 0 To m_numOfPolygonPoints - 1
                        m_PolygonPoints(i).x = cParams.GetDouble(14 + i * 2)
                        m_PolygonPoints(i).y = cParams.GetDouble(15 + i * 2)
                    Next i
                End If
                
            Else
                ReDim m_PolygonPoints(0 To 127) As POINTFLOAT
            End If
        
        'Lasso/freehand selections have a variable number of parameters, based on the number of points in the lasso.
        ' Position (14) stores the number of points.
        Case sLasso
                        
            '12 - Smooth stroke
            setSelectionProperty SP_SMOOTH_STROKE, cParams.GetDouble(12)
            
            '13 - number of points in the lasso array
            m_numOfLassoPoints = cParams.GetLong(13)
            
            If m_numOfLassoPoints > 0 Then
            
                ReDim m_LassoPoints(0 To m_numOfLassoPoints - 1) As POINTFLOAT
                
                'Retrieve all remaining lasso points
                If m_numOfLassoPoints > 0 Then
                    For i = 0 To m_numOfLassoPoints - 1
                        m_LassoPoints(i).x = cParams.GetDouble(14 + i * 2)
                        m_LassoPoints(i).y = cParams.GetDouble(15 + i * 2)
                    Next i
                End If
                
            Else
                ReDim m_LassoPoints(0 To 127) As POINTFLOAT
            End If
        
        Case sWand
            
            '12, 13 - (x,y) coordinates for the fill
            x1 = cParams.GetLong(12)
            y1 = cParams.GetLong(13)
            
            '14 - Tolerance
            setSelectionProperty SP_WAND_TOLERANCE, cParams.GetDouble(14)
            
            '15 - Sample Merged
            setSelectionProperty SP_WAND_SAMPLE_MERGED, cParams.GetLong(15)
            
            '16 - Search Mode (contiguous or global)
            setSelectionProperty SP_WAND_SEARCH_MODE, cParams.GetLong(16)
            
            '17 - Compare Mode (one of several values, e.g. "Luminance", "Red", etc
            setSelectionProperty SP_WAND_COMPARE_METHOD, cParams.GetLong(17)
            
        'Other types (e.g. raster selections) cannot be initiated this way, so we can ignore them.
        Case Else
    
    End Select
    
    'Setting selection parameters always forces a redraw of the selection mask, which is a performance-intensive operation.
    ' To avoid that unless absolutely necessary, compare the pre-initialization param string to the one we were passed.
    ' If they match, we don't need to regenerate the mask, because it would only be identical to the current one.
    If origMaskState And (StrComp(paramString, backupParamString, vbBinaryCompare) <> 0) Then isMaskReady = False
        
End Sub

'Return all of this selection's important settings as a specifically formatted parameter string.  This string can be passed to a selection
' object at any future point to recreate this selection exactly.
' (See also: "initFromParamString" above)
Public Function getSelectionParamString() As String
    
    Dim i As Long
    
    'All selection types store a preset list of standardized values
    getSelectionParamString = buildParams(sShape, getSelectionProperty(SP_AREA), getSelectionProperty(SP_SMOOTHING), getSelectionProperty(SP_FEATHERING_RADIUS), getSelectionProperty(SP_BORDER_WIDTH), getSelectionProperty(SP_ROUNDED_CORNER_RADIUS), getSelectionProperty(SP_LINE_WIDTH), selLeft, selTop, selWidth, selHeight)
    
    'After the preset values, each selection type can write its own coords, relevant to its particular shape
    Select Case sShape
    
        'Rectangles, ellipses, and lines require only (x1, y1) and (x2, y2) coordinates
        Case sRectangle, sCircle, sLine
            getSelectionParamString = getSelectionParamString & "|" & buildParams(x1, y1, x2, y2)
        
        'Polygon selections are trickier; they have a dynamic length based on the number of points in the shape.
        ' (TODO: use a string builder class for this, to improve performance)
        Case sPolygon
            
            getSelectionParamString = getSelectionParamString & "|" & getSelectionProperty(SP_POLYGON_CURVATURE) & "|" & CStr(m_numOfPolygonPoints)
            
            If m_numOfPolygonPoints > 0 Then
                For i = 0 To m_numOfPolygonPoints - 1
                    getSelectionParamString = getSelectionParamString & "|" & Trim$(Str(m_PolygonPoints(i).x)) & "|" & Trim$(Str(m_PolygonPoints(i).y))
                Next i
            End If
        
        'Lasso selections are trickier; they have a dynamic length based on the number of drawn points.
        ' (TODO: use a string builder class for this, to improve performance)
        Case sLasso
            
            getSelectionParamString = getSelectionParamString & "|" & getSelectionProperty(SP_SMOOTH_STROKE) & "|" & CStr(m_numOfLassoPoints)
            
            If m_numOfLassoPoints > 0 Then
                For i = 0 To m_numOfLassoPoints - 1
                    getSelectionParamString = getSelectionParamString & "|" & Trim$(Str(m_LassoPoints(i).x)) & "|" & Trim$(Str(m_LassoPoints(i).y))
                Next i
            End If
            
        'Wand selections only have a few extra params
        Case sWand
        
            getSelectionParamString = getSelectionParamString & "|" & CStr(x1) & "|" & CStr(y1) & "|" & getSelectionProperty(SP_WAND_TOLERANCE) & "|" & getSelectionProperty(SP_WAND_SAMPLE_MERGED) & "|" & getSelectionProperty(SP_WAND_SEARCH_MODE) & "|" & getSelectionProperty(SP_WAND_COMPARE_METHOD)
            
        'Raster selections cannot be initiated this way!
        Case sRaster
    
    End Select
    
    getSelectionParamString = translateParamStringToRelativeFormat(getSelectionParamString, True)
    
End Function

'Simple - use this to select the entire image attached to this selection object.  Note that this DOES NOT render the new selection on-screen,
' and it doesn't render a matching selection mask.  The calling function is responsible for that.
Public Sub selectAll()

    'Set basic information about this selection
    setSelectionShape sRectangle
    setSelectionProperty SP_AREA, sInterior
    setSelectionProperty SP_FEATHERING_RADIUS, 0
    setSelectionProperty SP_SMOOTHING, sAntialiased
    setSelectionProperty SP_ROUNDED_CORNER_RADIUS, 0
    
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
    If (sShape <> sRaster) And (sShape <> sLasso) And (sShape <> sPolygon) And (sShape <> sWand) Then
    
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

Public Sub overrideTransformMode(ByVal newOverride As Boolean)
    transformModeActive = newOverride
End Sub

'Get/set the selection's current coordinates.  At present, this is only used by line selections, as part of "compare mouse
' coordinates against selection coordinates" functionality.
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
Public Function getSelectionShape() As PD_SELECTION_SHAPE
    getSelectionShape = sShape
End Function

Public Sub setSelectionShape(ByVal selShape As PD_SELECTION_SHAPE)
    
    'Certain types of shapes are transformable.  Mark those now.
    Select Case selShape
        
        Case sRectangle, sCircle, sLine
            isTransformable = True
            
        Case sLasso, sPolygon
            isTransformable = True
            
        Case sWand
            isTransformable = False
        
    End Select
    
    If (sShape <> sRaster) Then isMaskReady = False
    sShape = selShape
    
End Sub

'Return the number of polygon points
Public Function getNumOfPolygonPoints() As Long
    getNumOfPolygonPoints = m_numOfPolygonPoints
End Function

'Copy the current polygon point collection into an outside array.  This is used by mouse coordinate checking functions.
Friend Sub getPolygonPoints(ByRef ptFloatArray() As POINTFLOAT)
    
    If m_numOfPolygonPoints > 0 Then
        ReDim ptFloatArray(0 To m_numOfPolygonPoints - 1) As POINTFLOAT
        CopyMemory ByVal VarPtr(ptFloatArray(0)), ByVal VarPtr(m_PolygonPoints(0)), m_numOfPolygonPoints * 8
    Else
        ReDim ptFloatArray(0) As POINTFLOAT
    End If
    
End Sub

'For lasso selections, the canvas needs to know if the current lasso selection is open (e.g. still under construction) or closed.
Public Function getLassoClosedState() As Boolean
    getLassoClosedState = m_LassoClosed
End Function

Public Sub setLassoClosedState(ByVal newState As Boolean)
    m_LassoClosed = newState
End Sub

'For polygon selections, the canvas needs to know if the current polygon selection is open (e.g. still under construction) or closed.
Public Function getPolygonClosedState() As Boolean
    getPolygonClosedState = m_PolygonClosed
End Function

Public Sub setPolygonClosedState(ByVal newState As Boolean)
    m_PolygonClosed = newState
End Sub

'For lasso and polygon selections, this function will return the current selecton region as a GDI+ region handle.  The mouse coordinate
' checker code uses this to see if the mouse cursor is currently within the bounds of the selection area.
Public Function getGdipRegionForSelection() As Long
    If sShape = sPolygon Then
        getGdipRegionForSelection = GDI_Plus.getGDIPlusRegionFromPoints(m_numOfPolygonPoints, VarPtr(m_PolygonPoints(0)), FillModeWinding, True, getSelectionProperty_Double(SP_POLYGON_CURVATURE))
    Else
        getGdipRegionForSelection = GDI_Plus.getGDIPlusRegionFromPoints(m_numOfLassoPoints, VarPtr(m_LassoPoints(0)), FillModeWinding, True, getSelectionProperty_Double(SP_SMOOTH_STROKE))
    End If
End Function

'For lasso and polygon selections, this function will calculate a boundary rect using GDI+.  This is relevant when curvature is active,
' because the curvature bounds may extend outside the points used to create the selection shape.
Public Sub fillBoundsUsingGDIPlus()
    
    'This function is only relevant for lasso and polygon selections
    If ((sShape <> sLasso) And (sShape <> sPolygon)) Then Exit Sub
    
    'GDI+ returns boundaries as a floating-point rect
    Dim boundRect As RECTF
    
    If sShape = sPolygon Then
        boundRect = GDI_Plus.getGDIPlusUnionFromPointsAndImage(m_numOfPolygonPoints, VarPtr(m_PolygonPoints(0)), containingPDImage, FillModeWinding, True, getSelectionProperty_Double(SP_POLYGON_CURVATURE))
    Else
        boundRect = GDI_Plus.getGDIPlusUnionFromPointsAndImage(m_numOfLassoPoints, VarPtr(m_LassoPoints(0)), containingPDImage, FillModeWinding, True, getSelectionProperty_Double(SP_SMOOTH_STROKE))
    End If
    
    'The returned boundaries encompass the intersection of our containing image, and the selection region.
    
    'All that's left to do is to copy the new rect values into our internal selLeft/selTop/etc and boundLeft/boundTop/etc values
    With boundRect
        selLeft = .Left
        selTop = .Top
        selWidth = .Width
        selHeight = .Height
    End With
    
    boundLeft = selLeft
    boundTop = selTop
    boundWidth = selWidth
    boundHeight = selHeight
    
End Sub

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
                
        'Polygons can be both moved and resized (by click-dragging individual points)
        Case sPolygon
        
            'Based on the transform mode, set the initial points accordingly
            Select Case selTransformationType
        
                'numOfPolygonPoints - interior of the selection, which will trigger a move transformation
                Case m_numOfPolygonPoints
                    moveXDist = x
                    moveYDist = y
                    
                    'For the transform to work, we need to apply any transformed data to the *original polygon points*.
                    ' Make a backup copy of the original array now.
                    ReDim m_PolygonPointsBackup(0 To m_numOfPolygonPoints - 1) As POINTFLOAT
                    CopyMemory ByVal VarPtr(m_PolygonPointsBackup(0)), ByVal VarPtr(m_PolygonPoints(0)), m_numOfPolygonPoints * 8
                    
                'Other transforms don't actually require special handling, as all the work is done in the
                ' setAdditionalTransformCoordinates function.
                Case Else
                    
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
            
        'Wand selections don't care about transformation
        Case sWand
            x1 = x
            y1 = y
        
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
    
    'If we're setting new initial coordinates, the mask is not (by definition) ready
    isMaskReady = False
    
    x1 = x
    y1 = y
    
    'Set the second set of point to match the first set
    Select Case sShape
        
        Case sRectangle, sCircle, sLine
            x2 = x
            y2 = y
            
            'Also, mark selections of any of these shapes (rectangle, circle, line) as transformable
            isTransformable = True
        
        'Polygon selections require us to initialize a whole bunch of polygon tracking variables
        Case sPolygon
            isTransformable = True
            
            m_numOfPolygonPoints = 1
            ReDim m_PolygonPoints(0 To 127) As POINTFLOAT
            
            m_PolygonPoints(0).x = x
            m_PolygonPoints(0).y = y
            
            'Also set up a dummy set of initial boundary coordinates
            selLeft = x
            selTop = y
            selWidth = 1
            selHeight = 1
            
        'Lasso selections have limited transform capabilities, and they have a variable number of points
        Case sLasso
            isTransformable = True
            
            m_numOfLassoPoints = 1
            ReDim m_LassoPoints(0 To 127) As POINTFLOAT
            
            m_LassoPoints(0).x = x
            m_LassoPoints(0).y = y
            
            'Also set up a dummy set of initial boundary coordinates
            selLeft = x
            selTop = y
            selWidth = 1
            selHeight = 1
        
        'Wand selections don't support transforms
        Case sWand
            isTransformable = False
            x1 = x
            y1 = y
        
        'Other selection types (e.g. raster selections) cannot be created this way
        Case Else
        
    End Select
    
    updateInternalCoords
    
End Sub

'Takes x and y coordinates (from a _MouseMove event, typically) and stores them internally
Public Sub setAdditionalCoordinates(ByVal x As Double, y As Double)
    
    Dim newMaskRequired As Boolean
    newMaskRequired = True
    
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
            
            'Polygon transforms consist of either moving an individual polygon point, or moving the entire polygon array
            Case sPolygon
            
                'Based on the transform mode, set the initial points accordingly
                Select Case selTransformationType
            
                    'Case -1 should never occur.  (-1 represents an invalid transformation request)
                    Case -1
                    
                    'Move transform
                    Case m_numOfPolygonPoints
                        
                        'Rebuild the main polygon array by copying all points from the backup array, and applying the current
                        ' x/y transformation distance to them.
                        For i = 0 To m_numOfPolygonPoints - 1
                            m_PolygonPoints(i).x = m_PolygonPointsBackup(i).x + (x - moveXDist)
                            m_PolygonPoints(i).y = m_PolygonPointsBackup(i).y + (y - moveYDist)
                        Next i
                    
                    'Anything else is just moving a polygon point
                    Case Else
                    
                        With m_PolygonPoints(selTransformationType)
                            .x = x
                            .y = y
                        End With
                        
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
                
                'If a transform mode is active, re-mark the selection as being transformable
                isTransformable = True
            
            'Wand selections are not transformable
            Case sWand
                Debug.Print "Transform initiated on a wand selection - FIX THIS!"
            
        End Select
                
    'This is not a transform, meaning the selection is being created for the first time.  For standard selection types,
    ' this simply means copying the passed (x, y) values.  Polygon/lasso selections are more complicated.
    Else
        
        Select Case sShape
        
            'Rectangle, ellipse, and line selections are easy - just copy the passed (x, y) values into (x2, y2)
            Case sRectangle, sCircle, sLine
                x2 = x
                y2 = y
                
            'Polygon selections will increment the current polygon array by one, adding the new point as the latest polygon coordinate
            Case sPolygon
            
                'Make room in the point array
                m_numOfPolygonPoints = m_numOfPolygonPoints + 1
                If m_numOfPolygonPoints > UBound(m_PolygonPoints) Then ReDim Preserve m_PolygonPoints(0 To m_numOfPolygonPoints * 2 - 1) As POINTFLOAT
                    
                'Store the new point
                m_PolygonPoints(m_numOfPolygonPoints - 1).x = x
                m_PolygonPoints(m_numOfPolygonPoints - 1).y = y
                
            'Lasso selections will increment the current lasso array, and add the new point to its list
            Case sLasso
            
                'Perform a quick check to make sure this point isn't a duplicate of the previous point.  With high-DPI mice,
                ' this is a distinct possibility.
                If m_LassoPoints(m_numOfLassoPoints - 1).x <> x Or m_LassoPoints(m_numOfLassoPoints - 1).y <> y Then
            
                    'Make room in the point array
                    m_numOfLassoPoints = m_numOfLassoPoints + 1
                    If m_numOfLassoPoints > UBound(m_LassoPoints) Then ReDim Preserve m_LassoPoints(0 To m_numOfLassoPoints * 2 - 1) As POINTFLOAT
                    
                    'Store the new point
                    m_LassoPoints(m_numOfLassoPoints - 1).x = x
                    m_LassoPoints(m_numOfLassoPoints - 1).y = y
                    
                End If
            
            'Wand selections can have their point of interest moved, but they don't actually support "additional" coordinates
            Case sWand
                If (x1 <> x) Or (y1 <> y) Then
                    x1 = x
                    y1 = y
                Else
                    newMaskRequired = False
                End If
            
            Case Else
            
        End Select
        
    End If
    
    'Update the bounding rect for the selection as a whole, based on the new coordinates
    updateInternalCoords newMaskRequired
    
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
' values accordingly.  Note that for some selection types - e.g. lasso - a full scan of all available coordinates must be performed,
' which can be performance-intensive if the shape is complex.  As such, try not to call this function any more than is necessary.
Private Sub updateInternalCoords(Optional ByVal forciblyResetMask As Boolean = True)

    'This function only needs to be run if the selection is stored in vector format.  If it is not, a bounding rect
    ' will already be correctly set.
    If (sShape <> sRaster) Then
    
        Dim i As Long
        Dim selMaxX As Long, selMaxY As Long
    
        'Mark the selection mask as "not ready", as it will need to be redrawn after new coordinates are set
        If forciblyResetMask Then isMaskReady = False

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
            
            'Polygon selections require us to search all polygon points in order to construct a bounding rect.
            Case sPolygon
                
                selLeft = LONG_MAX
                selTop = LONG_MAX
                selMaxX = -LONG_MAX
                selMaxY = -LONG_MAX
                
                'Search the entire polygon array for new left/top/width/height values
                For i = 0 To m_numOfPolygonPoints - 1
                    
                    With m_PolygonPoints(i)
                        If .x < selLeft Then selLeft = .x
                        If .y < selTop Then selTop = .y
                        If .x > selMaxX Then selMaxX = .x
                        If .y > selMaxY Then selMaxY = .y
                    End With
                    
                Next i
                
                selWidth = selMaxX - selLeft
                selHeight = selMaxY - selTop
                
                'If curvature is active, we'll use GDI+ to find our bounds more precisely
                If isLockedIn And (getSelectionProperty_Double(SP_POLYGON_CURVATURE) > 0) Then fillBoundsUsingGDIPlus
            
            'Like polygon selection, lasso selections require us to search all lasso points in order to construct
            ' a bounding rect.
            Case sLasso
            
                selLeft = LONG_MAX
                selTop = LONG_MAX
                selMaxX = -LONG_MAX
                selMaxY = -LONG_MAX
                
                'Search the entire lasso array for new left/top/width/height values
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
                
                'If curvature is active, we'll use GDI+ to find our bounds more precisely
                If isLockedIn And (getSelectionProperty_Double(SP_SMOOTH_STROKE) > 0) Then fillBoundsUsingGDIPlus
            
            'Wand selections require manual bounds-checking
            Case sWand
                If isLockedIn Then findNewBoundsManually True
                
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
        Case sRectangle, sCircle, sLine
            
            If (selLeft + selWidth <= 0) Then areAllCoordinatesInvalid = True
            If (selLeft > containingPDImage.Width) Then areAllCoordinatesInvalid = True
            If (selTop + selHeight <= 0) Then areAllCoordinatesInvalid = True
            If (selTop > containingPDImage.Height) Then areAllCoordinatesInvalid = True
           
        'Polygon and Lasso are a bit more complicated.  If curvature is not active, we can use existing selLeft/Top etc bounds,
        ' but if it is active, we need to perform a manual search for boundaries.
        Case sPolygon, sLasso
        
            'If curvature is active, use GDI+ to update the current boundaries.
            If ((sShape = sPolygon) And (getSelectionProperty_Double(SP_POLYGON_CURVATURE) <> 0)) Or ((sShape = sLasso) And (getSelectionProperty_Double(SP_SMOOTH_STROKE) <> 0)) Then fillBoundsUsingGDIPlus
            
            If (selLeft + selWidth <= 0) Then areAllCoordinatesInvalid = True
            If (selLeft >= containingPDImage.Width) Then areAllCoordinatesInvalid = True
            If (selTop + selHeight <= 0) Then areAllCoordinatesInvalid = True
            If (selTop >= containingPDImage.Height) Then areAllCoordinatesInvalid = True
        
        'Wand selections are valid if the (x1, y1) coordinate pair falls inside the image
        Case sWand
            If x1 < 0 Then areAllCoordinatesInvalid = True
            If y1 < 0 Then areAllCoordinatesInvalid = True
            If x1 >= containingPDImage.Width Then areAllCoordinatesInvalid = True
            If y1 >= containingPDImage.Height Then areAllCoordinatesInvalid = True
                    
        'In the future, additional selection types can be handled here.
        Case Else
            
    End Select
    
End Function

'Nudge the selection in a given direction.  This function is supplied as a convenience for SELECTION TYPES WHOSE POSITION CANNOT
' BE MODIFIED BY TEXT BOX.  If a selection's position can be modified via text box (e.g. rectangle, ellipse, etc), you need to
' use the updateViaTextBox() function instead, to ensure that text box and internal position are properly synched.
Public Sub nudgeSelection(Optional ByVal hOffset As Double = 0#, Optional ByVal vOffset As Double = 0#)

    Dim i As Long

    Select Case sShape
    
        Case sRectangle, sCircle, sLine
            Debug.Print "nudgeSelection function was used on an invalid selection type - FIX THIS!"
    
        Case sPolygon
        
            'Apply the new offsets to all points in the polygon
            For i = 0 To m_numOfPolygonPoints - 1
                With m_PolygonPoints(i)
                    .x = .x + hOffset
                    .y = .y + vOffset
                End With
            Next i
            
            'We also need to update the selection's bounding rect
            updateInternalCoords
        
        Case sLasso
            
            'Apply the new offsets to all points in the lasso array
            For i = 0 To m_numOfLassoPoints - 1
                With m_LassoPoints(i)
                    .x = .x + hOffset
                    .y = .y + vOffset
                End With
            Next i
            
            'We also need to update the selection's bounding rect
            updateInternalCoords
            
        Case sWand
            x1 = x1 + hOffset
            y1 = y1 + vOffset
            
            If x1 < 0 Then x1 = 0
            If y1 < 0 Then y1 = 0
            If x1 >= containingPDImage.Width Then x1 = containingPDImage.Width - 1
            If y1 >= containingPDImage.Height Then y1 = containingPDImage.Height - 1
            
            'We also need to update the selection's bounding rect
            updateInternalCoords
            
    End Select

End Sub

'Polygon selections are not as fiddly as lasso selections, but it can still be useful to remove the last-clicked point.
' This function (typically triggered via the BACKSPACE key) can be used to remove the last-created polygon point.
Public Sub removeLastPolygonPoint()

    'First, make sure we actually have points to remove
    If m_numOfPolygonPoints > 1 Then m_numOfPolygonPoints = m_numOfPolygonPoints - 1
    
End Sub

'Lasso selections have the unique burden of being somewhat unfavorable to user error.  This function (typically triggered via
' the BACKSPACE key) can be used to retreat the lasso position, and potentially correct any errors.  The calling function must
' supply two Double-type variables, which will receive the new cursor position IN IMAGE COORDINATES.  The calling function is
' responsible for translating these to screen coordinates and actually applying the cursor repositioning.
Public Sub retreatLassoPosition(ByRef newCursorX_ImgCoords As Double, ByRef newCursorY_ImgCoords As Double)

    'Determine a point to retreat to.  We could do this a number of different ways, but because PD generally favors quality
    ' over all else, we're going to complicate it a bit.  The goal is to retreat a distance of ten pixels, which may correspond
    ' to any number of actual lasso points.
    Dim newLassoIndex As Long
    newLassoIndex = m_numOfLassoPoints - 1
    
    Dim netDistance As Double
    netDistance = 0
    
    'Start calculating the net distance traveled by the lasso.  Once a distance of 10 pixels is exceeded, set that as our new
    ' lasso position.
    Do While (netDistance < 10) And (newLassoIndex > 0)
    
        'Calculate a distance between this coordinate and the previous one.
        netDistance = netDistance + Math_Functions.distanceTwoPoints(m_LassoPoints(newLassoIndex).x, m_LassoPoints(newLassoIndex).y, m_LassoPoints(newLassoIndex - 1).x, m_LassoPoints(newLassoIndex - 1).y)
    
        'Decrement the test index and repeat
        newLassoIndex = newLassoIndex - 1
    
    Loop
    
    'Reposition the lasso point index to match
    m_numOfLassoPoints = newLassoIndex + 1
    If m_numOfLassoPoints < 1 Then m_numOfLassoPoints = 1
    
    'Return the new cursor coordinates at this position, then exit
    newCursorX_ImgCoords = m_LassoPoints(m_numOfLassoPoints - 1).x
    newCursorY_ImgCoords = m_LassoPoints(m_numOfLassoPoints - 1).y

End Sub

'Update this selection using the values in the main form's selection text boxes
Public Sub updateViaTextBox()

    'Ignore text box update requests until the selection is locked in
    If Not isLockedIn Then Exit Sub

    rejectRefreshRequests = True
    
    Dim subpanelOffset As Long
    subpanelOffset = Selection_Handler.getSelectionSubPanelFromSelectionShape(pdImages(g_CurrentImage)) * 4
    
    'Check all text box entries for validity, then update the corresponding selection values.
    Select Case sShape
    
        'Rectangles, ellipses, and lines all support movement via text box
        Case sRectangle, sCircle, sLine
            isMaskReady = False
            If toolpanel_Selections.tudSel(subpanelOffset + 0).IsValid(False) Then selLeft = toolpanel_Selections.tudSel(subpanelOffset + 0)
            If toolpanel_Selections.tudSel(subpanelOffset + 1).IsValid(False) Then selTop = toolpanel_Selections.tudSel(subpanelOffset + 1)
            If toolpanel_Selections.tudSel(subpanelOffset + 2).IsValid(False) Then selWidth = toolpanel_Selections.tudSel(subpanelOffset + 2)
            If toolpanel_Selections.tudSel(subpanelOffset + 3).IsValid(False) Then selHeight = toolpanel_Selections.tudSel(subpanelOffset + 3)
            
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
        
        'updateInternalCoords
    
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

'For polygon and lasso selection, it reduces risk to reset a few custom trackers when a selection is erased.
Public Sub eraseCustomTrackers()
    m_numOfLassoPoints = 0
    m_numOfPolygonPoints = 0
    m_PolygonClosed = False
    m_LassoClosed = False
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

    Debug.Print "(" & Timer & ") Creating mask now..."
    
    'Note that a mask has been created for this image.  This is important for saving/loading selections, as a new mask must be generated
    ' if one isn't already present.
    ' TODO: see if we can move this to the end of the function without consequences.
    maskHasBeenCreated = True
    
    'Debug msg: trying to minimize redundant mask creation requests
    If isMaskReady Then Debug.Print "Selection mask is marked as READY, but a new mask was requested.  FIX THIS!"
    
    'If the current selection is raster-type, this function should not have been called!
    If sShape = sRaster Then
        Debug.Print "Mask redraw requested for raster-type selection - FIX THIS!"
        isMaskReady = True
        Exit Sub
    End If
    
    Dim maskBackColor As Long, maskForeColor As Long, maskBackByte As Byte
    
    'Interior and exterior selections are rendered using identical code; the only difference is the colors used
    If getSelectionProperty_Long(SP_AREA) = sExterior Then
        maskBackColor = RGB(255, 255, 255)
        maskBackByte = 255
        maskForeColor = RGB(0, 0, 0)
    Else
        maskBackColor = RGB(0, 0, 0)
        maskBackByte = 0
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
    If (selMask.getDIBWidth <> containingPDImage.Width) Or (selMask.getDIBHeight <> containingPDImage.Height) Then
        selMask.createBlank containingPDImage.Width, containingPDImage.Height, 24, maskBackColor
    Else
        selMask.resetDIB maskBackByte
    End If
        
    'Border selections are automatically disabled if the border size exceeds the smallest dimension (width or height)
    ' of the image
    Dim minDimension As Long
    If selWidth < selHeight Then minDimension = selWidth Else minDimension = selHeight
            
    'When rendering rectangle and ellipse selections in BORDERED mode, we calculate some values in advance.  Basically, these shapes
    ' handle border selections by rendering a smaller version of the selection shape within itself.  We never want the smaller copy
    ' of the shape to have negative width or height, so we check its dimensions in advance to make sure the border value is acceptable.
    ' If it isn't, we'll simply ignore it, and render the shape as a solid.
    Dim ignoreBorderValue As Boolean
    Dim rectModifier As Long
    
    ignoreBorderValue = False
    
    If (getSelectionProperty_Long(SP_AREA) = sBorder) And ((sShape = sRectangle) Or (sShape = sCircle)) Then
    
        rectModifier = getSelectionProperty_Long(SP_BORDER_WIDTH)
        If rectModifier < 1 Then rectModifier = 1
        If rectModifier > (minDimension \ 2) Then ignoreBorderValue = True
            
    End If
                
    'The actual rendering of the selection will vary based on the current selection type (obviously).
    ' TODO: move all rendering into a separate child class.  This is necessary for union/intersect/etc with multiple selections.
    Select Case sShape
    
        Case sRectangle
        
            'RECTANGLE SELECTION, NO ROUNDED CORNERS
            ' (Note: rectangular selections ignore the current antialiasing setting, as there's no point)
            If getSelectionProperty_Long(SP_ROUNDED_CORNER_RADIUS) = 0 Then
            
                'Interior/exterior selections, and bordered selections with too large of a width use identical rendering code.
                If (getSelectionProperty_Long(SP_AREA) = sInterior) Or (getSelectionProperty_Long(SP_AREA) = sExterior) Or ignoreBorderValue Then
                
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
                If (getSelectionProperty_Long(SP_AREA) = sInterior) Or (getSelectionProperty_Long(SP_AREA) = sExterior) Or ignoreBorderValue Then
                    
                    'Use GDI+ for antialiasing, regular GDI for non-antialiased
                    GDIPlusDrawRoundRect selMask, selLeft, selTop, selWidth, selHeight, getSelectionProperty_Long(SP_ROUNDED_CORNER_RADIUS), maskForeColor, (getSelectionProperty_Long(SP_SMOOTHING) > sNone)
                    
                'Border selections are a bit different; we must draw two rounded rectangles - an exterior and an interior one.
                Else
                    
                    'Draw the exterior rectangle
                    GDIPlusDrawRoundRect selMask, selLeft, selTop, selWidth, selHeight, getSelectionProperty_Long(SP_ROUNDED_CORNER_RADIUS), maskForeColor, (getSelectionProperty_Long(SP_SMOOTHING) > sNone)
                    
                    'Next, draw the interior one.
                    GDIPlusDrawRoundRect selMask, selLeft + rectModifier, selTop + rectModifier, selWidth - rectModifier * 2, selHeight - rectModifier * 2, getSelectionProperty_Long(SP_ROUNDED_CORNER_RADIUS), maskBackColor, (getSelectionProperty_Long(SP_SMOOTHING) > sNone)
                    
                End If
                
            End If
                    
        'CIRCLES / ELLIPSES
        Case sCircle
        
            'Interior/exterior selections use identical rendering code.
            If (getSelectionProperty_Long(SP_AREA) = sInterior) Or (getSelectionProperty_Long(SP_AREA) = sExterior) Or ignoreBorderValue Then
            
                GDIPlusFillEllipseToDC selMask.getDIBDC, selLeft, selTop, selWidth - 1, selHeight - 1, maskForeColor, (getSelectionProperty_Long(SP_SMOOTHING) > sNone)
                
            'Border selections are a bit different; we must draw two ellipses - an exterior and an interior one.
            Else
            
                'Draw the exterior oval
                GDIPlusFillEllipseToDC selMask.getDIBDC, selLeft, selTop, selWidth, selHeight, maskForeColor, (getSelectionProperty_Long(SP_SMOOTHING) > sNone)
                
                'Next, draw the interior one
                GDIPlusFillEllipseToDC selMask.getDIBDC, selLeft + rectModifier, selTop + rectModifier, selWidth - rectModifier * 2, selHeight - rectModifier * 2, maskBackColor, (getSelectionProperty_Long(SP_SMOOTHING) > sNone)
                        
            End If
            
        
        'Line selections
        ' TODO: let the user specify line caps
        Case sLine
        
            'Interior/exterior line selections use identical rendering code.
            If (getSelectionProperty_Long(SP_AREA) = sInterior) Or (getSelectionProperty_Long(SP_AREA) = sExterior) Or ignoreBorderValue Then
            
                GDIPlusDrawLineToDC selMask.getDIBDC, x1, y1, x2, y2, maskForeColor, 255, getSelectionProperty_Long(SP_LINE_WIDTH), (getSelectionProperty_Long(SP_SMOOTHING) > sNone), LineCapRound
                
            'Border selections are a bit different; we must draw two lines - an interior and an exterior one.
            ' TODO: GDI+ supplies a function that generates a new path, using the outline of an existing path as drawn with a specified pen.
            '        That may be a more convenient way to render bordered line selections.  (GdipWidenPath FYI)
            Else
                
                'Draw the exterior line
                GDIPlusDrawLineToDC selMask.getDIBDC, x1, y1, x2, y2, maskForeColor, 255, getSelectionProperty_Long(SP_LINE_WIDTH), (getSelectionProperty_Long(SP_SMOOTHING) > sNone), LineCapRound
                
                'Next, draw the interior one.  Note that some additional calculations are required to
                Dim xDiff As Double, yDiff As Double
                xDiff = Abs(x1 - x2)
                yDiff = Abs(y1 - y2)
                Dim xModifier As Long, yModifier As Long
                If xDiff > yDiff Then
                    If xDiff = 0 Then xDiff = 0.00001
                    xModifier = getSelectionProperty_Long(SP_BORDER_WIDTH)
                    yModifier = getSelectionProperty_Long(SP_BORDER_WIDTH) * (yDiff / xDiff)
                Else
                    If yDiff = 0 Then yDiff = 0.00001
                    xModifier = getSelectionProperty_Long(SP_BORDER_WIDTH) * (xDiff / yDiff)
                    yModifier = getSelectionProperty_Long(SP_BORDER_WIDTH)
                End If
                
                Dim newLineWidth As Double
                newLineWidth = getSelectionProperty_Long(SP_LINE_WIDTH) - getSelectionProperty_Long(SP_BORDER_WIDTH)
                
                If newLineWidth > 0 Then
                    If x1 = minX Then
                        If y1 = MinY Then
                            GDIPlusDrawLineToDC selMask.getDIBDC, x1 + xModifier, y1 + yModifier, x2 - xModifier, y2 - yModifier, maskBackColor, 255, newLineWidth, (getSelectionProperty_Long(SP_SMOOTHING) > sNone), LineCapRound
                        Else
                            GDIPlusDrawLineToDC selMask.getDIBDC, x1 + xModifier, y1 - yModifier, x2 - xModifier, y2 + yModifier, maskBackColor, 255, newLineWidth, (getSelectionProperty_Long(SP_SMOOTHING) > sNone), LineCapRound
                        End If
                    Else
                        If y1 = MinY Then
                            GDIPlusDrawLineToDC selMask.getDIBDC, x1 - xModifier, y1 + yModifier, x2 + xModifier, y2 - yModifier, maskBackColor, 255, newLineWidth, (getSelectionProperty_Long(SP_SMOOTHING) > sNone), LineCapRound
                        Else
                            GDIPlusDrawLineToDC selMask.getDIBDC, x1 - xModifier, y1 - yModifier, x2 + xModifier, y2 + yModifier, maskBackColor, 255, newLineWidth, (getSelectionProperty_Long(SP_SMOOTHING) > sNone), LineCapRound
                        End If
                    End If
                End If
                
            End If
        
        'Polygon selections are easy - simply close the current point collection, then fill (or stroke) it via GDI+.
        Case sPolygon
        
            'Interior/exterior lasso selections use identical rendering code.
            If (getSelectionProperty_Long(SP_AREA) = sInterior) Or (getSelectionProperty_Long(SP_AREA) = sExterior) Then
                GDIPlusDrawFilledShapeToDC selMask.getDIBDC, m_numOfPolygonPoints, VarPtr(m_PolygonPoints(0)), maskForeColor, 255, (getSelectionProperty_Long(SP_SMOOTHING) > sNone), True, getSelectionProperty_Double(SP_POLYGON_CURVATURE), FillModeWinding
                
            'Border selections stroke the path instead of filling it
            Else
                GDI_Plus.GDIPlusStrokePathToDC selMask.getDIBDC, m_numOfPolygonPoints, VarPtr(m_PolygonPoints(0)), True, maskForeColor, 255, (getSelectionProperty_Long(SP_SMOOTHING) > sNone), getSelectionProperty_Long(SP_BORDER_WIDTH), LineCapFlat, True, getSelectionProperty_Double(SP_POLYGON_CURVATURE)
            End If
        
        'Strangely enough, lasso selections are arguably the simplest selection type to render, as we simply close the lasso shape,
        ' then fill it as if it represents an arbitrary region.
        Case sLasso
        
            'Interior/exterior lasso selections use identical rendering code.
            If (getSelectionProperty_Long(SP_AREA) = sInterior) Or (getSelectionProperty_Long(SP_AREA) = sExterior) Then
                GDIPlusDrawFilledShapeToDC selMask.getDIBDC, m_numOfLassoPoints, VarPtr(m_LassoPoints(0)), maskForeColor, 255, (getSelectionProperty_Long(SP_SMOOTHING) > sNone), True, getSelectionProperty_Double(SP_SMOOTH_STROKE), FillModeWinding
                
            'Border selections stroke the path instead of filling it
            Else
                GDI_Plus.GDIPlusStrokePathToDC selMask.getDIBDC, m_numOfLassoPoints, VarPtr(m_LassoPoints(0)), True, maskForeColor, 255, (getSelectionProperty_Long(SP_SMOOTHING) > sNone), getSelectionProperty_Long(SP_BORDER_WIDTH), LineCapFlat, True, getSelectionProperty_Double(SP_SMOOTH_STROKE)
            End If
        
        'Wand selections use a custom class to perform a flood fill
        Case sWand
        
            'Create a flood fill class to help us process the wand
            Dim cFloodFill As pdFloodFill
            Set cFloodFill = New pdFloodFill
            
            'Set all initial parameters
            cFloodFill.setAntialiasingMode (getSelectionProperty_Long(SP_SMOOTHING) > sNone)
            cFloodFill.setInitialPoint x1, y1
            cFloodFill.setTolerance getSelectionProperty_Double(SP_WAND_TOLERANCE)
            cFloodFill.setSearchMode getSelectionProperty_Long(SP_WAND_SEARCH_MODE)
            cFloodFill.setCompareMode getSelectionProperty_Long(SP_WAND_COMPARE_METHOD)
            
            'Based on the flood fill type (layer vs image), pass a different source layer to the flood fill class
            Dim tmpLayer As pdLayer
            Set tmpLayer = New pdLayer
            
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            
            If getSelectionProperty_Long(SP_WAND_SAMPLE_MERGED) = 0 Then
                containingPDImage.getCompositedImage tmpDIB
            Else
                tmpLayer.CopyExistingLayer containingPDImage.getActiveLayer
                tmpLayer.convertToNullPaddedLayer containingPDImage.Width, containingPDImage.Height
                Set tmpDIB = tmpLayer.layerDIB
            End If
            
            'Apply the flood fill
            cFloodFill.initiateFloodFill tmpDIB, selMask
            
            'Erase our temporary objects
            Set tmpDIB = Nothing
            Set tmpLayer = Nothing
                        
        'Other selection types will be added in the future
        Case Else
    
    End Select
    
    'Mark the mask as ready for use
    isMaskReady = True
    
    'We now need to establish a bounding region for the selection.  For certain types of selections, we can do this with existing knowledge
    ' (e.g. the selLeft/Top/Width/Height values may reflect this).  For other types of selections, we need to find bounds via pixel searching.
    '
    'Why not just use selLeft/Top/Width/Height?  The bounding rect may differ from those values if feathering is in use.
    ' The viewport renderer uses the actual bounding rect to optimize its rendering of the selection effect, so it needs values that
    ' incorporate the full affected area, including any feathering or other modifications.
    
    'As a rule, exterior selections enclose the entire image boundary, so we sort by getSelectionProperty_Long(SP_AREA) first
    Select Case getSelectionProperty_Long(SP_AREA)
    
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
                    If getSelectionProperty_Long(SP_SMOOTHING) = sFullyFeathered Then
                        boundLeft = selLeft - getSelectionProperty_Long(SP_FEATHERING_RADIUS)
                        boundTop = selTop - getSelectionProperty_Long(SP_FEATHERING_RADIUS)
                        boundWidth = selWidth + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2
                        boundHeight = selHeight + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2
                    Else
                        boundLeft = selLeft
                        boundTop = selTop
                        boundWidth = selWidth
                        boundHeight = selHeight
                    End If
                    
                'Lines are a bit weirder.  Because the two points are not sorted by distance (e.g. x1 may be greater than x2), we have to check that now.
                Case sLine
                
                    If getSelectionProperty_Long(SP_SMOOTHING) = sFullyFeathered Then
                        boundLeft = minX - (getSelectionProperty_Long(SP_LINE_WIDTH) / 2) - getSelectionProperty_Long(SP_FEATHERING_RADIUS)
                        boundTop = MinY - (getSelectionProperty_Long(SP_LINE_WIDTH) / 2) - getSelectionProperty_Long(SP_FEATHERING_RADIUS)
                        boundWidth = (maxX - minX) + getSelectionProperty_Long(SP_LINE_WIDTH) + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2
                        boundHeight = (MaxY - MinY) + getSelectionProperty_Long(SP_LINE_WIDTH) + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2
                    Else
                        boundLeft = minX - (getSelectionProperty_Long(SP_LINE_WIDTH) / 2)
                        boundTop = MinY - (getSelectionProperty_Long(SP_LINE_WIDTH) / 2)
                        boundWidth = (maxX - minX) + getSelectionProperty_Long(SP_LINE_WIDTH)
                        boundHeight = (MaxY - MinY) + getSelectionProperty_Long(SP_LINE_WIDTH)
                    End If
                
                'Polygon selections support curvature via cardinal spline.  If active, we must manually calculate bounds,
                ' because the spline may extend beyond the current coordinate collection.
                Case sPolygon
                
                    'GDI+ can manually calculate bounds for us.  These bounds are not guaranteed to be the smallest possible area,
                    ' but they are guaranteed to enclose the entire path.
                    fillBoundsUsingGDIPlus
                    
                    'If feathering is active, we must manually increase the bounding area to account for the feathering radius.
                    If getSelectionProperty_Long(SP_SMOOTHING) = sFullyFeathered Then
                    
                        'Because of the way GDI+ pens work, they will be centered on the stroked path.  This means that we must
                        ' factor the pen width into our boundary calculations.
                        If getSelectionProperty_Long(SP_AREA) = sBorder Then
                            boundLeft = selLeft - getSelectionProperty_Long(SP_FEATHERING_RADIUS) - (getSelectionProperty_Long(SP_BORDER_WIDTH) \ 2)
                            boundTop = selTop - getSelectionProperty_Long(SP_FEATHERING_RADIUS) - (getSelectionProperty_Long(SP_BORDER_WIDTH) \ 2)
                            boundWidth = selWidth + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2 + getSelectionProperty_Long(SP_BORDER_WIDTH)
                            boundHeight = selHeight + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2 + getSelectionProperty_Long(SP_BORDER_WIDTH)
                        Else
                            boundLeft = selLeft - getSelectionProperty_Long(SP_FEATHERING_RADIUS)
                            boundTop = selTop - getSelectionProperty_Long(SP_FEATHERING_RADIUS)
                            boundWidth = selWidth + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2
                            boundHeight = selHeight + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2
                        End If
                        
                    Else
                        boundLeft = selLeft
                        boundTop = selTop
                        boundWidth = selWidth
                        boundHeight = selHeight
                    End If
                    
                    
                'Lasso selections support artificial smoothing via cardinal spline.  If active, we must manually calculate bounds,
                ' because the spline may extend beyond the drawn coordinates.
                Case sLasso
                
                    'GDI+ can manually calculate bounds for us.  These bounds are not guaranteed to be the smallest possible area,
                    ' but they are guaranteed to enclose the entire path.
                    fillBoundsUsingGDIPlus
                    
                    'If feathering is active, we must manually increase the bounding area to account for the feathering radius.
                    If getSelectionProperty_Long(SP_SMOOTHING) = sFullyFeathered Then
                    
                        'Because of the way GDI+ pens work, they will be centered on the stroked path.  This means that we must
                        ' factor the pen width into our boundary calculations.
                        If getSelectionProperty_Long(SP_AREA) = sBorder Then
                            boundLeft = selLeft - getSelectionProperty_Long(SP_FEATHERING_RADIUS) - (getSelectionProperty_Long(SP_BORDER_WIDTH) \ 2)
                            boundTop = selTop - getSelectionProperty_Long(SP_FEATHERING_RADIUS) - (getSelectionProperty_Long(SP_BORDER_WIDTH) \ 2)
                            boundWidth = selWidth + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2 + getSelectionProperty_Long(SP_BORDER_WIDTH)
                            boundHeight = selHeight + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2 + getSelectionProperty_Long(SP_BORDER_WIDTH)
                        Else
                            boundLeft = selLeft - getSelectionProperty_Long(SP_FEATHERING_RADIUS)
                            boundTop = selTop - getSelectionProperty_Long(SP_FEATHERING_RADIUS)
                            boundWidth = selWidth + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2
                            boundHeight = selHeight + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2
                        End If
                        
                    Else
                        boundLeft = selLeft
                        boundTop = selTop
                        boundWidth = selWidth
                        boundHeight = selHeight
                    End If
                
                'Magic wand requires manual bounds-finding
                Case sWand
                    findNewBoundsManually True
                    
                    'If feathering is active, increase bounds accordingly
                    If getSelectionProperty_Long(SP_SMOOTHING) = sFullyFeathered Then
                        boundLeft = selLeft - getSelectionProperty_Long(SP_FEATHERING_RADIUS)
                        boundTop = selTop - getSelectionProperty_Long(SP_FEATHERING_RADIUS)
                        boundWidth = selWidth + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2
                        boundHeight = selHeight + getSelectionProperty_Long(SP_FEATHERING_RADIUS) * 2
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
    
    'Finally, if the selection is locked and feathering has been requested, apply it now.
    ' (We only apply feathering when locked-in, as the performance penalty is severe.)
    If lockedIn And (getSelectionProperty_Long(SP_SMOOTHING) = sFullyFeathered) And (getSelectionProperty_Long(SP_FEATHERING_RADIUS) > 0) Then applyFeatheringToMask
    
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
' TODO: replace this with QuickBlur
Private Sub applyFeatheringToMask()
    
    'If GDI+ 1.1 exists, use it for a faster blur operation.
    If g_GDIPlusFXAvailable Then
        
        'GDIPlusBlurDIB selMask, getSelectionProperty_Long(SP_FEATHERING_RADIUS), boundLeft, boundTop, boundWidth, boundHeight
        quickBlurDIB selMask, getSelectionProperty_Long(SP_FEATHERING_RADIUS)
    
    'If GDI+ v1.0 is found, blurring is not supported, so fall back to PD's internal Gaussian blur function.
    Else
    
        'For exterior selections, the whole mask needs to be blurred.  No way around that.  But for interior and bordered selections,
        ' we can blur only the active area of the selection - saving precious time in the process.
        If (getSelectionProperty_Long(SP_AREA) = sExterior) Or ((boundLeft = 0) And (boundTop = 0) And (boundWidth = containingPDImage.Width - 1) And (boundHeight = containingPDImage.Height - 1)) Then
            
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB selMask
        
            Message "Feathering selection..."
            CreateApproximateGaussianBlurDIB getSelectionProperty_Long(SP_FEATHERING_RADIUS), tmpDIB, selMask, 1, True
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
            CreateApproximateGaussianBlurDIB getSelectionProperty_Long(SP_FEATHERING_RADIUS), srcDIB, dstDIB, 1, True
            
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
' TODO: replace the radius and alpha values with a central interface node function (that also accounts for high-DPI considerations)
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
        
        Dim i As Long
        
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
            
            Case sPolygon
            
                For i = 0 To m_numOfPolygonPoints - 1
                    Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), m_PolygonPoints(i).x, m_PolygonPoints(i).y, dstLeft, dstTop
                    GDIPlusDrawCanvasCircle dstCanvas.hDC, dstLeft, dstTop, circRadius, circAlpha
                Next i
                
            Case sWand
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstLeft, dstTop, circRadius, circAlpha
                
            'Other selection types will require more involved transform code
            Case Else
                
        End Select
        
        dstCanvas.RequestBufferSync
                
    End If
    
End Sub

'Render the current selection mask using one of several methods.  Note that this code operates entirely independent of selection type.
' It only requires that valid selection boundaries be specified in advance, which it uses to accelerate the render process.
' (Also note that the renderColor default constant evaluates to RGB(112, 183, 255), but VB doesn't let us use RGB inside a function dec.)
' TODO: allow light box translucency to be set by the user
Public Sub renderCustom(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, ByVal imgLeft As Long, ByVal imgTop As Long, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal renderType As SelectionRender, Optional ByVal renderColor As Long = 16758640)
    
    'If an update has been requested, but we are already in the middle of one, deny subsequent requests
    If rejectRefreshRequests Then Exit Sub

    'Is the selection mask ready to be used?  If not, prepare it now.
    If Not isMaskReady Then createSelectionMask
    
    'A few different criteria can cause us to abandon this function in favor of renderOutlineOnly (which, per its name,
    ' renders the active selection as just an outline):
    ' 1) The user has specified outline-only mode, and this is not a raster or magic wand selection.  (Those don't support outline mode at present.)
    ' 2) A lasso selection is active but not locked-in (meaning it's still being drawn)
    ' 3) A polygon selection is active but not locked-in (meaning it's still being drawn)
    '
    'If any of the above are true, we will reroute control to renderOutlineOnly
    If ((renderType = SELECTION_RENDER_OUTLINE) And (sShape <> sWand) And (sShape <> sRaster)) Or _
        ((sShape = sLasso) And (Not m_LassoClosed)) Or _
        ((sShape = sPolygon) And (Not m_PolygonClosed)) Then
        
        renderOutlineOnly dstDIB, srcImage, dstCanvas
        Exit Sub
            
    End If
    
    'Wand and raster selections cannot be rendered in outline mode; if this happens, forcibly switch the mode to HIGHLIGHT.
    If (renderType = SELECTION_RENDER_OUTLINE) And ((sShape = sWand) Or (sShape = sRaster)) Then renderType = SELECTION_RENDER_HIGHLIGHT
    
    
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
            
    If rightBound > srcImage.canvasBuffer.getDIBWidth - 1 Then rightBound = srcImage.canvasBuffer.getDIBWidth - 1
    If bottomBound > srcImage.canvasBuffer.getDIBHeight - 1 Then bottomBound = srcImage.canvasBuffer.getDIBHeight - 1
    
    'Build look-up tables for both the x and y direction.  This is faster than constantly calculating new values for
    ' every pixel inside the loop.  Note also that this code is identical among various selection render methods.
    Dim newX As Double, newY As Double, newXInt As Long
    Dim xLookup() As Long, yLookup() As Long
    
    ReDim xLookup(imgLeft To rightBound) As Long
    For x = imgLeft To rightBound
        Drawing.convertCanvasCoordsToImageCoords FormMain.mainCanvas(0), srcImage, x, imgTop, newX, newY, True
        newXInt = newX
        xLookup(x) = newXInt * 3&
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
' TODO: figure out how to render magic wand via outline??
' TODO: figure out how to draw bordered selections.  This is easy for rects, ellipses, and lines, but polygon and lasso selections are
'       much trickier.  GdipWidenPath doesn't work because it crosses itself on the inner loop, which makes for a messy, jumbled outline.
'       We'd probably need a custom solution that removes inner loops - ugly project!
Private Sub renderOutlineOnly(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas)
    
    'To render an outline correctly, we must first translate all relevant points from the image coordinate space
    ' to the viewport coordinate space.  The number of points required obviously varies by selection shape.
    Dim tmpViewportSpace() As POINTFLOAT
    Dim numOfRenderPoints As Long

    'Simple shapes (rectangles, ellipses), will use these dstLeft/Right/Top/Bottom values instead of a custom point array
    Dim dstLeft As Double, dstRight As Double
    Dim dstTop As Double, dstBottom As Double
    
    If (sShape = sRectangle) Or (sShape = sCircle) Then
        Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, selLeft, selTop, dstLeft, dstTop
        Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, selLeft + selWidth, selTop + selHeight, dstRight, dstBottom
    ElseIf (sShape = sLine) Then
        Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, x1, y1, dstLeft, dstTop
        Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, x2, y2, dstRight, dstBottom
    End If
    
    'New in version 6.6 is an awesome new path class, which makes the rendering steps much easier than they have historically been.
    Dim pdPath As pdGraphicsPath
    Set pdPath = New pdGraphicsPath
    
    Select Case sShape
    
        Case sRectangle
            
            'Add the selection rectangle to the path
            If getSelectionProperty_Long(SP_ROUNDED_CORNER_RADIUS) > 0 Then
                pdPath.addRoundedRectangle_Relative dstLeft, dstTop, dstRight - dstLeft, dstBottom - dstTop, getSelectionProperty_Long(SP_ROUNDED_CORNER_RADIUS) * g_Zoom.getZoomValue(containingPDImage.currentZoomValue)
            Else
                pdPath.addRectangle_Absolute dstLeft, dstTop, dstRight, dstBottom
            End If
            
            'If the selection is bordered, widen it now
            'If (getSelectionProperty_Long(SP_AREA) = sBorder) And (Not getSelectionProperty_Boolean(SP_BORDER_WIDTH_IGNORE)) Then
            '    pdPath.convertPathToBorderType_Simple getSelectionProperty_Long(SP_BORDER_WIDTH)
            'End If
                        
        Case sCircle
        
            'Add the ellipse to the path
            pdPath.addEllipse_Absolute dstLeft, dstTop, dstRight, dstBottom
                    
        Case sLine
        
            'Add the line to the path
            pdPath.addLine dstLeft, dstTop, dstRight, dstBottom
            
            'Widen the path to encompass only the border of the line
            pdPath.convertPathToBorderType_Simple getSelectionProperty_Long(SP_LINE_WIDTH) * g_Zoom.getZoomValue(containingPDImage.currentZoomValue), LineCapRound
            
        'Polygon and lasso selections use roughly identical outline code; the only difference is the reference array used and the number of points
        Case sPolygon, sLasso
        
            If sShape = sPolygon Then numOfRenderPoints = m_numOfPolygonPoints Else numOfRenderPoints = m_numOfLassoPoints
            
            If numOfRenderPoints > 0 Then
            
                'Convert the lasso or polygon array to viewport coordinate space
                ReDim tmpViewportSpace(0 To numOfRenderPoints - 1) As POINTFLOAT
                Dim tmpX As Double, tmpY As Double
                
                Dim i As Long
                
                For i = 0 To numOfRenderPoints - 1
                    
                    If (sShape = sPolygon) Then
                        Drawing.convertImageCoordsToCanvasCoords dstCanvas, srcImage, m_PolygonPoints(i).x, m_PolygonPoints(i).y, tmpX, tmpY
                    Else
                        Drawing.convertImageCoordsToCanvasCoords dstCanvas, srcImage, m_LassoPoints(i).x, m_LassoPoints(i).y, tmpX, tmpY
                    End If
                    
                    tmpViewportSpace(i).x = tmpX
                    tmpViewportSpace(i).y = tmpY
                Next i
                
                'Add the converted shape to the path object
                If sShape = sPolygon Then
                    pdPath.addPolygon numOfRenderPoints, VarPtr(tmpViewportSpace(0)), m_PolygonClosed, True, getSelectionProperty_Double(SP_POLYGON_CURVATURE)
                Else
                    pdPath.addPolygon numOfRenderPoints, VarPtr(tmpViewportSpace(0)), m_LassoClosed, True, getSelectionProperty_Double(SP_SMOOTH_STROKE)
                End If
                
            End If
    
    End Select
    
    'Draw the assembled path onto the canvas
    pdPath.strokePathToDIB_UIStyle dstDIB
    
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
    
    'Prepare the property dictionary
    Set m_PropertyDict = New pdDictionary
    
    'No lasso or polygon points yet
    m_numOfLassoPoints = 0
    ReDim m_LassoPoints(0) As POINTFLOAT
    
    m_numOfPolygonPoints = 0
    ReDim m_PolygonPoints(0) As POINTFLOAT
    
    'Polygon selections are (obviously) not yet closed
    m_PolygonClosed = False
        
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
    
    'All file interactions are handled via pdFSO
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If Not cFile.FileExist(srcFilename) Then
        Message "Requested selection file could not be located.  Selection load abandoned."
        readSelectionFromFile = False
        Exit Function
    End If
    
    'Note that our mask is not yet ready
    isMaskReady = False
    
    Dim dibLoadedSuccessfully As Boolean
    dibLoadedSuccessfully = False
    
    'Open the file
    Dim hFile As Long
    If cFile.CreateFileHandle(srcFilename, hFile, True, False, OptimizeSequentialAccess) Then
    
        'Make sure this is actually a PhotoDemon selection file
        Dim selectionIDCheck As Long
        cFile.ReadDataFromFile hFile, VarPtr(selectionIDCheck), 4&
        
        If (selectionIDCheck <> SELECTION_IDENTIFIER) Then
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  pdSelection.readSelectionFromFile was passed an invalid file: " & srcFilename
            #End If
            
            cFile.CloseFileHandle hFile
            readSelectionFromFile = False
            
            Exit Function
            
        End If
        
        'Check to make sure this selection file verison is supported (not implemented right now, because there's only one selection version!)
        Dim selectionVersionCheck As Long
        cFile.ReadDataFromFile hFile, VarPtr(selectionVersionCheck), 4&
        
        'Size of the original image associated with this selection.  These values can be used to resize the selection mask as necessary if
        ' the selection is saved in raster format, and its original image was a different size from this one.
        Dim origWidth As Long, origHeight As Long
        cFile.ReadDataFromFile hFile, VarPtr(origWidth), 4&
        cFile.ReadDataFromFile hFile, VarPtr(origHeight), 4&
        
        'Retrieve the selection parameters into a string.
        ' (TODO: make this Unicode-compatible, but note that we have to change the file version number to support this.)
        Dim paramString As String, lenParamString As Long
        cFile.ReadDataFromFile hFile, VarPtr(lenParamString), 4&
        
        Dim tmpParamBuffer() As Byte
        ReDim tmpParamBuffer(0 To lenParamString - 1) As Byte
        cFile.ReadDataFromFile hFile, VarPtr(tmpParamBuffer(0)), lenParamString
        paramString = StrConv(tmpParamBuffer, vbUnicode)
        
        'Parse the parameter string into its individual components, and use those to initialize the rest of the selection
        Me.initFromParamString paramString
        
        'Was the selection locked in?  What about transformable?
        cFile.ReadDataFromFile hFile, VarPtr(lockedIn), 2&
        cFile.ReadDataFromFile hFile, VarPtr(isTransformable), 2&
        
        'Raster selections (e.g. selections that can't be described with vectors) store a full copy of their selection mask
        ' in the selection file.  Load it conditionally.
        If (sShape = sRaster) Then
            
            'If the original containing image had the same dimensions as the current image (e.g. during Undo/Redo operations),
            ' we can create the selection mask directly from the file data.  Otherwise, we have no choice but to perform a
            ' resize so that the old mask matches the new dimensions.
            If (origWidth = containingPDImage.Width) And (origHeight = containingPDImage.Height) Then
                dibLoadedSuccessfully = selMask.createFromFile("", True, hFile)
            Else
            
                Dim tmpDIB As pdDIB
                Set tmpDIB = New pdDIB
                
                dibLoadedSuccessfully = tmpDIB.createFromFile("", True, hFile)
                
                'Create a mask for this selection that is the size of the current image.  (Remember: the selection mask is always the full size
                ' of the image.)  Then, copy the contents of the embedded selection mask into this selection mask, with resizing as necessary.
                If dibLoadedSuccessfully Then selMask.createFromExistingDIB tmpDIB, containingPDImage.Width, containingPDImage.Height, True
            
            End If
            
            maskHasBeenCreated = True
            isMaskReady = True
            
        End If
        
        'We have everything we need from the file!
        cFile.CloseFileHandle hFile
        
    End If
    
    'MsgBox "read from file: " & vbCrLf & srcFilename & vbCrLf & lockedIn & vbCrLf & paramString
    
    'If this selection isn't a vector, we need to manually find its boundaries now
    If (sShape = sRaster) And dibLoadedSuccessfully Then findNewBoundsManually
    
    'Polygon selections loaded from file are assumed to be closed
    m_PolygonClosed = True
    
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

    'We use pdFSO for all file interactions, so Unicode paths are correctly supported
    Dim cFile As pdFSO
    Set cFile = New pdFSO
        
    'Delete any existing file (overwrite)
    If cFile.FileExist(dstFilename) Then cFile.KillFile dstFilename
    
    'Open the target file
    Dim hFile As Long
    If cFile.CreateFileHandle(dstFilename, hFile, True, True, OptimizeSequentialAccess) Then
    
        'Write identifiers first
        Dim tmpLong As Long
        
        tmpLong = SELECTION_IDENTIFIER
        cFile.WriteDataToFile hFile, VarPtr(tmpLong), 4&
        
        tmpLong = SELECTION_FILE_VERSION_2014
        cFile.WriteDataToFile hFile, VarPtr(tmpLong), 4&
        
        'Write the size of the attached DIB.  This is used to resize the selection mask (as necessary) if the selection is saved
        ' in raster format, and its original image was a different size from this one.
        cFile.WriteDataToFile hFile, VarPtr(containingPDImage.Width), 4&
        cFile.WriteDataToFile hFile, VarPtr(containingPDImage.Height), 4&
        
        'Sometimes we write the selection's current state to file.  Other times (Undo/Redo, specifically) we write out the
        ' selection's *previous* state.
        Dim writeSelectionState As Boolean, writeTransformable As Boolean
        Dim writeParamString As String
        
        writeSelectionState = lockedIn
        writeTransformable = isTransformable
        writeParamString = getSelectionParamString()
        
        tmpLong = CLng(Len(writeParamString))
        cFile.WriteDataToFile hFile, VarPtr(tmpLong), 4&
        
        'For historic reasons, we write the param string as an ASCII string.  I would like to change this, but it requires a new
        ' selection file version - TODO!
        Dim tmpBytes() As Byte
        tmpBytes = StrConv(writeParamString, vbFromUnicode)
        cFile.WriteDataToFile hFile, VarPtr(tmpBytes(0)), tmpLong
        
        'Write out the selection's current state (locked or not), as well as its transformability.  At load-time, we'll use these
        ' to determine if we need to load raster data or if we can recreate the mask using just the param string.
        cFile.WriteDataToFile hFile, VarPtr(writeSelectionState), 2&
        cFile.WriteDataToFile hFile, VarPtr(writeTransformable), 2&
        
        'If this selection is *not* transformable, we must also embed the full selection mask.
        If (sShape = sRaster) And maskHasBeenCreated Then selMask.writeToFile "", True, True, hFile
        
        'All done!
        cFile.CloseFileHandle hFile
        
        writeSelectionToFile = True
            
    Else
        writeSelectionToFile = False
    End If
    
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
            
            'Lasso and polygon selections are more complicated.  Position [14] contains the number of points in both shapes.
            ' Every parameter past that point needs to be transformed in an x/y pattern.
            Case sLasso, sPolygon
            
                If tmpParser.GetLong(13) > 0 Then
                    For i = 0 To tmpParser.GetLong(13) - 1
                        tmpParser.setNewValueAtLocation 14 + i * 2, Trim$(Str(CDbl(tmpParser.GetDouble(14 + i * 2)) / CDbl(containingPDImage.Width)))
                        tmpParser.setNewValueAtLocation 15 + i * 2, Trim$(Str(CDbl(tmpParser.GetDouble(15 + i * 2)) / CDbl(containingPDImage.Height)))
                    Next i
                End If
            
            'Magic wands only have a single x/y coordinate pair
            Case sWand
                tmpParser.setNewValueAtLocation 12, Trim$(Str(CDbl(tmpParser.GetLong(12)) / CDbl(containingPDImage.Width)))
                tmpParser.setNewValueAtLocation 13, Trim$(Str(CDbl(tmpParser.GetLong(13)) / CDbl(containingPDImage.Height)))
                
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
                
            'Lasso and polygon selections are more complicated.  Position [14] contains the number of points in each shape.
            ' Every parameter past that point needs to be transformed in an x/y pattern.
            Case sLasso, sPolygon
            
                If tmpParser.GetLong(13) > 0 Then
                    For i = 0 To tmpParser.GetLong(13) - 1
                        tmpParser.setNewValueAtLocation 14 + i * 2, Trim$(Str(CDbl(tmpParser.GetDouble(14 + i * 2)) * CDbl(containingPDImage.Width)))
                        tmpParser.setNewValueAtLocation 15 + i * 2, Trim$(Str(CDbl(tmpParser.GetDouble(15 + i * 2)) * CDbl(containingPDImage.Height)))
                    Next i
                End If
                
            'Magic wands only have a single x/y coordinate pair
            Case sWand
            
                tmpParser.setNewValueAtLocation 12, Trim$(Str(CLng(tmpParser.GetDouble(12) * CDbl(containingPDImage.Width))))
                tmpParser.setNewValueAtLocation 13, Trim$(Str(CLng(tmpParser.GetDouble(13) * CDbl(containingPDImage.Height))))
            
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
'
'If boundaries are successfully found, this function will return TRUE.  Otherwise, it will return FALSE, which effectively means the
' selection mask is blank.
'
'TODO: only check the blue channel, as the other channels are irrelevant
Public Function findNewBoundsManually(Optional ByVal overrideRasterState As Boolean = False) As Boolean

    'Message "Finding new selection boundaries..."

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
    
    findNewBoundsManually = True
    
    'Make sure the mask is ready for processing
    If (Not isMaskReady) Then
        Debug.Print "Creating new selection mask at the request of findNewBoundsManually()"
        createSelectionMask
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
            findNewBoundsManually = False
        End If
        
        y = y + 1
    
    Loop While (Not boundFound)
    
    'If the selection mask is empty, abandon ship
    If (Not findNewBoundsManually) Then
        CopyMemory ByVal VarPtrArray(selMaskData), 0&, 4
        Exit Function
    End If
    
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
        If (Not boundFound) And (y <= 0) Then
            Debug.Print "No bottom boundary found - is mask blank?"
            boundFound = True
            selHeight = maskHeight - selTop + 1
            boundHeight = selHeight
            findNewBoundsManually = False
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
        If (Not boundFound) And (x >= maskWidth * 3) Then
            Debug.Print "No left boundary found - is mask blank?"
            boundFound = True
            selLeft = 0
            boundLeft = selLeft
            findNewBoundsManually = False
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
        If (Not boundFound) And (x <= 0) Then
            Debug.Print "No right boundary found - is mask blank?"
            boundFound = True
            selWidth = maskWidth - selLeft + 1
            boundWidth = selWidth
            findNewBoundsManually = False
        End If
        
        x = x - 3
        
    Loop While (Not boundFound)
    
    'All selection boundaries have now been located
    
    'Release our temporary byte array and exit
    CopyMemory ByVal VarPtrArray(selMaskData), 0&, 4
    Erase selMaskData
    
End Function

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
