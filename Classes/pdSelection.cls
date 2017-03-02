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
'Copyright 2012-2017 by Tanner Helland
'Created: 25/September/12
'Last updated: 21/February/17
'Last update: total refactor to prep for paint tool integration
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
' selection.  From these, a uniform set of m_CornersLocked.Left, m_CornersLocked.Top, m_CornersLocked.Width, and m_CornersLocked.Height values are calculated, which define the
' boundaries of the raw geometric shape constructed by the selection.  These four values are primarily used internally, or to
' show the user the size of their created shape (rectangle and elliptical selections allow the user to modify these values
' via text box), but generally they should not be used externally.
'
'At selection mask creation time, a final bounding rect (used by variables m_Bounds.Left, m_Bounds.Top, m_Bounds.Width, and m_Bounds.Height)
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
Private Const SELECTION_FILE_VERSION_2017 As Long = &H1004

'What shape does this selection have?
Public Enum PD_SelectionShape
    ss_Unknown = -1
    ss_Rectangle = 0
    ss_Circle = 1
    ss_Line = 2
    ss_Polygon = 3
    ss_Lasso = 4
    ss_Wand = 5
    ss_Raster = 6
End Enum

#If False Then
    Private Const ss_Unknown = -1, ss_Rectangle = 0, ss_Circle = 1, ss_Line = 2, ss_Polygon = 3, ss_Lasso = 4, ss_Wand = 5, ss_Raster = 6
#End If

Private m_SelectionShape As PD_SelectionShape

'What area does this selection encompass?  (Most selection shapes allow the user to change the selection state between interior, exterior,
' and bordered on the fly.
Public Enum PD_SelectionArea
    sa_Interior = 0
    sa_Exterior = 1
    sa_Border = 2
End Enum

#If False Then
    Private Const sa_Interior = 0, sa_Exterior = 1, sa_Border = 2
#End If

'What kind of smoothing (feathering) will be applied to this selection?
Public Enum PD_SelectionSmoothing
    ss_None = 0
    ss_Antialiased = 1
    ss_FullyFeathered = 2
End Enum

#If False Then
    Private Const ss_None = 0, ss_Antialiased = 1, ss_FullyFeathered = 2
#End If

'Is a special drawing display mode in use?  Depending on the selection type + display mode, we may be able to skip certain
' time-consuming steps (like generating a path outline for a wand-type selection).
Public Enum PD_SelectionDisplay
    sd_Normal = 0
    sd_Outline = 1
End Enum

#If False Then
    Private Const sd_Normal = 0, sd_Outline = 1
#End If

'Points of interest.  These are typically used to associate a current mouse position with a relevant point in the active object
' (e.g. a selection or layer).  Note that the hard-coded constants are negative, by design; complex shapes (like arbitrary
' polygons) will use >= 0 values to identify actual indices in their point array, which may also be valid POIs in some contexts.
Public Enum PD_SelectionPOI

    'An undefined POI just means "the mouse isn't over this object at all"
    poi_Undefined = -1
    
    'The mouse is somewhere in the interior of this object, but not near a corner or edge
    poi_Interior = -2
    
    'Corner POI constants.  Depending on context, this may mean the corner of a complex shape's bounding box (vs the shape itself).
    poi_CornerNW = -3
    poi_CornerNE = -4
    poi_CornerSE = -5
    poi_CornerSW = -6
    
    'Edge POI constants.  Depending on context, this may mean the edge of a complex shape's bounding box (vs the shape itself).
    poi_EdgeN = -7
    poi_EdgeE = -8
    poi_EdgeS = -9
    poi_EdgeW = -10
    
End Enum

#If False Then
    Private Const poi_Undefined = -1, poi_Interior = -2, poi_CornerNW = -3, poi_CornerNE = -4, poi_CornerSE = -5, poi_CornerSW = -6, poi_EdgeN = -7, poi_EdgeE = -8, poi_EdgeS = -9, poi_EdgeW = -10
#End If

'When accessing properties, use the following enum.  The property dictionary automatically handles the interface between these
' values and the dictionary itself.
Public Enum PD_SelectionProperty
    sp_Area = 0
    sp_Smoothing = 1
    sp_BorderWidth = 2
    sp_FeatheringRadius = 3
    sp_RoundedCornerRadius = 4
    sp_LineWidth = 5
    sp_SmoothStroke = 6
    sp_PolygonCurvature = 7
    sp_WandTolerance = 8
    sp_WandSearchMode = 9
    sp_WandSampleMerged = 10
    sp_WandCompareMethod = 11
End Enum

#If False Then
    Private Const sp_Area = 0, sp_Smoothing = 1, sp_BorderWidth = 2, sp_FeatheringRadius = 3, sp_RoundedCornerRadius = 4, sp_LineWidth = 5, sp_SmoothStroke = 6, sp_PolygonCurvature = 7, sp_WandTolerance = 8, sp_WandSearchMode = 9, sp_WandSampleMerged = 10, sp_WandCompareMethod = 11
#End If

'The coordinates of the current selection (if rectangular)
Private m_CornersUnlocked As RECTF_RB

'When a selection is "locked in", the x and y values of corner points are converted to these values
Private m_CornersLocked As RECTF

'And finally, for nonstandard or inverted selections, an additional set of coordinates is required to track the bounding area of the
' selection as a whole.  External functions need these values to determine how to handle the selection's data.
Private m_Bounds As RECTF

'Is this selection "locked in"?
Private m_IsLocked As Boolean

'When the selection is moved, it's necessary to know the difference between the current mouse point and the original left / top values
Private m_MoveXDist As Single, m_MoveYDist As Single

'Is this selection requesting text box updates? Because if it is, ignore external requests to update.
Private m_RejectRefreshRequests As Boolean

'Is a current "point of interest" selected?  (Note that different selection shapes may use this value differently; only certain values
' are hard-coded across multiple shapes.)
Private m_CurrentPOI As PD_SelectionPOI

'Is transformation mode active?
Private m_TransformModeActive As Boolean

'What image does this selection belong to?  We use this to simplify things like bounds-checking against the base image.
Private m_parentPDImage As pdImage

'This DIB contains the selection mask for this selection object.  Black pixels in the selection mask represent unselected pixels in the image.
' White pixels represent selected ones.  Non-white and non-black pixels describe aliasing (or "partial" selections).
Private m_SelMask As pdDIB

'Until a mask has been created, this value will remain "false".  It is used to optimize operations on empty selections.
Private m_MaskHasBeenCreated As Boolean

'After a mask has been generated, its status will be set as "ready".  Any function that changes the selected area will reset this to "false".
' The selection rendering code can check this value to see if the mask needs to be recreated before rendering it to screen.  If external
' functions modify the selection in some way (e.g. Select menu dialogs), they need to set this to TRUE to prevent the engine from attempting
' to recreate a mask from existing vector data.
Private m_IsMaskReady As Boolean

'Some types of transformations (basic shapes, like rectangles, etc), are transformable, meaning that after they are created, the user can
' click on them again to resize or move them.  Complex transformations (magic wand, lasso, etc), may not be transformable.  This boolean
' is read by the mouse tracker in the image form, and it uses it to determine if the user is allowed to transform the current selection.
Private m_IsTransformable As Boolean

'If the user holds "Shift" while moving the selection, it will be forced to a 1:1 aspect ratio.
Private m_IsSquare As Boolean

'Lasso and polygon selections have a variable number of points.  As such, we have to track their contents dynamically.
Private m_NumOfLassoPoints As Long
Private m_LassoPoints() As POINTFLOAT

Private m_NumOfPolygonPoints As Long
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

'Retrieve various boundary rects.  Full descriptions of these rects are given above.
Friend Function GetCornersUnlockedRect() As RECTF_RB
    GetCornersUnlockedRect = m_CornersUnlocked
End Function

Friend Function GetCornersLockedRect() As RECTF
    GetCornersLockedRect = m_CornersLocked
End Function

Friend Function GetBoundaryRect() As RECTF
    GetBoundaryRect = m_Bounds
End Function

Friend Function GetMaskDC() As Long
    If (Not m_IsMaskReady) Then CreateSelectionMask
    If (Not m_SelMask Is Nothing) Then GetMaskDC = m_SelMask.GetDIBDC
End Function

Friend Function GetMaskDIB() As pdDIB
    If (Not m_IsMaskReady) Then CreateSelectionMask
    Set GetMaskDIB = m_SelMask
End Function

Friend Sub SetParentReference(ByRef srcImage As pdImage)
    Set m_parentPDImage = srcImage
End Sub

Friend Sub SuspendAutoRefresh(ByVal newSetting As Boolean)
    m_RejectRefreshRequests = newSetting
End Sub

Friend Function GetAutoRefreshSuspend() As Boolean
    GetAutoRefreshSuspend = m_RejectRefreshRequests
End Function

'Selection properties can be retrieved via these functions.  Note that the default function returns a STRING; this is because
' properties are saved/loaded as XML.  To retrieve a numeric value, use the _Long or _Float function variants.
Friend Function GetSelectionProperty(ByVal propertyName As PD_SelectionProperty) As String
    GetSelectionProperty = m_PropertyDict.GetEntry_String(propertyName, vbNullString, True)
End Function

Friend Function GetSelectionProperty_Long(ByVal propertyName As PD_SelectionProperty, Optional ByVal defaultValueIfMissing As Long = 0) As Long
    GetSelectionProperty_Long = m_PropertyDict.GetEntry_Long(propertyName, defaultValueIfMissing)
End Function

Friend Function GetSelectionProperty_Float(ByVal propertyName As PD_SelectionProperty, Optional ByVal defaultValueIfMissing As Single = 0#) As Single
    GetSelectionProperty_Float = m_PropertyDict.GetEntry_Double(propertyName, defaultValueIfMissing)
End Function

Friend Function GetSelectionProperty_Boolean(ByVal propertyName As PD_SelectionProperty, Optional ByVal defaultValueIfMissing As Boolean = False) As Boolean
    GetSelectionProperty_Boolean = m_PropertyDict.GetEntry_Boolean(propertyName, defaultValueIfMissing)
End Function

'Selection properties can be set via this function
Friend Sub SetSelectionProperty(ByVal propertyName As PD_SelectionProperty, ByVal propertyValue As Variant)
    
    'If this key+value pair already exists in the collection, don't waste time adding it
    If m_PropertyDict.DoesKeyExist(propertyName) Then
        If (m_PropertyDict.GetEntry_Variant(propertyName) = propertyValue) Then Exit Sub
    End If
    
    m_PropertyDict.AddEntry propertyName, propertyValue
    
    'Some property changes require us to redraw the selection mask.  To maximize property change performance, we only refresh
    ' the mask if absolutely necessary.
    Select Case propertyName
    
        Case sp_Area
            If (m_SelectionShape <> ss_Raster) Then m_IsMaskReady = False
        
        Case sp_Smoothing
            If (m_SelectionShape <> ss_Raster) Then m_IsMaskReady = False
        
        Case sp_BorderWidth
            If (m_SelectionShape <> ss_Raster) And (m_SelectionShape <> ss_Wand) Then m_IsMaskReady = False
        
        Case sp_FeatheringRadius
            If (m_SelectionShape <> ss_Raster) Then m_IsMaskReady = False
        
        Case sp_RoundedCornerRadius
            If (m_SelectionShape = ss_Rectangle) Then m_IsMaskReady = False
        
        Case sp_LineWidth
            If (m_SelectionShape = ss_Line) Then m_IsMaskReady = False
        
        Case sp_SmoothStroke
            If (m_SelectionShape = ss_Lasso) Then m_IsMaskReady = False
        
        Case sp_PolygonCurvature
            If (m_SelectionShape = ss_Polygon) Then m_IsMaskReady = False
        
        Case sp_WandTolerance
            If (m_SelectionShape = ss_Wand) Then m_IsMaskReady = False
        
        Case sp_WandSearchMode
            If (m_SelectionShape = ss_Wand) Then m_IsMaskReady = False
        
        Case sp_WandSampleMerged
            If (m_SelectionShape = ss_Wand) Then m_IsMaskReady = False
        
        Case sp_WandCompareMethod
            If (m_SelectionShape = ss_Wand) Then m_IsMaskReady = False
        
    End Select
    
End Sub

'To save code elsewhere, a selection can be initialized by an XML string generated by a pdParamXML object.
' (This function is primarily used when writing full selection data to/from file.)
Friend Sub InitFromXML(ByVal srcXML As String)

    Dim i As Long
    
    'Start by creating a parameter parser to handle the parameter string.  This class will parse out individual parameters
    ' as specific data types upon request.
    If (Len(srcXML) <> 0) Then
    
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        cParams.SetParamString srcXML
        
        'All selections share a core set of universal properties
        SetSelectionShape cParams.GetLong("SelectionShape", ss_Rectangle)
        SetSelectionProperty sp_Area, cParams.GetLong("SelectionArea", sa_Interior)
        SetSelectionProperty sp_Smoothing, cParams.GetLong("SelectionSmoothing", ss_None)
        SetSelectionProperty sp_FeatheringRadius, cParams.GetLong("SelectionFeatheringRadius", 0)
        SetSelectionProperty sp_BorderWidth, cParams.GetLong("SelectionBorderWidth", 1)
        SetSelectionProperty sp_RoundedCornerRadius, cParams.GetLong("SelectionRoundedCornerRadius", 0)
        SetSelectionProperty sp_LineWidth, cParams.GetLong("SelectionLineWidth", 1)
        
        'All coordinates must be transformed from absolute values to relative ones (on the scale [0, 1])
        With m_CornersLocked
            .Left = TranslateValueRelToAbs(cParams.GetDouble("SelectionLeft", 0), True)
            .Top = TranslateValueRelToAbs(cParams.GetDouble("SelectionTop", 0), False)
            .Width = TranslateValueRelToAbs(cParams.GetDouble("SelectionWidth", 0), True)
            .Height = TranslateValueRelToAbs(cParams.GetDouble("SelectionHeight", 0), False)
        End With
        
        'Additional parameters vary by selection type
        Select Case m_SelectionShape
        
            'Rectangles, ellipses, and lines only require bounding rect values (stored as two (x, y) pairs).  Retrieve these
            ' values from positions [12, 15] inclusive.
            Case ss_Rectangle, ss_Circle, ss_Line
            
                With m_CornersUnlocked
                    .Left = TranslateValueRelToAbs(cParams.GetDouble("SelectionVectorLeft", 0#), True)
                    .Top = TranslateValueRelToAbs(cParams.GetDouble("SelectionVectorTop", 0#), False)
                    .Right = TranslateValueRelToAbs(cParams.GetDouble("SelectionVectorRight", 0#), True)
                    .Bottom = TranslateValueRelToAbs(cParams.GetDouble("SelectionVectorBottom", 0#), False)
                End With
            
            'Polygon selections have a variable number of parameters, based on the number of points in the polygon.
            Case ss_Polygon
                
                SetSelectionProperty sp_PolygonCurvature, cParams.GetDouble("SelectionPolygonCurvature", 0#)
                m_NumOfPolygonPoints = cParams.GetLong("SelectionNumOfPoints", 0)
                
                If (m_NumOfPolygonPoints > 0) Then
                
                    ReDim m_PolygonPoints(0 To m_NumOfPolygonPoints - 1) As POINTFLOAT
                    
                    'Retrieve all polygon points from the param string
                    If (m_NumOfPolygonPoints > 0) Then
                        For i = 0 To m_NumOfPolygonPoints - 1
                            With m_PolygonPoints(i)
                                .x = TranslateValueRelToAbs(cParams.GetDouble("SelectionPointX" & CStr(i + 1), 0), True)
                                .y = TranslateValueRelToAbs(cParams.GetDouble("SelectionPointY" & CStr(i + 1), 0), False)
                            End With
                        Next i
                    End If
                    
                Else
                    ReDim m_PolygonPoints(0 To 127) As POINTFLOAT
                End If
            
            'Lasso/freehand selections have a variable number of parameters, based on the number of points in the lasso.
            Case ss_Lasso
                
                SetSelectionProperty sp_SmoothStroke, cParams.GetDouble("SelectionSmoothStroke", 0#)
                m_NumOfLassoPoints = cParams.GetLong("SelectionNumOfPoints", 0)
                
                If (m_NumOfLassoPoints > 0) Then
                
                    ReDim m_LassoPoints(0 To m_NumOfLassoPoints - 1) As POINTFLOAT
                    
                    'Retrieve all remaining lasso points
                    If (m_NumOfLassoPoints > 0) Then
                        For i = 0 To m_NumOfLassoPoints - 1
                            With m_LassoPoints(i)
                                .x = TranslateValueRelToAbs(cParams.GetDouble("SelectionPointX" & CStr(i + 1), 0#), True)
                                .y = TranslateValueRelToAbs(cParams.GetDouble("SelectionPointY" & CStr(i + 1), 0#), False)
                            End With
                        Next i
                    End If
                    
                Else
                    ReDim m_LassoPoints(0 To 127) As POINTFLOAT
                End If
            
            'Almost all of the "magic wand" selection properties are unique (relative to other selection shapes)
            Case ss_Wand
                
                'For weird historical reasons, wand selections store the wand origin point inside the unlocked selection rect
                With m_CornersUnlocked
                    .Left = TranslateValueRelToAbs(cParams.GetDouble("SelectionWandX", 0), True)
                    .Top = TranslateValueRelToAbs(cParams.GetDouble("SelectionWandY", 0), False)
                End With
                
                SetSelectionProperty sp_WandTolerance, cParams.GetDouble("SelectionWandTolerance", 0#)
                SetSelectionProperty sp_WandSampleMerged, cParams.GetLong("SelectionWandSampleMerged", 0)
                SetSelectionProperty sp_WandSearchMode, cParams.GetLong("SelectionWandSearchMode", 0)
                SetSelectionProperty sp_WandCompareMethod, cParams.GetLong("SelectionWandCompareMode", 0)
                
            'Other types (e.g. raster selections) cannot be initiated via XML parameters
            Case Else
        
        End Select
        
        'Initializing a selection via XML always forces a redraw of the selection mask
        m_IsMaskReady = False
        
    'End non-zero XML length check
    End If
        
End Sub

'Return all of this selection's important settings as an XML packet.  If the current selection shape is vector-based,
' this packet can be used to exactly re-create this selection.  (See also: "InitFromXML()" above)
Friend Function GetSelectionAsXML() As String
    
    Dim i As Long
    
    'All selection types store a preset list of standardized values
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    cParams.AddParam "SelectionShape", Me.GetSelectionShape
    cParams.AddParam "SelectionArea", Me.GetSelectionProperty(sp_Area)
    cParams.AddParam "SelectionSmoothing", Me.GetSelectionProperty(sp_Smoothing)
    cParams.AddParam "SelectionFeatheringRadius", Me.GetSelectionProperty(sp_FeatheringRadius)
    cParams.AddParam "SelectionBorderWidth", Me.GetSelectionProperty(sp_BorderWidth)
    cParams.AddParam "SelectionRoundedCornerRadius", Me.GetSelectionProperty(sp_RoundedCornerRadius)
    cParams.AddParam "SelectionLineWidth", Me.GetSelectionProperty(sp_LineWidth)
    
    With m_CornersLocked
        cParams.AddParam "SelectionLeft", TranslateValueAbsToRel(.Left, True)
        cParams.AddParam "SelectionTop", TranslateValueAbsToRel(.Top, False)
        cParams.AddParam "SelectionWidth", TranslateValueAbsToRel(.Width, True)
        cParams.AddParam "SelectionHeight", TranslateValueAbsToRel(.Height, False)
    End With
    
    'After universal values, we add any additional properties specific to the current selection shape
    Select Case m_SelectionShape
    
        'Rectangles, ellipses, and lines store corner coordinates
        Case ss_Rectangle, ss_Circle, ss_Line
            With m_CornersUnlocked
                cParams.AddParam "SelectionVectorLeft", TranslateValueAbsToRel(.Left, True)
                cParams.AddParam "SelectionVectorTop", TranslateValueAbsToRel(.Top, False)
                cParams.AddParam "SelectionVectorRight", TranslateValueAbsToRel(.Right, True)
                cParams.AddParam "SelectionVectorBottom", TranslateValueAbsToRel(.Bottom, False)
            End With
            
        'Polygon selections are trickier; they have a dynamic parameter list based on the number of points in the shape.
        Case ss_Polygon
            
            cParams.AddParam "SelectionPolygonCurvature", Me.GetSelectionProperty(sp_PolygonCurvature)
            cParams.AddParam "SelectionNumOfPoints", m_NumOfPolygonPoints
            
            If (m_NumOfPolygonPoints > 0) Then
                For i = 0 To m_NumOfPolygonPoints - 1
                    With m_PolygonPoints(i)
                        cParams.AddParam "SelectionPointX" & CStr(i + 1), TranslateValueAbsToRel(.x, True)
                        cParams.AddParam "SelectionPointY" & CStr(i + 1), TranslateValueAbsToRel(.y, False)
                    End With
                Next i
            End If
            
        'Lasso selections are similar to polygon selections
        Case ss_Lasso
            
            cParams.AddParam "SelectionSmoothStroke", Me.GetSelectionProperty(sp_SmoothStroke)
            cParams.AddParam "SelectionNumOfPoints", m_NumOfLassoPoints
            
            If (m_NumOfLassoPoints > 0) Then
                For i = 0 To m_NumOfLassoPoints - 1
                    With m_LassoPoints(i)
                        cParams.AddParam "SelectionPointX" & CStr(i + 1), TranslateValueAbsToRel(.x, True)
                        cParams.AddParam "SelectionPointY" & CStr(i + 1), TranslateValueAbsToRel(.y, False)
                    End With
                Next i
            End If
            
        'Wand selections have a few wand-specific fixed params
        Case ss_Wand
        
            'For weird historical reasons, wand selections store the wand origin point inside the unlocked selection rect
            With m_CornersUnlocked
                cParams.AddParam "SelectionWandX", TranslateValueAbsToRel(.Left, True)
                cParams.AddParam "SelectionWandY", TranslateValueAbsToRel(.Top, False)
            End With
            
            cParams.AddParam "SelectionWandTolerance", Me.GetSelectionProperty(sp_WandTolerance)
            cParams.AddParam "SelectionWandSampleMerged", Me.GetSelectionProperty(sp_WandSampleMerged)
            cParams.AddParam "SelectionWandSearchMode", Me.GetSelectionProperty(sp_WandSearchMode)
            cParams.AddParam "SelectionWandCompareMode", Me.GetSelectionProperty(sp_WandCompareMethod)
            
        'Raster selections cannot be returned this way
        Case Else
    
    End Select
    
    GetSelectionAsXML = cParams.GetParamString()
    
End Function

'By default, selection properties use absolute coordinates (e.g. [50, 50] for the center of a [100, 100] image).  When saving
' to file or recording in a macro, we want to switch to relative formats (e.g. [0.5, 0.5] for a coordinate at the center of any image).
' This allows selections to be correctly stretched to images of any arbitrary size.
'
'This function handles translation of coordinate values between absolute and relative geometries.
Private Function TranslateValueAbsToRel(ByVal srcFloat As Double, ByVal useWidth As Boolean) As Double
    Dim tmpDimension As Double
    If useWidth Then tmpDimension = m_parentPDImage.Width Else tmpDimension = m_parentPDImage.Height
    If (tmpDimension <> 0#) Then TranslateValueAbsToRel = srcFloat / tmpDimension Else TranslateValueAbsToRel = 0#
End Function

Private Function TranslateValueRelToAbs(ByVal srcFloat As Double, ByVal useWidth As Boolean) As Double
    Dim tmpDimension As Double
    If useWidth Then tmpDimension = m_parentPDImage.Width Else tmpDimension = m_parentPDImage.Height
    TranslateValueRelToAbs = srcFloat * tmpDimension
End Function

'Simple - use this to select the entire image attached to this selection object.  Note that this DOES NOT render the new selection on-screen,
' and it doesn't render a matching selection mask.  The caller is responsible for that.
Friend Sub SelectAll()

    'Set basic information about this selection
    SetSelectionShape ss_Rectangle
    SetSelectionProperty sp_Area, sa_Interior
    SetSelectionProperty sp_FeatheringRadius, 0
    SetSelectionProperty sp_Smoothing, ss_Antialiased
    SetSelectionProperty sp_RoundedCornerRadius, 0
    
    With m_CornersUnlocked
        .Left = 0
        .Top = 0
        .Right = m_parentPDImage.Width
        .Bottom = m_parentPDImage.Height
    End With
    
    With m_CornersLocked
        .Left = 0
        .Top = 0
        .Width = m_parentPDImage.Width
        .Height = m_parentPDImage.Height
    End With
    
    m_IsTransformable = True
    m_IsMaskReady = False
    
End Sub

'Request a 1:1 aspect ratio selection (squares, circles)
Friend Sub RequestSquare(ByVal requestChoice As Boolean, Optional ByVal forceSelectionRefresh As Boolean = False)
    
    m_IsSquare = requestChoice
    If (m_SelectionShape <> ss_Raster) And (m_SelectionShape <> ss_Lasso) And (m_SelectionShape <> ss_Polygon) And (m_SelectionShape <> ss_Wand) Then
    
        m_IsMaskReady = False
    
        'Redraw the selection if necessary
        If forceSelectionRefresh Then
            If (Not m_parentPDImage Is Nothing) Then
                UpdateInternalCoords
                CreateSelectionMask
            End If
        End If
        
    End If
    
End Sub

'Request a redraw of the selection mask.  We must do this when loading an Undo or Redo request after the image size has
' been changed; otherwise, OOB errors can occur (because the selection mask will be a different size than the image).
Friend Sub RequestNewMask()

    If (Not m_parentPDImage Is Nothing) Then
        
        UpdateInternalCoords
        
        'Transformable selections are rendered using polygon geometry; as such, redrawing them requires invoking
        ' whatever polygon code we used previously.
        If (m_SelectionShape <> ss_Raster) Then
            CreateSelectionMask
            
        'Raster selections are simply bitmaps.  We will pad them to match the new image size, but other
        ' than that, we do not change their existing composition.
        Else
        
        End If
        
    End If

End Sub

'Get/set a transformation type.  Note that this class only caches POIs *while a transformation is active*.  For generic things like
' MouseOver events, the current POI (if any) must be obtained manually via Selections.IsCoordSelectionPOI().
Friend Function GetActiveSelectionPOI() As PD_SelectionPOI
    If m_TransformModeActive Then
        GetActiveSelectionPOI = m_CurrentPOI
    Else
        GetActiveSelectionPOI = poi_Undefined
    End If
End Function

Friend Sub SetActiveSelectionPOI(ByVal newPOI As PD_SelectionPOI)
    m_CurrentPOI = newPOI
End Sub

Friend Sub OverrideTransformMode(ByVal newOverride As Boolean)
    m_TransformModeActive = newOverride
End Sub

'Get/set the selection's current coordinates.  At present, this is only used by line selections, as part of "compare mouse
' coordinates against selection coordinates" functionality.
'TODO: remove this
Friend Sub GetSelectionCoordinates(ByVal numPoint As Long, ByRef xDst As Double, ByRef yDst As Double)
    If (numPoint = 1) Then
        xDst = m_CornersUnlocked.Left
        yDst = m_CornersUnlocked.Top
    Else
        xDst = m_CornersUnlocked.Right
        yDst = m_CornersUnlocked.Bottom
    End If
End Sub

'Get/set a selection shape
Friend Function GetSelectionShape() As PD_SelectionShape
    GetSelectionShape = m_SelectionShape
End Function

Friend Sub SetSelectionShape(ByVal selShape As PD_SelectionShape)
    
    'Certain types of shapes are transformable.  Mark those now.
    Select Case selShape
        
        Case ss_Rectangle, ss_Circle, ss_Line
            m_IsTransformable = True
            
        Case ss_Lasso, ss_Polygon
            m_IsTransformable = True
            
        Case ss_Wand
            m_IsTransformable = False
        
    End Select
    
    If (m_SelectionShape <> ss_Raster) Then m_IsMaskReady = False
    m_SelectionShape = selShape
    
End Sub

'Return the number of polygon points
Friend Function GetNumOfPolygonPoints() As Long
    GetNumOfPolygonPoints = m_NumOfPolygonPoints
End Function

'Copy the current polygon point collection into an arbitrary destination array.  This is used by mouse coordinate checking functions.
Friend Sub GetPolygonPoints(ByRef ptFloatArray() As POINTFLOAT)
    
    If (m_NumOfPolygonPoints > 0) Then
        ReDim ptFloatArray(0 To m_NumOfPolygonPoints - 1) As POINTFLOAT
        CopyMemory ByVal VarPtr(ptFloatArray(0)), ByVal VarPtr(m_PolygonPoints(0)), m_NumOfPolygonPoints * 8
    Else
        ReDim ptFloatArray(0) As POINTFLOAT
    End If
    
End Sub

'For lasso selections, the canvas needs to know if the current lasso selection is open (e.g. still under construction) or closed.
Friend Function GetLassoClosedState() As Boolean
    GetLassoClosedState = m_LassoClosed
End Function

Friend Sub SetLassoClosedState(ByVal newState As Boolean)
    m_LassoClosed = newState
End Sub

'For polygon selections, the canvas needs to know if the current polygon selection is open (e.g. still under construction) or closed.
Friend Function GetPolygonClosedState() As Boolean
    GetPolygonClosedState = m_PolygonClosed
End Function

Friend Sub SetPolygonClosedState(ByVal newState As Boolean)
    m_PolygonClosed = newState
End Sub

'For lasso and polygon selections, this function will return the current selecton region as a GDI+ region handle.  The mouse coordinate
' checker code uses this to see if the mouse cursor is currently within the bounds of the selection area.
Friend Function GetGdipRegionForSelection() As Long
    If (m_SelectionShape = ss_Polygon) Then
        GetGdipRegionForSelection = GDI_Plus.GetGDIPlusRegionFromPoints(m_NumOfPolygonPoints, VarPtr(m_PolygonPoints(0)), GP_FM_Winding, True, GetSelectionProperty_Float(sp_PolygonCurvature))
    Else
        GetGdipRegionForSelection = GDI_Plus.GetGDIPlusRegionFromPoints(m_NumOfLassoPoints, VarPtr(m_LassoPoints(0)), GP_FM_Winding, True, GetSelectionProperty_Float(sp_SmoothStroke))
    End If
End Function

'For lasso and polygon selections, this function will calculate a boundary rect using GDI+.  This is relevant when curvature is active,
' because the curvature bounds may extend outside the points used to create the selection shape.
Friend Sub FillBoundsUsingGDIPlus()
    
    'This function is only relevant for lasso and polygon selections
    If ((m_SelectionShape = ss_Lasso) Or (m_SelectionShape = ss_Polygon)) Then
    
        If (m_SelectionShape = ss_Polygon) Then
            m_CornersLocked = GDI_Plus.GetGDIPlusUnionFromPointsAndImage(m_NumOfPolygonPoints, VarPtr(m_PolygonPoints(0)), m_parentPDImage, GP_FM_Winding, True, GetSelectionProperty_Float(sp_PolygonCurvature))
        Else
            m_CornersLocked = GDI_Plus.GetGDIPlusUnionFromPointsAndImage(m_NumOfLassoPoints, VarPtr(m_LassoPoints(0)), m_parentPDImage, GP_FM_Winding, True, GetSelectionProperty_Float(sp_SmoothStroke))
        End If
        
        m_Bounds = m_CornersLocked
        
    End If
    
End Sub

'Takes x and y coordinates (from a _MouseDown event, typically) and uses them in a manner specified by the current transform operation.
' Note that this this should only be called after a transformation type has been set (via SetActiveSelectionPOI(), above).
Friend Sub SetInitialTransformCoordinates(ByVal x As Double, ByVal y As Double)

    'If new transform coordinates are being set, this selection must be "unlocked" first
    m_IsLocked = False
    m_TransformModeActive = True
    
    'Different selection types handle transformation differently.  For example, rectangular selections can be resized
    ' in multiple directions, but a line selection will only move its endpoints.  So we must sort input twice:
    ' 1) by selection type
    ' 2) by transformation type.
    
    'Rectangular and elliptical selections are handled identically
    If ((m_SelectionShape = ss_Rectangle) Or (m_SelectionShape = ss_Circle)) Then
        
        'Based on the transform mode, set the initial points accordingly
        Select Case m_CurrentPOI
        
            'Failsafe check for undefined transforms
            Case poi_Undefined
                Debug.Print "Selection transform initiated on a non-existent point - FIX THIS!"
            
            'Corners
            Case poi_CornerNW
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left + m_CornersLocked.Width
                    .Top = m_CornersLocked.Top + m_CornersLocked.Height
                    .Right = m_CornersLocked.Left
                    .Bottom = m_CornersLocked.Top
                End With
            
            Case poi_CornerNE
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left
                    .Top = m_CornersLocked.Top + m_CornersLocked.Height
                    .Right = x
                    .Bottom = y
                End With
            
            Case poi_CornerSE
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left
                    .Top = m_CornersLocked.Top
                    .Right = x
                    .Bottom = y
                End With
            
            Case poi_CornerSW
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left + m_CornersLocked.Width
                    .Top = m_CornersLocked.Top
                    .Right = x
                    .Bottom = y
                End With
            
            Case poi_EdgeN
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left
                    .Right = m_CornersLocked.Left + m_CornersLocked.Width
                    .Top = m_CornersLocked.Top + m_CornersLocked.Height
                    .Bottom = y
                End With
            
            Case poi_EdgeE
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left
                    .Right = x
                    .Top = m_CornersLocked.Top
                    .Bottom = m_CornersLocked.Top + m_CornersLocked.Height
                End With
            
            Case poi_EdgeS
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left
                    .Right = m_CornersLocked.Left + m_CornersLocked.Width
                    .Top = m_CornersLocked.Top
                    .Bottom = y
                End With
            
            Case poi_EdgeW
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left + m_CornersLocked.Width
                    .Right = x
                    .Top = m_CornersLocked.Top
                    .Bottom = m_CornersLocked.Top + m_CornersLocked.Height
                End With
            
            '8 - interior of selection, not near a corner or edge
            Case poi_Interior
                With m_CornersUnlocked
                    m_MoveXDist = x - m_CornersLocked.Left
                    m_MoveYDist = y - m_CornersLocked.Top
                End With
        
        End Select
        
    'Line selections
    ElseIf (m_SelectionShape = ss_Line) Then
    
        'Based on the transform mode, set the initial points accordingly
        Select Case m_CurrentPOI
    
            'Case "undefined" should never occur, but if it does - treat this like a normal initial coordinate call
            Case poi_Undefined
                With m_CornersUnlocked
                    .Left = x
                    .Top = y
                    .Right = x
                    .Bottom = y
                End With
            
            'First point is being moved
            Case 0
                With m_CornersUnlocked
                    .Left = x
                    .Top = y
                End With
            
            'Second point is being moved
            Case 1
                With m_CornersUnlocked
                    .Right = x
                    .Bottom = y
                End With
        
        End Select
                
    'Polygons can be both moved and resized (by click-dragging individual points)
    ElseIf (m_SelectionShape = ss_Polygon) Then
    
        'Based on the transform mode, set the initial points accordingly
        Select Case m_CurrentPOI
    
            'Interior of the selection, which triggers a move transformation
            Case m_NumOfPolygonPoints
                m_MoveXDist = x
                m_MoveYDist = y
                
                'For the transform to work, we need to apply any transformed data to the *original polygon points*.
                ' Make a backup copy of the original array now.
                ReDim m_PolygonPointsBackup(0 To m_NumOfPolygonPoints - 1) As POINTFLOAT
                CopyMemory ByVal VarPtr(m_PolygonPointsBackup(0)), ByVal VarPtr(m_PolygonPoints(0)), m_NumOfPolygonPoints * 8
                
            'Other transforms don't actually require special handling, as all the work is done in the
            ' SetAdditionalTransformCoordinates function.
            Case Else
                
        End Select
        
    'Lasso selections can only be moved, so the only valid transform type is type 0
    ElseIf (m_SelectionShape = ss_Lasso) Then
    
        'Based on the transform mode, set the initial points accordingly
        Select Case m_CurrentPOI
    
            'Interior of the selection, which triggers a move transformation
            Case 0
                m_MoveXDist = x
                m_MoveYDist = y
                
                'For the transform to work, we need to apply any transformed data to the *original lasso points*.
                ' Make a backup copy of the original array now.
                ReDim m_LassoPointsBackup(0 To m_NumOfLassoPoints - 1) As POINTFLOAT
                CopyMemory ByVal VarPtr(m_LassoPointsBackup(0)), ByVal VarPtr(m_LassoPoints(0)), m_NumOfLassoPoints * 8
                
        End Select
            
    'Wand selections don't care about transformation type
    ElseIf (m_SelectionShape = ss_Wand) Then
        With m_CornersUnlocked
            .Left = x
            .Top = y
        End With
        
    'Any other selection types cannot be transformed
    End If
    
    UpdateInternalCoords
    
End Sub

'Takes x and y coordinates (from a _MouseDown event, typically) and stores them internally
Friend Sub SetInitialCoordinates(ByVal x As Double, ByVal y As Double)
    
    'If new initial coordinates are being set, this selection must be "unlocked"
    m_IsLocked = False
    
    'The use of setInitialCoordinates means this is not a transformation
    m_TransformModeActive = False
    
    'If we're setting new initial coordinates, the mask is not (by definition) ready
    m_IsMaskReady = False
    
    With m_CornersUnlocked
        .Left = x
        .Top = y
    End With
    
    'Set the second set of point to match the first set
    Select Case m_SelectionShape
        
        Case ss_Rectangle, ss_Circle, ss_Line
            
            With m_CornersUnlocked
                .Right = x
                .Bottom = y
            End With
            
            'Also, mark selections of any of these shapes (rectangle, circle, line) as transformable
            m_IsTransformable = True
        
        'Polygon selections require us to initialize a whole bunch of polygon tracking variables
        Case ss_Polygon
            m_IsTransformable = True
            
            m_NumOfPolygonPoints = 1
            ReDim m_PolygonPoints(0 To 127) As POINTFLOAT
            
            m_PolygonPoints(0).x = x
            m_PolygonPoints(0).y = y
            
            'Also set up a dummy set of initial boundary coordinates
            With m_CornersLocked
                .Left = x
                .Top = y
                .Width = 1
                .Height = 1
            End With
            
        'Lasso selections have limited transform capabilities, and they have a variable number of points
        Case ss_Lasso
            m_IsTransformable = True
            
            m_NumOfLassoPoints = 1
            ReDim m_LassoPoints(0 To 127) As POINTFLOAT
            
            m_LassoPoints(0).x = x
            m_LassoPoints(0).y = y
            
            'Also set up a dummy set of initial boundary coordinates
            With m_CornersLocked
                .Left = x
                .Top = y
                .Width = 1
                .Height = 1
            End With
        
        'Wand selections don't support transforms
        Case ss_Wand
            m_IsTransformable = False
            With m_CornersUnlocked
                .Left = x
                .Top = y
            End With
        
        'Other selection types (e.g. raster selections) cannot be created this way
        Case Else
        
    End Select
    
    UpdateInternalCoords
    
End Sub

'Takes x and y coordinates (from a _MouseMove event, typically) and stores them internally
Friend Sub SetAdditionalCoordinates(ByVal x As Double, ByVal y As Double)
    
    Dim newMaskRequired As Boolean
    newMaskRequired = True
    
    'Check for an active transformation mode.  (A transformation is something like resizing or moving an existing selection,
    ' versus drawing a new one from scratch.)
    If m_TransformModeActive Then
        
        Dim i As Long
        
        Select Case m_SelectionShape
    
            Case ss_Rectangle, ss_Circle
            
                'Based on the transform mode, set the initial points accordingly
                Select Case m_CurrentPOI
                
                    'Case -1 should never occur, but if it does - treat this like a normal subsequent coordinate call
                    Case poi_Undefined
                    
                    'Corners
                    Case poi_CornerNW, poi_CornerNE, poi_CornerSE, poi_CornerSW
                        m_CornersUnlocked.Right = x
                        m_CornersUnlocked.Bottom = y
                    
                    'Edges
                    Case poi_EdgeN
                        m_CornersUnlocked.Bottom = y
                        
                    Case poi_EdgeE
                        m_CornersUnlocked.Right = x
                        
                    Case poi_EdgeS
                        m_CornersUnlocked.Bottom = y
                        
                    Case poi_EdgeW
                        m_CornersUnlocked.Right = x
                        
                    'Interior
                    Case poi_Interior
                        With m_CornersUnlocked
                            .Left = x - m_MoveXDist
                            .Top = y - m_MoveYDist
                            .Right = .Left + m_CornersLocked.Width
                            .Bottom = .Top + m_CornersLocked.Height
                        End With
                        
                End Select
                
                'If a transform mode is active, re-mark the selection as being transformable
                m_IsTransformable = True
        
            Case ss_Line
            
                'Based on the transform mode, set the initial points accordingly
                Select Case m_CurrentPOI
            
                    'Case -1 should never occur.  (-1 represents an invalid transformation request)
                    Case -1
                    
                    'First point is being moved
                    Case 0
                        m_CornersUnlocked.Left = x
                        m_CornersUnlocked.Top = y
                        
                    'Second point is being moved
                    Case 1
                        m_CornersUnlocked.Right = x
                        m_CornersUnlocked.Bottom = y
                        
                End Select
                
                'If a transform mode is active, re-mark the selection as being transformable
                m_IsTransformable = True
            
            'Polygon transforms consist of either moving an individual polygon point, or moving the entire polygon array
            Case ss_Polygon
            
                'Based on the transform mode, set the initial points accordingly
                Select Case m_CurrentPOI
            
                    'Case -1 should never occur.  (-1 represents an invalid transformation request)
                    Case -1
                    
                    'Move transform
                    Case m_NumOfPolygonPoints
                        
                        'Rebuild the main polygon array by copying all points from the backup array, and applying the current
                        ' x/y transformation distance to them.
                        For i = 0 To m_NumOfPolygonPoints - 1
                            m_PolygonPoints(i).x = m_PolygonPointsBackup(i).x + (x - m_MoveXDist)
                            m_PolygonPoints(i).y = m_PolygonPointsBackup(i).y + (y - m_MoveYDist)
                        Next i
                    
                    'Anything else is just moving a polygon point
                    Case Else
                    
                        With m_PolygonPoints(m_CurrentPOI)
                            .x = x
                            .y = y
                        End With
                        
                End Select
                
                'If a transform mode is active, re-mark the selection as being transformable
                m_IsTransformable = True
            
            'Lasso transforms require us to transform the entire lasso array
            Case ss_Lasso
            
                'Based on the transform mode, set the initial points accordingly
                Select Case m_CurrentPOI
            
                    'Case -1 should never occur.  (-1 represents an invalid transformation request)
                    Case -1
                    
                    'Move transform
                    Case 0
                        
                        'Rebuild the main lasso array by copying all points from the backup array, and applying the current
                        ' x/y transformation distance to them.
                        For i = 0 To m_NumOfLassoPoints - 1
                            m_LassoPoints(i).x = m_LassoPointsBackup(i).x + (x - m_MoveXDist)
                            m_LassoPoints(i).y = m_LassoPointsBackup(i).y + (y - m_MoveYDist)
                        Next i
                        
                End Select
                
                'If a transform mode is active, re-mark the selection as being transformable
                m_IsTransformable = True
            
            'Wand selections are not transformable
            Case ss_Wand
                Debug.Print "Transform initiated on a wand selection - FIX THIS!"
            
        End Select
                
    'This is not a transform, meaning the selection is being created for the first time.  For standard selection types,
    ' this simply means copying the passed (x, y) values.  Polygon/lasso selections are more complicated.
    Else
        
        Select Case m_SelectionShape
        
            'Rectangle, ellipse, and line selections are easy - just copy the passed (x, y) values into (x2, y2)
            Case ss_Rectangle, ss_Circle, ss_Line
                m_CornersUnlocked.Right = x
                m_CornersUnlocked.Bottom = y
                
            'Polygon selections will increment the current polygon array by one, adding the new point as the latest polygon coordinate
            Case ss_Polygon
            
                'Make room in the point array
                m_NumOfPolygonPoints = m_NumOfPolygonPoints + 1
                If (m_NumOfPolygonPoints > UBound(m_PolygonPoints)) Then ReDim Preserve m_PolygonPoints(0 To m_NumOfPolygonPoints * 2 - 1) As POINTFLOAT
                    
                'Store the new point
                m_PolygonPoints(m_NumOfPolygonPoints - 1).x = x
                m_PolygonPoints(m_NumOfPolygonPoints - 1).y = y
                
            'Lasso selections will increment the current lasso array, and add the new point to its list
            Case ss_Lasso
            
                'Perform a quick check to make sure this point isn't a duplicate of the previous point.  With high-DPI mice,
                ' this is a distinct possibility.
                If ((m_LassoPoints(m_NumOfLassoPoints - 1).x <> x) Or (m_LassoPoints(m_NumOfLassoPoints - 1).y <> y)) Then
            
                    'Make room in the point array
                    m_NumOfLassoPoints = m_NumOfLassoPoints + 1
                    If (m_NumOfLassoPoints > UBound(m_LassoPoints)) Then ReDim Preserve m_LassoPoints(0 To m_NumOfLassoPoints * 2 - 1) As POINTFLOAT
                    
                    'Store the new point
                    m_LassoPoints(m_NumOfLassoPoints - 1).x = x
                    m_LassoPoints(m_NumOfLassoPoints - 1).y = y
                    
                End If
            
            'Wand selections can have their point of interest moved, but they don't actually support "additional" coordinates
            Case ss_Wand
                If (m_CornersUnlocked.Left <> x) Or (m_CornersUnlocked.Top <> y) Then
                    m_CornersUnlocked.Left = x
                    m_CornersUnlocked.Top = y
                Else
                    newMaskRequired = False
                End If
            
            Case Else
            
        End Select
        
    End If
    
    'Update the bounding rect for the selection as a whole, based on the new coordinates
    UpdateInternalCoords newMaskRequired
    
End Sub

'Has this selection been locked in?
Friend Function IsLockedIn() As Boolean
    IsLockedIn = m_IsLocked
End Function

'Is this selection transformable?
Friend Function IsTransformable() As Boolean
    IsTransformable = m_IsTransformable
End Function

'If the user is using the SHIFT key to request a square-shaped (or circle-shaped) selection, this function will be called.
Private Sub MakeCoordinatesSquare()

    Select Case m_SelectionShape
                
        Case ss_Rectangle, ss_Circle, ss_Line
            
            With m_CornersUnlocked
                
                If (.Left < .Right) Then
                    If (.Top < .Bottom) Then
                        If (Abs(.Left - .Right) > Abs(.Top - .Bottom)) Then
                            .Bottom = .Top + Abs(.Left - .Right)
                        Else
                            .Right = .Left + Abs(.Top - .Bottom)
                        End If
                    Else
                        If (Abs(.Left - .Right) > Abs(.Top - .Bottom)) Then
                            .Bottom = .Top - Abs(.Left - .Right)
                        Else
                            .Right = .Left + Abs(.Top - .Bottom)
                        End If
                    End If
                Else
                    If (.Top < .Bottom) Then
                        If (Abs(.Left - .Right) > Abs(.Top - .Bottom)) Then
                            .Bottom = .Top + Abs(.Left - .Right)
                        Else
                            .Right = .Left - Abs(.Top - .Bottom)
                        End If
                    Else
                        If (Abs(.Left - .Right) > Abs(.Top - .Bottom)) Then
                            .Bottom = .Top - Abs(.Left - .Right)
                        Else
                            .Right = .Left - Abs(.Top - .Bottom)
                        End If
                    End If
                End If
                
            End With
            
        'Other selection types do not currently support square modifiers
        Case Else
        
    End Select
    
End Sub

'Whenever internal vector or coordinate values are changed, this sub needs to be called to update the left/right/width/height
' values accordingly.  Note that for some selection types - e.g. lasso - a full scan of all available coordinates must be performed,
' which can be performance-intensive if the shape is complex.  As such, try not to call this function any more than is necessary.
Private Sub UpdateInternalCoords(Optional ByVal forciblyResetMask As Boolean = True)

    'This function only needs to be run if the selection is stored in vector format.  If it is not, a bounding rect
    ' will already be correctly set.
    If (m_SelectionShape <> ss_Raster) Then
    
        Dim i As Long
        Dim selMaxX As Long, selMaxY As Long
    
        'Mark the selection mask as "not ready", as it will need to be redrawn after new coordinates are set
        If forciblyResetMask Then m_IsMaskReady = False

        'If a square (1:1 aspect ratio) selection has been requested, calculate new coordinates now.
        
        ' (This set of if/then blocks looks complicated, but it's actually very simple - we simply have to account for every variation
        '  of quadrants, because the selection can be drawn up or down in both directions, giving eight possible variants of x1 </> x2
        '  and y1 </> y2.  By covering all those cases, square selections can be drawn in any direction.)
        If m_IsSquare Then MakeCoordinatesSquare
        
        'Finally, calculate a left, top, width and height for this selection based off the current individual coordinate values
        Select Case m_SelectionShape
        
            'Rectangles, ellipses, and lines all use the same (x1, y1) - (x2, y2) coordinate system, so finding selection bounds is easy.
            Case ss_Rectangle, ss_Circle, ss_Line
                
                With m_CornersUnlocked
                    
                    If (.Left < .Right) Then
                        m_CornersLocked.Left = .Left
                        m_CornersLocked.Width = .Right - .Left
                    Else
                        m_CornersLocked.Left = .Right
                        m_CornersLocked.Width = .Left - .Right
                    End If
                    
                    If (.Top < .Bottom) Then
                        m_CornersLocked.Top = .Top
                        m_CornersLocked.Height = .Bottom - .Top
                    Else
                        m_CornersLocked.Top = .Bottom
                        m_CornersLocked.Height = .Top - .Bottom
                    End If
                    
                End With
            
            'Polygon selections require us to search all polygon points in order to construct a bounding rect.
            Case ss_Polygon
                
                m_CornersLocked.Left = LONG_MAX
                m_CornersLocked.Top = LONG_MAX
                selMaxX = -LONG_MAX
                selMaxY = -LONG_MAX
                
                'Search the entire polygon array for new left/top/width/height values
                For i = 0 To m_NumOfPolygonPoints - 1
                    
                    With m_PolygonPoints(i)
                        If (.x < m_CornersLocked.Left) Then m_CornersLocked.Left = .x
                        If (.y < m_CornersLocked.Top) Then m_CornersLocked.Top = .y
                        If (.x > selMaxX) Then selMaxX = .x
                        If (.y > selMaxY) Then selMaxY = .y
                    End With
                    
                Next i
                
                m_CornersLocked.Width = selMaxX - m_CornersLocked.Left
                m_CornersLocked.Height = selMaxY - m_CornersLocked.Top
                
                'If curvature is active, we'll use GDI+ to find our bounds more precisely
                If (Me.IsLockedIn And (GetSelectionProperty_Float(sp_PolygonCurvature) > 0)) Then Me.FillBoundsUsingGDIPlus
            
            'Like polygon selection, lasso selections require us to search all lasso points in order to construct
            ' a bounding rect.
            Case ss_Lasso
            
                m_CornersLocked.Left = LONG_MAX
                m_CornersLocked.Top = LONG_MAX
                selMaxX = -LONG_MAX
                selMaxY = -LONG_MAX
                
                'Search the entire lasso array for new left/top/width/height values
                For i = 0 To m_NumOfLassoPoints - 1
                    
                    With m_LassoPoints(i)
                        If (.x < m_CornersLocked.Left) Then m_CornersLocked.Left = .x
                        If (.y < m_CornersLocked.Top) Then m_CornersLocked.Top = .y
                        If (.x > selMaxX) Then selMaxX = .x
                        If (.y > selMaxY) Then selMaxY = .y
                    End With
                    
                Next i
                
                m_CornersLocked.Width = selMaxX - m_CornersLocked.Left
                m_CornersLocked.Height = selMaxY - m_CornersLocked.Top
                
                'If curvature is active, we'll use GDI+ to find our bounds more precisely
                If (Me.IsLockedIn And (GetSelectionProperty_Float(sp_SmoothStroke) > 0)) Then Me.FillBoundsUsingGDIPlus
            
            'Wand selections require manual bounds-checking
            Case ss_Wand
                If Me.IsLockedIn Then Me.FindNewBoundsManually True
                
            Case Else
            
        End Select
        
    End If

End Sub

'Because selections can be created beyond the parent image's borders, it is sometimes necessary to check if ALL selection coordinates
' lie off the image.  If this is the case, we don't want to finalize the current selection - we want to forget it.
'
'Note that this function requires the m_CornersLocked.Left/Top/Width/Height values to be correctly set prior to calling.
'
'Returns: TRUE if this selection lies completely outside its parent image boundaries
Friend Function AreAllCoordinatesInvalid() As Boolean
    
    Select Case m_SelectionShape
    
        'Rectangles, ellipses, and lines are easy - check the bounding box, and if it lies completely outside the image,
        ' reject it.  Note that this occurs before a final bound rect has been calculated, so you have no choice but to
        ' rely on intermediate x/y coords instead of the m_Bounds.Left/m_Bounds.Top etc values.
        Case ss_Rectangle, ss_Circle, ss_Line
            
            If (m_CornersLocked.Left + m_CornersLocked.Width <= 0) Then AreAllCoordinatesInvalid = True
            If (m_CornersLocked.Left > m_parentPDImage.Width) Then AreAllCoordinatesInvalid = True
            If (m_CornersLocked.Top + m_CornersLocked.Height <= 0) Then AreAllCoordinatesInvalid = True
            If (m_CornersLocked.Top > m_parentPDImage.Height) Then AreAllCoordinatesInvalid = True
           
        'Polygon and Lasso are a bit more complicated.  If curvature is not active, we can use existing m_CornersLocked.Left/Top etc bounds,
        ' but if it is active, we need to perform a manual search for boundaries.
        Case ss_Polygon, ss_Lasso
        
            'If curvature is active, use GDI+ to update the current boundaries.
            If ((m_SelectionShape = ss_Polygon) And (GetSelectionProperty_Float(sp_PolygonCurvature) <> 0)) Or ((m_SelectionShape = ss_Lasso) And (GetSelectionProperty_Float(sp_SmoothStroke) <> 0)) Then Me.FillBoundsUsingGDIPlus
            
            If (m_CornersLocked.Left + m_CornersLocked.Width <= 0) Then AreAllCoordinatesInvalid = True
            If (m_CornersLocked.Left >= m_parentPDImage.Width) Then AreAllCoordinatesInvalid = True
            If (m_CornersLocked.Top + m_CornersLocked.Height <= 0) Then AreAllCoordinatesInvalid = True
            If (m_CornersLocked.Top >= m_parentPDImage.Height) Then AreAllCoordinatesInvalid = True
        
        'Wand selections are valid if the (x1, y1) coordinate pair falls inside the image
        Case ss_Wand
            If (m_CornersUnlocked.Left < 0) Then AreAllCoordinatesInvalid = True
            If (m_CornersUnlocked.Top < 0) Then AreAllCoordinatesInvalid = True
            If (m_CornersUnlocked.Left >= m_parentPDImage.Width) Then AreAllCoordinatesInvalid = True
            If (m_CornersUnlocked.Top >= m_parentPDImage.Height) Then AreAllCoordinatesInvalid = True
            
        'In the future, additional selection types can be handled here.
        Case Else
            
    End Select
    
End Function

'Nudge the selection in a given direction.  This function is supplied as a convenience for SELECTION TYPES WHOSE POSITION CANNOT
' BE MODIFIED BY TEXT BOX.  If a selection's position can be modified via text box (e.g. rectangle, ellipse, etc), you should
' use the UpdateViaTextBox() function instead, to ensure that text box and internal positions are properly synched.
Friend Sub NudgeSelection(Optional ByVal hOffset As Double = 0#, Optional ByVal vOffset As Double = 0#)

    Dim i As Long

    Select Case m_SelectionShape
    
        Case ss_Rectangle, ss_Circle, ss_Line
            Debug.Print "nudgeSelection function was used on an invalid selection type - FIX THIS!"
    
        Case ss_Polygon
        
            'Apply the new offsets to all points in the polygon
            For i = 0 To m_NumOfPolygonPoints - 1
                With m_PolygonPoints(i)
                    .x = .x + hOffset
                    .y = .y + vOffset
                End With
            Next i
            
            'We also need to update the selection's bounding rect
            UpdateInternalCoords
        
        Case ss_Lasso
            
            'Apply the new offsets to all points in the lasso array
            For i = 0 To m_NumOfLassoPoints - 1
                With m_LassoPoints(i)
                    .x = .x + hOffset
                    .y = .y + vOffset
                End With
            Next i
            
            'We also need to update the selection's bounding rect
            UpdateInternalCoords
            
        Case ss_Wand
            m_CornersUnlocked.Left = m_CornersUnlocked.Left + hOffset
            m_CornersUnlocked.Top = m_CornersUnlocked.Top + vOffset
            
            If (m_CornersUnlocked.Left < 0) Then m_CornersUnlocked.Left = 0
            If (m_CornersUnlocked.Top < 0) Then m_CornersUnlocked.Top = 0
            If (m_CornersUnlocked.Left >= m_parentPDImage.Width) Then m_CornersUnlocked.Left = m_parentPDImage.Width - 1
            If (m_CornersUnlocked.Top >= m_parentPDImage.Height) Then m_CornersUnlocked.Top = m_parentPDImage.Height - 1
            
            'We also need to update the selection's bounding rect
            UpdateInternalCoords
            
    End Select

End Sub

'Polygon selections are not as fiddly as lasso selections, but it can still be useful to remove the last-clicked point.
' This function (typically triggered via the BACKSPACE key) can be used to remove the last-created polygon point.
Friend Sub RemoveLastPolygonPoint()
    If (m_NumOfPolygonPoints > 1) Then m_NumOfPolygonPoints = m_NumOfPolygonPoints - 1
End Sub

'Lasso selections have the unique burden of being somewhat unfavorable to user error.  This function (typically triggered via
' the BACKSPACE key) can be used to retreat the lasso position and potentially correct any positioning errors.  The calling
' function must supply two Double-type variables, which will receive the new cursor position IN IMAGE COORDINATES.  The calling
' function is responsible for translating these to screen coordinates and actually applying the cursor repositioning.
Friend Sub RetreatLassoPosition(ByRef newCursorX_ImgCoords As Double, ByRef newCursorY_ImgCoords As Double)

    'Determine a point to retreat to.  We could do this a number of different ways, but because PD generally favors quality
    ' over all else, we're going to complicate it a bit.  The goal is to retreat a distance of ten pixels, which may correspond
    ' to any number of actual lasso points.
    Dim newLassoIndex As Long
    newLassoIndex = m_NumOfLassoPoints - 1
    
    Dim netDistance As Double
    netDistance = 0
    
    'Start calculating the net distance traveled by the lasso.  Once a distance of 10 pixels is exceeded, set that as our new
    ' lasso position.
    Do While (netDistance < 10) And (newLassoIndex > 0)
    
        'Calculate a distance between this coordinate and the previous one.
        netDistance = netDistance + Math_Functions.DistanceTwoPoints(m_LassoPoints(newLassoIndex).x, m_LassoPoints(newLassoIndex).y, m_LassoPoints(newLassoIndex - 1).x, m_LassoPoints(newLassoIndex - 1).y)
    
        'Decrement the test index and repeat
        newLassoIndex = newLassoIndex - 1
    
    Loop
    
    'Reposition the lasso point index to match
    m_NumOfLassoPoints = newLassoIndex + 1
    If (m_NumOfLassoPoints < 1) Then m_NumOfLassoPoints = 1
    
    'Return the new cursor coordinates at this position, then exit
    newCursorX_ImgCoords = m_LassoPoints(m_NumOfLassoPoints - 1).x
    newCursorY_ImgCoords = m_LassoPoints(m_NumOfLassoPoints - 1).y

End Sub

'Update this selection using the values in the main form's selection text boxes
Friend Sub UpdateViaTextBox()

    'Ignore text box update requests until the selection is locked in
    If (Not Me.IsLockedIn) Then Exit Sub

    m_RejectRefreshRequests = True
    
    Dim subpanelOffset As Long
    subpanelOffset = Selections.GetSelectionSubPanelFromSelectionShape(pdImages(g_CurrentImage)) * 4
    
    'Check all text box entries for validity, then update the corresponding selection values.
    Select Case m_SelectionShape
    
        'Rectangles, ellipses, and lines all support movement via text box
        Case ss_Rectangle, ss_Circle, ss_Line
            m_IsMaskReady = False
            If toolpanel_Selections.tudSel(subpanelOffset + 0).IsValid(False) Then m_CornersLocked.Left = toolpanel_Selections.tudSel(subpanelOffset + 0)
            If toolpanel_Selections.tudSel(subpanelOffset + 1).IsValid(False) Then m_CornersLocked.Top = toolpanel_Selections.tudSel(subpanelOffset + 1)
            If toolpanel_Selections.tudSel(subpanelOffset + 2).IsValid(False) Then m_CornersLocked.Width = toolpanel_Selections.tudSel(subpanelOffset + 2)
            If toolpanel_Selections.tudSel(subpanelOffset + 3).IsValid(False) Then m_CornersLocked.Height = toolpanel_Selections.tudSel(subpanelOffset + 3)
            
        'I haven't decided if other selection types will support movement via text box...
        Case Else
        
    End Select
    
    'For some selection types, we need to update more than just the m_CornersLocked.Left/Top/Width/Height values.
    Select Case m_SelectionShape
    
        'Adjust the x1, y1, x2, y2 values to match the retrieved text box values
        Case ss_Rectangle, ss_Circle
            m_CornersUnlocked.Left = m_CornersLocked.Left
            m_CornersUnlocked.Top = m_CornersLocked.Top
            m_CornersUnlocked.Right = m_CornersLocked.Left + m_CornersLocked.Width
            m_CornersUnlocked.Bottom = m_CornersLocked.Top + m_CornersLocked.Height
        
        'Adjust the x1, y1, x2, y2 values to match the retrieved text box values, but note that x2 and y2 do not correspond to
        ' width and height on line selections; they correspond to x2/y2 coordinates.
        Case ss_Line
            m_CornersUnlocked.Left = m_CornersLocked.Left
            m_CornersUnlocked.Top = m_CornersLocked.Top
            m_CornersUnlocked.Right = m_CornersLocked.Width
            m_CornersUnlocked.Bottom = m_CornersLocked.Height
            
        Case Else
        
    End Select
    
    m_RejectRefreshRequests = False
    
End Sub

'"Lock-in" a selection. Typically this is prompted by a _MouseUp event
Friend Sub LockIn()
    
    'Mark this selection as locked-in
    m_IsLocked = True
        
    'For vector selections, update the internal coordinates one final time
    If (m_SelectionShape <> ss_Raster) Then
        
        'The final thing we need to check for is the width and height, which may be still be zero at this point.
        ' Due to the way outside filters and effects use selection bounding rects, we can't allow selections of size 0.
        If (m_CornersLocked.Width < 1) Then m_CornersLocked.Width = 1
        If (m_CornersLocked.Height < 1) Then m_CornersLocked.Height = 1
        
    End If
        
End Sub

'"Unlock" a selection
Friend Sub LockRelease()
    m_IsLocked = False
End Sub

'For polygon and lasso selection, it reduces risk to reset a few custom trackers when a selection is erased.
Friend Sub EraseCustomTrackers()
    m_NumOfLassoPoints = 0
    m_NumOfPolygonPoints = 0
    m_PolygonClosed = False
    m_LassoClosed = False
End Sub

'Create a selection mask based on the current selection type.  A few items to note:
' 1) The selection mask is always the size of the full image.  This makes transforms (e.g. "grow selection") much easier to handle.
' 2) Black pixels (0) in the mask represent unselected pixels in the image.  White (255) represents selected.  Other values can be
'     used to specify aliasing or partial selections.
' 3) The selection mask is stored as a pdDIB object, so any image filters can be applied to it.
' 4) For shape-based selections (rectangle, square, etc), the selection's dimensions need to be set BEFORE calling this function.  This function
'     relies on things like m_CornersLocked.Left and m_CornersLocked.Width to know where to render the mask.  For non-shape-based selections, this function will call a
'     separate function to find the bounding rect.
Private Sub CreateSelectionMask()

    Debug.Print "(" & Timer & ") Creating mask now..."
    
    'Note that a mask has been created for this image.  This is important for saving/loading selections, as a new mask must be generated
    ' if one isn't already present.
    ' TODO: see if we can move this to the end of the function without consequences.
    m_MaskHasBeenCreated = True
    
    'Debug msg: trying to minimize redundant mask creation requests
    If m_IsMaskReady Then Debug.Print "Selection mask is marked as READY, but a new mask was requested.  FIX THIS!"
    
    'If the current selection is raster-type, this function should not have been called!
    If (m_SelectionShape = ss_Raster) Then
        Debug.Print "Mask redraw requested for raster-type selection - FIX THIS!"
        m_IsMaskReady = True
        Exit Sub
    End If
    
    Dim maskBackColor As Long, maskForeColor As Long, maskBackByte As Byte
    
    'Interior and exterior selections are rendered using identical code; the only difference is the colors used
    If (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Then
        maskBackColor = RGB(255, 255, 255)
        maskBackByte = 255
        maskForeColor = RGB(0, 0, 0)
    Else
        maskBackColor = RGB(0, 0, 0)
        maskBackByte = 0
        maskForeColor = RGB(255, 255, 255)
    End If

    'Some selection types (line selections) need to know max/min values, which are separate from left/top/width/height
    Dim minX As Long, maxX As Long, minY As Long, maxY As Long
    If (m_CornersUnlocked.Left < m_CornersUnlocked.Right) Then
        minX = m_CornersUnlocked.Left
        maxX = m_CornersUnlocked.Right
    Else
        minX = m_CornersUnlocked.Right
        maxX = m_CornersUnlocked.Left
    End If
    If (m_CornersUnlocked.Top < m_CornersUnlocked.Bottom) Then
        minY = m_CornersUnlocked.Top
        maxY = m_CornersUnlocked.Bottom
    Else
        minY = m_CornersUnlocked.Bottom
        maxY = m_CornersUnlocked.Top
    End If

    'At present, mask creation is only applicable for certain transformable shapes (rectangles, ellipses, lines).  Other functions,
    ' like "Invert selection", rely on an already-created mask - so attempting to create a mask again will have undesirable behavior.
    ' As such, use caution when calling this function, as the existing mask will be completely erased.
    
    'Start by creating a blank mask (this will also erase any existing mask)
    If ((m_SelMask.GetDIBWidth <> m_parentPDImage.Width) Or (m_SelMask.GetDIBHeight <> m_parentPDImage.Height)) Then
        m_SelMask.CreateBlank m_parentPDImage.Width, m_parentPDImage.Height, 32, maskBackColor, maskBackByte
        m_SelMask.SetInitialAlphaPremultiplicationState True
    Else
        m_SelMask.ResetDIB maskBackByte
    End If
        
    'Border selections are automatically disabled if the border size exceeds the smallest dimension (width or height)
    ' of the image
    Dim minDimension As Long
    If (m_CornersLocked.Width < m_CornersLocked.Height) Then minDimension = m_CornersLocked.Width Else minDimension = m_CornersLocked.Height
            
    'When rendering rectangle and ellipse selections in BORDERED mode, we calculate some values in advance.  Basically, these shapes
    ' handle border selections by rendering a smaller version of the selection shape within itself.  We never want the smaller copy
    ' of the shape to have negative width or height, so we check its dimensions in advance to make sure the border value is acceptable.
    ' If it isn't, we'll simply ignore it, and render the shape as a solid.
    Dim ignoreBorderValue As Boolean
    Dim rectModifier As Long
    
    ignoreBorderValue = False
    
    If (GetSelectionProperty_Long(sp_Area) = sa_Border) And ((m_SelectionShape = ss_Rectangle) Or (m_SelectionShape = ss_Circle)) Then
        rectModifier = GetSelectionProperty_Long(sp_BorderWidth)
        If (rectModifier < 1) Then rectModifier = 1
        If (rectModifier > (minDimension \ 2)) Then ignoreBorderValue = True
    End If
                
    'The actual rendering of the selection will vary based on the current selection type (obviously).
    ' TODO: move all rendering into a separate child class.  This is necessary for union/intersect/etc with multiple selections.
    Select Case m_SelectionShape
    
        Case ss_Rectangle
        
            'RECTANGLE SELECTION, NO ROUNDED CORNERS
            ' (Note: rectangular selections ignore the current antialiasing setting, as there's no point)
            If (Me.GetSelectionProperty_Long(sp_RoundedCornerRadius) = 0) Then
            
                'Interior/exterior selections, and bordered selections with too large of a width use identical rendering code.
                If (GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Or ignoreBorderValue Then
                
                    GDI_Plus.GDIPlusFillDIBRect m_SelMask, m_CornersLocked.Left, m_CornersLocked.Top, m_CornersLocked.Width, m_CornersLocked.Height, maskForeColor
                    
                'Border selections are a bit different; we must draw two rectangles - an exterior and an interior one.
                Else
                
                    'Draw the exterior rectangle
                    GDI_Plus.GDIPlusFillDIBRect m_SelMask, m_CornersLocked.Left, m_CornersLocked.Top, m_CornersLocked.Width, m_CornersLocked.Height, maskForeColor
                    
                    'Next, draw the interior one
                    GDI_Plus.GDIPlusFillDIBRect m_SelMask, m_CornersLocked.Left + rectModifier, m_CornersLocked.Top + rectModifier, m_CornersLocked.Width - rectModifier * 2, m_CornersLocked.Height - rectModifier * 2, maskBackColor
                    
                End If
            
            'RECTANGLE SELECTION *WITH* ROUNDED CORNERS
            Else
                
                'Interior/exterior selections use identical rendering code.  The only difference is the render color (whether the
                ' rounded rect represents the selected or non-selected region, basically).
                If (GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Or ignoreBorderValue Then
                    
                    'Use GDI+ for antialiasing, regular GDI for non-antialiased
                    GDIPlusDrawRoundRect m_SelMask, m_CornersLocked.Left, m_CornersLocked.Top, m_CornersLocked.Width, m_CornersLocked.Height, GetSelectionProperty_Long(sp_RoundedCornerRadius), maskForeColor, (GetSelectionProperty_Long(sp_Smoothing) > ss_None)
                    
                'Border selections are a bit different; we must draw two rounded rectangles - an exterior and an interior one.
                Else
                    
                    'Draw the exterior rectangle
                    GDIPlusDrawRoundRect m_SelMask, m_CornersLocked.Left, m_CornersLocked.Top, m_CornersLocked.Width, m_CornersLocked.Height, GetSelectionProperty_Long(sp_RoundedCornerRadius), maskForeColor, (GetSelectionProperty_Long(sp_Smoothing) > ss_None)
                    
                    'Next, draw the interior one.
                    GDIPlusDrawRoundRect m_SelMask, m_CornersLocked.Left + rectModifier, m_CornersLocked.Top + rectModifier, m_CornersLocked.Width - rectModifier * 2, m_CornersLocked.Height - rectModifier * 2, GetSelectionProperty_Long(sp_RoundedCornerRadius), maskBackColor, (GetSelectionProperty_Long(sp_Smoothing) > ss_None)
                    
                End If
                
            End If
                    
        'CIRCLES / ELLIPSES
        Case ss_Circle
        
            'Interior/exterior selections use identical rendering code.
            If (GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Or ignoreBorderValue Then
            
                GDIPlusFillEllipseToDC m_SelMask.GetDIBDC, m_CornersLocked.Left, m_CornersLocked.Top, m_CornersLocked.Width - 1, m_CornersLocked.Height - 1, maskForeColor, (GetSelectionProperty_Long(sp_Smoothing) > ss_None)
                
            'Border selections are a bit different; we must draw two ellipses - an exterior and an interior one.
            Else
            
                'Draw the exterior oval
                GDIPlusFillEllipseToDC m_SelMask.GetDIBDC, m_CornersLocked.Left, m_CornersLocked.Top, m_CornersLocked.Width, m_CornersLocked.Height, maskForeColor, (GetSelectionProperty_Long(sp_Smoothing) > ss_None)
                
                'Next, draw the interior one
                GDIPlusFillEllipseToDC m_SelMask.GetDIBDC, m_CornersLocked.Left + rectModifier, m_CornersLocked.Top + rectModifier, m_CornersLocked.Width - rectModifier * 2, m_CornersLocked.Height - rectModifier * 2, maskBackColor, (GetSelectionProperty_Long(sp_Smoothing) > ss_None)
                        
            End If
            
        
        'Line selections
        ' TODO: let the user specify line caps
        Case ss_Line
        
            'Interior/exterior line selections use identical rendering code.
            If (GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Or ignoreBorderValue Then
            
                GDIPlusDrawLineToDC m_SelMask.GetDIBDC, m_CornersUnlocked.Left, m_CornersUnlocked.Top, m_CornersUnlocked.Right, m_CornersUnlocked.Bottom, maskForeColor, 255, GetSelectionProperty_Long(sp_LineWidth), (GetSelectionProperty_Long(sp_Smoothing) > ss_None), GP_LC_Round
                
            'Border selections are a bit different; we must draw two lines - an interior and an exterior one.
            ' TODO: GDI+ supplies a function that generates a new path, using the outline of an existing path as drawn with a specified pen.
            '        That may be a more convenient way to render bordered line selections.  (GdipWidenPath FYI)
            Else
                
                'Draw the exterior line
                GDIPlusDrawLineToDC m_SelMask.GetDIBDC, m_CornersUnlocked.Left, m_CornersUnlocked.Top, m_CornersUnlocked.Right, m_CornersUnlocked.Bottom, maskForeColor, 255, GetSelectionProperty_Long(sp_LineWidth), (GetSelectionProperty_Long(sp_Smoothing) > ss_None), GP_LC_Round
                
                'Next, draw the interior one.  Note that some additional calculations are required to
                Dim xDiff As Double, yDiff As Double
                xDiff = Abs(m_CornersUnlocked.Left - m_CornersUnlocked.Right)
                yDiff = Abs(m_CornersUnlocked.Top - m_CornersUnlocked.Bottom)
                Dim xModifier As Long, yModifier As Long
                If (xDiff > yDiff) Then
                    If (xDiff = 0) Then xDiff = 0.00001
                    xModifier = GetSelectionProperty_Long(sp_BorderWidth)
                    yModifier = GetSelectionProperty_Long(sp_BorderWidth) * (yDiff / xDiff)
                Else
                    If (yDiff = 0) Then yDiff = 0.00001
                    xModifier = GetSelectionProperty_Long(sp_BorderWidth) * (xDiff / yDiff)
                    yModifier = GetSelectionProperty_Long(sp_BorderWidth)
                End If
                
                Dim newLineWidth As Double
                newLineWidth = GetSelectionProperty_Long(sp_LineWidth) - GetSelectionProperty_Long(sp_BorderWidth)
                
                If (newLineWidth > 0) Then
                    If (m_CornersUnlocked.Left = minX) Then
                        If (m_CornersUnlocked.Top = minY) Then
                            GDIPlusDrawLineToDC m_SelMask.GetDIBDC, m_CornersUnlocked.Left + xModifier, m_CornersUnlocked.Top + yModifier, m_CornersUnlocked.Right - xModifier, m_CornersUnlocked.Bottom - yModifier, maskBackColor, 255, newLineWidth, (GetSelectionProperty_Long(sp_Smoothing) > ss_None), GP_LC_Round
                        Else
                            GDIPlusDrawLineToDC m_SelMask.GetDIBDC, m_CornersUnlocked.Left + xModifier, m_CornersUnlocked.Top - yModifier, m_CornersUnlocked.Right - xModifier, m_CornersUnlocked.Bottom + yModifier, maskBackColor, 255, newLineWidth, (GetSelectionProperty_Long(sp_Smoothing) > ss_None), GP_LC_Round
                        End If
                    Else
                        If (m_CornersUnlocked.Top = minY) Then
                            GDIPlusDrawLineToDC m_SelMask.GetDIBDC, m_CornersUnlocked.Left - xModifier, m_CornersUnlocked.Top + yModifier, m_CornersUnlocked.Right + xModifier, m_CornersUnlocked.Bottom - yModifier, maskBackColor, 255, newLineWidth, (GetSelectionProperty_Long(sp_Smoothing) > ss_None), GP_LC_Round
                        Else
                            GDIPlusDrawLineToDC m_SelMask.GetDIBDC, m_CornersUnlocked.Left - xModifier, m_CornersUnlocked.Top - yModifier, m_CornersUnlocked.Right + xModifier, m_CornersUnlocked.Bottom + yModifier, maskBackColor, 255, newLineWidth, (GetSelectionProperty_Long(sp_Smoothing) > ss_None), GP_LC_Round
                        End If
                    End If
                End If
                
            End If
        
        'Polygon selections are easy - simply close the current point collection, then fill (or stroke) it via GDI+.
        Case ss_Polygon
        
            'Interior/exterior lasso selections use identical rendering code.
            If (GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Then
                GDIPlusDrawFilledShapeToDC m_SelMask.GetDIBDC, m_NumOfPolygonPoints, VarPtr(m_PolygonPoints(0)), maskForeColor, 255, (GetSelectionProperty_Long(sp_Smoothing) > ss_None), True, GetSelectionProperty_Float(sp_PolygonCurvature), GP_FM_Winding
                
            'Border selections stroke the path instead of filling it
            Else
                GDI_Plus.GDIPlusStrokePathToDC m_SelMask.GetDIBDC, m_NumOfPolygonPoints, VarPtr(m_PolygonPoints(0)), True, maskForeColor, 255, (GetSelectionProperty_Long(sp_Smoothing) > ss_None), GetSelectionProperty_Long(sp_BorderWidth), GP_LC_Flat, True, GetSelectionProperty_Float(sp_PolygonCurvature)
            End If
        
        'Strangely enough, lasso selections are arguably the simplest selection type to render, as we simply close the lasso shape,
        ' then fill it as if it represents an arbitrary region.
        Case ss_Lasso
        
            'Interior/exterior lasso selections use identical rendering code.
            If (GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Then
                GDIPlusDrawFilledShapeToDC m_SelMask.GetDIBDC, m_NumOfLassoPoints, VarPtr(m_LassoPoints(0)), maskForeColor, 255, (GetSelectionProperty_Long(sp_Smoothing) > ss_None), True, GetSelectionProperty_Float(sp_SmoothStroke), GP_FM_Winding
                
            'Border selections stroke the path instead of filling it
            Else
                GDI_Plus.GDIPlusStrokePathToDC m_SelMask.GetDIBDC, m_NumOfLassoPoints, VarPtr(m_LassoPoints(0)), True, maskForeColor, 255, (GetSelectionProperty_Long(sp_Smoothing) > ss_None), GetSelectionProperty_Long(sp_BorderWidth), GP_LC_Flat, True, GetSelectionProperty_Float(sp_SmoothStroke)
            End If
        
        'Wand selections use a custom class to perform a flood fill
        Case ss_Wand
        
            'Create a flood fill class to help us process the wand
            Dim cFloodFill As pdFloodFill
            Set cFloodFill = New pdFloodFill
            
            'Set all initial parameters
            cFloodFill.SetAntialiasingMode CBool(GetSelectionProperty_Long(sp_Smoothing) > ss_None)
            cFloodFill.SetInitialPoint m_CornersUnlocked.Left, m_CornersUnlocked.Top
            cFloodFill.SetTolerance GetSelectionProperty_Float(sp_WandTolerance)
            cFloodFill.SetSearchMode GetSelectionProperty_Long(sp_WandSearchMode)
            cFloodFill.SetCompareMode GetSelectionProperty_Long(sp_WandCompareMethod)
            
            'Based on the flood fill type (layer vs image), pass a different source layer to the flood fill class
            Dim tmpLayer As pdLayer
            Set tmpLayer = New pdLayer
            
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            
            If (GetSelectionProperty_Long(sp_WandSampleMerged) = 0) Then
                m_parentPDImage.GetCompositedImage tmpDIB
            Else
                tmpLayer.CopyExistingLayer m_parentPDImage.GetActiveLayer
                tmpLayer.ConvertToNullPaddedLayer m_parentPDImage.Width, m_parentPDImage.Height
                Set tmpDIB = tmpLayer.layerDIB
            End If
            
            'Apply the flood fill
            cFloodFill.InitiateFloodFill tmpDIB, m_SelMask
            
            'Erase our temporary objects
            Set tmpDIB = Nothing
            Set tmpLayer = Nothing
                        
        'Other selection types will be added in the future
        Case Else
    
    End Select
    
    'Mark the mask as ready for use
    m_IsMaskReady = True
    
    'We now need to establish a bounding region for the selection.  For certain types of selections, we can do this with existing knowledge
    ' (e.g. the m_CornersLocked RectF may reflect this).  For other types of selections, we need to find bounds via pixel searching.
    '
    'Why not just use the m_CornersLocked RectF?  The bounding rect may differ from those values if feathering is in use.
    ' The viewport renderer uses the actual bounding rect to optimize its rendering of the selection effect, so it needs values that
    ' incorporate the full affected area, including any feathering or other modifications.
    
    'As a rule, exterior selections enclose the entire image boundary, so we sort by getSelectionProperty_Long(SP_AREA) first
    Select Case GetSelectionProperty_Long(sp_Area)
    
        'Exterior selections typically bound the entire image.  We could search for a smaller area, but at present the costs
        ' of this outweigh any potential benefits.
        Case sa_Exterior
            With m_Bounds
                .Left = 0
                .Top = 0
                .Width = m_parentPDImage.Width
                .Height = m_parentPDImage.Height
            End With
        
        'Interior and bordered selections are handled more normally
        Case Else
        
            'Next, we sort bound calculations by selection shape
            Select Case m_SelectionShape
            
                'Rectangle, ellipse, and lasso selections are easy; bounds have already been set by the width/height values
                Case ss_Rectangle, ss_Circle
                    If (GetSelectionProperty_Long(sp_Smoothing) = ss_FullyFeathered) Then
                        With m_Bounds
                            .Left = m_CornersLocked.Left - GetSelectionProperty_Long(sp_FeatheringRadius)
                            .Top = m_CornersLocked.Top - GetSelectionProperty_Long(sp_FeatheringRadius)
                            .Width = m_CornersLocked.Width + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                            .Height = m_CornersLocked.Height + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                        End With
                    Else
                        With m_Bounds
                            .Left = m_CornersLocked.Left
                            .Top = m_CornersLocked.Top
                            .Width = m_CornersLocked.Width
                            .Height = m_CornersLocked.Height
                        End With
                    End If
                    
                'Lines are a bit weirder.  Because the two points are not sorted by distance (e.g. x1 may be greater than x2), we have to check that now.
                Case ss_Line
                
                    If (GetSelectionProperty_Long(sp_Smoothing) = ss_FullyFeathered) Then
                        With m_Bounds
                            .Left = minX - (GetSelectionProperty_Long(sp_LineWidth) / 2) - GetSelectionProperty_Long(sp_FeatheringRadius)
                            .Top = minY - (GetSelectionProperty_Long(sp_LineWidth) / 2) - GetSelectionProperty_Long(sp_FeatheringRadius)
                            .Width = (maxX - minX) + GetSelectionProperty_Long(sp_LineWidth) + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                            .Height = (maxY - minY) + GetSelectionProperty_Long(sp_LineWidth) + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                        End With
                    Else
                        With m_Bounds
                            .Left = minX - (GetSelectionProperty_Long(sp_LineWidth) / 2)
                            .Top = minY - (GetSelectionProperty_Long(sp_LineWidth) / 2)
                            .Width = (maxX - minX) + GetSelectionProperty_Long(sp_LineWidth)
                            .Height = (maxY - minY) + GetSelectionProperty_Long(sp_LineWidth)
                        End With
                    End If
                
                'Polygon selections support curvature via cardinal spline.  If active, we must manually calculate bounds,
                ' because the spline may extend beyond the current coordinate collection.
                Case ss_Polygon
                
                    'GDI+ can manually calculate bounds for us.  These bounds are not guaranteed to be the smallest possible area,
                    ' but they are guaranteed to enclose the entire path.
                    Me.FillBoundsUsingGDIPlus
                    
                    'If feathering is active, we must manually increase the bounding area to account for the feathering radius.
                    If (GetSelectionProperty_Long(sp_Smoothing) = ss_FullyFeathered) Then
                    
                        'Because of the way GDI+ pens work, they will be centered on the stroked path.  This means that we must
                        ' factor the pen width into our boundary calculations.
                        If (GetSelectionProperty_Long(sp_Area) = sa_Border) Then
                            With m_Bounds
                                .Left = m_CornersLocked.Left - GetSelectionProperty_Long(sp_FeatheringRadius) - (GetSelectionProperty_Long(sp_BorderWidth) \ 2)
                                .Top = m_CornersLocked.Top - GetSelectionProperty_Long(sp_FeatheringRadius) - (GetSelectionProperty_Long(sp_BorderWidth) \ 2)
                                .Width = m_CornersLocked.Width + GetSelectionProperty_Long(sp_FeatheringRadius) * 2 + GetSelectionProperty_Long(sp_BorderWidth)
                                .Height = m_CornersLocked.Height + GetSelectionProperty_Long(sp_FeatheringRadius) * 2 + GetSelectionProperty_Long(sp_BorderWidth)
                            End With
                        Else
                            With m_Bounds
                                .Left = m_CornersLocked.Left - GetSelectionProperty_Long(sp_FeatheringRadius)
                                .Top = m_CornersLocked.Top - GetSelectionProperty_Long(sp_FeatheringRadius)
                                .Width = m_CornersLocked.Width + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                                .Height = m_CornersLocked.Height + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                            End With
                        End If
                        
                    Else
                        With m_Bounds
                            .Left = m_CornersLocked.Left
                            .Top = m_CornersLocked.Top
                            .Width = m_CornersLocked.Width
                            .Height = m_CornersLocked.Height
                        End With
                    End If
                    
                    
                'Lasso selections support artificial smoothing via cardinal spline.  If active, we must manually calculate bounds,
                ' because the spline may extend beyond the drawn coordinates.
                Case ss_Lasso
                
                    'GDI+ can manually calculate bounds for us.  These bounds are not guaranteed to be the smallest possible area,
                    ' but they are guaranteed to enclose the entire path.
                    Me.FillBoundsUsingGDIPlus
                    
                    'If feathering is active, we must manually increase the bounding area to account for the feathering radius.
                    If (GetSelectionProperty_Long(sp_Smoothing) = ss_FullyFeathered) Then
                    
                        'Because of the way GDI+ pens work, they will be centered on the stroked path.  This means that we must
                        ' factor the pen width into our boundary calculations.
                        If GetSelectionProperty_Long(sp_Area) = sa_Border Then
                            With m_Bounds
                                .Left = m_CornersLocked.Left - GetSelectionProperty_Long(sp_FeatheringRadius) - (GetSelectionProperty_Long(sp_BorderWidth) \ 2)
                                .Top = m_CornersLocked.Top - GetSelectionProperty_Long(sp_FeatheringRadius) - (GetSelectionProperty_Long(sp_BorderWidth) \ 2)
                                .Width = m_CornersLocked.Width + GetSelectionProperty_Long(sp_FeatheringRadius) * 2 + GetSelectionProperty_Long(sp_BorderWidth)
                                .Height = m_CornersLocked.Height + GetSelectionProperty_Long(sp_FeatheringRadius) * 2 + GetSelectionProperty_Long(sp_BorderWidth)
                            End With
                        Else
                            With m_Bounds
                                .Left = m_CornersLocked.Left - GetSelectionProperty_Long(sp_FeatheringRadius)
                                .Top = m_CornersLocked.Top - GetSelectionProperty_Long(sp_FeatheringRadius)
                                .Width = m_CornersLocked.Width + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                                .Height = m_CornersLocked.Height + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                            End With
                        End If
                        
                    Else
                        With m_Bounds
                            .Left = m_CornersLocked.Left
                            .Top = m_CornersLocked.Top
                            .Width = m_CornersLocked.Width
                            .Height = m_CornersLocked.Height
                        End With
                    End If
                
                'Magic wand requires manual bounds-finding
                Case ss_Wand
                    
                    Me.FindNewBoundsManually True
                    
                    'If feathering is active, increase bounds accordingly
                    If (GetSelectionProperty_Long(sp_Smoothing) = ss_FullyFeathered) Then
                        With m_Bounds
                            .Left = m_CornersLocked.Left - GetSelectionProperty_Long(sp_FeatheringRadius)
                            .Top = m_CornersLocked.Top - GetSelectionProperty_Long(sp_FeatheringRadius)
                            .Width = m_CornersLocked.Width + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                            .Height = m_CornersLocked.Height + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                        End With
                    End If
                
                'Other shapes currently rely on manual bounds-checking, using the rendered mask as the guide
                Case Else
                    Me.FindNewBoundsManually
                    
            End Select
                
    End Select
    
    'Do some basic bounds checking on the bound values to make sure they lie inside the image.  This is important because the
    ' selection mask (and any code that operates on it) assumes a match to image boundaries, despite the fact that selection
    ' points can actually lie anywhere on the canvas - even outside the image!
    FixBoundsToImageSize
    
    'Finally, if the selection is locked and feathering has been requested, apply it now.
    ' (We only apply feathering when locked-in, as the performance penalty is severe.)
    If m_IsLocked And (GetSelectionProperty_Long(sp_Smoothing) = ss_FullyFeathered) And (GetSelectionProperty_Long(sp_FeatheringRadius) > 0) Then ApplyFeatheringToMask
    
End Sub

'When calculating selection boundaries, all bounds must ultimately lie on or inside image borders.  Use this function to verify that.
' (Because bounds have already been precisely calculated, this function's behavior does not need to differ by selection type.)
Private Sub FixBoundsToImageSize()
    
    With m_Bounds
        
        If (.Left < 0) Then
            .Width = .Width + .Left
            .Left = 0
        End If
        
        If (.Top < 0) Then
            .Height = .Height + .Top
            .Top = 0
        End If
        
        If (.Left + .Width > m_parentPDImage.Width) Then .Width = m_parentPDImage.Width - .Left
        If (.Top + .Height > m_parentPDImage.Height) Then .Height = m_parentPDImage.Height - .Top
        
    End With
    
End Sub

'Apply feathering to the current selection mask.  If the user is on Win 7 or later, we may choose to do this via GDI+.
Private Sub ApplyFeatheringToMask()
    QuickBlurDIB m_SelMask, GetSelectionProperty_Long(sp_FeatheringRadius), False
End Sub

'While a selection tool is active, we draw transform nodes around it.  The viewport renderer invokes this function as necessary.
Friend Sub RenderTransformNodes(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, ByVal imgLeft As Long, ByVal imgTop As Long)
    
    'If an update has been requested, but we are already in the middle of one, deny subsequent requests
    If m_RejectRefreshRequests Then Exit Sub
    
    'Before drawing the nodes, we need to convert the selection's coordinates from "image coordinates" to "viewport coordinates".
    ' (Remember that the selection's location is stored as coordinates relative to the IMAGE ITSELF - but to render the transform nodes
    '   on the screen, we need to translate the coordinates to the current viewport - this makes them display accurately when scroll
    '   and zoom parameters are in effect.)
    Dim dstLeft As Double, dstRight As Double
    Dim dstTop As Double, dstBottom As Double
    
    If (m_SelectionShape = ss_Rectangle) Or (m_SelectionShape = ss_Circle) Then
        Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), m_CornersLocked.Left, m_CornersLocked.Top, dstLeft, dstTop
        Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), m_CornersLocked.Left + m_CornersLocked.Width, m_CornersLocked.Top + m_CornersLocked.Height, dstRight, dstBottom
    Else
        Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), m_CornersUnlocked.Left, m_CornersUnlocked.Top, dstLeft, dstTop
        Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), m_CornersUnlocked.Right, m_CornersUnlocked.Bottom, dstRight, dstBottom
    End If
    
    
    'If this selection is in vector format, we have the option of drawing transform circles at appropriate locations
    ' for the selection's shape.
    If (m_SelectionShape <> ss_Raster) Then
        
        Dim circRadius As Long
        circRadius = 7
                
        Dim circAlpha As Long
        circAlpha = 190
        
        Dim i As Long
        
        Select Case m_SelectionShape
        
            Case ss_Rectangle, ss_Circle
        
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
                
            Case ss_Line
            
                'Lines only need transform nodes at the end points of the line
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstLeft, dstTop, circRadius, circAlpha
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstRight, dstBottom, circRadius, circAlpha
            
            Case ss_Polygon
            
                For i = 0 To m_NumOfPolygonPoints - 1
                    Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), m_PolygonPoints(i).x, m_PolygonPoints(i).y, dstLeft, dstTop
                    GDIPlusDrawCanvasCircle dstCanvas.hDC, dstLeft, dstTop, circRadius, circAlpha
                Next i
                
            Case ss_Wand
                GDIPlusDrawCanvasCircle dstCanvas.hDC, dstLeft, dstTop, circRadius, circAlpha
                
            'Other selection types will require more involved transform code
            Case Else
                
        End Select
        
        dstCanvas.RequestViewportRedraw
                
    End If
    
End Sub

'Render the current selection mask using one of several methods.  Note that this code operates entirely independent of selection type.
' It only requires that valid selection boundaries be specified in advance, which it uses to accelerate the render process.
' (Also note that the renderColor default constant evaluates to RGB(112, 183, 255), but VB doesn't let us use RGB inside a function dec.)
' TODO: allow light box translucency to be set by the user
Friend Sub RenderCustom(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, ByVal imgLeft As Long, ByVal imgTop As Long, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal renderType As PD_SelectionRender, Optional ByVal renderColor As Long = 16758640)
    
    'If an update has been requested, but we are already in the middle of one, deny subsequent requests
    If m_RejectRefreshRequests Then Exit Sub

    'Is the selection mask ready to be used?  If not, prepare it now.
    If (Not m_IsMaskReady) Then CreateSelectionMask
    
    'A few different criteria can cause us to abandon this function in favor of renderOutlineOnly (which, per its name,
    ' renders the active selection as just an outline):
    ' 1) The user has specified outline-only mode, and this is not a raster or magic wand selection.  (Those don't support outline mode at present.)
    ' 2) A lasso selection is active but not locked-in (meaning it's still being drawn)
    ' 3) A polygon selection is active but not locked-in (meaning it's still being drawn)
    '
    'If any of the above are true, we will reroute control to renderOutlineOnly
    If ((renderType = PDSR_Outline) And (m_SelectionShape <> ss_Wand) And (m_SelectionShape <> ss_Raster)) Or _
        ((m_SelectionShape = ss_Lasso) And (Not m_LassoClosed)) Or _
        ((m_SelectionShape = ss_Polygon) And (Not m_PolygonClosed)) Then
        
        RenderOutlineOnly dstDIB, srcImage, dstCanvas
        Exit Sub
            
    End If
    
    'Wand and raster selections cannot be rendered in outline mode; if this happens, forcibly switch the mode to HIGHLIGHT.
    If (renderType = PDSR_Outline) And ((m_SelectionShape = ss_Wand) Or (m_SelectionShape = ss_Raster)) Then renderType = PDSR_Highlight
    
    
    'Before rendering the selection, we need to convert the its coordinates from "image coordinates" to "viewport coordinates".
    ' (Remember that the selection's location is stored as coordinates relative to the IMAGE ITSELF - but to render it on the screen,
    '   we need to translate the coordinates to the current viewport - this makes them display accurately when scroll and zoom are
    '   in effect.)
    Dim dstLeft As Double, dstRight As Double
    Dim dstTop As Double, dstBottom As Double
    Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, m_Bounds.Left, m_Bounds.Top, dstLeft, dstTop
    Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, m_Bounds.Left + m_Bounds.Width, m_Bounds.Top + m_Bounds.Height, dstRight, dstBottom
    
    'Several different selection rendering methods use the same variables.  I've declared the common ones here.
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    Dim x As Long, y As Long
    Dim xStride As Long
    
    Dim selMaskData() As Byte
    Dim selMaskSA As SAFEARRAY2D
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = m_parentPDImage.Width - 1
    srcHeight = m_parentPDImage.Height - 1
    
    'Just to be safe, declare some bounding variables in advance as well.  PhotoDemon now allows selections to lie off the image, so
    ' we must make sure we don't attempt to render selection pixels in any out-of-bound regions.
    Dim rightBound As Long, bottomBound As Long
    rightBound = imgLeft + imgWidth - 1
    bottomBound = imgTop + imgHeight - 1
            
    If (rightBound > srcImage.canvasBuffer.GetDIBWidth - 1) Then rightBound = srcImage.canvasBuffer.GetDIBWidth - 1
    If (bottomBound > srcImage.canvasBuffer.GetDIBHeight - 1) Then bottomBound = srcImage.canvasBuffer.GetDIBHeight - 1
    
    'Build look-up tables for both the x and y direction.  This is faster than constantly calculating new values for
    ' every pixel inside the loop.  Note also that this code is identical among various selection render methods.
    Dim newX As Double, newY As Double, newXInt As Long
    Dim xLookup() As Long, yLookup() As Long
    
    Dim xPxSizeInBytes As Long
    If (dstDIB.GetDIBColorDepth = 32) Then xPxSizeInBytes = 4 Else xPxSizeInBytes = 3
    
    ReDim xLookup(imgLeft To rightBound) As Long
    For x = imgLeft To rightBound
        Drawing.ConvertCanvasCoordsToImageCoords FormMain.mainCanvas(0), srcImage, x, imgTop, newX, newY, True
        newXInt = Int(newX)
        xLookup(x) = newXInt * xPxSizeInBytes
    Next x
    
    ReDim yLookup(imgTop To bottomBound) As Long
    For y = imgTop To bottomBound
        Drawing.ConvertCanvasCoordsToImageCoords FormMain.mainCanvas(0), srcImage, imgLeft, y, newX, newY, True
        yLookup(y) = Int(newY)
    Next y
    
    Dim tmpSelHolder As Byte
    Dim tmpSelDivider As Double
    
    'Start by acquiring a pointer to the target DIB's pixel bits
    PrepDIBSafeArray dstDIB, tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Next, acquire a pointer to the selection mask's pixel bits
    PrepDIBSafeArray m_SelMask, selMaskSA
    CopyMemory ByVal VarPtrArray(selMaskData()), VarPtr(selMaskSA), 4
    
    'Based on the selection type, draw a corresponding selection visualization.
    Select Case renderType
    
        'Lightbox style
        Case PDSR_Lightbox
            
            'Use a look-up table to dim the non-active parts of the image
            Dim hData(0 To 255) As Byte
            
            For x = 0 To 255
                hData(x) = x \ 3
            Next x
                                    
            'Dim the non-active parts of the image
            For x = imgLeft To rightBound
                xStride = x * xPxSizeInBytes
            For y = imgTop To bottomBound
            
                tmpSelHolder = selMaskData(xLookup(x), yLookup(y))
                
                'Remember that lightbox works a little weirdly - UNSELECTED pixels are marked, while selected ones are left untouched
                If tmpSelHolder = 0 Then
                    iData(xStride + 2, y) = hData(iData(xStride + 2, y))
                    iData(xStride + 1, y) = hData(iData(xStride + 1, y))
                    iData(xStride, y) = hData(iData(xStride, y))
                ElseIf tmpSelHolder < 255 Then
                    tmpSelDivider = tmpSelHolder / 255
                    iData(xStride + 2, y) = BlendColors(hData(iData(xStride + 2, y)), iData(xStride + 2, y), tmpSelDivider)
                    iData(xStride + 1, y) = BlendColors(hData(iData(xStride + 1, y)), iData(xStride + 1, y), tmpSelDivider)
                    iData(xStride, y) = BlendColors(hData(iData(xStride, y)), iData(xStride, y), tmpSelDivider)
                End If
                
            Next y
            Next x
            
        'Highlight style (Paint.NET)
        Case PDSR_Highlight
        
            'Extract RGB values from the passed color
            Dim newR As Long, newG As Long, newB As Long
            newR = Colors.ExtractRed(renderColor)
            newG = Colors.ExtractGreen(renderColor)
            newB = Colors.ExtractBlue(renderColor)
            
            'Use look-up tables to highlight the active parts of the image
            Dim rData(0 To 255) As Byte, gData(0 To 255) As Byte, bData(0 To 255) As Byte
            
            For x = 0 To 255
                rData(x) = (newR + x) \ 2
                gData(x) = (newG + x) \ 2
                bData(x) = (newB + x) \ 2
            Next x
                        
            'Highlight the selected parts of the image
            For x = imgLeft To rightBound
                xStride = x * xPxSizeInBytes
            For y = imgTop To bottomBound
            
                tmpSelHolder = selMaskData(xLookup(x), yLookup(y))
                
                'If this pixels is over an active pixel in the selection mask, mark it with the selected color
                If tmpSelHolder = 255 Then
                    iData(xStride + 2, y) = rData(iData(xStride + 2, y))
                    iData(xStride + 1, y) = gData(iData(xStride + 1, y))
                    iData(xStride, y) = bData(iData(xStride, y))
                    
                'Partially selected pixels are rendered with antialiasing
                ElseIf tmpSelHolder > 0 Then
                    tmpSelDivider = tmpSelHolder / 255
                    iData(xStride + 2, y) = BlendColors(iData(xStride + 2, y), rData(iData(xStride + 2, y)), tmpSelDivider)
                    iData(xStride + 1, y) = BlendColors(iData(xStride + 1, y), gData(iData(xStride + 1, y)), tmpSelDivider)
                    iData(xStride, y) = BlendColors(iData(xStride, y), bData(iData(xStride, y)), tmpSelDivider)
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
Private Sub RenderOutlineOnly(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas)
    
    'To render an outline correctly, we must first translate all relevant points from the image coordinate space
    ' to the viewport coordinate space.  The number of points required obviously varies by selection shape.
    Dim tmpViewportSpace() As POINTFLOAT
    Dim numOfRenderPoints As Long

    'Simple shapes (rectangles, ellipses), will use these dstLeft/Right/Top/Bottom values instead of a custom point array
    Dim dstLeft As Double, dstRight As Double
    Dim dstTop As Double, dstBottom As Double
    
    If (m_SelectionShape = ss_Rectangle) Or (m_SelectionShape = ss_Circle) Then
        Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, m_CornersLocked.Left, m_CornersLocked.Top, dstLeft, dstTop
        Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, m_CornersLocked.Left + m_CornersLocked.Width, m_CornersLocked.Top + m_CornersLocked.Height, dstRight, dstBottom
    ElseIf (m_SelectionShape = ss_Line) Then
        Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, m_CornersUnlocked.Left, m_CornersUnlocked.Top, dstLeft, dstTop
        Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), srcImage, m_CornersUnlocked.Right, m_CornersUnlocked.Bottom, dstRight, dstBottom
    End If
    
    'New in version 6.6 is an awesome new path class, which makes the rendering steps much easier than they have historically been.
    Dim pdPath As pd2DPath
    Set pdPath = New pd2DPath
    
    Select Case m_SelectionShape
    
        Case ss_Rectangle
            
            'Add the selection rectangle to the path
            If (GetSelectionProperty_Long(sp_RoundedCornerRadius) > 0) Then
                pdPath.AddRoundedRectangle_Relative dstLeft, dstTop, dstRight - dstLeft, dstBottom - dstTop, GetSelectionProperty_Long(sp_RoundedCornerRadius) * g_Zoom.GetZoomValue(m_parentPDImage.GetZoom)
            Else
                pdPath.AddRectangle_Absolute dstLeft, dstTop, dstRight, dstBottom
            End If
            
            'If the selection is bordered, widen it now
            'If (getSelectionProperty_Long(SP_AREA) = sa_Border) And (Not getSelectionProperty_Boolean(sp_BorderWidth_IGNORE)) Then
            '    pdPath.ConvertPath_PenTrace getSelectionProperty_Long(sp_BorderWidth)
            'End If
                        
        Case ss_Circle
        
            'Add the ellipse to the path
            pdPath.AddEllipse_Absolute dstLeft, dstTop, dstRight, dstBottom
                    
        Case ss_Line
        
            'Add the line to the path
            pdPath.AddLine dstLeft, dstTop, dstRight, dstBottom
            
            'Widen the path to encompass only the border of the line
            Dim penPathBorder As pd2DPen
            Drawing2D.QuickCreateSolidPen penPathBorder, GetSelectionProperty_Long(sp_LineWidth) * g_Zoom.GetZoomValue(m_parentPDImage.GetZoom), , , P2_LJ_Round, P2_LC_Round
            pdPath.ConvertPath_PenTrace penPathBorder
            
        'Polygon and lasso selections use roughly identical outline code; the only difference is the reference array used and the number of points
        Case ss_Polygon, ss_Lasso
        
            If (m_SelectionShape = ss_Polygon) Then numOfRenderPoints = m_NumOfPolygonPoints Else numOfRenderPoints = m_NumOfLassoPoints
            
            If (numOfRenderPoints > 0) Then
            
                'Convert the lasso or polygon array to viewport coordinate space
                ReDim tmpViewportSpace(0 To numOfRenderPoints - 1) As POINTFLOAT
                Dim tmpX As Double, tmpY As Double
                
                Dim i As Long
                
                For i = 0 To numOfRenderPoints - 1
                    
                    If (m_SelectionShape = ss_Polygon) Then
                        Drawing.ConvertImageCoordsToCanvasCoords dstCanvas, srcImage, m_PolygonPoints(i).x, m_PolygonPoints(i).y, tmpX, tmpY
                    Else
                        Drawing.ConvertImageCoordsToCanvasCoords dstCanvas, srcImage, m_LassoPoints(i).x, m_LassoPoints(i).y, tmpX, tmpY
                    End If
                    
                    tmpViewportSpace(i).x = tmpX
                    tmpViewportSpace(i).y = tmpY
                Next i
                
                'Add the converted shape to the path object
                If (m_SelectionShape = ss_Polygon) Then
                    pdPath.AddPolygon numOfRenderPoints, VarPtr(tmpViewportSpace(0)), m_PolygonClosed, True, GetSelectionProperty_Float(sp_PolygonCurvature)
                Else
                    pdPath.AddPolygon numOfRenderPoints, VarPtr(tmpViewportSpace(0)), m_LassoClosed, True, GetSelectionProperty_Float(sp_SmoothStroke)
                End If
                
            End If
    
    End Select
    
    'Draw the assembled path onto the canvas
    Dim cPainter As pd2DPainter, cSurface As pd2DSurface, cPenUIBase As pd2DPen, cPenUITop As pd2DPen
    Drawing2D.QuickCreatePainter cPainter
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstDIB.GetDIBDC, True
    Drawing2D.QuickCreatePairOfUIPens cPenUIBase, cPenUITop, , P2_LJ_Round, P2_LC_Round
    
    cPainter.DrawPath cSurface, cPenUIBase, pdPath
    cPainter.DrawPath cSurface, cPenUITop, pdPath
    
End Sub

Private Sub Class_Initialize()
    
    m_IsLocked = False
    
    'Initialize the selection mask for this object and mark it as "not ready" (because no mask has been drawn yet)
    Set m_SelMask = New pdDIB
    m_IsMaskReady = False
    
    'Note that a mask has never been created for this selection
    m_MaskHasBeenCreated = False
    
    'Mark it as not transformable... yet
    m_IsTransformable = False
    
    'Prepare the property dictionary
    Set m_PropertyDict = New pdDictionary
    
    'No lasso or polygon points yet
    m_NumOfLassoPoints = 0
    ReDim m_LassoPoints(0) As POINTFLOAT
    
    m_NumOfPolygonPoints = 0
    ReDim m_PolygonPoints(0) As POINTFLOAT
    
    'Polygon selections are (obviously) not yet closed
    m_PolygonClosed = False
        
End Sub

Private Sub Class_Terminate()
    
    If (Not m_SelMask Is Nothing) Then
        m_SelMask.EraseDIB
        Set m_SelMask = Nothing
    End If

End Sub

'Sometimes a selection needs to directly access the bits of a target DIB. This routine sets up the corresponding SafeArray.
Private Sub PrepDIBSafeArray(ByRef srcDIB As pdDIB, ByRef tmpSA As SAFEARRAY2D)
    With tmpSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcDIB.GetDIBHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcDIB.GetDIBStride
        .pvData = srcDIB.GetDIBPointer
    End With
End Sub

'Create a selection from selection data previously saved to file
' (NOTE: this function will not generate a selection mask or render the selection on-screen.  The calling function must explicitly
'        request a render if they want one.)
Friend Function ReadSelectionFromFile(ByVal srcFilename As String, Optional ignoreLockStatus As Boolean = False) As Boolean
    
    'Like all other PD-specific files, selection files are just pdPackage instances
    Dim cPackage As pdPackager
    Set cPackage = New pdPackager
    If cPackage.ReadPackageFromFile(srcFilename, SELECTION_IDENTIFIER, PD_SM_FileBacked) Then
    
        'Reset some of our internal trackers prior to reading the selection data
        
        'First, note that our mask is not yet ready.  (Masks are generated on-demand, unless this is a raster selection,
        ' in which case we'll load the actual mask straight from the file.)
        m_IsMaskReady = False
        
        Dim dibLoadedSuccessfully As Boolean
        dibLoadedSuccessfully = False
        
        'Retrieve the file-specific header from the package.  This contains details like selection format version, which we need
        ' before we proceed with full parsing.
        Dim minorHeader As String
        If cPackage.GetNodeDataByName_String("SelHeader", True, minorHeader) Then
            
            'Copy the string into an XML parser
            Dim headerXML As pdParamXML
            Set headerXML = New pdParamXML
            headerXML.SetParamString minorHeader
            
            'Verify selection version.  (At present, there's only one possible version.)
            If (headerXML.GetLong("SelVersion", 0) = SELECTION_FILE_VERSION_2017) Then
            
                'This is enough to validate the file.  Load the full selection header and initialize this object accordingly.
                Dim majorHeader As String
                If cPackage.GetNodeDataByName_String("SelHeader", False, majorHeader) Then
                    
                    Me.InitFromXML majorHeader
                    
                    'Use the minor header to populate a few extra settings
                    m_IsLocked = headerXML.GetBool("SelIsLocked", False)
                    m_IsTransformable = headerXML.GetBool("SelIsTransformable", True)
                    
                    'Raster selections (e.g. selections that can't be described with vectors) store a full copy of their selection mask
                    ' in the selection file.  Load it conditionally.
                    If (m_SelectionShape = ss_Raster) Then
                        
                        'Retrieve the original parent image's width and height.  If the original containing image had the same dimensions
                        ' as our current image (e.g. during Undo/Redo operations), we can create the selection mask directly from the
                        ' file data.  Otherwise, we have no choice but to perform a resize so that the old mask matches the new dimensions.
                        Dim origWidth As Long, origHeight As Long
                        origWidth = headerXML.GetLong("SelMaskWidth", m_parentPDImage.Width)
                        origHeight = headerXML.GetLong("SelMaskHeight", m_parentPDImage.Height)
                        
                        Dim tmpDIBPointer As Long, tmpDIBLength As Long
                        
                        'Dimensions match!  Load the source data directly into our selection mask
                        If (origWidth = m_parentPDImage.Width) And (origHeight = m_parentPDImage.Height) Then
                            
                            With headerXML
                                m_SelMask.CreateBlank origWidth, origHeight, .GetLong("SelMaskDepth", 32), 0, 0
                            End With
                            
                            m_SelMask.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                            m_SelMask.SetInitialAlphaPremultiplicationState headerXML.GetBool("SelMaskAlphaPremultiplied", True)
                            cPackage.GetNodeDataByName_UnsafeDstPointer "SelMask", False, tmpDIBPointer
                        
                        'Dimensions do not match.  Use an intermediary DIB to cache the original raster data, then resize it to match
                        ' our current image.
                        Else
                        
                            Dim tmpDIB As pdDIB
                            Set tmpDIB = New pdDIB
                            
                            With headerXML
                                tmpDIB.CreateBlank origWidth, origHeight, .GetLong("SelMaskDepth", 32), 0, 0
                            End With
                            
                            tmpDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                            tmpDIB.SetInitialAlphaPremultiplicationState headerXML.GetBool("SelMaskAlphaPremultiplied", True)
                            If cPackage.GetNodeDataByName_UnsafeDstPointer("SelMask", False, tmpDIBPointer) Then
                                m_SelMask.CreateFromExistingDIB tmpDIB, m_parentPDImage.Width, m_parentPDImage.Height, True
                            End If
                            
                        End If
                        
                        'Note that the mask has been created; this is crucial for successful raster loads
                        m_MaskHasBeenCreated = True
                        m_IsMaskReady = True
                        
                    End If
                    
                    'If this selection isn't a vector, we need to manually find its boundaries now
                    If ((m_SelectionShape = ss_Raster) And (Not m_SelMask Is Nothing)) Then Me.FindNewBoundsManually
                    
                    'Polygon selections loaded from file are assumed to be closed
                    m_PolygonClosed = True
                    
                    'If this selection was locked in at the time of its save, lock this selection in as well
                    If (Not ignoreLockStatus) Then
                        If m_IsLocked Then
                            Me.LockIn
                            m_parentPDImage.SetSelectionActive True
                        Else
                            Me.LockRelease
                            m_parentPDImage.SetSelectionActive False
                        End If
                    End If
                    
                    ReadSelectionFromFile = True
                    
                Else
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "WARNING!  pdSelection failed to retrieve the actual header string for this file."
                    #End If
                    ReadSelectionFromFile = False
                End If
            
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  pdSelection found an unknown version in this file header."
                #End If
                ReadSelectionFromFile = False
            End If
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  pdSelection failed to load a valid header from this saved selection file."
            #End If
            ReadSelectionFromFile = False
        End If
        
    End If
        
End Function

'As of v7.0, pdPackager is used to read/write selection files.  Note that this function will blindly overwrite the destination
' file if it exists; it's up to the caller to plan for this.
Friend Function WriteSelectionToFile(ByVal dstFilename As String, Optional ByVal compressXML As PD_COMPRESSION_ENGINES = PD_CE_Zstd, Optional ByVal xmlCompressionLevel As Long = -1, Optional ByVal compressRaster As PD_COMPRESSION_ENGINES = PD_CE_Zstd, Optional ByVal rasterCompressionLevel As Long = -1) As Boolean
    
    Dim cPackage As pdPackager
    Set cPackage = New pdPackager
    cPackage.PrepareNewPackage , SELECTION_IDENTIFIER, , PD_SM_MemoryBacked
    
    'We now prepare two different XML packets for the packager.  One is the default vector selection descriptor PD uses internally;
    ' the other is a set of file-specific descriptors that make it easier to recreate this selection at load-time.
    Dim nodeIndex As Long
    nodeIndex = cPackage.AddNode("SelHeader")
    
    If (xmlCompressionLevel = -1) Then xmlCompressionLevel = Compression.GetDefaultCompressionLevel(compressXML)
    If (rasterCompressionLevel = -1) Then rasterCompressionLevel = Compression.GetDefaultCompressionLevel(compressRaster)
    
    'Write the default vector selection descriptor first
    cPackage.AddNodeDataFromString nodeIndex, False, Me.GetSelectionAsXML(), compressXML, xmlCompressionLevel
    
    'Prep a secondary descriptor
    Dim maskWillBeEmbedded As Boolean
    maskWillBeEmbedded = ((m_SelectionShape = ss_Raster) And m_MaskHasBeenCreated)
    Dim tmpXML As pdParamXML
    Set tmpXML = New pdParamXML
    With tmpXML
        .AddParam "SelVersion", SELECTION_FILE_VERSION_2017
        .AddParam "SelParentImageWidth", m_parentPDImage.Width
        .AddParam "SelParentImageHeight", m_parentPDImage.Height
        .AddParam "SelIsLocked", Me.IsLockedIn
        .AddParam "SelIsTransformable", m_IsTransformable
        .AddParam "SelMaskEmbedded", maskWillBeEmbedded
        
        'If we're going to embed a mask, we also need to store some mask-specific data
        If maskWillBeEmbedded Then
            .AddParam "SelMaskDepth", m_SelMask.GetDIBColorDepth
            .AddParam "SelMaskWidth", m_SelMask.GetDIBWidth
            .AddParam "SelMaskHeight", m_SelMask.GetDIBHeight
            .AddParam "SelMaskStride", m_SelMask.GetDIBStride
            .AddParam "SelMaskAlphaPremultiplied", m_SelMask.GetAlphaPremultiplication
        End If
        
    End With
    
    'Add the secondary descriptor to the same node (in the header chunk)
    cPackage.AddNodeDataFromString nodeIndex, True, tmpXML.GetParamString(), compressXML, xmlCompressionLevel
    
    'Finally, if this is a raster-type selection, add the full mask as well
    If maskWillBeEmbedded Then
        Dim maskDIBPointer As Long, maskDIBLength As Long
        m_SelMask.RetrieveDIBPointerAndSize maskDIBPointer, maskDIBLength
        nodeIndex = cPackage.AddNode("SelMask")
        cPackage.AddNodeDataFromPointer nodeIndex, False, maskDIBPointer, maskDIBLength, compressRaster, rasterCompressionLevel
    End If
    
    WriteSelectionToFile = cPackage.WritePackageToFile(dstFilename)
    
End Function

'When working with raster selections (e.g. non-transformable ones), we still want to minimize selection processing time by processing the
' smallest possible rectangle that includes all selected pixels.  This function will scan the selection mask and populate the
' m_CornersLocked.Left/Top/Width/Height and m_Bounds.Left/Top/Width/Height values automatically, based on the mask's contents.
'
'By default, this function will mark a selection as type "Raster", since it's assumed that a vector selection would already have the data
' necessary to determine its own bounds.  This behavior can be overridden, but do so with caution, because incorrectly calculated bounds
' will cause errors on tools that try to map the selection back to the active layer.
'
'If boundaries are successfully found, this function will return TRUE.  Otherwise, it will return FALSE, which effectively means the
' selection mask is blank.
'
'TODO: only check the blue channel, as the other channels are irrelevant
Friend Function FindNewBoundsManually(Optional ByVal overrideRasterState As Boolean = False) As Boolean

    'Message "Finding new selection boundaries..."

    'This function assumes two things: that a selection is not transformable (otherwise we'd know the boundaries already), and
    ' a mask has already been created.  If either of these two conditions is not met, this function may fail.
    '
    'Because these states can be inferred, this function will automatically set certain selection parameters.  This behavior can
    ' be overriden, BUT DO NOT OVERRIDE IT without understanding the consequences - in particular, that a mismatch between
    ' vector selection data and the selection mask will cause errors in the mask-to-layer mapping code.
    If (Not overrideRasterState) Then
    
        m_IsTransformable = False
        m_IsMaskReady = True
        
        'Because the selection is being converted to pure raster data, we must also update its shape
        m_SelectionShape = ss_Raster
        
    End If
    
    FindNewBoundsManually = True
    
    'Make sure the mask is ready for processing
    If (Not m_IsMaskReady) Then
        Debug.Print "Creating new selection mask at the request of FindNewBoundsManually()"
        CreateSelectionMask
    End If
    
    'Point a standard int array at the selection mask
    Dim x As Long, y As Long
    
    Dim selMaskData() As Long
    Dim selMaskSA As SAFEARRAY2D
    m_SelMask.WrapLongArrayAroundDIB selMaskData, selMaskSA
    
    Dim maskWidth As Long, maskHeight As Long
    maskWidth = m_SelMask.GetDIBWidth - 1
    maskHeight = m_SelMask.GetDIBHeight - 1
    
    Dim boundFound As Boolean
    
    'Find the top bound first.
    boundFound = False
    y = 0
    Do
    
        For x = 0 To maskWidth
            
            'In the future, if we decide to let the user select individual color channels, this code will need to be reworked.
            ' For now, however, all pixels are rendered as grayscale, so we can shortcut and just check for non-zero entries.
            If (selMaskData(x, y) <> 0) Then
                boundFound = True
                m_CornersLocked.Top = y
                m_Bounds.Top = m_CornersLocked.Top
                Exit For
            End If
            
        Next x
        
        'Blank masks shouldn't be possible, but if they are, mark the boundary as found to prevent a program lock
        If (Not boundFound) And (y >= maskHeight) Then
            Debug.Print "No top boundary found - is mask blank?"
            boundFound = True
            m_CornersLocked.Top = 0
            m_Bounds.Top = m_CornersLocked.Top
            FindNewBoundsManually = False
        End If
        
        y = y + 1
    
    Loop While (Not boundFound)
    
    'If the selection mask is empty, abandon ship
    If (Not FindNewBoundsManually) Then
        m_SelMask.UnwrapLongArrayFromDIB selMaskData
        Exit Function
    End If
    
    'Next, find the bottom bound.  Note that we skip the "is mask blank" check, as we've already handled that case above.
    boundFound = False
    y = maskHeight
    
    Do
    
        For x = 0 To maskWidth
            If (selMaskData(x, y) <> 0) Then
                boundFound = True
                m_CornersLocked.Height = y - m_CornersLocked.Top + 1
                m_Bounds.Height = m_CornersLocked.Height
                Exit For
            End If
        Next x
        
        y = y - 1
    
    Loop While (Not boundFound)
    
    'Next, find the left bound
    boundFound = False
    x = 0
    Do
    
        For y = 0 To maskHeight
            If (selMaskData(x, y) <> 0) Then
                boundFound = True
                m_CornersLocked.Left = x
                m_Bounds.Left = m_CornersLocked.Left
                Exit For
            End If
            
        Next y
        
        x = x + 1
    
    Loop While (Not boundFound)
    
    'Finally, find the right bound
    boundFound = False
    x = maskWidth
    
    Do
    
        For y = 0 To maskHeight
            If (selMaskData(x, y) <> 0) Then
                boundFound = True
                m_CornersLocked.Width = x - m_CornersLocked.Left + 1
                m_Bounds.Width = m_CornersLocked.Width
                Exit For
            End If
        Next y
        
        x = x - 1
        
    Loop While (Not boundFound)
    
    'All selection boundaries have now been located
    
    'Release our temporary byte array and exit
    m_SelMask.UnwrapLongArrayFromDIB selMaskData
    
End Function

'External functions can use this function to request a thumbnail version of the selection mask.
Friend Function RequestThumbnail(ByRef dstThumbnailDIB As pdDIB, Optional ByVal thumbnailSize As Long = 64) As Boolean
    
    If (m_SelMask Is Nothing) Then
        RequestThumbnail = False
        Exit Function
    End If
    
    'Thumbnails have some interesting requirements.  We always want them to be square, with the image set in the middle
    ' of the thumbnail (with aspect ratio preserved) and any empty edges made transparent.
    
    'We also need to determine the thumbnail's actual width and height, and any x and y offset necessary to preserve the
    ' aspect ratio and center the image on the thumbnail.
    Dim tIcoWidth As Long, tIcoHeight As Long, tX As Double, tY As Double
    
    'Start by determining proper dimensions for the resized thumbnail image.
    ConvertAspectRatio m_SelMask.GetDIBWidth, m_SelMask.GetDIBHeight, thumbnailSize, thumbnailSize, tIcoWidth, tIcoHeight
    
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
    dstThumbnailDIB.CreateBlank thumbnailSize, thumbnailSize, 32, 0
    
    'Note that the user's thumbnail performance setting affects the quality used here.
    RequestThumbnail = GDIPlusResizeDIB(dstThumbnailDIB, tX, tY, tIcoWidth, tIcoHeight, m_SelMask, 0, 0, m_SelMask.GetDIBWidth, m_SelMask.GetDIBHeight, g_UserPreferences.GetThumbnailInterpolationPref())
    
    'NOTE: now that selection masks are 32-bpp, the destination thumbnail DIB is going to have wonky alpha values in the selection
    ' mask region.  Premultiply alpha manually to compensate.
    dstThumbnailDIB.SetAlphaPremultiplication True
    
End Function
