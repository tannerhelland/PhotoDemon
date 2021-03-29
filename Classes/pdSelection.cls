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
'Copyright 2012-2021 by Tanner Helland
'Created: 25/September/12
'Last updated: 03/May/18
'Last update: add support for locking width, height, aspect-ratio
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
' selection.  From these, a uniform set of boundaries are calculated and stored inside the m_CornersLocked RectF; these define
' the boundaries of the raw geometric shape constructed by the selection.  These boundary values are primarily used internally,
' or to show the user the size of their created shape (rectangle and elliptical selections allow the user to modify these values
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Used when writing/reading selection data to/from a file
Private Const SELECTION_IDENTIFIER As Long = &H64734450     'PDsd
Private Const SELECTION_FILE_VERSION_2017 As Long = &H1004
Private Const SELECTION_FILE_VERSION_2019 As Long = &H1005&

'What shape does this selection have?
Public Enum PD_SelectionShape
    ss_Unknown = -1
    ss_Rectangle = 0
    ss_Circle = 1
    ss_Polygon = 2
    ss_Lasso = 3
    ss_Wand = 4
    ss_Raster = 5
End Enum

#If False Then
    Private Const ss_Unknown = -1, ss_Rectangle = 0, ss_Circle = 1, ss_Polygon = 2, ss_Lasso = 3, ss_Wand = 4, ss_Raster = 5
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

'When accessing properties, use the following enum.  The property dictionary automatically handles the interface between these
' values and the dictionary itself.
Public Enum PD_SelectionProperty
    sp_Area = 0
    sp_Smoothing = 1
    sp_BorderWidth = 2
    sp_FeatheringRadius = 3
    sp_RoundedCornerRadius = 4
    sp_SmoothStroke = 5
    sp_PolygonCurvature = 6
    sp_WandTolerance = 7
    sp_WandSearchMode = 8
    sp_WandSampleMerged = 9
    sp_WandCompareMethod = 10
End Enum

#If False Then
    Private Const sp_Area = 0, sp_Smoothing = 1, sp_BorderWidth = 2, sp_FeatheringRadius = 3, sp_RoundedCornerRadius = 4, sp_SmoothStroke = 5, sp_PolygonCurvature = 6, sp_WandTolerance = 7, sp_WandSearchMode = 8, sp_WandSampleMerged = 9, sp_WandCompareMethod = 10
#End If

Public Enum PD_SelectionLockable
    pdsl_Width = 0
    pdsl_Height = 1
    pdsl_AspectRatio = 2
End Enum

#If False Then
    Private Const pdsl_Width = 0, pdsl_Height = 1, pdsl_AspectRatio = 2
#End If

'The coordinates of the current selection (if rectangular)
Private m_CornersUnlocked As RectF_RB

'When a selection is "locked in", the x and y values of corner points are converted to these values
Private m_CornersLocked As RectF

'And finally, for nonstandard or inverted selections, an additional set of coordinates is required to track the bounding area of the
' selection as a whole.  External functions need these values to determine how to handle the selection's data.
Private m_Bounds As RectF

'Is this selection "locked in"?
Private m_IsLocked As Boolean

'Some selection attributes can be independently locked (e.g. aspect ratio).  Note also that the
' "m_LockedAspectRatio" value is only valid WHEN the aspect ratio is locked.
Private m_IsWidthLocked As Boolean, m_IsHeightLocked As Boolean, m_IsAspectLocked As Boolean
Private m_LockedWidth As Long, m_LockedHeight As Long, m_LockedAspectRatio As Double

'When the selection is moved, it's necessary to know the difference between the current mouse point and the original left / top values
Private m_MoveXDist As Single, m_MoveYDist As Single

'Is this selection requesting text box updates? Because if it is, ignore external requests to update.
Private m_RejectRefreshRequests As Boolean

'Is a current "point of interest" selected?  (Note that different selection shapes may use this value differently; only certain values
' are hard-coded across multiple shapes.)
Private m_CurrentPOI As PD_PointOfInterest

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

'When in outline or marching ant rendering mode, the current on-screen outline (e.g. the selection outline *already transformed*
' to viewport coordinates) is cached to improve performance.  Similarly, certain outline rendering properties are also cached.
Private m_FinalOutline As pd2DPath, m_FinalOutlineIsReady As Boolean
Private Const ANT_DASH_SIZE As Single = 4#
Private Const ANT_DASH_SPEED As Long = 280
Private m_AntDashes() As Single, m_AntDashOffset As Single

'External forces may need to suspend animations for various reasons (e.g. a separate dialog is being raised).
' This value will be set to TRUE if animations are currently allowed; FALSE otherwise.  Note that this value is
' only checked *inside* the timer event, so the timer itself still needs to be activated/deactivated properly.
Private m_AnimationsAllowed As Boolean

'Some types of transformations (basic shapes, like rectangles, etc), are transformable, meaning that after they are created, the user can
' click on them again to resize or move them.  Complex transformations (magic wand, lasso, etc), may not be transformable.  This boolean
' is read by the mouse tracker in the image form, and it uses it to determine if the user is allowed to transform the current selection.
Private m_IsTransformable As Boolean

'If the user holds "Shift" while moving the selection, it will be forced to a 1:1 aspect ratio.
Private m_IsSquare As Boolean

'Lasso and polygon selections have a variable number of points.  As such, we have to track their contents dynamically.
Private m_NumOfLassoPoints As Long, m_LassoPoints() As PointFloat
Private m_NumOfPolygonPoints As Long, m_PolygonPoints() As PointFloat

'When move transforms are applied to lasso or polygon selections, we must maintain a backup copy of the original point array.
' Any transformation values are then applied to this backup data, and the backup is erased when the mouse is released.
Private m_LassoPointsBackup() As PointFloat, m_PolygonPointsBackup() As PointFloat

'When a polygon or lasso selection is closed, these values will be set to TRUE
Private m_PolygonClosed As Boolean, m_LassoClosed As Boolean

'Magic wand selections supply their own custom wand outline, and they are handled specially through a pdFloodFill instance.
Private m_FloodFill As pdFloodFill, m_WandOutline As pd2DPath

'Magic wand selections also require a special copy of the current relevant image data.  (This may be a fully composited
' image copy, or a null-padded version of the current layer.)  We cache this locally to improve wand performance, and we
' use the notification timestamp from the parent image to determine when it's time to update our local copy.
Private m_WandImage As pdDIB, m_WandImageTimestamp As Currency

'To improve viewport rendering performance, selection objects store a viewport-sized buffer of the current selection mask.
' This viewport-sized DIB can be used as a reference for calculating any on-screen elements, e.g. the actual "highlight"
' or "outline" objects that we render onto the viewport.  This reference DIB must be recreated whenever the viewport size
' and/or position changes (e.g. when zoom or scroll are invoked); to that end, we store several viewport rect copies,
' which we can compare against current viewport settings to detect changes.
'
'Because of its universal helpfulness, this overlay is always synchronized, regardless of the current selection rendering
' technique.  (Note that the position of the overlay is also cached in a RECTF object, although - importantly! - this RectF
' is *not* boundary-checked against the size of the viewport.  Its coordinates may lay outside the overlay reference DIB,
' by design.)
Private m_ViewportReference As pdDIB, m_ViewportRefRectF As RectF, m_ViewportRefReady As Boolean
Private m_LastViewportRectF As RectF, m_LastImageRectF As RectF, m_LastRenderMode As PD_SelectionRender

'To further improve viewport rendering performance when "highlight" and "lightbox" rendering modes are active, this class
' also caches such overlays inside a persistent DIB.  Because this DIB is inextricably tied to the current viewport,
' it must be updated whenever the cached overlay reference DIB, above, changes.
Private m_ViewportOverlay As pdDIB, m_OverlayIsReady As Boolean

'To further improve viewport rendering when the "outline" and "marching ant" rendering modes are active, this class caches
' an extra copy of the selection mask, converted to a single-color byte array and clipped to [0, 255] values.  This array
' also guarantees empty boundary pixels, and it must be regenerated whenever the reference DIB, above, changes.
'
'(Note that we also store the dimensions of this array, and its (x, y) offset relative to the underlying image.)
Private m_ViewportByteCopy() As Byte
Private m_ViewportByteCopyRect As RectL_WH
Private m_ViewportEdges As pdEdgeDetector

'When selection UI settings change, we generate a local "hash" of the current settings.  Because we cache UI overlays
' locally (to improve rendering time), we need to regenerate the cache whenever rendering settings change, independent of
' any changes to the actual underlying selection object.
Private m_LastOverlayCache As String

'Some selection use a timestamp to improve UI behavior
Private m_LastTime As Currency

'Each different selection shape has a number of properties specific to that shape.  Instead of storing these in open variables,
' they are stored in a dictionary, and only created/accessed as necessary.
Private m_PropertyDict As pdDictionary

'Marching ant outlines require an animation timer
Private WithEvents m_AntTimer As pdTimer
Attribute m_AntTimer.VB_VarHelpID = -1

'New project: I'm going to attempt to migrate selection behavior to individual classes.  This is a prerequisite
' for supporting multiple selections at once, since we need a way to manage multiple selections simultaneously.
' (This could also greatly simplify and streamline future selection enhancements.)
'
'For now, I'm starting with lasso selections *only*.  Once I've nailed down the required interface definitions,
' I'll extend the new strategy to *all* selection types.
Private m_SelectionLasso As pdSelection_Lasso

'Retrieve various boundary rects.  Full descriptions of these rects are given above.
Friend Function GetCornersUnlockedRect() As RectF_RB
    GetCornersUnlockedRect = m_CornersUnlocked
End Function

Friend Function GetCornersLockedRect() As RectF
    GetCornersLockedRect = m_CornersLocked
End Function

Friend Function GetBoundaryRect() As RectF
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

Friend Function GetMaskDIB_ViewportCopy() As pdDIB
    If (Not m_IsMaskReady) Then CreateSelectionMask
    Set GetMaskDIB_ViewportCopy = m_ViewportReference
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
        
        Case sp_SmoothStroke
            If (m_SelectionShape = ss_Lasso) Then m_IsMaskReady = False
        
        Case sp_PolygonCurvature
            If (m_SelectionShape = ss_Polygon) Then m_IsMaskReady = False
        
        Case sp_WandTolerance
            If (m_SelectionShape = ss_Wand) Then m_IsMaskReady = False
        
        Case sp_WandSearchMode
            If (m_SelectionShape = ss_Wand) Then m_IsMaskReady = False
        
        Case sp_WandSampleMerged
            m_WandImageTimestamp = 0@
            If (m_SelectionShape = ss_Wand) Then m_IsMaskReady = False
        
        Case sp_WandCompareMethod
            If (m_SelectionShape = ss_Wand) Then m_IsMaskReady = False
        
    End Select
    
End Sub

'To save code elsewhere, a selection can be initialized by an XML string generated by a pdSerialize object.
' (This function is primarily used when writing full selection data to/from file.)
Friend Sub InitFromXML(ByRef srcXML As String)

    Dim i As Long
    
    'Start by creating a parameter parser to handle the parameter string.  This class will parse out individual parameters
    ' as specific data types upon request.
    If (LenB(srcXML) <> 0) Then
        
        'Do a quick failsafe check to make sure the passed string differs from our current settings.  If it doesn't,
        ' we can fully skip this step.
        If Strings.StringsNotEqual(srcXML, Me.GetSelectionAsXML, False) Then
            
            Dim cParams As pdSerialize
            Set cParams = New pdSerialize
            cParams.SetParamString srcXML
            
            'All selections share a core set of universal properties
            SetSelectionShape cParams.GetLong("SelectionShape", ss_Rectangle)
            SetSelectionProperty sp_Area, cParams.GetLong("SelectionArea", sa_Interior)
            SetSelectionProperty sp_Smoothing, cParams.GetLong("SelectionSmoothing", es_None)
            SetSelectionProperty sp_FeatheringRadius, cParams.GetLong("SelectionFeatheringRadius", 0)
            SetSelectionProperty sp_BorderWidth, cParams.GetLong("SelectionBorderWidth", 1)
            SetSelectionProperty sp_RoundedCornerRadius, cParams.GetDouble("SelectionRoundedCornerRadius", 0#)
            
            'All coordinates must be transformed from absolute values to relative ones (on the scale [0, 1])
            With m_CornersLocked
                .Left = PDMath.TranslateValue_RelToAbs(cParams.GetSingle("SelectionLeft", 0!), m_parentPDImage.Width)
                .Top = PDMath.TranslateValue_RelToAbs(cParams.GetSingle("SelectionTop", 0!), m_parentPDImage.Height)
                .Width = PDMath.TranslateValue_RelToAbs(cParams.GetSingle("SelectionWidth", 0!), m_parentPDImage.Width)
                .Height = PDMath.TranslateValue_RelToAbs(cParams.GetSingle("SelectionHeight", 0!), m_parentPDImage.Height)
            End With
            
            'Additional parameters vary by selection type
            Select Case m_SelectionShape
            
                'Rectangles, ellipses, and lines only require bounding rect values (stored as two (x, y) pairs).  Retrieve these
                ' values from positions [12, 15] inclusive.
                Case ss_Rectangle, ss_Circle
                
                    With m_CornersUnlocked
                        .Left = PDMath.TranslateValue_RelToAbs(cParams.GetSingle("SelectionVectorLeft", 0!), m_parentPDImage.Width)
                        .Top = PDMath.TranslateValue_RelToAbs(cParams.GetSingle("SelectionVectorTop", 0!), m_parentPDImage.Height)
                        .Right = PDMath.TranslateValue_RelToAbs(cParams.GetSingle("SelectionVectorRight", 0!), m_parentPDImage.Width)
                        .Bottom = PDMath.TranslateValue_RelToAbs(cParams.GetSingle("SelectionVectorBottom", 0!), m_parentPDImage.Height)
                    End With
                
                'Polygon selections have a variable number of parameters, based on the number of points in the polygon.
                Case ss_Polygon
                    
                    SetSelectionProperty sp_PolygonCurvature, cParams.GetDouble("SelectionPolygonCurvature", 0#)
                    m_NumOfPolygonPoints = cParams.GetLong("SelectionNumOfPoints", 0)
                    
                    If (m_NumOfPolygonPoints > 0) Then
                    
                        ReDim m_PolygonPoints(0 To m_NumOfPolygonPoints - 1) As PointFloat
                        
                        'Retrieve all polygon points from the param string
                        If (m_NumOfPolygonPoints > 0) Then
                            For i = 0 To m_NumOfPolygonPoints - 1
                                With m_PolygonPoints(i)
                                    .x = PDMath.TranslateValue_RelToAbs(cParams.GetDouble("SelectionPointX" & CStr(i + 1), 0), m_parentPDImage.Width)
                                    .y = PDMath.TranslateValue_RelToAbs(cParams.GetDouble("SelectionPointY" & CStr(i + 1), 0), m_parentPDImage.Height)
                                End With
                            Next i
                        End If
                        
                        'Create a copy of the current point collection; this will be used for transforms
                        BackupCurrentSelectionPoints
                        
                    Else
                        ReDim m_PolygonPoints(0 To 127) As PointFloat
                        ReDim m_PolygonPointsBackup(0 To 127) As PointFloat
                    End If
                
                'Lasso/freehand selections have a variable number of parameters, based on the number of points in the lasso.
                Case ss_Lasso
                    
                    SetSelectionProperty sp_SmoothStroke, cParams.GetDouble("SelectionSmoothStroke", 0#)
                    m_NumOfLassoPoints = cParams.GetLong("SelectionNumOfPoints", 0)
                    
                    If (m_NumOfLassoPoints > 0) Then
                    
                        ReDim m_LassoPoints(0 To m_NumOfLassoPoints - 1) As PointFloat
                        
                        'Retrieve all remaining lasso points
                        If (m_NumOfLassoPoints > 0) Then
                            For i = 0 To m_NumOfLassoPoints - 1
                                With m_LassoPoints(i)
                                    .x = PDMath.TranslateValue_RelToAbs(cParams.GetDouble("SelectionPointX" & CStr(i + 1), 0#), m_parentPDImage.Width)
                                    .y = PDMath.TranslateValue_RelToAbs(cParams.GetDouble("SelectionPointY" & CStr(i + 1), 0#), m_parentPDImage.Height)
                                End With
                            Next i
                        End If
                        
                    Else
                        ReDim m_LassoPoints(0 To 127) As PointFloat
                    End If
                
                'Almost all of the "magic wand" selection properties are unique (relative to other selection shapes)
                Case ss_Wand
                    
                    'For weird historical reasons, wand selections store the wand origin point inside the unlocked selection rect
                    With m_CornersUnlocked
                        .Left = PDMath.TranslateValue_RelToAbs(cParams.GetDouble("SelectionWandX", 0), m_parentPDImage.Width)
                        .Top = PDMath.TranslateValue_RelToAbs(cParams.GetDouble("SelectionWandY", 0), m_parentPDImage.Height)
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
            
        'End matching param string check
        End If
        
    'End non-zero XML length check
    End If
        
End Sub

'Return all of this selection's important settings as an XML packet.  If the current selection shape is vector-based,
' this packet can be used to exactly re-create this selection.  (See also: "InitFromXML()" above)
Friend Function GetSelectionAsXML() As String
    
    Dim i As Long
    
    'All selection types store a preset list of standardized values
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    cParams.AddParam "SelectionShape", Me.GetSelectionShape
    cParams.AddParam "SelectionArea", Me.GetSelectionProperty(sp_Area)
    cParams.AddParam "SelectionSmoothing", Me.GetSelectionProperty(sp_Smoothing)
    cParams.AddParam "SelectionFeatheringRadius", Me.GetSelectionProperty(sp_FeatheringRadius)
    cParams.AddParam "SelectionBorderWidth", Me.GetSelectionProperty(sp_BorderWidth)
    cParams.AddParam "SelectionRoundedCornerRadius", Me.GetSelectionProperty(sp_RoundedCornerRadius)
    
    With m_CornersLocked
        cParams.AddParam "SelectionLeft", PDMath.TranslateValue_AbsToRel(.Left, m_parentPDImage.Width)
        cParams.AddParam "SelectionTop", PDMath.TranslateValue_AbsToRel(.Top, m_parentPDImage.Height)
        cParams.AddParam "SelectionWidth", PDMath.TranslateValue_AbsToRel(.Width, m_parentPDImage.Width)
        cParams.AddParam "SelectionHeight", PDMath.TranslateValue_AbsToRel(.Height, m_parentPDImage.Height)
    End With
    
    'After universal values, we add any additional properties specific to the current selection shape
    Select Case m_SelectionShape
    
        'Rectangles, ellipses, and lines store corner coordinates
        Case ss_Rectangle, ss_Circle
            With m_CornersUnlocked
                cParams.AddParam "SelectionVectorLeft", PDMath.TranslateValue_AbsToRel(.Left, m_parentPDImage.Width)
                cParams.AddParam "SelectionVectorTop", PDMath.TranslateValue_AbsToRel(.Top, m_parentPDImage.Height)
                cParams.AddParam "SelectionVectorRight", PDMath.TranslateValue_AbsToRel(.Right, m_parentPDImage.Width)
                cParams.AddParam "SelectionVectorBottom", PDMath.TranslateValue_AbsToRel(.Bottom, m_parentPDImage.Height)
            End With
            
        'Polygon selections are trickier; they have a dynamic parameter list based on the number of points in the shape.
        Case ss_Polygon
            
            cParams.AddParam "SelectionPolygonCurvature", Me.GetSelectionProperty(sp_PolygonCurvature)
            cParams.AddParam "SelectionNumOfPoints", m_NumOfPolygonPoints
            
            If (m_NumOfPolygonPoints > 0) Then
                For i = 0 To m_NumOfPolygonPoints - 1
                    With m_PolygonPoints(i)
                        cParams.AddParam "SelectionPointX" & CStr(i + 1), PDMath.TranslateValue_AbsToRel(.x, m_parentPDImage.Width), True
                        cParams.AddParam "SelectionPointY" & CStr(i + 1), PDMath.TranslateValue_AbsToRel(.y, m_parentPDImage.Height), True
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
                        cParams.AddParam "SelectionPointX" & CStr(i + 1), PDMath.TranslateValue_AbsToRel(.x, m_parentPDImage.Width), True
                        cParams.AddParam "SelectionPointY" & CStr(i + 1), PDMath.TranslateValue_AbsToRel(.y, m_parentPDImage.Height), True
                    End With
                Next i
            End If
            
        'Wand selections have a few wand-specific fixed params
        Case ss_Wand
        
            'For weird historical reasons, wand selections store the wand origin point inside the unlocked selection rect
            With m_CornersUnlocked
                cParams.AddParam "SelectionWandX", PDMath.TranslateValue_AbsToRel(.Left, m_parentPDImage.Width)
                cParams.AddParam "SelectionWandY", PDMath.TranslateValue_AbsToRel(.Top, m_parentPDImage.Height)
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

'Simple - use this to select the entire image attached to this selection object.  Note that this DOES NOT render the new selection on-screen,
' and it doesn't render a matching selection mask.  The caller is responsible for that.
Friend Sub SelectAll()

    'Set basic information about this selection
    SetSelectionShape ss_Rectangle
    SetSelectionProperty sp_Area, sa_Interior
    SetSelectionProperty sp_FeatheringRadius, 0
    SetSelectionProperty sp_Smoothing, es_Antialiased
    SetSelectionProperty sp_RoundedCornerRadius, 0#
    
    With m_CornersUnlocked
        .Left = 0!
        .Top = 0!
        .Right = m_parentPDImage.Width
        .Bottom = m_parentPDImage.Height
    End With
    
    With m_CornersLocked
        .Left = 0!
        .Top = 0!
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
Friend Function GetActiveSelectionPOI(Optional ByVal onlyReportActiveTransform As Boolean = True) As PD_PointOfInterest
    If onlyReportActiveTransform Then
        If m_TransformModeActive Then GetActiveSelectionPOI = m_CurrentPOI Else GetActiveSelectionPOI = poi_Undefined
    Else
        GetActiveSelectionPOI = m_CurrentPOI
    End If
End Function

Friend Sub SetActiveSelectionPOI(ByVal newPOI As PD_PointOfInterest)
    m_CurrentPOI = newPOI
End Sub

'If the current selection shape supports "resize via corner node", this function will return TRUE.  Generally speaking, corner-based
' resize is handled "for free" by the selection engine (meaning it handles all UI rendering and input related to said corners).
'
'Certain selection types, like "line" or "magic wand", may not support this behavior.
Friend Function DoesShapeSupportCornerResize() As Boolean
    If (m_SelectionShape = ss_Rectangle) Then
        DoesShapeSupportCornerResize = True
    ElseIf (m_SelectionShape = ss_Circle) Then
        DoesShapeSupportCornerResize = True
    Else
        DoesShapeSupportCornerResize = False
    End If
End Function

'If the current selection shape supports custom points of interest (e.g. points other than just the corners), this function will
' return TRUE.  You can retrieve the actual POIs with the GetCurrentPOIList() function, below.
Friend Function DoesShapeSupportCustomPOIs() As Boolean
    If (m_SelectionShape = ss_Rectangle) Then
        DoesShapeSupportCustomPOIs = False
    ElseIf (m_SelectionShape = ss_Circle) Then
        DoesShapeSupportCustomPOIs = False
    ElseIf (m_SelectionShape = ss_Lasso) Then
        DoesShapeSupportCustomPOIs = False
    ElseIf (m_SelectionShape = ss_Polygon) Then
        DoesShapeSupportCustomPOIs = True
    ElseIf (m_SelectionShape = ss_Wand) Then
        DoesShapeSupportCustomPOIs = True
    Else
        DoesShapeSupportCustomPOIs = False
    End If
End Function

'Return a list of the current selection "points of interest".  Not all selection types support custom POIs.
' This function returns TRUE if POIs are both supported and available; you can also check this via DoesShapeSupportCustomPOIs(), above.
Friend Function GetCurrentPOIList(ByRef dstPoints() As PointFloat) As Boolean
    
    GetCurrentPOIList = Me.DoesShapeSupportCustomPOIs()
    
    If GetCurrentPOIList Then
        If (m_SelectionShape = ss_Wand) Then
            ReDim dstPoints(0) As PointFloat
            dstPoints(0).x = m_CornersUnlocked.Left
            dstPoints(0).y = m_CornersUnlocked.Top
        End If
    End If
    
End Function

Friend Sub OverrideTransformMode(ByVal newOverride As Boolean)
    m_TransformModeActive = newOverride
End Sub

'Get/set a selection shape
Friend Function GetSelectionShape() As PD_SelectionShape
    GetSelectionShape = m_SelectionShape
End Function

Friend Sub SetSelectionShape(ByVal selShape As PD_SelectionShape)
    
    'Certain types of shapes are transformable.  Mark those now.
    Select Case selShape
        
        Case ss_Rectangle, ss_Circle
            m_IsTransformable = True
            
        Case ss_Lasso, ss_Polygon
            m_IsTransformable = True
            
        Case ss_Wand
            m_IsTransformable = False
        
    End Select
    
    'Invalidate any internal trackers that are reliant on specific shapes
    If (m_SelectionShape <> ss_Wand) Then
        Set m_WandImage = Nothing
        If (Not m_FloodFill Is Nothing) Then m_FloodFill.FreeUpResources
    End If
    
    If (m_SelectionShape <> ss_Raster) Then m_IsMaskReady = False
    m_SelectionShape = selShape
    
End Sub

'Return the number of polygon points
Friend Function GetNumOfPolygonPoints() As Long
    GetNumOfPolygonPoints = m_NumOfPolygonPoints
End Function

'Copy the current polygon point collection into an arbitrary destination array.  This is used by mouse coordinate checking functions.
Friend Sub GetPolygonPoints(ByRef ptFloatArray() As PointFloat)
    
    If (m_NumOfPolygonPoints > 0) Then
        ReDim ptFloatArray(0 To m_NumOfPolygonPoints - 1) As PointFloat
        CopyMemoryStrict VarPtr(ptFloatArray(0)), VarPtr(m_PolygonPoints(0)), m_NumOfPolygonPoints * 8
    Else
        ReDim ptFloatArray(0) As PointFloat
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
    m_ViewportRefReady = False
End Sub

'Some properties can be independently locked (e.g. size).  When locked, these properties cannot be changed by UI inputs.
Friend Sub LockProperty(ByVal selProperty As PD_SelectionLockable, ByVal lockedValue As Variant)

    If (selProperty = pdsl_Width) Then
        m_IsWidthLocked = True
        m_LockedWidth = lockedValue
        m_IsHeightLocked = False
        m_IsAspectLocked = False
    ElseIf (selProperty = pdsl_Height) Then
        m_IsHeightLocked = True
        m_LockedHeight = lockedValue
        m_IsWidthLocked = False
        m_IsAspectLocked = False
    ElseIf (selProperty = pdsl_AspectRatio) Then
        m_IsAspectLocked = True
        m_LockedAspectRatio = lockedValue
        m_IsWidthLocked = False
        m_IsHeightLocked = False
    End If
    
End Sub

Friend Sub UnlockProperty(ByVal selProperty As PD_SelectionLockable)
    If (selProperty = pdsl_Width) Then
        m_IsWidthLocked = False
    ElseIf (selProperty = pdsl_Height) Then
        m_IsHeightLocked = False
    ElseIf (selProperty = pdsl_AspectRatio) Then
        m_IsAspectLocked = False
    End If
End Sub

Friend Function GetPropertyLockedState(ByVal selProperty As PD_SelectionLockable) As Boolean
    If (selProperty = pdsl_Width) Then
        GetPropertyLockedState = m_IsWidthLocked
    ElseIf (selProperty = pdsl_Height) Then
        GetPropertyLockedState = m_IsHeightLocked
    ElseIf (selProperty = pdsl_AspectRatio) Then
        GetPropertyLockedState = m_IsAspectLocked
    End If
End Function

'For lasso and polygon selections, this function will return the current selecton region
' as a pd2DRegion instance.  The mouse coordinate checker code uses this to see if the
' mouse cursor is currently within the bounds of the selection area.  (This is only
' implemented for polygon and lasso selections; other selection types use much simpler
' interior detection code.)
Friend Function GetSelectionAsRegion() As pd2DRegion
    
    If (m_SelectionShape = ss_Polygon) Then
        If (m_NumOfPolygonPoints < 3) Then Exit Function
    ElseIf (m_SelectionShape = ss_Lasso) Then
        If (m_NumOfLassoPoints < 3) Then Exit Function
    Else
        PDDebug.LogAction "WARNING!  Not implemented for this selection shape!"
        Exit Function
    End If
    
    Set GetSelectionAsRegion = New pd2DRegion
    
    'Use a path to construct the initial region
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    tmpPath.SetFillRule P2_FR_Winding
    
    If (m_SelectionShape = ss_Polygon) Then
        tmpPath.AddPolygon m_NumOfPolygonPoints, VarPtr(m_PolygonPoints(0)), True, GetSelectionProperty_Float(sp_PolygonCurvature)
    Else
        tmpPath.AddPolygon m_NumOfLassoPoints, VarPtr(m_LassoPoints(0)), GetSelectionProperty_Float(sp_SmoothStroke)
    End If
    
    GetSelectionAsRegion.AddPath tmpPath, P2_CM_Replace
    
End Function

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
                    If (Not m_IsWidthLocked) Then .Right = x Else .Right = m_CornersLocked.Left + m_CornersLocked.Width
                    If (Not m_IsHeightLocked) Then .Bottom = y Else .Bottom = m_CornersLocked.Top
                End With
            
            Case poi_CornerSE
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left
                    .Top = m_CornersLocked.Top
                    If (Not m_IsWidthLocked) Then .Right = x Else .Right = m_CornersLocked.Left + m_CornersLocked.Width
                    If (Not m_IsHeightLocked) Then .Bottom = y Else .Bottom = m_CornersLocked.Top + m_CornersLocked.Height
                End With
            
            Case poi_CornerSW
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left + m_CornersLocked.Width
                    .Top = m_CornersLocked.Top
                    If (Not m_IsWidthLocked) Then .Right = x Else .Right = m_CornersLocked.Left
                    If (Not m_IsHeightLocked) Then .Bottom = y Else .Bottom = m_CornersLocked.Top + m_CornersLocked.Height
                End With
            
            Case poi_EdgeN
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left
                    .Right = m_CornersLocked.Left + m_CornersLocked.Width
                    .Top = m_CornersLocked.Top + m_CornersLocked.Height
                    If (Not m_IsHeightLocked) Then .Bottom = y Else .Bottom = m_CornersLocked.Top
                End With
            
            Case poi_EdgeE
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left
                    If (Not m_IsWidthLocked) Then .Right = x Else .Right = m_CornersLocked.Left + m_CornersLocked.Width
                    .Top = m_CornersLocked.Top
                    .Bottom = m_CornersLocked.Top + m_CornersLocked.Height
                End With
            
            Case poi_EdgeS
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left
                    .Right = m_CornersLocked.Left + m_CornersLocked.Width
                    .Top = m_CornersLocked.Top
                    If (Not m_IsHeightLocked) Then .Bottom = y Else .Bottom = m_CornersLocked.Top + m_CornersLocked.Height
                End With
            
            Case poi_EdgeW
                With m_CornersUnlocked
                    .Left = m_CornersLocked.Left + m_CornersLocked.Width
                    If (Not m_IsWidthLocked) Then .Right = x Else .Right = m_CornersLocked.Left
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
        
    'Polygons can be both moved and resized (by click-dragging individual points)
    ElseIf (m_SelectionShape = ss_Polygon) Then
    
        'Based on the transform mode, set the initial points accordingly
        Select Case m_CurrentPOI
    
            'Interior of the selection, which triggers a move transformation
            Case poi_Interior
                m_MoveXDist = x
                m_MoveYDist = y
                
                'Create a safe copy of the current point collection
                BackupCurrentSelectionPoints
                
            'Other transforms don't actually require special handling, as all the work is done in the
            ' SetAdditionalTransformCoordinates function.
            Case Else
                
        End Select
        
    'Lasso selections can only be moved, so the only valid transform type is type 0
    ElseIf (m_SelectionShape = ss_Lasso) Then
    
        'Based on the transform mode, set the initial points accordingly
        Select Case m_CurrentPOI
    
            'Interior of the selection, which triggers a move transformation
            Case poi_Interior
                m_MoveXDist = x
                m_MoveYDist = y
                
                'Create a safe copy of the current point collection
                BackupCurrentSelectionPoints
                
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
        
        Case ss_Rectangle, ss_Circle
            
            With m_CornersUnlocked
                If m_IsWidthLocked Then .Right = x + m_LockedWidth Else .Right = x
                If m_IsHeightLocked Then .Bottom = y + m_LockedHeight Else .Bottom = y
            End With
            
            'Also, mark selections of any of these shapes (rectangle, circle, line) as transformable
            m_IsTransformable = True
        
        'Polygon selections require us to initialize a whole bunch of polygon tracking variables
        Case ss_Polygon
            m_IsTransformable = True
            
            m_NumOfPolygonPoints = 1
            ReDim m_PolygonPoints(0 To 127) As PointFloat
            
            m_PolygonPoints(0).x = x
            m_PolygonPoints(0).y = y
            
            'Also set up a dummy set of initial boundary coordinates
            With m_CornersLocked
                .Left = x
                .Top = y
                .Width = 1
                .Height = 1
            End With
            
            'Create a copy of the current point collection; this will be used for transforms
            BackupCurrentSelectionPoints
            
        'Lasso selections have limited transform capabilities, and they have a variable number of points
        Case ss_Lasso
            m_IsTransformable = True
            
            m_NumOfLassoPoints = 1
            ReDim m_LassoPoints(0 To 127) As PointFloat
            
            m_LassoPoints(0).x = x
            m_LassoPoints(0).y = y
            
            'Also set up a dummy set of initial boundary coordinates
            With m_CornersLocked
                .Left = x
                .Top = y
                .Width = 1
                .Height = 1
            End With
            
            'Create a safe copy of the current point collection
            BackupCurrentSelectionPoints
        
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
                        
                        'Locked aspect ratio requires a more detailed approach
                        If m_IsAspectLocked And (m_LockedAspectRatio > 0#) Then
                            m_CornersUnlocked.Right = x
                            If (m_CurrentPOI = poi_CornerNW) Or (m_CurrentPOI = poi_CornerSE) Then
                                m_CornersUnlocked.Bottom = m_CornersUnlocked.Top + (m_CornersUnlocked.Right - m_CornersUnlocked.Left) * (1# / m_LockedAspectRatio)
                            ElseIf (m_CurrentPOI = poi_CornerNE) Or (m_CurrentPOI = poi_CornerSW) Then
                                m_CornersUnlocked.Bottom = m_CornersUnlocked.Top - (m_CornersUnlocked.Right - m_CornersUnlocked.Left) * (1# / m_LockedAspectRatio)
                            End If
                        Else
                            If (Not m_IsWidthLocked) Then m_CornersUnlocked.Right = x
                            If (Not m_IsHeightLocked) Then m_CornersUnlocked.Bottom = y
                        End If
                        
                    'Edges
                    Case poi_EdgeN
                        If m_IsAspectLocked And (m_LockedAspectRatio > 0#) Then
                            m_CornersUnlocked.Bottom = y
                            m_CornersUnlocked.Right = m_CornersUnlocked.Left - (m_CornersUnlocked.Bottom - m_CornersUnlocked.Top) * m_LockedAspectRatio
                        Else
                            If (Not m_IsHeightLocked) Then m_CornersUnlocked.Bottom = y
                        End If
                        
                    Case poi_EdgeE
                        If m_IsAspectLocked And (m_LockedAspectRatio > 0#) Then
                            m_CornersUnlocked.Right = x
                            m_CornersUnlocked.Bottom = m_CornersUnlocked.Top + (m_CornersUnlocked.Right - m_CornersUnlocked.Left) * (1# / m_LockedAspectRatio)
                        Else
                            If (Not m_IsWidthLocked) Then m_CornersUnlocked.Right = x
                        End If
                        
                    Case poi_EdgeS
                        If m_IsAspectLocked And (m_LockedAspectRatio > 0#) Then
                            m_CornersUnlocked.Bottom = y
                            m_CornersUnlocked.Right = m_CornersUnlocked.Left + (m_CornersUnlocked.Bottom - m_CornersUnlocked.Top) * m_LockedAspectRatio
                        Else
                            If (Not m_IsHeightLocked) Then m_CornersUnlocked.Bottom = y
                        End If
                        
                    Case poi_EdgeW
                        If m_IsAspectLocked And (m_LockedAspectRatio > 0#) Then
                            m_CornersUnlocked.Right = x
                            m_CornersUnlocked.Bottom = m_CornersUnlocked.Top - (m_CornersUnlocked.Right - m_CornersUnlocked.Left) * (1# / m_LockedAspectRatio)
                        Else
                            If (Not m_IsWidthLocked) Then m_CornersUnlocked.Right = x
                        End If
                        
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
            
            'Polygon transforms consist of either moving an individual polygon point, or moving the entire polygon array
            Case ss_Polygon
            
                'Ignore events if they happen too close together.  (This is part of a fix for https://github.com/tannerhelland/PhotoDemon/issues/284)
                If ((VBHacks.GetHighResTimeInMSEx() - m_LastTime) < 500) Then Exit Sub
            
                'Based on the transform mode, set the initial points accordingly
                Select Case m_CurrentPOI
            
                    'Case -1 should never occur.  (-1 represents an invalid transformation request)
                    Case poi_Undefined
                    
                    'Move transform
                    Case poi_Interior
                        
                        'Failsafe check for rapid clicks
                        If (m_NumOfPolygonPoints - 1 <= UBound(m_PolygonPointsBackup)) Then
                            
                            'Rebuild the main polygon array by copying all points from the backup array, and applying the current
                            ' x/y transformation distance to them.
                            For i = 0 To m_NumOfPolygonPoints - 1
                                m_PolygonPoints(i).x = m_PolygonPointsBackup(i).x + (x - m_MoveXDist)
                                m_PolygonPoints(i).y = m_PolygonPointsBackup(i).y + (y - m_MoveYDist)
                            Next i
                            
                        End If
                    
                    'Anything else is just moving a polygon point
                    Case Else
                    
                        With m_PolygonPoints(m_CurrentPOI)
                            .x = x
                            .y = y
                        End With
                        
                        'Create a safe copy of the current point collection
                        BackupCurrentSelectionPoints
                        
                End Select
                
                'PDDebug.LogAction "Polygon transform active: " & m_NumOfPolygonPoints
                
                'If a transform mode is active, re-mark the selection as being transformable
                m_IsTransformable = True
            
            'Lasso transforms require us to transform the entire lasso array
            Case ss_Lasso
            
                'Based on the transform mode, set the initial points accordingly
                Select Case m_CurrentPOI
            
                    'Case -1 should never occur.  (-1 represents an invalid transformation request)
                    Case poi_Undefined
                    
                    'Move transform
                    Case poi_Interior
                        
                        'Failsafe check for rapid clicks
                        If (m_NumOfLassoPoints - 1 <= UBound(m_LassoPointsBackup)) Then
                                
                            'Rebuild the main lasso array by copying all points from the backup array, and applying the current
                            ' x/y transformation distance to them.
                            For i = 0 To m_NumOfLassoPoints - 1
                                m_LassoPoints(i).x = m_LassoPointsBackup(i).x + (x - m_MoveXDist)
                                m_LassoPoints(i).y = m_LassoPointsBackup(i).y + (y - m_MoveYDist)
                            Next i
                            
                        End If
                            
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
        
        'Note the current time; only polygon selections use this value (to improve UI behavior)
        m_LastTime = VBHacks.GetHighResTimeInMSEx()
        
        Select Case m_SelectionShape
        
            'Rectangle and ellipse selections may have locked aspect ratios
            Case ss_Rectangle, ss_Circle
                If m_IsAspectLocked And (m_LockedAspectRatio > 0#) Then
                    m_CornersUnlocked.Right = x
                    If (y > m_CornersUnlocked.Top) Then
                        m_CornersUnlocked.Bottom = m_CornersUnlocked.Top + Abs(m_CornersUnlocked.Right - m_CornersUnlocked.Left) * (1# / m_LockedAspectRatio)
                    Else
                        m_CornersUnlocked.Bottom = m_CornersUnlocked.Top - Abs(m_CornersUnlocked.Right - m_CornersUnlocked.Left) * (1# / m_LockedAspectRatio)
                    End If
                Else
                    If m_IsWidthLocked Then m_CornersUnlocked.Right = (m_CornersLocked.Left + m_CornersLocked.Width) Else m_CornersUnlocked.Right = x
                    If m_IsHeightLocked Then m_CornersUnlocked.Bottom = (m_CornersLocked.Top + m_CornersLocked.Height) Else m_CornersUnlocked.Bottom = y
                End If
                
            'Polygon selections will increment the current polygon array by one, adding the new point as the latest polygon coordinate
            Case ss_Polygon
            
                'Make room in the point array
                m_NumOfPolygonPoints = m_NumOfPolygonPoints + 1
                If (m_NumOfPolygonPoints > UBound(m_PolygonPoints)) Then ReDim Preserve m_PolygonPoints(0 To m_NumOfPolygonPoints * 2 - 1) As PointFloat
                    
                'Store the new point
                m_PolygonPoints(m_NumOfPolygonPoints - 1).x = x
                m_PolygonPoints(m_NumOfPolygonPoints - 1).y = y
                
                'Create a copy of the current point collection
                BackupCurrentSelectionPoints
                
                'PDDebug.LogAction "Polygon transform NOT active: " & m_NumOfPolygonPoints
                
            'Lasso selections will increment the current lasso array, and add the new point to its list
            Case ss_Lasso
            
                'Perform a quick check to make sure this point isn't a duplicate of the previous point.  With high-DPI mice,
                ' this is a distinct possibility.
                If ((m_LassoPoints(m_NumOfLassoPoints - 1).x <> x) Or (m_LassoPoints(m_NumOfLassoPoints - 1).y <> y)) Then
            
                    'Make room in the point array
                    m_NumOfLassoPoints = m_NumOfLassoPoints + 1
                    If (m_NumOfLassoPoints > UBound(m_LassoPoints)) Then ReDim Preserve m_LassoPoints(0 To m_NumOfLassoPoints * 2 - 1) As PointFloat
                    
                    'Store the new point
                    m_LassoPoints(m_NumOfLassoPoints - 1).x = x
                    m_LassoPoints(m_NumOfLassoPoints - 1).y = y
                    
                    'Create a copy of the current point collection
                    BackupCurrentSelectionPoints
                    
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
                
        Case ss_Rectangle, ss_Circle
            
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
            Case ss_Rectangle, ss_Circle
                
                'Our internal outline path cache needs to be regenerated whenever we update internal coordinates
                m_FinalOutlineIsReady = False
                
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
            
                'Our internal outline path cache needs to be regenerated whenever we update internal coordinates
                m_FinalOutlineIsReady = False
                
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
                
            'Like polygon selection, lasso selections require us to search all lasso points in order to construct
            ' a bounding rect.
            Case ss_Lasso
            
                'Our internal outline path cache needs to be regenerated whenever we update internal coordinates
                m_FinalOutlineIsReady = False
                
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
                
            'Wand selections require manual bounds-checking
            Case ss_Wand
                If Me.IsLockedIn Then
                    m_ViewportRefReady = False
                    m_OverlayIsReady = False
                    m_FinalOutlineIsReady = False
                    Me.FindNewBoundsManually True
                End If
                
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
        Case ss_Rectangle, ss_Circle
            
            If (m_CornersLocked.Left + m_CornersLocked.Width <= 0) Then AreAllCoordinatesInvalid = True
            If (m_CornersLocked.Left > m_parentPDImage.Width) Then AreAllCoordinatesInvalid = True
            If (m_CornersLocked.Top + m_CornersLocked.Height <= 0) Then AreAllCoordinatesInvalid = True
            If (m_CornersLocked.Top > m_parentPDImage.Height) Then AreAllCoordinatesInvalid = True
           
        'Polygon and Lasso are a bit more complicated.  If curvature is not active, we can use existing m_CornersLocked.Left/Top etc bounds,
        ' but if it is active, we need to perform a manual search for boundaries.
        Case ss_Polygon, ss_Lasso
        
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
    
        Case ss_Rectangle, ss_Circle
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
            
            'Create a safe copy of the current point collection
            BackupCurrentSelectionPoints
        
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
            
            'Create a safe copy of the current point collection
            BackupCurrentSelectionPoints
            
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
    If (m_NumOfPolygonPoints > 1) Then
        m_NumOfPolygonPoints = m_NumOfPolygonPoints - 1
        BackupCurrentSelectionPoints
        m_ViewportRefReady = False
        m_OverlayIsReady = False
        m_LastOverlayCache = vbNullString
        m_FinalOutlineIsReady = False
        m_IsMaskReady = False
        If (m_NumOfPolygonPoints < 2) Then Set m_FinalOutline = Nothing
    End If
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
        netDistance = netDistance + PDMath.DistanceTwoPoints(m_LassoPoints(newLassoIndex).x, m_LassoPoints(newLassoIndex).y, m_LassoPoints(newLassoIndex - 1).x, m_LassoPoints(newLassoIndex - 1).y)
    
        'Decrement the test index and repeat
        newLassoIndex = newLassoIndex - 1
    
    Loop
    
    'Reposition the lasso point index to match
    m_NumOfLassoPoints = newLassoIndex + 1
    If (m_NumOfLassoPoints < 1) Then m_NumOfLassoPoints = 1
    m_ViewportRefReady = False
    m_FinalOutlineIsReady = False
    m_IsMaskReady = False
    If (m_NumOfLassoPoints < 2) Then Set m_FinalOutline = Nothing
    
    'Return the new cursor coordinates at this position, then exit
    newCursorX_ImgCoords = m_LassoPoints(m_NumOfLassoPoints - 1).x
    newCursorY_ImgCoords = m_LassoPoints(m_NumOfLassoPoints - 1).y
    
    'Create a safe copy of the current point collection
    BackupCurrentSelectionPoints

End Sub

'Update this selection using the values in the main form's selection text boxes
Friend Sub UpdateViaTextBox(Optional ByVal indexOfSrcControl As Long = -1)

    'Ignore text box update requests until the selection is locked in
    If (Not Me.IsLockedIn) Then Exit Sub
    
    If m_RejectRefreshRequests Then Exit Sub
    m_RejectRefreshRequests = True
    
    Dim subpanelOffset As Long, subpanelCtlOffset As Long
    subpanelOffset = Selections.GetSelectionSubPanelFromSelectionShape(PDImages.GetActiveImage())
    subpanelCtlOffset = subpanelOffset * 2
    
    'Check all text box entries for validity, then update the corresponding selection values.
    Select Case m_SelectionShape
    
        'Rectangles are currently getting experimental support for new features!
        Case ss_Rectangle, ss_Circle
            m_IsMaskReady = False
            
            'Modify position
            If (toolpanel_Selections.cboSize(subpanelOffset).ListIndex = 0) Then
                If toolpanel_Selections.tudSel(subpanelCtlOffset + 0).IsValid(False) Then m_CornersLocked.Left = toolpanel_Selections.tudSel(subpanelCtlOffset + 0)
                If toolpanel_Selections.tudSel(subpanelCtlOffset + 1).IsValid(False) Then m_CornersLocked.Top = toolpanel_Selections.tudSel(subpanelCtlOffset + 1)
                
            'Modify size
            ElseIf (toolpanel_Selections.cboSize(subpanelOffset).ListIndex = 1) Then
            
                'If aspect ratio is locked, we must cater to it during changes
                If m_IsAspectLocked And (m_LockedAspectRatio > 0) Then
                    
                    'User is changing width
                    If ((indexOfSrcControl And 1) = 0) Then
                        If toolpanel_Selections.tudSel(subpanelCtlOffset + 0).IsValid(False) Then
                            m_CornersLocked.Width = toolpanel_Selections.tudSel(subpanelCtlOffset + 0)
                            m_CornersLocked.Height = m_CornersLocked.Width * (1# / m_LockedAspectRatio)
                            toolpanel_Selections.tudSel(subpanelCtlOffset + 1) = m_CornersLocked.Height
                        End If
                        
                    'User is changing height
                    Else
                        If toolpanel_Selections.tudSel(subpanelCtlOffset + 1).IsValid(False) Then
                            m_CornersLocked.Height = toolpanel_Selections.tudSel(subpanelCtlOffset + 1)
                            m_CornersLocked.Width = m_CornersLocked.Height * m_LockedAspectRatio
                            toolpanel_Selections.tudSel(subpanelCtlOffset + 0) = m_CornersLocked.Width
                        End If
                    End If
                
                'If aspect ratio is *not* locked, freely modify either value, and cache any locked values so they
                ' can be used elsewhere.
                Else
                    If toolpanel_Selections.tudSel(subpanelCtlOffset + 0).IsValid(False) Then m_CornersLocked.Width = toolpanel_Selections.tudSel(subpanelCtlOffset + 0)
                    If toolpanel_Selections.tudSel(subpanelCtlOffset + 1).IsValid(False) Then m_CornersLocked.Height = toolpanel_Selections.tudSel(subpanelCtlOffset + 1)
                    If m_IsWidthLocked Then m_LockedWidth = m_CornersLocked.Width
                    If m_IsHeightLocked Then m_LockedHeight = m_CornersLocked.Height
                End If
            
            'Modify aspect ratio
            ElseIf (toolpanel_Selections.cboSize(subpanelOffset).ListIndex = 2) Then
                
                'Because aspect ratio calculations involve division, ensure validity before continuing
                Dim aspW As Double, aspH As Double, aspFinal As Double
                If toolpanel_Selections.tudSel(subpanelCtlOffset + 0).IsValid(False) Then aspW = toolpanel_Selections.tudSel(subpanelCtlOffset + 0)
                If toolpanel_Selections.tudSel(subpanelCtlOffset + 1).IsValid(False) Then aspH = toolpanel_Selections.tudSel(subpanelCtlOffset + 1)
                If (aspW > 0!) And (aspH > 0!) Then
                    
                    'We preferentially change the width of the selection to match the new aspect ratio if...
                    ' 1) The first aspect ratio parameter is changing (e.g. the 4 in 4:3)
                    ' 2) The selection's width is *not* locked
                    ' 3) The selection's height *is* locked
                    '
                    'Otherwise, we change height preferentially.
                    If (((indexOfSrcControl And 1) = 0) And (Not m_IsWidthLocked)) Or m_IsHeightLocked Then
                        aspFinal = aspW / aspH
                        m_CornersLocked.Width = m_CornersLocked.Height * aspFinal
                    Else
                        aspFinal = aspH / aspW
                        m_CornersLocked.Height = m_CornersLocked.Width * aspFinal
                    End If
                    
                    'Calculate aspect ratio and cache it; we need it in a variety of places
                    If (m_CornersLocked.Width <> 0!) And (m_CornersLocked.Height <> 0!) Then m_LockedAspectRatio = m_CornersLocked.Width / m_CornersLocked.Height
                    
                End If
                
            End If
            
        'I haven't decided if other selection types will support movement via text box...
        Case Else
        
    End Select
    
    'For some selection types, we need to update more than just the m_CornersLocked.Left/Top/Width/Height values.
    Select Case m_SelectionShape
    
        'Adjust the x1, y1, x2, y2 values to match any changes made via text box values
        Case ss_Rectangle, ss_Circle
            m_CornersUnlocked.Left = m_CornersLocked.Left
            m_CornersUnlocked.Top = m_CornersLocked.Top
            m_CornersUnlocked.Right = m_CornersLocked.Left + m_CornersLocked.Width
            m_CornersUnlocked.Bottom = m_CornersLocked.Top + m_CornersLocked.Height
        
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
        
        'Create a safe copy of the current point collection
        BackupCurrentSelectionPoints
        
        'The final thing we need to check for is the width and height, which may be still be zero at this point.
        ' Due to the way outside filters and effects use selection bounding rects, we can't allow selections of size 0.
        If (m_CornersLocked.Width < 1) Then m_CornersLocked.Width = 1
        If (m_CornersLocked.Height < 1) Then m_CornersLocked.Height = 1
        
        'If this selection has properties that are not applied during UI interactions (like feathering,
        ' which requires expensive gaussian blur operations), we need to apply them now.
        If (Me.GetSelectionProperty_Long(sp_Smoothing) = es_FullyFeathered) Then
            If (Me.GetSelectionProperty_Long(sp_FeatheringRadius) > 0) Then
                m_IsMaskReady = False
            End If
        End If
        
    End If
        
End Sub

'"Unlock" a selection
Friend Sub LockRelease()
    m_IsLocked = False
End Sub

'When a selection is erased, we can free up some internal tracking metrics.  If we don't do this, immediately starting a
' new selection of identical type (e.g. removing a polygon selection, starting a polygon selection) may result in some
' bits of the previous selection getting reused.
Friend Sub EraseCustomTrackers()
    
    m_NumOfLassoPoints = 0
    m_NumOfPolygonPoints = 0
    m_PolygonClosed = False
    m_LassoClosed = False
    
    m_FinalOutlineIsReady = False
    Set m_FinalOutline = Nothing
    
    m_OverlayIsReady = False
    m_LastOverlayCache = vbNullString
    If (Not m_ViewportOverlay Is Nothing) Then m_ViewportOverlay.ResetDIB 0
    
    m_IsTransformable = False
    m_IsMaskReady = False
    
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

    'Debug.Print "(" & Timer & ") Creating mask now..."
    
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
    
    'Whenever a new mask is created, any viewport-specific overlays and outlines must be re-created to match
    m_ViewportRefReady = False
    m_FinalOutlineIsReady = False
    m_OverlayIsReady = False
    
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
    
    'Before proceeding, we need to establish mask rendering colors.  PD supports both "interior" and "exterior" selections (and some
    ' tools support a hybrid, called "bordered" selections), and we use identical code to render such selections.  The only difference
    ' is the back and forecolors used; these colors are *inverted* when rendering exterior selections, which makes mask generation
    ' extremely simple.
    
    '(Note also that alpha bytes are generally ignored in mask operations; it is mask *color* that matters.)
    Dim maskBackColor As Long, maskForeColor As Long, maskBackOpacity As Byte, maskForeOpacity As Byte
    
    'Interior and exterior selections are rendered using identical code; the only difference is the colors used
    If (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Then
        maskBackColor = RGB(255, 255, 255)
        maskBackOpacity = 255
        maskForeColor = RGB(0, 0, 0)
        maskForeOpacity = 0
    Else
        maskBackColor = RGB(0, 0, 0)
        maskBackOpacity = 0
        maskForeColor = RGB(255, 255, 255)
        maskForeOpacity = 255
    End If
    
    'Start by creating a blank mask (this will also erase any existing mask)
    If ((m_SelMask.GetDIBWidth <> m_parentPDImage.Width) Or (m_SelMask.GetDIBHeight <> m_parentPDImage.Height)) Then
        m_SelMask.CreateBlank m_parentPDImage.Width, m_parentPDImage.Height, 32, maskBackColor, maskBackOpacity
        m_SelMask.SetInitialAlphaPremultiplicationState True
    Else
        m_SelMask.ResetDIB maskBackOpacity
    End If
    
    'We use pd2D for as much of the actual selection rendering as we can.  (The plan is still to migrate this code
    ' to dedicated per-shape child classes, but for now, this gives us a clean way to standardize rendering.)
    Dim dstSurface As pd2DSurface
    Set dstSurface = New pd2DSurface
    
    Dim surfaceOK As Boolean
    surfaceOK = dstSurface.WrapSurfaceAroundPDDIB(m_SelMask)
    
    'Set some surface properties in advance, like antialiasing and compositing.
    Dim useAA As Boolean
    
    If (Me.GetSelectionProperty_Long(sp_Smoothing) > es_None) Then
        
        'Rectangle selections are an exception here; we don't want to antialias them at all, even if antialiasing is enabled.
        ' (GDI+ tends to mess up antialiasing for perfect rects.)
        useAA = Not ((m_SelectionShape = ss_Rectangle) And (Me.GetSelectionProperty_Float(sp_RoundedCornerRadius) = 0#))
    Else
        useAA = False
    End If
    
    'Some setting combinations result in us *not* using antialiasing.  Alongside this, we must also manually
    ' clamp all coordinates to integer boundaries or GDI+ may still output antialiased-like results.
    Dim fixedCornersLocked As RectF
    fixedCornersLocked = m_CornersLocked
    
    If useAA Then
        dstSurface.SetSurfaceAntialiasing P2_AA_HighQuality
    Else
        
        dstSurface.SetSurfaceAntialiasing P2_AA_None
        
        'If we're *not* using antialiasing, clip the selection to fixed integer boundaries.
        ' (Note that this only applies to rectangular and elliptical selection shapes.)
        If (Me.GetSelectionShape = ss_Circle) Or (Me.GetSelectionShape = ss_Rectangle) Then
        
            Dim lFrac As Single, tFrac As Single
            lFrac = PDMath.Frac(m_CornersLocked.Left)
            tFrac = PDMath.Frac(m_CornersLocked.Top)
            With fixedCornersLocked
                .Left = Int(.Left)
                .Top = Int(.Top)
                .Width = Int(.Width + lFrac)
                .Height = Int(.Height + tFrac)
            End With
            
        End If
            
    End If
    
    dstSurface.SetSurfaceCompositing P2_CM_Overwrite
    
    If (Not surfaceOK) Then PDDebug.LogAction "WARNING!  pdSelection.CreateSelectionMask() failed to wrap the destination mask surface."
    
    'Interior/exterior selections use brushes to fill the appropriate region; bordered selections use a pen to *stroke*
    ' the selection outline.  (For bordered selections, note that we also generate a matching pd2DPath object, which is
    ' then used to determine selection boundaries (which vary due to pen mitre settings).)
    Dim backBrush As pd2DBrush, foreBrush As pd2DBrush
    Drawing2D.QuickCreateSolidBrush backBrush, maskBackColor, CSng(maskBackOpacity) / 2.55!
    Drawing2D.QuickCreateSolidBrush foreBrush, maskForeColor, CSng(maskForeOpacity) / 2.55!
    
    Dim maskPen As pd2DPen, maskPath As pd2DPath
    If (Me.GetSelectionProperty_Long(sp_Area) = sa_Border) Then
                
        'Border-type selections need to calculate a border size in advance.  If the user's specified border size is too large
        ' (e.g. it is larger than the selection's width or height), we will manually rein it in.
        Dim actualBorderSize As Single
        If (GetSelectionProperty_Long(sp_Area) = sa_Border) Then
            actualBorderSize = PDMath.Max2Float_Single(GetSelectionProperty_Long(sp_BorderWidth), 1#)
            
            If (Me.GetSelectionShape = ss_Rectangle) Or (Me.GetSelectionShape = ss_Circle) Then
                If (actualBorderSize > PDMath.Min2Float_Single(m_CornersLocked.Width + 1#, m_CornersLocked.Height + 1#)) Then actualBorderSize = PDMath.Min2Float_Single(m_CornersLocked.Width + 1#, m_CornersLocked.Height + 1#)
            End If
            
        End If
        
        'Note that we technically would want to use the specified foreground opacity, but because GDI+ won't give us correct
        ' antialiasing with a fully transparent brush, we rely on an opaque brush instead.
        If (Me.GetSelectionShape = ss_Rectangle) Then
            Drawing2D.QuickCreateSolidPen maskPen, actualBorderSize, maskForeColor, 100#, P2_LJ_Miter, P2_LC_Flat
            maskPen.SetPenMiterLimit 10#
        Else
            Drawing2D.QuickCreateSolidPen maskPen, actualBorderSize, maskForeColor, 100#, P2_LJ_Round, P2_LC_Flat
        End If
        
        Set maskPath = New pd2DPath
        
    End If
    
    'The actual rendering of the selection will vary based on the current selection type (obviously).
    ' TODO: move all rendering into a separate child class.  This is necessary for union/intersect/etc with multiple selections.
    Select Case m_SelectionShape
    
        Case ss_Rectangle
        
            'RECTANGLE SELECTION, NO ROUNDED CORNERS
            If (Me.GetSelectionProperty_Float(sp_RoundedCornerRadius) = 0#) Then
            
                If (Me.GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (Me.GetSelectionProperty_Long(sp_Area) = sa_Exterior) Then
                    PD2D.FillRectangleF_FromRectF dstSurface, foreBrush, fixedCornersLocked
                Else
                    maskPath.AddRectangle_RectF m_CornersLocked
                    PD2D.DrawRectangleF_FromRectF dstSurface, maskPen, fixedCornersLocked
                End If
            
            'RECTANGLE SELECTION *WITH* ROUNDED CORNERS
            Else
                
                'The round-rectangle radius property is stored as a ratio on the scale [0, 100.0].  We want to scale this value to
                ' an absolute value, based on the current size of the rectangle.  (Where 100.0 = maximum corner curvature.)
                Dim actualRoundedRadius As Double
                actualRoundedRadius = PDMath.Min2Float_Single(fixedCornersLocked.Width, fixedCornersLocked.Height) * (GetSelectionProperty_Float(sp_RoundedCornerRadius) * 0.01)
                
                dstSurface.SetSurfacePixelOffset P2_PO_Half
                
                If (GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Then
                    PD2D.FillRoundRectangleF_FromRectF dstSurface, foreBrush, fixedCornersLocked, actualRoundedRadius
                Else
                    maskPath.AddRoundedRectangle_RectF m_CornersLocked, actualRoundedRadius
                    PD2D.DrawRoundRectangleF_FromRectF dstSurface, maskPen, fixedCornersLocked, actualRoundedRadius
                End If
                
            End If
                    
        'CIRCLES / ELLIPSES
        Case ss_Circle
        
            'Note that ellipses require us to manually reduce the specified rect by 1px on the right and bottom border; I'm not sure
            ' why GDI+ renders this way, but it's possible that they did it to make it semi-"backwards compatible" with GDI methods
            ' (which treat the bottom-right point as exclusive).
            Dim newEllipseWidth As Single, newEllipseHeight As Single
            newEllipseWidth = PDMath.Max2Float_Single(fixedCornersLocked.Width - 1#, 1#)
            newEllipseHeight = PDMath.Max2Float_Single(fixedCornersLocked.Height - 1#, 1#)
                
            If (GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Then
                PD2D.FillEllipseF dstSurface, foreBrush, fixedCornersLocked.Left, fixedCornersLocked.Top, newEllipseWidth, newEllipseHeight
            Else
                maskPath.AddEllipse_Relative m_CornersLocked.Left, m_CornersLocked.Top, newEllipseWidth, newEllipseHeight
                PD2D.DrawEllipseF dstSurface, maskPen, fixedCornersLocked.Left, fixedCornersLocked.Top, newEllipseWidth, newEllipseHeight
            End If
            
        
        'Polygon selections are easy - simply close the current point collection, then fill (or stroke) it via GDI+.
        Case ss_Polygon
            
            'Don't create a polygon mask until at least *3* points exist!
            If (m_NumOfPolygonPoints > 2) Then
                
                'Start by storing the current polygon in a path object
                Dim polyPath As pd2DPath
                Set polyPath = New pd2DPath
                polyPath.SetFillRule P2_FR_Winding
                polyPath.AddPolygon m_NumOfPolygonPoints, VarPtr(m_PolygonPoints(0)), True, (GetSelectionProperty_Float(sp_PolygonCurvature) > 0#), GetSelectionProperty_Float(sp_PolygonCurvature)
                
                If ((GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (GetSelectionProperty_Long(sp_Area) = sa_Exterior)) Then
                    PD2D.FillPath dstSurface, foreBrush, polyPath
                    m_Bounds = polyPath.GetPathBoundariesF()
                Else
                    PD2D.DrawPath dstSurface, maskPen, polyPath
                    m_Bounds = polyPath.GetPathBoundariesF(maskPen)
                End If
                
            End If
            
            'To ensure clipping is correct, make sure our corner coordinates match our boundary coordinates
            m_CornersLocked = m_Bounds
            
        'Strangely enough, lasso selections are arguably the simplest selection type to render, as we simply close the lasso shape,
        ' then fill it as if it represents an arbitrary region.
        Case ss_Lasso
            
            'Don't create a lasso mask until at least *2* points exist!
            If (m_NumOfLassoPoints > 1) Then
            
                'Start by storing the current polygon in a path object
                Dim lassoPath As pd2DPath
                Set lassoPath = New pd2DPath
                lassoPath.SetFillRule P2_FR_Winding
                lassoPath.AddPolygon m_NumOfLassoPoints, VarPtr(m_LassoPoints(0)), True, (GetSelectionProperty_Float(sp_SmoothStroke) > 0#), GetSelectionProperty_Float(sp_SmoothStroke)
                
                If (GetSelectionProperty_Long(sp_Area) = sa_Interior) Or (GetSelectionProperty_Long(sp_Area) = sa_Exterior) Then
                    PD2D.FillPath dstSurface, foreBrush, lassoPath
                    m_Bounds = lassoPath.GetPathBoundariesF()
                Else
                    PD2D.DrawPath dstSurface, maskPen, lassoPath
                    m_Bounds = lassoPath.GetPathBoundariesF(maskPen)
                End If
            
            End If
            
            'To ensure clipping is correct, make sure our corner coordinates match our boundary coordinates
            m_CornersLocked = m_Bounds
            
        'Wand selections use a custom class to perform a flood fill
        Case ss_Wand
        
            'Based on the flood fill type (layer vs image), pass a different source layer to the flood fill class.
            ' Note that we try to only regenerate this image as absolutely necessary.
            Dim wandImageRefreshRequired As Boolean
            wandImageRefreshRequired = (m_WandImage Is Nothing)
            If (Not wandImageRefreshRequired) Then wandImageRefreshRequired = (m_WandImageTimestamp <> m_parentPDImage.GetTimeOfLastChange())
            
            If wandImageRefreshRequired Then
                
                'A new image copy is required.  As much as possible, still try to minimize the work we do here
                If (m_WandImage Is Nothing) Then Set m_WandImage = New pdDIB
                If (m_WandImage.GetDIBWidth <> m_parentPDImage.Width) Or (m_WandImage.GetDIBHeight <> m_parentPDImage.Height) Then
                    m_WandImage.CreateBlank m_parentPDImage.Width, m_parentPDImage.Height, 32, 0, 0
                Else
                    m_WandImage.ResetDIB 0
                End If
                
                'Note the timestamp; this may allow us to skip subsequent copy requests
                m_WandImageTimestamp = m_parentPDImage.GetTimeOfLastChange()
                
                'Merged image data is far more cumbersome to generate
                If (GetSelectionProperty_Long(sp_WandSampleMerged) = 0) Then
                    m_parentPDImage.GetCompositedImage m_WandImage
                    
                Else
                    
                    Dim tmpLayer As pdLayer
                    Set tmpLayer = New pdLayer
                    tmpLayer.CopyExistingLayer m_parentPDImage.GetActiveLayer
                    tmpLayer.ConvertToNullPaddedLayer m_parentPDImage.Width, m_parentPDImage.Height
                    Set m_WandImage = tmpLayer.layerDIB
                    Set tmpLayer = Nothing
                    
                End If
                
            End If
            
            'pdFloodFill handles most the heavy lifting here
            If (m_FloodFill Is Nothing) Then Set m_FloodFill = New pdFloodFill
            
            'Set all initial parameters
            m_FloodFill.SetAntialiasingMode (GetSelectionProperty_Long(sp_Smoothing) > es_None)
            m_FloodFill.SetInitialPoint m_CornersUnlocked.Left, m_CornersUnlocked.Top
            m_FloodFill.SetTolerance GetSelectionProperty_Float(sp_WandTolerance)
            m_FloodFill.SetSearchMode GetSelectionProperty_Long(sp_WandSearchMode)
            m_FloodFill.SetCompareMode GetSelectionProperty_Long(sp_WandCompareMethod)
            
            'Apply the flood fill, and make sure we retrieve an outline while we're at it
            If (m_WandOutline Is Nothing) Then Set m_WandOutline = New pd2DPath Else m_WandOutline.ResetPath
            m_FloodFill.InitiateFloodFill m_WandImage, m_SelMask, m_WandOutline
                        
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
            
            'If feathering is in play, we'll account for it programmatically.
            Dim needToAddFeathering As Boolean
            needToAddFeathering = (GetSelectionProperty_Long(sp_Smoothing) = es_FullyFeathered)
            
            'Next, we sort bound calculations by selection shape
            Select Case m_SelectionShape
            
                'Rectangle, ellipse, and lasso selections are easy; bounds have already been set by the width/height values
                Case ss_Rectangle, ss_Circle
                
                    If (GetSelectionProperty_Long(sp_Area) <> sa_Border) Then
                        m_Bounds = m_CornersLocked
                    Else
                        
                        'Retrieve boundaries directly from the relevant path object
                        m_Bounds = maskPath.GetPathBoundariesF(maskPen)
                        
                        'When a selection is first created, the clicked point may result in a selection of size [0x0] which
                        ' will fail to generate a path, and thus the boundary check will fail.  Check for this potential case
                        ' and handle it manually.
                        If (m_Bounds.Width = 0#) And (m_Bounds.Height = 0#) Then m_Bounds = m_CornersLocked
                        
                    End If
                    
                'Polygon boundaries are already calculated in a previous step
                Case ss_Polygon
                    
                'Lasso boundaries are already calculated in a previous step
                Case ss_Lasso
                
                'Magic wand requires manual bounds-finding
                Case ss_Wand
                    Me.FindNewBoundsManually True
                    
                'Other shapes currently rely on manual bounds-checking, using the rendered mask as the guide
                Case Else
                    Me.FindNewBoundsManually
                    needToAddFeathering = False
                    
            End Select
            
            'Finally, add the current feathering size, if any, to the selection boundaries
            If needToAddFeathering Then
                With m_Bounds
                    .Left = m_Bounds.Left - GetSelectionProperty_Long(sp_FeatheringRadius)
                    .Top = m_Bounds.Top - GetSelectionProperty_Long(sp_FeatheringRadius)
                    .Width = m_Bounds.Width + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                    .Height = m_Bounds.Height + GetSelectionProperty_Long(sp_FeatheringRadius) * 2
                End With
            End If
                
    End Select
    
    'Do some basic bounds checking on the bound values to make sure they lie inside the image.  This is important because the
    ' selection mask (and any code that operates on it) assumes a match to image boundaries, despite the fact that selection
    ' points can actually lie anywhere on the canvas - even outside the image!
    FixBoundsToImageSize
    
    'Finally, if the selection is locked and feathering has been requested, apply it now.
    ' (We only apply feathering when locked-in, as the performance penalty is severe.)
    If m_IsLocked And (GetSelectionProperty_Long(sp_Smoothing) = es_FullyFeathered) And (GetSelectionProperty_Long(sp_FeatheringRadius) > 0) Then ApplyFeatheringToMask
    
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
    Filters_Layers.QuickBlurDIBRegion m_SelMask, GetSelectionProperty_Long(sp_FeatheringRadius), m_Bounds
End Sub

'Helper function for retrieving the current selection corners (*not* boundary corners; this means things like feathering are
' *not* taken into account, by design).
'
'The destination array must be one-dimensional, zero-dimensioned, and at least four items long (e.g. [0, 3] or larger).
Private Sub GetSelectionCorners(ByRef dstPoints() As PointFloat)
    
    With m_CornersLocked
        dstPoints(0).x = .Left
        dstPoints(0).y = .Top
        dstPoints(1).x = .Left + .Width
        dstPoints(1).y = .Top
        dstPoints(2).x = .Left
        dstPoints(2).y = .Top + .Height
        dstPoints(3).x = .Left + .Width
        dstPoints(3).y = .Top + .Height
    End With
    
End Sub

'While a selection tool is active, we draw transform nodes around it.  The viewport renderer invokes this function as necessary.
Friend Sub RenderTransformNodes(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, ByVal curToolID As PDTools)
    
    'If an update has been requested, but we are already in the middle of one, deny subsequent requests
    If m_RejectRefreshRequests Then Exit Sub
    
    'Transform nodes are only rendered if the current selection shape matches the current selection tool.  (Otherwise, the
    ' selection is treated as "static", and only the default outline/highlight/whatever is drawn.)
    If (Selections.GetRelevantToolFromSelectShape() = curToolID) Then
        
        'Before drawing the nodes, we need to convert any relevant selection coordinates from "image coordinates" to "viewport coordinates".
        ' (Remember that the selection's location is stored as coordinates relative to the IMAGE ITSELF - but to render the transform nodes
        '   on the screen, we need to figure out where they lie in the current viewport, accounting for things like zoom and scroll.)
        '
        'We always generate this list of points, regardless of selection type.
        Dim selCorners() As PointFloat
        ReDim selCorners(0 To 3) As PointFloat
        GetSelectionCorners selCorners
        
        'Transform those corners into the destination viewport coordinate space
        Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, selCorners, False
        
        'Convert the current point of interest value, if any, to an index into our list of selection corners
        Dim curPOI As PD_PointOfInterest
        curPOI = m_CurrentPOI
        
        Dim i As Long
        Dim tmpX As Double, tmpY As Double
        
        Dim circRadius As Single
        circRadius = 7!
            
        'pd2D is used for all rendering.  Note that a "highlight" pen is only created if a POI is currently in use
        Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True, dstCanvas.hWnd
        
        Dim cPenBaseNormal As pd2DPen, cPenTopNormal As pd2DPen, cPenBaseHighlight As pd2DPen, cPenTopHighlight As pd2DPen
        Drawing2D.QuickCreatePairOfUIPens cPenBaseNormal, cPenTopNormal, False
        If (curPOI <> poi_Undefined) Then Drawing2D.QuickCreatePairOfUIPens cPenBaseHighlight, cPenTopHighlight, True
        
        Dim cBrushFill As pd2DBrush
        
        'If the current selection supports "resize by corner node", we want to handle that case now.
        If Me.DoesShapeSupportCornerResize() Then
        
            'Convert the current POI, if any, to its matching position in the "corner node point array"
            Dim cornerPOI As PD_PointOfInterest
            
            If (curPOI <> poi_Undefined) Then
                If (curPOI = poi_CornerNW) Then
                    cornerPOI = 0
                ElseIf (curPOI = poi_CornerNE) Then
                    cornerPOI = 1
                ElseIf (curPOI = poi_CornerSW) Then
                    cornerPOI = 2
                ElseIf (curPOI = poi_CornerSE) Then
                    cornerPOI = 3
                Else
                    cornerPOI = poi_Undefined
                End If
            Else
                cornerPOI = poi_Undefined
            End If
            
            'Render the corner transform nodes, with the currently interactive node (if any) highlighted to match
            Dim cornerSize As Single, halfCornerSize As Single
            cornerSize = 12#
            halfCornerSize = cornerSize * 0.5
            
            'In keeping with convention, corner nodes are rendered as squares centered over each node
            For i = 0 To 3
                If (i = cornerPOI) Then
                    PD2D.DrawRectangleF cSurface, cPenBaseHighlight, selCorners(i).x - halfCornerSize, selCorners(i).y - halfCornerSize, cornerSize, cornerSize
                    PD2D.DrawRectangleF cSurface, cPenTopHighlight, selCorners(i).x - halfCornerSize, selCorners(i).y - halfCornerSize, cornerSize, cornerSize
                Else
                    PD2D.DrawRectangleF cSurface, cPenBaseNormal, selCorners(i).x - halfCornerSize, selCorners(i).y - halfCornerSize, cornerSize, cornerSize
                    PD2D.DrawRectangleF cSurface, cPenTopNormal, selCorners(i).x - halfCornerSize, selCorners(i).y - halfCornerSize, cornerSize, cornerSize
                End If
            Next i
            
        End If
        
        'If the current selection also supports custom points of interest, let's draw those next
        If Me.DoesShapeSupportCustomPOIs() Then
            
            'Some shapes have been migrated to a nice, uniform set of POI handlers.  These can be handled using universal code.
            If (m_SelectionShape = ss_Wand) Then
                
                Dim poiList() As PointFloat
                Me.GetCurrentPOIList poiList
                
                Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, poiList, False
                
                For i = 0 To UBound(poiList)
                    If (i = curPOI) Then
                        PD2D.DrawCircleF cSurface, cPenBaseHighlight, poiList(i).x, poiList(i).y, circRadius
                        PD2D.DrawCircleF cSurface, cPenTopHighlight, poiList(i).x, poiList(i).y, circRadius
                    Else
                        PD2D.DrawCircleF cSurface, cPenBaseNormal, poiList(i).x, poiList(i).y, circRadius
                        PD2D.DrawCircleF cSurface, cPenTopNormal, poiList(i).x, poiList(i).y, circRadius
                    End If
                Next i
            
            'Legacy rendering path for outlier selection types follows
            Else
                
                Dim polyX As Double, polyY As Double
                
                If (m_SelectionShape = ss_Polygon) Then
                    
                    For i = 0 To m_NumOfPolygonPoints - 1
                        
                        Drawing.ConvertImageCoordsToCanvasCoords dstCanvas, srcImage, m_PolygonPoints(i).x, m_PolygonPoints(i).y, polyX, polyY
                        
                        'If this is the initial point in a polygon selection, fill it to help orient the user
                        If (i = 0) And (Not m_PolygonClosed) Then
                            Drawing2D.QuickCreateSolidBrush cBrushFill, g_Themer.GetGenericUIColor(UI_Accent), 67!
                            PD2D.FillCircleF cSurface, cBrushFill, polyX, polyY, circRadius
                        End If
                        
                        If (i = curPOI) Then
                            PD2D.DrawCircleF cSurface, cPenBaseHighlight, polyX, polyY, circRadius
                            PD2D.DrawCircleF cSurface, cPenTopHighlight, polyX, polyY, circRadius
                        Else
                            PD2D.DrawCircleF cSurface, cPenBaseNormal, polyX, polyY, circRadius
                            PD2D.DrawCircleF cSurface, cPenTopNormal, polyX, polyY, circRadius
                        End If
                        
                    Next i
                    
                    'After drawing all polygon points, we lastly want to draw a dotted line to
                    ' the current mouse position.
                    If dstCanvas.IsMouseOverCanvas And (m_NumOfPolygonPoints > 0) And (Not Me.GetPolygonClosedState()) Then
                        
                        'Remove antialiasing to match the default polygon renderer
                        cSurface.SetSurfaceAntialiasing P2_AA_None
                        
                        Dim canvasMouseX As Long, canvasMouseY As Long
                        canvasMouseX = dstCanvas.GetLastMouseX()
                        canvasMouseY = dstCanvas.GetLastMouseY()
                        
                        Dim cBlackPen As pd2DPen, cDottedPen As pd2DPen
                        Set cBlackPen = New pd2DPen
                        cBlackPen.SetPenColor vbBlack
                        cBlackPen.SetPenWidth 1!
                        
                        Set cDottedPen = New pd2DPen
                        cDottedPen.SetPenWidth 1!
                        cDottedPen.SetPenColor vbWhite
                        cDottedPen.CreatePen
                        cDottedPen.SetPenStyle P2_DS_Custom
                        cDottedPen.SetPenDashes_UNSAFE VarPtr(m_AntDashes(0)), 2
                        cDottedPen.SetPenDashOffset m_AntDashOffset
                        
                        PD2D.DrawLineF cSurface, cBlackPen, polyX, polyY, canvasMouseX, canvasMouseY
                        PD2D.DrawLineF cSurface, cDottedPen, polyX, polyY, canvasMouseX, canvasMouseY
                        
                        'Note that - by design - the standard selection renderer does not render
                        ' polygon lines until there are at least 3 points in the selection, as that's
                        ' the minimum required for a valid polygon (which the underlying renderer
                        ' requires as it actually renders the polygon, then scans boundaries).
                        
                        'To make the transition from 2 points to 3 point less jarring, we manually
                        ' render the first line segment when there are just 2 points in the polygon.
                        If (m_NumOfPolygonPoints = 2) Then
                            Drawing.ConvertImageCoordsToCanvasCoords dstCanvas, srcImage, m_PolygonPoints(0).x, m_PolygonPoints(0).y, tmpX, tmpY
                            PD2D.DrawLineF cSurface, cBlackPen, polyX, polyY, tmpX, tmpY
                            PD2D.DrawLineF cSurface, cDottedPen, polyX, polyY, tmpX, tmpY
                        End If
                        
                    End If
                    
                End If
            
            End If
            
        End If
        
        'Finally, for lasso selections, if the lasso is currently being drawn, we want to highlight
        ' the starting point of the lasso with a circle.
        If (m_SelectionShape = ss_Lasso) And (Not Me.GetLassoClosedState()) And (m_NumOfLassoPoints > 0) Then
            Drawing.ConvertImageCoordsToCanvasCoords dstCanvas, srcImage, m_LassoPoints(0).x, m_LassoPoints(0).y, tmpX, tmpY
            
            'If this is the initial point in a polygon selection, fill it to help orient the user
            If (i = 0) And (Not m_PolygonClosed) Then
                Drawing2D.QuickCreateSolidBrush cBrushFill, g_Themer.GetGenericUIColor(UI_Accent), 67!
                PD2D.FillCircleF cSurface, cBrushFill, tmpX, tmpY, circRadius
            End If
            
            PD2D.DrawCircleF cSurface, cPenBaseNormal, tmpX, tmpY, circRadius
            PD2D.DrawCircleF cSurface, cPenTopNormal, tmpX, tmpY, circRadius
            
        End If
        
        'We don't need to manually free these pd2D objects, obviously, but I've added the code here just in case
        ' this function gets expanded in the future.
        Set cSurface = Nothing
        Set cPenBaseNormal = Nothing: Set cPenTopNormal = Nothing
        Set cPenBaseHighlight = Nothing: Set cPenTopHighlight = Nothing
        
    End If
    
End Sub

'The selection engine caches a lot of local objects in an effort to improve performance.  To detect setting changes,
' it's convenient to generate arbitrary string hashes between "current settings" and "settings used at last cache generation."
' (Note that this function requires that all passed variants can be coerced into strings.)
Private Function GenerateArbitraryHash(ParamArray srcStuff() As Variant) As String
    
    If (UBound(srcStuff) >= LBound(srcStuff)) Then
    
        Dim i As Long
        For i = LBound(srcStuff) To UBound(srcStuff)
            GenerateArbitraryHash = GenerateArbitraryHash & CStr(srcStuff(i))
        Next i
    
    Else
        GenerateArbitraryHash = vbNullString
    End If
    
End Function

'Render the current selection mask using one of several methods.  Required inputs include:
' 1) Destination DIB, assumed to be the front buffer (e.g. the full composite image has already been rendered to this surface, in its
'     proper location, and any color management has already been applied to it)
' 2) Source image (we could technically pull from our local copy, but I'd like to remove that dependency in the future)
' 3) Destination canvas (which determines things like the current zoom and scroll position)
'
'Note that this function is the parent for *all* selection type rendering.  (After performing some required setup steps,
' it will automatically forward subsequent rendering to an appropriate child function, depending on the current UI settings.)
Friend Sub RenderSelectionToViewport(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas)
    
    'If an update has been requested, but we are already in the middle of an update, deny subsequent requests
    If m_RejectRefreshRequests Then Exit Sub
    
    'Is the underlying selection mask ready?  If not, prepare it first.  (This allows us to suspend selection mask creation
    ' until the last possible moment, reducing the potential of generating the mask more times than we actually require.)
    If (Not m_IsMaskReady) Then CreateSelectionMask
    
    'If the current selection lies fully off-image, skip rendering entirely.
    If Me.AreAllCoordinatesInvalid Then Exit Sub
    
    'Debug builds may want to measure time taken; I've currently commented-out reporting capabilities, but this function
    ' is fully wired for detailed reporting as future needs may require.
    Dim startTime As Currency
    
    'Regardless of what type of selection overlay we're rendering (highlight, outline, marching ants, etc), we now want to
    ' generate a copy of the selection mask, scaled and cropped to match the current viewport.  (Subsequent rendering will
    ' rely on this selection mask copy, instead of the full mask, which greatly accelerates rendering when the base image
    ' is very large but zoomed-out.)
    
    'Start by requesting the current canvas intersection rect from the source image
    Dim viewportIntersectRectCanvas As RectF
    srcImage.ImgViewport.GetIntersectRectCanvas viewportIntersectRectCanvas
    
    'We now need to calculate an integer-only version of the "viewport space" coordinates.  (The width and height of
    ' the overlay DIB are integer-only, obviously, and we can improve performance in our final render if we use
    ' integer coordinates there as well.)
    Dim imgViewportRectL As RectL_WH
    With viewportIntersectRectCanvas
        imgViewportRectL.Left = Int(.Left)
        imgViewportRectL.Top = Int(.Top)
        imgViewportRectL.Width = Int(PDMath.Frac(.Left) + .Width + 0.999999!)
        imgViewportRectL.Height = Int(PDMath.Frac(.Top) + .Height + 0.999999!)
    End With
    
    'Next, we need to convert the current selection boundary coordinates from "image coordinates" to "viewport coordinates".
    ' (Remember that the selection's location is stored as coordinates relative to the IMAGE ITSELF - but to render it on
    '  the screen, we need to translate the coordinates to the current viewport - this makes them display accurately when
    '  scroll and zoom are in effect.)
    Dim dstRectF As RectF
    Drawing.ConvertImageCoordsToCanvasCoords_RectF dstCanvas, srcImage, m_Bounds, dstRectF
    
    'EXPERIMENTAL!  Clamp the converted rect to integer coordinates, to improve rendering along pixel boundaries.
    ' (I still need to do something similar in the central viewport; hence why this is labeled as "experimental".)
    'PDMath.GetIntClampedRectF dstRectF
    
    'If we've already cached a viewport-specific mask copy on a previous run, we now want to perform two additional checks:
    ' one against the source image region, and another against the destination viewport region.  If either of these rects
    ' have changed since we generated our existing overlay copy, we need to create a new overlay reference (e.g. a new copy
    ' of the selection mask, pre-zoomed-and-translated-and-cropped to match the current viewport).
    '
    'This class will automatically detect internal changes that require a new viewport DIB (and thus set m_ViewportRefReady
    ' to FALSE), but viewport-specific changes are not detected until this function is actually invoked.
    If m_ViewportRefReady Then
        
        'First, compare current and previous viewport rects
        If (Not PDMath.AreRectFsEqual(m_LastViewportRectF, viewportIntersectRectCanvas)) Then
            m_LastViewportRectF = viewportIntersectRectCanvas
            m_ViewportRefReady = False
        End If
        
        'Next, compare current and previous image rects
        If m_ViewportRefReady Then
            
            Dim curImgRectF As RectF
            srcImage.ImgViewport.GetIntersectRectImage curImgRectF
            
            If (Not PDMath.AreRectFsEqual(curImgRectF, m_LastImageRectF)) Then
                m_LastImageRectF = curImgRectF
                m_ViewportRefReady = False
            End If
            
        End If
        
        'Next, compare previous and current rendering modes
        If (m_LastRenderMode <> Selections.GetSelectionRenderMode()) Then m_ViewportRefReady = False
        
    End If
    
    'If we regenerate our cached viewport image, we want to notify subsequent child functions, as they may need to
    ' regenerate their own caches against the new temporary image.
    Dim overlayWasRegenerated As Boolean: overlayWasRegenerated = False
    
    'If any of our "viewport cache is ready" checks failed, generate a new cache now
    If (Not m_ViewportRefReady) Then
        
        VBHacks.GetHighResTime startTime
        
        'Prep a target DIB; if at all possible, we want to reuse our existing DIB (as allocating new memory can be
        ' prohibitively expensive, given how frequently this function may be called).
        If (m_ViewportReference Is Nothing) Then Set m_ViewportReference = New pdDIB
        If (m_ViewportReference.GetDIBWidth <> imgViewportRectL.Width) Or (m_ViewportReference.GetDIBHeight <> imgViewportRectL.Height) Then
            m_ViewportReference.CreateBlank imgViewportRectL.Width, imgViewportRectL.Height, 32, 0, 0
            m_ViewportReference.SetInitialAlphaPremultiplicationState True
        Else
            m_ViewportReference.ResetDIB 0
        End If
        
        'Paint a resized+translated selection mask into our cached DIB.  Note that the interpolation method used varies
        ' by user setting; this is done to match the main viewport pipeline, and ensure accurate results when the viewport
        ' is greatly zoomed-out.
        Dim isZoomedIn As Boolean
        isZoomedIn = (Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex()) > 1#)
        
        Dim interpolationType As GP_InterpolationMode
        If isZoomedIn Then
            interpolationType = GP_IM_NearestNeighbor
        Else
            If (g_ViewportPerformance = PD_PERF_BALANCED) Then interpolationType = GP_IM_Bilinear Else interpolationType = GP_IM_HighQualityBicubic
        End If
        
        With m_ViewportRefRectF
            .Left = dstRectF.Left - viewportIntersectRectCanvas.Left
            .Top = dstRectF.Top - viewportIntersectRectCanvas.Top
            .Width = dstRectF.Width
            .Height = dstRectF.Height
            GDI_Plus.GDIPlus_StretchBlt m_ViewportReference, .Left, .Top, .Width, .Height, m_SelMask, m_Bounds.Left, m_Bounds.Top, m_Bounds.Width, m_Bounds.Height, , interpolationType, , , isZoomedIn, False  'True
        End With
        
        'Mark the finished viewport cache as "ready"; we won't regenerate it unless the underlying selection mask
        ' (or the current viewport position/zoom) changes
        m_ViewportRefReady = True
        
        'Also note that the overlay reference was regenerated; this means we have to generate a new UI DIB to match
        overlayWasRegenerated = True
        
        'Timing report:
        'pdDebug.LogAction "Selection viewport reference DIB generated in " & Format$(VBHacks.GetTimerDifferenceNow(startTime) * 1000, "#0") & " ms"
        
    End If
    
    'We have now covered all situations where the underlying mask reference may have changed.  (If it has,
    ' the local "overlayWasRegenerated" value will be set to TRUE.)
    
    'A few different criteria can cause us to silently switch between this function and RenderOutlineOnly (which, per its name,
    ' renders the active selection's outline by treating the underlying selection data as a vector path, not a raster mask).
    Dim renderType As PD_SelectionRender, renderColor As Long, renderOpacity As Long
    renderType = Selections.GetSelectionRenderMode()
    m_LastRenderMode = renderType
    
    If (renderType = PDSR_Highlight) Then
        renderColor = Selections.GetSelectionColor_Highlight()
        renderOpacity = Selections.GetSelectionOpacity_Highlight()
    Else
        renderColor = Selections.GetSelectionColor_Lightbox()
        renderOpacity = Selections.GetSelectionOpacity_Lightbox()
    End If
    
    'We use the RASTER OVERLAY FUNCTION if...
    ' 1) The user has requested "highlight" or "lightbox" rendering mode.
    
    'We use "outline"-style rendering if...
    ' 1) The user has requested outline-only rendering, or...
    ' 2) A lasso selection is active but not locked-in (meaning it's still being drawn), or...
    ' 3) A polygon selection is active but not locked-in (meaning it's still being drawn)
    Dim useOutlineRenderingInstead As Boolean
    useOutlineRenderingInstead = (renderType = PDSR_Outline) Or (renderType = PDSR_Ants)
    If (Not useOutlineRenderingInstead) Then useOutlineRenderingInstead = ((m_SelectionShape = ss_Lasso) And (Not m_LassoClosed))
    If (Not useOutlineRenderingInstead) Then useOutlineRenderingInstead = ((m_SelectionShape = ss_Polygon) And (Not m_PolygonClosed))
    
    'If we are not using marching ants, disable the associated ant timer now (as a failsafe; this is typically solved elsewhere,
    ' but better safe than sorry).
    If (renderType <> PDSR_Ants) Then
        
        'As a failsafe, forcibly disable animations prior to stopping the timer; this helps against any pending timer events
        ' firing as a result of stopping the timer
        Dim initAnimState As Boolean: initAnimState = m_AnimationsAllowed
        m_AnimationsAllowed = False
        m_AntTimer.StopTimer
        m_AnimationsAllowed = initAnimState
        
    End If
    
    If useOutlineRenderingInstead Then
        RenderSelectionAsOutline dstDIB, srcImage, dstCanvas, imgViewportRectL, overlayWasRegenerated, renderType
    Else
        RenderSelectionAsOverlay dstDIB, srcImage, dstCanvas, imgViewportRectL, overlayWasRegenerated, renderType, renderColor, renderOpacity
    End If
    
End Sub

'Child function of RenderSelectionToViewport(), above.  Renders the current selection mask to a target viewport,
' using the HIGHLIGHT or LIGHTBOX selection styles.
Private Sub RenderSelectionAsOverlay(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, ByRef imgViewportRectL As RectL_WH, ByRef overlayWasRegenerated As Boolean, ByVal renderType As PD_SelectionRender, Optional ByVal renderColor As Long = 16758640, Optional ByVal renderOpacity As Single = 50!)
    
    'To improve performance, this function caches the current selection overlay image as a persistent DIB.  (This lets
    ' us cheat and directly AlphaBlend the cached result during intensive operations like paint strokes.)  If the associated
    ' viewport cached has changed, we obviously need to regenerate our overlay cache - but we also need to regenerate the
    ' cache if the current selection UI settings have changed (such as selection rendering color).
    Dim renderingSettingsChanged As Boolean: renderingSettingsChanged = False
    If (Not overlayWasRegenerated) Then renderingSettingsChanged = Strings.StringsNotEqual(m_LastOverlayCache, GenerateArbitraryHash(renderType, renderColor, renderOpacity), False)
    
    'If any preceding checks failed, generate a new UI DIB to match current settings.
    If (overlayWasRegenerated Or renderingSettingsChanged Or (Not m_OverlayIsReady)) Then
        
        If UserPrefs.GenerateDebugLogs Then Debug.Print "rebuilding selection overlay (" & Timer & ")"
        
        'The following chunk of code generates the actual 32-bpp selection image "overlaid" on the viewport.
        ' (Said another way, this image reflects things like the user's currently selected highlight or lightbox UI settings.)
        
        'Start by making sure our cached overlay dimensions match the temporary viewport DIB generated by our parent function.
        If (m_ViewportOverlay Is Nothing) Then Set m_ViewportOverlay = New pdDIB
        If (m_ViewportOverlay.GetDIBWidth <> imgViewportRectL.Width) Or (m_ViewportOverlay.GetDIBHeight <> imgViewportRectL.Height) Then
            m_ViewportOverlay.CreateBlank imgViewportRectL.Width, imgViewportRectL.Height, 32, 0, 0
            m_ViewportOverlay.SetInitialAlphaPremultiplicationState True
        Else
            m_ViewportOverlay.ResetDIB 0
        End If
        
        'In the future, I may allow selection masks to mask specific color channels (which would be awesome), but for now,
        ' we treat the mask as if it contains grayscale entities only.  That allows us to cheat and use byte-sized lookup tables.
        
        'Start by generating a lookup table of all 256 possible overlay colors.
        Dim overlayLookup() As Long
        ReDim overlayLookup(0 To 255) As Long
        
        Dim tmpRGBA As RGBQuad, i As Long
        Dim newR As Long, newG As Long, newB As Long, newA As Single
        Dim defaultFillColor As Long
        
        'Extract RGB values from the passed color
        newR = Colors.ExtractRed(renderColor)
        newG = Colors.ExtractGreen(renderColor)
        newB = Colors.ExtractBlue(renderColor)
        
        'Convert opacity to a [0, 1] scale, and because we're multiplying it by (color / 255), we can go a step
        ' further and just pre-multiply it against (1 / 255)
        Dim opacityMultiplier As Single
        opacityMultiplier = renderOpacity * 0.01! * (1! / 255!)
            
        'The constructed lookup table varies by overlay type
        If (renderType = PDSR_Highlight) Then
            
            'Build a lookup table of opacity-corrected values
            For i = 0 To 255
                newA = CSng(i) * opacityMultiplier
                With tmpRGBA
                    .Red = newR * newA
                    .Green = newG * newA
                    .Blue = newB * newA
                    .Alpha = newA * 255
                End With
                CopyMemoryStrict VarPtr(overlayLookup(i)), VarPtr(tmpRGBA), 4&
            Next i
            
            'Calculate a default fill color and opacity.  (This is the default color we fill the overlay with
            ' before doing our actual per-pixel rendering.)  Note that we deliberately use premultiplied alpha,
            ' despite GDI+ operating on non-premultiplied alpha by default.
            If (Me.GetSelectionProperty_Long(sp_Area) = sa_Exterior) Then
                newA = (renderOpacity * 0.01)
                defaultFillColor = RGB(newR * newA, newG * newA, newB * newA)
            Else
                newA = 0#
                defaultFillColor = 0
            End If
        
        ElseIf (renderType = PDSR_Lightbox) Then
            
            For i = 0 To 255
                newA = CSng(255 - i) * opacityMultiplier
                With tmpRGBA
                    .Red = newR * newA
                    .Green = newG * newA
                    .Blue = newB * newA
                    .Alpha = newA * 255
                End With
                CopyMemoryStrict VarPtr(overlayLookup(i)), VarPtr(tmpRGBA), 4&
            Next i
            
            'Calculate a default fill color and opacity.
            If (Me.GetSelectionProperty_Long(sp_Area) = sa_Exterior) Then
                newA = 0#
                defaultFillColor = 0
            Else
                newA = (renderOpacity * 0.01)
                defaultFillColor = RGB(newR * newA, newG * newA, newB * newA)
            End If
            
        End If
        
        'To shortcut per-pixel fill operations, fill the entire DIB with the default "background" color before we start.
        Dim cSurface As pd2DSurface, cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, m_ViewportOverlay.GetDIBDC, False
        GDI_Plus.GDIPlus_GraphicsSetCompositingMode cSurface.GetHandle, GP_CM_SourceCopy
        Drawing2D.QuickCreateSolidBrush cBrush, defaultFillColor, newA * 100
        PD2D.FillRectangleI cSurface, cBrush, 0, 0, m_ViewportOverlay.GetDIBWidth, m_ViewportOverlay.GetDIBHeight
        Set cBrush = Nothing: Set cSurface = Nothing
        
        'With the lookup table successfully generated, we can generate the overlay using identical code (regardless of the
        ' actual overlay UI style)
        Dim x As Long, y As Long
        Dim overlayPixels() As Long, maskPixels() As Long
        Dim overlaySALine As SafeArray1D, maskSALine As SafeArray1D
        
        'Finally, calculate loop bounds, and make sure there is no possible chance of OOB errors
        Dim viewportIntersectRectCanvas As RectF
        srcImage.ImgViewport.GetIntersectRectCanvas viewportIntersectRectCanvas
    
        Dim loopStartX As Long, loopStartY As Long, loopEndX As Long, loopEndY As Long
        loopStartX = Int(m_ViewportRefRectF.Left)
        loopStartY = Int(m_ViewportRefRectF.Top)
        
        '(When the viewport is deeply zoomed in, we need larger compensation boundaries around the image's edges.)
        Dim extendX As Single, extendY As Single
        extendX = 1.999999
        extendY = 1.999999
        loopEndX = loopStartX + Int(m_ViewportRefRectF.Width + extendX)
        loopEndY = loopStartY + Int(m_ViewportRefRectF.Height + extendY)
        
        If (loopEndX > m_ViewportOverlay.GetDIBWidth - 1) Then loopEndX = m_ViewportOverlay.GetDIBWidth - 1
        If (loopEndY > m_ViewportOverlay.GetDIBHeight - 1) Then loopEndY = m_ViewportOverlay.GetDIBHeight - 1
        If (loopStartX < 0) Then loopStartX = 0
        If (loopStartY < 0) Then loopStartY = 0
        If (loopStartX > loopEndX) Then loopStartX = loopEndX
        If (loopStartY > loopEndY) Then loopStartY = loopEndY
        
        'Failsafe check for disastrous inputs
        If (loopStartX >= 0) And (loopStartY >= 0) Then
        
            'With everything prepared, we can now (finally) apply the results to the overlay DIB.  Note that we use a
            ' per-scanline approach for accessing pixels.  This lets us use 1-dimensional arrays instead of 2-dimensional
            ' arrays for a slight performance boost.
            m_ViewportReference.WrapLongArrayAroundScanline maskPixels, maskSALine, 0
            Dim viewRefPtr As Long, viewRefStride As Long
            viewRefPtr = maskSALine.pvData
            viewRefStride = maskSALine.cElements * 4
            
            m_ViewportOverlay.WrapLongArrayAroundScanline overlayPixels, overlaySALine, 0
            Dim viewOverPtr As Long, viewOverStride As Long
            viewOverPtr = overlaySALine.pvData
            viewOverStride = overlaySALine.cElements * 4
            
            For y = loopStartY To loopEndY
                maskSALine.pvData = viewRefPtr + viewRefStride * y
                overlaySALine.pvData = viewOverPtr + viewOverStride * y
            For x = loopStartX To loopEndX
                overlayPixels(x) = overlayLookup(maskPixels(x) And 255)
            Next x
            Next y
            
            'Because the arrays wrapped around the target DIBs are scoped locally, we have to manually free them before
            ' the function exits.
            m_ViewportReference.UnwrapLongArrayFromDIB maskPixels
            m_ViewportOverlay.UnwrapLongArrayFromDIB overlayPixels
            
            'Whenever we generate a new UI overlay, we need to generate a hash of the rendering settings used.  (If rendering
            ' UI settings change, we'll need to generate a new overlay, even if the selection itself hasn't changed.)
            
            'IMPORTANT NOTE: if new UI settings are added in the future, make sure they are reflected here, and also in the
            ' previously overlay cache check, above.
            m_LastOverlayCache = GenerateArbitraryHash(renderType, renderColor)
            m_OverlayIsReady = True
            
        Else
            PDDebug.LogAction "WARNING!  pdSelection.RenderSelectionAsOverlay received catastrophically bad inputs."
        End If
        
    End If
    
    'Final task: rendering the actual overlay onto the destination viewport.  If rendering and viewport settings haven't
    ' changed since the last render, this step may be the only thing this function actually performs (which is why it's
    ' so speedy during paint ops).
    If (Not m_ViewportOverlay Is Nothing) Then m_ViewportOverlay.AlphaBlendToDCEx dstDIB.GetDIBDC, imgViewportRectL.Left, imgViewportRectL.Top, imgViewportRectL.Width, imgViewportRectL.Height, 0, 0, imgViewportRectL.Width, imgViewportRectL.Height
    
End Sub

'Child function of RenderSelectionToViewport(), above.  Renders the current selection mask to a target viewport,
' using the OUTLINE or MARCHING ANT selection styles.
Private Sub RenderSelectionAsOutline(ByRef dstDIB As pdDIB, ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, ByRef imgViewportRectL As RectL_WH, ByRef overlayWasRegenerated As Boolean, ByVal renderType As PD_SelectionRender)
    
    'Outline rendering isn't free.  In debug builds you may want to profile it.
    ' (At present, profile reports are commented-out, but timing is still tracked - so just uncomment the
    '  various pdDebug.LogAction() calls to gather performance data.)
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Rendering an outline requires two broad stages:
    ' 1a) See if our currently cached outline path is accurate.
    ' 1b) If (1a) fails, we need to generate a new outline path, specific to the current viewport settings.
    '     (This provides a huge speed boost when the current image is large and zoomed-out, as we only generate a path
    '      relevant to the current on-screen representation.)
    ' 2) Render the actual outline onto the target viewport DIB
    
    'Let's begin with task (1): figuring out if our existing outline path is both:
    ' 1) available, and...
    ' 2) up-to-date.
    If (Not m_FinalOutlineIsReady) Or (m_FinalOutline Is Nothing) Or overlayWasRegenerated Then
        
        'Debug.Print "rebuilding selection outline path (" & Timer & ")"
        
        If (m_FinalOutline Is Nothing) Then Set m_FinalOutline = New pd2DPath Else m_FinalOutline.ResetPath
        
        'PD supports two different paths for populating an "outline path":
        ' 1) 99% of the time, we want to perform a manual outline scan of the selection mask
        '    (scaled and zoomed to match the current viewport).  This works for any selection tool,
        '    regardless of shape, feathering, or other esoteric settings, which makes it universally
        '    awesome.
        '
        ' 2) As always, there is an exception, however: when rendering a lasso or polygon selection,
        '    if the selection has *not yet been completed*, we want to simply render the current
        '    selection outline "as-is" - this is important for conveying that the selection itself is
        '    still "under construction".  For this, we use a separate, dedicated rendering function
        '    based on the current in-progress polygon or lasso points.
        Dim useLegacyVectorMode As Boolean
        useLegacyVectorMode = (m_SelectionShape = ss_Lasso) And (Not m_LassoClosed)
        If (Not useLegacyVectorMode) Then useLegacyVectorMode = ((m_SelectionShape = ss_Polygon) And (Not m_PolygonClosed))
        
        If useLegacyVectorMode Then
            
            'Polygon and lasso selections use roughly identical outline code; the only difference is the reference array used
            ' and the number of points.
            If (m_SelectionShape = ss_Polygon) Or (m_SelectionShape = ss_Lasso) Then
                
                Dim numOfRenderPoints As Long
                If (m_SelectionShape = ss_Polygon) Then numOfRenderPoints = m_NumOfPolygonPoints Else numOfRenderPoints = m_NumOfLassoPoints
                
                If (numOfRenderPoints > 1) Then
                
                    'Convert the lasso or polygon array to viewport coordinate space
                    Dim tmpViewportSpace() As PointFloat
                    ReDim tmpViewportSpace(0 To numOfRenderPoints - 1) As PointFloat
                    
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
                        m_FinalOutline.AddPolygon numOfRenderPoints, VarPtr(tmpViewportSpace(0)), m_PolygonClosed, True, GetSelectionProperty_Float(sp_PolygonCurvature)
                    Else
                        m_FinalOutline.AddPolygon numOfRenderPoints, VarPtr(tmpViewportSpace(0)), m_LassoClosed, True, GetSelectionProperty_Float(sp_SmoothStroke)
                    End If
                    
                    'Translate the final path into its correct on-screen position, while also accounting for the 1px border we
                    ' manually add to the transparent "overlay" DIB (which gives us room for outlines that extend "outside"
                    ' the first row and column of pixels in the image).
                    m_FinalOutline.TranslatePath -imgViewportRectL.Left + 1, -imgViewportRectL.Top + 1
                    
                Else
                    Set m_FinalOutline = Nothing
                End If
                
            End If
        
        '99% of the time, we'll be using our standard, "universal" outline detector.
        Else
            
            'm_ViewportRefRectF stores the coordinates of where the current selection was rendered onto the cached image.
            ' (Note that it may not be located at position (0, 0), especially if the image is zoomed out!)
            '
            'From its values, we want to construct an integer-only set of boundaries, appropriate for a per-pixel loop.
            Dim loopStartX As Long, loopStartY As Long, loopEndX As Long, loopEndY As Long
            loopStartX = Int(m_ViewportRefRectF.Left)
            loopStartY = Int(m_ViewportRefRectF.Top)
            If (loopStartX < 0) Then loopStartX = 0
            If (loopStartY < 0) Then loopStartY = 0
            loopEndX = loopStartX + Int(m_ViewportRefRectF.Width + 1.999999)
            loopEndY = loopStartY + Int(m_ViewportRefRectF.Height + 1.999999)
            
            If (loopEndX > imgViewportRectL.Width - 1) Then loopEndX = imgViewportRectL.Width - 1
            If (loopEndY > imgViewportRectL.Height - 1) Then loopEndY = imgViewportRectL.Height - 1
            If (loopStartX > loopEndX) Then loopStartX = loopEndX
            If (loopStartY > loopEndY) Then loopStartY = loopEndY
            
            'We now want to use those loop boundaries to produce a byte-only copy of the selection mask.  (This is required
            ' by the edge-detection engine, and it also gives a chance to better handle things like feathered selection edges.)
            Dim newWidth As Long, newHeight As Long
            newWidth = (loopEndX - loopStartX) + 4
            newHeight = (loopEndY - loopStartY) + 4
            
            If (newWidth <> m_ViewportByteCopyRect.Width) Or (newHeight <> m_ViewportByteCopyRect.Height) Then
                ReDim m_ViewportByteCopy(0 To newWidth, 0 To newHeight) As Byte
                m_ViewportByteCopyRect.Width = newWidth
                m_ViewportByteCopyRect.Height = newHeight
            End If
            
            'We're going to offset each line by several pixels on either side; this guarantees blank boundary pixels,
            ' which lets us skip time-consuming boundary checks.
            Dim xModifier As Long, yModifier As Long
            xModifier = (-loopStartX + 2)
            yModifier = (-loopStartY + 2)
            
            Dim maskPixels() As Long, maskSALine As SafeArray1D
            
            Dim x As Long, y As Long, tmpXPosition As Long, tmpYPosition As Long
            
            For y = loopStartY To loopEndY
                m_ViewportReference.WrapLongArrayAroundScanline maskPixels, maskSALine, y
                tmpYPosition = y + yModifier
            For x = loopStartX To loopEndX
                If ((maskPixels(x) And 255) > 64) Then
                    m_ViewportByteCopy(x + xModifier, tmpYPosition) = 255
                Else
                    m_ViewportByteCopy(x + xModifier, tmpYPosition) = 0
                End If
            Next x
            Next y
            
            'Next, we must manually mirror edge pixels along a 1px border on the top- and left- sides; this guarantees that
            ' our outline path remains *outside* the current viewport, when the selection is greatly zoomed-in.
            
            'The top-left corner pixel is handled specially
            m_ViewportByteCopy(loopStartX + xModifier - 1, loopStartY + yModifier - 1) = m_ViewportByteCopy(loopStartX + xModifier, loopStartY + yModifier)
            
            'Top row
            tmpYPosition = loopStartY + yModifier
            For x = loopStartX To loopEndX
                m_ViewportByteCopy(x + xModifier, tmpYPosition - 1) = m_ViewportByteCopy(x + xModifier, tmpYPosition)
            Next x
            
            'Left row
            tmpXPosition = loopStartX + xModifier
            For y = loopStartY To loopEndY
                m_ViewportByteCopy(tmpXPosition - 1, y + yModifier) = m_ViewportByteCopy(tmpXPosition, y + yModifier)
            Next y
            
            'Because the array wrapped around the reference DIB is scoped locally, we need to free it before exiting
            m_ViewportReference.UnwrapLongArrayFromDIB maskPixels
            
            'Find the edges of the temporary byte-map we've just generated
            If (m_ViewportEdges Is Nothing) Then Set m_ViewportEdges = New pdEdgeDetector
            m_ViewportEdges.FindAllEdges m_FinalOutline, m_ViewportByteCopy, 1, 1, m_ViewportByteCopyRect.Width - 1, m_ViewportByteCopyRect.Height - 1, loopStartX - 1, loopStartY - 1
            
        End If
        
        'Mark the finished overlay as "ready"; we won't regenerate it unless the underlying selection mask (or the current
        ' viewport position/zoom) changes.
        m_FinalOutlineIsReady = True
        
    End If
        
    'We are now ready to render the actual outline onto the overlay DIB.
    If (Not m_FinalOutline Is Nothing) Then
    
        'Start by making sure our cached overlay dimensions are correct.  We need the overlay DIB to be at least 1-px larger
        ' in each dimension than the temporary viewport DIB generated by our parent function; this allows us to paint our
        ' outline "outside" the image itself, as appropriate.
        If (m_ViewportOverlay Is Nothing) Then Set m_ViewportOverlay = New pdDIB
        If (m_ViewportOverlay.GetDIBWidth <> imgViewportRectL.Width + 2) Or (m_ViewportOverlay.GetDIBHeight <> imgViewportRectL.Height + 2) Then
            m_ViewportOverlay.CreateBlank imgViewportRectL.Width + 2, imgViewportRectL.Height + 2, 32, 0, 0
            m_ViewportOverlay.SetInitialAlphaPremultiplicationState True
        Else
            m_ViewportOverlay.ResetDIB 0
        End If
        
        'Draw the assembled path onto the canvas
        Dim cSurface As pd2DSurface, cPenUIBase As pd2DPen, cPenUITop As pd2DPen
        
        'I haven't made a final decision on using antialiasing for the rendering outline.  At present,
        ' antialiasing is used for "outline" mode, but not "marching ant" mode.
        Dim useAA As Boolean
        useAA = (renderType = PDSR_Outline)
        
        Drawing2D.QuickCreateSurfaceFromDC cSurface, m_ViewportOverlay.GetDIBDC, useAA
        
        'Half-pixel offsets solve some rendering problems, while also introducing other ones.  I've left it disabled for now,
        ' pending additional testing.
        'cSurface.SetSurfacePixelOffset P2_PO_Half
        
        'Render the outline using PD's standard "white-on-black" outline technique
        If (renderType = PDSR_Outline) Then
            
            Drawing2D.QuickCreatePairOfUIPens cPenUIBase, cPenUITop, , P2_LJ_Round, P2_LC_Round
            PD2D.DrawPath cSurface, cPenUIBase, m_FinalOutline
            PD2D.DrawPath cSurface, cPenUITop, m_FinalOutline
        
        'Render the outline using an animated "marching ants" approach
        Else
            
            Dim cBlackPen As pd2DPen, cDottedPen As pd2DPen
            Set cBlackPen = New pd2DPen
            cBlackPen.SetPenColor vbBlack
            cBlackPen.SetPenWidth 1!
            
            Set cDottedPen = New pd2DPen
            cDottedPen.SetPenWidth 1!
            cDottedPen.SetPenColor vbWhite
            cDottedPen.CreatePen
            cDottedPen.SetPenStyle P2_DS_Custom
            cDottedPen.SetPenDashes_UNSAFE VarPtr(m_AntDashes(0)), 2
            cDottedPen.SetPenDashOffset m_AntDashOffset
            
            PD2D.DrawPath cSurface, cBlackPen, m_FinalOutline
            PD2D.DrawPath cSurface, cDottedPen, m_FinalOutline
                        
            'If marching ants are being used, make sure our timer is enabled now
            If (Not m_AntTimer.IsActive) Then
                m_AntTimer.Interval = ANT_DASH_SPEED
                m_AntTimer.StartTimer
            End If
        
        End If
        
        'Update the overlay hash; even though we don't track this locally, it is necessary for overlay and lightbox rendering
        ' to detect changes to their rendering mode options.
        m_LastOverlayCache = GenerateArbitraryHash(renderType)
        
    Else
        If (Not m_ViewportOverlay Is Nothing) Then m_ViewportOverlay.ResetDIB 0
    End If
    
    'Final task: rendering the final overlay onto the destination viewport.  If rendering and viewport settings haven't
    ' changed since the last render, this step may be the only thing this function actually performs (which is why it's
    ' so speedy during paint ops).
    If (Not m_ViewportOverlay Is Nothing) Then m_ViewportOverlay.AlphaBlendToDCEx dstDIB.GetDIBDC, imgViewportRectL.Left - 1, imgViewportRectL.Top - 1, imgViewportRectL.Width + 2, imgViewportRectL.Height + 2, 0, 0, imgViewportRectL.Width + 2, imgViewportRectL.Height + 2
    
    'Timing reports:
    'pdDebug.LogAction "Selection outline rendered in " & Format$(VBHacks.GetTimerDifferenceNow(startTime) * 1000, "#0") & " ms."
    
End Sub

'When resources are tight, you can call this sub to free some non-essential internal caches.  This will cause a performance hit
' on subsequent selection actions, so please only do it if the savings are relevant.
Friend Sub FreeNonEssentialResources()
    
    'The flood fill manager caches a lot of internal resources (like the fill stack), so we may be able to save some memory there.
    If (Not m_FloodFill Is Nothing) Then m_FloodFill.FreeUpResources
    
    'If the current selection mask isn't actively being rendered, it doesn't need a DC attached.
    m_SelMask.FreeFromDC
    
End Sub

'Create a selection from selection data previously saved to file
' (NOTE: this function will not generate a selection mask or render the selection on-screen.  The calling function must explicitly
'        request a render if they want one.)
Friend Function ReadSelectionFromFile(ByRef srcFilename As String, Optional ignoreLockStatus As Boolean = False) As Boolean
    
    'Check for modern format first (legacy format is still supported)
    Dim packageIsModern As Boolean, cPackage As pdPackageChunky
    Set cPackage = New pdPackageChunky
    packageIsModern = cPackage.OpenPackage_File(srcFilename, "PDSF")
    Set cPackage = Nothing
    
    If packageIsModern Then
        ReadSelectionFromFile = ReadSelectionFromFile_New(srcFilename, ignoreLockStatus)
    Else
        ReadSelectionFromFile = ReadSelectionFromFile_Legacy(srcFilename, ignoreLockStatus)
    End If
    
End Function
        
'Create a selection from selection data previously saved to file
' (NOTE: this function will not generate a selection mask or render the selection on-screen.  The calling function must explicitly
'        request a render if they want one.)
Friend Function ReadSelectionFromFile_New(ByRef srcFilename As String, Optional ignoreLockStatus As Boolean = False) As Boolean
    
    ReadSelectionFromFile_New = False
    PDDebug.LogAction "Reading selection data from file: " & srcFilename
    
    'Like all other PD-specific files, selection files are just pdPackage instances
    Dim cPackage As pdPackageChunky
    Set cPackage = New pdPackageChunky
    If cPackage.OpenPackage_File(srcFilename, "PDSF") Then
    
        'Reset some of our internal trackers prior to reading the selection data
        Me.EraseCustomTrackers
        
        'First, note that our mask is not yet ready.  (Masks are generated on-demand, unless this is
        ' a raster selection, in which case we'll load the mask bits directly from the file.)
        m_IsMaskReady = False
        
        Dim chnkName As String, chnkSize As Long, chnkStream As pdStream
        Dim xmlString As String
        
        Do While cPackage.ChunksRemain()
            
            Dim origWidth As Long, origHeight As Long
            Dim tmpDIB As pdDIB, tmpDIBPointer As Long, tmpDIBLength As Long
            
            'Internal selection data
            If (cPackage.GetChunkName() = "SELI") Then
                
                If cPackage.GetNextChunk(chnkName, chnkSize, chnkStream) Then
                    xmlString = chnkStream.ReadString_UTF8(chnkSize)
                    Me.InitFromXML xmlString
                Else
                    PDDebug.LogAction "WARNING!  Selection reader couldn't read internal selection XML."
                End If
            
            'External selection data
            ElseIf (cPackage.GetChunkName() = "SELE") Then
                
                If cPackage.GetNextChunk(chnkName, chnkSize, chnkStream) Then
                    
                    Dim headerXML As pdSerialize
                    Set headerXML = New pdSerialize
                    headerXML.SetParamString chnkStream.ReadString_UTF8(chnkSize)
                    
                    'Use the minor header to populate a few extra settings
                    m_IsLocked = headerXML.GetBool("sel-locked", False)
                    m_IsTransformable = headerXML.GetBool("sel-transformable", True)
                    
                    'Certain selection types (e.g. selections that can't be fully described with vectors) store a full copy of their
                    ' selection mask in the selection file.  Check for these and load them conditionally.
                    If (m_SelectionShape = ss_Raster) Or (m_SelectionShape = ss_Wand) Then
                        
                        'Retrieve the original parent image's width and height.  If the original containing image had the same dimensions
                        ' as our current image (e.g. during Undo/Redo operations), we can create the selection mask directly from the
                        ' file data.  Otherwise, we have no choice but to perform a resize so that the old mask matches the new dimensions.
                        origWidth = headerXML.GetLong("sel-mask-width", m_parentPDImage.Width)
                        origHeight = headerXML.GetLong("sel-mask-height", m_parentPDImage.Height)
                        
                        'Dimensions match!  Load the source data directly into our selection mask
                        If (origWidth = m_parentPDImage.Width) And (origHeight = m_parentPDImage.Height) Then
                            
                            With headerXML
                                m_SelMask.CreateBlank origWidth, origHeight, .GetLong("SelMaskDepth", 32), 0, 0
                            End With
                            
                            m_SelMask.SetInitialAlphaPremultiplicationState headerXML.GetBool("sel-mask-alpha-premultiplied", True)
                            
                        'Dimensions do not match.  Use an intermediary DIB to cache the original raster data, then resize it to match
                        ' our current image.
                        Else
                        
                            Set tmpDIB = New pdDIB
                            
                            With headerXML
                                tmpDIB.CreateBlank origWidth, origHeight, .GetLong("sel-mask-depth", 32), 0, 0
                            End With
                            
                            tmpDIB.SetInitialAlphaPremultiplicationState headerXML.GetBool("sel-mask-alpha-premultiplied", True)
                            
                        End If
                        
                        'Note that a mask has been created; this prevents PD from attempting to create a new copy.
                        ' (For magic wand selections, this is particularly bad as the wand may generate totally new boundaries
                        '  on a new image.)
                        m_MaskHasBeenCreated = True
                        m_IsMaskReady = True
                        
                    End If
                        
                Else
                    PDDebug.LogAction "WARNING!  Selection reader couldn't read external selection XML."
                End If
            
            'Mask data (not included in all files - only those that require it)
            ElseIf (cPackage.GetChunkName() = "SELM") Then
            
                'Only load the mask for certain selection types (e.g. ones that can't be defined as vectors)
                If (m_SelectionShape = ss_Raster) Or (m_SelectionShape = ss_Wand) Then
                    
                    'If this image is the same size as the embedded mask, we don't need to stretch anything;
                    ' load the bits directly into our (already created) mask DIB
                    If (origWidth = m_parentPDImage.Width) And (origHeight = m_parentPDImage.Height) Then
                        m_SelMask.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                        If (Not cPackage.GetNextChunk(chnkName, chnkSize, Nothing, tmpDIBPointer, tmpDIBLength)) Then PDDebug.LogAction "WARNING!  Couldn't retrieve selection mask!"
                    Else
                        
                        tmpDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                        
                        If cPackage.GetNextChunk(chnkName, chnkSize, Nothing, tmpDIBPointer, tmpDIBLength) Then
                            m_SelMask.CreateBlank m_parentPDImage.Width, m_parentPDImage.Height, tmpDIB.GetDIBColorDepth, 0, 0
                            GDI_Plus.GDIPlusResizeDIB m_SelMask, 0, 0, m_parentPDImage.Width, m_parentPDImage.Height, tmpDIB, 0, 0, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, GP_IM_HighQualityBicubic
                            m_SelMask.SetInitialAlphaPremultiplicationState tmpDIB.GetAlphaPremultiplication()
                            Set tmpDIB = Nothing
                        Else
                            PDDebug.LogAction "WARNING!  Couldn't retrieve selection mask!"
                        End If
                            
                    End If
                    
                End If
            
            End If
            
        Loop
        
        'Polygon selections loaded from file are assumed to be closed
        m_PolygonClosed = True
        
        'If this selection was locked in at the time of its save, lock this selection in as well
        If (Not ignoreLockStatus) Then
            If m_IsLocked Then Me.LockIn Else Me.LockRelease
            m_parentPDImage.SetSelectionActive m_IsLocked
        End If
                    
        'If this selection isn't a vector, we need to manually find its boundaries now
        If (((m_SelectionShape = ss_Raster) Or (m_SelectionShape = ss_Wand)) And (Not m_SelMask Is Nothing)) Then
            Me.NotifyRasterDataChanged
            Me.FindNewBoundsManually
        End If
        
        ReadSelectionFromFile_New = True
        PDDebug.LogAction "pdSelection.ReadSelectionFromFile complete.  Exiting now..."
    
    End If
    
End Function

'Create a selection from selection data previously saved to file
' (NOTE: this function will not generate a selection mask or render the selection on-screen.  The calling function must explicitly
'        request a render if they want one.)
Friend Function ReadSelectionFromFile_Legacy(ByRef srcFilename As String, Optional ignoreLockStatus As Boolean = False) As Boolean
    
    ReadSelectionFromFile_Legacy = False
    PDDebug.LogAction "Reading legacy selection data from file: " & srcFilename
    
    'Like all other PD-specific files, selection files are just pdPackage instances
    Dim cPackage As pdPackageLegacyV2
    Set cPackage = New pdPackageLegacyV2
    If cPackage.ReadPackageFromFile(srcFilename, SELECTION_IDENTIFIER, PD_SM_FileBacked) Then
    
        'Reset some of our internal trackers prior to reading the selection data
        
        'First, note that our mask is not yet ready.  (Masks are generated on-demand, unless this is a raster selection,
        ' in which case we'll load the actual mask straight from the file.)
        m_IsMaskReady = False
        
        'Retrieve the file-specific header from the package.  This contains details like selection format version, which we need
        ' before we proceed with full parsing.
        Dim minorHeader As String
        If cPackage.GetNodeDataByName_String("SelHeader", True, minorHeader) Then
            
            'Copy the string into an XML parser
            Dim headerXML As pdSerialize
            Set headerXML = New pdSerialize
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
                    
                    'Certain selection types (e.g. selections that can't be fully described with vectors) store a full copy of their
                    ' selection mask in the selection file.  Check for these and load them conditionally.
                    If (m_SelectionShape = ss_Raster) Or (m_SelectionShape = ss_Wand) Then
                        
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
                                m_SelMask.CreateBlank m_parentPDImage.Width, m_parentPDImage.Height, tmpDIB.GetDIBColorDepth, 0, 0
                                GDI_Plus.GDIPlusResizeDIB m_SelMask, 0, 0, m_parentPDImage.Width, m_parentPDImage.Height, tmpDIB, 0, 0, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, GP_IM_HighQualityBicubic
                                m_SelMask.SetInitialAlphaPremultiplicationState tmpDIB.GetAlphaPremultiplication()
                                Set tmpDIB = Nothing
                            End If
                            
                        End If
                        
                        'Note that a mask has been created; this prevents PD from attempting to create a new copy.
                        ' (For magic wand selections, this is particularly bad as the wand may generate totally new boundaries
                        '  on a new image.)
                        m_MaskHasBeenCreated = True
                        m_IsMaskReady = True
                        
                    End If
                    
                    'If this selection isn't a vector, we need to manually find its boundaries now
                    If (((m_SelectionShape = ss_Raster) Or (m_SelectionShape = ss_Wand)) And (Not m_SelMask Is Nothing)) Then
                        Me.NotifyRasterDataChanged
                        Me.FindNewBoundsManually
                    End If
                    
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
                    
                    ReadSelectionFromFile_Legacy = True
                    
                Else
                    PDDebug.LogAction "WARNING!  pdSelection failed to retrieve the actual header string for this file."
                End If
            
            Else
                PDDebug.LogAction "WARNING!  pdSelection found an unknown version in this file header; header XML follows:"
                PDDebug.LogAction headerXML.GetParamString()
            End If
            
        Else
            PDDebug.LogAction "WARNING!  pdSelection failed to load a valid header from this saved selection file."
        End If
        
    End If
    
    PDDebug.LogAction "pdSelection.ReadSelectionFromFile complete.  Exiting now..."
    
End Function

'As of v8.0, pdPackageChunky is used to read/write selection files.  Note that this function will
' blindly overwrite the destination file if it exists; it's up to the caller to plan for this.
Friend Function WriteSelectionToFile(ByVal dstFilename As String, Optional ByVal compressXML As PD_CompressionFormat = cf_Zstd, Optional ByVal xmlCompressionLevel As Long = -1, Optional ByVal compressRaster As PD_CompressionFormat = cf_Zstd, Optional ByVal rasterCompressionLevel As Long = -1, Optional ByRef dstUndoFileSize As Long) As Boolean
    
    Dim cPackage As pdPackageChunky
    Set cPackage = New pdPackageChunky
    cPackage.StartNewPackage_File dstFilename, packageID:="PDSF"
    
    'Selection files are pdPackage (chunky) files with three pieces of data:
    ' 1) A standard vector selection descriptor.  This is an XML string, used internally by PD, that holds
    '    all information necessary to create the current selection (except raster selections; see below).
    ' 2) A "special" file-specific XML string that contains additional internal program state data;
    '    for example, it saves the image width/height associated with the selection - this lets us
    '    dynamically rescale saved selections to new images.
    ' 3) For raster masks only, we also save pixel data from the active selection mask.  (For vector
    '    selections, we dynamically create the mask at load-time, using the data from the XML strings.)
    
    'Before writing any data, ensure default compression levels are handled correctly
    If (xmlCompressionLevel = -1) Then xmlCompressionLevel = Compression.GetDefaultCompressionLevel(compressXML)
    If (rasterCompressionLevel = -1) Then rasterCompressionLevel = Compression.GetDefaultCompressionLevel(compressRaster)
    
    'Start by adding the internal selection XML data
    cPackage.StartChunk "SELI"
    cPackage.GetInProgressChunk.WriteString_UTF8 Me.GetSelectionAsXML()
    cPackage.EndChunk compressXML, xmlCompressionLevel
    
    'Next, prepare a second XML string, this one containing additional internal session data.
    ' (The selection engine only maintains selection-specific data; we need to also know some
    ' external selection-adjacent data, like the size of the current image, so that we can
    ' properly scale the selection to other images in the future.)
    Dim maskWillBeEmbedded As Boolean
    maskWillBeEmbedded = ((m_SelectionShape = ss_Raster) And m_MaskHasBeenCreated) Or ((m_SelectionShape = ss_Wand) And m_MaskHasBeenCreated)
    
    Dim tmpXML As pdSerialize
    Set tmpXML = New pdSerialize
    With tmpXML
        .AddParam "sel-version", SELECTION_FILE_VERSION_2019
        .AddParam "sel-parent-image-width", m_parentPDImage.Width
        .AddParam "sel-parent-image-height", m_parentPDImage.Height
        .AddParam "sel-locked", Me.IsLockedIn
        .AddParam "sel-transformable", m_IsTransformable
        .AddParam "sel-mask-embedded", maskWillBeEmbedded
        
        'If we're going to embed a mask, we also need to store some mask-specific data
        If maskWillBeEmbedded Then
            
            'Make sure the mask is ready before continuing!
            If (Not m_IsMaskReady) Then CreateSelectionMask
            
            .AddParam "sel-mask-depth", m_SelMask.GetDIBColorDepth
            .AddParam "sel-mask-width", m_SelMask.GetDIBWidth
            .AddParam "sel-mask-height", m_SelMask.GetDIBHeight
            .AddParam "sel-mask-stride", m_SelMask.GetDIBStride
            .AddParam "sel-mask-alpha-premultiplied", m_SelMask.GetAlphaPremultiplication
            
        End If
        
    End With
    
    'Add the secondary descriptor to the same node (in the header chunk)
    cPackage.StartChunk "SELE"
    cPackage.GetInProgressChunk.WriteString_UTF8 tmpXML.GetParamString()
    cPackage.EndChunk compressXML, xmlCompressionLevel
    
    'Finally, if this is a raster or magic-wand selection, add the full mask as well
    If maskWillBeEmbedded Then
        Dim maskDIBPointer As Long, maskDIBLength As Long
        m_SelMask.RetrieveDIBPointerAndSize maskDIBPointer, maskDIBLength
        cPackage.AddChunk_WholeFromPtr "SELM", maskDIBPointer, maskDIBLength, compressRaster, rasterCompressionLevel
    End If
    
    WriteSelectionToFile = cPackage.FinishPackage()
    
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
Friend Function FindNewBoundsManually(Optional ByVal overrideRasterState As Boolean = False) As Boolean

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
    
    Dim selMaskData() As Long, selMaskSA As SafeArray2D, selMaskSA1D As SafeArray1D
    
    'While here, grab a baseline DIB pointer and stride, so we can quickly wrap individual lines
    Dim maskPtr As Long, maskStride As Long
    m_SelMask.WrapLongArrayAroundScanline selMaskData, selMaskSA1D
    maskPtr = selMaskSA1D.pvData
    maskStride = selMaskSA1D.cElements * 4
    
    Dim maskWidth As Long, maskHeight As Long
    maskWidth = m_SelMask.GetDIBWidth - 1
    maskHeight = m_SelMask.GetDIBHeight - 1
    
    Dim boundFound As Boolean
    
    'Find the top bound first.
    boundFound = False
    y = 0
    Do
    
        selMaskSA1D.pvData = maskPtr + y * maskStride
    
        For x = 0 To maskWidth
            
            'In the future, if we decide to let the user select individual color channels, this code will need to be reworked.
            ' For now, however, all pixels are rendered as grayscale, so we can shortcut and just check for non-zero entries.
            If (selMaskData(x) <> 0) Then
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
    
        selMaskSA1D.pvData = maskPtr + y * maskStride
    
        For x = 0 To maskWidth
            If (selMaskData(x) <> 0) Then
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
    
    m_SelMask.UnwrapLongArrayFromDIB selMaskData
    m_SelMask.WrapLongArrayAroundDIB selMaskData, selMaskSA
    
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
        Debug.Print "WARNING!  pdSelection.RequestThumbnail was called, but no selection exists."
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
    If (tIcoWidth > tIcoHeight) Then
        tX = 0
        tY = (thumbnailSize - tIcoHeight) * 0.5
    
    '...otherwise, center it horizontally
    Else
        tY = 0
        tX = (thumbnailSize - tIcoWidth) * 0.5
    End If
    
    'Prepare the destination DIB
    If (dstThumbnailDIB Is Nothing) Then Set dstThumbnailDIB = New pdDIB
    If (dstThumbnailDIB.GetDIBWidth <> thumbnailSize) Or (dstThumbnailDIB.GetDIBHeight <> thumbnailSize) Then
        dstThumbnailDIB.CreateBlank thumbnailSize, thumbnailSize, 32, 0
    Else
        dstThumbnailDIB.ResetDIB 0
    End If
    
    'Note that the user's thumbnail performance setting affects the quality used here.
    RequestThumbnail = GDIPlusResizeDIB(dstThumbnailDIB, tX, tY, tIcoWidth, tIcoHeight, m_SelMask, 0, 0, m_SelMask.GetDIBWidth, m_SelMask.GetDIBHeight, UserPrefs.GetThumbnailInterpolationPref())
    
End Function

'Normally, this class can auto-detect any changes that affect the selection mask.  Raster selections are an exception to this,
' because external functions can "hook" into the raster and modify it at will.  If a function does this, it needs to notify us,
' so we can generate a new viewport-specific overlay matching the updated raster data.
Friend Sub NotifyRasterDataChanged()
    m_FinalOutlineIsReady = False
    m_OverlayIsReady = False
    m_ViewportRefReady = False
End Sub

'Want to suspend (or subsequently re-enable) animations?  Use this sub.  Note that it does *not* touch the timer object,
' by design - the activation and deactivation of the timer object is handled elsewhere.
Friend Sub NotifyAnimationsAllowed(ByVal allowedState As Boolean)
    m_AnimationsAllowed = allowedState
End Sub

Private Sub Class_Initialize()
    
    m_IsLocked = False
    
    'Initialize the selection mask for this object and mark it as "not ready" (because no mask has been drawn yet)
    Set m_SelMask = New pdDIB
    m_IsMaskReady = False
    m_LastRenderMode = -1
    
    'Note that a mask has never been created for this selection
    m_MaskHasBeenCreated = False
    
    'Mark it as not transformable... yet
    m_IsTransformable = False
    
    'Prepare the property dictionary
    Set m_PropertyDict = New pdDictionary
    
    'No lasso or polygon points yet
    m_NumOfLassoPoints = 0
    ReDim m_LassoPoints(0) As PointFloat
    ReDim m_LassoPointsBackup(0) As PointFloat
    
    m_NumOfPolygonPoints = 0
    ReDim m_PolygonPoints(0) As PointFloat
    ReDim m_PolygonPointsBackup(0) As PointFloat
    
    'Polygon selections are (obviously) not yet closed
    m_PolygonClosed = False
    
    ReDim m_AntDashes(0 To 1) As Single
    m_AntDashes(0) = ANT_DASH_SIZE
    m_AntDashes(1) = ANT_DASH_SIZE
    
    'We may not need a marching ants timer, but instantiate one regardless.  (This simplifies checking animation state
    ' inside the rendering loop.)
    Set m_AntTimer = New pdTimer
    m_AnimationsAllowed = True
    
End Sub

Private Sub Class_Terminate()
    
    If (Not m_SelMask Is Nothing) Then
        m_SelMask.EraseDIB
        Set m_SelMask = Nothing
    End If

End Sub

Private Sub m_AntTimer_Timer()
    
    'Disallow animations if an external object has forcibly suspended them
    If m_AnimationsAllowed Then
        
        'Advance the current offset
        m_AntDashOffset = m_AntDashOffset + 1#
        If (m_AntDashOffset >= ANT_DASH_SIZE * 2) Then m_AntDashOffset = 0#
        
        'Perform a failsafe check to see if we're even needed
        Dim okayToRender As Boolean
        okayToRender = Me.IsLockedIn
        
        If okayToRender And (Not m_parentPDImage Is Nothing) Then
            okayToRender = (m_parentPDImage.imageID = PDImages.GetActiveImageID())
        Else
            okayToRender = False
        End If
        
        If okayToRender Then
            Dim tmpViewportParams As PD_ViewportParams
            tmpViewportParams = Viewport.GetDefaultParamObject()
            tmpViewportParams.curPOI = poi_ReuseLast
            Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0), VarPtr(tmpViewportParams)
        Else
            m_AntTimer.StopTimer
        End If
        
    End If
    
End Sub
'For transforms to work, we need to apply any transformation matrices to *the original* polygon/lasso points.
' As such, we maintain a persistent second copy of the current point collection, backed up whenever the
' current point collection changes.
Private Sub BackupCurrentSelectionPoints()

    Select Case m_SelectionShape
    
        Case ss_Polygon
            If (UBound(m_PolygonPointsBackup) < m_NumOfPolygonPoints - 1) Then ReDim m_PolygonPointsBackup(0 To m_NumOfPolygonPoints - 1) As PointFloat
            CopyMemoryStrict VarPtr(m_PolygonPointsBackup(0)), VarPtr(m_PolygonPoints(0)), m_NumOfPolygonPoints * 8
            
        Case ss_Lasso
            If (UBound(m_LassoPointsBackup) < m_NumOfLassoPoints - 1) Then ReDim m_LassoPointsBackup(0 To m_NumOfLassoPoints - 1) As PointFloat
            CopyMemoryStrict VarPtr(m_LassoPointsBackup(0)), VarPtr(m_LassoPoints(0)), m_NumOfLassoPoints * 8
            
    End Select

End Sub
