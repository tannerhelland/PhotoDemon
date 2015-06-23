VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdVoronoi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Voronoi class
'Copyright 2014-2015 by Tanner Helland
'Created: 14/July/14
'Last updated: 23/June/15
'Last update: move randomizing to pdRandomize
'Dependencies: pdRandomize (for randomizing point distribution)
'
'While this class is called pdVoronoi, it's primarily a Worley Noise implementation (http://en.wikipedia.org/wiki/Worley_noise),
' that uses only the Voronoi aspects most relevant to image processing (http://en.wikipedia.org/wiki/Voronoi_diagram).
' For a full Voronoi toolkit, you'll need to look elsewhere, but if all you need is the bits relevant to generating Worley Noise,
' you're in for a treat, because this class is completely free of dependencies to other PhotoDemon code - so have at it!
'
'Voronoi diagrams work by taking a series of points, and for the relevant space around them, finding the nearest Voronoi
' point to each location.  When performed at a sufficiently detailed level (e.g. for each pixel in a grid), you are left
' with a puzzle-like appearance a la http://en.wikipedia.org/wiki/Voronoi_diagram#mediaviewer/File:Euclidean_Voronoi_Diagram.png
'
'For image processing, Voronoi noise is an especially helpful technique for image segmentation.  Its biggest issue is that
' it is painfully slow - comparing each pixel in an image to thousands (or more) Voronoi points, until you find the nearest one,
' is an unpleasant exercise that will bring even high-end PCs to a grinding halt.
'
'This class circumvents this is by carefully partitioning the initial array of Voronoi points into a rough grid, then randomizing
' each point within a predetermined grid radius only.  While this limits the exoticness of the resulting Voronoi diagram, it allows
' us to search only a relevant neighborhood for each pixel, which in turn makes processing a photograph-sized image a task that
' takes several seconds instead of several hours - a worthwhile improvement!
'
'Three standard distance methods (cartesian, manhattan, and chebyshev) are provided.  Set the desired distance method using
' the setDistanceMode function, and make sure to set it *before* calling the getNearestPointIndex function!
'
'Some helper functions exist to aid with image-processing-specific tasks like cell shading, and these rely on the caching of
' certain values within the heavily used getNearestPointIndex function.  If you don't need the data these caches provide, feel free
' to comment out everything related to the vPointsMaxDistance() arrays.  It won't result in a huge speed gain, but you'll save a bit
' of cache thrashing, which never hurts.
'
'Finally, many thanks to Robert Rayment, who did extensive profiling and research on Voronoi implementations before I ever
' began work on this class.  His comments were invaluable in determining the shape and style of this class's interface.
' FYI, Robert's PaintRR project has a faster and simpler version of this routine worth checking out if PD's method seems like
' overkill! Link here: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Available distance calculation methods
Public Enum VORONOI_DISTANCE
    DIST_CARTESIAN = 0
    DIST_MANHATTAN = 1
    DIST_CHEBYSHEV = 2
End Enum

#If False Then
    Const DIST_CARTESIAN = 0, DIST_MANHATTAN = 1, DIST_CHEBYSHEV = 2
#End If

'Available shading methods
Public Enum VORONOI_SHADING
    NO_SHADE = 1
    SHADE_F1 = 2
    SHADE_F2 = 3
    SHADE_F2_MINUS_F1 = 4
    SHADE_F1_DIVIDE_F2 = 5
End Enum

#If False Then
    Const NO_SHADE = 1, SHADE_F1 = 2, SHADE_F2 = 3, SHADE_F2_MINUS_F1 = 4, SHADE_F1_DIVIDE_F2 = 5
#End If

'This vPoints() array will store the coordinates of each point in the Voronoi diagram
Private vPoints() As POINTAPI

'The vPointsMaxDistance() array will store the maximum value of the shade calculation for a given cell, during the nearest
' point search stage.  This allows us to later return shade on the scale [0, 1] regardless of shading method, without having
' to perform a second distance check.
Private vPointsMaxDistance() As Single

'Size of a given cell (in one dimension), as supplied by the user
Private m_cellSize As Long

'Number of rows and columns in the diagram.
Private m_numRows As Long, m_numColumns As Long

'Size of the image associated with this diagram
Private m_imgWidth As Long, m_imgHeight As Long

'Points have been successfully initialized
Private m_PointsInitialized As Boolean

'Technique used to calculate distance between points; this is set via the setDistanceMode() sub
Private m_distanceMode As VORONOI_DISTANCE

'Technique used to calculate shading; this is set via the setShadingMode() sub
Private m_shadingMode As VORONOI_SHADING

'Turbulence currently used for the function
Private m_Turbulence As Single

'Radius of nearby points to search.  For certain functions, we can get away with only searching the nearest 9 points;
' unfortunately, certain combinations of patterns require a larger search area.
Private m_VoronoiSearchRadius As Long

'PRNG
Private cRandom As pdRandomize

'Given a DIB and the user's requested number of rows and columns, populate an initial vPoints() array.
' In the future, I may allow the user to supply a specific point pattern, e.g. "Triangle" or "Square" or "Hexagon".
' Right now, squares are assumed, and the passed pointPattern value affects nothinh.
Public Sub initPoints(ByVal cellSize As Long, ByVal srcImageWidth As Long, srcImageHeight As Long, Optional ByVal pointPattern As Long = 0)

    'Store the cell size
    m_cellSize = cellSize
    
    'Cache the image's width and height, as we'll need them multiple times in the future.  (Because these values are cached,
    ' this initialization function must be called again if the image's dimensions change!)
    m_imgWidth = srcImageWidth
    m_imgHeight = srcImageHeight
    
    'Calculate the number of rows and columns in the array, as a function of cell size and image size
    m_numRows = m_imgHeight \ m_cellSize
    m_numColumns = m_imgWidth \ m_cellSize
    
    'Resize all arrays to match
    ReDim vPoints(0 To m_numColumns - 1, 0 To m_numRows - 1) As POINTAPI
    ReDim vPointsMaxDistance(0 To getTotalNumOfVoronoiPoints() - 1) As Single
    
    'So that each point is centered nicely inside a cell, we'll make use of half width and height values
    Dim hCellSize As Long
    hCellSize = cellSize \ 2
    
    'Populate initial point positions
    Dim x As Long, y As Long
    For x = 0 To m_numColumns - 1
    For y = 0 To m_numRows - 1
        vPoints(x, y).x = (x * m_imgWidth) \ m_numColumns + hCellSize
        vPoints(x, y).y = (y * m_imgHeight) \ m_numRows + hCellSize
    Next y
    Next x
    
    'Note that initialization was succesful
    m_PointsInitialized = True

End Sub

'Randomize the stored point array by some set amount.
' Turbulence is a value on the scale [0, 1]; 1 will result in maximum randomization
' Seed is optional; the same seed will result in the same diagram
Public Function randomizePoints(ByVal fxTurbulence As Single, Optional ByVal rndSeed As Double = 0) As Boolean

    'Make sure the point array was successfully initialized
    If Not m_PointsInitialized Then
        Debug.Print "You haven't initialized the point array yet!"
        Exit Function
    End If
    
    'Seed the randomizer
    Set cRandom = New pdRandomize
    If rndSeed = 0 Then
        cRandom.setSeed_AutomaticAndRandom
    Else
        cRandom.setSeed_Float rndSeed
    End If
    
    'Cache the turbulence value
    m_Turbulence = fxTurbulence
    
    'Perturb each point in the array by an amount proportional to the cell size; at max turbulence, points can
    ' be perturbed by "cell size / 2", times the user's turbulence parameter (which is on the range [0, 1])
    Dim tmpCellSize As Single
    tmpCellSize = (m_cellSize \ 2) * fxTurbulence
    
    Dim x As Long, y As Long
    For x = 0 To m_numColumns - 1
    For y = 0 To m_numRows - 1
        vPoints(x, y).x = vPoints(x, y).x + (1 - (cRandom.getRandomFloat_WH * 2)) * tmpCellSize
        vPoints(x, y).y = vPoints(x, y).y + (1 - (cRandom.getRandomFloat_WH * 2)) * tmpCellSize
    Next y
    Next x
    
    'Update the Voronoi search radius to account for the new turbulence parameter
    determineVoronoiSearchRadius
    
    randomizePoints = True

End Function

'Internal function for determining optimal Voronoi search radius.  If an F2 distance is involved, or the turbulence parameter
' is quite high, the search radius must be extended.  This function will automatically be called after the shading mode or
' turbulence parameter is changed.
' (NOTE: in PhotoDemon, I limit the turbulence factor to 1/2 of the cell size, which saves us from having to extend the
'        search radius due to turbulence.  Individuals needing greater turbulence will thus need to uncomment the
'        "If m_Turbulence..." line below.)
Private Sub determineVoronoiSearchRadius()
    
    'Regardless of input values, we must always search at least 1px in all directions.
    m_VoronoiSearchRadius = 1
    
    If m_shadingMode > SHADE_F1 Then m_VoronoiSearchRadius = m_VoronoiSearchRadius + 1
    'If m_Turbulence > 0.5 Then m_VoronoiSearchRadius = m_VoronoiSearchRadius + 1
    
End Sub

'Set the mode used to calculate distance
Public Sub setDistanceMode(ByVal newMode As VORONOI_DISTANCE)
    m_distanceMode = newMode
End Sub

'Set the mode used to calculate shading
Public Sub setShadingMode(ByVal newMode As VORONOI_SHADING)

    m_shadingMode = newMode
    
    'Update the Voronoi search radius to account for the new shading mode
    determineVoronoiSearchRadius
    
End Sub

'Given a location IN THE SOURCE IMAGE, return the INDEX of the nearest point in the Voronoi diagram.
Public Function getNearestPointIndex(ByVal srcX As Long, ByVal srcY As Long, Optional ByRef secondNearestPointIndex As Long = -1) As Long
    
    'Start by finding the (x, y) coordinates of the relevant cell
    Dim cellX As Long, cellY As Long
    cellX = (srcX * m_numColumns) \ m_imgWidth
    cellY = (srcY * m_numRows) \ m_imgHeight
    
    'Search neighboring cells to find the closest point, and possibly the second-closest point as well.
    ' (Note: assigning Long-type hex declarations to Singles makes me uncomfortable, but VB doesn't seem to mind,
    '        so I'm running with it.)
    Dim minDistance As Single, minDistance2 As Single
    minDistance = &HEFFFFFF
    minDistance2 = &HEFFFFFF
    
    Dim nearestIndex As Long, nearestIndex2 As Long
    
    'Start by determining the valid min/max indices for our search.  The search radius required for proper operation
    ' varies according to certain input parameters; the m_VoronoiSearchRadius will have been automatically updated by
    ' any relevant functions prior to being utilized here.
    Dim searchMinX As Long, searchMaxX As Long, searchMinY As Long, searchMaxY As Long
    
    searchMinX = cellX - m_VoronoiSearchRadius
    If searchMinX < 0 Then searchMinX = 0
    
    searchMaxX = cellX + m_VoronoiSearchRadius
    If searchMaxX > m_numColumns - 1 Then searchMaxX = m_numColumns - 1
    
    searchMinY = cellY - m_VoronoiSearchRadius
    If searchMinY < 0 Then searchMinY = 0
    
    searchMaxY = cellY + m_VoronoiSearchRadius
    If searchMaxY > m_numRows - 1 Then searchMaxY = m_numRows - 1
    
    'Search all neighboring cells for the nearest Voronoi point
    Dim curDistance As Single, curShadeDistance As Single
    
    Dim x As Long, y As Long
    For x = searchMinX To searchMaxX
    For y = searchMinY To searchMaxY
    
        'Find the distance to this point, using the method requested by the user
        Select Case m_distanceMode
        
            Case DIST_CARTESIAN
                If m_shadingMode = NO_SHADE Then
                    curDistance = cartesianDistanceShortcut(srcX, srcY, vPoints(x, y).x, vPoints(x, y).y)
                Else
                    curDistance = cartesianDistance(srcX, srcY, vPoints(x, y).x, vPoints(x, y).y)
                End If
            
            Case DIST_MANHATTAN
                curDistance = manhattanDistance(srcX, srcY, vPoints(x, y).x, vPoints(x, y).y)
            
            Case DIST_CHEBYSHEV
                curDistance = chebyshevDistance(srcX, srcY, vPoints(x, y).x, vPoints(x, y).y)
            
        End Select
        
        'Check to see if this is the minimum recorded distance for this Voronoi cell
        If (curDistance <= minDistance) Then
        
            'As we are now updating the nearest point, we can also update the second-nearest point using the existing
            ' nearest-point distance.
            If minDistance < minDistance2 Then
                minDistance2 = minDistance
                nearestIndex2 = nearestIndex
            End If
            
            'Update the nearest distance and index markers
            minDistance = curDistance
            nearestIndex = getCellIndex(x, y)
            
        ElseIf (curDistance <= minDistance2) Then
            minDistance2 = curDistance
            nearestIndex2 = getCellIndex(x, y)
        End If
        
    Next y
    Next x
    
    'Update the maximum distance cache for this cell as necessary
    curShadeDistance = getShadingDistance(minDistance, minDistance2)
    If curShadeDistance > vPointsMaxDistance(nearestIndex) Then vPointsMaxDistance(nearestIndex) = curShadeDistance
    
    'Return the nearest and second-nearest indices for this pixel
    secondNearestPointIndex = nearestIndex2
    getNearestPointIndex = nearestIndex
    
End Function

'Given a pixel location (x, y) and a Voronoi point index, return the distance between the two using the current
' class-wide distance formula.
Public Function getDistance(ByVal srcX As Long, ByVal srcY As Long, ByVal vPointIndex As Long) As Single

    'Start by finding the (x, y) coordinates of the relevant cell
    Dim cellCoords As POINTAPI
    cellCoords = getVoronoiCoordinates(vPointIndex)
    
    'Find the distance to this point, using the method requested by the user
    Select Case m_distanceMode
    
        Case DIST_CARTESIAN
            If m_shadingMode = NO_SHADE Then
                getDistance = cartesianDistanceShortcut(srcX, srcY, cellCoords.x, cellCoords.y)
            Else
                getDistance = cartesianDistance(srcX, srcY, cellCoords.x, cellCoords.y)
            End If
        
        Case DIST_MANHATTAN
            getDistance = manhattanDistance(srcX, srcY, cellCoords.x, cellCoords.y)
        
        Case DIST_CHEBYSHEV
            getDistance = chebyshevDistance(srcX, srcY, cellCoords.x, cellCoords.y)
        
    End Select

End Function

'Different shade methods calculate shading distance differently.  Call this function with the distance to the nearest
' and second-nearest Voronoi points, and it will return a shade calculation based on the currently selected shade method.
Private Function getShadingDistance(Optional ByRef nearestDistance As Single, Optional ByRef secondNearestDistance As Single) As Single

     Select Case m_shadingMode
        
            Case NO_SHADE
                getShadingDistance = nearestDistance
        
            Case SHADE_F1
                getShadingDistance = nearestDistance
            
            Case SHADE_F2
                getShadingDistance = secondNearestDistance
            
            Case SHADE_F2_MINUS_F1
                getShadingDistance = secondNearestDistance - nearestDistance
            
            Case SHADE_F1_DIVIDE_F2
                If secondNearestDistance <> 0 Then
                    getShadingDistance = nearestDistance / secondNearestDistance
                Else
                    getShadingDistance = nearestDistance
                End If
        
        End Select

End Function

'Given a 2D coordinate in the master point array, return a unique 1-D index for that point
Private Function getCellIndex(ByVal cellX As Long, ByVal cellY As Long)
    getCellIndex = cellX + (cellY * m_numColumns)
End Function

'If external functions need to know how many Voronoi points are possible, they can use this function
Public Function getTotalNumOfVoronoiPoints() As Long
    getTotalNumOfVoronoiPoints = m_numRows * m_numColumns
End Function

'If external functions want to know the maximum distance for a given cell, they can use this function
Public Function getMaxDistanceForCell(ByVal pointIndex As Long) As Single
    getMaxDistanceForCell = vPointsMaxDistance(pointIndex)
End Function

'Given a pixel location (x, y), and optionally the nearest and/or second nearest Voronoi point index, return a shading value
' for that pixel using the class-wide current shading method.
Public Function getShadingValue(ByVal srcX As Long, ByVal srcY As Long, Optional ByVal nearestPointIndex As Long = -1, Optional ByVal secondNearestPointIndex As Long = -1) As Single

    'If shading is not active, return 1 and exit
    If m_shadingMode = NO_SHADE Then
        getShadingValue = 1
        Exit Function
    End If

    Dim vDistance1 As Single, vDistance2 As Single
    
    'Retrieve the (x, y) coordinates of the nearest Voronoi point.  Note that if caller already cached this value in a previous
    ' step, they can supply it as the nearestPointIndex parameter - that makes this function much faster!
    
    'Note that in the special case of shading using only the F2 value (the 2nd-closest point), we can completely
    ' avoid calculating the distance to the nearest point!
    If m_shadingMode <> SHADE_F2 Then
    
        'If the user did not supply a point index, calculate one now
        If nearestPointIndex = -1 Then nearestPointIndex = getNearestPointIndex(srcX, srcY)
        
        'Find the distance to this point, using the method requested by the user.  (This step could be avoided
        ' by allowing the user to also cache distance - I haven't done that at present, but it's a possibility for
        ' future optimizations.)
        vDistance1 = getDistance(srcX, srcY, nearestPointIndex)
        
    End If
    
    'If the class is using a shading method that also requires knowledge of F2, find that now, using steps
    ' identical to those above, but with F2 instead of F1.  (As before, if the shading method relies only the distance
    ' to the nearest point, we can completely skip this step.
    If m_shadingMode <> SHADE_F1 Then
    
        If secondNearestPointIndex = -1 Then
            getNearestPointIndex srcX, srcY, secondNearestPointIndex
        End If
        
        'Find the distance to this point, using the method requested by the user.  (This step could be avoided
        ' by allowing the user to also cache distance - I haven't done that at present, but it's a possibility for
        ' future optimizations.)
        vDistance2 = getDistance(srcX, srcY, secondNearestPointIndex)
    
    End If
    
    'If the max distance is greater than 0, return it.  If it equals zero (e.g. the source location lies directly
    ' atop a Voronoi point), the calculation will hit a divide-by-zero error, so catch that case and return 1 explicitly.
    If vPointsMaxDistance(nearestPointIndex) > 0 Then
    
        'Use the getShadingDistance function to calculate a shade value for this pixel, per the user's requested
        ' shading method.
        getShadingValue = (getShadingDistance(vDistance1, vDistance2) / vPointsMaxDistance(nearestPointIndex))
        
        'Certain shade functions return inverted results; catch these in advance, so the caller can deal with all
        ' output universally, instead of having to manually reverse values.
        If (m_shadingMode = SHADE_F1) Or (m_shadingMode = SHADE_F1_DIVIDE_F2) Then getShadingValue = 1 - getShadingValue
        
    Else
        getShadingValue = 1
    End If
    
End Function

'If external functions need the coordinates of a given Voronoi points, they can use this function
Friend Function getVoronoiCoordinates(ByVal pointIndex As Long) As POINTAPI
    
    If pointIndex < getTotalNumOfVoronoiPoints() Then
        
        'Solve for the row and column of the index in question
        Dim vRow As Long, vColumn As Long
        vRow = pointIndex \ m_numColumns
        vColumn = pointIndex Mod m_numColumns
        
        'Return the coordinate at that point
        getVoronoiCoordinates = vPoints(vColumn, vRow)
        
    End If
    
End Function

Private Sub Class_Initialize()
    m_PointsInitialized = False
    m_distanceMode = DIST_CARTESIAN
    m_shadingMode = NO_SHADE
End Sub

'Return the distance between two points, but ignore the square root function; if calculating something simple, like "minimum distance only",
' we only need relative values - not absolute ones - so we can skip that step for a slight performance boost.
Private Function cartesianDistanceShortcut(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Single
    cartesianDistanceShortcut = (x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)
End Function

Private Function cartesianDistance(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Single
    cartesianDistance = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
End Function

'"Manhattan Distance" is the walking distance between two points, no angular traveling allowed (e.g. only pure
' horizontal and vertical movement).
Private Function manhattanDistance(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Single
    manhattanDistance = Abs(x1 - x2) + Abs(y1 - y2)
End Function

'"Chebyshev Distance" is similar to "Manhattan Distance", above, except that pure diagonal movement is now allowed.
' It is sometimes called "Chessboard Distance", since it matches the number of moves a King piece must make to travel
' between two points.
Private Function chebyshevDistance(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Single
    If Abs(x1 - x2) > Abs(y1 - y2) Then
        chebyshevDistance = Abs(x1 - x2)
    Else
        chebyshevDistance = Abs(y1 - y2)
    End If
End Function
