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
'Copyright ©2013-2014 by Tanner Helland
'Created: 14/July/14
'Last updated: 14/July/14
'Last update: initial build
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
'For image processing, Voronoi noise is an especially helpful technique for image segmentation.  It's biggest issue is that
' it is painfully slow - comparing each of many millions of pixels to thousands (or more) Voronoi points, searching for the
' nearest one, is an exercise that will bring even modern PCs to a grinding halt.
'
'The way we circumvent this is by carefully partitioning our Voronoi points into a grid shape, then randomizing each point
' within a predetermined radius only.  While this limits the exoticness of the resulting Voronoi diagram, it allows us to
' search only a relevant neighborhood for each pixel, which allows us to process a photograph in several seconds instead of
' several hours - a worthwhile improvement!
'
'Three standard distance methods (cartesian, manhattan, and chebyshev) are provided.  Set the desired distance method using
' the setDistanceMode function, and make sure to use it *before* calling the getNearestPointIndex function!
'
'Some helper functions exist to aid with image-processing-specific tasks like cell shading, and these relying on the caching
' of certain values within the heavily used getNearestPointIndex function.  If you don't need the data these caches provide,
' feel free to comment out everything related to the vPointsMaxDistance() array.  It won't result in a huge speed gain, but
' you'll save a bit of cache thrashing, which never hurts.
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

'This vPoints() array will store the coordinates of each point in the Voronoi diagram
Private vPoints() As POINTAPI
Private vPointsMaxDistance() As Single

'Size of a given cell (in one dimension), as supplied by the user
Private m_cellSize As Long

'Number of rows and columns in the diagram.
Private m_numRows As Long, m_numColumns As Long

'Size of the image associated with this diagram
Private m_imgWidth As Long, m_imgHeight As Long

'Points have been successfully initialized
Private m_PointsInitialized As Boolean

'When calculating distances, the class will automatically update these values as necessary.
Private m_maxDistance As Double

'Technique used to calculate distance between points; this is set via the setDistanceMode() sub
Private m_distanceMode As VORONOI_DISTANCE

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
    m_numRows = m_imgWidth / m_cellSize
    m_numColumns = m_imgHeight / m_cellSize
    
    'Resize all arrays to match
    ReDim vPoints(0 To m_numRows - 1, 0 To m_numColumns - 1) As POINTAPI
    ReDim vPointsMaxDistance(0 To getTotalNumOfVoronoiPoints() - 1) As Single
    
    'So that each point is centered nicely inside a cell, we'll make use of half width and height values
    Dim hCellSize As Long
    hCellSize = cellSize \ 2
    
    'Populate initial point positions
    Dim x As Long, y As Long
    For x = 0 To m_numRows - 1
    For y = 0 To m_numColumns - 1
        vPoints(x, y).x = (x / m_numRows) * m_imgWidth + hCellSize
        vPoints(x, y).y = (y / m_numColumns) * m_imgHeight + hCellSize
    Next y
    Next x
    
    'Reset certain cached values
    m_maxDistance = 0
    
    'Note that initialization was succesful
    m_PointsInitialized = True

End Sub

'Randomize the stored point array by some set amount.
' Turbulence is a value on the scale [0, 1]; 1 will result in maximum randomization
' Seed is optional; the same seed will result in the same diagram
Public Function randomizePoints(ByVal fxTurbulence As Double, Optional ByVal rndSeed As Double = 0) As Boolean

    'Make sure the point array was successfully initialized
    If Not m_PointsInitialized Then
        Debug.Print "You haven't initialized the point array yet!"
        Exit Function
    End If
    
    'Seed the randomizer
    Rnd -1
    Randomize rndSeed
    
    'Perturb each point in the array by an amount proportional to the cell size; at max turbulence, points can
    ' be perturbed by "cell size / 2".
    Dim tmpCellSize As Long
    tmpCellSize = m_cellSize \ 2
    
    Dim x As Long, y As Long
    For x = 0 To m_numRows - 1
    For y = 0 To m_numColumns - 1
        vPoints(x, y).x = vPoints(x, y).x + (1 - (Rnd * 2)) * fxTurbulence * tmpCellSize
        vPoints(x, y).y = vPoints(x, y).y + (1 - (Rnd * 2)) * fxTurbulence * tmpCellSize
    Next y
    Next x

    randomizePoints = True

End Function

'Set the mode used to calculate distance
Public Sub setDistanceMode(ByVal newMode As VORONOI_DISTANCE)
    m_distanceMode = newMode
End Sub

'Given a location IN THE SOURCE IMAGE, return the INDEX of the nearest point in the Voronoi diagram.
Public Function getNearestPointIndex(ByVal srcX As Long, ByVal srcY As Long) As Long
    
    'Start by finding the (x, y) coordinates of the relevant cell
    Dim cellX As Long, cellY As Long
    cellX = (srcX / m_imgWidth) * m_numRows
    cellY = (srcY / m_imgHeight) * m_numColumns
    
    'Search the neighboring cells to find the closest point
    Dim minDistance As Double
    minDistance = &HEFFFFFF
    
    Dim nearestIndex As Long
    
    'Start by determining the valid min/max indices for our search
    Dim searchMinX As Long, searchMaxX As Long, searchMinY As Long, searchMaxY As Long
    
    searchMinX = cellX - 1
    If searchMinX < 0 Then searchMinX = 0
    
    searchMinY = cellY - 1
    If searchMinY < 0 Then searchMinY = 0
    
    searchMaxX = cellX + 1
    If searchMaxX > m_numRows - 1 Then searchMaxX = m_numRows - 1
    
    searchMaxY = cellY + 1
    If searchMaxY > m_numColumns - 1 Then searchMaxY = m_numColumns - 1
    
    'Search all neighboring celles for the nearest Voronoi point
    Dim curDistance As Double
    
    Dim x As Long, y As Long
    For x = searchMinX To searchMaxX
    For y = searchMinY To searchMaxY
    
        'Find the distance to this point, using the method requested by the user
        Select Case m_distanceMode
        
            Case DIST_CARTESIAN
                curDistance = cartesianDistanceShortcut(srcX, srcY, vPoints(x, y).x, vPoints(x, y).y)
            
            Case DIST_MANHATTAN
                curDistance = manhattanDistance(srcX, srcY, vPoints(x, y).x, vPoints(x, y).y)
            
            Case DIST_CHEBYSHEV
                curDistance = chebyshevDistance(srcX, srcY, vPoints(x, y).x, vPoints(x, y).y)
            
        End Select
                
        If (curDistance < minDistance) Then
            minDistance = curDistance
            nearestIndex = getCellIndex(x, y)
        End If
    
    Next y
    Next x
        
    'Cache max distance as necessary
    If minDistance > vPointsMaxDistance(nearestIndex) Then vPointsMaxDistance(nearestIndex) = minDistance
        
    'Return the nearest index
    getNearestPointIndex = nearestIndex
    
End Function

'Given a 2D coordinate in the master point array, return a unique 1-D index for that point
Private Function getCellIndex(ByVal cellX As Long, ByVal cellY As Long)
    getCellIndex = (cellY * m_numRows) + cellX
End Function

'If external functions need to know how many Voronoi points are possible, they can use this function
Public Function getTotalNumOfVoronoiPoints() As Long
    getTotalNumOfVoronoiPoints = m_numRows * m_numColumns
End Function

Public Function getShadingValue(ByVal srcX As Long, ByVal srcY As Long, ByRef shadeMethod As Long, Optional ByVal pointIndex As Long = -1) As Double

    Dim vDistance As Double
    
    'For now, assume that the nearest point value is known
    Dim srcPoint As POINTAPI
    srcPoint = getVoronoiCoordinates(pointIndex)
    
    'Find the distance to this point, using the method requested by the user
    Select Case m_distanceMode
    
        Case DIST_CARTESIAN
            vDistance = cartesianDistanceShortcut(srcX, srcY, srcPoint.x, srcPoint.y)
        
        Case DIST_MANHATTAN
            vDistance = manhattanDistance(srcX, srcY, srcPoint.x, srcPoint.y)
        
        Case DIST_CHEBYSHEV
            vDistance = chebyshevDistance(srcX, srcY, srcPoint.x, srcPoint.y)
        
    End Select
    
    getShadingValue = 1 - (vDistance / vPointsMaxDistance(pointIndex))

End Function

'If external functions need the coordinates of a given Voronoi points, they can use this function
Friend Function getVoronoiCoordinates(ByVal pointIndex As Long) As POINTAPI
    
    If pointIndex < getTotalNumOfVoronoiPoints() Then
        
        'Solve for the row and column of the index in question
        Dim vRow As Long, vColumn As Long
        vColumn = pointIndex \ m_numRows
        vRow = pointIndex Mod m_numRows
        
        getVoronoiCoordinates = vPoints(vRow, vColumn)
        
    End If
    
End Function

Private Sub Class_Initialize()
    m_PointsInitialized = False
    m_distanceMode = DIST_CARTESIAN
End Sub

'Return the distance between two points, but ignore the square root function; if calculating something simple, like "minimum distance only",
' we only need relative values - not absolute ones - so we can skip that step for a slight performance boost.
Private Function cartesianDistanceShortcut(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    cartesianDistanceShortcut = (x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)
End Function

'"Manhattan Distance" is the walking distance between two points, no angular traveling allowed (e.g. only pure
' horizontal and vertical movement).
Private Function manhattanDistance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    manhattanDistance = Abs(x1 - x2) + Abs(y1 - y2)
End Function

'"Chebyshev Distance" is similar to "Manhattan Distance", above, except that pure diagonal movement is now allowed.
' It is sometimes called "Chessboard Distance", since it matches the number of moves a King piece must make to travel
' between two points.
Private Function chebyshevDistance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    If Abs(x1 - x2) > Abs(y1 - y2) Then
        chebyshevDistance = Abs(x1 - x2)
    Else
        chebyshevDistance = Abs(y1 - y2)
    End If
End Function
