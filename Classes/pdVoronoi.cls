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
'Copyright 2014-2026 by Tanner Helland
'Created: 14/July/14
'Last updated: 07/December/20
'Last update: 2x performance improvements! yay!
'Dependencies: pdRandomize (for randomizing point distribution)
'
'While this class is called pdVoronoi, it's primarily a Worley Noise implementation...
' (https://en.wikipedia.org/wiki/Worley_noise),
' ...that uses only the Voronoi aspects most relevant to image processing...
' (https://en.wikipedia.org/wiki/Voronoi_diagram).
'
'For a full Voronoi toolkit (including Delaunay triangulation), you'll need to look elsewhere,
' but if all you need is the bits relevant to generating Worley Noise, you're in for a treat,
' because this class is mostly free of dependencies to other PhotoDemon code - so have at it!
'
'Voronoi diagrams work by taking a series of points, and for the relevant space around them,
' finding the nearest Voronoi point to each location.  When performed at a sufficiently detailed
' level (e.g. for each pixel in a grid), you are left with a puzzle-like appearance a la
' https://en.wikipedia.org/wiki/Voronoi_diagram#mediaviewer/File:Euclidean_Voronoi_Diagram.png
'
'For image processing, Voronoi noise is an especially helpful technique for image segmentation.
' Its biggest issue is that it is painfully slow - comparing each pixel in an image to thousands
' (or more) Voronoi points, until you find the nearest one, is an unpleasant exercise that brings
' even high-end PCs to a grinding halt.
'
'This class circumvents this is by partitioning the initial array of Voronoi points into a grid,
' then randomizing each point within a predetermined grid radius only.  While this limits the
' exoticness of the resulting Voronoi diagram, it allows us to search only a relevant
' neighborhood for each pixel, which in turn makes processing a photograph-sized image a task
' that takes several seconds instead of several hours - a worthwhile improvement!
'
'Note also that this class has been aggressively optimized for performance, which explains some
' of the otherwise poor implementation decisions (e.g. many tasks were originally separate into
' standalone functions, but some of those functions are called inside tight per-pixel loops so
' I've manually inlined them since it makes a *big* difference on 20+ megapixel images).
'
'Three standard distance methods (cartesian, manhattan, and chebyshev) are provided.
' Set the desired distance method using the SetDistanceMode function, and make sure to set it
' *before* calling the GetNearestPointIndex() function!
'
'Some helper functions exist to aid with image-processing-specific tasks like cell shading,
' and these rely on caching certain values within the heavily used GetNearestPointIndex()
' function.  If you don't need the data these caches provide, feel free to comment out everything
' related to the m_vPointsMaxDistance() arrays.  It won't result in a huge speed gain, but you'll
' reduce cache thrashing (which never hurts).
'
'Finally, many thanks to Robert Rayment, who did extensive profiling and research on various
' Voronoi implementations before I started work on this class.  His comments were invaluable in
' determining the shape and style of this class's interface.  (FYI, Robert's PaintRR app has a
' much simpler approach to this filter - check it out if PD's method seems like overkill!
' Link here: http://rrprogs.com/)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Available distance calculation methods
Public Enum PD_VoronoiDistance
    vd_Cartesian = 0
    vd_Manhattan = 1
    vd_Chebyshev = 2
End Enum

#If False Then
    Private Const vd_Cartesian = 0, vd_Manhattan = 1, vd_Chebyshev = 2
#End If

'pdVoronoi supports a number of initial point patterns, which makes for more interesting effects
Public Enum PD_VoronoiPattern
    vp_Square = 0
    vp_Diamond = 1
    vp_Hexagon = 2
    vp_Braid = 3
    vp_Chain = 4
    vp_Cross = 5
    vp_Leaves = 6
    vp_Quilt = 7
    vp_Ragged = 8
    vp_Weave = 9
    vp_Experimental = 10     'Don't expose in production!
    [vp_Count] = 10
End Enum

#If False Then
    Private Const vp_Square = 0, vp_Diamond = 1, vp_Hexagon = 2, vp_Braid = 3, vp_Chain = 4, vp_Cross = 5, vp_Leaves = 6, vp_Quilt = 7, vp_Ragged = 8, vp_Weave = 9
    Private Const vp_Experimental = 10, vp_Count = 11
#End If

'Available shading methods
Public Enum PD_VoronoiShading
    vs_NoShade = 1
    vs_ShadeF1 = 2
    vs_ShadeF2 = 3
    vs_ShadeF2MinusF1 = 4
    vs_ShadeF1DivF2 = 5
    vs_ShadeSqrF1DivF2 = 6
End Enum

#If False Then
    Private Const vs_NoShade = 1, vs_ShadeF1 = 2, vs_ShadeF2 = 3, vs_ShadeF2MinusF1 = 4, vs_ShadeF1DivF2 = 5, vs_ShadeSqrF1DivF2 = 6
#End If

Private Type VoronoiData
    x As Single     'center-point of this Voronoi cell
    y As Single
    maxDistance As Single   'max distance for this cell, used to easily calculate shading on a [0, 1] scale
    IsActive As Boolean     'points can be "deactivated", which creates more varied output
End Type

'This m_vPoints() array will store the coordinates of each point in the Voronoi diagram
Private m_vPoints() As VoronoiData

'Size of a given cell (in one dimension), as supplied by the user
Private m_cellSize As Long

'Number of rows and columns in the diagram.
Private m_numRows As Long, m_numColumns As Long, m_NumPoints As Long

'Size of the image associated with this diagram
Private m_imgWidth As Long, m_imgHeight As Long

'Points have been successfully initialized
Private m_PointsInitialized As Boolean

'Technique used to calculate distance between points; this is set via the SetDistanceMode() sub
Private m_distanceMode As PD_VoronoiDistance

'Technique used to calculate shading; this is set via the SetShadingMode() sub
Private m_shadingMode As PD_VoronoiShading

'Because VB6 doesn't short-circuit, it's faster to set some combination-style flags in advance,
' to avoid the need for complex branches inside loops
Private m_invertShading As Boolean

'Turbulence currently used for the function
Private m_Turbulence As Single

'Radius of nearby points to search.  For certain functions, we can get away with only searching the nearest 9 points;
' unfortunately, certain combinations of patterns require a larger search area.
Private m_VoronoiSearchRadius As Long

'Search boundaries (in x and y directions) are pre-calculated; this spares a ton of bounds-checking
' inside pixel scans.
Private Type BoundsCheck
    bcMin As Long
    bcMax As Long
End Type

Private m_XBounds() As BoundsCheck, m_YBounds() As BoundsCheck

'Row/column lookups are also pre-cached (because \ and % are expensive)
Private m_XLut() As Long, m_YLut() As Long

'PRNG
Private m_Random As pdRandomize

'Given a DIB and the user's requested number of rows and columns, populate an initial m_vPoints() array.
' In the future, I may allow the user to supply a specific point pattern, e.g. "Triangle" or "Square" or "Hexagon".
' Right now, squares are assumed, and the passed pointPattern value affects nothing.
Friend Sub InitPoints(ByVal cellSize As Long, ByVal srcImageWidth As Long, srcImageHeight As Long, Optional ByVal pointPattern As Long = 0)

    'Store the cell size
    m_cellSize = cellSize
    
    'Cache the image's width and height, as we'll need them multiple times in the future.  (Because these values are cached,
    ' this initialization function must be called again if the image's dimensions change!)
    m_imgWidth = srcImageWidth
    m_imgHeight = srcImageHeight
    
    'Calculate the number of rows and columns in the array, as a function of cell size and image size
    m_numRows = m_imgHeight \ m_cellSize
    m_numColumns = m_imgWidth \ m_cellSize
    
    'Enforce a minimum row and column size
    If (m_numRows < 1) Then m_numRows = 1
    If (m_numColumns < 1) Then m_numColumns = 1
    m_NumPoints = m_numRows * m_numColumns
    
    'Resize central Voronoi data collection
    If (UBound(m_vPoints) <> m_NumPoints - 1) Then ReDim m_vPoints(0 To m_NumPoints - 1) As VoronoiData
    
    'Manually fill the max distance array with near-zero-but-not-actually-zero values.
    ' (This incurs a static startup cost, but it lets us skip DBZ checks in the inner pixel loop,
    ' where perf is much more sensitive.)
    '
    'We also want to mark all points as initially active.  Some may be deactivated as part of optimizations.
    Dim x As Long, y As Long, targetCell As Long
    For x = 0 To m_NumPoints - 1
        m_vPoints(x).maxDistance = 0.000001!
        m_vPoints(x).IsActive = True
    Next x
    
    'So that each point is centered nicely inside a cell, we'll initialize them to half-width/height values
    Dim hCellSize As Single
    hCellSize = cellSize / 2
    
    'Populate initial point positions
    Dim invCol As Single, invRow As Single
    If (m_numRows <> 0) Then invRow = 1! / m_numRows
    If (m_numColumns <> 0) Then invCol = 1! / m_numColumns
    
    For y = 0 To m_numRows - 1
    For x = 0 To m_numColumns - 1
        targetCell = y * m_numColumns + x
        m_vPoints(targetCell).x = (x * m_imgWidth) * invCol + hCellSize
        m_vPoints(targetCell).y = (y * m_imgHeight) * invRow + hCellSize
    Next x
    Next y
    
    'Initialize row/column lookup tables
    If (UBound(m_XLut) <> m_imgWidth - 1) Then ReDim m_XLut(0 To m_imgWidth - 1) As Long
    If (UBound(m_YLut) <> m_imgHeight - 1) Then ReDim m_YLut(0 To m_imgHeight - 1) As Long
    For x = 0 To m_imgWidth - 1
        m_XLut(x) = (x * m_numColumns) \ m_imgWidth
    Next x
    For y = 0 To m_imgHeight - 1
        m_YLut(y) = (y * m_numRows) \ m_imgHeight
    Next y
    
    'Note that initialization was succesful
    m_PointsInitialized = True

End Sub

'Randomize the stored point array by some set amount.
' Turbulence is a value on the scale [0, 1]; 1 will result in maximum randomization
' Seed is optional; the same seed will result in the same diagram, with the exception of seed value "0.0" - this tells the
' randomizer to choose its own random seed.
Friend Function RandomizePoints(ByVal fxTurbulence As Double, Optional ByVal rndSeed As Double = 0#) As Boolean

    'Make sure the point array was successfully initialized
    If (Not m_PointsInitialized) Then Exit Function
    
    'Seed the randomizer
    If (rndSeed = 0#) Then
        m_Random.SetSeed_AutomaticAndRandom
    Else
        m_Random.SetSeed_Float rndSeed
    End If
    
    'Cache the turbulence value
    m_Turbulence = fxTurbulence
    
    'Perturb each point in the array by an amount proportional to the cell size; at max turbulence, points can
    ' be perturbed by "cell size / 2", times the user's turbulence parameter (which is on the range [0, 1])
    Dim tmpCellSize As Single
    tmpCellSize = (m_cellSize / 2) * fxTurbulence
    
    Dim x As Long, y As Long
    For y = 0 To m_numRows - 1
    For x = 0 To m_numColumns - 1
        With m_vPoints(y * m_numColumns + x)
            .x = .x + (1# - (m_Random.GetRandomFloat_WH * 2#)) * tmpCellSize
            .y = .y + (1# - (m_Random.GetRandomFloat_WH * 2#)) * tmpCellSize
        End With
    Next x
    Next y
    
    'Update the Voronoi search radius to account for the new turbulence parameter
    DetermineVoronoiSearchRadius
    
    RandomizePoints = True

End Function

'Internal function for determining optimal Voronoi search radius.  If an F2 distance is involved, or the turbulence parameter
' is quite high, the search radius must be extended.  This function will automatically be called after the shading mode or
' turbulence parameter is changed.
' (NOTE: in PhotoDemon, I limit the turbulence factor to 1/2 of the cell size, which saves us from having to extend the
'        search radius due to turbulence.  Individuals needing greater turbulence will thus need to uncomment the
'        "If m_Turbulence..." line below.)
Private Sub DetermineVoronoiSearchRadius(Optional ByVal startingRadius As Long = 1)
    
    'Regardless of input values, we must always search at least +/-1 cell in all directions.
    m_VoronoiSearchRadius = startingRadius
    
    If (m_shadingMode > vs_ShadeF1) Then m_VoronoiSearchRadius = m_VoronoiSearchRadius + 1
    'If (m_Turbulence > 0.5) Then m_VoronoiSearchRadius = m_VoronoiSearchRadius + 1
    
    'After determining Voronoi search radius, we can precalculate specific bounds for
    ' each x/y position in the table.  This spares the need for complex If/Then checks
    ' on the inner loop
    Dim x As Long, y As Long
    Dim lbCheck As Long, ubCheck As Long
    If (UBound(m_XBounds) <> m_numColumns - 1) Then ReDim m_XBounds(0 To m_numColumns - 1) As BoundsCheck
    For x = 0 To m_numColumns - 1
        lbCheck = x - m_VoronoiSearchRadius
        If (lbCheck < 0) Then lbCheck = 0
        m_XBounds(x).bcMin = lbCheck
        ubCheck = x + m_VoronoiSearchRadius
        If (ubCheck > m_numColumns - 1) Then ubCheck = m_numColumns - 1
        m_XBounds(x).bcMax = ubCheck
    Next x
    
    If (UBound(m_YBounds) <> m_numRows - 1) Then ReDim m_YBounds(0 To m_numRows - 1) As BoundsCheck
    For y = 0 To m_numRows - 1
        lbCheck = y - m_VoronoiSearchRadius
        If (lbCheck < 0) Then lbCheck = 0
        m_YBounds(y).bcMin = lbCheck
        ubCheck = y + m_VoronoiSearchRadius
        If (ubCheck > m_numRows - 1) Then ubCheck = m_numRows - 1
        m_YBounds(y).bcMax = ubCheck
    Next y
    
End Sub

'Set the mode used to calculate distance
Friend Sub SetDistanceMode(ByVal newMode As PD_VoronoiDistance)
    m_distanceMode = newMode
End Sub

Friend Sub SetInitialPattern(ByVal newPattern As PD_VoronoiPattern)
    
    Dim i As Long, x As Long, y As Long
    
    Select Case newPattern
    
        Case vp_Square
            For i = 0 To m_NumPoints - 1
                m_vPoints(i).IsActive = True
            Next i
        
        Case vp_Diamond
            For i = 0 To m_NumPoints - 1
                y = i \ m_numColumns
                x = i Mod m_numColumns
                m_vPoints(i).IsActive = ((x + y) And 1)
            Next i
        
        Case vp_Hexagon
            For i = 0 To m_NumPoints - 1
                y = i \ m_numColumns
                x = i Mod m_numColumns
                If (y And 1) Then
                    m_vPoints(i).IsActive = ((x And 3) = 0)
                Else
                    m_vPoints(i).IsActive = (((x + 2) And 3) = 0)
                End If
            Next i
            
        Case vp_Braid
            For i = 0 To m_NumPoints - 1
                y = i \ m_numColumns
                x = i Mod m_numColumns
                If (y And 1) Then
                    m_vPoints(i).IsActive = (x And 1)
                Else
                    m_vPoints(i).IsActive = ((x + 1) And 3)
                End If
            Next i
            
        Case vp_Chain
            For i = 0 To m_NumPoints - 1
                y = i \ m_numColumns
                x = i Mod m_numColumns
                If (y And 1) = 0 Then m_vPoints(i).IsActive = (x And 2)
            Next i
            
        Case vp_Cross
            For i = 0 To m_NumPoints - 1
                y = i \ m_numColumns
                x = i Mod m_numColumns
                If (y And 1) = 0 Then
                    m_vPoints(i).IsActive = (x And 3)
                Else
                    m_vPoints(i).IsActive = ((x + 2) And 3)
                End If
            Next i
        
        Case vp_Leaves
            For i = 0 To m_NumPoints - 1
                y = i \ m_numColumns
                x = i Mod m_numColumns
                If (x Mod 3) = 1 Then
                    m_vPoints(i).IsActive = ((y Mod 3) = 0)
                Else
                    m_vPoints(i).IsActive = ((y Mod 3) = 1)
                End If
            Next i
            
        Case vp_Quilt
            For i = 0 To m_NumPoints - 1
                y = i \ m_numColumns
                x = i Mod m_numColumns
                m_vPoints(i).IsActive = (y And 1) Or (x And 1)
            Next i
            
        Case vp_Ragged
            For i = 0 To m_NumPoints - 1
                y = i \ m_numColumns
                x = i Mod m_numColumns
                If (y Mod 3) = 0 Then
                    m_vPoints(i).IsActive = (x And 3)
                ElseIf (y Mod 3) = 1 Then
                    m_vPoints(i).IsActive = ((x + 1) And 3)
                Else
                    m_vPoints(i).IsActive = (x And 1)
                End If
            Next i
            
        Case vp_Weave
            For i = 0 To m_NumPoints - 1
                y = i \ m_numColumns
                x = i Mod m_numColumns
                If (y Mod 3) = 0 Then
                    m_vPoints(i).IsActive = (x And 1)
                ElseIf (y Mod 3) = 1 Then
                    m_vPoints(i).IsActive = ((x + 1) And 1)
                Else
                    m_vPoints(i).IsActive = False
                End If
            Next i
            
        Case vp_Experimental
            For i = 0 To m_NumPoints - 1
                y = i \ m_numColumns
                x = i Mod m_numColumns
                If (y Mod 3) = 0 Then
                    m_vPoints(i).IsActive = ((x Mod 3) = 0)
                ElseIf (y Mod 3) = 1 Then
                    m_vPoints(i).IsActive = ((x Mod 3) = 1)
                Else
                    m_vPoints(i).IsActive = ((x Mod 3) = 0)
                End If
                
            Next i
        
    End Select
    
End Sub

'Set the mode used to calculate shading
Friend Sub SetShadingMode(ByVal newMode As PD_VoronoiShading)
    m_shadingMode = newMode
    m_invertShading = (m_shadingMode = vs_ShadeF1) Or (m_shadingMode = vs_ShadeF1DivF2) Or (m_shadingMode = vs_ShadeSqrF1DivF2)
End Sub

'NOTE: you *MUST* call SetInitialPattern(), above, before calling this function for
' correct results.
'
'Set density on a [0, 1] scale.  At full density, all grid cells will contain 1 point.
' As you reduce density, cells will be "blanked out".  Note, however, that each cell must
' have at least two neighboring cells that contain points; without this, the search radius
' of each cell would need to be made much more aggressive, which has serious perf implications.
Friend Sub SetDensity(ByVal newDensity As Single)
    
    Dim i As Long
    
    If (newDensity < 1!) Then
        
        'For perf reasons, don't allow more than 90% of points to be removed
        If (newDensity < 0.1) Then newDensity = 0.1
        
        'Randomly deactivate points at a ratio corresponding to the specified density
        For i = 0 To m_NumPoints - 1
            If m_vPoints(i).IsActive Then m_vPoints(i).IsActive = (m_Random.GetRandomFloat_WH() <= newDensity)
        Next i
        
    End If
    
    'Note: you MUST call FinalizeParameters() after calling this function, or the class
    ' will not function correctly.

End Sub

'This function *MUST* be called at least once prior to retrieving actual Voronoi values.
' It will automatically determine things like grid search radius, without which this
' class will not function correctly.
Friend Sub FinalizeParameters()

    'Every cell needs to touch at least two active cells (without this, some rendering models
    ' break down as they need both an f1 and f2 distance).  To ensure this condition is met,
    ' we're gonna quickly iterate the table and find the smallest search radius that ensures
    ' all cells in the collection touch at least two active cells.
    Dim x As Long, y As Long, i As Long, j As Long, numCellsFound As Long
    
    Dim startingRadius As Long
    startingRadius = 1
    
    'Before searching, reset the search radius to its default value
    DetermineVoronoiSearchRadius startingRadius
    
    'Iterate all cells
    Dim searchMinX As Long, searchMaxX As Long, searchMinY As Long, searchMaxY As Long
    For y = 0 To m_numRows - 1
    For x = 0 To m_numColumns - 1
        
StartSearchAgain:
        
        searchMinX = m_XBounds(x).bcMin
        searchMaxX = m_XBounds(x).bcMax
        searchMinY = m_YBounds(y).bcMin
        searchMaxY = m_YBounds(y).bcMax
        
        'Count active cells
        numCellsFound = 0
        For j = searchMinY To searchMaxY
        For i = searchMinX To searchMaxX
            If m_vPoints(j * m_numColumns + i).IsActive Then
                numCellsFound = numCellsFound + 1
                If (numCellsFound >= 2) Then GoTo ContinueSearching
            End If
        Next i
        Next j
        
        'If this cell failed the search criteria, increase search radius and immediately try again
        If (numCellsFound < 2) Then
            startingRadius = startingRadius + 1
            DetermineVoronoiSearchRadius startingRadius
            GoTo StartSearchAgain
        End If

ContinueSearching:

    Next x
    Next y

End Sub

'Given a location IN THE SOURCE IMAGE, return the INDEX of the nearest point in the Voronoi diagram.
Friend Function GetNearestPointIndex(ByVal srcX As Long, ByVal srcY As Long, Optional ByRef secondNearestPointIndex As Long = -1) As Long
    
    'Start by finding the (x, y) coordinates of the relevant cell
    Dim cellX As Long, cellY As Long
    cellX = m_XLut(srcX)
    cellY = m_YLut(srcY)
    
    'Search neighboring cells to find the closest point, and possibly the second-closest point as well.
    ' (Note: assigning Long-type hex declarations to Singles makes me uncomfortable, but VB doesn't seem to mind,
    '        so I'm running with it.)
    Dim minDistance As Single, minDistance2 As Single
    minDistance = &HEFFFFFF
    minDistance2 = &HEFFFFFF
    
    'Start by determining the valid min/max indices for our search.  The search radius required for proper operation
    ' varies according to certain input parameters; the m_VoronoiSearchRadius will have been automatically updated by
    ' any relevant functions prior to being utilized here.
    Dim searchMinX As Long, searchMaxX As Long, searchMinY As Long, searchMaxY As Long
    searchMinX = m_XBounds(cellX).bcMin
    searchMaxX = m_XBounds(cellX).bcMax
    searchMinY = m_YBounds(cellY).bcMin
    searchMaxY = m_YBounds(cellY).bcMax
    
    'Search all neighboring cells for the nearest Voronoi point
    Dim curDistance As Single, curShadeDistance As Single, xDist As Single, yDist As Single
    Dim targetCell As Long
    
    Dim x As Long, y As Long
    For y = searchMinY To searchMaxY
    For x = searchMinX To searchMaxX
    
        targetCell = y * m_numColumns + x
        
        'Only check active points; inactive ones are ignored
        If m_vPoints(targetCell).IsActive Then
            
            'Find the distance to this point, using the method requested by the user.
            ' (Note that these distance calculations have been manually in-lined for performance reasons)
            Select Case m_distanceMode
                
                Case vd_Cartesian
                    xDist = srcX - m_vPoints(targetCell).x
                    yDist = srcY - m_vPoints(targetCell).y
                    curDistance = xDist * xDist + yDist * yDist
                    If (m_shadingMode <> vs_NoShade) Then curDistance = Sqr(curDistance)
                    
                Case vd_Manhattan
                    curDistance = Abs(srcX - m_vPoints(targetCell).x) + Abs(srcY - m_vPoints(targetCell).y)
                    
                Case vd_Chebyshev
                    xDist = Abs(srcX - m_vPoints(targetCell).x)
                    yDist = Abs(srcY - m_vPoints(targetCell).y)
                    If (xDist > yDist) Then curDistance = xDist Else curDistance = yDist
                
            End Select
            
            'Check to see if this is the minimum recorded distance for this Voronoi cell
            If (curDistance <= minDistance) Then
            
                'As we are now updating the nearest point, we can also update the second-nearest point using the existing
                ' nearest-point distance.
                If (minDistance < minDistance2) Then
                    minDistance2 = minDistance
                    secondNearestPointIndex = GetNearestPointIndex
                End If
                
                'Update the nearest distance and index markers
                minDistance = curDistance
                GetNearestPointIndex = targetCell
            
            'Second check for 2nd-nearest point (used to calculate cell shading)
            ElseIf (curDistance <= minDistance2) Then
                minDistance2 = curDistance
                secondNearestPointIndex = targetCell
            End If
            
        End If
        
    Next x
    Next y
    
    'After finding the nearest point, update the maximum distance cache for this cell (as necessary)
    ' (This is an in-lined version of "curShadeDistance = GetShadingDistance(minDistance, minDistance2)")
    Select Case m_shadingMode
        Case vs_NoShade
            curShadeDistance = minDistance
        Case vs_ShadeF1
            curShadeDistance = minDistance
        Case vs_ShadeF2
            curShadeDistance = minDistance2
        Case vs_ShadeF2MinusF1
            curShadeDistance = minDistance2 - minDistance
        Case vs_ShadeF1DivF2
            'Branchless DBZ fix
            curShadeDistance = minDistance / (minDistance2 + 0.0000001!)
        Case vs_ShadeSqrF1DivF2
            curShadeDistance = minDistance / (minDistance2 + 0.0000001!)
            curShadeDistance = curShadeDistance * curShadeDistance
    End Select
    
    If (curShadeDistance > m_vPoints(GetNearestPointIndex).maxDistance) Then m_vPoints(GetNearestPointIndex).maxDistance = curShadeDistance
    
    'Return values were already calculated
    
End Function

'Given a pixel location (x, y) and a Voronoi point index, return the distance between the two using the current
' class-wide distance formula.
Friend Function GetDistance(ByVal srcX As Single, ByVal srcY As Single, ByVal vPointIndex As Long) As Single

    'Find the distance to this point, using the method requested by the user
    Select Case m_distanceMode
    
        Case vd_Cartesian
            srcX = srcX - m_vPoints(vPointIndex).x
            srcY = srcY - m_vPoints(vPointIndex).y
            GetDistance = srcX * srcX + srcY * srcY
            If (m_shadingMode <> vs_NoShade) Then GetDistance = Sqr(GetDistance)
            
        Case vd_Manhattan
            GetDistance = Abs(srcX - m_vPoints(vPointIndex).x) + Abs(srcY - m_vPoints(vPointIndex).y)
            
        Case vd_Chebyshev
            srcX = Abs(srcX - m_vPoints(vPointIndex).x)
            srcY = Abs(srcY - m_vPoints(vPointIndex).y)
            If (srcX > srcY) Then GetDistance = srcX Else GetDistance = srcY
            
    End Select

End Function

'If external functions need to know how many Voronoi points are possible, they can use this function
Friend Function GetTotalNumOfVoronoiPoints() As Long
    GetTotalNumOfVoronoiPoints = m_NumPoints
End Function

'If external functions want to know the maximum distance for a given cell, they can use this function
Friend Function GetMaxDistanceForCell(ByVal pointIndex As Long) As Single
    GetMaxDistanceForCell = m_vPoints(pointIndex).maxDistance
End Function

'Given a pixel location (x, y), and the nearest and/or second nearest Voronoi point index,
' return a shading value for that pixel using the current shading method.
Friend Function GetShadingValue(ByVal srcX As Long, ByVal srcY As Long, ByVal nearestPointIndex As Long, ByVal secondNearestPointIndex As Long) As Single

    'If shading is not active, return 1 and exit
    If (m_shadingMode = vs_NoShade) Then
        GetShadingValue = 1!
        
    Else

        Dim vDistance1 As Single, vDistance2 As Single
            
        'Note that in the special case of shading using only the F2 value (the 2nd-closest point),
        ' we can completely avoid calculating the distance to the nearest point!
        If (m_shadingMode <> vs_ShadeF2) Then
        
            'Find the distance to this point, using the method requested by the user.  (This step could be avoided
            ' by allowing the user to also cache distance - I haven't done that at present, but it's a possibility for
            ' future optimizations.)
            vDistance1 = GetDistance(srcX, srcY, nearestPointIndex)
            
        End If
        
        'If the class is using a shading method that also requires knowledge of F2, find that now, using steps
        ' identical to those above, but with F2 instead of F1.  (As before, if the shading method relies only the distance
        ' to the nearest point, we can completely skip this step.
        If (m_shadingMode <> vs_ShadeF1) Then
        
            'Find the distance to this point, using the method requested by the user.  (This step could be avoided
            ' by allowing the user to also cache distance - I haven't done that at present, but it's a possibility for
            ' future optimizations.)
            vDistance2 = GetDistance(srcX, srcY, secondNearestPointIndex)
        
        End If
        
        'Use the GetShadingDistance function to calculate a shade value for this pixel, per the user's requested
        ' shading method.  (Note that the max distance array was intentionally initialized to non-zero values,
        ' so we do *not* need to perform a DBZ check here.)
        Select Case m_shadingMode
            Case vs_NoShade
                GetShadingValue = vDistance1
            Case vs_ShadeF1
                GetShadingValue = vDistance1
            Case vs_ShadeF2
                GetShadingValue = vDistance2
            Case vs_ShadeF2MinusF1
                GetShadingValue = vDistance2 - vDistance1
            Case vs_ShadeF1DivF2
                GetShadingValue = vDistance1 / (vDistance2 + 0.0000001!) 'Branchless DBZ failsafe
            Case vs_ShadeSqrF1DivF2
                GetShadingValue = vDistance1 / (vDistance2 + 0.0000001!)
                GetShadingValue = GetShadingValue * GetShadingValue
        End Select
    
        GetShadingValue = GetShadingValue / m_vPoints(nearestPointIndex).maxDistance
        
        'Certain shade functions return inverted results; catch these in advance, so the caller can deal with all
        ' output universally, instead of having to manually reverse values on their end.
        If m_invertShading Then GetShadingValue = 1! - GetShadingValue
        
    End If
    
End Function

'If external functions need the coordinates of a given Voronoi points, they can use this function.
' (For perf reasons, the incoming index is *NOT* bounds-checked.  It will crash if you pass a bad index!)
Friend Function GetVoronoiCoordinates(ByVal pointIndex As Long) As PointFloat
    GetVoronoiCoordinates.x = m_vPoints(pointIndex).x
    GetVoronoiCoordinates.y = m_vPoints(pointIndex).y
End Function

'Serialization helpers
Friend Function GetPatternCount() As Long
    GetPatternCount = vp_Count
End Function

Friend Function GetPatternIDFromName(ByRef srcPatternName As String) As PD_VoronoiPattern
    If Strings.StringsEqual(srcPatternName, "square", True) Then
        GetPatternIDFromName = vp_Square
    ElseIf Strings.StringsEqual(srcPatternName, "diamond", True) Then
        GetPatternIDFromName = vp_Diamond
    ElseIf Strings.StringsEqual(srcPatternName, "hexagon", True) Then
        GetPatternIDFromName = vp_Hexagon
    ElseIf Strings.StringsEqual(srcPatternName, "braid", True) Then
        GetPatternIDFromName = vp_Braid
    ElseIf Strings.StringsEqual(srcPatternName, "chain", True) Then
        GetPatternIDFromName = vp_Chain
    ElseIf Strings.StringsEqual(srcPatternName, "cross", True) Then
        GetPatternIDFromName = vp_Cross
    ElseIf Strings.StringsEqual(srcPatternName, "leaves", True) Then
        GetPatternIDFromName = vp_Leaves
    ElseIf Strings.StringsEqual(srcPatternName, "quilt", True) Then
        GetPatternIDFromName = vp_Quilt
    ElseIf Strings.StringsEqual(srcPatternName, "ragged", True) Then
        GetPatternIDFromName = vp_Ragged
    ElseIf Strings.StringsEqual(srcPatternName, "weave", True) Then
        GetPatternIDFromName = vp_Weave
    ElseIf Strings.StringsEqual(srcPatternName, "experimental", True) Then
        GetPatternIDFromName = vp_Experimental
    Else
        GetPatternIDFromName = vp_Square
    End If
End Function

Friend Function GetPatternNameFromID(ByVal srcPattern As PD_VoronoiPattern) As String
    Select Case srcPattern
        Case vp_Square
            GetPatternNameFromID = "square"
        Case vp_Diamond
            GetPatternNameFromID = "diamond"
        Case vp_Hexagon
            GetPatternNameFromID = "hexagon"
        Case vp_Braid
            GetPatternNameFromID = "braid"
        Case vp_Chain
            GetPatternNameFromID = "chain"
        Case vp_Cross
            GetPatternNameFromID = "cross"
        Case vp_Leaves
            GetPatternNameFromID = "leaves"
        Case vp_Quilt
            GetPatternNameFromID = "quilt"
        Case vp_Ragged
            GetPatternNameFromID = "ragged"
        Case vp_Weave
            GetPatternNameFromID = "weave"
        Case vp_Experimental
            GetPatternNameFromID = "experimental"
        Case Else
            GetPatternNameFromID = "square"
    End Select
End Function

Friend Function GetPatternUINameFromID(ByVal srcPattern As PD_VoronoiPattern) As String
    If (g_Language Is Nothing) Then Exit Function
    Select Case srcPattern
        Case vp_Square
            GetPatternUINameFromID = g_Language.TranslateMessage("square")
        Case vp_Diamond
            GetPatternUINameFromID = g_Language.TranslateMessage("diamond")
        Case vp_Hexagon
            GetPatternUINameFromID = g_Language.TranslateMessage("hexagon")
        Case vp_Braid
            GetPatternUINameFromID = g_Language.TranslateMessage("braid")
        Case vp_Chain
            GetPatternUINameFromID = g_Language.TranslateMessage("chain")
        Case vp_Cross
            GetPatternUINameFromID = g_Language.TranslateMessage("cross")
        Case vp_Leaves
            GetPatternUINameFromID = g_Language.TranslateMessage("leaves")
        Case vp_Quilt
            GetPatternUINameFromID = g_Language.TranslateMessage("quilt")
        Case vp_Ragged
            GetPatternUINameFromID = g_Language.TranslateMessage("ragged")
        Case vp_Weave
            GetPatternUINameFromID = g_Language.TranslateMessage("weave")
        Case vp_Experimental
            GetPatternUINameFromID = "experimental"
        Case Else
            GetPatternUINameFromID = g_Language.TranslateMessage("square")
    End Select
End Function

Private Sub Class_Initialize()
    
    m_PointsInitialized = False
    m_distanceMode = vd_Cartesian
    m_shadingMode = vs_NoShade
    m_invertShading = False
    Set m_Random = New pdRandomize
    
    'Initialize class-level arrays to simplify bounds-checking on subsequent runs
    ReDim m_vPoints(0) As VoronoiData
    ReDim m_XBounds(0) As BoundsCheck
    ReDim m_YBounds(0) As BoundsCheck
    ReDim m_XLut(0) As Long
    ReDim m_YLut(0) As Long
    
End Sub
