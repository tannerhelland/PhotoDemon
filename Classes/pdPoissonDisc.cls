VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPoissonDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Poisson Disc Sampler
'Copyright 2019-2026 by Tanner Helland
'Created: 14/November/19
'Last updated: 14/November/19
'Last update: initial build
'Dependencies: pdRandomize (for random number generation)
'
'Poisson disc sampling (PDS) is a fast way to supersample a region of 2D space:
' https://en.wikipedia.org/wiki/Supersampling#Poisson_disc
'
'This class uses a fast generative technique known as Bridson's algorithm:
' https://www.cs.ubc.ca/~rbridson/docs/bridson-siggraph07-poissondisk.pdf
'
'What's the relevance of this to a photo editor?  On some adjustments and effects,
' PD uses supersampling to accelerate the pixel sampling process.  Some classes of
' filters work just fine with a representative (but not comprehensive) sample of pixels,
' and this class produces a supersample mapping guaranteed to produce a representative
' sample for any arbitrary radius (r).
'
'An easy example of is the bilateral filter, a la the 2012 paper "A Low-Memory,
' Straightforward and Fast Bilateral Filter Through Subsampling in Spatial Domain":
' http://vcg.isti.cnr.it/Publications/2012/BCCS12/j.1467-8659.2011.02078.x.pdf
'
'Those authors compared PNSR between true bilateral results and supersampled results,
' and PDS was found to be the highest-quality sampler - hence why I use it here.
' Specifically, they said: "Note that the use of Poisson-disk sampling produces the
' closest approximation to the full bilateral filter. The other sampling strategies create
' visual artefacts and pattern-like artefacts, particularly when a regular sampling pattern
' is used."
'
'Distances are currently calculated using standard cartesian distance.  In the future,
' it may be interesting to add manhattan and chebyshev to the mix.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Use our own internal randomizer for point distribution
Private m_Randomize As pdRandomize

'Produce a list of points (indexed using the passed grid), evenly sampled at min radius r,
' for a source surface of width WxH.
'
'Returns: FALSE if the requested radius is too small to produce a usable grid
Friend Function GetDisc(ByRef dstPoints() As PointFloat, ByRef dstNumPoints As Long, ByRef outGrid() As Long, ByRef outGridWidth As Long, ByRef outGridHeight As Long, ByVal radius As Double, ByVal srcWidth As Long, ByVal srcHeight As Long) As Boolean
    
    GetDisc = True
    
    'Initialize the destination point list to a reasonable default
    dstNumPoints = 64
    ReDim dstPoints(0 To dstNumPoints - 1) As PointFloat
    
    'We also need an internal set of "active" points (points for which we are still trying to assign
    ' new neighbors).  This is sorta like a stack.
    Dim activePoints() As PointFloat
    ReDim activePoints(0 To dstNumPoints - 1) As PointFloat
    Dim numActivePoints As Long
    
    Dim maxWidth As Single, maxHeight As Single
    maxWidth = srcWidth - 1!
    maxHeight = srcHeight - 1!
    
    'Next, we need to set up the grid.  Per Bridson's algorithm, the size of each cell should be
    ' r / sqrt(2).
    Dim cellSize As Long
    cellSize = Int(radius / Sqr(2#))
    
    'If cells are too small, the caller needs to brute-force the algorithm instead
    If (cellSize < 1) Then
        GetDisc = False
        Exit Function
    End If
    
    Dim rSquared As Double
    rSquared = radius * radius
    
    'Use the cell size and the source width/height to calculate width/height for the lookup grid.
    ' (Note that we allocate an extra cell in either direction to improve boundary behavior;
    ' in the future, something similar could be done in the left/top direction as well, though this
    ' would add a minor perf hit for array lookups.)
    outGridWidth = Int(srcWidth / cellSize + 0.9999999) + 1
    outGridHeight = Int(srcHeight / cellSize + 0.9999999) + 1
    ReDim outGrid(0 To outGridWidth - 1, 0 To outGridHeight - 1) As Long
    
    'Initialize the grid to some kind of "null" flag.  (Cells can be empty.)
    Dim x As Long, y As Long
    For y = 0 To outGridHeight - 1
    For x = 0 To outGridWidth - 1
        outGrid(x, y) = -1
    Next x
    Next y
    
    'Add an initial random point to both the the active collection.
    ' (Note that we deliberately initialize it to be somewhere in the middle quadrant of
    ' the grid; this improves performance over an initial point added to the periphery,
    ' where points are more likely to fail multiple boundary checks on each addition.)
    Dim tmpPoint As PointFloat
    tmpPoint.x = (outGridWidth \ 4) + (m_Randomize.GetRandomFloat_WH() * (outGridWidth \ 2))
    tmpPoint.y = (outGridHeight \ 4) + (m_Randomize.GetRandomFloat_WH() * (outGridHeight \ 2))
    
    dstPoints(0) = tmpPoint
    dstNumPoints = 1
    
    activePoints(0) = tmpPoint
    numActivePoints = 1
    
    Dim apIndex As Long, ptAdded As Boolean
    Dim rndTheta As Double, rndRadius As Double
    Dim xMin As Long, xMax As Long, yMin As Long, yMax As Long, xGrd As Long, yGrd As Long
    
    'Now we loop endlessly; as long as there are active points to process, we'll try and
    ' add more points to the collection
    Do While (numActivePoints > 0)
        
        'Select a new active point
        apIndex = m_Randomize.GetRandomIntRange_WH(0, numActivePoints - 1)
        
        'We are now going to attempt to add new points around the current one.  If we succeed,
        ' this point gets to stay in the active list, as do any point(s) we add.
        
        'In his original paper (https://www.cs.ubc.ca/~rbridson/docs/bridson-siggraph07-poissondisk.pdf)
        ' Bridson suggests - without evidence, for better or worse - 30 sampling attempts per pixel.
        ' You can obviously try fewer and risk a poorer distribution, at some boost to perf.
        Const MAX_SAMPLING_ATTEMPTS As Long = 30
        ptAdded = False
        
        Dim k As Long
        For k = 1 To MAX_SAMPLING_ATTEMPTS
            
            'Generate a random theta and radius for this test point
            rndTheta = m_Randomize.GetRandomFloat_WH() * PI_DOUBLE
            rndRadius = radius + m_Randomize.GetRandomFloat_WH() * radius
            
            'Construct a matching cartesian position
            tmpPoint.x = activePoints(apIndex).x + rndRadius * Cos(rndTheta)
            tmpPoint.y = activePoints(apIndex).y + rndRadius * Sin(rndTheta)
            
            'Test this point; if it's good, add it and exit the loop
            
            'First, test it against source image dimensions
            If (tmpPoint.x < 0) Or (tmpPoint.y < 0) Then GoTo TryAnotherPoint
            If (tmpPoint.x > maxWidth) Or (tmpPoint.y > maxHeight) Then GoTo TryAnotherPoint
            
            'Next, figure out where this points (hypothetical) grid index lies
            xGrd = Int(tmpPoint.x / cellSize)
            yGrd = Int(tmpPoint.y / cellSize)
            
            'Calculate corresponding loop intervals
            xMin = xGrd - 1
            If (xMin < 0) Then xMin = 0
            yMin = yGrd - 1
            If (yMin < 0) Then yMin = 0
            xMax = xMin + 2
            If (xMax >= outGridWidth) Then xMax = outGridWidth - 1
            yMax = yMin + 2
            If (yMax >= outGridHeight) Then yMax = outGridHeight - 1
            
            For y = yMin To yMax
            For x = xMin To xMax
            
                'Only test non-null points
                If (outGrid(x, y) >= 0) Then
                    If (PDMath.DistanceTwoPointsShortcut(tmpPoint.x, tmpPoint.y, dstPoints(outGrid(x, y)).x, dstPoints(outGrid(x, y)).y) < rSquared) Then GoTo TryAnotherPoint
                End If
            
            Next x
            Next y
            
            'If we're still here, this point is valid!  Add it to both lists and mark its
            ' corresponding grid index.
            outGrid(xGrd, yGrd) = dstNumPoints
            
            If (dstNumPoints > UBound(dstPoints)) Then ReDim Preserve dstPoints(0 To dstNumPoints * 2 - 1) As PointFloat
            dstPoints(dstNumPoints) = tmpPoint
            dstNumPoints = dstNumPoints + 1
            
            If (numActivePoints > UBound(activePoints)) Then ReDim Preserve activePoints(0 To numActivePoints * 2 - 1) As PointFloat
            activePoints(numActivePoints) = tmpPoint
            numActivePoints = numActivePoints + 1
            
            ptAdded = True
            Exit For
            
TryAnotherPoint:
        Next k
        
        'If a point was *not* added, this point should no longer be active.  Remove it from the list.
        If (Not ptAdded) Then
            
            'Fast remove; swap with the trailing point
            activePoints(apIndex) = activePoints(numActivePoints - 1)
            numActivePoints = numActivePoints - 1
            
        End If
        
    Loop
    
End Function

'The point of this class is to return
Private Sub Class_Initialize()
    Set m_Randomize = New pdRandomize
    m_Randomize.SetSeed_AutomaticAndRandom
End Sub
