VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdEdgeDetector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Edge Detector ("Marching squares") Engine
'Copyright 2017-2017 by Tanner Helland
'Created: 05/January/17
'Last updated: 05/January/17
'Last update: initial build
'
'This class supplies the "marching squares" algorithm used by a number of different tools in PD.
'
'To simplify the class (and improve performance), all operations are performed on byte arrays filled
' with interior/exterior data.  This class doesn't care how you generate those arrays, and it treats
' zeroes as exterior markers, and non-zeroes as interior ones.  (For a 32-bpp image, for example,
' you could generate the required array by setting transparent pixel coordinates to 0 and all other
' pixels to 1.)
'
'Also, note that this class does *not* perform edge-handling, by design.  It is up to the caller to
' decide how they want edges handled (typically by adding a row/column of padding around all sides).
'
'This class returns a list of points defining the polygon that encloses the first discovered interior
' region.  Note that the caller supplies the starting point, by design - if you want to match multiple
' regions, you need to manually update the byte array between calls (by removing the "discovered"
' exterior points), or you need to supply different starting points.  This class only ever returns *one*
' region.
'
'Like any array-based tool, this class will be slow inside the IDE.  Please use only when compiled.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'A stack is used to track polygon points enclosing the first discovered region
Private m_Stack() As POINTAPI
Private m_StackPosition As Long
Private m_StackHeight As Long
Private Const INITIAL_STACK_HEIGHT As Long = 4096

'Obviously the function needs a starting x/y position
Private m_InitPoint As POINTAPI

'Marching squares uses predefined direction descriptors.  See https://en.wikipedia.org/wiki/Marching_squares
' These descriptors are populated in Class_Initialize().
Private m_XDirs() As Long, m_YDirs() As Long

'If you don't want to find a starting point yourself, use this helper function to do it for you.
' The starting point is searched for starting in the top-left corner and moving LTR.
'
' RETURNS: TRUE if a starting point is found; FALSE for an empty array.
'          (If FALSE is returned, do *not* attempt further edge analysis of the shape, as the class will crash.)
Friend Function FindStartingPoint(ByRef srcArray() As Byte, ByVal initX As Long, ByVal initY As Long, ByVal finalX As Long, ByVal finalY As Long, ByRef dstStartX As Long, ByRef dstStartY As Long) As Boolean
    
    Dim x As Long, y As Long
    Dim keepSearching As Boolean
    keepSearching = True
    
    For y = initY To finalY
    For x = initX To finalX
    
        If (srcArray(x, y) <> 0) Then
            dstStartX = x
            dstStartY = y
            keepSearching = False
            Exit For
        End If
    
    Next x
        If (Not keepSearching) Then Exit For
    Next y
    
    FindStartingPoint = CBool(Not keepSearching)
    
End Function

'Return a polygon that defines the edges of a region starting from point (startX, startY).  If (startX, startY)
' do not define a valid edge point, the function will return meaningless data.
'
'RETURNS: TRUE if a valid polygon is generated; FALSE otherwise.
Friend Function FindEdges(ByRef srcArray() As Byte, ByVal startX As Long, ByVal startY As Long, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As Boolean

    'Current (x, y) positions.  (These start at startX and startY.)
    Dim x As Long, y As Long
    x = startX
    y = startY
    
    'Next (x, y) *direction*.  These are always one of [-1, 0, 1].
    Dim nextX As Long, nextY As Long
    nextX = 0
    nextY = 0
    
    'Previous (x, y) *direction*.  These are always one of [-1, 0, 1].
    ' We track these so that we can ignore points that lie on an existing line "as we go", instead of
    ' wasting time and energy pushing them onto the stack.
    Dim prevX As Long, prevY As Long
    prevX = 1000000000
    prevY = 1000000000
    
    'Index into our precalculated direction arrays.
    Dim i As Long
    
    'Start searching for contour points!
    Do
        
        'For the current pixel, determine an index into the marching squares lookup table.  Note that this
        ' table is constructed very specifically, which is what makes the following lookup technique work.
        ' (See: https://en.wikipedia.org/wiki/Marching_squares#Basic_algorithm)
        ' Note also that other languages could compute this more efficiently using bit-shifting..
        i = 0
        If (srcArray(x - 1, y - 1) <> 0) Then i = i + 1
        If (srcArray(x, y - 1) <> 0) Then i = i + 2
        If (srcArray(x - 1, y) <> 0) Then i = i + 4
        If (srcArray(x, y) <> 0) Then i = i + 8
        
        'Now that we have an index into the table, use that index to compute the actual physical
        ' location of the next pixel in line.  Note that we also enforce clockwise path behavior.
        If (i = 6) Then
            If (prevY = -1) Then nextX = -1 Else nextX = 1
            nextY = 0
        ElseIf (i = 9) Then
            nextX = 0
            If (prevX = 1) Then nextY = -1 Else nextY = 1
        Else
            nextX = m_XDirs(i)
            nextY = m_YDirs(i)
        End If
        
        'Add this point to our running polygon list, but *only* if this point is moving in a different
        ' direction from our previous point!
        If ((nextX <> prevX) Or (nextY <> prevY)) Then
            PushOntoStack x + xOffset, y + yOffset
            prevX = nextX
            prevY = nextY
        End If
        
        'Traverse to the next point in line
        x = x + nextX
        y = y + nextY
        
    'Loop checks are easy - continue until we return to our starting point
    Loop While ((x <> startX) Or (y <> startY))
    
    FindEdges = True
    
End Function

'Stack helper functions
Private Sub PushOntoStack(ByVal x As Long, ByVal y As Long)
    
    'Resize the stack as necessary
    If (m_StackPosition > m_StackHeight) Then
        m_StackHeight = m_StackHeight * 2 + 1
        ReDim Preserve m_Stack(0 To m_StackHeight) As POINTAPI
    End If
    
    With m_Stack(m_StackPosition)
        .x = x
        .y = y
    End With
    
    m_StackPosition = m_StackPosition + 1

End Sub

'After the path algorithm finishes, the caller needs to retrieve the final point list.  Because PD performs
' all rendering in floating-point coordinates, this function exists to make transfers easier.
Friend Sub RetrieveFinalPolygon(ByRef dstPoints() As POINTFLOAT, ByRef numOfPoints As Long)
    
    numOfPoints = m_StackPosition
    
    ReDim dstPoints(0 To numOfPoints - 1) As POINTFLOAT
    
    Dim i As Long
    For i = 0 To numOfPoints - 1
        dstPoints(i).x = m_Stack(i).x
        dstPoints(i).y = m_Stack(i).y
    Next i

End Sub

Private Sub Class_Initialize()

    'Reset all stack values
    m_StackPosition = 0
    m_StackHeight = INITIAL_STACK_HEIGHT - 1
    ReDim m_Stack(0 To m_StackHeight) As POINTAPI
    
    'Populate hard-coded direction descriptors
    ReDim m_XDirs(0 To 15) As Long: ReDim m_YDirs(0 To 15) As Long
    m_XDirs(0) = 1
    m_XDirs(1) = 0
    m_XDirs(2) = 1
    m_XDirs(3) = 1
    m_XDirs(4) = -1
    m_XDirs(5) = 0
    m_XDirs(6) = -1
    m_XDirs(7) = 1
    m_XDirs(8) = 0
    m_XDirs(9) = 0
    m_XDirs(10) = 0
    m_XDirs(11) = 0
    m_XDirs(12) = -1
    m_XDirs(13) = 0
    m_XDirs(14) = -1
    m_XDirs(15) = 1000000000
    
    m_YDirs(0) = 0
    m_YDirs(1) = -1
    m_YDirs(2) = 0
    m_YDirs(3) = 0
    m_YDirs(4) = 0
    m_YDirs(5) = -1
    m_YDirs(6) = 0
    m_YDirs(7) = 0
    m_YDirs(8) = 1
    m_YDirs(9) = -1
    m_YDirs(10) = 1
    m_YDirs(11) = 1
    m_YDirs(12) = 0
    m_YDirs(13) = -1
    m_YDirs(14) = 0
    m_YDirs(15) = 1000000000
    
End Sub
