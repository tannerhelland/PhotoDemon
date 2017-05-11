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
'Last updated: 09/May/17
'Last update: add comprehensive edge detection (e.g. edges of *multiple* objects in one image)
'
'This class supplies the "marching squares" algorithm used by a number of different tools in PD.
'
'To simplify the class (and improve performance), all operations are performed on byte arrays filled
' with interior/exterior data.  This class doesn't care how you generate those arrays, and it treats
' zeroes as exterior markers, and non-zeroes as interior ones.  (For a 32-bpp image, for example,
' you could generate the required array by setting transparent pixel coordinates to 0 and all other
' pixels to non-zero values.)
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

'Fast memory ops
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal dstPointer As Long, ByVal srcPointer As Long, ByVal numOfBytes As Long)

'Faster WAPI interop
Private Declare Function GdipIsVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal y As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisibleRegionPointI Lib "gdiplus" (ByVal hRegion As Long, ByVal x As Long, ByVal y As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long

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

'Want to retrieve the finished point list into a path object instead?  Use this convenient function.  (Note that this function
' does *not* erase the destination path, if one exists; instead, it adds the edge as a new, self-contained polygon.)
Friend Sub RetrieveFinalPolygonIntoPath(ByRef dstPath As pd2DPath, Optional ByVal closeShape As Boolean = True)
    
    If (dstPath Is Nothing) Then Set dstPath = New pd2DPath
    
    Dim numOfPoints As Long
    numOfPoints = m_StackPosition
    
    ReDim dstPoints(0 To numOfPoints - 1) As POINTFLOAT
    
    Dim i As Long
    For i = 0 To numOfPoints - 1
        dstPoints(i).x = m_Stack(i).x
        dstPoints(i).y = m_Stack(i).y
    Next i
    
    dstPath.AddPolygon numOfPoints, VarPtr(dstPoints(0)), closeShape
    
End Sub

'NEW TEST: remove polygon interiors via GDI+.  Performance is TBD.
'
'Want to find every last edge in an image, including non-standard ones or ones that may not close properly?  Use this function instead.
' Note that (by design) it does *not* forcibly close paths.  This allows it to return line segments or other non-standard edges that
' would otherwise be extremely difficult to define as polygons.
'
'The approach required by this kind of algorithm makes it unsuitable for subsequent polygon operations, but for raster-style operations
' (or just rendering overlays on affected edges), it is ideal and much much faster than a mathematically rigorous solution.
Friend Function FindAllEdges(ByRef dstPath As pd2DPath, ByRef srcArray() As Byte, ByVal startX As Long, ByVal startY As Long, ByVal maxX As Long, ByVal maxY As Long, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As Boolean
    
    If (dstPath Is Nothing) Then Set dstPath = New pd2DPath
    
    'Current (x, y) positions.  (These start at startX and startY.)  After a shape is completed, these will advance to the next
    ' non-zero position in the source array, and repeat a scan from there.
    Dim x As Long, y As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = startX
    initY = startY
    finalX = maxX
    finalY = maxY
    
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
    
    Dim dstPoints() As POINTFLOAT
    ReDim dstPoints(0) As POINTFLOAT
    
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    
    Dim tmpRegion As pd2DRegion
    Set tmpRegion = New pd2DRegion
    
    'Index into our precalculated direction arrays.
    Dim i As Long, j As Long
    Dim intBounds As RECTL
    Dim intInitX As Long, intInitY As Long, intFinalX As Long, intFinalY As Long
    Dim tmpGdipHandle As Long, tmpResult As Long
    
    'Starting at (x, y), look for the next valid edge pixel.  Once one is found, start tracing it.
    For y = initY To finalY
    For x = initX To finalX
        
        'Is this an edge pixel?
        If (srcArray(x, y) <> 0) Then
            
            Debug.Print "Starting contour search at (" & x & ", " & y & ")"
            
            'It is!  Start tracing its contour.
            startX = x
            startY = y
        
            'Start searching for contour points!
            Do
                
                'For the current pixel, determine an index into the marching squares lookup table.  Note that this
                ' table is constructed very specifically, which is what makes the following lookup technique work.
                ' (See: https://en.wikipedia.org/wiki/Marching_squares#Basic_algorithm)
                ' Note also that other languages could probably compute this more efficiently using bit-shifting...
                i = 0
                If (srcArray(x - 1, y - 1) <> 0) Then i = i Or 1
                If (srcArray(x, y - 1) <> 0) Then i = i Or 2
                If (srcArray(x - 1, y) <> 0) Then i = i Or 4
                If (srcArray(x, y) <> 0) Then i = i Or 8
                
                'Now that we have an index into the table, use that index to compute the actual physical
                ' location of the next pixel in line.  Note that we also enforce clockwise path behavior.
                ' (The special checks for 6 and 9 deal with ambiguous saddle points; by tracking previous
                '  direction, we can resolve the ambiguity.)
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
            
            'The stack now describes a completed polygon.  Add it to a temporary path object.
            tmpPath.ResetPath
            tmpPath.AddPolygonInt m_StackPosition, VarPtr(m_Stack(0)), True
            
            'Add the temporary polgyon to our running collection
            dstPath.AddPath tmpPath
            
            Debug.Print "Searching path interior..."
            tmpRegion.AddPath tmpPath, P2_CM_Replace
            tmpGdipHandle = tmpRegion.GetHandle()
            
            'We now want to search the interior of the path region, and invert any/all points lying *inside* the polygon.
            ' (This allows us to isolate holes in the region.)
            intBounds = tmpRegion.GetRegionBoundsL()
            intInitX = intBounds.Left
            intInitY = intBounds.Top
            intFinalX = intBounds.Left + intBounds.Right
            intFinalY = intBounds.Top + intBounds.Bottom
            
            For j = intInitY To intFinalY
            For i = intInitX To intFinalX
                GdipIsVisibleRegionPointI tmpGdipHandle, i, j, 0&, tmpResult
                If (tmpResult <> 0) Then
                    srcArray(i - xOffset, j - yOffset) = Not srcArray(i - xOffset, j - yOffset)
                End If
            Next i
            Next j
            
            'Reset the stack pointer
            m_StackPosition = 0
            
        End If
    
    'Continue with the next segment
    Next x
    Next y
    
    FindAllEdges = True
    
End Function

'It's generally trivial to create threshold maps of an image (where pixels below a threshold are marked "0", while pixels above a threshold
' are marked "non-zero"), but it's a bit more obnoxious to convert this to an "edge" representation, where interior areas have been removed.
' This helper function will complete such a task for you, but note that you must supply *both* a source and destination array.
' (The destination array will be enlarged by 1 pixel on each side, to ensure no boundary issues.)
'
'IMPORTANT NOTE!  The source array *will be erased* by this function, by design.  If you want it to be retained, you *must* make a
' manual copy prior to invoking this function.
Friend Sub ConvertThresholdMapToEdgeMap(ByRef srcArray() As Byte, ByRef dstArray() As Byte, ByVal srcXBound As Long, ByVal srcYBound As Long)

    'First thing we want to do is initialize the destination array.  Note that this array has a guaranteed blank 1-px border around
    ' all sides; this is crucial for simplifying our edge-detection techniques.
    Dim xBound As Long, yBound As Long
    xBound = srcXBound + 2
    yBound = srcYBound + 2
    
    ReDim dstArray(0 To xBound, 0 To yBound) As Byte
    
    'Next, we need to copy all lines from the boundary check array to our outline array, offsetting them by (1) in each direction.
    ' This guarantees a boundary of zeroes around the target image, which greatly accelerates edge detection.
    Dim x As Long, y As Long
    For y = 0 To srcYBound
        CopyMemory_Strict VarPtr(dstArray(0, y + 1)) + 1, VarPtr(srcArray(0, y)), srcXBound
    Next y
    
    'The destination array now contains an exact copy of the source array, but with a 1-px border added.  Wipe the source array.
    FillMemory VarPtr(srcArray(0, 0)), srcXBound * srcYBound, 0
    
    Dim pxIsEdge As Boolean
    
    'Next, we need to scan the destination array, and mark any outline pixels (*not* interior pixels) back onto the source map,
    ' accounting for that 1-px border offset as we do so.
    For y = 1 To srcYBound + 1
    For x = 1 To srcXBound + 1
        
        pxIsEdge = False
        
        'We only care about pixels that are *not* thresholded.
        If (dstArray(x, y) = 0) Then
            
            If (dstArray(x - 1, y - 1) <> 0) Then
                pxIsEdge = True
            ElseIf (dstArray(x, y - 1) <> 0) Then
                pxIsEdge = True
            ElseIf (dstArray(x + 1, y - 1) <> 0) Then
                pxIsEdge = True
            ElseIf (dstArray(x - 1, y) <> 0) Then
                pxIsEdge = True
            ElseIf (dstArray(x + 1, y) <> 0) Then
                pxIsEdge = True
            ElseIf (dstArray(x - 1, y + 1) <> 0) Then
                pxIsEdge = True
            ElseIf (dstArray(x, y + 1) <> 0) Then
                pxIsEdge = True
            ElseIf (dstArray(x + 1, y + 1) <> 0) Then
                pxIsEdge = True
            End If
            
            'If this is an edge pixel, mark it in the source array
            If pxIsEdge Then srcArray(x - 1, y - 1) = 1
        
        End If
        
    Next x
    Next y
    
    'The source array has now been re-processed to reflect *only* edge pixels.  Re-copy it into the destination, overwriting our
    ' original copy as we go.
    For y = 0 To srcYBound
        CopyMemory_Strict VarPtr(dstArray(0, y + 1)) + 1, VarPtr(srcArray(0, y)), srcXBound
    Next y
    
End Sub

'To improve edge-detection performance, a source array needs to have a guaranteed blank set of pixels around its outside edge.
' Use this helper function to make that happen.
Friend Sub MakeArrayEdgeSafe(ByRef srcArray() As Byte, ByRef dstArray() As Byte, ByVal srcXBound As Long, ByVal srcYBound As Long)

    'First thing we want to do is initialize the destination array.  Note that this array has a guaranteed blank 1-px border around
    ' all sides; this is crucial for simplifying our edge-detection techniques.
    Dim xBound As Long, yBound As Long
    xBound = srcXBound + 2
    yBound = srcYBound + 2
    
    ReDim dstArray(0 To xBound, 0 To yBound) As Byte
    
    'Next, we need to copy all lines from the boundary check array to our outline array, offsetting them by (1) in each direction.
    ' This guarantees a boundary of zeroes around the target image, which greatly accelerates edge detection.
    Dim x As Long, y As Long
    For y = 0 To srcYBound
        CopyMemory_Strict VarPtr(dstArray(0, y + 1)) + 1, VarPtr(srcArray(0, y)), srcXBound
    Next y
    
    'The destination array now contains an exact copy of the source array, but with a 1-px border added.  Wipe the source array.
    
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
