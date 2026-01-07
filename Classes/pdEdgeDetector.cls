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
'Copyright 2017-2026 by Tanner Helland
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'A stack is used to track polygon points enclosing the first discovered region
Private m_Stack() As PointAPI
Private m_StackPosition As Long
Private m_StackHeight As Long
Private Const INITIAL_STACK_HEIGHT As Long = 4096

'Marching squares uses predefined direction descriptors.  See https://en.wikipedia.org/wiki/Marching_squares
' These descriptors are populated in Class_Initialize().
Private m_XDirs() As Long, m_YDirs() As Long

Private Declare Function GdipGetRegionScansCount Lib "gdiplus" (ByVal hRegion As Long, ByRef dstRectCount As Long, ByVal optTransformMatrix As Long) As GP_Result
Private Declare Function GdipGetRegionScansI Lib "gdiplus" (ByVal hRegion As Long, ByVal ptrToRectLArray As Long, ByRef dstNumOfRects As Long, ByVal optTransformMatrix As Long) As GP_Result
Private m_RectLCache() As RectL, m_SizeOfRectCache As Long

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
    
    FindStartingPoint = (Not keepSearching)
    
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
        ReDim Preserve m_Stack(0 To m_StackHeight) As PointAPI
    End If
    
    With m_Stack(m_StackPosition)
        .x = x
        .y = y
    End With
    
    m_StackPosition = m_StackPosition + 1

End Sub

'After the path algorithm finishes, the caller needs to retrieve the final point list.  Because PD performs
' all rendering in floating-point coordinates, this function exists to make transfers easier.
Friend Sub RetrieveFinalPolygon(ByRef dstPoints() As PointFloat, ByRef numOfPoints As Long)
    
    numOfPoints = m_StackPosition
    
    ReDim dstPoints(0 To numOfPoints - 1) As PointFloat
    
    Dim i As Long
    For i = 0 To numOfPoints - 1
        dstPoints(i).x = m_Stack(i).x
        dstPoints(i).y = m_Stack(i).y
    Next i

End Sub

'Want to find every last edge in an image, including non-standard ones or ones that may not close properly?
' Then use this function.
'
'Complex polygons (including concave shapes, shapes with holes, or entirely separate polygons) are all handled
' correctly and returned as a single unified pd2DPath object.
Friend Function FindAllEdges(ByRef dstPath As pd2DPath, ByRef srcArray() As Byte, ByVal startX As Long, ByVal startY As Long, ByVal maxX As Long, ByVal maxY As Long, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As Boolean
    
    If (dstPath Is Nothing) Then Set dstPath = New pd2DPath Else dstPath.ResetPath
    dstPath.SetFillRule P2_FR_OddEven
    m_StackPosition = 0
            
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
    
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    
    Dim tmpRegion As pd2DRegion
    Set tmpRegion = New pd2DRegion
    
    'In debug mode, we want to track how much time we spend tracing outlines and removing filled shapes.  This gives us a
    ' better idea of where we might improve performance.
    Dim startTime As Currency, traceTime As Currency, fillTime As Currency
    Dim gdipCreateTime As Currency, gdipCopyTime As Currency, gdipRegionTime As Currency
    
    'Index into our precalculated direction arrays.
    Dim i As Long, j As Long, n As Long
    Dim intInitX As Long, intInitY As Long, intFinalX As Long, intFinalY As Long
    Dim numOfRects As Long
    Dim tmpGdipHandle As Long
    
    'To improve performance, we're going to point a transient 1D array at certain rows during processing.
    ' This requires unsafe array manipulation, but it can be meaningfully faster than 2D array accesses.
    Dim tmpArray() As Byte, tmpSA As SafeArray1D
    
    'Populate the safearray struct's unchanging values
    Dim srcArrayBasePointer As Long, srcArrayStride As Long
    srcArrayBasePointer = VarPtr(srcArray(0, 0))
    srcArrayStride = UBound(srcArray(), 1) + 1
    
    With tmpSA
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcArrayStride
        
        'pvData *will* change as the function goes along, but let's at least start with a safe value
        .pvData = srcArrayBasePointer
        
    End With
    
    'Point the uninitialized temporary array at our custom-built SafeArray struct
    PutMem4 VarPtrArray(tmpArray()), VarPtr(tmpSA)
    
    'And now, a duplicate of the above steps, but designed for the outer loop
    Dim tmpPixelArray() As Byte, tmpPixelSA As SafeArray1D
    Dim srcPixelPointer As Long, srcPixelStride As Long
    srcPixelPointer = VarPtr(srcArray(0, 0))
    srcPixelStride = UBound(srcArray(), 1) + 1
    
    With tmpPixelSA
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcPixelStride
        
        'pvData *will* change as the function goes along, but let's at least start with a safe value
        .pvData = srcPixelPointer
        
    End With
    
    'Point the uninitialized temporary array at our custom-built SafeArray struct
    PutMem4 VarPtrArray(tmpPixelArray()), VarPtr(tmpPixelSA)
    
    'Because we've typically added a 1-px blank byte border around our image data (during pre-processing), all edge coordinates
    ' are going to be offset by (1, 1).  Create a generic transform now, which we'll use to offset edge coordinates before
    ' reporting them back to our parent object.
    Dim tmpTransform As pd2DTransform
    Set tmpTransform = New pd2DTransform
    tmpTransform.ApplyTranslation -xOffset, -yOffset
    
    Dim tmpMatrix As Long
    tmpMatrix = tmpTransform.GetHandle
    
    'Starting at (x, y), look for the next valid edge pixel.  Once one is found, start tracing it.
    For y = initY To finalY
        tmpPixelSA.pvData = srcPixelPointer + (y * srcPixelStride)
    For x = initX To finalX
        
        'Is this an edge pixel?
        If (tmpPixelArray(x) <> 0) Then
            
            'It is!  Start tracing its contour.
            startX = x
            startY = y
            VBHacks.GetHighResTime startTime
            
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
            
            traceTime = traceTime + VBHacks.GetTimerDifferenceNow(startTime)
            VBHacks.GetHighResTime startTime
            
            'The stack now describes a completed polygon.  Add it to a temporary path object.
            tmpPath.ResetPath
            tmpPath.SetFillRule P2_FR_OddEven
            tmpPath.AddPolygonInt m_StackPosition, VarPtr(m_Stack(0)), True
            
            gdipCreateTime = gdipCreateTime + VBHacks.GetTimerDifferenceNow(startTime)
            VBHacks.GetHighResTime startTime
            
            'Add the temporary polgyon to our running collection.  (NOTE!  You might be tempted to use dstPath.AddPath here,
            ' passing the path object that we've already assembled.  Do *not* do this.  For some reason, .AddPath is insanely
            ' slow compared to just adding the raw data points.  (Like, 1000x slower for complicated paths.)
            dstPath.AddPolygonInt m_StackPosition, VarPtr(m_Stack(0)), True
            
            gdipCopyTime = gdipCopyTime + VBHacks.GetTimerDifferenceNow(startTime)
            VBHacks.GetHighResTime startTime
            
            'Create a region object from our outline.  Why do this when we already have a perfectly good path?  Because regions
            ' are *significantly* faster to hit-test than paths.
            
            ' (ALSO NOTE: I have tested both GDI and GDI+ regions here.  For reasons I don't understand, GDI+ regions are much
            ' faster to hit-test than GDI regions.  (At least several times faster, in fact, with a gap that widens as the region
            ' gets more complex.)  I would have expected the opposite, but data doesn't lie - so we stick with GDI+ regions instead.
            tmpRegion.AddPath tmpPath, P2_CM_Replace
            tmpGdipHandle = tmpRegion.GetHandle()
            
            'We now want to search the interior of the path region, and invert any/all points lying *inside* the polygon.
            ' (This allows us to isolate holes inside the path we just traced.)
            
            'A naive approach to this might look something like this:
            
'            intBounds = tmpRegion.GetRegionBoundsL()
'            intInitX = intBounds.Left
'            intInitY = intBounds.Top
'            intFinalX = intBounds.Left + intBounds.Right
'            intFinalY = intBounds.Top + intBounds.Bottom
'
'            For j = intInitY To intFinalY
'            For i = intInitX To intFinalX
'                GdipIsVisibleRegionPointI tmpGdipHandle, i, j, 0&, tmpResult
'                If (tmpResult <> 0) Then
'                    srcArray(i - xOffset, j - yOffset) = Not srcArray(i - xOffset, j - yOffset)
'                End If
'            Next i
'            Next j
            
            'Unfortunately, hit-testing individual points against a complex region is pretty damn slow, so while this technique
            ' does work - and the code is nice and simple - we want to use something faster.
            
            'Start by seeing how many basic rectangles are required to accurately represent the entirety of this region.
            GdipGetRegionScansCount tmpGdipHandle, numOfRects, tmpMatrix
            If (numOfRects > 0) Then
                
                'Next, ask GDI+ to give us a copy of the current region, constructed as a (potentially large) series of
                ' nothing but basic rectangles.
                If (numOfRects > m_SizeOfRectCache) Then
                    m_SizeOfRectCache = numOfRects
                    ReDim m_RectLCache(0 To m_SizeOfRectCache - 1) As RectL
                End If
            
                If (GdipGetRegionScansI(tmpGdipHandle, VarPtr(m_RectLCache(0)), numOfRects, tmpMatrix) = GP_OK) Then
                    
                    gdipRegionTime = gdipRegionTime + VBHacks.GetTimerDifferenceNow(startTime)
                    VBHacks.GetHighResTime startTime
                    
                    'We now have a full copy of this region, in the form of individual rectangles (the sum total of which
                    ' represent the underlying region).
                    
                    'For each rect in this collection, we now need to manually invert all bytes.  This turns holes inside
                    ' the region into a new sub-region, which we can then detect on subsequent passes of our trace algorithm.
                    For n = 0 To numOfRects - 1
                        
                        With m_RectLCache(n)
                            intInitX = .Left
                            intInitY = .Top
                            intFinalX = .Left + .Right - 1
                            intFinalY = .Top + .Bottom - 1
                        End With
                        
                        For j = intInitY To intFinalY
                            tmpSA.pvData = srcArrayBasePointer + (j * srcArrayStride)
                        For i = intInitX To intFinalX
                            tmpArray(i) = Not tmpArray(i)
                        Next i
                        Next j
                        
                    Next n
                
                End If

            End If
            
            fillTime = fillTime + VBHacks.GetTimerDifferenceNow(startTime)
            
            'Reset the stack pointer
            m_StackPosition = 0
            
        End If
    
    'Continue with the next segment
    Next x
    Next y
    
    'Before exiting, free any temporary source array pointers
    PutMem4 VarPtrArray(tmpArray()), 0&
    PutMem4 VarPtrArray(tmpPixelArray()), 0&
    
    'Want detailed timing reports?  Uncomment this line:
    'pdDebug.LogAction "Detailed pdEdgeDetector report || Trace: " & Format$(traceTime * 1000, "#0") & ", GdiP/Create: " & Format$(gdipCreateTime * 1000, "#0") & ", GdiP/Copy: " & Format$(gdipCopyTime * 1000, "#0") & ", GdiP/Region: " & Format$(gdipRegionTime * 1000, "#0") & ", Fill: " & Format$(fillTime * 1000, "#0")
    
    FindAllEdges = True
    
End Function

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
    Dim y As Long
    For y = 0 To srcYBound
        CopyMemoryStrict VarPtr(dstArray(1, y + 1)), VarPtr(srcArray(0, y)), srcXBound + 1
    Next y
    
    'The destination array now contains an exact copy of the source array, but with a 1-px border added.  Wipe the source array.
    
End Sub

Private Sub Class_Initialize()

    'Reset all stack values
    m_StackPosition = 0
    m_StackHeight = INITIAL_STACK_HEIGHT - 1
    ReDim m_Stack(0 To m_StackHeight) As PointAPI
    
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
