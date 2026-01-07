VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdKDTreeArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon array-backed KD-Tree (slower than normal KD-tree, but much faster tear-down due to no child classes)
'Copyright 2018-2026 by Tanner Helland
'Created: 28/January/18
'Last updated: 23/June/22
'Last update: add support for tracking/returning original palette index
'
'This class contains a specialized K-D tree implementation.  Please see the normal pdKDTree class for a generic
' KD-tree implementation, including full implementation details.  This class exists purely to work around some
' painful issues with COM teardown performance in VB6 when large numbers of classes are created.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Constructing the tree now uses a stack instead of recursion; this improves performance when node
' count is astronomically high
Private Type SortStack
    sLB As Long
    sUB As Long
End Type

'Each tree element is stored in a backing array.  This backing array removes the need for child classes,
' which changes the performance trade-offs involved - this class is much faster to build and teardown than
' the normal KD-tree class, but it requires more complicated code because we can't just "walk" through
' natural parent-child class relationships.  Because the teardown penalty of a standard implementation is
' so severe, however, (we're talking a teardown time in *minutes* if the class count exceeds several hundred
' thousand instances), this class is absolutely worth the extra code when I need a KD-tree with tens or
' hundreds of thousands of nodes.
Private Type KDNode
    nColor As RGBQuad
    nChildLeft As Long
    nChildRight As Long
    nDepth As Integer           'VB will pad the struct to 4-byte alignment regardless, so may as well make this a Long
    nOriginalIndex As Integer   'Original index mode is only supported up to 32k indices; this is by design given PD's usage
End Type

Private m_numNodes As Long
Private m_Nodes() As KDNode

'When retrieving values from the tree, we search through it recursively.  To minimize the stack size
' hit caused by recursion, we declare as many recursive function variables at class-level as we can.
Private m_bestDistance As Long, m_testDistance As Long
Private m_testColor As RGBQuad, m_srcColor As RGBQuad, m_bestColor As RGBQuad
Private m_Depth As Long

'This class only matches RGB values at present; RGBA could easily be added, but I do not have a use for it
' at present (the regular pdKDTree implementation handles this instead)
Private m_curR As Long, m_curG As Long, m_curB As Long
Private m_testR As Long, m_testG As Long, m_testB As Long

Private m_numComparisons As Long

'Build a generic KD-tree as fast as possible.  This will insert colors in palette order, meaning the tree is likely to
' be grossly unbalanced.  This builder is meant only for scenarios where the source palette is enormous (10,000+ colors)
' and subsequent match requirements are small, as in PD's color LUT constructor.  For normal palette-matching behavior,
' with palette sizes <= 256, use the balanced tree builder below.
Friend Function BuildTree(ByRef srcPalette() As RGBQuad, ByVal lowBound As Long, ByVal highBound As Long, Optional ByVal useAlphaToo As Boolean = False) As Boolean

    BuildTree = (highBound >= lowBound)
    
    If BuildTree Then
        
        'Initialize the backing array as precisely as possible.  (This spares us from needing additional
        ' array bound checks on the insert loop.)
        m_numNodes = 0
        ReDim m_Nodes(0 To (highBound - lowBound)) As KDNode
        
        'Manually populate the first node
        With m_Nodes(0)
            .nColor = srcPalette(0)
            .nDepth = 0
            .nOriginalIndex = 0
        End With
        
        m_numNodes = 1
        
        'Build an addition stack
        Const INIT_STACK_SIZE As Long = 256
        Dim qsRemaining() As SortStack, qsStackPtr As Long
        ReDim qsRemaining(0 To INIT_STACK_SIZE - 1) As SortStack
        qsStackPtr = 0
        qsRemaining(0).sLB = lowBound + 1
        qsRemaining(0).sUB = highBound
        
        Do
            
            lowBound = qsRemaining(qsStackPtr).sLB
            highBound = qsRemaining(qsStackPtr).sUB
            qsStackPtr = qsStackPtr - 1
            
            'Add the requested node (the center of the current list)
            Dim curIndex As Long
            curIndex = lowBound + (highBound - lowBound) \ 2
            If useAlphaToo Then
                InsertNodeRGBA srcPalette(curIndex), curIndex
            Else
                InsertNodeRGB srcPalette(curIndex), curIndex
            End If
            
            'Partition the high and low lists, add to the stack, then repeat
            If (lowBound <= curIndex - 1) Then
                qsStackPtr = qsStackPtr + 1
                If (qsStackPtr > UBound(qsRemaining)) Then ReDim Preserve qsRemaining(0 To qsStackPtr * 2 - 1) As SortStack
                qsRemaining(qsStackPtr).sLB = lowBound
                qsRemaining(qsStackPtr).sUB = curIndex - 1
            End If
            
            If (curIndex + 1 <= highBound) Then
                qsStackPtr = qsStackPtr + 1
                If (qsStackPtr > UBound(qsRemaining)) Then ReDim Preserve qsRemaining(0 To qsStackPtr * 2 - 1) As SortStack
                qsRemaining(qsStackPtr).sLB = curIndex + 1
                qsRemaining(qsStackPtr).sUB = highBound
            End If
            
        Loop While (qsStackPtr >= 0)
        
    End If
        
End Function

Friend Function BuildTreeBalanced(ByRef srcPalette() As RGBQuad, ByVal lowBound As Long, ByVal highBound As Long, Optional ByVal useAlphaToo As Boolean = False) As Boolean

    BuildTreeBalanced = (highBound >= lowBound)
    If BuildTreeBalanced Then
        
        'Initialize the backing array as precisely as possible.  (This spares us from needing additional
        ' array bound checks on the insert loop.)
        m_numNodes = 0
        ReDim m_Nodes(0 To (highBound - lowBound)) As KDNode
        
        'Transfer the source palette into a "cache" type, so we can store original palette indices
        ' into the KD-tree (instead of just bare RGBQuads)
        Dim tmpCache() As PDPaletteCache
        ReDim tmpCache(0 To highBound - lowBound) As PDPaletteCache
        
        Dim i As Long
        For i = lowBound To highBound
            tmpCache(i - lowBound).ColorValue = srcPalette(i)
            tmpCache(i - lowBound).OrigIndex = i + lowBound
        Next i
        
        'For best performance, users should request a balanced tree.  This imposes a tree creation penalty,
        ' but subsequent nearest-neighbor queries are *much* faster.
        InsertNodeBalanced tmpCache, 0, useAlphaToo
        
    End If
    
End Function

'Given some source color, return the best color match (RGB) from the tree.
' RGBA matching is not currently implemented, by design, but easily could be in the future.
Friend Function GetNearestColor(ByRef srcColor As RGBQuad) As RGBQuad
    
    'Failsafe only; build the damn tree before trying to retrieve colors from it!
    If (m_numNodes > 0) Then
        
        m_bestDistance = LONG_MAX
        m_srcColor = srcColor
        
        'Start searching at the root node
        CompareNode 0
        
        'Return the best-match color from the recursive descent through the tree
        GetNearestColor = m_bestColor
        
    End If
        
End Function

Private Sub CompareNode(ByVal idxTest As Long)
    
    m_numComparisons = m_numComparisons + 1
    
    'Before checking child nodes, compare the target color against this node's color.
    m_testColor = m_Nodes(idxTest).nColor
    m_curR = m_testColor.Red
    m_curG = m_testColor.Green
    m_curB = m_testColor.Blue
    
    m_testR = m_curR - m_srcColor.Red
    m_testG = m_curG - m_srcColor.Green
    m_testB = m_curB - m_srcColor.Blue
    m_testDistance = m_testR * m_testR + m_testG * m_testG + m_testB * m_testB
    
    'Store the best (closest) result so far
    If (m_testDistance < m_bestDistance) Then
        m_bestDistance = m_testDistance
        m_bestColor = m_testColor
    End If
    
    'Next, we want to determine if any of this node's child nodes contain potential points "closer to"
    ' or "further from" our current best-match color.  Note that we need to cache these values in
    ' local variables because we're going to potentially initiate a new layer of recursion here.
    '
    'Also, we use our previously stored "depth" tracker to determine which axis to use for comparisons.
    Dim srcComponent As Long, targetComponent As Long
    Select Case m_Nodes(idxTest).nDepth
        Case 0
            srcComponent = m_srcColor.Red
            targetComponent = m_curR
        Case 1
            srcComponent = m_srcColor.Green
            targetComponent = m_curG
        Case 2
            srcComponent = m_srcColor.Blue
            targetComponent = m_curB
    End Select
    
    'If the target color is *less than* this node's color, better matches will be found in
    ' the left tree.  (Conversely, if it is *greater than or equal to* this node's color,
    ' search the right tree first.)
    If (srcComponent < targetComponent) Then
    
        'Search the left tree *immediately*
        If (m_Nodes(idxTest).nChildLeft > 0) Then CompareNode m_Nodes(idxTest).nChildLeft
        
        'Now we need to repeat some ugly steps (but writing it this way minimizes branching for
        ' improved perf).  We next need to see if it's necessary to check the right branch
        ' of ths tree as well.  We do this by testing the theoretical "closest point" possible in
        ' the right branch, and if that "theoretical" point is closer to the target color than our
        ' current best match, we need to search the right branch for possible targets as well.
        If (m_Nodes(idxTest).nChildRight > 0) Then
            
            'We know that the best value for this tree *likely* lies in the left branch.  In order for
            ' a value in the right branch to be closer than the current value, it would need to be the
            ' *smallest possible value* in that tree - so at depth 0, this means a color with an r value
            ' as low as possible.
            '
            'Because our KD-tree implementation uses "greater-than-or-equal-to" for right branch
            ' determination, the lowest possible value in right branches is an r-value equal to the
            ' target color's.  (We ignore green and blue because they could potentially be *equal* to
            ' the target color, but we have no way of knowing that as this node only branches on red!)
            m_testDistance = srcComponent - targetComponent
            
            'If the closest "theoretical" point in the right branch is closer than the current best match,
            ' we must also search that sub-tree for a best match.
            If ((m_testDistance * m_testDistance) < m_bestDistance) Then CompareNode m_Nodes(idxTest).nChildRight
            
        End If
        
    Else
        
        'Search the right tree
        If (m_Nodes(idxTest).nChildRight > 0) Then CompareNode m_Nodes(idxTest).nChildRight
        If (m_Nodes(idxTest).nChildLeft > 0) Then
            
            'Because we're querying the left tree, the nearest possible color would have to be at least
            ' one less than this node's color.  As such, if this node has a value of 0, there is no
            ' possible way that the left node could contain a closer color (as it can't contain
            ' *any* colors less than zero!)
            If (srcComponent > 0) Then
                m_testDistance = srcComponent - 1 - targetComponent
                If ((m_testDistance * m_testDistance) < m_bestDistance) Then CompareNode m_Nodes(idxTest).nChildLeft
            End If
            
        End If
        
    End If
    
End Sub

Private Sub InsertNodeRGB(ByRef srcColor As RGBQuad, ByVal originalIndex As Long)
    
    'We always start on the R-axis (note that the root node must already be filled for this function to work)
    m_Depth = 0
    
    Dim curNode As Long
    curNode = 0
    
    'We'll iterate until we find an empty node; loop termination is manual via Exit Do on appropriate paths
    Do
        
        'Depth comes from the current child node
        m_Depth = m_Nodes(curNode).nDepth
        
        'Compare the target color against this node's color.  If the passed color is "higher" or "lower" than
        ' the target color (which component we compare depends on depth) we'll traverse down the respective
        ' branch of the tree.
        m_testColor = m_Nodes(curNode).nColor
        
        'Do a quick search for equality; if found, return our color and exit.
        ' (Note the manual short-circuiting to work around VB deficiencies.)
        If (srcColor.Red = m_testColor.Red) Then
            If (srcColor.Green = m_testColor.Green) Then
                If (srcColor.Blue = m_testColor.Blue) Then
                    Exit Do
                End If
            End If
        End If
        
        'Rotate between RGB axes
        Dim useLeftBranch As Boolean
        
        Select Case m_Depth
            Case 0
                useLeftBranch = (srcColor.Red < m_testColor.Red)
            Case 1
                useLeftBranch = (srcColor.Green < m_testColor.Green)
            Case 2
                useLeftBranch = (srcColor.Blue < m_testColor.Blue)
        End Select
        
        'Increment depth before continuing with traversal
        m_Depth = m_Depth + 1
        If (m_Depth > 2) Then m_Depth = 0
        
        'We will now assign the color to empty child indices, or continue traversing the tree as necessary
        
        'Left branch
        If useLeftBranch Then
            
            'No child
            If (m_Nodes(curNode).nChildLeft = 0) Then
                
                'Add this color as a new child, then exit
                m_Nodes(curNode).nChildLeft = m_numNodes
                With m_Nodes(m_numNodes)
                    .nColor = srcColor
                    .nOriginalIndex = originalIndex And &H7FFF&
                    .nDepth = m_Depth
                End With
                m_numNodes = m_numNodes + 1
                
                Exit Do
                
            'Has a child
            Else
                curNode = m_Nodes(curNode).nChildLeft
            End If
            
        'Right branch
        Else
            
            'No child
            If (m_Nodes(curNode).nChildRight = 0) Then
                
                'Add this color as a new child, then exit
                m_Nodes(curNode).nChildRight = m_numNodes
                With m_Nodes(m_numNodes)
                    .nColor = srcColor
                    .nOriginalIndex = originalIndex And &H7FFF&
                    .nDepth = m_Depth
                End With
                m_numNodes = m_numNodes + 1
                
                Exit Do
                
            'Has a child
            Else
                curNode = m_Nodes(curNode).nChildRight
            End If
            
        End If
        
    Loop
    
End Sub

Private Sub InsertNodeRGBA(ByRef srcColor As RGBQuad, ByVal originalIndex As Long)
    
    'We always start on the R-axis (note that the root node must already be filled for this function to work)
    m_Depth = 0
    
    Dim curNode As Long
    curNode = 0
    
    'We'll iterate until we find an empty node; loop termination is manual via Exit Do on appropriate paths
    Do
        
        'Depth comes from the current child node
        m_Depth = m_Nodes(curNode).nDepth
        
        'Compare the target color against this node's color.  If the passed color is "higher" or "lower" than
        ' the target color (which component we compare depends on depth) we'll traverse down the respective
        ' branch of the tree.
        m_testColor = m_Nodes(curNode).nColor
        
        'Do a quick search for equality; if found, return our color and exit.
        ' (Note the manual short-circuiting to work around VB deficiencies.)
        If (srcColor.Red = m_testColor.Red) Then
            If (srcColor.Green = m_testColor.Green) Then
                If (srcColor.Blue = m_testColor.Blue) Then
                    If (srcColor.Alpha = m_testColor.Alpha) Then Exit Do
                End If
            End If
        End If
        
        'Rotate between RGB axes
        Dim useLeftBranch As Boolean
        
        Select Case m_Depth
            Case 0
                useLeftBranch = (srcColor.Red < m_testColor.Red)
            Case 1
                useLeftBranch = (srcColor.Green < m_testColor.Green)
            Case 2
                useLeftBranch = (srcColor.Blue < m_testColor.Blue)
            Case 3
                useLeftBranch = (srcColor.Alpha < m_testColor.Alpha)
        End Select
        
        'Increment depth before continuing with traversal
        m_Depth = m_Depth + 1
        If (m_Depth > 3) Then m_Depth = 0
        
        'We will now assign the color to empty child indices, or continue traversing the tree as necessary
        
        'Left branch
        If useLeftBranch Then
            
            'No child
            If (m_Nodes(curNode).nChildLeft = 0) Then
                
                'Add this color as a new child, then exit
                m_Nodes(curNode).nChildLeft = m_numNodes
                With m_Nodes(m_numNodes)
                    .nColor = srcColor
                    .nOriginalIndex = originalIndex And &H7FFF&
                    .nDepth = m_Depth
                End With
                m_numNodes = m_numNodes + 1
                
                Exit Do
                
            'Has a child
            Else
                curNode = m_Nodes(curNode).nChildLeft
            End If
            
        'Right branch
        Else
            
            'No child
            If (m_Nodes(curNode).nChildRight = 0) Then
                
                'Add this color as a new child, then exit
                m_Nodes(curNode).nChildRight = m_numNodes
                With m_Nodes(m_numNodes)
                    .nColor = srcColor
                    .nOriginalIndex = originalIndex And &H7FFF&
                    .nDepth = m_Depth
                End With
                m_numNodes = m_numNodes + 1
                
                Exit Do
                
            'Has a child
            Else
                curNode = m_Nodes(curNode).nChildRight
            End If
            
        End If
        
    Loop
    
End Sub

'Like most (all?) tree structures, KD trees achieve maximum performance when the tree is constructed
' as close to balanced as possible.  The only straightforward way to accomplish this is when all tree
' contents are known in advance - something that is fortunately true for PD's primary usage of
' palette-matching.
'
'So I've created this function, which is guaranteed to create a balanced tree.  As you'd expect, it is
' slower than a default Insert method would be.  However, subsequent queries on a tree created by
' *this* method will be much faster than a tree created by the default Insert method.  In PD, it is an
' excellent trade-off to use this tree creation method (which on a 256-color palette still takes < 1ms)
' in return for a massive improvement during palette matching (on a 256-color palette and ~10 mb image,
' the performance improvement is many *actual* seconds!)
'
'This function returns the index into the master KD-tree table of the assigned node, which the parent can
' use to correctly assign child node indices.
Friend Function InsertNodeBalanced(ByRef srcPalette() As PDPaletteCache, Optional ByVal depth As Long = 0, Optional ByVal useAlphaToo As Boolean = False) As Long
    
    'MOD is slower than a simple branch; cycle between values 0/1/2/3 for red/green/blue/alpha axes
    If useAlphaToo Then
        If (depth > 3) Then depth = 0
    Else
        If (depth > 2) Then depth = 0
    End If
    
    'In this insertion mode, we *always* start by assigning the current node the median value for the
    ' current depth (0, 1, 2 for red, green, blue).  After assigning this node its value, we will
    ' split the remaining palette entries into two groups (one lower, one greater-than-or-equal);
    ' these groups are then passed to left and right child nodes, who will proceed identically.
    
    'Normally, finding the median of a data set requires sorting, but sorting is expensive.
    ' RGB data is discrete and on a fixed range, so we can cheat and use histograms to find the
    ' median much more quickly.
    
    ' (Note that this step can be skipped entirely if the source palette only contains one color;
    ' this is likely for leaf nodes at the bottom of the tree.)
    If (UBound(srcPalette) > 0) Then
    
        Dim palHistogram(0 To 255) As Long
        
        'Build a histogram for the current palette, using the color definition appropriate for this depth.
        Dim i As Long
        For i = 0 To UBound(srcPalette)
            If (depth = 0) Then
                palHistogram(srcPalette(i).ColorValue.Red) = palHistogram(srcPalette(i).ColorValue.Red) + 1
            ElseIf (depth = 1) Then
                palHistogram(srcPalette(i).ColorValue.Green) = palHistogram(srcPalette(i).ColorValue.Green) + 1
            ElseIf (depth = 2) Then
                palHistogram(srcPalette(i).ColorValue.Blue) = palHistogram(srcPalette(i).ColorValue.Blue) + 1
            Else
                palHistogram(srcPalette(i).ColorValue.Alpha) = palHistogram(srcPalette(i).ColorValue.Alpha) + 1
            End If
        Next i
        
        'Now that the histogram is known, find the median for the current channel
        Dim pxCount As Long, numPixelsReq As Long
        numPixelsReq = Int((CDbl(UBound(srcPalette) + 1) + 0.5) / 2!)
        If (numPixelsReq < 1) Then numPixelsReq = 1
        
        For i = 0 To 255
            pxCount = pxCount + palHistogram(i)
            If (pxCount >= numPixelsReq) Then Exit For
        Next i
        
        'i now points at the median histogram index.  Find the first color in the table that matches
        ' this entry, and make it this node's color.
        Dim targetValue As Long
        targetValue = i
        
        For i = 0 To UBound(srcPalette)
            If (depth = 0) Then
                If (srcPalette(i).ColorValue.Red = targetValue) Then Exit For
            ElseIf (depth = 1) Then
                If (srcPalette(i).ColorValue.Green = targetValue) Then Exit For
            ElseIf (depth = 2) Then
                If (srcPalette(i).ColorValue.Blue = targetValue) Then Exit For
            Else
                If (srcPalette(i).ColorValue.Alpha = targetValue) Then Exit For
            End If
        Next i
        
        'i now points at the palette value we want for this node.  Assign it.
        Dim targetIndex As Long
        targetIndex = i
        
        With m_Nodes(m_numNodes)
            .nColor = srcPalette(targetIndex).ColorValue
            .nDepth = depth
            .nOriginalIndex = srcPalette(targetIndex).OrigIndex And &H7FFF&
        End With
        InsertNodeBalanced = m_numNodes
        m_numNodes = m_numNodes + 1
        
        'We now want to construct two new sub-palettes: one for colors *less than* the current entry
        ' (these go in the left node), and one for colors *greater than or equal to* the current entry
        ' (these go in the right node).  It is possible - and likely, as we move down the tree - that one
        ' of these new palettes will not be required.
        
        'Note also that this function is written in a way that allows for palettes larger than 256 colors.
        Dim leftPalette() As PDPaletteCache, rightPalette() As PDPaletteCache
        ReDim leftPalette(0 To 3) As PDPaletteCache
        ReDim rightPalette(0 To 3) As PDPaletteCache
        Dim leftCount As Long, rightCount As Long
        Dim placeLeft As Boolean
        
        For i = 0 To UBound(srcPalette)
            If (i <> targetIndex) Then
                
                If (depth = 0) Then
                    placeLeft = (srcPalette(i).ColorValue.Red < targetValue)
                ElseIf (depth = 1) Then
                    placeLeft = (srcPalette(i).ColorValue.Green < targetValue)
                ElseIf (depth = 2) Then
                    placeLeft = (srcPalette(i).ColorValue.Blue < targetValue)
                Else
                    placeLeft = (srcPalette(i).ColorValue.Alpha < targetValue)
                End If
                
                If placeLeft Then
                    If (leftCount > UBound(leftPalette)) Then ReDim Preserve leftPalette(0 To UBound(leftPalette) * 2 + 1) As PDPaletteCache
                    leftPalette(leftCount) = srcPalette(i)
                    leftCount = leftCount + 1
                Else
                    If (rightCount > UBound(rightPalette)) Then ReDim Preserve rightPalette(0 To UBound(rightPalette) * 2 + 1) As PDPaletteCache
                    rightPalette(rightCount) = srcPalette(i)
                    rightCount = rightCount + 1
                End If
                
            End If
        Next i
        
        'Trim the left and right palettes as necessary, then pass them to freshly created child nodes.
        If (leftCount > 0) Then
            If (UBound(leftPalette) <> leftCount - 1) Then ReDim Preserve leftPalette(0 To leftCount - 1) As PDPaletteCache
            m_Nodes(InsertNodeBalanced).nChildLeft = InsertNodeBalanced(leftPalette, depth + 1, useAlphaToo)
        Else
            m_Nodes(InsertNodeBalanced).nChildLeft = 0
        End If
        
        If (rightCount > 0) Then
            If (UBound(rightPalette) <> rightCount - 1) Then ReDim Preserve rightPalette(0 To rightCount - 1) As PDPaletteCache
            m_Nodes(InsertNodeBalanced).nChildRight = InsertNodeBalanced(rightPalette, depth + 1, useAlphaToo)
        Else
            m_Nodes(InsertNodeBalanced).nChildRight = 0
        End If
        
    'This palette only contains one color, meaning we can skip all the "create children node(s)" steps
    Else
        With m_Nodes(m_numNodes)
            .nColor = srcPalette(0).ColorValue
            .nChildLeft = 0
            .nChildRight = 0
            .nDepth = depth
            .nOriginalIndex = srcPalette(0).OrigIndex And &H7FFF&
        End With
        InsertNodeBalanced = m_numNodes
        m_numNodes = m_numNodes + 1
    End If
    
End Function

Friend Sub PrintDebugInfo()
    Debug.Print m_numComparisons
End Sub
