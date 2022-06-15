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
'Copyright 2018-2022 by Tanner Helland
'Created: 28/January/18
'Last updated: 15/June/22
'Last update: copy "fast tree construction but slower tree traversal" mode out of the base KD-tree class, and rewrite
'             it here against a backing array (to improve horrifically slow class teardown when 100,000+ instances are
'             necessary).
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

'The "fast build but slow retrieve" version uses a stack to construct the tree
Private Type SortStack
    sLB As Long
    sUB As Long
End Type

'Each tree element is stored in a backing array.  This backing array removes the need for child classes,
' which changes the performance trade-offs involved - this class is much faster to build and teardown than
' the normal KD-tree class, but slower to traverse.  (This is possibly caused by locality - normal class
' creation may produce left/right child classes that are physically closer together in memory vs a linear
' array like this.)  Because the teardown penalty of the normal class is so severe, however, (we're talking
' a teardown time in *minutes* if the class count exceeds several hundred thousand instances), this is a
' critical trade-off when I need a KD-tree with many tens of thousands of individual nodes.
Private Type KDNode
    nColor As RGBQuad
    nChildLeft As Long
    nChildRight As Long
    nDepth As Long      'VB will pad the struct to 4-byte alignment regardless, so may as well make this a Long
End Type

Private m_numNodes As Long
Private m_Nodes() As KDNode

'When retrieving colors, I assume the caller is gonna retrieve a bunch of colors in the row.
' (This class makes no sense otherwise.)  To improve performance, let's reuse a single search stack
' across all retrieval calls.
Private m_SearchStack() As Long

Friend Sub BuildTree(ByRef srcPalette() As RGBQuad, ByVal lowBound As Long, ByVal highBound As Long)

    If (highBound < lowBound) Then Exit Sub
    
    'Initialize the backing array as precisely as possible.  (This spares us from needing additional
    ' array bound checks on the insert loop.)
    m_numNodes = 0
    ReDim m_Nodes(0 To (highBound - lowBound)) As KDNode
    
    'Manually populate the first node
    With m_Nodes(0)
        .nColor = srcPalette(0)
        .nDepth = 0
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
        InsertNode srcPalette(curIndex)
        
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
    
End Sub

'Given some source color, return the best color match from the tree
Friend Function GetNearestColor(ByRef srcColor As RGBQuad) As RGBQuad
    
    'Failsafe only; build the damn tree before trying to retrieve colors from it!
    If (m_numNodes <= 0) Then Exit Function
    
    Dim bestDistance As Long, testDistance As Long
    bestDistance = LONG_MAX
    
    Dim testColor As RGBQuad
    
    'We keep a running stack of nodes to check.  When we exhaust this list, we stop iterating nodes.
    
    'The first node we check is always the root node
    m_SearchStack(0) = 0
    
    Dim numStackNodes As Long
    numStackNodes = 1
    
    'RGBQuads store color values as bytes.  These overflow when we try to do basic arithmetic with them,
    ' so we must copy them into save Long values.
    Dim curR As Long, curG As Long, curB As Long
    Dim testR As Long, testG As Long, testB As Long
    
    Dim idxTest As Long, testDepth As Long
    
    Do While (numStackNodes > 0)
        
        'Decrement stack counter and make that our current index
        numStackNodes = numStackNodes - 1
        idxTest = m_SearchStack(numStackNodes)
        testDepth = m_Nodes(idxTest).nDepth
        
        'Before checking child nodes, compare the target color against this node's color.
        testColor = m_Nodes(idxTest).nColor
        curR = testColor.Red
        curG = testColor.Green
        curB = testColor.Blue
        
        testR = curR - srcColor.Red
        testG = curG - srcColor.Green
        testB = curB - srcColor.Blue
        testDistance = testR * testR + testG * testG + testB * testB
        
        'Store the best (closest) result so far
        If (testDistance < bestDistance) Then
            bestDistance = testDistance
            GetNearestColor = testColor
        End If
        
        'Next, we want to determine if any of our child nodes contain potential points "closer to" or
        ' "further from" our current best-match color.  (Excuse the ugly code here; I've manually inlined
        ' code for better performance.)
        '
        'Note that we use our previously stored depth to determine which axis to use for comparisons,
        ' and to simplify the code, we use the existing variables "r" and "g" to hold "current" and
        ' "module" level values, e.g. at m_Depth = 0, r = srcColor.Red and g = m_Red.
        Select Case testDepth
            Case 0
                testR = srcColor.Red
                testG = curR
            Case 1
                testR = srcColor.Green
                testG = curG
            Case 2
                testR = srcColor.Blue
                testG = curB
        End Select
        
        'We will potentially need to search the sub-trees beneath both child nodes, so to simplify
        ' stack boundary checks, ensure room for at least two more stack entries.
        If (numStackNodes >= UBound(m_SearchStack) - 1) Then ReDim Preserve m_SearchStack(0 To numStackNodes * 2 + 1) As Long
        
        'If the target color is *less than* this node's color, better matches will be found in
        ' the left tree.  (Conversely, if it is *greater than or equal to* this node's color,
        ' search the right tree first.)
        If (testR < testG) Then
        
            'Search the left tree
            testB = m_Nodes(idxTest).nChildLeft
            If (testB > 0) Then
                m_SearchStack(numStackNodes) = testB
                numStackNodes = numStackNodes + 1
            End If
            
            'Now we need to repeat some ugly steps (but writing it this way minimizes branching for
            ' improved perf).  We next need to see if it's necessary to check the right branch
            ' of ths tree as well.  We do this by testing the theoretical "closest point" possible in
            ' the right branch, and if that "theoretical" point is closer to the target color than our
            ' current best match, we need to search the right branch for possible targets as well.
            testB = m_Nodes(idxTest).nChildRight
            If (testB > 0) Then
                
                'We know that the best value for this tree *likely* lies in the left branch.  In order for
                ' a value in the right branch to be closer than the current value, it would need to be the
                ' *smallest possible value* in that tree - so at depth 0, this means a color with an r value
                ' as low as possible.
                '
                'Because our KD-tree implementation uses "greater-than-or-equal-to" for right branch
                ' determination, the lowest possible value in right branches is an r-value equal to the
                ' target color's.  (We ignore green and blue because they could potentially be *equal* to
                ' the target color, but we have no way of knowing that as this node only branches on red!)
                testDistance = testR - testG
                
                'If the closest "theoretical" point in the right branch is closer than the current best match,
                ' we must also search that sub-tree for a best match.
                If ((testDistance * testDistance) < bestDistance) Then
                    m_SearchStack(numStackNodes) = testB
                    numStackNodes = numStackNodes + 1
                End If
                
            End If
            
        Else
            
            'Search the right tree
            testB = m_Nodes(idxTest).nChildRight
            If (testB > 0) Then
                m_SearchStack(numStackNodes) = testB
                numStackNodes = numStackNodes + 1
            End If
            
            testB = m_Nodes(idxTest).nChildLeft
            If (testB > 0) Then
                
                'Because we're querying the left tree, the nearest possible color would have to be at least
                ' one less than this node's color.  As such, if this node has a value of 0, there is no
                ' possible way that the left node could contain a closer color (as it can't contain
                ' *any* colors less than zero!)
                If (testR > 0) Then
                    testDistance = testR - 1 - testG
                    If ((testDistance * testDistance) < bestDistance) Then
                        m_SearchStack(numStackNodes) = testB
                        numStackNodes = numStackNodes + 1
                    End If
                End If
                
            End If
            
        End If
            
    Loop
    
End Function

Private Sub InsertNode(ByRef srcColor As RGBQuad)
    
    'We always start on the R-axis (note that the root node must already be filled for this function to work)
    Dim depth As Long
    
    Dim curNode As Long
    curNode = 0
    
    Dim tmpColor As RGBQuad
    
    'We'll iterate until we find an empty node; loop termination is manual via Exit Do on appropriate paths
    Do
        
        'Depth comes from the current child node
        depth = m_Nodes(curNode).nDepth
        
        'Compare the target color against this node's color.  If the passed color is "higher" or "lower" than
        ' the target color (which component we compare depends on depth) we'll traverse down the respective
        ' branch of the tree.
        tmpColor = m_Nodes(curNode).nColor
        
        'Do a quick search for equality; if found, return our color and exit.
        ' (Note the manual short-circuiting to work around VB deficiencies.)
        If (srcColor.Red = tmpColor.Red) Then
            If (srcColor.Green = tmpColor.Green) Then
                If (srcColor.Blue = tmpColor.Blue) Then
                    Exit Do
                End If
            End If
        End If
        
        'Rotate between RGB axes
        Dim useLeftBranch As Boolean
        
        Select Case depth

            'Red
            Case 0
                useLeftBranch = (srcColor.Red < tmpColor.Red)
                
            'Green
            Case 1
                useLeftBranch = (srcColor.Green < tmpColor.Green)

            'Blue
            Case 2
                useLeftBranch = (srcColor.Blue < tmpColor.Blue)
                
        End Select
        
        'Increment depth before continuing with traversal
        depth = depth + 1
        If (depth > 2) Then depth = 0
        
        'We will now assign the color to empty child indices, or continue traversing the tree as necessary
        
        'Left branch
        If useLeftBranch Then
            
            'No child
            If (m_Nodes(curNode).nChildLeft = 0) Then
                
                'Add this color as a new child, then exit
                m_Nodes(curNode).nChildLeft = m_numNodes
                With m_Nodes(m_numNodes)
                    .nColor = srcColor
                    .nDepth = depth
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
                    .nDepth = depth
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

Private Sub Class_Initialize()
    
    'Prevent the need for "is array initialized "checks
    Const INIT_STACK_SIZE As Long = 256
    ReDim m_SearchStack(0 To INIT_STACK_SIZE - 1) As Long
    
End Sub
