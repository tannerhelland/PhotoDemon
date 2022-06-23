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

Private m_curR As Long, m_curG As Long, m_curB As Long, m_curA As Long
Private m_testR As Long, m_testG As Long, m_testB As Long, m_testA As Long

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
        InsertNode srcPalette(curIndex), curIndex
        
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

Private Sub InsertNode(ByRef srcColor As RGBQuad, ByVal originalIndex As Long)
    
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
            Case 0
                useLeftBranch = (srcColor.Red < tmpColor.Red)
            Case 1
                useLeftBranch = (srcColor.Green < tmpColor.Green)
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
                    .nOriginalIndex = originalIndex And &H7FFF&
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
                    .nOriginalIndex = originalIndex And &H7FFF&
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
