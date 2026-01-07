VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdKDTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon KD-Tree for fast palette matching (child node)
'Copyright 2018-2026 by Tanner Helland
'Created: 28/January/18
'Last updated: 16/September/21
'Last update: simplify code where possible
'Dependencies: pdKDTree
'
'A length description of this class (and what it does) can be found in the pdKDTree class.  This class is
' not designed to be used as a standalone class; instead, it should be managed by a pdKDTree instance.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This node's color.  For performance purposes, we also cache Long-type versions so we can use them in
' various signed arithmetic operations (e.g. m_Color.Red - someOtherColor.Red can underflow).
Private m_Color As RGBQuad
Private m_Red As Long, m_Green As Long, m_Blue As Long, m_Alpha As Long

'This node's original index in the initial source palette we were passed.  Storing this allows us
' to retrieve color data as palette indices instead of just raw RGBQuads.
Private m_OriginalIndex As Long

'Lookups into tree nodes requires recursion (not within the same instance, but within parent/child
' nodes of this class).  Because we have to perform millions of lookups when applying a palette to
' an image, and each lookup requires a number of local variables, it's faster to declare those
' variables once - here - than within each lookup function.  This gives them a fixed initialization
' cost.
Private m_testDistance As Long, m_testR As Long, m_testG As Long, m_testB As Long, m_testA As Long

'The depth of this node determines how we split it when assigning children (as we must rotate
' between x/y/z axes, per the KD-tree definition).
Private m_Depth As Long
Private Const NODE_EMPTY As Long = -1

'Child nodes.  These are not guaranteed to be initialized, so please check for "Is Nothing" before using.
Private m_Left As pdKDTreeNode, m_Right As pdKDTreeNode

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
Friend Sub InsertNodeBalanced(ByRef srcPalette() As PDPaletteCache, Optional ByVal depth As Long = 0)
    
    'MOD is slower than a simple branch; cycle between values 0/1/2 for red/green/blue axes
    If (depth > 2) Then depth = 0
    
    'In this insertion mode, we *always* start by assigning this node the median value for the
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
            Else
                palHistogram(srcPalette(i).ColorValue.Blue) = palHistogram(srcPalette(i).ColorValue.Blue) + 1
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
            Else
                If (srcPalette(i).ColorValue.Blue = targetValue) Then Exit For
            End If
        Next i
        
        'i now points at the palette value we want for this node
        Dim targetIndex As Long
        targetIndex = i
        m_Color = srcPalette(targetIndex).ColorValue
        m_Red = m_Color.Red
        m_Green = m_Color.Green
        m_Blue = m_Color.Blue
        m_Depth = depth
        m_OriginalIndex = srcPalette(targetIndex).OrigIndex
        
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
                Else
                    placeLeft = (srcPalette(i).ColorValue.Blue < targetValue)
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
            If (m_Left Is Nothing) Then Set m_Left = New pdKDTreeNode
            m_Left.InsertNodeBalanced leftPalette, depth + 1
        Else
            Set m_Left = Nothing
        End If
        
        If (rightCount > 0) Then
            If (UBound(rightPalette) <> rightCount - 1) Then ReDim Preserve rightPalette(0 To rightCount - 1) As PDPaletteCache
            If (m_Right Is Nothing) Then Set m_Right = New pdKDTreeNode
            m_Right.InsertNodeBalanced rightPalette, depth + 1
        Else
            Set m_Right = Nothing
        End If
        
    'This palette only contains one color, meaning we can skip all the "create children node(s)" steps
    Else
        m_Color = srcPalette(0).ColorValue
        m_Red = m_Color.Red
        m_Green = m_Color.Green
        m_Blue = m_Color.Blue
        m_OriginalIndex = srcPalette(0).OrigIndex
        m_Depth = depth
    End If
    
End Sub

'Same as InsertNodeBalanced, but with alpha included.
Friend Sub InsertNodeBalancedIncAlpha(ByRef srcPalette() As PDPaletteCache, Optional ByVal depth As Long = 0)
    
    'MOD is slower than a simple branch; cycle between values 0/1/2/3 for red/green/blue/alpha axes
    If (depth > 3) Then depth = 0
    
    'In this insertion mode, we *always* start by assigning this node the median value for the
    ' current depth (0, 1, 2, 3 for red, green, blue, alpha).  After assigning this node its value,
    ' we will split the remaining palette entries into two groups (one lower, one greater-than-or-equal);
    ' these groups are then passed to left and right child nodes, who will proceed identically.
    
    'Normally, finding the median of a data set requires sorting, but sorting is expensive.
    ' RGBA data is discrete and on a fixed range, so we can cheat and use histograms to find the
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
        
        'i now points at the palette value we want for this node
        Dim targetIndex As Long
        targetIndex = i
        m_Color = srcPalette(targetIndex).ColorValue
        m_Red = m_Color.Red
        m_Green = m_Color.Green
        m_Blue = m_Color.Blue
        m_Alpha = m_Color.Alpha
        m_Depth = depth
        m_OriginalIndex = srcPalette(targetIndex).OrigIndex
        
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
            If (m_Left Is Nothing) Then Set m_Left = New pdKDTreeNode
            m_Left.InsertNodeBalancedIncAlpha leftPalette, depth + 1
        Else
            Set m_Left = Nothing
        End If
        
        If (rightCount > 0) Then
            If (UBound(rightPalette) <> rightCount - 1) Then ReDim Preserve rightPalette(0 To rightCount - 1) As PDPaletteCache
            If (m_Right Is Nothing) Then Set m_Right = New pdKDTreeNode
            m_Right.InsertNodeBalancedIncAlpha rightPalette, depth + 1
        Else
            Set m_Right = Nothing
        End If
        
    'This palette only contains one color, meaning we can skip all the "create children node(s)" steps
    Else
        m_Color = srcPalette(0).ColorValue
        m_Red = m_Color.Red
        m_Green = m_Color.Green
        m_Blue = m_Color.Blue
        m_Alpha = m_Color.Alpha
        m_OriginalIndex = srcPalette(0).OrigIndex
        m_Depth = depth
    End If
    
End Sub

'Return the color from this tree (and its children) that is closest to some specified source color.
' This function is called recursively, which is why everything must be passed to it as ByRef.
' (That said, the srcColor value can be treated as a CONST, but VB doesn't allow const pointers, so...
' ...yeah.  It is what it is.)
Friend Sub NearestColor(ByRef srcColor As RGBQuad, ByRef curBestColor As RGBQuad, ByRef curBestDistance As Long)
    
    'Before checking child nodes, compare the target color against this node's color.
    m_testR = m_Red - srcColor.Red
    m_testG = m_Green - srcColor.Green
    m_testB = m_Blue - srcColor.Blue
    m_testDistance = m_testR * m_testR + m_testG * m_testG + m_testB * m_testB
    
    'Store the best (closest) result so far
    If (m_testDistance < curBestDistance) Then
        curBestDistance = m_testDistance
        curBestColor = m_Color
    End If
    
    'Next, we want to determine if any of our child nodes contain points "closer" or "further" from
    ' our current best-match color.  (NOTE: in early builds, this was coded more cleanly, with local
    ' nearDirection and farDirection pdKDTreeNode objects being assigned to left or right children,
    ' accordingly - but this code is used inside tight per-pixel loops, so I've manually expanded
    ' and in-lined a bunch of code for improved performance.)
    
    'Note that we use our previously stored depth to determine which axis to use for comparisons,
    ' and to simplify the code, we use the existing variables "r" and "g" to hold "current" and
    ' "module" level values, e.g. at m_Depth = 0, r = srcColor.Red and g = m_Red.
    Select Case m_Depth
        Case 0
            m_testR = srcColor.Red
            m_testG = m_Red
        Case 1
            m_testR = srcColor.Green
            m_testG = m_Green
        Case 2
            m_testR = srcColor.Blue
            m_testG = m_Blue
    End Select
    
    'If the target color is *less than* this node's color, better matches will be found in
    ' the left tree.  (Conversely, if it is *greater than or equal to* this node's color,
    ' search the right tree first.)
    If (m_testR < m_testG) Then
        If (Not m_Left Is Nothing) Then m_Left.NearestColor srcColor, curBestColor, curBestDistance
        
        'Now we need to repeat some ugly steps (but writing it this way avoids unnecessary branching,
        ' which kills performance).  We next need to see if it's necessary to check the right branch
        ' of ths tree as well.  We do this by testing the theoretical "closest point" possible in
        ' the right branch, and if that "theoretical" point is closer to the target color than our
        ' current best match, we need to search the right branch for possible targets as well.
        If (Not m_Right Is Nothing) Then
            
            'We know that the best value for this tree currently lies in the left branch.  In order for
            ' a value in the right branch to be closer than the current value, it would need to be the
            ' *smallest possible value* in that tree - meaning a color with an r value as low as possible.
            ' Because our KD-tree implementation uses "greater-than-or-equal-to" for right branch
            ' determination, the lowest possible value in right branches is an r-value equal to the target
            ' color's.  (We ignore green and blue because they could potentially be *equal* to the target
            ' color, but we have no way of knowing that as this node only branches on red!)
            m_testDistance = m_testR - m_testG
        
            'If the closest "theoretical" point in the right branch is closer than the current best match,
            ' let's query that child for a best match.
            If ((m_testDistance * m_testDistance) < curBestDistance) Then m_Right.NearestColor srcColor, curBestColor, curBestDistance
            
        End If
        
    Else
        If (Not m_Right Is Nothing) Then m_Right.NearestColor srcColor, curBestColor, curBestDistance
        If (Not m_Left Is Nothing) Then
            
            'Because we're querying the left tree, the nearest possible color would have to be at least
            ' one less than this node's color.  As such, if this node has a value of 0, there is no
            ' possible way that the left node could contain a closer color (as it can't contain
            ' *any* colors less than zero!)
            If (m_testR > 0) Then
                m_testDistance = m_testR - 1 - m_testG
                If ((m_testDistance * m_testDistance) < curBestDistance) Then m_Left.NearestColor srcColor, curBestColor, curBestDistance
            End If
        End If
    End If
    
End Sub

'Same as NearestColor, but alpha values are included in the calculation.
Friend Sub NearestColorIncAlpha(ByRef srcColor As RGBQuad, ByRef curBestColor As RGBQuad, ByRef curBestDistance As Long)
    
    'Before checking child nodes, compare the target color against this node's color.
    m_testR = m_Red - srcColor.Red
    m_testG = m_Green - srcColor.Green
    m_testB = m_Blue - srcColor.Blue
    m_testA = m_Alpha - srcColor.Alpha
    m_testDistance = m_testR * m_testR + m_testG * m_testG + m_testB * m_testB + m_testA * m_testA
    
    If (m_testDistance < curBestDistance) Then
        curBestDistance = m_testDistance
        curBestColor = m_Color
    End If
    
    'Next, we want to determine if any of our child nodes contain points "closer" or "further" from
    ' our current best-match color.  (NOTE: in early builds, this was coded more cleanly, with local
    ' nearDirection and farDirection pdKDTreeNode objects being assigned to left or right children,
    ' accordingly - but this code is used inside tight per-pixel loops, so I've manually expanded
    ' and in-lined a bunch of code for improved performance.)
    
    'Note that we use our previously stored depth to determine which axis to use for comparisons,
    ' and to simplify the code, we use the existing variables "r" and "g" to hold "current" and
    ' "module" level values, e.g. at m_Depth = 0, r = srcColor.Red and g = m_Red.
    Select Case m_Depth
        Case 0
            m_testR = srcColor.Red
            m_testG = m_Red
        Case 1
            m_testR = srcColor.Green
            m_testG = m_Green
        Case 2
            m_testR = srcColor.Blue
            m_testG = m_Blue
        Case 3
            m_testR = srcColor.Alpha
            m_testG = m_Alpha
    End Select
    
    'If the target color is *less than* this node's color, better matches will be found in
    ' the left tree.  (Conversely, if it is *greater than or equal to* this node's color,
    ' search the right tree first.)
    If (m_testR < m_testG) Then
        If (Not m_Left Is Nothing) Then m_Left.NearestColorIncAlpha srcColor, curBestColor, curBestDistance
        
        'Now we need to repeat some ugly steps (but writing it this way avoids unnecessary branching,
        ' which kills performance).  We next need to see if it's necessary to check the right branch
        ' of ths tree as well.  We do this by testing the theoretical "closest point" possible in
        ' the right branch, and if that "theoretical" point is closer to the target color than our
        ' current best match, we need to search the right branch for possible targets as well.
        If (Not m_Right Is Nothing) Then
            
            'We know that the best value for this tree currently lies in the left branch.  In order for
            ' a value in the right branch to be closer than the current value, it would need to be the
            ' *smallest possible value* in that tree - meaning a color with an r value as low as possible.
            ' Because our KD-tree implementation uses "greater-than-or-equal-to" for right branch
            ' determination, the lowest possible value in right branches is an r-value equal to the target
            ' color's.  (We ignore green and blue because they could potentially be *equal* to the target
            ' color, but we have no way of knowing that as this node only branches on red!)
            m_testDistance = m_testR - m_testG
        
            'If the closest "theoretical" point in the right branch is closer than the current best match,
            ' let's query that child for a best match.
            If ((m_testDistance * m_testDistance) < curBestDistance) Then m_Right.NearestColorIncAlpha srcColor, curBestColor, curBestDistance
            
        End If
        
    Else
        If (Not m_Right Is Nothing) Then m_Right.NearestColorIncAlpha srcColor, curBestColor, curBestDistance
        If (Not m_Left Is Nothing) Then
            
            'Because we're querying the left tree, the nearest possible color would have to be at least
            ' one less than this node's color.  As such, if this node has a value of 0, there is no
            ' possible way that the left node could contain a closer color (as it can't contain
            ' *any* colors less than zero!)
            If (m_testR > 0) Then
                m_testDistance = m_testR - 1 - m_testG
                If ((m_testDistance * m_testDistance) < curBestDistance) Then m_Left.NearestColorIncAlpha srcColor, curBestColor, curBestDistance
            End If
        End If
    End If
    
End Sub

'Return the original palette index of the palette color from this tree (and its children) that is closest
' to some specified source color.  This function is called recursively, which is why everything must be
' passed to it as ByRef.  (That said, the srcColor value can be treated as a CONST, but VB doesn't allow
' const pointers, so... yeah.  It is what it is.)
Friend Sub NearestPaletteIndex(ByRef srcColor As RGBQuad, ByRef curBestColor As PDPaletteCache, ByRef curBestDistance As Long)
    
    'Before checking child nodes, compare the target color against this node's color.
    m_testR = m_Red - srcColor.Red
    m_testG = m_Green - srcColor.Green
    m_testB = m_Blue - srcColor.Blue
    m_testDistance = m_testR * m_testR + m_testG * m_testG + m_testB * m_testB
    
    If (m_testDistance < curBestDistance) Then
        curBestDistance = m_testDistance
        curBestColor.ColorValue = m_Color
        curBestColor.OrigIndex = m_OriginalIndex
    End If
    
    'Next, we want to determine which child node contains points "closer" or "further" from our
    ' current best-match color.  (NOTE: in early builds, this was coded more cleanly, with local
    ' nearDirection and farDirection pdKDTreeNode objects being assigned to left or right children,
    ' accordingly - but this code is used inside tight per-pixel loops, so I've manually expanded
    ' and in-lined a bunch of code to improve performance.)
    
    'Note that we use our previously stored depth to determine which axis to use for comparisons,
    ' and to simplify the code, we use the existing variables "r" and "g" to hold "current" and
    ' "module" level values, e.g. at m_Depth = 0, r = srcColor.Red and g = m_Red.
    Select Case m_Depth
        Case 0
            m_testR = srcColor.Red
            m_testG = m_Red
        Case 1
            m_testR = srcColor.Green
            m_testG = m_Green
        Case 2
            m_testR = srcColor.Blue
            m_testG = m_Blue
        Case 3
            m_testR = srcColor.Alpha
            m_testG = m_Alpha
    End Select
        
    'If the target color is *less than* this node's color, better matches will be found in
    ' the left tree.  (Conversely, if it is *greater than or equal to* this node's color,
    ' search the right tree first.)
    If (m_testR < m_testG) Then
        If (Not m_Left Is Nothing) Then m_Left.NearestPaletteIndex srcColor, curBestColor, curBestDistance
        
        'Now we need to repeat some ugly steps (but writing it this way avoids unnecessary branching,
        ' which kills performance).  We next need to see if it's necessary to check the right branch
        ' of ths tree as well.  We do this by testing the theoretical "closest point" possible in
        ' the right branch, and if that "theoretical" point is closer to the target color than our
        ' current best match, we need to search the right branch for possible targets as well.
        If (Not m_Right Is Nothing) Then
            
            'We know that the best value for this tree lay in the left branch, meaning the closest
            ' value in the right branch would be the *smallest possible value* - meaning a color
            ' with an r value as low as possible.  Because our KD-tree implementation uses
            ' "greater-than-or-equal-to" for right branch determination, the lowest possible value in
            ' the branch is an r-value equal to the target color's.
            m_testDistance = m_testR - m_testG
        
            'If the closest "theoretical" point in the right branch is closer than the current best match,
            ' let's query that child for a best match.
            If ((m_testDistance * m_testDistance) < curBestDistance) Then m_Right.NearestPaletteIndex srcColor, curBestColor, curBestDistance
            
        End If
        
    Else
        If (Not m_Right Is Nothing) Then m_Right.NearestPaletteIndex srcColor, curBestColor, curBestDistance
        If (Not m_Left Is Nothing) Then
            
            'Because we're querying the left tree, the nearest possible color would have to be at least
            ' one less than this node's color.  As such, if this node has a value of 0, there is no
            ' possible way that the left node could contain a closer color (as it can't contain
            ' *any* colors less than zero!)
            If (m_testR > 0) Then
                m_testDistance = m_testR - 1 - m_testG
                If ((m_testDistance * m_testDistance) < curBestDistance) Then m_Left.NearestPaletteIndex srcColor, curBestColor, curBestDistance
            End If
        End If
    End If
    
End Sub

'Same as NearestPaletteIndex, but with alpha included in the calculation
Friend Sub NearestPaletteIndexIncAlpha(ByRef srcColor As RGBQuad, ByRef curBestColor As PDPaletteCache, ByRef curBestDistance As Long)
    
    'Before checking child nodes, compare the target color against this node's color.
    m_testR = m_Red - srcColor.Red
    m_testG = m_Green - srcColor.Green
    m_testB = m_Blue - srcColor.Blue
    m_testA = m_Alpha - srcColor.Alpha
    m_testDistance = m_testR * m_testR + m_testG * m_testG + m_testB * m_testB + m_testA * m_testA
    
    If (m_testDistance < curBestDistance) Then
        curBestDistance = m_testDistance
        curBestColor.ColorValue = m_Color
        curBestColor.OrigIndex = m_OriginalIndex
    End If
    
    'Next, we want to determine which child node contains points "closer" or "further" from our
    ' current best-match color.  (NOTE: in early builds, this was coded more cleanly, with local
    ' nearDirection and farDirection pdKDTreeNode objects being assigned to left or right children,
    ' accordingly - but this code is used inside tight per-pixel loops, so I've manually expanded
    ' and in-lined a bunch of code to improve performance.)
    
    'Note that we use our previously stored depth to determine which axis to use for comparisons,
    ' and to simplify the code, we use the existing variables "r" and "g" to hold "current" and
    ' "module" level values, e.g. at m_Depth = 0, r = srcColor.Red and g = m_Red.
    Select Case m_Depth
        Case 0
            m_testR = srcColor.Red
            m_testG = m_Red
        Case 1
            m_testR = srcColor.Green
            m_testG = m_Green
        Case 2
            m_testR = srcColor.Blue
            m_testG = m_Blue
        Case 3
            m_testR = srcColor.Alpha
            m_testG = m_Alpha
    End Select
        
    'If the target color is *less than* this node's color, better matches will be found in
    ' the left tree.  (Conversely, if it is *greater than or equal to* this node's color,
    ' search the right tree first.)
    If (m_testR < m_testG) Then
        If (Not m_Left Is Nothing) Then m_Left.NearestPaletteIndexIncAlpha srcColor, curBestColor, curBestDistance
        
        'Now we need to repeat some ugly steps (but writing it this way avoids unnecessary branching,
        ' which kills performance).  We next need to see if it's necessary to check the right branch
        ' of ths tree as well.  We do this by testing the theoretical "closest point" possible in
        ' the right branch, and if that "theoretical" point is closer to the target color than our
        ' current best match, we need to search the right branch for possible targets as well.
        If (Not m_Right Is Nothing) Then
            
            'We know that the best value for this tree lay in the left branch, meaning the closest
            ' value in the right branch would be the *smallest possible value* - meaning a color
            ' with an r value as low as possible.  Because our KD-tree implementation uses
            ' "greater-than-or-equal-to" for right branch determination, the lowest possible value in
            ' the branch is an r-value equal to the target color's.
            m_testDistance = m_testR - m_testG
        
            'If the closest "theoretical" point in the right branch is closer than the current best match,
            ' let's query that child for a best match.
            If ((m_testDistance * m_testDistance) < curBestDistance) Then m_Right.NearestPaletteIndexIncAlpha srcColor, curBestColor, curBestDistance
            
        End If
        
    Else
        If (Not m_Right Is Nothing) Then m_Right.NearestPaletteIndexIncAlpha srcColor, curBestColor, curBestDistance
        If (Not m_Left Is Nothing) Then
            
            'Because we're querying the left tree, the nearest possible color would have to be at least
            ' one less than this node's color.  As such, if this node has a value of 0, there is no
            ' possible way that the left node could contain a closer color (as it can't contain
            ' *any* colors less than zero!)
            If (m_testR > 0) Then
                m_testDistance = m_testR - 1 - m_testG
                If ((m_testDistance * m_testDistance) < curBestDistance) Then m_Left.NearestPaletteIndexIncAlpha srcColor, curBestColor, curBestDistance
            End If
        End If
    End If
    
End Sub

Private Sub Class_Initialize()
    m_Depth = NODE_EMPTY
End Sub
