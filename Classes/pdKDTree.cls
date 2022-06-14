VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdKDTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon KD-Tree for fast palette matching
'Copyright 2018-2022 by Tanner Helland
'Created: 28/January/18
'Last updated: 14/June/22
'Last update: add a "fast tree construction but slower tree traversal" mode, which will be used for 3D LUT construction
'             (where a smaller number of LUT points must be matched against a huge tree of *all* colors in the image)
'Dependencies: pdKDTreeNode
'
'K-D trees (https://en.wikipedia.org/wiki/K-d_tree) are a data structure well-designed for sorting large collections
' of k-dimensional points.  One of their main advantages involves the ability to create well-balanced binary-type
' trees of k-dimensional data, which allows for very quick nearest-neighbor and range selection queries.
'
'In PD's case, I've been hunting for a better solution when matching image colors against palettes.  Reducing a
' 32-bit RGB image to some subset of (n) colors is a non-trivial task, especially as (n) grows.  A naive search is
' cost-prohibitive past ~16 colors, and given the impossibility of sorting RGB data in a meaningful way,
' standard search and/or sort approaches breakdown.
'
'KD trees, on the other hand, really shine with multidimensional data (like RGB/A coordinates).  For pathologically
' bad cases, PD's KD-tree color matcher is basically identical in performance to GDI's GetNearestPaletteIndex()
' function, while for well-distributed 3D data (e.g. palettes with many different shades of color) PD's
' implementation can be up to a full order of magnitude faster, with even larger gains if the palette size is
' larger than 256 colors, or if you need RGBA matching (which GDI doesn't implement at all).
'
'Anyway, while this implementation is specifically designed around 3D RGB and 4D RGBA data (for performance reasons),
' it would be trivial to adopt the code for any other k-dimensional data.  I may look at this in the future for
' effects built off coordinate math (like stained glass / crystallize and their underlying Voronoi engine) as it
' could provide meaningful perf improvements without some of the restrictions of our current Voronoi implementation.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_Root As pdKDTreeNode
Private m_BestDistance As Long

'The "fast build but slow retrieve" version uses a stack to construct the tree
Private Type SortStack
    sLB As Long
    sUB As Long
End Type

'Pass an entire palette to this method; the tree will be automatically constructed accordingly.
Friend Function BuildTree(ByRef srcPalette() As RGBQuad, ByVal numColorsToUse As Long) As Boolean
    
    BuildTree = (numColorsToUse > 0)
    
    If BuildTree Then
    
        Set m_Root = New pdKDTreeNode
        
        'Transfer the source palette into a "cache" type, so we can store original palette indices
        ' into the KD-tree (instead of just bare RGBQuads)
        Dim tmpCache() As PDPaletteCache
        ReDim tmpCache(0 To numColorsToUse - 1) As PDPaletteCache
        
        Dim i As Long
        For i = 0 To numColorsToUse - 1
            tmpCache(i).ColorValue = srcPalette(i)
            tmpCache(i).OrigIndex = i
        Next i
        
        'For best performance, users should request a balanced tree.  This imposes a tree creation penalty,
        ' but subsequent nearest-neighbor queries are *much* faster.
        m_Root.InsertNodeBalanced tmpCache, 0
        
    End If
    
End Function

'Pass an entire palette to this method; the tree will be automatically constructed accordingly.
Friend Function BuildTreeIncAlpha(ByRef srcPalette() As RGBQuad, ByVal numColorsToUse As Long) As Boolean
    
    BuildTreeIncAlpha = (numColorsToUse > 0)
    
    If BuildTreeIncAlpha Then
    
        Set m_Root = New pdKDTreeNode
        
        'Transfer the source palette into a "cache" type, so we can store original palette indices
        ' into the KD-tree (instead of just bare RGBQuads)
        Dim tmpCache() As PDPaletteCache
        ReDim tmpCache(0 To numColorsToUse - 1) As PDPaletteCache
        
        Dim i As Long
        For i = 0 To numColorsToUse - 1
            tmpCache(i).ColorValue = srcPalette(i)
            tmpCache(i).OrigIndex = i
        Next i
        
        'For best performance, users should request a balanced tree.  This imposes a tree creation penalty,
        ' but subsequent nearest-neighbor queries are *much* faster.
        m_Root.InsertNodeBalancedIncAlpha tmpCache, 0
        
    End If
    
End Function

'Given some source color, return the best color match from the tree
Friend Function GetNearestColor(ByRef srcColor As RGBQuad) As RGBQuad
    m_BestDistance = LONG_MAX
    If (Not m_Root Is Nothing) Then m_Root.NearestColor srcColor, GetNearestColor, m_BestDistance
End Function

Friend Function GetNearestColorIncAlpha(ByRef srcColor As RGBQuad) As RGBQuad
    m_BestDistance = LONG_MAX
    If (Not m_Root Is Nothing) Then m_Root.NearestColorIncAlpha srcColor, GetNearestColorIncAlpha, m_BestDistance
End Function

'Given some source color, return an index into the original source palette of the palette entry
' that most closely matches the requested color.
Friend Function GetNearestPaletteIndex(ByRef srcColor As RGBQuad) As Long
    m_BestDistance = LONG_MAX
    If (Not m_Root Is Nothing) Then
        Dim tmpResult As PDPaletteCache
        m_Root.NearestPaletteIndex srcColor, tmpResult, m_BestDistance
        GetNearestPaletteIndex = tmpResult.OrigIndex
    End If
End Function

'Given some source color, return an index into the original source palette of the palette entry
' that most closely matches the requested color.
Friend Function GetNearestPaletteIndexIncAlpha(ByRef srcColor As RGBQuad) As Long
    m_BestDistance = LONG_MAX
    If (Not m_Root Is Nothing) Then
        Dim tmpResult As PDPaletteCache
        m_Root.NearestPaletteIndexIncAlpha srcColor, tmpResult, m_BestDistance
        GetNearestPaletteIndexIncAlpha = tmpResult.OrigIndex
    End If
End Function

'This alternate build strategy was created during early testing phases.  It is still valid code,
' but I do not generally recommend using it in PD as the alternative implementation is *much* faster
' for normal palettes.
'
'One exception to this rule is LUT construction, where instead of solving the problem of many-colors-in-image
' to few-colors-in-palette, we have the opposite problem: few-colors-in-LUT to many-colors-in-image.  For this
' weird case, the penalty of constructing a perfectly balanced tree outweighs the perf hit of slow color retrieval,
' so it's better to use a fast build strategy (at some cost to retrieval perf).
Friend Sub BuildTreeFastButRetrievalIsSlow(ByRef srcPalette() As RGBQuad, ByVal lowBound As Long, ByVal highBound As Long)

    If (highBound < lowBound) Then Exit Sub
    Set m_Root = New pdKDTreeNode
    
    'Transfer the source palette into a "cache" type, so we can store original palette indices
    ' into the KD-tree (instead of just bare RGBQuads)
    Dim tmpCache() As PDPaletteCache
    ReDim tmpCache(0 To highBound) As PDPaletteCache
    
    Dim i As Long
    For i = lowBound To highBound
        tmpCache(i).ColorValue = srcPalette(i)
        tmpCache(i).OrigIndex = i
    Next i
    
    'Build an addition stack
    Const INIT_STACK_SIZE As Long = 256
    Dim qsRemaining() As SortStack, qsStackPtr As Long
    ReDim qsRemaining(0 To INIT_STACK_SIZE - 1) As SortStack
    qsStackPtr = 0
    qsRemaining(0).sLB = lowBound
    qsRemaining(0).sUB = highBound
    
    Do
        
        lowBound = qsRemaining(qsStackPtr).sLB
        highBound = qsRemaining(qsStackPtr).sUB
        qsStackPtr = qsStackPtr - 1
        
        'Add the requested node (the center of the current list)
        Dim curIndex As Long
        curIndex = lowBound + (highBound - lowBound) \ 2
        m_Root.InsertNode tmpCache(curIndex)
        
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
