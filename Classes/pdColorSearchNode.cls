VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdColorSearchNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Self-pruning Octree-based Color Lookup Class (Node only)
'Copyright 2017-2026 by Tanner Helland
'Created: 14/January/17
'Last updated: 15/January/17
'Last update: implement self-pruning behavior, by only passing colors down branches as absolutely necessary.
'Dependencies: pdColorSearch
'
'For details on how this class works, please refer to its parent class: pdColorSearch
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Used for testing color weighting by human eye sensitivity.  (Not used at present; results are unpredictable.)
'Private Const CUSTOM_WEIGHT_RED As Single = 0.299!
'Private Const CUSTOM_WEIGHT_GREEN As Single = 0.587!
'Private Const CUSTOM_WEIGHT_BLUE As Single = 0.114!

'Local copy of the source palette.  By caching this, we can perform all internal comparisons
' by palette index instead of by color.  This provides a meaningful performance boost,
' given that we'll traverse the tree millions of times on a typical image.
Private m_ColorList() As RGBQuad

'Depth of this node.  Nodes on the outermost level of the tree return their color directly,
' rather than querying child nodes.
Private m_Depth As Long

'Palette index of this node, if any, with a few constants to make searching the tree easier
Private m_PaletteIndex As Long
Private Const PALETTE_ENTRY_UNINITIALIZED As Long = -1
Private Const PALETTE_ENTRY_INVALID As Long = -2

'Cached bit masks.  Since VB doesn't support bit-shift operators, we have to optimize this
' where we can.
Private m_BitShiftMasks() As Long

'Child nodes.  These are not guaranteed to be initialized, so make sure to check against Nothing before accessing.
Private m_ChildNodes(0 To 7) As pdColorSearchNode

'Constructor
Friend Sub NodeInitialize(ByVal nodeDepth As Long, ByRef srcColorList() As RGBQuad)

    m_Depth = nodeDepth
    
    'Caching the full color list in each node requires little memory, and it potentially accelerates
    ' a number of search-related functions.
    ReDim m_ColorList(0 To UBound(srcColorList)) As RGBQuad
    CopyMemoryStrict VarPtr(m_ColorList(0)), VarPtr(srcColorList(0)), UBound(srcColorList) * 4 + 4

End Sub

'Add colors to this tree.  Because VB6 doesn't support custom constructors, make sure you manually call
' NodeInitialize() before adding anything!
Friend Sub AddColor(ByVal colorIndex As Long)

    'The final node level (node 7, which is the last because there are only 8 bits per color!)
    ' only contains direct palette indices, not children.
    If (m_Depth = 7) Then
        m_PaletteIndex = colorIndex
    
    'Interior nodes may point deeper into the tree.
    Else
        
        'If this node = -1, then it is freshly initialized, without any children.  Just store the
        ' palette index here, rather than plunging deeper into the tree.
        If (m_PaletteIndex = PALETTE_ENTRY_UNINITIALIZED) Then
            m_PaletteIndex = colorIndex
        
        'If this node <> -1, then we have child nodes, and we need to pass the color down the tree.
        Else
            
            Dim childIndex As Long
            
            'Are we currently storing a color?  If we are, we need to also pass *that* color down the tree.
            If (m_PaletteIndex <> PALETTE_ENTRY_INVALID) Then
            
                'Convert the palette color at this level to a bitmask representing a specific child node
                childIndex = GetOctIndex(m_ColorList(m_PaletteIndex), m_Depth)
                
                'Pass this color down the line
                If (m_ChildNodes(childIndex) Is Nothing) Then
                    Set m_ChildNodes(childIndex) = New pdColorSearchNode
                    m_ChildNodes(childIndex).NodeInitialize m_Depth + 1, m_ColorList
                End If
                
                m_ChildNodes(childIndex).AddColor m_PaletteIndex
                
                'Note that this node now has child nodes
                m_PaletteIndex = PALETTE_ENTRY_INVALID
            
            End If
            
            'Convert the newly added color to a bitmask specific to this level (e.g. one that properly
            ' represents the correct child node at this tree level).
            childIndex = GetOctIndex(m_ColorList(colorIndex), m_Depth)
            
            'Pass this color down the line
            If (m_ChildNodes(childIndex) Is Nothing) Then
                Set m_ChildNodes(childIndex) = New pdColorSearchNode
                m_ChildNodes(childIndex).NodeInitialize m_Depth + 1, m_ColorList
            End If
            
            m_ChildNodes(childIndex).AddColor colorIndex
            
        End If
    
    End If
    
End Sub

Friend Function GetNearestColorIndex(ByRef srcColor As RGBQuad, ByRef dstToColor As Long) As Long
    
    'There are two reasons to return a color immediately, rather than further traversing the tree.
    ' 1) This node directly stores a color (meaning it has no child nodes)
    ' 2) We are at the deepest node in the tree, so children are impossible
    If (m_PaletteIndex >= 0) Or (m_Depth = 7) Then
    
        GetNearestColorIndex = m_PaletteIndex
        
        'Calculate distance.  Our parent may be querying all child nodes for a best match, which means
        ' it needs distance reports to know which child is closest to the target.
        Dim rDist As Long, gDist As Long, bDist As Long
        With m_ColorList(m_PaletteIndex)
            rDist = CLng(srcColor.Red) - .Red
            gDist = CLng(srcColor.Green) - .Green
            bDist = CLng(srcColor.Blue) - .Blue
        End With
        
        'At present, a well-generated palette produces a sparse enough tree that color-weighting doesn't provide
        ' much benefit.  However, it may be beneficial to enable this in the future, particularly if we add the
        ' option to build palettes via octrees.
        'dstToColor = (rDist * rDist) * CUSTOM_WEIGHT_RED + (gDist * gDist) * CUSTOM_WEIGHT_GREEN + (bDist * bDist) * CUSTOM_WEIGHT_BLUE
        dstToColor = (rDist * rDist) + (gDist * gDist) + (bDist * bDist)
        
    'We are not at maximum depth, and we aren't directly storing a color.  Query our children.
    Else

        'Translate the requested color into a child index
        Dim childIndex As Long
        childIndex = GetOctIndex(srcColor, m_Depth)
        
        'See if a child exists at that index.  If it does, plunge deeper into the tree.
        If (Not m_ChildNodes(childIndex) Is Nothing) Then
            GetNearestColorIndex = m_ChildNodes(childIndex).GetNearestColorIndex(srcColor, dstToColor)
            
        'A child doesn't exist at the requested node.  The best we can do is return one of our children,
        ' specifically the one with the closest-matching color.
        Else
            
            Dim minDistance As Long: minDistance = 999999999
            Dim minColorIndex As Long
            Dim curDistance As Long, curNode As Long
            
            Dim i As Long
            For i = 0 To 7
                If (Not m_ChildNodes(i) Is Nothing) Then
                    
                    'Ask the child to return the best match from its sub-tree
                    curNode = m_ChildNodes(i).GetNearestColorIndex(srcColor, curDistance)
                    
                    'If this match is better than a previous one, cache it
                    If (curDistance <= minDistance) Then
                        minDistance = curDistance
                        minColorIndex = curNode
                    End If
                    
                End If
            Next i
            
            'Return the best match from our subtree
            GetNearestColorIndex = minColorIndex
            dstToColor = minDistance
            
        End If
        
    End If

End Function

'Given a color, return the relevant child node for that color at the requested tree level.
' (In a color-based octree, each level of the octree corresponds to a bit in each color channel.
'  Most significant bits are stored first, with lesser bits used as we go deeper into the tree.)
'
'Note that RGB order does not actually matter, meaning you could map RGB channels to any particular
' bit-order.  They are always mapped to identical subtrees (e.g. changing the order of the subtrees
' doesn't actually change the colors returned).
Private Function GetOctIndex(ByRef srcColor As RGBQuad, ByVal treeDepth As Long) As Long
    If (srcColor.Red And m_BitShiftMasks(treeDepth)) <> 0 Then GetOctIndex = 1
    If (srcColor.Green And m_BitShiftMasks(treeDepth)) <> 0 Then GetOctIndex = GetOctIndex Or 2
    If (srcColor.Blue And m_BitShiftMasks(treeDepth)) <> 0 Then GetOctIndex = GetOctIndex Or 4
End Function

'Because VB makes bit-shifting difficult, we pre-generate bit masks in advance.
Private Sub Class_Initialize()
    m_PaletteIndex = PALETTE_ENTRY_UNINITIALIZED
    ReDim m_BitShiftMasks(0 To 7) As Long
    Dim i As Long
    For i = 0 To 7
        m_BitShiftMasks(i) = 2 ^ (7 - i)
    Next i
End Sub
