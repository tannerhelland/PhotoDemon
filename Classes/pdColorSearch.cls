VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdColorSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Self-pruning Octree-based Color Lookup Class
'Copyright 2017-2026 by Tanner Helland
'Created: 14/January/17
'Last updated: 15/January/17
'Last update: implement the self-pruning behavior, by only passing colors down branches as absolutely necessary.
'Dependencies: pdColorSearchNode
'
'This class provides a fast way to perform lossy color-matching between source colors and a palette.  At present,
' PD uses it in the Palettes.ApplyPaletteToImage_Octree() function.
'
'For a nice overview of how octrees work, see https://en.wikipedia.org/wiki/Octree.
'
'Note that this class is *not* a color quantization class.  Instead, it uses octrees to address the
' "nearest neighbor" problem of matching colors against a size-limited list.  Octrees are a natural fit for
' color-matching, as their use of 8-children perfectly matches the number of combinations for each bit-level
' of an RGB color (e.g. the high-bit of red, green, and blue have 8 different combinations.  Same for the
' next-highest bit, and the next-highest bit, etc.)
'
'I don't know if "self-pruning" is the right name for what this class does, but in a typical color-matching
' octree, colors are only ever stored at the maximum tree-depth level.  Thus, you must traverse an entire
' branch to ultimately arrive at a color descriptor.  This class operates differently; colors are stored at
' the highest possible level, and only pushed "down" a branch if a new color passes through the same node.
' This process imposes relatively little penalty during the construction stage, while making the lookup
' stage significantly faster.
'
'To use this class, first add your palette of colors via the CreateColorTree() function.  Then, retrieve
' color matches using the GetNearestPaletteIndex() function.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Root node; all searches start here
Private m_RootNode As pdColorSearchNode

Private Sub Class_Initialize()
    Set m_RootNode = New pdColorSearchNode
End Sub

Friend Sub CreateColorTree(ByRef srcPalette() As RGBQuad)
    
    'Start by creating the root node
    m_RootNode.NodeInitialize 0, srcPalette
    
    'Add all source colors to the root node!  It handles the rest for us
    Dim i As Long
    For i = 0 To UBound(srcPalette)
        m_RootNode.AddColor i
    Next i
    
End Sub

Friend Function GetNearestPaletteIndex(ByRef srcColor As RGBQuad) As Long
    Dim tmpDistance As Long
    GetNearestPaletteIndex = m_RootNode.GetNearestColorIndex(srcColor, tmpDistance)
End Function
