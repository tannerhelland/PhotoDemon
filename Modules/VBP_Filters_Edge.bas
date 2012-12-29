Attribute VB_Name = "Filters_Edge"
'***************************************************************************
'Filter (Edge) Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 12/June/01
'Last updated: 05/September/12
'Last update: rewrote and optimized all filters against the new layer class.
'
'Runs all edge-related filters (edge detection, relief, etc.).
'
'***************************************************************************

Option Explicit

'Redraw the image using a pencil sketch effect.
Public Sub FilterPencil()
    
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    
    FM(-1, -1) = -1
    FM(-1, 0) = -1
    FM(-1, 1) = -1
    
    FM(0, -1) = -1
    FM(0, 0) = 8
    FM(0, 1) = -1
    
    FM(1, -1) = -1
    FM(1, 0) = -1
    FM(1, 1) = -1
    
    FilterWeight = 1
    FilterBias = 0
    
    DoFilter "pencil sketch", True

End Sub

'A typical relief filter, that makes the image seem pseudo-3D.
Public Sub FilterRelief()
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, -1) = 2
    FM(-1, 0) = 1
    FM(0, 1) = 1
    FM(0, 0) = 1
    FM(0, -1) = -1
    FM(1, 0) = -1
    FM(1, 1) = -2
    FilterWeight = 3
    FilterBias = 75
    DoFilter "Relief"
End Sub

'A lighter version of a traditional sharpen filter; it's designed to bring out edge detail without the blowout typical of sharpening
Public Sub FilterEdgeEnhance()
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, 0) = -1
    FM(1, 0) = -1
    FM(0, -1) = -1
    FM(0, 1) = -1
    FM(0, 0) = 8
    FilterWeight = 4
    FilterBias = 0
    DoFilter "Edge Enhance"
End Sub
