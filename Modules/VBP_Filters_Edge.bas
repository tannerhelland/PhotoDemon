Attribute VB_Name = "Filters_Edge"
'***************************************************************************
'Filter (Edge) Interface
'Copyright ©2000-2013 by Tanner Helland
'Created: 12/June/01
'Last updated: 05/September/12
'Last update: rewrote and optimized all filters against the new layer class.
'
'Runs all edge-related filters (edge detection, relief, etc.).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Redraw the image using a pencil sketch effect.
Public Sub FilterPencil()
    
    g_FilterSize = 3
    ReDim g_FM(-1 To 1, -1 To 1) As Long
    
    g_FM(-1, -1) = -1
    g_FM(-1, 0) = -1
    g_FM(-1, 1) = -1
    
    g_FM(0, -1) = -1
    g_FM(0, 0) = 8
    g_FM(0, 1) = -1
    
    g_FM(1, -1) = -1
    g_FM(1, 0) = -1
    g_FM(1, 1) = -1
    
    g_FilterWeight = 1
    g_FilterBias = 0
    
    DoFilter g_Language.TranslateMessage("pencil sketch"), True

End Sub

'A typical relief filter, that makes the image seem pseudo-3D.
Public Sub FilterRelief()
    g_FilterSize = 3
    ReDim g_FM(-1 To 1, -1 To 1) As Long
    g_FM(-1, -1) = 2
    g_FM(-1, 0) = 1
    g_FM(0, 1) = 1
    g_FM(0, 0) = 1
    g_FM(0, -1) = -1
    g_FM(1, 0) = -1
    g_FM(1, 1) = -2
    g_FilterWeight = 3
    g_FilterBias = 75
    DoFilter g_Language.TranslateMessage("Relief")
End Sub

'A lighter version of a traditional sharpen filter; it's designed to bring out edge detail without the blowout typical of sharpening
Public Sub FilterEdgeEnhance()
    g_FilterSize = 3
    ReDim g_FM(-1 To 1, -1 To 1) As Long
    g_FM(-1, 0) = -1
    g_FM(1, 0) = -1
    g_FM(0, -1) = -1
    g_FM(0, 1) = -1
    g_FM(0, 0) = 8
    g_FilterWeight = 4
    g_FilterBias = 0
    DoFilter g_Language.TranslateMessage("Edge Enhance")
End Sub
