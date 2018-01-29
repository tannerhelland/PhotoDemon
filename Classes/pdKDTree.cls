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
'KD Tree for fast RGB palette matching

Option Explicit

Private m_Root As pdKDTreeNode

'Pass an entire palette to this method; the tree will be automatically constructed accordingly.
Friend Function BuildTree(ByRef srcPalette() As RGBQuad, ByVal numColorsToUse As Long) As Boolean
    
    BuildTree = (numColorsToUse > 0)
    
    If BuildTree Then
    
        Set m_Root = New pdKDTreeNode
    
        Dim i As Long
        For i = 0 To numColorsToUse - 1
            m_Root.InsertNode srcPalette(i)
        Next i
    
    End If
    
End Function

'Given some source color, return the best color match from the tree
Friend Function GetNearestColor(ByRef srcColor As RGBQuad) As RGBQuad
    
    'Everything is passed ByRef, so these points will be updated as the function processes
    Dim bestDistance As Double, curBestColor As RGBQuad
    bestDistance = DOUBLE_MAX
    
    If (Not m_Root Is Nothing) Then
        m_Root.NearestColor srcColor, curBestColor, bestDistance
        GetNearestColor = curBestColor
    End If
    
End Function
