Attribute VB_Name = "Filters_Edge"
'***************************************************************************
'Filter (Edge) Interface
'©2000-2012 Tanner Helland
'Created: 6/12/01
'Last updated: 1/25/03
'
'Runs all edge-related filters (edge detection, relief, etc.).
'
'***************************************************************************

Option Explicit

Public Sub FilterPencil()
    Message "Generating pencil image..."
    Dim tR As Integer, tG As Integer, tB As Integer
    Dim c As Long, d As Long
    Dim tColor As Byte
    SetProgBarMax PicWidthL
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1)
    For x = 1 To PicWidthL - 1
    For y = 1 To PicHeightL - 1
        tR = 0
        tG = 0
        tB = 0
        For c = x - 1 To x + 1
        For d = y - 1 To y + 1
            If c = x And d = y Then
                tR = tR + 8 * ImageData(c * 3 + 2, d)
                tG = tG + 8 * ImageData(c * 3 + 1, d)
                tB = tB + 8 * ImageData(c * 3, d)
            Else
                tR = tR - ImageData(c * 3 + 2, d)
                tG = tG - ImageData(c * 3 + 1, d)
                tB = tB - ImageData(c * 3, d)
            End If
        Next d
        Next c
        ByteMe tR
        ByteMe tG
        ByteMe tB
        tColor = (tR + tG + tB) \ 3
        tColor = 255 - tColor
        tData(x * 3 + 2, y) = tColor
        tData(x * 3 + 1, y) = tColor
        tData(x * 3, y) = tColor
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetProgBarVal cProgBar.Max
    TransferImageData
    SetImageData
End Sub

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

Private Sub TransferImageData()
    Message "Transferring data..."
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        For z = 0 To 2
            ImageData(QuickVal + z, y) = tData(QuickVal + z, y)
        Next z
    Next y
    Next x
End Sub

