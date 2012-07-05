VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SavePCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Autor: ALKO
'e-mail: alfred.koppold@freenet.de
Option Explicit
  Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
   Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long

Private BitCnt As Long
Private Bildhöhe As Long
Private Bildbreite As Long
Private Bildbreiteright As Long
Private CompPCX() As Byte
Private Compressed As Boolean
Private Bitcbr As Long 'BPL

Private Const DIB_RGB_COLORS = 0
Private Const BMPBIT = 50
Private Const BI_RGB = 0&

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Private Type RGBTRE
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
End Type
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Private Type BITMAPINFO8Bit
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type
Private Type PCXHeader
    ZSoftID As Byte
    Version As Byte
    RLE As Byte
    BitsPerPixel As Byte
    XMin As Integer
    YMin As Integer
    xMax As Integer
    yMax As Integer
    HRes As Integer
    VRes As Integer
    Palette16(0 To 47) As Byte
    Reserved As Byte
    nPlanes As Byte
    BytesperLine As Integer
    PaletteType As Integer
    Info(0 To 57) As Byte
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

'HAcK BY TANNER
Dim lStart As Long

Public Sub SavePCXinFile(filename As String, Pic As PictureBox, Optional ByVal BitspPix As Long = 24, Optional ByVal compress As Boolean = True)
    lStart = 0
    Dim a As Long
    Dim i As Long
    Dim iBitmap As Long, iDC As Long
    Dim b As Long
    Dim TempPal(0 To 255) As RGBQUAD
    Dim bBits() As Byte
    Dim Dibbreite As Long
    Dim bi24BitInfo As BITMAPINFO
Dim Übergabe() As Byte
Compressed = compress
    BitCnt = BitspPix
    a = Pic.Parent.ScaleMode
    b = Pic.ScaleMode
    Pic.Parent.ScaleMode = 3
    Pic.ScaleMode = 3
    Bildhöhe = Pic.ScaleHeight
    Bildbreite = Pic.ScaleWidth
    Pic.Parent.ScaleMode = a
    Pic.ScaleMode = b
    Bildbreiteright = Bildbreite
    Select Case BitspPix
        Case 1
        If Bildbreite Mod 8 = 0 Then
        Bitcbr = Bildbreite \ 8
        Else
        Bitcbr = Bildbreite \ 8 + 1
        End If
        If Bitcbr Mod 2 <> 0 Then
        Bitcbr = Bitcbr + 1 'for PCX
        End If
        If Bitcbr Mod 4 = 0 Then
        Dibbreite = Bitcbr
        Else
        Dibbreite = (Bitcbr \ 4 + 1) * 4 'forBMPDIB
        End If
            BitCnt = 1
 
        Case 4
            If Bildbreite Mod 2 = 0 Then
            Bitcbr = Bildbreite \ 2
            Else
            Bildbreite = Bildbreite + 1
            Bitcbr = Bildbreite \ 2
            End If
            If Bitcbr Mod 2 <> 0 Then
            Bitcbr = Bitcbr + 1
            End If
        If Bitcbr Mod 4 = 0 Then
        Dibbreite = Bitcbr
        Else
        Dibbreite = (Bitcbr \ 4 + 1) * 4 'forBMPDIB
        End If


            BitCnt = 4

        Case 8
            If Bildbreite Mod 2 <> 0 Then
    Bildbreite = Bildbreite + 1
    End If

                Bitcbr = Bildbreite

            BitCnt = 8
    If Bildbreite Mod 4 = 0 Then '8 bit
    Dibbreite = Bildbreite
    Else
    Dibbreite = Bildbreite + 2
    End If

        Case 24
            If Bildbreite Mod 2 <> 0 Then
    Bildbreite = Bildbreite + 1
    End If

                Bitcbr = Bildbreite * 3

            BitCnt = 24
    If Bildbreite * 3 Mod 4 = 0 Then '24 bit
    Dibbreite = Bildbreite * 3
    Else
    Dibbreite = Bildbreite * 3 + 2
    End If
    End Select
        ReDim bBits((Dibbreite) * Bildhöhe - 1)

        With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = Bildbreiteright
        .biHeight = Bildhöhe
    End With

    iDC = CreateCompatibleDC(Pic.hdc)
    iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject iDC, iBitmap
    BitBlt iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, Pic.hdc, 0, 0, vbSrcCopy

    GetIconBits bBits(), BitCnt + BMPBIT, iDC, iBitmap, TempPal(), Bildhöhe, Bildbreite
    ReDim Übergabe(Bitcbr * Bildhöhe - 1)
  If Bitcbr <> Dibbreite Then
  For i = 0 To Bildhöhe - 1
  CopyMemory Übergabe(i * Bitcbr), bBits(i * Dibbreite), Bitcbr
  Next i
  Else
    CopyMemory Übergabe(0), bBits(0), UBound(bBits) + 1
  End If
        DeleteDC iDC
    DeleteObject iBitmap

    'HAcK BY TANNER
    'Dim TmpPalette() As RGBQUAD
    'ReDim TmpPalette(1 To UBound(iPalette)) As RGBQUAD
    'For x = 1 To UBound(iPalette)
    '    TmpPalette(x).rgbRed = iPalette(x).r
    '    TmpPalette(x).rgbGreen = iPalette(x).g
    '    TmpPalette(x).rgbBlue = iPalette(x).b
    'Next x
    'If HasBeenReduced = True And BitCnt = 8 Then
    '    SavePCX Filename, iData(), Bildbreite, Bildhöhe, BitCnt, TmpPalette()
    'Else
        SavePCX filename, Übergabe(), Bildbreite, Bildhöhe, BitCnt, TempPal()
    'End If

End Sub
Private Sub SavePCX(DSTFile As String, BitmapBytes() As Byte, PCXWidth As Long, PCXHeight As Long, BBP As Long, BitmapPalette() As RGBQUAD)
    Dim PCXMap() As Byte
    Dim FileHeader As PCXHeader
    Dim i As Long
    Dim PCXRed() As Byte
    Dim PCXGreen() As Byte
    Dim PCXBlue() As Byte
    Dim Übergabe() As Byte
    Dim a As Long
    
    FileHeader.ZSoftID = 10
    FileHeader.Version = 5
    Select Case Compressed
    Case False
    FileHeader.RLE = 0
    Case True
    FileHeader.RLE = 1
    End Select
    If BBP >= 8 Then
        FileHeader.BitsPerPixel = 8
        If BBP = 8 Then
            FileHeader.nPlanes = 1
        ElseIf BBP = 24 Then
            FileHeader.nPlanes = 3
        End If
    Else
        FileHeader.BitsPerPixel = BBP
        FileHeader.nPlanes = 1
    End If
    
    FileHeader.XMin = 1
    FileHeader.YMin = 1
    FileHeader.xMax = Bildbreiteright
    FileHeader.yMax = PCXHeight
    FileHeader.HRes = 96
    FileHeader.VRes = 96
    FileHeader.Reserved = 0
    FileHeader.PaletteType = 0
    
    Select Case BBP
        Case 1
            FileHeader.BytesperLine = Bitcbr
            
            MakePCXMap BitmapBytes, PCXMap, PCXWidth, PCXHeight, BBP
            If Compressed = True Then CompressPCX PCXMap, Bitcbr
            If Dir(DSTFile) <> "" Then Kill DSTFile
            
            Open DSTFile For Binary As #1
                Put #1, , FileHeader
                Put #1, , PCXMap
            Close
        Case 4
            FileHeader.BytesperLine = Bitcbr
            
            MakePCXPalette BitmapPalette, FileHeader.Palette16, 16
            MakePCXMap BitmapBytes, PCXMap, PCXWidth, PCXHeight, BBP
            If Compressed = True Then CompressPCX PCXMap, PCXWidth

            If Dir(DSTFile) <> "" Then Kill DSTFile
            
            Open DSTFile For Binary As #1
                Put #1, , FileHeader
                Put #1, , PCXMap
            Close
        Case 8
            FileHeader.BytesperLine = PCXWidth
            
            Dim PCXPalette(0 To 767) As Byte
            
            'HAcK BY TANNER
            'If HasBeenReduced = True Then
            '    lStart = 1
            '    MakePCXPalette BitmapPalette, PCXPalette, UBound(BitmapPalette) + 1
            'Else
                MakePCXPalette BitmapPalette, PCXPalette, 256
            'End If
            
            MakePCXMap BitmapBytes, PCXMap, PCXWidth, PCXHeight, BBP
            If Compressed = True Then CompressPCX PCXMap, PCXWidth

            If Dir(DSTFile) <> "" Then Kill DSTFile
            Dim PalId As Byte
            Open DSTFile For Binary As #1
                Put #1, , FileHeader
                Put #1, , PCXMap
                PalId = 12
                Put #1, , PalId
                Put #1, , PCXPalette
            Close
        Case 24
            FileHeader.BytesperLine = PCXWidth
    
            MakePCX24 BitmapBytes, PCXMap, PCXWidth, PCXHeight, PCXWidth
            If Dir(DSTFile) <> "" Then Kill DSTFile
            CompressPCX PCXMap, PCXWidth
            Open DSTFile For Binary As #1
                Put #1, , FileHeader
                Put #1, , CompPCX
            Close
        Case Else
            Exit Sub
    End Select
End Sub
Private Sub MakePCXMap(BitmapBytes() As Byte, PCXMap() As Byte, PCXWidth As Long, PCXHeight As Long, Bpp As Long)
Dim Größe As Long
'Copy only the bytes in PCXMap
'Select Case BPP
'Case 1
'Größe = PCXWidth * PCXHeight / 8 - 1
'Case 4
'Größe = PCXWidth * PCXHeight / 2 - 1
'Case 8
'Größe = PCXWidth * PCXHeight - 1
'End Select
Größe = UBound(BitmapBytes) + 1
ReDim PCXMap(Größe)
CopyMemory PCXMap(0), BitmapBytes(0), Größe + 1 'UBound(BitmapBytes) + 1
End Sub

Private Sub MakePCXPalette(BitmapPalette() As RGBQUAD, PCXPalette() As Byte, nColors As Long)
    Dim Cnt As Long
    
    For Cnt = lStart To nColors - 1
        PCXPalette(Cnt * 3) = BitmapPalette(Cnt).rgbRed
        PCXPalette(Cnt * 3 + 1) = BitmapPalette(Cnt).rgbGreen
        PCXPalette(Cnt * 3 + 2) = BitmapPalette(Cnt).rgbBlue
    Next
End Sub
Private Sub MakePCX24(Feld() As Byte, PCXMap() As Byte, nWidth As Long, nHeight As Long, BytesperLine)
    Dim x As Long, y As Long, nStartPos As Long
    Dim High As Long
        Dim quad() As RGBTRE
        ReDim quad(nWidth * nHeight)
        CopyMemory quad(0), Feld(0), (nWidth * nHeight * 3)
        ReDim PCXMap(nWidth * nHeight * 3 - 1)
    For y = 0 To nHeight - 1
        For x = 0 To nWidth - 1
        High = nHeight - y - 1
            nStartPos = High * BytesperLine * 3 + x '0-31 96-127 192-
            With quad((x + (nHeight - 1 - y) * nWidth)) ' zuerst 992- 1023 Ende 0-31
                 PCXMap(nStartPos) = .rgbRed
                PCXMap(nStartPos + BytesperLine) = .rgbGreen
                PCXMap(nStartPos + BytesperLine * 2) = .rgbBlue
            End With
        
        Next x
    Next y
    
End Sub
Private Sub CompressPCX(ArraytoComp() As Byte, Linewidth As Long)
Dim i As Long
Dim b As Long
Dim Zähler As Long
Dim Wo As Long
Dim Linezahl As Long

ReDim CompPCX((UBound(ArraytoComp) * 2) + ((UBound(ArraytoComp) \ Linewidth) * 2)) ' größtmögliche Größe
'Linezahl = 1
Do While i <= UBound(ArraytoComp)
If i + 1 <= UBound(ArraytoComp) Then
Do Until ArraytoComp(i) <> ArraytoComp(i + 1) Or Zähler = 62
If i + 1 = (Linewidth * Linezahl) And Linewidth Mod 2 <> 0 Then Exit Do ' +1 da Grund = 0
i = i + 1
Zähler = Zähler + 1
If i = UBound(ArraytoComp) Then Exit Do
Loop
Else

End If
If Zähler = 0 Then

If ArraytoComp(i) < 192 Then
CompPCX(b) = ArraytoComp(i) ' 1* kleiner 192
Wo = b
b = b + 1
If i + 1 = (Linewidth * Linezahl) And Linewidth Mod 2 <> 0 Then 'falls Zeilenende und ungerade Zahl
CompPCX(b) = 0
Wo = b
b = b + 1
Linezahl = Linezahl + 1
End If
Else
CompPCX(b) = 193
CompPCX(b + 1) = ArraytoComp(i)  '1* größer 192
Wo = b + 1
b = b + 2
If i + 1 = (Linewidth * Linezahl) And Linewidth Mod 2 <> 0 Then 'falls Zeilenende und ungerade Zahl
CompPCX(b) = 0
Wo = b
b = b + 1
Linezahl = Linezahl + 1
End If
End If

Else
CompPCX(b) = 192 + Zähler + 1 ' bis 255 ist Zähler = 62
CompPCX(b + 1) = ArraytoComp(i) 'Zähler * Byte
Wo = b + 1
b = b + 2
If i + 1 = (Linewidth * Linezahl) And Linewidth Mod 2 <> 0 Then 'falls Zeilenende und ungerade Zahl
CompPCX(b) = 0
Wo = b
b = b + 1
Linezahl = Linezahl + 1
End If
End If

Zähler = 0 'Zähler zurücksetzen
i = i + 1
Loop
ReDim Preserve CompPCX(Wo) 'zuschneiden
ReDim ArraytoComp(Wo)
CopyMemory ArraytoComp(0), CompPCX(0), Wo + 1
End Sub
Private Sub GetIconBits(bBits() As Byte, Bpp As Long, nDC As Long, nBitmap As Long, CopyArr() As RGBQUAD, Height As Long, Breite As Long)
    Dim BI As BITMAPINFO8Bit
    Dim Test As Long
    If Bpp > BMPBIT Then
        Bpp = Bpp - BMPBIT
        BI.bmiHeader.biHeight = 0 - Height
    Else
        BI.bmiHeader.biHeight = Height
    End If
    BI.bmiHeader.biBitCount = Bpp
    BI.bmiHeader.biCompression = BI_RGB
    BI.bmiHeader.biPlanes = 1
    BI.bmiHeader.biWidth = Breite
    BI.bmiHeader.biSize = Len(BI.bmiHeader)
    Test = GetDIBits(nDC, nBitmap, 0, Height, bBits(0), BI, DIB_RGB_COLORS)
    If Bpp = 1 Then
        CopyMemory CopyArr(0), BI.bmiColors(0), Len(CopyArr(0)) * 2
    ElseIf Bpp = 4 Then
        CopyMemory CopyArr(0), BI.bmiColors(0), Len(CopyArr(0)) * 16
    ElseIf Bpp = 8 Then
        CopyMemory CopyArr(0), BI.bmiColors(0), Len(CopyArr(0)) * 256
    End If
End Sub

