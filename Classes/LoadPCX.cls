VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadPCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'This code was originally written by Alfred Koppold, and was sent to me via personal correspondence.

'I believe you can download a publicly available version via this link:
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=56537&lngWId=1

' Many thanks to Alfred for providing a way to access PCX files natively in VB6

'Autor: ALKO
'e-mail: alfred.koppold@freenet.de

Option Explicit

'Types
Private Type ScTw
TwipWidth As Long
TwipHeight As Long
End Type

Private Type RGBQUAD
rgbBlue As Byte
rgbGreen As Byte
rgbRed As Byte
rgbReserved As Byte
End Type
Private Type RGBTRIPLE
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type PCXHeader
    Manufacturer As Byte  '10 = ZSoft
    Version As Byte 'Version
    Encoding As Byte    '1 = .PCX RLE
    Bpp As Byte    '1, 2, 4, 8
    XMin As Integer
    YMin As Integer
    xMax As Integer
    yMax As Integer
    HDpi As Integer
    VDpi As Integer
    ColourPalette(0 To 15) As RGBTRIPLE
    Reserved1 As Byte
    Planes As Byte
    BytesperLine As Integer
    PaletteInfo As Integer
    HScreenSize As Integer
    VScreenSize As Integer
    Reserved2(0 To 53)  As Byte
End Type

'Functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Variables
Private nLineSize As Long
Private BitmapData() As Byte
Private TwipWidth As Long
Private TwipHeight As Long
Private i As Long
Private Orientation As Integer
Private nWidth As Long
Private nHeight As Long
Private Scales As Integer
Private Automatic As Boolean
Private TW As ScTw
Private Header As PCXHeader

Public Function LoadPCX(ByVal FileName As String) As StdPicture
    
    Dim nFreefile As Integer
    
    nFreefile = FreeFile

    Open FileName For Binary Lock Write As #nFreefile
    
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    Close #nFreefile
    With Header
        nWidth = .xMax - .XMin + 1
        nHeight = .yMax - .YMin + 1
        nLineSize = .Planes * .BytesperLine
    End With
TW = PixelToTwips(nWidth, nHeight)

    If Header.Bpp = 1 Then
        
        If Header.Planes = 1 Then
        Read1Bit FileName
        
        End If
                If Header.Planes = 4 Then
                Read4bit FileName
                    
        
        End If

        End If
        If Header.Bpp = 4 Then
        If Header.Planes = 1 Then
        Read4bit FileName
        End If
        End If
        If Header.Bpp = 8 Then
        If Header.Planes = 1 Then
         Read8Bit FileName
         ElseIf Header.Planes = 3 Then
         Read24bit FileName
End If
End If
End Function

Private Sub Read8Bit(ByRef FileName As String)
    'Variables and arrays
    Dim nFreefile As Integer
    Dim Header As PCXHeader
    Dim Palette8(0 To 255) As RGBTRIPLE
    Dim PalByte As Byte
    Dim Übergabe() As Byte
    Const cStartOfPalette As Long = 12
    
    nFreefile = FreeFile
    'Open File
    Open FileName For Binary Lock Write As #nFreefile
        'Read the header
    Get #nFreefile, , Header
    'Get data
    ReDim BitmapData(LOF(nFreefile) - Len(Header))
    Get #nFreefile, , BitmapData()
    'Get palette indication byte
    Seek #nFreefile, LOF(nFreefile) - 768
    Get #nFreefile, , PalByte
        'Get Palette
    If PalByte = cStartOfPalette Then
        Seek #nFreefile, LOF(nFreefile) - 767
        Get #nFreefile, , Palette8()
    Else
        'Not correct.
        For i = 0 To 255
        Palette8(i).Blue = i
        Palette8(i).Green = i
        Palette8(i).Red = i
        Next i
    End If
    Close #nFreefile

'==================================================
    
    Dim nXSize As Long, NYSize As Long
    Dim nLineSize As Long
    
    nXSize = Header.xMax - Header.XMin + 1
    NYSize = Header.yMax - Header.YMin + 1
    nLineSize = Header.Planes * Header.BytesperLine

'==================================================
    PixelToTwips nXSize, NYSize
            ReDim Übergabe(UBound(Palette8) * 3 + 3)
        CopyMemory Übergabe(0), Palette8(0), UBound(Übergabe)
        InitColorTable_8 Übergabe
If Header.Encoding = 1 Then
    DecodePcx BitmapData
    End If
    MakeBitmap BitmapData, NYSize, nLineSize
    CreateBitmap_8 BitmapData, nXSize, NYSize, Orientation

End Sub

Private Function Read1Bit(FileName As String)
    'Variables and arrays
    Dim nFreefile As Integer
    Dim Header As PCXHeader
    nFreefile = FreeFile

'Read data
Open FileName For Binary Lock Write As #nFreefile
        'Read the header
    Get #nFreefile, , Header
    'Get data
    ReDim BitmapData(LOF(nFreefile) - Len(Header))
    Get #nFreefile, , BitmapData()
    Close #nFreefile
'==================================================
    
    Dim nXSize As Long, NYSize As Long
    Dim nLineSize As Long
    
    nXSize = Header.xMax - Header.XMin + 1
    NYSize = Header.yMax - Header.YMin + 1
    nLineSize = Header.Planes * Header.BytesperLine

'==================================================


    PixelToTwips nXSize, NYSize
    If Header.Encoding = 1 Then
DecodePcx BitmapData
End If
MakeBitmap BitmapData, NYSize, nLineSize
InitColorTable_1
CreateBitmap_1 BitmapData, nXSize, NYSize, Orientation

End Function

Private Function Read4bit(FileName As String)
    'Variables and arrays
    Dim nFreefile As Integer
    Dim Header As PCXHeader
    Dim Übergabe() As Byte
    Const cStartOfPalette As Long = 12
    'Find first free file handle.
    nFreefile = FreeFile
'Read the data
    Open FileName For Binary Lock Write As #nFreefile
    'Read the header
    Get #nFreefile, , Header
    'Get data
    ReDim BitmapData(LOF(nFreefile) - Len(Header))
    Get #nFreefile, , BitmapData()
    Close #nFreefile
'==================================================
    Dim nXSize As Long, NYSize As Long, Newsize As Long
    Dim nLineSize As Long
    
    nXSize = Header.xMax - Header.XMin + 1
    NYSize = Header.yMax - Header.YMin + 1
    If nXSize Mod 2 = 0 Then
    Newsize = nXSize
    Else
    Newsize = nXSize + 1
    End If
If (Newsize / 2) Mod 2 = 0 Then
'Bildbreite ist durch 2 teilbar
nLineSize = Newsize / 2
Else
nLineSize = (Newsize / 2) + 1 'gerade Zahl machen
End If

'==================================================
    PixelToTwips nXSize, NYSize
            ReDim Übergabe(UBound(Header.ColourPalette) * 3 + 3)
        CopyMemory Übergabe(0), Header.ColourPalette(0), UBound(Übergabe)
        InitColorTable_4 Übergabe

    If Header.Encoding = 1 Then
    DecodePcx BitmapData
    End If
    If Header.Planes = 4 Then
    Dim Test As Long
    Dim Übg() As Byte
    ReDim Übg(UBound(BitmapData))
    Test = BMPBreite(nXSize)
    Test = Test / 2 * NYSize - 1 ' Bildgröße
    Übg = SortPlanes(BitmapData, Header.BytesperLine, nXSize, NYSize)
        ReDim BitmapData(Test)
    CopyMemory BitmapData(0), Übg(0), Test + 1
    End If
    MakeBitmap BitmapData, NYSize, nLineSize
  CreateBitmap_4 BitmapData, nXSize, NYSize, Orientation
  
End Function

Private Function Read24bit(FileName As String)
    Dim nFreefile As Integer
    Dim Header As PCXHeader
    Dim Data() As Byte
    Dim quad() As RGBQUAD
    
    nFreefile = FreeFile
    Open FileName For Binary Lock Write As #nFreefile
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    
    ReDim Data(LOF(nFreefile) - Len(Header))
    
    Get #nFreefile, , Data()
    
    Close #nFreefile

'==================================================
    Dim nXSize As Long, NYSize As Long
    
    nXSize = Header.xMax - Header.XMin + 1
    NYSize = Header.yMax - Header.YMin + 1
     nLineSize = Header.Planes * Header.BytesperLine

'==================================================
    PixelToTwips nXSize, NYSize

    
    Dim nBitmapX As Long
    
    If (nWidth) Mod Len(nBitmapX) = 0 Then
   
        nBitmapX = nWidth
   
    Else
   
        nBitmapX = (nWidth \ 4) * 4 + 3
   
    End If


    ReDim quad(nBitmapX * nHeight)

Dim Größe As Long
Größe = nBitmapX * nHeight
If Header.Encoding = 1 Then
DecodePcx Data
End If
'Sort from planes into a bitmap
'==================================================

    Dim x As Long, y As Long, nStartPos As Long
    
    For y = 0 To nHeight - 1
        For x = 0 To nWidth - 1
        
            nStartPos = y * Header.BytesperLine * 3 + x
        
            With quad((x + (nHeight - 1 - y) * nWidth))
                .rgbRed = Data(nStartPos)
                .rgbGreen = Data(nStartPos + Header.BytesperLine)
                .rgbBlue = Data(nStartPos + Header.BytesperLine * 2)
            End With
        
        Next x
    Next y
    

   ReDim BitmapData(UBound(quad) * 4 + 4)
   CopyMemory BitmapData(0), quad(0), UBound(BitmapData)
CreateBitmap_24 BitmapData, nWidth, nHeight, Orientation

End Function

Private Sub DecodePcx(ImageArray() As Byte)
Dim Übergabe() As Byte
Dim Stand As Long
Dim i As Long
Dim x As Long
Dim n As Long
Dim c As Byte
Dim Länge As Long

ReDim Übergabe(UBound(ImageArray))

CopyMemory Übergabe(0), ImageArray(0), UBound(ImageArray)
ReDim ImageArray(0)
For Länge = 0 To UBound(Übergabe) - 1
x = Übergabe(Länge)
If x >= 192 Then
n = x - 192
c = Übergabe(Länge + 1)
Länge = Länge + 1
Else
n = 1
c = x
End If
For i = 1 To n
ReDim Preserve ImageArray(Stand)
ImageArray(Stand) = c
Stand = Stand + 1
Next i
Next Länge

End Sub

Private Sub MakeBitmap(ImageArray() As Byte, Lines As Long, BytesLine As Long)
Dim Übergabe() As Byte
Dim Größe As Long
Dim GrößeBMP As Long
Dim i As Long
Dim Zugabe As Integer
Dim Standort As Long
Dim nBitmapX As Long
    
    If (BytesLine) Mod Len(nBitmapX) = 0 Then
    
        nBitmapX = BytesLine - 1
   
    Else
   
        nBitmapX = (BytesLine \ 4) * 4 + 3
   
    End If
If nBitmapX + 1 <> BytesLine Then
Zugabe = nBitmapX - BytesLine + 1
End If
Größe = Lines * BytesLine

GrößeBMP = Lines * (nBitmapX + 1) - 1
ReDim Übergabe(UBound(ImageArray))

CopyMemory Übergabe(0), ImageArray(0), UBound(ImageArray) + 1

ReDim ImageArray(GrößeBMP)
For i = 0 To BytesLine * Lines - BytesLine Step BytesLine
CopyMemory ImageArray(Standort), Übergabe(Größe - i - BytesLine), BytesLine
Standort = Standort + nBitmapX + 1

Next i

End Sub



Private Function BMPBreite(Bildbreite As Long) As Long
If Bildbreite Mod 2 <> 0 Then
Bildbreite = Bildbreite + 1
End If
If Bildbreite / 2 Mod 2 = 0 Then
'Bildbreite (durch 2 da 2 Pixel pro Byte) ist durch 2 teilbar
'Bildbreite = Bildbreite - alles o. K.
Else
Bildbreite = Bildbreite + 2
End If
BMPBreite = Bildbreite
End Function
Private Function SortPlanes(Arr1() As Byte, BitesperLineroh As Integer, Bildbreite As Long, Bildhöhe As Long) As Byte()
Dim a As Long
Dim Bitesperline As Long
Dim fertig() As Byte
Dim i As Long
Dim Zähler As Long
Dim Byte1 As Byte
Dim Byte2 As Byte
Dim Byte3 As Byte
Dim Byte4 As Byte
Dim Größe As Long
On Error GoTo allesfertig
Dim Bytefertig() As Byte
ReDim Bytefertig(3)
If Bildbreite Mod 2 <> 0 Then
Bildbreite = Bildbreite + 1
End If
If Bildbreite / 2 Mod 2 = 0 Then
'Bildbreite (durch 2 da 2 Pixel pro Byte) ist durch 2 teilbar
'Bildbreite = Bildbreite - alles o. K.
Else
Bildbreite = Bildbreite + 2
End If
Größe = Bildbreite / 2 * Bildhöhe
Bitesperline = Bildbreite / 2
ReDim fertig(Größe)
Dim Bildroh As Long
Bildroh = Bildhöhe * BitesperLineroh * 4
For i = 0 To Bildroh - BitesperLineroh * 4 Step BitesperLineroh * 4 'Zeilen
Dim t As Long
Dim v As Long
If Bitesperline Mod 4 = 0 Then
v = Bitesperline / 4 - 1
Else
v = Bitesperline / 4
End If

Zähler = 0
For t = 0 To v
Byte1 = Arr1(i + t)
Byte2 = Arr1(i + BitesperLineroh + t)
Byte3 = Arr1(i + (2 * BitesperLineroh) + t)
Byte4 = Arr1(i + (3 * BitesperLineroh) + t)
Bytefertig = BytesSort(Byte1, Byte2, Byte3, Byte4)
Zähler = Zähler + 1
If Zähler <= Bitesperline Then
fertig(a) = Bytefertig(0)
a = a + 1
End If
Zähler = Zähler + 1
If Zähler <= Bitesperline Then
fertig(a) = Bytefertig(1)
If a = Größe Then GoTo allesfertig
a = a + 1
End If
Zähler = Zähler + 1
If Zähler <= Bitesperline Then
fertig(a) = Bytefertig(2)
If a = Größe Then GoTo allesfertig
a = a + 1
End If
Zähler = Zähler + 1
If Zähler <= Bitesperline Then
fertig(a) = Bytefertig(3)
If a = Größe Then GoTo allesfertig
a = a + 1
End If
Next t
Next i
allesfertig:

SortPlanes = fertig

End Function

Private Function GetByte(Bytes As Byte, Position As Long) As Integer
GetByte = 0
Select Case Position
Case 1
If Bytes And 128 Then GetByte = 1
Case 2
If Bytes And 64 Then GetByte = 1
Case 3
If Bytes And 32 Then GetByte = 1
Case 4
If Bytes And 16 Then GetByte = 1
Case 5
If Bytes And 8 Then GetByte = 1
Case 6
If Bytes And 4 Then GetByte = 1
Case 7
If Bytes And 2 Then GetByte = 1
Case 8
If Bytes And 1 Then GetByte = 1
End Select

End Function


Private Function BytesSort(Byte1 As Byte, Byte2 As Byte, Byte3 As Byte, Byte4 As Byte) As Byte()
Dim fertig(3) As Byte
Dim a As Long
Dim c As Long
For c = 1 To 7 Step 2
If GetByte(Byte1, c) = 1 Then fertig(a) = fertig(a) + 16
If GetByte(Byte2, c) = 1 Then fertig(a) = fertig(a) + 32
If GetByte(Byte3, c) = 1 Then fertig(a) = fertig(a) + 64
If GetByte(Byte4, c) = 1 Then fertig(a) = fertig(a) + 128

If GetByte(Byte1, c + 1) = 1 Then fertig(a) = fertig(a) + 1
If GetByte(Byte2, c + 1) = 1 Then fertig(a) = fertig(a) + 2
If GetByte(Byte3, c + 1) = 1 Then fertig(a) = fertig(a) + 4
If GetByte(Byte4, c + 1) = 1 Then fertig(a) = fertig(a) + 8
a = a + 1
Next c
BytesSort = fertig
End Function

Private Function PixelToTwips(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
ux = Screen.TwipsPerPixelX
PixelToTwips.TwipWidth = xwert * ux
uy = Screen.TwipsPerPixelY
PixelToTwips.TwipHeight = ywert * uy
End Function

Public Property Get ScaleMode() As Integer
ScaleMode = Scales
End Property

Public Property Let ScaleMode(ByVal vNewValue As Integer)
If vNewValue > 0 Then vNewValue = 1
If vNewValue <> Scales Then
Scales = vNewValue
End If
End Property
Public Property Get PCXWidth() As Long
Select Case Scales
Case 0
PCXWidth = nWidth
Case 1
PCXWidth = TW.TwipWidth
End Select
End Property



Public Property Get PCXHeight() As Long
Select Case Scales
Case 0
PCXHeight = nHeight
Case 1
PCXHeight = TW.TwipHeight
End Select
End Property
Private Sub Class_Initialize()
Automatic = True
Scales = 1
End Sub
Public Property Get Compressed() As Boolean
Select Case Header.Encoding
Case 1
Compressed = True
Case Else
Compressed = False
End Select

End Property
Public Property Get Bpp() As Integer
Bpp = CInt(Header.Bpp)
End Property
Public Property Get IsPCX() As Boolean
Dim Test As Boolean
Test = True
    Select Case Header.Manufacturer
    Case &HA
    Case Else
    Test = False
    End Select

    Select Case Header.Encoding
    Case Is < &H2
    Case Else
    Test = False
    End Select
    
    Select Case Header.Version
    Case &H0
    Case &H2
    Case &H3
    Case &H5
    Case Else
    Test = False
    End Select
    
    IsPCX = Test
End Property
Public Property Get Autoscale() As Boolean
Autoscale = Automatic
End Property

Public Property Let Autoscale(ByVal vNewValue As Boolean)
Automatic = vNewValue
End Property

Public Function DrawPCX(PicObj As Object)
DrawBitmap nWidth, nHeight, PicObj, Automatic
End Function

