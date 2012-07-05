VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadTGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Autor: ALKO
'e-mail: alfred.koppold@freenet.de

Option Explicit
'Types
Private Type ScTw
TwipWidth As Long
TwipHeight As Long
End Type

Private Type RGBTriple
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type RGBQUAD
rgbBlue As Byte
rgbGreen As Byte
rgbRed As Byte
rgbReserved As Byte
End Type

Private Type tgaheader                  'declare the header
  Info       As Byte      'length of image information block
  colortyp   As Byte      'DAC table or BGR format
  imagetyp   As Byte      'compressed or uncompressed
  origin     As Integer         'first entry in the DAC table
  colnumber  As Integer         'number of colors in the DAC table
  entrybits  As Byte      'entry size in the DAC table
  xvalue     As Integer         'x co-ordinate lower left corner
  yvalue     As Integer         'y co-ordinate lower left corner
  widt       As Integer         'image width
  Height     As Integer         'image height
  pixelsize  As Byte      'number of bits per pixel
  descriptor As Byte      'image descriptor
End Type

Private Type Palette2
Byte1 As Byte
Byte2 As Byte
End Type

Private Type Palette3
Blau As Byte
Grün As Byte
Rot As Byte
End Type

Private Type Palette4
Blau As Byte
Grün As Byte
Rot As Byte
Attribut As Byte
End Type

'Functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Variables
Private Orientation As Integer
Private Übergabe() As Byte
Private BitmapData() As Byte
Private Header As tgaheader         'define the header
Private Pal1() As Palette3
Private Pal2() As Palette4
Private Pal3() As Palette2
Private i As Long
Private nWidth As Long
Private nHeight As Long
Private Scales As Integer
Private Automatic As Boolean
Private TW As ScTw

Public Function LoadTGA(ByVal Filename As String) As StdPicture
    Dim nFreefile As Integer
    
    nFreefile = FreeFile

    Open Filename For Binary Lock Write As #nFreefile
    
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    Close #nFreefile
TW = PixelToTwips(CLng(Header.widt), CLng(Header.Height))
Select Case Header.pixelsize
Case 8
Read8Bit Filename
Case 16
Read16bit Filename
Case 24
Read24bit Filename
Case 32
Read32bit Filename
End Select
End Function




Private Sub Read8Bit(Filename As String)
    Dim nFreefile As Integer
    Dim nLineSize As Long
    Dim Palette8(0 To 255) As RGBTriple
    
    nFreefile = FreeFile

    Open Filename For Binary Lock Write As #nFreefile
    
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    
Select Case Header.entrybits / 8
    Case 2
    ReDim Pal3(Header.colnumber - 1)
    Get #nFreefile, , Pal3
    'Palette erstellen
    For i = 0 To UBound(Pal3)
    If GetByte(Pal3(i).Byte1, 1) = 1 Then Palette8(i).Blue = 16
    If GetByte(Pal3(i).Byte1, 2) = 1 Then Palette8(i).Blue = Palette8(i).Blue + 8
    If GetByte(Pal3(i).Byte1, 3) = 1 Then Palette8(i).Blue = Palette8(i).Blue + 4
    If GetByte(Pal3(i).Byte1, 4) = 1 Then Palette8(i).Blue = Palette8(i).Blue + 2
    If GetByte(Pal3(i).Byte1, 5) = 1 Then Palette8(i).Blue = Palette8(i).Blue + 1
    If GetByte(Pal3(i).Byte1, 6) = 1 Then Palette8(i).Green = 16
    If GetByte(Pal3(i).Byte1, 7) = 1 Then Palette8(i).Green = Palette8(i).Green + 8
    If GetByte(Pal3(i).Byte1, 8) = 1 Then Palette8(i).Green = Palette8(i).Green + 4
    If GetByte(Pal3(i).Byte2, 1) = 1 Then Palette8(i).Green = Palette8(i).Green + 2
    If GetByte(Pal3(i).Byte2, 2) = 1 Then Palette8(i).Green = Palette8(i).Green + 1
    If GetByte(Pal3(i).Byte2, 3) = 1 Then Palette8(i).Red = 16
    If GetByte(Pal3(i).Byte2, 4) = 1 Then Palette8(i).Red = Palette8(i).Red + 8
    If GetByte(Pal3(i).Byte2, 5) = 1 Then Palette8(i).Red = Palette8(i).Red + 4
    If GetByte(Pal3(i).Byte2, 6) = 1 Then Palette8(i).Red = Palette8(i).Red + 2
    If GetByte(Pal3(i).Byte2, 7) = 1 Then Palette8(i).Red = Palette8(i).Red + 1
Palette8(i).Red = Palette8(i).Red * 4
Palette8(i).Green = Palette8(i).Green * 4
Palette8(i).Blue = Palette8(i).Blue * 4
    Next i
    Case 3
    ReDim Pal1(Header.colnumber - 1)
    Get #nFreefile, , Pal1
    'Farbfolge ändern
    For i = 0 To UBound(Pal1)
    Palette8(i).Red = Pal1(i).Rot
    Palette8(i).Blue = Pal1(i).Blau
    Palette8(i).Green = Pal1(i).Grün
    Next i

    Case 4
    ReDim Pal2(Header.colnumber - 1)
    Get #nFreefile, , Pal2
    'Farbfolge ändern
    For i = 0 To UBound(Pal2)
    Palette8(i).Red = Pal2(i).Rot
    Palette8(i).Blue = Pal2(i).Blau
    Palette8(i).Green = Pal2(i).Grün
    Next i

End Select
    
    ReDim BitmapData(LOF(nFreefile) - Len(Header) - (UBound(Pal1) * 3))
     Get #nFreefile, , BitmapData()
        Orientation = GetByte(Header.descriptor, 3)

    Close #nFreefile
        With Header
        nWidth = .widt - .xvalue ' + 1
        nHeight = .Height - .yvalue ' + 1
        nLineSize = .widt
    End With
        ReDim Übergabe(UBound(Palette8) * 3 + 3)
        CopyMemory Übergabe(0), Palette8(0), UBound(Übergabe)
        InitColorTable_8 Übergabe
If Header.imagetyp = 9 Or Header.imagetyp = 10 Then
    BitmapData = DecompressTGA(BitmapData(), 8, nHeight, nWidth)
    End If
    MakeBitmap BitmapData, nHeight, nLineSize
    CreateBitmap_8 BitmapData, nWidth, nHeight, Orientation

End Sub


Private Sub Read16bit(Filename As String)
    Dim nFreefile As Integer
    Dim nLineSize As Long
    Dim Data() As Byte
    
    nFreefile = FreeFile

    Open Filename For Binary Lock Write As #nFreefile
    
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    
    With Header
        nWidth = .widt - .xvalue ' + 1
        nHeight = .Height - .yvalue ' + 1
        nLineSize = .widt * .pixelsize
    End With
    ReDim Data(LOF(nFreefile) - Len(Header))
    
    Get #nFreefile, , Data()
       Orientation = GetByte(Header.descriptor, 3)

    Close #nFreefile
If Header.imagetyp = 9 Or Header.imagetyp = 10 Then
    Data = DecompressTGA(Data(), 24, nHeight, nWidth)
    End If
CreateBitmap_16 Data, nWidth, nHeight, Orientation

End Sub

Private Sub Read24bit(Filename As String)
    Dim nFreefile As Integer
    Dim nLineSize As Long
    Dim Data() As Byte
    Dim quad() As RGBQUAD
    
    nFreefile = FreeFile

    Open Filename For Binary Lock Write As #nFreefile
    
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    
    With Header
        nWidth = .widt - .xvalue ' + 1
        nHeight = .Height - .yvalue ' + 1
        nLineSize = .widt * .pixelsize
    End With
    ReDim Data(LOF(nFreefile) - Len(Header))
    
    Get #nFreefile, , Data()
   Orientation = GetByte(Header.descriptor, 3)
    Close #nFreefile
If Header.imagetyp = 9 Or Header.imagetyp = 10 Then
    Data = DecompressTGA(Data(), 24, nHeight, nWidth)
    End If
    'Sort from planes into a bitmap
'==================================================
    Dim x As Long, y As Long, nStartPos As Long
    ReDim quad(UBound(Data) / 3)
    For x = 0 To UBound(Data) / 3 - 1
        
        
            With quad(x)
                .rgbBlue = Data(x * 3)
                .rgbGreen = Data(x * 3 + 1)
                .rgbRed = Data(x * 3 + 2)
            End With
        
        Next x
           ReDim BitmapData(UBound(quad) * 4 + 4)
   CopyMemory BitmapData(0), quad(0), UBound(BitmapData)

CreateBitmap_24 BitmapData, nWidth, nHeight, Orientation

End Sub

Private Sub Read32bit(Filename As String)
    Dim nFreefile As Integer
    Dim nLineSize As Long
    Dim Data() As Byte
    Dim quad() As RGBQUAD
    nFreefile = FreeFile

    Open Filename For Binary Lock Write As #nFreefile
    
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    
    With Header
        nWidth = .widt - .xvalue ' + 1
        nHeight = .Height - .yvalue ' + 1
        nLineSize = .widt * .pixelsize
    End With
    ReDim Data(LOF(nFreefile) - Len(Header))
    
    Get #nFreefile, , Data()
       Orientation = GetByte(Header.descriptor, 3)

    Close #nFreefile
If Header.imagetyp = 9 Or Header.imagetyp = 10 Then
    Data = DecompressTGA(Data(), 32, nHeight, nWidth)
    End If
    'Sort from planes into a bitmap
'==================================================
    Dim x As Long, y As Long, nStartPos As Long
    ReDim quad(UBound(Data) / 4)
    For x = 0 To UBound(Data) / 4 - 1
        
        
            With quad(x)
                .rgbBlue = Data(x * 4)
                .rgbGreen = Data(x * 4 + 1)
                .rgbRed = Data(x * 4 + 2)
            End With
        
        Next x
           ReDim BitmapData(UBound(quad) * 4 + 4)
   CopyMemory BitmapData(0), quad(0), UBound(BitmapData)

CreateBitmap_24 BitmapData, nWidth, nHeight, Orientation

End Sub


Private Function DecompressTGA(RLEStream() As Byte, Bits As Long, Höhe As Long, Breite As Long) As Byte()
    Dim InitSize As Long
    Dim Temp() As Byte
    Dim n As Long
    Dim k As Boolean
    Dim b As Long
    Dim l As Long
    Dim fertig As Long
    Dim z As Long
    Dim Länge As Long
    Dim Byteanzahl As Long
    
    Byteanzahl = Bits / 8
       InitSize = CLng(Höhe * Breite * Byteanzahl)
    ReDim Temp(0 To InitSize)
    Do While fertig < InitSize
    If l > UBound(RLEStream) Then GoTo Ende
        z = 0
        If RLEStream(l) > 127 Then
            n = RLEStream(l) - 127
            For b = 0 To n - 1
            
            CopyMemory Temp(fertig), RLEStream(l + 1), Byteanzahl
                           fertig = fertig + Byteanzahl
            
            Next b
        k = True
        Else
           n = RLEStream(l) + 1
           Länge = n * Byteanzahl

            CopyMemory Temp(fertig), RLEStream(l + 1), Länge
            
            k = False
            z = z + Länge

            fertig = fertig + z
            End If
            If k = True Then
                l = Byteanzahl + 1 + l
                n = z + (n * Byteanzahl) + 1
            Else
                l = (n * Byteanzahl + 1) + l
                n = z + n
            End If
    Loop
Ende:

    DecompressTGA = Temp
    
End Function

Private Sub MakeBitmap(ImageArray() As Byte, Lines As Long, BytesLine As Long)
Dim Übergabe() As Byte
Dim GrößeBMP As Long
Dim Zugabe As Integer
Dim Standort As Long
Dim nBitmapX As Long
Dim l As Long

    If (BytesLine) Mod 4 = 0 Then
            'wenn duch 4 teilbar dann ok
        nBitmapX = BytesLine - 1
       Else
           'ansonsten ändern
        nBitmapX = (BytesLine \ 4) * 4 + 3
    End If
If nBitmapX + 1 <> BytesLine Then
Zugabe = nBitmapX - BytesLine + 1
End If
GrößeBMP = Lines * (nBitmapX + 1) - 1
ReDim Übergabe(UBound(ImageArray))

CopyMemory Übergabe(0), ImageArray(0), UBound(ImageArray) + 1
ReDim ImageArray(GrößeBMP)
For i = 0 To BytesLine * Lines - BytesLine Step BytesLine
CopyMemory ImageArray(Standort), Übergabe(i), BytesLine
Standort = Standort + nBitmapX + 1

Next i

End Sub


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



Public Function DrawTGA(PicObj As Object)
DrawBitmap nWidth, nHeight, PicObj, Automatic
End Function

Public Property Get TGAWidth() As Long
Select Case Scales
Case 0
TGAWidth = nWidth
Case 1
TGAWidth = TW.TwipWidth
End Select
End Property



Public Property Get TGAHeight() As Long
Select Case Scales
Case 0
TGAHeight = nHeight
Case 1
TGAHeight = TW.TwipHeight
End Select
End Property

Public Property Get Compressed() As Boolean
Select Case Header.imagetyp
Case 9
Compressed = True
Case 10
Compressed = True
Case Else
Compressed = False
End Select

End Property


Public Property Get BPP() As Integer
BPP = CInt(Header.pixelsize)
End Property


Public Property Get IsTGA() As Boolean
Dim Test As Boolean
Test = True
    Select Case Header.imagetyp
    Case 1
    Case 2
    Case 3
    Case 9
    Case 10
    Case 11
    Case 32
    Case 33
    Case Else
    Test = False
    End Select

    Select Case Header.pixelsize
    Case &H8
    '8bpp
    Case &H10
    '16bpp
    Case &H18
    '24bpp
    Case &H20
    '32bpp
    Case Else
    Test = False
    End Select
    IsTGA = Test
End Property


Public Property Get ScaleMode() As Integer
ScaleMode = Scales
End Property

Public Property Let ScaleMode(ByVal vNewValue As Integer)
If vNewValue > 0 Then vNewValue = 1
If vNewValue <> Scales Then
Scales = vNewValue
End If
End Property


Public Property Get Autoscale() As Boolean
Autoscale = Automatic
End Property

Public Property Let Autoscale(ByVal vNewValue As Boolean)
Automatic = vNewValue
End Property

Private Sub Class_Initialize()
Automatic = True
Scales = 1
End Sub
Private Function PixelToTwips(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
Dim XWert1 As Long
Dim yWert1 As Long
ux = Screen.TwipsPerPixelX
PixelToTwips.TwipWidth = xwert * ux
uy = Screen.TwipsPerPixelY
PixelToTwips.TwipHeight = ywert * uy
End Function

