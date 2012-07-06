Attribute VB_Name = "Outside_ModFRX"
'Note: this file has been modified for use within PhotoDemon.

'IF YOU WANT TO USE THIS CODE IN YOUR OWN PROJECT, PLEASE DOWNLOAD THE ORIGINAL FROM THIS LINK:
' http://btmtz.mvps.org/gfxfromfrx/

'Many thanks to Brad Martinez for this excellent set of code.

Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' ============================================================================
' VB (and COM) recognize the following graphic files: BMP, DIB, GIF, JPG, WMF, EMF, ICO, CUR
' each containing the following proprietary image file format signatures:
' (all IMGSIG_* and IMGTERM_* constants are user-defined)

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' BMP, DIB (bitmap):

'typedef struct tagBITMAPFILEHEADER {
'    WORD      bfType;   // "BM"
'    DWORD   bfSize;    // size of file, should match FRXITEMHDR*.dwSizeImage
'    WORD      bfReserved1;
'    WORD      bfReserved2;
'    DWORD   bfOffBits;
'} BITMAPFILEHEADER;

Public Const IMGSIG_BMPDIB = &H4D42   ' "BM" ('424D') WORD @ image offset 0

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' GIF:

' First 3 bytes is "GIF", next 3 bytes is version, '87a', '89a', etc.

Public Const IMGSIG_GIF = &H464947   ' "GIF" ('4749 | 46') masked DWORD @ image offset 0
Public Const IMGTERM_GIF = &H3B      ' ";" (semicolon), WORD @ offset Len(image) - 1

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' JPG:

' SOI = Start Of Image = 'FFD8'
'   This marker must be present in any JPG file *once* at the beginning of the file.
'    (Any JPG file starts with the sequence FFD8.)
' EOI = End Of Image = 'FFD9'
'    Similar to EOI: any JPG file ends with FFD9.
' APP0 = it's the marker used to identify a JPG file which uses the JFIF specification = FFE0

' integers
Public Const IMGSIG_JPG = &HD8FF       ' ('FFD8') WORD @ offset image 0, may have APP0
Public Const IMGTERM_JPG = &HD9FF   ' ('FFD9') WORD @ offset Len(image) - 2

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' WMF, EMF:

' first try to read the DWORD enhanced metafile signature @ image offset 40 (&H28)
' (ENHMETAHEADER.dSignature member)

Public Const ENHMETA_SIGNATURE = &H464D4520   ' ('2045 | 4D46') " EMF" (in wingdi.h)

' If that fails, try to read the DWORD METAHEADER.mtSize member @ image offset 6
' (it should equal FRXITEMHDR*.dwSizeImage), and check mtHeaderSize too.

Public Type METAHEADER   ' mh
  mtType As Integer
  mtHeaderSize As Integer   ' Len(mh)
  mtVersion As Integer
  mtSize As Long   ' size of image
  mtNoObjects As Integer
  mtMaxRecord As Long
  mtNoParameters As Integer
End Type

' If that fails, read the 16bit Aldus Placeable metafile header key:

' "Q129658 SAMPLE: Reading and Writing Aldus Placeable Metafiles" or
' "Q66949: INFO: Windows Metafile Functions & Aldus Placeable Metafiles"
'typedef struct {
'    DWORD           dwKey;   // 0x9AC6CDD7
'    WORD              hmf;
'    SMALL_RECT  bbox;
'    WORD              wInch;
'    DWORD           dwReserved;
'    WORD              wCheckSum;
'} APMHEADER, *PAPMHEADER;  // APMFILEHEADER

Public Const IMGSIG_WMF_APM = &H9AC6CDD7   ' ('D7CD | C69A') DWORD @ image offset 0

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ICO, CUR:

' First check NEWHEADER.ResType, then, since there may be a discrepency in
' the cursor's CURSORDIRENTRY and CURSORDIR structs, read the NEWHEADER
' ResCount member, multiply that by Len(ICONDIRENTRY) (or 16 bytes) to find
' the BITMAPINFOHEADER, then read its biSize member, which should be
' Len(BITMAPINFOHEADER) (or 40 bytes)

Public Const RES_ICON = 1
Public Const RES_CURSOR = 2

Public Type NEWHEADER   ' was ICONDIR (ICONHEADER?)
  Reserved As Integer   ' must be 0
  ResType As Integer    ' RES_ICON or RES_CURSOR
  ResCount As Integer   ' number of images (ICONDIRENTRYs) in the file (group)
End Type

'Public Type ICONDIRENTRY
'  bWidth As Byte                ' Width, in pixels, of the image
'  bHeight As Byte               ' Height, in pixels, of the image
'  bColorCount As Byte        ' Number of colors in image (0 if >=8bpp)
'  bReserved As Byte          ' Reserved ( must be 0)
'  wPlanes As Integer           ' Color Planes
'  wBitCount As Integer        ' Bits per pixel
'  dwBytesInRes As Long    ' How many bytes in this resource?
'  dwImageOffset As Long   ' Where in the file is this image?
'End Type
'
'Public Type CURSORDIRENTRY
'  ' The new CURSORDIR struct defines the first 4 Byte members instead as: (!!??)
''  wWidth As Integer
''  wHeight As Integer
'  bWidth As Byte                ' Width, in pixels, of the image
'  bHeight As Byte               ' Height, in pixels, of the image
'  bColorCount As Byte        ' Number of colors in image (0 if >=8bpp)
'  bReserved As Byte          ' Reserved ( must be 0)
'  wXHotspot As Integer      ' x-coordinate, in pixels, of the cursor hot spot.
'  wYHotspot As Integer      ' y-coordinate, in pixels, of the cursor hot spot.
'  dwBytesInRes As Long    ' How many bytes in this resource?
'  dwImageOffset As Long   ' Where in the file is this image?
'End Type

' user-defined struct sizes
Public Const SIZEOFDIRENTRY = 16
Public Const SIZEOFBITMAPINFOHEADER = 40

' assumes that NEWHEADER.Reserved is 0 (which it's supposed to be)
Public Const IMGSIG_ICO = &H10000      ' see above, DWORD @ image offset 0
Public Const IMGSIG_CUR = &H20000    ' see above, DWORD @ image offset 0

' ============================================================================
' VB FRX/CTX/DSX/DOX/PGX binary file item header formats:

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' TexBox.Text when Multiline = True has WORD text size value

Public Type FRXITEMHDRW   ' fihw
  dwSizeText As Integer   ' size of text
End Type

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Label.Caption and VB3 frx has DWORD image/text size value

Public Type FRXITEMHDRDW   ' fihdw
  dwSizeImage As Long   ' size of image/text
End Type

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' VB intrinsic control StdPictures (other blobs?) use FRXITEMHDR

Public Type FRXITEMHDR   ' fih
  dwSizeImageEx As Long   ' = dwSizeImage + 8
  dwKey As Long                 ' &H746C "lt" ( | 6C74 | )
  dwSizeImage As Long       ' size of image (= dwSizeImageEx - 8)
End Type

' frx binary when Form.Icon is deleted in designtime:
'   0800 0000 6C74 0000 0000 0000   ....lt......  (just the FRXITEMHDR, no data)

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Comctl32.ocx, Mscomctl.ocx StdPictures (other blobs?) use FRXITEMHDREX

Public Type Guid    ' 16 bytes (128 bits)
  dwData1 As Long      ' 4 bytes
  wData2 As Integer     ' 2 bytes
  wData3 As Integer     ' 2 bytes
  abData4(7) As Byte   ' 8 bytes, zero based
End Type

Public Type FRXITEMHDREX   ' fihex, 28 bytes
  dwSizeImageEx As Long   ' = dwSizeImage + 24
  clsid As Guid                    ' CLSID_StdPicture, CLSID_?
  dwKey As Long                 ' &H746C "lt" ( | 6C74 | )
  dwSizeImage As Long       ' size of image (= dwSizeImageEx - 24)
End Type

Public Const FIH_Key = &H746C

' ============================================================================

Public Enum CBoolean   ' enum members are Long data types
  CFalse = 0
  CTrue = 1
End Enum

Public Const S_OK = 0    ' indicates successful HRESULT

'WINOLEAPI CreateStreamOnHGlobal(
'    HGLOBAL hGlobal,            // Memory handle for the stream object
'    BOOL fDeleteOnRelease,  // Whether to free memory when the object is released
'    LPSTREAM * ppstm           // Indirect pointer to the new stream object
');
Declare Function CreateStreamOnHGlobal Lib "ole32" _
                              (ByVal hGlobal As Long, _
                              ByVal fDeleteOnRelease As CBoolean, _
                              ppstm As Any) As Long

'STDAPI OleLoadPicture(
'    IStream * pStream,  // Pointer to the stream that contains picture's data
'    LONG lSize,            // Number of bytes read from the stream
'    BOOL fRunmode,   // The opposite of the initial value of the picture's property
'    REFIID riid,             // interface identifier describing the type of interface pointer to return
'    VOID ppvObj          // Indirect pointer to the object, not AddRef'd!!
');
Declare Function OleLoadPicture Lib "olepro32" _
                              (pStream As Any, _
                              ByVal lSize As Long, _
                              ByVal fRunmode As CBoolean, _
                              riid As Guid, _
                              ppvObj As Any) As Long

Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Guid) As Long
Declare Function IsEqualGUID Lib "ole32" (rguid1 As Guid, rguid2 As Guid) As Boolean

Public Const sCLSID_StdPicture = "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
Public Const sIID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"

Public Const GMEM_MOVEABLE = &H2
Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

' ============================================================================
' GetOpen/SaveFileName

' Returns the string before first null char (if any) in an ANSII string.

Public Function GetStrFromBufferA(szA As String) As String
  If InStr(szA, vbNullChar) Then
    GetStrFromBufferA = Left$(szA, InStr(szA, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would rtn a zero length string ("").
    GetStrFromBufferA = szA
  End If
End Function

' Returns the low-order word from the given 32-bit value.

Public Function LOWORD(dwValue As Long) As Integer
  MoveMemory LOWORD, dwValue, 2
End Function

Public Function PictureFromBits(abPic() As Byte) As IPicture  ' not a StdPicture!!
  Dim nLow As Long
  Dim cbMem  As Long
  Dim hMem  As Long
  Dim lpMem  As Long
  Dim IID_IPicture As Guid
  Dim istm As stdole.IUnknown '  IStream
  
  ' Get the size of the picture's bits
  On Error GoTo Out
  nLow = LBound(abPic)
  On Error GoTo 0
  cbMem = (UBound(abPic) - nLow) + 1
  
  ' Allocate a global memory object
  hMem = GlobalAlloc(GMEM_MOVEABLE, cbMem)
  If hMem Then
    
    ' Lock the memory object and get a pointer to it.
    lpMem = GlobalLock(hMem)
    If lpMem Then
      
      ' Copy the picture bits to the memory pointer and unlock the handle.
      MoveMemory ByVal lpMem, abPic(nLow), cbMem
      Call GlobalUnlock(hMem)
      
      ' Create an ISteam from the pictures bits (we can explicitly free hMem
      ' below, but we'll have the call do it here...)
      If (CreateStreamOnHGlobal(hMem, CTrue, istm) = S_OK) Then
        If (CLSIDFromString(StrPtr(sIID_IPicture), IID_IPicture) = S_OK) Then
          
          ' Create an IPicture from the IStream (the docs say the call does not
          ' AddRef its last param, but it looks like the reference counts are correct..)
          Call OleLoadPicture(ByVal ObjPtr(istm), cbMem, CFalse, IID_IPicture, PictureFromBits)
          
        End If   ' CLSIDFromString
      End If   ' CreateStreamOnHGlobal
    End If   ' lpMem
    
'    Call GlobalFree(hMem)
  End If   ' hMem
      
Out:
End Function
