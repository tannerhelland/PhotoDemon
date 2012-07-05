VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFrxFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
Public Enum FrxStructTypes
  fstUnknown = 0
  fstItemHdr = 1         ' FRXITEMHDR
  fstItemHdrEx = 2     ' FRXITEMHDREX
  fstItemHdrDW = 3  ' FRXITEMHDRDW
End Enum

Private Type FRXITEMHDRINFO
  dwStructType As FrxStructTypes
  dwSizeHdr As Long
  dwSizeImage As Long
End Type

Private m_sPath As String
Private m_cbFile As Long

Private m_colFrxItems As New Collection
'

' ============================================================================
' public members

Public Property Get Path() As String
  Path = m_sPath
End Property

Public Property Let Path(s As String)
  If LoadItems(s) Then m_sPath = s
End Property

Public Property Get FileSize() As Long
  FileSize = m_cbFile
End Property

Public Property Get Item(Index As Integer) As cFrxItem   ' proc ID = Default
Attribute Item.VB_UserMemId = 0
  Set Item = m_colFrxItems(Index)
End Property

Public Property Get Count() As Long
   Count = m_colFrxItems.Count
End Property

Public Property Get NewEnum() As IUnknown   ' Hidden, proc ID = -4
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = m_colFrxItems.[_NewEnum]
End Property

' ============================================================================
' private members

Private Function LoadItems(sPath As String) As Boolean
  Dim ff As Integer
  Dim abFile() As Byte
  
  ' First make sure that we're about to look in a VB binary file
  If (InStr(1, ".frx.ctx.dsx.dox.pgx", Right(sPath, 4), vbTextCompare) = 0) Then
    Exit Function
  End If
  
  ' Load the entire file into a byte array
  ff = FreeFile
  On Error GoTo Out
  Open sPath For Binary As ff
  m_cbFile = LOF(ff)
  On Error GoTo 0
  ReDim abFile(m_cbFile - 1)
  Get #ff, , abFile
  Close ff
  
  Dim dwFileOffset As Long
  Dim fihi As FRXITEMHDRINFO
  
  Dim abItem() As Byte
  Dim dwPicType As PictureTypes
  Dim cfi As cFrxItem
  Dim cbBlob As Long
  
  ' De-reference all existing cFrxItems
  Set m_colFrxItems = Nothing
  
  ' Keep going while file pointer is less than the file size
  Do While (dwFileOffset < m_cbFile)
    
    ' Get the item header at the current file pointer
    Call GetFrxItemHeaderInfo(abFile, m_cbFile, dwFileOffset, fihi)
    
    dwPicType = ptNone
    
    ' If a valid item header was found, load the item's bits into an array and
    ' get its picture type (if any). Validate fihi.dwSizeImage before performing
    ' math on it (which may cause an overflow if an invalid header).
    If (fihi.dwSizeImage > 0) And (fihi.dwSizeImage < m_cbFile) Then
      If ((dwFileOffset + fihi.dwSizeHdr + fihi.dwSizeImage) <= m_cbFile) Then
        ReDim abItem(fihi.dwSizeImage - 1)
        MoveMemory abItem(0), abFile(dwFileOffset + fihi.dwSizeHdr), fihi.dwSizeImage
        dwPicType = GetPictureType(abItem, fihi.dwSizeImage)
      End If
    End If
    
    ' If a valid picture's header was found... (a form icon that was deleted
    ' in designtime has a FRXITEMHDR header, but no image).
    If dwPicType Or ((fihi.dwStructType = fstItemHdr) And (fihi.dwSizeImage = 0)) Then
      
      ' If we've been reading non-picture bits, add a cFrxItem for the entire
      ' blob before adding the picture below.
      If cbBlob Then
        Call LoadBlob(VarPtr(abFile(dwFileOffset - cbBlob)), cbBlob, dwFileOffset - cbBlob)
        cbBlob = 0
      End If
      
      ' Initialize a new cFrxItem for the picture and add it to the collection
      Set cfi = New cFrxItem
      Call cfi.Init(abItem, fihi.dwSizeHdr, dwFileOffset + fihi.dwSizeHdr, fihi.dwSizeImage, dwPicType)
      m_colFrxItems.Add cfi
    
      ' Move the file pointer
      dwFileOffset = dwFileOffset + fihi.dwSizeHdr + fihi.dwSizeImage
    
    Else
      ' Probably hit some text or other difficult to extract blob, walk through
      ' it byte-by-byte while counting the number of blob bytes read.
      dwFileOffset = dwFileOffset + 1
      cbBlob = cbBlob + 1
    
    End If   ' dwPicType...
  
  Loop   ' m_cbFile > dwFileOffset)

  ' If a blob is at the end of the file, or the entire file is made up of a blob
  ' (some apps create files with the same extensions as VB binary files),
  ' add a cFrxItem for it.
  If cbBlob Then
    Call LoadBlob(VarPtr(abFile(dwFileOffset - cbBlob)), cbBlob, dwFileOffset - cbBlob)
  End If
  
Out:
  With Err
    If .Number Then Call .Raise(.Number, "cFrxFile.LoadItems", .Description)
  End With
  
End Function

Private Sub LoadBlob(lpBlob As Long, cbBlob As Long, dwFileOffset As Long)
  Dim abBlob() As Byte
  Dim cfi As New cFrxItem

  ReDim abBlob(cbBlob - 1)
  MoveMemory abBlob(0), ByVal lpBlob, cbBlob
  
  Call cfi.Init(abBlob(), 0, dwFileOffset, cbBlob, ptNone)
  m_colFrxItems.Add cfi

End Sub

Private Function GetFrxItemHeaderInfo(abFrx() As Byte, _
                                                              cbFrx As Long, _
                                                              dwFileOffset As Long, _
                                                              fihi As FRXITEMHDRINFO) As Boolean
  Dim fih As FRXITEMHDR
  Dim fihex As FRXITEMHDREX
  Dim fihdw As FRXITEMHDRDW
  ' The dwSizeImageEx struct members may overflow if they're
  ' < &H800000018 since we're subtracting from them below
  On Error GoTo Out
  
  fihi.dwStructType = fstUnknown
  
  ' FRXITEMHDR
  If (cbFrx >= (dwFileOffset + Len(fih))) Then
    MoveMemory fih, abFrx(dwFileOffset), Len(fih)
    If (fih.dwSizeImageEx - 8 = fih.dwSizeImage) And (fih.dwKey = FIH_Key) Then
      fihi.dwStructType = fstItemHdr
      fihi.dwSizeHdr = Len(fih)
      fihi.dwSizeImage = fih.dwSizeImage
    End If
  End If
  
  ' FRXITEMHDREX
  If (fihi.dwStructType = fstUnknown) And (cbFrx >= (dwFileOffset + Len(fihex))) Then
    MoveMemory fihex, abFrx(dwFileOffset), Len(fihex)
    If (fihex.dwSizeImageEx - 24 = fihex.dwSizeImage) And (fihex.dwKey = FIH_Key) Then
      fihi.dwStructType = fstItemHdrEx
      fihi.dwSizeHdr = Len(fihex)
      fihi.dwSizeImage = fihex.dwSizeImage
    End If
  End If
  
  ' FRXITEMHDRDW
  If (fihi.dwStructType = fstUnknown) And (cbFrx >= (dwFileOffset + Len(fihdw))) Then
    MoveMemory fihdw, abFrx(dwFileOffset), Len(fihdw)   ' 4
    fihi.dwStructType = fstItemHdrDW
    fihi.dwSizeHdr = Len(fihdw)
    fihi.dwSizeImage = fihdw.dwSizeImage
  End If
  
  GetFrxItemHeaderInfo = (fihi.dwStructType <> fstUnknown)

Out:
End Function

Private Function GetPictureType(abPict() As Byte, cbPic As Long) As PictureTypes
  Dim dwImgSig As Long
  
  If (cbPic > 4) Then
    MoveMemory dwImgSig, abPict(0), 4
    Select Case True
    
      ' BMP, DIB: check image signature, and that BITMAPFILEHEADER.bfSize matches cbPic
      Case LOWORD(dwImgSig) = IMGSIG_BMPDIB
        Dim bfSize As Long
        MoveMemory bfSize, abPict(2), 4
        If (bfSize = cbPic) Then GetPictureType = ptBMP
      
      ' GIF: check image signature and terminator
      Case (dwImgSig And &HFFFFFF) = IMGSIG_GIF
        If (abPict(cbPic - 1) = IMGTERM_GIF) Then GetPictureType = ptGIF
      
      ' JPG: check image signature and terminator
      Case LOWORD(dwImgSig) = IMGSIG_JPG
        Dim dwImgTerm As Integer
        MoveMemory dwImgTerm, abPict(cbPic - 2), 2
        If (dwImgTerm = IMGTERM_JPG) Then GetPictureType = ptJPG
      
      ' ICO, CUR: check image signature, then see inline below
      Case (dwImgSig = IMGSIG_ICO), (dwImgSig = IMGSIG_CUR)
        Dim nh As NEWHEADER
        Dim bihOffset As Long
        Dim biSize As Long
        
        ' Fill the NEWHEADER struct
        MoveMemory nh, abPict(0), Len(nh)
        On Error Resume Next
        ' Calculate the position of the BITMAPINFOHEADER struct
        bihOffset = Len(nh) + (nh.ResCount * SIZEOFDIRENTRY)
        ' Get the size of the BITMAPINFOHEADER struct from it's first biSize member
        MoveMemory biSize, abPict(bihOffset), 4
        On Error GoTo 0
        ' If biSize = 40 bytes, then we'll call it valid.
        If (biSize = SIZEOFBITMAPINFOHEADER) Then
          If (dwImgSig = IMGSIG_ICO) Then
            GetPictureType = ptICO
          Else   ' (dwImgSig = IMGSIG_CUR)
            GetPictureType = ptCUR
          End If
        End If
        
      ' WMF, Aldus Placeable metafile: check image signature
      Case dwImgSig = IMGSIG_WMF_APM
        GetPictureType = ptWMF
      
      ' WMF, EMF
      Case Else
        ' EMF: check image signature
        If (cbPic > &H2B) Then
          MoveMemory dwImgSig, abPict(&H28), 4
          If (dwImgSig = ENHMETA_SIGNATURE) Then
            GetPictureType = ptEMF
          Else
            
            ' WMF, no Aldus: check METAHEADER.mtHeaderSize and mtSize members
            Dim mh As METAHEADER
            MoveMemory mh, abPict(0), Len(mh)
            If (mh.mtHeaderSize = Len(mh)) And (mh.mtSize = cbPic) Then
              GetPictureType = ptWMF
            End If
          
          End If
        End If   ' (cbPic > &H2B)
        
    End Select   ' True
  End If   ' (cbPic > 4)
  
End Function
