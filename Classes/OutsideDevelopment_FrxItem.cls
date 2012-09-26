VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFrxItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
Private Const MAX_PATH = 260

Private Type SHFILEINFO   ' sfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

' Retrieves information about an object in the file system, such as a file,
' a folder, a directory, or a drive root.
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" _
                              (ByVal pszPath As Any, _
                              ByVal dwFileAttributes As Long, _
                              psfi As SHFILEINFO, _
                              ByVal cbFileInfo As Long, _
                              ByVal uFlags As Long) As Long

' SHGetFileInfo uFlags
Private Const SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
Private Const SHGFI_TYPENAME = &H400                  ' sfi.szTypeName is filled, rtns BOOL

' ============================================================================

Public Enum PictureTypes
  ptNone = 0
  ptBMP = 1
  ptDIB = 2
  ptGIF = 3
  ptJPG = 4
  ptWMF = 5
  ptEMF = 6
  ptICO = 7
  ptCUR = 8
End Enum

Private m_abItem() As Byte
Private m_dwSizeHdr As Long
Private m_dwFileOffset As Long
Private m_dwSizeImage As Long
Private m_dwPicType As PictureTypes

Private m_sTypeName As String
Private m_sExtension As String

Friend Sub Init(abBits() As Byte, dwSizeHdr, dwOffset As Long, dwSizeImage As Long, dwPicType As PictureTypes)
  m_abItem = abBits
  m_dwSizeHdr = dwSizeHdr
  m_dwFileOffset = dwOffset
  m_dwSizeImage = dwSizeImage
  m_dwPicType = dwPicType
  Call FillPictureInfo
End Sub

Public Property Get Bits() As Variant
  Bits = m_abItem
End Property

Public Property Get HeaderSize() As Long
  HeaderSize = m_dwSizeHdr
End Property

Public Property Get FileOffset() As Long
  FileOffset = m_dwFileOffset
End Property

Public Property Get ImageSize() As Long
  ImageSize = m_dwSizeImage
End Property

Public Property Get PictureType() As PictureTypes
  PictureType = m_dwPicType
End Property

Public Property Get FileTypeName() As String
  FileTypeName = m_sTypeName
End Property

Public Property Get FileExtension() As String
  FileExtension = m_sExtension
End Property

Public Property Get Picture() As StdPicture
  If m_dwPicType Then
    
        Set Picture = New StdPicture
        Set Picture = PictureFromBits(m_abItem)

  End If
End Property

Private Sub FillPictureInfo()
  Dim sTypeName As String
  
  Select Case m_dwPicType
    Case ptBMP, ptDIB
      m_sTypeName = "Bitmap"
      m_sExtension = "bmp"
    Case ptGIF
      m_sTypeName = "GIF"
      m_sExtension = "gif"
    Case ptJPG
      m_sTypeName = "JPEG"
      m_sExtension = "jpg"
    Case ptWMF
      m_sTypeName = "WMF"
      m_sExtension = "wmf"
    Case ptEMF
      m_sTypeName = "EMF"
      m_sExtension = "emf"
    Case ptICO
      m_sTypeName = "Icon"
      m_sExtension = "ico"
    Case ptCUR
      m_sTypeName = "Cursor"
      m_sExtension = "cur"
    Case ptNone
      m_sTypeName = "Binary"
      m_sExtension = "txt"
  End Select
  
  If m_dwPicType Then
    sTypeName = GetFileTypeName("." & m_sExtension)
    If Len(sTypeName) Then m_sTypeName = sTypeName
  End If
  
End Sub

' If successful returns the specified file's typename, returns an empty string otherwise.

Private Function GetFileTypeName(sFile As String) As String
  Dim sfi As SHFILEINFO
  If SHGetFileInfo(sFile, 0, sfi, Len(sfi), SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
    GetFileTypeName = GetStrFromBufferA(sfi.szTypeName)
  End If
End Function
