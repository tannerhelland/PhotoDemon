VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'Name:          ClsImageControls
'Autor:         Leandro I Ascierto
'Web:           www.leandroascierto.com.ar
'Date:          10/07/2011
'Descripción:   agrega imágenes en los controles cuando se encuentran los temas de windows presente
'Nota:          Esta es algo asi como una segunda versión de esta http://www.recursosvisualbasic.com.ar/htm/ocx-componentes-activex-dll/217-modulo-de-clase-cimage-button.htm
'               Desventajas: solo muestra un icono y no para los diferentes estados y en caso de que no esten presente los temas de windows no se muestra nada
'               Ventajas: Solo vasta una clase para agregar todos los iconos de el formulario o la aplicación, acepta imagenes PNG, JPG, BMP y ICO, y varios tipos de lectura.
'-------------------------------------------------------------------------------------
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal hImagelist As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_SetIconSize Lib "comctl32.dll" (ByVal himl As Long, ByVal cx As Long, ByVal cy As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32.dll" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long

Private Const BCM_FIRST As Long = &H1600
Private Const BCM_GETIMAGELIST As Long = (BCM_FIRST + &H3)
Private Const BCM_SETIMAGELIST As Long = (BCM_FIRST + &H2)

Private Const ILC_MASK          As Long = &H1
Private Const ILC_COLOR32       As Long = &H20

Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

Public Enum IconAlign
    Icon_Left = 0
    Icon_Right = 1
    Icon_Top = 2
    Icon_Bottom = 3
    Icon_Center = 4
End Enum

Private Type BUTTON_IMAGELIST
   himl     As Long
   rc       As RECT
   uAlign   As IconAlign
End Type

Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, hbmReturn As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipGetImageBounds Lib "gdiplus.dll" (ByVal mImage As Long, ByRef mSrcRect As RECTF, ByRef mSrcUnit As Long) As Long
Private Declare Function GdipDrawImageRect Lib "gdiplus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal Format As Long, ByRef Scan0 As Any, ByRef BITMAP As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type IconHeader
    ihReserved          As Integer
    ihType              As Integer
    ihCount             As Integer
End Type

Private Type IconEntry
    ieWidth             As Byte
    ieHeight            As Byte
    ieColorCount        As Byte
    ieReserved          As Byte
    iePlanes            As Integer
    ieBitCount          As Integer
    ieBytesInRes        As Long
    ieImageOffset       As Long
End Type

Private Const InterpolationModeHighQuality  As Long = &H2
Private Const IconVersion                   As Long = &H30000
Private Const PixelFormat32bppARGB          As Long = &H26200A
Private Const UnitPixel                     As Long = &H2&
Private Const OBJ_BITMAP                    As Long = 7

Private cColl   As Collection
Private BI      As BUTTON_IMAGELIST

Public Function SetMargins(hWnd As Long, Optional ByVal Left As Long, Optional ByVal Top As Long, Optional ByVal Right As Long, Optional ByVal Bottom As Long)
    SendMessage hWnd, BCM_GETIMAGELIST, 0&, BI
    SetRect BI.rc, Left, Top, Right, Bottom
    SendMessage hWnd, BCM_SETIMAGELIST, 0&, BI
End Function

Public Property Get Align(hWnd As Long) As IconAlign
    SendMessage hWnd, BCM_GETIMAGELIST, 0&, BI
    Align = BI.uAlign
End Property

Public Property Let Align(hWnd As Long, uAlign As IconAlign)
    SendMessage hWnd, BCM_GETIMAGELIST, 0&, BI
    BI.uAlign = uAlign
    SendMessage hWnd, BCM_SETIMAGELIST, 0&, BI
End Property

Public Function RemoveImage(hWnd As Long) As Boolean
    On Error Resume Next
   
    SendMessage hWnd, BCM_GETIMAGELIST, 0&, BI
    If BI.himl Then
        cColl.Remove Str(BI.himl)
        ImageList_Destroy (BI.himl): BI.himl = 0
        RemoveImage = SendMessage(hWnd, BCM_SETIMAGELIST, 0&, BI)
    End If
End Function

Private Sub Class_Initialize()
    Set cColl = New Collection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Dim i As Long
    For i = cColl.Count To 1 Step -1
        ImageList_Destroy cColl(i)
    Next
End Sub

'------------------------------------------------------------------------
'pone una Imágen GDI standar desde su handle al boton
'------------------------------------------------------------------------
Public Function SetImageFromHandle(hWnd As Long, hImage As Long, Optional ByVal cx As Long = 16, Optional ByVal cy As Long = 16) As Boolean
    On Error GoTo SetImageFromHandle_Error
    
    If hImage = 0 Then Exit Function

    If GetObjectType(hImage) = OBJ_BITMAP Then
        Dim GDIsi       As GdiplusStartupInput
        Dim gToken      As Long
        Dim hBitmap      As Long
        
        GDIsi.GdiplusVersion = 1&
        If GdiplusStartup(gToken, GDIsi) = 0 Then
            If GdipCreateBitmapFromHBITMAP(hImage, 0&, hBitmap) = 0 Then
                SetImageFromHandle = SetGdiPlusImagen(hWnd, hBitmap, cx, cy)
                Call GdipDisposeImage(hBitmap)
            End If
            GdiplusShutdown gToken: gToken = 0
        End If
        Exit Function
    End If
    
    SendMessage hWnd, BCM_GETIMAGELIST, 0&, BI
     
    If BI.himl Then
        ImageList_SetIconSize BI.himl, cx, cy
        SetImageFromHandle = ImageList_AddIcon(BI.himl, hImage) <> -1
    Else
        BI.himl = ImageList_Create(cx, cy, ILC_COLOR32 Or ILC_MASK, 1, 1)
        If BI.himl Then
            cColl.Add BI.himl, Str(BI.himl)
            If ImageList_AddIcon(BI.himl, hImage) <> -1 Then
                SetImageFromHandle = SendMessage(hWnd, BCM_SETIMAGELIST, 0&, BI)
            End If
        End If
    End If
    
    Exit Function
    
SetImageFromHandle_Error:
    If gToken Then GdiplusShutdown gToken
End Function

'------------------------------------------------------------------------
'pone una imagen desde archivo(Ico,Png,jpg,bmp, etc.)  al boton
'------------------------------------------------------------------------
Public Function LoadImageFromFile(hWnd As Long, ByVal FileName As String, Optional ByVal cx As Long = 16, Optional ByVal cy As Long = 16) As Boolean

    On Local Error GoTo LoadImageFromFile_Error
    
    Dim bvData() As Byte
    
    'Change by Tanner: support Unicode paths via pdFSO
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.LoadFileAsByteArray(FileName, bvData) Then
        LoadImageFromFile = LoadImageFromStream(hWnd, bvData, cx, cy)
    Else
        LoadImageFromFile = False
    End If
    
LoadImageFromFile_Error:
    LoadImageFromFile = False
    
End Function

'------------------------------------------------------------------------
'pone una imagen desde recurso(Ico,Png,jpg,bmp, etc.) al boton
'------------------------------------------------------------------------
Public Function LoadImageFromRes(hWnd As Long, ByVal ResIndex As Variant, ByVal ResSection As Variant, Optional ByVal cx As Long = 16, Optional ByVal cy As Long = 16) As Boolean
    On Local Error GoTo LoadImageFromRes_Error
    
    Dim bvData()    As Byte

    bvData = LoadResData(ResIndex, ResSection)

    LoadImageFromRes = LoadImageFromStream(hWnd, bvData, cx, cy)

LoadImageFromRes_Error:

End Function

'------------------------------------------------------------------------
'pone una imagen desde un array de bits(Ico,Png,jpg,bmp, etc.) al boton
'------------------------------------------------------------------------
Public Function LoadImageFromStream(hWnd As Long, ByRef bvData() As Byte, Optional ByVal cx As Long = 16, Optional ByVal cy As Long = 16) As Boolean
    
    On Local Error GoTo LoadImageFromStream_Error
    
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    If bvData(2) = vbResIcon Or bvData(2) = vbResCursor Then
        Dim hIcon           As Long
        Dim tIconHeader     As IconHeader
        Dim tIconEntry()    As IconEntry
        Dim MaxBitCount     As Long
        Dim MaxSize         As Long
        Dim Aproximate      As Long
        Dim IconID          As Long
        Dim i               As Long
    
        Call CopyMemory(tIconHeader, bvData(0), Len(tIconHeader))
        If tIconHeader.ihCount >= 1 Then
        
            ReDim tIconEntry(tIconHeader.ihCount - 1)
            Call CopyMemory(tIconEntry(0), bvData(Len(tIconHeader)), Len(tIconEntry(0)) * tIconHeader.ihCount)
            IconID = -1
               
            For i = 0 To tIconHeader.ihCount - 1
                If tIconEntry(i).ieBitCount > MaxBitCount Then MaxBitCount = tIconEntry(i).ieBitCount
            Next
    
            For i = 0 To tIconHeader.ihCount - 1
                If MaxBitCount = tIconEntry(i).ieBitCount Then
                    MaxSize = CLng(tIconEntry(i).ieWidth) + CLng(tIconEntry(i).ieHeight)
                    If MaxSize > Aproximate And MaxSize <= (cx + cy) Then
                        Aproximate = MaxSize
                        IconID = i
                    End If
                End If
            Next
                       
            If IconID = -1 Then Exit Function
           
            With tIconEntry(IconID)
                hIcon = CreateIconFromResourceEx(bvData(.ieImageOffset), .ieBytesInRes, 1, IconVersion, cx, cy, &H0)
                If hIcon <> 0 Then
                    LoadImageFromStream = SetImageFromHandle(hWnd, hIcon, cx, cy)
                    DestroyIcon hIcon
                End If
            End With
           
        End If
    
    Else
    
        Dim IStream     As IUnknown
        Dim GDIsi       As GdiplusStartupInput
        Dim hBitmap     As Long
        Dim gToken      As Long
        
        Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
       
        If Not IStream Is Nothing Then
            GDIsi.GdiplusVersion = 1&
            If GdiplusStartup(gToken, GDIsi) = 0 Then
                If GdipLoadImageFromStream(IStream, hBitmap) = 0 Then
                    LoadImageFromStream = SetGdiPlusImagen(hWnd, hBitmap, cx, cy)
                    Call GdipDisposeImage(hBitmap)
                End If
                GdiplusShutdown gToken: gToken = 0
            End If
        End If
    
        Set IStream = Nothing
    
    End If
    Exit Function
LoadImageFromStream_Error:
    If gToken Then GdiplusShutdown gToken
    
End Function

'------------------------------------------------------------------------
'pone una Imágen desde un hBitmap de GDI+ al boton
'------------------------------------------------------------------------
Public Function SetGdiPlusImagen(hWnd As Long, hBitmap As Long, ByVal cx As Long, ByVal cy As Long) As Boolean
    Dim tR          As RECTF
    Dim ResizeBmp   As Long
    Dim ResizeGra   As Long
    Dim hIcon       As Long
    
    If hBitmap = 0 Then Exit Function
    
    Call GdipGetImageBounds(hBitmap, tR, UnitPixel)
    
    If cx <> tR.Width Or cy <> tR.Height Then
        If GdipCreateBitmapFromScan0(cx, cy, 0&, PixelFormat32bppARGB, ByVal 0&, ResizeBmp) = 0 Then
            If GdipGetImageGraphicsContext(ResizeBmp, ResizeGra) = 0 Then
                GdipSetInterpolationMode ResizeGra, InterpolationModeHighQuality
                If GdipDrawImageRect(ResizeGra, hBitmap, 0, 0, cx, cy) = 0 Then
                    If GdipCreateHICONFromBitmap(ResizeBmp, hIcon) = 0 Then
                        SetGdiPlusImagen = SetImageFromHandle(hWnd, hIcon, cx, cy)
                        DestroyIcon hIcon
                    End If
                 End If
                Call GdipDeleteGraphics(ResizeGra)
            End If
            Call GdipDisposeImage(ResizeBmp)
        End If
    Else
       If GdipCreateHICONFromBitmap(hBitmap, hIcon) = 0 Then
            SetGdiPlusImagen = SetImageFromHandle(hWnd, hIcon, cx, cy)
            DestroyIcon hIcon
       End If
    End If
    
End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function


