VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This code was originally written by Leandro I. Ascierto.
' Many thanks to Leandro for providing a way to apply PNG images to program menus in VB6
'
'NOTE: this file has been modified for use within PhotoDemon.  Modifications include bugfixes relating to use under non-standard
' Windows themes (e.g. "Classic Theme" on Win 7 and earlier), a full conversion of the subclassing system to an IDE-stable alternative,
' and rewriting the class so that persistent DCs are not required for each menu image.  (Instead, we generate them on-the-fly.)
'
'These changes rely on a number of internal PhotoDemon checks, variables, and objects, so I do not recommend using PD's
' version of the class in your own projects.  You may download the original version of this code from the following link
' (good as of June 2012): http://leandroascierto.com/blog/clsmenuimage/

Option Explicit

' ---------------------------------------------------
' Autor:            Leandro I. Ascierto
' Date:             17 de Julio de 2010
' Web:              www.leandroascierto.com.ar
' Requirements:     Windows XP or Later
' History:          17/07/2010 First
' ---------------------------------------------------


'Tanner's addition 20 September '12:
' We must check if theming is enabled. On Vista/7, use of the "Classic Theme" disables theming and screws up menu icon placement.
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As Long) As Long

' ------------
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, ByRef lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuInfo Lib "user32" (ByVal hMenu As Long, ByRef LPMENUINFO As MENUINFO) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, ByRef LPCMENUINFO As MENUINFO) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
'Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' ------------------------------------------------------- GDI Plus -------------------------------------------------------------------
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef Image As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef Token As Long, ByRef lpInput As GdiplusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, BITMAP As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal mHicon As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal Image As Long, ByRef PixelFormat As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal BITMAP As Long, ByRef RECT As RECT, ByVal Flags As ImageLockMode, ByVal PixelFormat As Long, ByRef LockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal BITMAP As Long, ByRef LockedBitmapData As BitmapData) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBmp As Long, ByVal hPal As Long, ByRef BITMAP As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, MatrixGray As Any, ByVal Flags As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    RhbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
    hbmpItem As Long
End Type

Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    ItemState As Long
    hWndItem As Long
    hDC As Long
    rcItem As RECT
    itemData As Long
End Type

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Type IconHeader
    ihReserved As Integer
    ihType As Integer
    ihCount As Integer
End Type

Private Type IconEntry
    ieWidth As Byte
    ieHeight As Byte
    ieColorCount As Byte
    ieReserved As Byte
    iePlanes As Integer
    ieBitCount As Integer
    ieBytesInRes As Long
    ieImageOffset As Long
End Type

Private Type ARGB
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Private Type BitmapData
    Width As Long
    Height As Long
    Stride As Long
    PixelFormat As Long
    Scan0 As Long
    Reserved As Long
End Type

Private Type BITMAPINFOHEADER
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

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As ARGB
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type COLORMATRIX
    m(0 To 4, 0 To 4) As Single
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type


Private Type MemoDIB
    hDIB As Long
    Ptr As Long
End Type

Private Enum ImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum

Private Enum ColorAdjustType
    ColorAdjustTypeDefault = 0
    ColorAdjustTypeBitmap = 1
    ColorAdjustTypeBrush = 2
    ColorAdjustTypePen = 3
    ColorAdjustTypeText = 4
    ColorAdjustTypeCount = 5
    ColorAdjustTypeAny = 6
End Enum

Private Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = 0
    ColorMatrixFlagsSkipGrays = 1
    ColorMatrixFlagsAltGray = 2
End Enum

Private Const PixelFormat32bppARGB As Long = &H26200A
Private Const PixelFormat32bppRGB As Long = &H22009

Private Const GdiplusVersion As Long = 1&
Private Const IconVersion As Long = &H30000
    
Private Const DIB_RGB_COLORS As Long = 0&

Private Const LR_LOADFROMFILE As Long = &H10
Private Const IMAGE_ICON As Long = 1

Private Const MIIM_ID As Long = &H2
Private Const MIIM_DATA As Long = &H20
Private Const MIIM_BITMAP As Long = &H80
    
Private Const MIM_APPLYTOSUBMENUS = &H80000000
Private Const MIM_STYLE As Long = &H10

Private Const ODT_MENU As Long = 1
Private Const ODS_GRAYED As Long = &H2
Private Const ODS_CHECKED As Long = &H8
Private Const MNS_CHECKORBMP As Long = &H4000000
Private Const MNS_NOCHECK As Long = &H80000000
Private Const HBMMENU_CALLBACK As Long = -1


Private Const NULL_BRUSH As Long = 5
Private Const COLOR_GRAYTEXT As Long = 17
Private Const COLOR_APPWORKSPACE As Long = 12
Private Const SM_CXMENUCHECK As Long = 71

Private Const WM_DESTROY As Long = &H2
Private Const WM_MEASUREITEM As Long = &H2C
Private Const WM_DRAWITEM As Long = &H2B
Private Const WM_MENUSELECT As Long = &H11F
Private Const WM_ENTERIDLE As Long = &H121

Public Event MenuMessages(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Private m_hWnd As Long

Private GdipToken As Long
Private m_lWidth As Long
Private m_lHeight As Long
Private mDIB() As MemoDIB
Private IsWinVistaOrLater As Boolean
Private cColl As Collection

Private classicThemeWorkaround As Boolean

'Tanner's addition: convert this class to ise IDE-safe subclassing
Private cSubclass As cSelfSubHookCallback

Public Property Get IsWindowVistaOrLater() As Boolean
    IsWindowVistaOrLater = IsWinVistaOrLater
End Property

Public Property Get CanWeTheme() As Boolean
    CanWeTheme = Not classicThemeWorkaround
End Property

Public Property Get ImageCount() As Long
    ImageCount = UBound(mDIB) - 1
End Property


Public Function RemoveImage(ByVal Index As Long) As Boolean
    
    If Index < 0 Or Index > Me.ImageCount Then Exit Function
    
    Call DeleteObject(mDIB(Index).hDIB)
        
    Dim i As Long
    For i = Index To Me.ImageCount
        mDIB(i) = mDIB(i + 1)
    Next

    ReDim Preserve mDIB(Me.ImageCount)
    
    RemoveImage = True
    
End Function

Public Function PutImageToVBMenu(ByVal imageID As Long, ByVal MenuPos As Long, ParamArray vSubMenuPos() As Variant) As Boolean
    
    On Error Resume Next
    
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim MII As MENUITEMINFO
    Dim v As Variant
    Dim sKey As String

    hMenu = GetMenu(m_hWnd)
    
    Dim tmpInfo As MENUINFO
    
    If classicThemeWorkaround Then
        tmpInfo.cbSize = Len(tmpInfo)
        tmpInfo.fMask = MIM_APPLYTOSUBMENUS Or MIM_STYLE
        tmpInfo.dwStyle = MNS_CHECKORBMP
        SetMenuInfo hMenu, tmpInfo
    End If
    
    hSubMenu = hMenu
    
    For Each v In vSubMenuPos
        hSubMenu = GetSubMenu(hSubMenu, v)
    Next
    
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_ID
    End With
    
    If GetMenuItemInfo(hSubMenu, MenuPos, True, MII) = 0 Then Exit Function
        
    sKey = hSubMenu & "-" & MII.wID

    With MII
        '.cbSize = Len(MII)
        .fMask = MIIM_BITMAP 'Or MIIM_DATA
        
        If imageID = -1 Then
            .hbmpItem = 0
            If KeyExists(sKey) Then Call cColl.Remove(sKey)
        Else
            If IsWinVistaOrLater Then
                .hbmpItem = mDIB(imageID).hDIB
            Else
                .hbmpItem = HBMMENU_CALLBACK
                If KeyExists(sKey) Then Call cColl.Remove(sKey)
                cColl.Add imageID, sKey
            End If
        End If
        
        '.dwItemData = ImageID
        
    End With
    
    PutImageToVBMenu = SetMenuItemInfo(hSubMenu, MenuPos, True, MII)
   
    If hSubMenu = hMenu Then DrawMenuBar m_hWnd
    
End Function

Public Function PutImageToApiMenu(ByVal imageID As Long, ByVal hMenu As Long, ByVal MenuPos As Long, Optional ByVal itemData As Long) As Boolean

    Dim MII As MENUITEMINFO
    Dim sKey As String

    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_ID
    End With
    
    If GetMenuItemInfo(hMenu, MenuPos, True, MII) = 0 Then Exit Function
        
    sKey = hMenu & "-" & MII.wID

    With MII
        .fMask = MIIM_BITMAP Or MIIM_DATA
        
        If imageID = -1 Then
            .hbmpItem = 0
            If KeyExists(sKey) Then Call cColl.Remove(sKey)
        Else
            If IsWinVistaOrLater Then
                .hbmpItem = mDIB(imageID).hDIB
            Else
                .hbmpItem = HBMMENU_CALLBACK
                If KeyExists(sKey) Then Call cColl.Remove(sKey)
                cColl.Add imageID, sKey
            End If
        End If
        .dwItemData = itemData
    End With
    
    PutImageToApiMenu = SetMenuItemInfo(hMenu, MenuPos, True, MII)

End Function

Public Sub RemoveMenuCheckApi(ByVal hMenu As Long)
    Dim MI As MENUINFO
    
    With MI
        .cbSize = Len(MI)
        .fMask = MIM_STYLE
        .dwStyle = MNS_NOCHECK
    End With

    SetMenuInfo hMenu, MI
End Sub

Public Sub RemoveMenuCheckVB(ParamArray vSubMenuPos() As Variant)
    Dim MI As MENUINFO
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim v As Variant

    hMenu = GetMenu(m_hWnd)
    
    hSubMenu = hMenu
    
    For Each v In vSubMenuPos
        hSubMenu = GetSubMenu(hSubMenu, v)
    Next
    
    With MI
        .cbSize = Len(MI)
        .fMask = MIM_STYLE
        .dwStyle = MNS_NOCHECK
    End With

    SetMenuInfo hSubMenu, MI
End Sub


Private Sub DrawCheck(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long, bDisabled As Boolean)
    Dim hPen As Long, oldPen As Long
    Dim hBrush As Long, oldBrush As Long
    
    hPen = CreatePen(0, 1, GetSysColor(IIf(bDisabled, COLOR_GRAYTEXT, COLOR_APPWORKSPACE)))
    hBrush = GetStockObject(NULL_BRUSH)
    
    oldPen = SelectObject(hDC, hPen)
    oldBrush = SelectObject(hDC, hBrush)
    
    Rectangle hDC, x, y, x + x2, y + y2
    
    DeleteObject SelectObject(hDC, oldPen)
    Call SelectObject(hDC, oldBrush)
    
End Sub

Private Function CreateNewDib() As Long

    'Prepare a bmp header for the DIB
    Dim tBITMAPINFO As BITMAPINFO
    With tBITMAPINFO.bmiHeader
        .biSize = Len(tBITMAPINFO.bmiHeader)
        .biBitCount = 32
        .biHeight = m_lWidth
        .biWidth = m_lHeight
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * 4&
    End With
    
    Dim Index As Long
    Index = UBound(mDIB)
    
    With mDIB(Index)
    
        'Create a temporary memory DC, because CreateDIBSection requires a DC for palette matching
        Dim tmpDC As Long
        tmpDC = GetDC(0&)
        .hDIB = CreateDIBSection(tmpDC, tBITMAPINFO, DIB_RGB_COLORS, .Ptr, 0&, 0&)
        ReleaseDC 0&, tmpDC
        
        If .hDIB <> 0 Then
            CreateNewDib = Index
        Else
            CreateNewDib = -1
            Exit Function
        End If
        
    End With

    ReDim Preserve mDIB(Index + 1)

End Function

Public Function Clear()

    Dim i As Long
    
    For i = 0 To UBound(mDIB) - 1
        Call DeleteObject(mDIB(i).hDIB)
    Next
    
    ReDim mDIB(0)
    
End Function

Private Sub DrawDIB(ByVal DestHdc As Long, ByVal x As Long, ByVal y As Long, ByVal Index As Long, Disabled As Boolean)

    Dim hGraphics As Long
    Dim hImage As Long
    Dim hAttributes As Long
    Dim tMatrixColor As COLORMATRIX
    Dim tMatrixGray As COLORMATRIX
    
    If Index < 0 Or Index > Me.ImageCount Then Exit Sub

    If GdipCreateBitmapFromScan0(m_lWidth, m_lHeight, m_lWidth * 4&, PixelFormat32bppARGB, ByVal mDIB(Index).Ptr, hImage) = 0 Then

        If GdipCreateFromHDC(DestHdc, hGraphics) = 0 Then
        
            GdipImageRotateFlip hImage, &H6
            
            If Disabled Then
            
                GdipCreateImageAttributes hAttributes
                
                With tMatrixColor
                    .m(0, 0) = 0.299
                    .m(1, 0) = .m(0, 0)
                    .m(2, 0) = .m(0, 0)
                    .m(0, 1) = 0.587
                    .m(1, 1) = .m(0, 1)
                    .m(2, 1) = .m(0, 1)
                    .m(0, 2) = 0.114
                    .m(1, 2) = .m(0, 2)
                    .m(2, 2) = .m(0, 2)
                    .m(3, 3) = 0.5
                    .m(4, 4) = 1
                End With
                
                GdipSetImageAttributesColorMatrix hAttributes, ColorAdjustTypeDefault, True, tMatrixColor, tMatrixGray, ColorMatrixFlagsDefault
    
            End If
            
            GdipDrawImageRectRectI hGraphics, hImage, x, y, m_lWidth, m_lHeight, 0, 0, m_lWidth, m_lHeight, &H2, hAttributes, 0&, 0&
            
            If hAttributes Then Call GdipDisposeImageAttributes(hAttributes)
    
            GdipDeleteGraphics hGraphics
        
        End If
        
        GdipDisposeImage hImage
    
    End If
    
End Sub


Private Function AlphaIconToBmp(ByVal IconHandle As Long, ByRef RefBmp As Long) As Long
    
    Dim tRECT As RECT
    Dim tICONINFO As ICONINFO
    Dim tBitmapData As BitmapData
    Dim lPixelFormat As Long
    Dim sngWidth As Single
    Dim sngHeight As Single

    If GetIconInfo(IconHandle, tICONINFO) <> 0 Then

        If GdipCreateBitmapFromHBITMAP(tICONINFO.hbmColor, 0&, RefBmp) = 0 Then

            If GdipGetImagePixelFormat(RefBmp, lPixelFormat) = 0 Then

                If GdipGetImageDimension(RefBmp, sngWidth, sngHeight) = 0 Then

                    With tRECT
                        .Right = CLng(sngWidth)
                        .Bottom = CLng(sngHeight)
                    End With

                    If GdipBitmapLockBits(RefBmp, tRECT, ImageLockModeRead, lPixelFormat, tBitmapData) = 0 Then
                        
                        Call GdipCreateBitmapFromScan0(tRECT.Right, tRECT.Bottom, tBitmapData.Stride, PixelFormat32bppARGB, ByVal tBitmapData.Scan0, AlphaIconToBmp)

                        Call GdipBitmapUnlockBits(RefBmp, tBitmapData)
                            
                    End If
                    
                End If
                
            End If
            
        End If

        Call DeleteObject(tICONINFO.hbmMask)
        Call DeleteObject(tICONINFO.hbmColor)
        
    End If
    
End Function


Public Function AddIconFromHandle(ByVal hIcon As Long, Optional bGhosted As Boolean) As Boolean

    Dim hBmp As Long
    Dim hImage As Long
    
    On Local Error GoTo AddIconFromHandle_Error

        
    If hIcon <> 0 Then
        If IsAlphaIcon(hIcon) Then
           hImage = AlphaIconToBmp(hIcon, hBmp)
        Else
            GdipCreateBitmapFromHICON hIcon, hImage
        End If
    End If

    AddIconFromHandle = pvAddImagen(hImage, bGhosted)
    
    If hBmp <> 0 Then GdipDisposeImage hBmp

AddIconFromHandle_Error:
    
End Function


Public Function AddImageFromFile(ByVal FileName As String, Optional bGhosted As Boolean) As Boolean

    On Local Error GoTo AddImageFromFile_Error
    
    Dim hIcon As Long
    Dim FileType As String
    Dim hBmp As Long
    Dim hImage As Long
    
    FileType = UCase(Right(FileName, 3))
    
    If FileType = "ICO" Or FileType = "CUR" Then

        hIcon = LoadImage(App.hInstance, FileName, IMAGE_ICON, m_lWidth, m_lHeight, LR_LOADFROMFILE)
        
        If hIcon <> 0 Then
            If IsAlphaIcon(hIcon) Then
               hImage = AlphaIconToBmp(hIcon, hBmp)
            Else
                GdipCreateBitmapFromHICON hIcon, hImage
            End If
            
            DestroyIcon hIcon
        End If
        
    Else
    
        Call GdipLoadImageFromFile(StrPtr(FileName), hImage)
        
    End If
    
    AddImageFromFile = pvAddImagen(hImage, bGhosted)
    
    If hBmp <> 0 Then GdipDisposeImage hBmp

AddImageFromFile_Error:
    
End Function


Public Function AddImageFromStream(ByRef bvData() As Byte, Optional bGhosted As Boolean) As Boolean
    
    On Local Error GoTo AddImageFromStream_Error
    
    Dim hImage As Long
    Dim hIcon As Long
    Dim hBmp As Long

    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function

    If bvData(2) = vbResIcon Or bvData(2) = vbResCursor Then
    
        hIcon = LoadIconFromStream(bvData)
        
        If hIcon = 0 Then Exit Function
        
        If IsAlphaIcon(hIcon) Then
           hImage = AlphaIconToBmp(hIcon, hBmp)
        Else
            GdipCreateBitmapFromHICON hIcon, hImage
        End If
        
        DestroyIcon hIcon
    Else
        
        LoadImageFromStream bvData, hImage
    
    End If
    
    AddImageFromStream = pvAddImagen(hImage, bGhosted)
    
    If hBmp <> 0 Then GdipDisposeImage hBmp


AddImageFromStream_Error:

End Function

Private Function pvAddImagen(ByVal hImage As Long, Optional bGhosted As Boolean) As Boolean
    Dim hGraphics As Long
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim Index As Long
    Dim OldhDib As Long

    
    If hImage <> 0 Then
    
        Index = CreateNewDib()
        
        If Index <> -1 Then
            
            Dim tmpDC As Long
            tmpDC = Drawing.GetMemoryDC()
            
            OldhDib = SelectObject(tmpDC, mDIB(Index).hDIB)
        
            GdipCreateFromHDC tmpDC, hGraphics
        
            GdipGetImageDimension hImage, imgWidth, imgHeight
            
            If bGhosted Then
                Dim tMatrixColor As COLORMATRIX
                Dim tMatrixGray As COLORMATRIX
                Dim hAttributes As Long
            
                GdipCreateImageAttributes hAttributes
                
                With tMatrixColor
                    .m(0, 0) = 1
                    .m(1, 1) = 1
                    .m(2, 2) = 1
                    .m(3, 3) = 0.7
                    .m(4, 4) = 1

                End With
                
                GdipSetImageAttributesColorMatrix hAttributes, ColorAdjustTypeDefault, True, tMatrixColor, tMatrixGray, ColorMatrixFlagsDefault
            End If

            GdipDrawImageRectRectI hGraphics, hImage, 0, 0, m_lWidth, m_lHeight, 0, 0, imgWidth, imgHeight, &H2, hAttributes, 0&, 0&
            
            If hAttributes Then Call GdipDisposeImageAttributes(hAttributes)
    
            GdipDisposeImage hImage
    
            GdipDeleteGraphics hGraphics
            
            Call SelectObject(tmpDC, OldhDib)
            Drawing.FreeMemoryDC tmpDC
            
            pvAddImagen = True
        
        End If

    End If
End Function
 
Private Function LoadImageFromStream(ByRef bvData() As Byte, ByRef hImage As Long) As Boolean

    On Local Error GoTo LoadImageFromStream_Error
    
    Dim IStream As IUnknown
    
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
   
    If Not IStream Is Nothing Then
        
        If GdipLoadImageFromStream(IStream, hImage) = 0 Then
            
            LoadImageFromStream = True

        End If
            
    End If

    Set IStream = Nothing
    
LoadImageFromStream_Error:

End Function

Private Function LoadIconFromStream(ByRef bytIcoData() As Byte) As Long

    On Local Error GoTo LoadIconFromStream_Error
    
    Dim tIconHeader As IconHeader
    Dim tIconEntry() As IconEntry
    Dim MaxBitCount As Long
    Dim MaxSize As Long
    Dim Aproximate As Long
    Dim IconID As Long
    Dim hIcon As Long
    Dim i As Long
  
    
    Call CopyMemory(tIconHeader, bytIcoData(0), Len(tIconHeader))

    If tIconHeader.ihCount >= 1 Then
    
        ReDim tIconEntry(tIconHeader.ihCount - 1)
        
        Call CopyMemory(tIconEntry(0), bytIcoData(Len(tIconHeader)), Len(tIconEntry(0)) * tIconHeader.ihCount)
        
        IconID = -1
           
        For i = 0 To tIconHeader.ihCount - 1
            If tIconEntry(i).ieBitCount > MaxBitCount Then MaxBitCount = tIconEntry(i).ieBitCount
        Next

       
        For i = 0 To tIconHeader.ihCount - 1
            If MaxBitCount = tIconEntry(i).ieBitCount Then
                MaxSize = CLng(tIconEntry(i).ieWidth) + CLng(tIconEntry(i).ieHeight)
                If MaxSize > Aproximate And MaxSize <= (m_lWidth + m_lHeight) Then
                    Aproximate = MaxSize
                    IconID = i
                End If
            End If
        Next
                   
        If IconID = -1 Then Exit Function
       
        With tIconEntry(IconID)
            hIcon = CreateIconFromResourceEx(bytIcoData(.ieImageOffset), .ieBytesInRes, 1, IconVersion, m_lWidth, m_lHeight, &H0)
            If hIcon <> 0 Then
                LoadIconFromStream = hIcon
            End If
        End With
       
    End If

LoadIconFromStream_Error:

End Function

Public Function Init(ByVal hWnd As Long, imgWidth As Long, ByVal imgHeight As Long, Optional ByVal bRaiseEvent As Boolean = False) As Boolean

    If GdipToken Then
    
        m_lWidth = imgWidth
        m_lHeight = imgHeight
        m_hWnd = hWnd
        
        If UBound(mDIB) > 0 Then
           Me.Clear
        End If
        
        If Not IsWinVistaOrLater Or bRaiseEvent Then
            If hWnd <> 0 Then
                Call StopSubclassing
                Init = SetSubclassing(hWnd)
            End If
        Else
            Init = True
        End If
    
    End If
    
End Function


Private Function SetSubclassing(ByVal hWnd As Long) As Boolean

    If cSubclass.ssc_Subclass(hWnd, , , Me) Then
        cSubclass.ssc_AddMsg hWnd, MSG_AFTER, WM_MENUSELECT, WM_ENTERIDLE, WM_MEASUREITEM, WM_DRAWITEM, WM_DESTROY
        SetSubclassing = True
    Else
        SetSubclassing = False
    End If

End Function

Private Function StopSubclassing() As Boolean

    If m_hWnd <> 0 Then
        cSubclass.ssc_UnSubclass m_hWnd <> 0
        StopSubclassing = True
    Else
        StopSubclassing = False
    End If

End Function

Private Sub Class_Initialize()
    
    Set cColl = New Collection
    
    Dim tOSVI As OSVERSIONINFO
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    Call GetVersionEx(tOSVI)
    
    If tOSVI.dwMajorVersion < 5 Then Exit Sub
    IsWinVistaOrLater = (tOSVI.dwMajorVersion >= 6)
    
    'Tanner's addition 20 September '12:
    ' Double-check that theming is enabled. If it isn't, fall back to subclassing (e.g. Windows XP style)
    Dim hTheme As Long
    Dim sClass As String
    sClass = "Window"
    hTheme = OpenThemeData(FormMain.hWnd, StrPtr(sClass))
    If hTheme = 0 Then
        classicThemeWorkaround = True
    Else
        CloseThemeData hTheme
    End If
    
    ReDim mDIB(0)

    InitGDI

    'Tanner's fix: initialize an IDE-safe subclasser
    Set cSubclass = New cSelfSubHookCallback
    
End Sub


Private Sub Class_Terminate()

    cSubclass.ssc_Terminate
    Clear
    TerminateGDI
    
End Sub

Private Sub InitGDI()
    Dim GdipStartupInput As GdiplusStartupInput
    GdipStartupInput.GdiplusVersion = GdiplusVersion
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0&)
End Sub
 
Private Sub TerminateGDI()
    If GdipToken <> 0 Then Call GdiplusShutdown(GdipToken)
End Sub

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Function KeyExists(ByVal sKey As String) As Boolean
    On Error GoTo HandleError:
    Dim tVal As String 'Variant
    
    tVal = cColl(sKey)
  
    KeyExists = True
    Exit Function

HandleError:
    Err.Clear
End Function

Private Function IsAlphaIcon(ByVal IconHandle As Long) As Boolean

    Dim tARGB() As ARGB
    Dim tRECT As RECT
    Dim tICONINFO As ICONINFO
    Dim tBitmapData As BitmapData
    Dim lPixelFormat As Long
    Dim lngX As Long
    Dim lngY As Long
    Dim sngWidth As Single
    Dim sngHeight As Single
    Dim lngArgbBmp As Long
    Dim lngColorBmp As Long
    Dim bolRet As Boolean

    If GetIconInfo(IconHandle, tICONINFO) <> 0 Then
        
        If tICONINFO.hbmColor <> 0 Then

            If GdipCreateBitmapFromHBITMAP(tICONINFO.hbmColor, 0&, lngColorBmp) = 0 Then

                If GdipGetImagePixelFormat(lngColorBmp, lPixelFormat) = 0 Then

                    If lPixelFormat <> PixelFormat32bppRGB Then

                        bolRet = False
                        
                    Else

                        If GdipGetImageDimension(lngColorBmp, sngWidth, sngHeight) = 0 Then
                            
                            With tRECT
                                .Right = CLng(sngWidth)
                                .Bottom = CLng(sngHeight)
                            End With

                            ReDim tARGB(tRECT.Right - 1&, tRECT.Bottom - 1&)
                            
                            With tBitmapData
                                .Scan0 = VarPtr(tARGB(0&, 0&))
                                .Stride = 4& * tRECT.Right
                            End With
                            
                            If GdipBitmapLockBits(lngColorBmp, tRECT, ImageLockModeRead Or ImageLockModeUserInputBuf, lPixelFormat, tBitmapData) = 0 Then

                                For lngY = 0 To tBitmapData.Height - 1
                                    For lngX = 0 To tBitmapData.Width - 1
                                    
                                        If tARGB(lngX, lngY).Alpha > 0 Then
                                            If tARGB(lngX, lngY).Alpha < 255 Then
                                            
                                                bolRet = True
                                                Exit For
                                                
                                            End If
                                        End If
                                        
                                    Next lngX

                                    If bolRet Then Exit For
                                    
                                Next lngY
                                
                                Call GdipDisposeImage(lngArgbBmp)

                                Call GdipBitmapUnlockBits(lngColorBmp, tBitmapData)
                                    
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
                Call GdipDisposeImage(lngColorBmp)
                
            End If
            
            Call DeleteObject(tICONINFO.hbmColor)
        
        End If
        
        Call DeleteObject(tICONINFO.hbmMask)
        
    Else
    
        bolRet = False
        
    End If

    IsAlphaIcon = bolRet
    
End Function

'All events subclassed by this window are processed here.
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
    On Error Resume Next
    
    Select Case uMsg
    
        Case WM_MENUSELECT, WM_ENTERIDLE 'put here what you need :)
            RaiseEvent MenuMessages(lng_hWnd, uMsg, wParam, lParam)
            
        Case WM_MEASUREITEM
        
            Dim MIS As MEASUREITEMSTRUCT
            
            CopyMemory MIS, ByVal lParam, Len(MIS)

            If MIS.CtlType = ODT_MENU Then

                If MIS.itemHeight < m_lHeight + 4 Then
                    MIS.itemHeight = m_lHeight + 4
                End If
                
                If g_IsVistaOrLater Then
                    If classicThemeWorkaround Then MIS.itemWidth = MIS.itemWidth + m_lWidth + 12
                Else
                    MIS.itemWidth = MIS.itemWidth + m_lWidth + 2
                End If
                
                CopyMemory ByVal lParam, MIS, Len(MIS)
                
                lReturn = 1
                bHandled = True
            Else
                bHandled = False
            End If
        
        
        Case WM_DRAWITEM
    
            Dim DIS As DRAWITEMSTRUCT
            Dim IsDisabled As Boolean
            Dim IsCheckStyle As Boolean
            Dim IsChecked As Boolean
            Dim lLeft As Long
            Dim MI As MENUINFO
            
            CopyMemory DIS, ByVal lParam, Len(DIS)

            If DIS.CtlType = ODT_MENU Then
            
                If DIS.hWndItem <> GetMenu(lng_hWnd) Then
                        
                    With MI
                        .cbSize = Len(MI)
                        .fMask = MIM_STYLE
                    End With
    
                    GetMenuInfo DIS.hWndItem, MI
  
                    IsCheckStyle = (MI.dwStyle And MNS_NOCHECK) <> MNS_NOCHECK
    
                    IsChecked = (DIS.ItemState And ODS_CHECKED) = ODS_CHECKED
    
                    lLeft = IIf(IsCheckStyle, GetSystemMetrics(SM_CXMENUCHECK), 0)
                
                End If
                
                IsDisabled = (DIS.ItemState And ODS_GRAYED) = ODS_GRAYED
                
                If Not IsCheckStyle And IsChecked Then
                
                    Call DrawCheck(DIS.hDC, lLeft, DIS.rcItem.Top, m_lWidth + 4, m_lHeight + 4, IsDisabled)
                    
                End If
                
                'Tanner edit: this is used to prevent a bug where the hDC value is invalid for disabled and non-hovered menu entries
                ' for users of the "classic theme"
                If Not (classicThemeWorkaround And IsDisabled) Then
                    DrawDIB DIS.hDC, lLeft - 4, DIS.rcItem.Top + 2, cColl(DIS.hWndItem & "-" & DIS.itemID), IsDisabled
                End If
                
                lReturn = 1
                bHandled = True
            Else
                bHandled = False
            End If
            
        Case WM_DESTROY
            Call StopSubclassing
            
    End Select
    
' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub
