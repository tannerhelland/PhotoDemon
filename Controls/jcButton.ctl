VERSION 5.00
Begin VB.UserControl jcbutton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ToolboxBitmap   =   "jcButton.ctx":0000
End
Attribute VB_Name = "jcbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'***************************************************************************
'*  Title:      JC button
'*  Function:   An ownerdrawn multistyle button
'*  Author:     Juned Chhipa
'*  Created:    November 2008
'*  Dedicated:  To my Parents and my Teachers :-)
'*  Contact me: juned.chhipa@yahoo.com
'*
'*  Copyright � 2008-2009 Juned Chhipa. All rights reserved.
'****************************************************************************
'* This control can be used as an alternative to Command Button. It is      *
'* a lightweight button control which will emulate new button styles.       *
'*                                                                          *
'* This control uses self-subclassing routines of Paul Caton.               *
'* Feel free to use this control. Please read Documentation.chm             *
'* Please send comments/suggestions/bug reports to mail address stated above*
'****************************************************************************
'*
'* - CREDITS:
'* - Dana Seaman :-  Worked much for this control (Thanks a million)
'* - Paul Caton  :-  Self-Subclass Routines
'* - Noel Dacara :-  Inspiration for DropDown menu support
'* - Tuan Hai    :-  Numerous Suggestions and appreciating me ;)
'* - Fred.CPP    :-  For the amazing Aqua Style and for flexible tooltips
'* - Gonkuchi    :-  For his sub TransBlt to make grayscale pictures
'* - Carles P.V. :-  For fastest gradient routines
'*
'* I have tested this control painstakingly and tried my best to make
'* it work as a real command button.
'*
'****************************************************************************
'* This software is provided "as-is" without any express/implied warranty.  *
'* In no event shall the author be held liable for any damages arising      *
'* from the use of this software.                                           *
'* If you do not agree with these terms, do not install "JCButton". Use     *
'* of the program implicitly means you have agreed to these terms.          *
'*                                                                          *
'* Permission is granted to anyone to use this software for any purpose,    *
'* including commercial use, and to alter and redistribute it, provided     *
'* that the following conditions are met:                                   *
'*                                                                          *
'* 1.All redistributions of source code files must retain all copyright     *
'*   notices that are currently in place, and this list of conditions       *
'*   without any modification.                                              *
'*                                                                          *
'* 2.All redistributions in binary form must retain all occurrences of      *
'*   above copyright notice and web site addresses that are currently in    *
'*   place (for example, in the About boxes).                               *
'*                                                                          *
'* 3.Modified versions in source or binary form must be plainly marked as   *
'*   such, and must not be misrepresented as being the original software.   *
'****************************************************************************

'* N'joy ;)

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ByRef pccolorref As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As tLOGFONT) As Long
Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

'User32 Declares
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function TransparentBlt Lib "MSIMG32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' --for tooltips
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' --Theme Stuff
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function GetThemeBackgroundRegion Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pRegion As Long) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long

Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetLayout Lib "gdi32" (ByVal hdc As Long, ByVal dwLayout As Long) As Long

'==========================================================================================================================================================================================================================================================================================
' Subclassing Declares
Private Enum MsgWhen
   MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
   MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
   MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
   TME_HOVER = &H1
   TME_LEAVE = &H2
   TME_QUERY = &H40000000
   TME_CANCEL = &H80000000
End Enum

'Windows Messages
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_MOVING As Long = &H216
Private Const WM_NCACTIVATE As Long = &H86
Private Const WM_ACTIVATE As Long = &H6

Private Const ALL_MESSAGES As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04  As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05  As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08  As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09  As Long = 137                                      'Table A (after) entry count patch offset

Private Type TRACKMOUSEEVENT_STRUCT
   cbSize               As Long
   dwFlags              As TRACKMOUSEEVENT_FLAGS
   hwndTrack            As Long
   dwHoverTime          As Long
End Type

'for subclass
Private Type SubClassDatatype
   hWnd                         As Long
   nAddrSclass                  As Long
   nAddrOrig                    As Long
   nMsgCountA                   As Long
   nMsgCountB                   As Long
   aMsgTabelA()                 As Long
   aMsgTabelB()                 As Long
End Type

'for subclass
Private SubclassData()          As SubClassDatatype                                       'Subclass data array
Private TrackUser32     As Boolean

'Kernel32 declares used by the Subclasser
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'  End of Subclassing Declares
'==========================================================================================================================================================================================================================================================================================================

'[Enumerations]
Public Enum enumButtonStlyes
   [eStandard]                 '1) Standard VB Button
   [eFlat]                     '2) Standard Toolbar Button
   [eWindowsXP]                '3) Famous Win XP Button
   [eVistaAero]                '5) The New Vista Aero Button
   [eOfficeXP]
   [eOffice2003]              '13) Office 2003 Style
   [eXPToolbar]                '4) XP Toolbar
   [eVistaToolbar]             '9) Vista Toolbar Button
   [eOutlook2007]              '8) Office 2007 Outlook Button
   [eInstallShield]            '7) InstallShield?!?~?
   [eGelButton]               '11) Gel Button
   [e3DHover]                 '13) 3D Hover Button
   [eFlatHover]               '12) Flat Hover Button
   [eWindowsTheme]
End Enum

#If False Then
   Private eStandard, eFlat, eVistaAero, eVistaToolbar, eInstallShield, eFlatHover, eOffice2003
   Private eWindowsXP, eXPToolbar, e3DHover, eGelButton, eOutlook2007, eOfficeXP, eWindowsTheme
#End If

Public Enum enumButtonModes
   [ebmCommandButton]
   [ebmCheckBox]
   [ebmOptionButton]
End Enum

#If False Then
   Private ebmCommandButton, ebmCheckBox, ebmOptionButton
#End If

Public Enum enumButtonStates
   [eStateNormal]             'Normal State
   [eStateOver]               'Hover State
   [eStateDown]               'Down State
End Enum

#If False Then
   'A trick to preserve casing when typing in IDE
   Private eStateNormal, eStateOver, eStateDown, eStateFocused
#End If

Public Enum enumCaptionAlign
   [ecLeftAlign]
   [ecCenterAlign]
   [ecRightAlign]
End Enum

#If False Then
   'A trick to preserve casing when typing in IDE
   Private ecLeftAlign, ecCenterAlign, ecRightAlign
#End If

Public Enum enumPictureAlign
   [epLeftEdge]
   [epLeftOfCaption]
   [epRightEdge]
   [epRightOfCaption]
   [epBackGround]
   [epTopEdge]
   [epTopOfCaption]
   [epBottomEdge]
   [epBottomOfCaption]
End Enum

#If False Then
   Private epLeftEdge, epRightEdge, epRightOfCaption, epLeftOfCaption, epBackGround
   Private epTopEdge, epTopOfCaption, epBottomEdge, epBottomOfCaption
#End If

' --Tooltip Icons
Public Enum enumIconType
   TTNoIcon
   TTIconInfo
   TTIconWarning
   TTIconError
End Enum

#If False Then
   Private TTNoIcon, TTIconInfo, TTIconWarning, TTIconError
#End If

' --Tooltip [ Balloon / Standard ]
Public Enum enumTooltipStyle
   TooltipStandard
   TooltipBalloon
End Enum

#If False Then
   Private TooltipStandard, TooltipBalloon
#End If

' --Caption effects
Public Enum enumCaptionEffects
   [eseNone]
   [eseEmbossed]
   [eseEngraved]
   [eseShadowed]
   [eseOutline]
   [eseCover]
End Enum

#If False Then
   Private eseNone, eseEmbossed, eseEngraved, eseShadowed, eseOutline, eseCover
#End If

Public Enum enumPicEffect
   [epeNone]
   [epeLighter]
   [epeDarker]
End Enum

#If False Then
   Private epeNone, epeLighter, epeDarker, epePushUp
#End If

' --For dropdown symbols
Public Enum enumSymbol
   ebsNone
   ebsArrowUp = 5
   ebsArrowDown = 6
   ebsArrowRight = 4
End Enum

#If False Then
   Private ebsArrowUp, ebsArrowDown, ebsNone
#End If

Public Enum enumXPThemeColors
   [ecsBlue] = 0
   [ecsOliveGreen] = 1
   [ecsSilver] = 2
   [ecsCustom] = 3
End Enum

' --A trick to preserve casing of enums while typing in IDE
#If False Then
   Private ecsBlue, ecsOliveGreen, ecsSilver, ecsCustom
#End If

Public Enum enumDisabledPicMode
   [edpBlended]
   [edpGrayed]
End Enum

#If False Then
   Private edpBlended, edpGrayed
#End If

' --For gradient subs
Public Enum GradientDirectionCts
   [gdHorizontal] = 0
   [gdVertical] = 1
   [gdDownwardDiagonal] = 2
   [gdUpwardDiagonal] = 3
End Enum

' --A trick to preserve casing of enums when typing in IDE
#If False Then
   Private gdHorizontal, gdVertical, gdDownwardDiagonal, gdUpwardDiagonal
#End If

Public Enum enumMenuAlign
   [edaBottom] = 0
   [edaTop] = 1
   [edaLeft] = 2
   [edaRight] = 3
   [edaTopLeft] = 4
   [edaBottomLeft] = 5
   [edaTopRight] = 6
   [edaBottomRight] = 7
End Enum

#If False Then
   Private edaBottom, edaTop, edaTopLeft, edaBottomLeft, edaTopRight, edaBottomRight
#End If

'  used for Button colors
Private Type tButtonColors
   tBackColor           As Long
   tDisabledColor       As Long
   tForeColor           As Long
   tForeColorOver       As Long
   tGreyText            As Long
End Type

'  used to define various graphics areas
Private Type RECT
   Left                 As Long
   Top                  As Long
   Right                As Long
   Bottom               As Long
End Type

''Tooltip Window Types
Private Type TOOLINFO
   lSize                As Long
   lFlags               As Long
   lHwnd                As Long
   lId                  As Long
   lpRect               As RECT
   hInstance            As Long
   lpStr                As String
   lParam               As Long
End Type

''Tooltip Window Types [for UNICODE support]
Private Type TOOLINFOW
   lSize                As Long
   lFlags               As Long
   lHwnd                As Long
   lId                  As Long
   lpRect               As RECT
   hInstance            As Long
   lpStrW               As Long
   lParam               As Long
End Type

Private Type POINT
   x                    As Long
   Y                    As Long
End Type

' --Used for creating a drop down symbol
' --I m using Marlett Font to create that symbol
Private Type tLOGFONT
   lfHeight             As Long
   lfWidth              As Long
   lfEscapement         As Long
   lfOrientation        As Long
   lfWeight             As Long
   lfItalic             As Byte
   lfUnderline          As Byte
   lfStrikeOut          As Byte
   lfCharSet            As Byte
   lfOutPrecision       As Byte
   lfClipPrecision      As Byte
   lfQuality            As Byte
   lfPitchAndFamily     As Byte
   lfFaceName           As String * 32
End Type

'  RGB Colors structure
Private Type RGBColor
   r                    As Single
   g                    As Single
   B                    As Single
End Type

Private Type BITMAP
   bmType               As Long
   bmWidth              As Long
   bmHeight             As Long
   bmWidthBytes         As Long
   bmPlanes             As Integer
   bmBitsPixel          As Integer
   bmBits               As Long
End Type

'  for gradient painting and bitmap tiling
Private Type BITMAPINFOHEADER
   biSize               As Long
   biWidth              As Long
   biHeight             As Long
   biPlanes             As Integer
   biBitCount           As Integer
   biCompression        As Long
   biSizeImage          As Long
   biXPelsPerMeter      As Long
   biYPelsPerMeter      As Long
   biClrUsed            As Long
   biClrImportant       As Long
End Type

Private Type ICONINFO
   fIcon                As Long
   xHotspot             As Long
   yHotspot             As Long
   hbmMask              As Long
   hbmColor             As Long
End Type

Private Type RGBTRIPLE
   rgbBlue              As Byte
   rgbGreen             As Byte
   rgbRed               As Byte
End Type

Private Type RGBQUAD
   rgbBlue              As Byte
   rgbGreen             As Byte
   rgbRed               As Byte
   rgbAlpha             As Byte
End Type

Private Type BITMAPINFO
   bmiHeader            As BITMAPINFOHEADER
   bmiColors            As RGBTRIPLE
End Type

Private Type OSVERSIONINFO
   dwOSVersionInfoSize  As Long
   dwMajorVersion       As Long
   dwMinorVersion       As Long
   dwBuildNumber        As Long
   dwPlatformId         As Long
   szCSDVersion         As String * 128 '* Maintenance string for PSS usage.
End Type

' --constants for unicode support
Private Const VER_PLATFORM_WIN32_NT = 2

' --constants for  Flat Button
Private Const BDR_RAISEDINNER As Long = &H4

' --constants for Win 98 style buttons
Private Const BDR_SUNKEN95 As Long = &HA
Private Const BDR_RAISED95 As Long = &H5

Private Const BF_LEFT   As Long = &H1
Private Const BF_TOP    As Long = &H2
Private Const BF_RIGHT  As Long = &H4
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RECT   As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

' --System Hand Pointer
Private Const IDC_HAND  As Long = 32649

' --Color Constant
Private Const COLOR_BTNFACE As Long = 15
Private Const COLOR_BTNHIGHLIGHT As Long = 20
Private Const COLOR_BTNSHADOW As Long = 16
Private Const COLOR_HIGHLIGHT As Long = 13
Private Const COLOR_GRAYTEXT As Long = 17
Private Const CLR_INVALID As Long = &HFFFF
Private Const DIB_RGB_COLORS As Long = 0

' --Windows Messages
Private Const WM_USER   As Long = &H400
Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const SWP_REFRESH As Long = (&H1 Or &H2 Or &H4 Or &H20)
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const HWND_TOPMOST As Long = -&H1
Private Const CW_USEDEFAULT As Long = &H80000000

''Tooltip Window Constants
Private Const TTS_NOPREFIX As Long = &H2
Private Const TTF_TRANSPARENT As Long = &H100
Private Const TTF_IDISHWND As Long = &H1
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const TTF_CENTERTIP As Long = &H2
Private Const TTM_ADDTOOLA As Long = (WM_USER + 4)
Private Const TTM_ADDTOOLW As Long = (WM_USER + 50)
Private Const TTM_ACTIVATE As Long = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA As Long = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR As Long = (WM_USER + 20)
Private Const TTM_SETTITLE As Long = (WM_USER + 32)
Private Const TTM_SETTITLEW As Long = (WM_USER + 33)
Private Const TTS_BALLOON As Long = &H40
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTF_SUBCLASS As Long = &H10
Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"

' --Formatting Text Consts
Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_RTLREADING = &H20000              ' Right to left
Private Const DT_DRAWFLAG As Long = DT_CENTER Or DT_WORDBREAK

' --drawing Icon Constants
Private Const DI_NORMAL As Long = &H3

' --Property Variables:

Private m_ButtonStyle   As enumButtonStlyes     'Choose your Style
Private m_Buttonstate   As enumButtonStates     'Normal / Over / Down

Private m_bIsDown       As Boolean              'Is button is pressed?
Private m_bMouseInCtl   As Boolean              'Is Mouse in Control
Private m_bHasFocus     As Boolean              'Has focus?
Private m_bHandPointer  As Boolean              'Use Hand Pointer
Private m_lCursor       As Long
Private m_bDefault      As Boolean              'Is Default?
Private m_DropDownSymbol As enumSymbol
Private m_bDropDownSep  As Boolean
Private m_ButtonMode    As enumButtonModes      'Command/Check/Option button
Private m_CaptionEffects As enumCaptionEffects
Private m_bValue        As Boolean              'Value (Checked/Unchekhed)
Private m_bShowFocus    As Boolean              'Bool to show focus
Private m_bParentActive As Boolean              'Parent form Active or not
Private m_lParenthWnd   As Long                 'Is parent active?
Private m_WindowsNT     As Long                 'OS Supports Unicode?
Private m_bEnabled      As Boolean              'Enabled/Disabled
Private m_Caption       As String               'String to draw caption
Private m_CaptionAlign  As enumCaptionAlign
Private m_bColors       As tButtonColors        'Button Colors
Private m_bUseMaskColor As Boolean              'Transparent areas
Private m_lMaskColor    As Long                 'Set Transparent color
Private m_lButtonRgn    As Long                 'Button Region
Private m_bIsSpaceBarDown As Boolean              'Space bar down boolean
Private m_ButtonRect    As RECT                 'Button Position
Private m_FocusRect     As RECT
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private m_lXPColor      As enumXPThemeColors
Private m_bIsThemed     As Boolean
Private m_bHasUxTheme   As Boolean

Private m_lDownButton   As Integer              'For click/Dblclick events
Private m_lDShift       As Integer              'A flag for dblClick
Private m_lDX           As Single
Private m_lDY           As Single

' --Popup menu variables
Private m_bPopupEnabled As Boolean              'Popus is enabled
Private m_bPopupShown   As Boolean              'Popupmenu is shown
Private m_bPopupInit    As Boolean              'Flag to prevent WM_MOUSLEAVE to redraw the button
Private DropDownMenu    As VB.Menu              'Popupmenu to be shown
Private MenuAlign       As enumMenuAlign        'PopupMenu Alignments
Private MenuFlags       As Long                 'PopupMenu Flags
Private DefaultMenu     As VB.Menu              'Default menu in the popupmenu

' --Tooltip variables
Private m_sTooltipText  As String
Private m_sTooltiptitle As String
Private m_lToolTipIcon  As enumIconType
Private m_lTooltipType  As enumTooltipStyle
Private m_lttBackColor  As Long
Private m_lttCentered   As Boolean
Private m_bttRTL        As Boolean              'Right to Left reading
Private m_lttHwnd       As Long
Private m_hMode         As Long                 'Added this, as tooltips
                                                'were not displayed in
                                                'compiled exe. (Thanks to Jim Jose)
' --Caption variables
Private CaptionW        As Long                 'Width of Caption
Private CaptionH        As Long                 'Height of Caption
Private CaptionX        As Long                 'Left of Caption
Private CaptionY        As Long                 'Top of Caption
Private lpSignRect      As RECT                 'Drop down Symbol rect
Private m_bRTL          As Boolean
Private m_TextRect      As RECT                 'Caption drawing area

' --Picture variables
Private m_Picture       As StdPicture
Private m_PictureHot    As StdPicture
Private m_PictureDown   As StdPicture
Private m_PictureOpacity As Byte
Private m_PicOpacityOnOver As Byte
Private m_PicDisabledMode As enumDisabledPicMode
Private m_PictureShadow As Boolean
Private m_PictureAlign  As enumPictureAlign     'Picture Alignments
Private m_PicEffectonOver As enumPicEffect
Private m_PicEffectonDown As enumPicEffect
Private m_bPicPushOnHover As Boolean
Private PicH            As Long
Private PicW            As Long
Private aLighten(255)   As Byte                 'Light Picture
Private aDarken(255)    As Byte                 'Dark Picture

Private tmppic          As New StdPicture       'Temp picture
Private PicX            As Long                 'X position of picture
Private PicY            As Long                 'Y Position of Picture
Private m_PicRect       As RECT                 'Picture drawing area
Private lh              As Long                 'ScaleHeight of button
Private lw              As Long                 'ScaleWidth of button

'  Events
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over the button."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user clicks over the button twice."
Attribute DblClick.VB_UserMemId = -601
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occrus when the cursor moves around the button for the first time."
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Occurs when the cursor leaves/moves outside the button."
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the cursor moves over the button."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while the button has the focus."
Attribute MouseUp.VB_UserMemId = -607
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while the button has the focus."
Attribute MouseDown.VB_UserMemId = -605
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while the button has the focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while the button has the focus."
Attribute KeyUp.VB_UserMemId = -604
Public Event KeyPress(KeyAcsii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603

Private Sub DrawLineApi(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)

   '****************************************************************************
   '*  draw lines
   '****************************************************************************

   Dim pt               As POINT
   Dim hPen             As Long
   Dim hPenOld          As Long

   hPen = CreatePen(0, 1, Color)
   hPenOld = SelectObject(hdc, hPen)
   MoveToEx hdc, X1, Y1, pt
   LineTo hdc, X2, Y2
   SelectObject hdc, hPenOld
   DeleteObject hPen

End Sub

Private Function BlendColorEx(Color1 As Long, Color2 As Long, Optional Percent As Long) As Long

   'Combines two colors together by how many percent.

   Dim r1               As Long, g1 As Long, b1 As Long
   Dim r2               As Long, g2 As Long, b2 As Long
   Dim r3               As Long, g3 As Long, b3 As Long

   If Percent <= 0 Then Percent = 0
   If Percent >= 100 Then Percent = 100

   r1 = Color1 And 255
   g1 = (Color1 \ 256) And 255
   b1 = (Color1 \ 65536) And 255

   r2 = Color2 And 255
   g2 = (Color2 \ 256) And 255
   b2 = (Color2 \ 65536) And 255

   r3 = r1 + (r1 - r2) * Percent \ 100
   g3 = g1 + (g1 - g2) * Percent \ 100
   b3 = b1 + (b1 - b2) * Percent \ 100

   BlendColorEx = r3 + 256& * g3 + 65536 * b3

End Function

Private Function BlendColors(ByVal lBackColorFrom As Long, ByVal lBackColorTo As Long) As Long

   '***************************************************************************
   '*  Combines (mix) two colors                                              *
   '*  This is another method in which you can't specify percentage
   '***************************************************************************

   BlendColors = RGB(((lBackColorFrom And &HFF) + (lBackColorTo And &HFF)) / 2, (((lBackColorFrom \ &H100) And &HFF) + ((lBackColorTo \ &H100) And &HFF)) / 2, (((lBackColorFrom \ &H10000) And &HFF) + ((lBackColorTo \ &H10000) And &HFF)) / 2)

End Function

Private Sub DrawRectangle(ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long)

   '****************************************************************************
   '*  Draws a rectangle specified by coords and color of the rectangle        *
   '****************************************************************************

   Dim brect            As RECT
   Dim hBrush           As Long
   Dim ret              As Long

   With brect
      .Left = x
      .Top = Y
      .Right = x + Width
      .Bottom = Y + Height
   End With

   hBrush = CreateSolidBrush(Color)

   ret = FrameRect(hdc, brect, hBrush)

   ret = DeleteObject(hBrush)

End Sub

Private Sub DrawFocusRectangle(ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)

   '****************************************************************************
   '*  Draws a Focus Rectangle inside button if m_bShowFocus property is True  *
   '****************************************************************************

   Dim brect            As RECT
   Dim RetVal           As Long

   With brect
      .Left = x
      .Top = Y
      .Right = x + Width
      .Bottom = Y + Height
   End With

   RetVal = DrawFocusRect(hdc, brect)

End Sub

Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal TransColor As Long = -1, Optional ByVal BrushColor As Long = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False)

   '****************************************************************************
   '* Routine : To make transparent and grayscale images
   '* Author  : Gonkuchi
   '
   '* Modified by Dana Seaman
   '****************************************************************************

   Dim B                As Long, h As Long, F As Long, i As Long, newW As Long
   Dim TmpDC            As Long, TmpBmp As Long, TmpObj As Long
   Dim Sr2DC            As Long, Sr2Bmp As Long, Sr2Obj As Long
   Dim DataDest()       As RGBTRIPLE, DataSrc() As RGBTRIPLE
   Dim Info             As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
   Dim hOldOb           As Long, PicEffect As enumPicEffect
   Dim SrcDC            As Long, tObj As Long, ttt As Long
   Dim bDisOpacity      As Byte
   Dim OverOpacity      As Byte
   Dim a2               As Long
   Dim a1               As Long

   If DstW = 0 Or DstH = 0 Then Exit Sub
   If SrcPic Is Nothing Then Exit Sub

   If m_Buttonstate = eStateOver Then
      PicEffect = m_PicEffectonOver
   ElseIf m_Buttonstate = eStateDown Then
      PicEffect = m_PicEffectonDown
   End If
   
   If Not m_bEnabled Then
      Select Case m_PicDisabledMode
      Case edpBlended
         bDisOpacity = 52
      Case edpGrayed
         bDisOpacity = m_PictureOpacity * 0.75
         isGreyscale = True
      End Select
   End If
   
   If m_Buttonstate = eStateOver Then
      OverOpacity = m_PicOpacityOnOver
   End If

   SrcDC = CreateCompatibleDC(hdc)

   If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
   If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)

   If SrcPic.Type = vbPicTypeBitmap Then 'check if it's an icon or a bitmap
      tObj = SelectObject(SrcDC, SrcPic)
   Else
      Dim hBrush           As Long
      tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
      hBrush = CreateSolidBrush(TransColor)
      DrawIconEx SrcDC, 0, 0, SrcPic.Handle, DstW, DstH, 0, hBrush, DI_NORMAL
      DeleteObject hBrush
   End If

   TmpDC = CreateCompatibleDC(SrcDC)
   Sr2DC = CreateCompatibleDC(SrcDC)
   TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
   Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
   TmpObj = SelectObject(TmpDC, TmpBmp)
   Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
   ReDim DataDest(DstW * DstH * 3 - 1)
   ReDim DataSrc(UBound(DataDest))
   With Info.bmiHeader
      .biSize = Len(Info.bmiHeader)
      .biWidth = DstW
      .biHeight = DstH
      .biPlanes = 1
      .biBitCount = 24
   End With

   BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
   BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
   GetDIBits TmpDC, TmpBmp, 0, DstH, DataDest(0), Info, 0
   GetDIBits Sr2DC, Sr2Bmp, 0, DstH, DataSrc(0), Info, 0

   If BrushColor > 0 Then
      BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
      BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
      BrushRGB.rgbRed = BrushColor And &HFF
   End If

   ' --No Maskcolor to use
   If Not m_bUseMaskColor Then TransColor = -1

   newW = DstW - 1

   For h = 0 To DstH - 1
      F = h * DstW
      For B = 0 To newW
         i = F + B
         If m_Buttonstate = eStateOver Then
            a1 = OverOpacity
         Else
            a1 = IIf(m_bEnabled, m_PictureOpacity, bDisOpacity)
         End If
         a2 = 255 - a1
         If GetNearestColor(hdc, CLng(DataSrc(i).rgbRed) + 256& * DataSrc(i).rgbGreen + 65536 * DataSrc(i).rgbBlue) <> TransColor Then
            With DataDest(i)
               If BrushColor > -1 Then
                  If MonoMask Then
                     If (CLng(DataSrc(i).rgbRed) + DataSrc(i).rgbGreen + DataSrc(i).rgbBlue) <= 384 Then DataDest(i) = BrushRGB
                  Else
                     If a1 = 255 Then
                        DataDest(i) = BrushRGB
                     ElseIf a1 > 0 Then
                        .rgbRed = (a2 * .rgbRed + a1 * BrushRGB.rgbRed) \ 256
                        .rgbGreen = (a2 * .rgbGreen + a1 * BrushRGB.rgbGreen) \ 256
                        .rgbBlue = (a2 * .rgbBlue + a1 * BrushRGB.rgbBlue) \ 256
                     End If
                  End If
               Else
                  If isGreyscale Then
                     gCol = CLng(DataSrc(i).rgbRed * 0.3) + DataSrc(i).rgbGreen * 0.59 + DataSrc(i).rgbBlue * 0.11
                     If a1 = 255 Then
                        .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                     ElseIf a1 > 0 Then
                        .rgbRed = (a2 * .rgbRed + a1 * gCol) \ 256
                        .rgbGreen = (a2 * .rgbGreen + a1 * gCol) \ 256
                        .rgbBlue = (a2 * .rgbBlue + a1 * gCol) \ 256
                     End If
                  Else
                     If a1 = 255 Then
                        If PicEffect = epeLighter Then
                           .rgbRed = aLighten(DataSrc(i).rgbRed)
                           .rgbGreen = aLighten(DataSrc(i).rgbGreen)
                           .rgbBlue = aLighten(DataSrc(i).rgbBlue)
                        ElseIf PicEffect = epeDarker Then
                           .rgbRed = aDarken(DataSrc(i).rgbRed)
                           .rgbGreen = aDarken(DataSrc(i).rgbGreen)
                           .rgbBlue = aDarken(DataSrc(i).rgbBlue)
                        Else
                           DataDest(i) = DataSrc(i)
                        End If
                     ElseIf a1 > 0 Then
                        If (PicEffect = epeLighter) Then
                           .rgbRed = (a2 * .rgbRed + a1 * aLighten(DataSrc(i).rgbRed)) \ 256
                           .rgbGreen = (a2 * .rgbGreen + a1 * aLighten(DataSrc(i).rgbGreen)) \ 256
                           .rgbBlue = (a2 * .rgbBlue + a1 * aLighten(DataSrc(i).rgbBlue)) \ 256
                        ElseIf PicEffect = epeDarker Then
                           .rgbRed = (a2 * .rgbRed + a1 * aDarken(DataSrc(i).rgbRed)) \ 256
                           .rgbGreen = (a2 * .rgbGreen + a1 * aDarken(DataSrc(i).rgbGreen)) \ 256
                           .rgbBlue = (a2 * .rgbBlue + a1 * aDarken(DataSrc(i).rgbBlue)) \ 256
                        Else
                           .rgbRed = (a2 * .rgbRed + a1 * DataSrc(i).rgbRed) \ 256
                           .rgbGreen = (a2 * .rgbGreen + a1 * DataSrc(i).rgbGreen) \ 256
                           .rgbBlue = (a2 * .rgbBlue + a1 * DataSrc(i).rgbBlue) \ 256
                        End If
                     End If
                  End If
               End If
            End With
         End If
      Next B
   Next h

   ' /--Paint it!
   SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, DataDest(0), Info, 0

   Erase DataDest, DataSrc
   DeleteObject SelectObject(TmpDC, TmpObj)
   DeleteObject SelectObject(Sr2DC, Sr2Obj)
   If SrcPic.Type = vbPicTypeIcon Then DeleteObject SelectObject(SrcDC, tObj)
   DeleteDC TmpDC
   DeleteDC Sr2DC
   DeleteObject tObj
   DeleteDC SrcDC

End Sub

Private Sub TransBlt32(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal BrushColor As Long = -1, Optional ByVal isGreyscale As Boolean = False)

   '****************************************************************************
   '* Routine : Renders 32 bit Bitmap                                          *
   '* Author  : Dana Seaman                                                    *
   '****************************************************************************

   Dim B                As Long, h As Long, F As Long, i As Long, newW As Long
   Dim TmpDC            As Long, TmpBmp As Long, TmpObj As Long
   Dim Sr2DC            As Long, Sr2Bmp As Long, Sr2Obj As Long
   Dim DataDest()       As RGBQUAD, DataSrc() As RGBQUAD
   Dim Info             As BITMAPINFO, BrushRGB As RGBQUAD, gCol As Long
   Dim hOldOb           As Long, PicEffect As enumPicEffect
   Dim SrcDC            As Long, tObj As Long, ttt As Long
   Dim bDisOpacity      As Byte
   Dim OverOpacity      As Byte
   Dim a2               As Long
   Dim a1               As Long

   If DstW = 0 Or DstH = 0 Then Exit Sub
   If SrcPic Is Nothing Then Exit Sub

   If m_Buttonstate = eStateOver Then
      PicEffect = m_PicEffectonOver
   ElseIf m_Buttonstate = eStateDown Then
      PicEffect = m_PicEffectonDown
   End If
   
   If Not m_bEnabled Then
      Select Case m_PicDisabledMode
      Case edpBlended
         bDisOpacity = 52
      Case edpGrayed
         bDisOpacity = m_PictureOpacity * 0.75
         isGreyscale = True
      End Select
   End If
      
   If m_Buttonstate = eStateOver Then
      OverOpacity = m_PicOpacityOnOver
   End If
   
   SrcDC = CreateCompatibleDC(hdc)

   If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
   If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)

   tObj = SelectObject(SrcDC, SrcPic)

   TmpDC = CreateCompatibleDC(SrcDC)
   Sr2DC = CreateCompatibleDC(SrcDC)

   TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
   Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
   TmpObj = SelectObject(TmpDC, TmpBmp)
   Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)

   With Info.bmiHeader
      .biSize = Len(Info.bmiHeader)
      .biWidth = DstW
      .biHeight = DstH
      .biPlanes = 1
      .biBitCount = 32
      .biSizeImage = 4 * ((DstW * .biBitCount + 31) \ 32) * DstH
   End With
   ReDim DataDest(Info.bmiHeader.biSizeImage - 1)
   ReDim DataSrc(UBound(DataDest))

   BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
   BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
   GetDIBits TmpDC, TmpBmp, 0, DstH, DataDest(0), Info, 0
   GetDIBits Sr2DC, Sr2Bmp, 0, DstH, DataSrc(0), Info, 0

   If BrushColor <> -1 Then
      BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
      BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
      BrushRGB.rgbRed = BrushColor And &HFF
   End If

   newW = DstW - 1

   For h = 0 To DstH - 1
      F = h * DstW
      For B = 0 To newW
         i = F + B
         If m_bEnabled Then
             If m_Buttonstate = eStateOver Then
                a1 = (CLng(DataSrc(i).rgbAlpha) * OverOpacity) \ 255
             Else
                a1 = (CLng(DataSrc(i).rgbAlpha) * m_PictureOpacity) \ 255
             End If
         Else
            a1 = (CLng(DataSrc(i).rgbAlpha) * bDisOpacity) \ 255
         End If
         a2 = 255 - a1
         With DataDest(i)
            If BrushColor <> -1 Then
               If a1 = 255 Then
                  DataDest(i) = BrushRGB
               ElseIf a1 > 0 Then
                  .rgbRed = (a2 * .rgbRed + a1 * BrushRGB.rgbRed) \ 256
                  .rgbGreen = (a2 * .rgbGreen + a1 * BrushRGB.rgbGreen) \ 256
                  .rgbBlue = (a2 * .rgbBlue + a1 * BrushRGB.rgbBlue) \ 256
               End If
            Else
               If isGreyscale Then
                  gCol = CLng(DataSrc(i).rgbRed * 0.3) + DataSrc(i).rgbGreen * 0.59 + DataSrc(i).rgbBlue * 0.11
                  If a1 = 255 Then
                     .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                  ElseIf a1 > 0 Then
                     .rgbRed = (a2 * .rgbRed + a1 * gCol) \ 256
                     .rgbGreen = (a2 * .rgbGreen + a1 * gCol) \ 256
                     .rgbBlue = (a2 * .rgbBlue + a1 * gCol) \ 256
                  End If
               Else
                  If a1 = 255 Then
                     If (PicEffect = epeLighter) Then
                        .rgbRed = aLighten(DataSrc(i).rgbRed)
                        .rgbGreen = aLighten(DataSrc(i).rgbGreen)
                        .rgbBlue = aLighten(DataSrc(i).rgbBlue)
                     ElseIf PicEffect = epeDarker Then
                        .rgbRed = aDarken(DataSrc(i).rgbRed)
                        .rgbGreen = aDarken(DataSrc(i).rgbGreen)
                        .rgbBlue = aDarken(DataSrc(i).rgbBlue)
                     Else
                        DataDest(i) = DataSrc(i)
                     End If
                  ElseIf a1 > 0 Then
                     If (PicEffect = epeLighter) Then
                        .rgbRed = (a2 * .rgbRed + a1 * aLighten(DataSrc(i).rgbRed)) \ 256
                        .rgbGreen = (a2 * .rgbGreen + a1 * aLighten(DataSrc(i).rgbGreen)) \ 256
                        .rgbBlue = (a2 * .rgbBlue + a1 * aLighten(DataSrc(i).rgbBlue)) \ 256
                     ElseIf PicEffect = epeDarker Then
                        .rgbRed = (a2 * .rgbRed + a1 * aDarken(DataSrc(i).rgbRed)) \ 256
                        .rgbGreen = (a2 * .rgbGreen + a1 * aDarken(DataSrc(i).rgbGreen)) \ 256
                        .rgbBlue = (a2 * .rgbBlue + a1 * aDarken(DataSrc(i).rgbBlue)) \ 256
                     Else
                        .rgbRed = (a2 * .rgbRed + a1 * DataSrc(i).rgbRed) \ 256
                        .rgbGreen = (a2 * .rgbGreen + a1 * DataSrc(i).rgbGreen) \ 256
                        .rgbBlue = (a2 * .rgbBlue + a1 * DataSrc(i).rgbBlue) \ 256
                     End If
                  End If
               End If
            End If
         End With
      Next B
   Next h

   ' /--Paint it!
   SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, DataDest(0), Info, 0

   Erase DataDest, DataSrc
   DeleteObject SelectObject(TmpDC, TmpObj)
   DeleteObject SelectObject(Sr2DC, Sr2Obj)
   If SrcPic.Type = vbPicTypeIcon Then DeleteObject SelectObject(SrcDC, tObj)
   DeleteDC TmpDC
   DeleteDC Sr2DC
   DeleteObject tObj
   DeleteDC SrcDC

End Sub

' --By Dana Seaman
Private Function Lighten(ByVal Color As Byte) As Byte

   Dim lColor           As Long
   lColor = (293& * Color) \ 255
   If lColor > 255 Then
      Lighten = 255
   Else
      Lighten = lColor
   End If

End Function

' --By Dana Seaman
Private Function Darken(ByVal Color As Byte) As Byte

   Darken = (217& * Color) \ 255

End Function

Private Sub DrawGradientEx(ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal GradientDirection As GradientDirectionCts)

   '****************************************************************************
   '* Draws very fast Gradient in four direction.                              *
   '* Author: Carles P.V (Gradient Master)                                     *
   '* This routine works as a heart for this control.                          *
   '* Thank you so much Carles.                                                *
   '****************************************************************************

   Dim uBIH             As BITMAPINFOHEADER
   Dim lBits()          As Long
   Dim lGrad()          As Long

   Dim r1               As Long
   Dim g1               As Long
   Dim b1               As Long
   Dim r2               As Long
   Dim g2               As Long
   Dim b2               As Long
   Dim dR               As Long
   Dim dG               As Long
   Dim dB               As Long

   Dim Scan             As Long
   Dim i                As Long
   Dim iEnd             As Long
   Dim iOffset          As Long
   Dim j                As Long
   Dim jEnd             As Long
   Dim iGrad            As Long

   '-- A minor check

   'If (Width < 1 Or Height < 1) Then Exit Sub
   If (Width < 1 Or Height < 1) Then
      Exit Sub
   End If

   '-- Decompose colors
   Color1 = Color1 And &HFFFFFF
   r1 = Color1 Mod &H100&
   Color1 = Color1 \ &H100&
   g1 = Color1 Mod &H100&
   Color1 = Color1 \ &H100&
   b1 = Color1 Mod &H100&
   Color2 = Color2 And &HFFFFFF
   r2 = Color2 Mod &H100&
   Color2 = Color2 \ &H100&
   g2 = Color2 Mod &H100&
   Color2 = Color2 \ &H100&
   b2 = Color2 Mod &H100&

   '-- Get color distances
   dR = r2 - r1
   dG = g2 - g1
   dB = b2 - b1

   '-- Size gradient-colors array
   Select Case GradientDirection
      Case [gdHorizontal]
         ReDim lGrad(0 To Width - 1)
      Case [gdVertical]
         ReDim lGrad(0 To Height - 1)
      Case Else
         ReDim lGrad(0 To Width + Height - 2)
   End Select

   '-- Calculate gradient-colors
   iEnd = UBound(lGrad())
   If (iEnd = 0) Then
      '-- Special case (1-pixel wide gradient)
      lGrad(0) = (b1 \ 2 + b2 \ 2) + 256 * (g1 \ 2 + g2 \ 2) + 65536 * (r1 \ 2 + r2 \ 2)
   Else
      For i = 0 To iEnd
         lGrad(i) = b1 + (dB * i) \ iEnd + 256 * (g1 + (dG * i) \ iEnd) + 65536 * (r1 + (dR * i) \ iEnd)
      Next i
   End If

   '-- Size DIB array
   ReDim lBits(Width * Height - 1) As Long
   iEnd = Width - 1
   jEnd = Height - 1
   Scan = Width

   '-- Render gradient DIB
   Select Case GradientDirection

      Case [gdHorizontal]

         For j = 0 To jEnd
            For i = iOffset To iEnd + iOffset
               lBits(i) = lGrad(i - iOffset)
            Next i
            iOffset = iOffset + Scan
         Next j

      Case [gdVertical]

         For j = jEnd To 0 Step -1
            For i = iOffset To iEnd + iOffset
               lBits(i) = lGrad(j)
            Next i
            iOffset = iOffset + Scan
         Next j

      Case [gdDownwardDiagonal]

         iOffset = jEnd * Scan
         For j = 1 To jEnd + 1
            For i = iOffset To iEnd + iOffset
               lBits(i) = lGrad(iGrad)
               iGrad = iGrad + 1
            Next i
            iOffset = iOffset - Scan
            iGrad = j
         Next j

      Case [gdUpwardDiagonal]

         iOffset = 0
         For j = 1 To jEnd + 1
            For i = iOffset To iEnd + iOffset
               lBits(i) = lGrad(iGrad)
               iGrad = iGrad + 1
            Next i
            iOffset = iOffset + Scan
            iGrad = j
         Next j
   End Select

   '-- Define DIB header
   With uBIH
      .biSize = 40
      .biPlanes = 1
      .biBitCount = 32
      .biWidth = Width
      .biHeight = Height
   End With

   '-- Paint it!
   StretchDIBits hdc, x, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy

End Sub

Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional ByRef hPalette As Long = 0) As Long

   '****************************************************************************
   '*  System color code to long rgb                                           *
   '****************************************************************************

   If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
      TranslateColor = CLR_INVALID
   End If

End Function

Private Sub RedrawButton()

   '****************************************************************************
   '*  The main routine of this usercontrol. Everything is drawn here.         *
   '****************************************************************************

   UserControl.Cls                                'Clears usercontrol
   lh = ScaleHeight
   lw = ScaleWidth

   SetRect m_ButtonRect, 0, 0, lw, lh             'Sets the button rectangle

   If (m_ButtonMode <> ebmCommandButton) Then                        'If Checkboxmode True
      If Not (m_ButtonStyle = eStandard Or m_ButtonStyle = eXPToolbar) Then
         If m_bValue Then m_Buttonstate = eStateDown
      End If
   End If
   
   Select Case m_ButtonStyle

      Case eStandard
         DrawStandardButton m_Buttonstate
      Case e3DHover
         DrawStandardButton m_Buttonstate
      Case eFlat
         DrawStandardButton m_Buttonstate
      Case eFlatHover
         DrawStandardButton m_Buttonstate
      Case eWindowsXP
         DrawWinXPButton m_Buttonstate
      Case eXPToolbar
         DrawXPToolbar m_Buttonstate
      Case eGelButton
         DrawGelButton m_Buttonstate
      Case eOfficeXP
         DrawOfficeXP m_Buttonstate
      Case eInstallShield
         DrawInstallShieldButton m_Buttonstate
      Case eVistaAero
         DrawVistaButton m_Buttonstate
      Case eVistaToolbar
         DrawVistaToolbarStyle m_Buttonstate
      Case eOutlook2007
         DrawOutlook2007 m_Buttonstate
      Case eOffice2003
         DrawOffice2003 m_Buttonstate
      Case eWindowsTheme
         If IsThemed Then
            ' --Theme can be applied
            WindowsThemeButton m_Buttonstate
         Else
            ' --Fallback to ownerdraw WinXP Button
            m_ButtonStyle = eWindowsXP
            m_lXPColor = ecsBlue
            SetThemeColors
            DrawWinXPButton m_Buttonstate
         End If
   End Select

End Sub

Private Sub CreateRegion()

   '***************************************************************************
   '*  Create region everytime you redraw a button.                           *
   '*  Because some settings may have changed the button regions              *
   '***************************************************************************

   If m_lButtonRgn Then DeleteObject m_lButtonRgn
   Select Case m_ButtonStyle
      Case eWindowsXP, eVistaAero, eVistaToolbar, eInstallShield
         m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 3, 3)
      Case eGelButton, eXPToolbar
         m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 4, 4)
      Case Else
         m_lButtonRgn = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
   End Select
   SetWindowRgn UserControl.hWnd, m_lButtonRgn, True       'Set Button Region
   DeleteObject m_lButtonRgn                               'Free memory

End Sub

Private Sub DrawSymbol(ByVal eArrow As enumSymbol)

   Dim hOldFont         As Long
   Dim hNewFont         As Long
   Dim sSign            As String
   Dim BtnSymbol        As enumSymbol

   hNewFont = BuildSymbolFont(14)
   hOldFont = SelectObject(hdc, hNewFont)

   sSign = eArrow
   DrawText hdc, sSign, 1, lpSignRect, DT_WORDBREAK '!!
   DeleteObject hNewFont

End Sub

Private Function BuildSymbolFont(lFontSize As Long) As Long

   Const SYMBOL_CHARSET = 2
   Dim lpFont           As tLOGFONT

   With lpFont
      .lfFaceName = "Marlett" + vbNullChar    'Standard Marlett Font
      .lfHeight = lFontSize                   'I was using Webdings first,
      .lfCharSet = SYMBOL_CHARSET             'but I am not sure whether
   End With                                    'it is installed in every machine!
   'Still Im not sure about Marlet :)
   BuildSymbolFont = CreateFontIndirect(lpFont) 'I got inspirations from
   'Light Templer's Project
End Function

Private Sub DrawPicwithCaption()

'****************************************************************************
' Calculate Caption rects and draw the pictures and caption                 *
'****************************************************************************
   Dim lpRect           As RECT            'RECT to draw caption
   Dim pRect            As RECT
   Dim lShadowClr       As Long
   Dim lPixelClr        As Long
   lw = ScaleWidth                         'ScaleHeight of Button
   lh = ScaleHeight                        'ScaleWidth of Button

   If (m_Buttonstate = eStateDown Or (m_ButtonMode <> ebmCommandButton And m_bValue = True)) Then
      '-- Mouse down
      If Not m_PictureDown Is Nothing Then
         Set tmppic = m_PictureDown
      Else
         If Not m_PictureHot Is Nothing Then
            Set tmppic = m_PictureHot
         Else
            Set tmppic = m_Picture
         End If
      End If
   ElseIf (m_Buttonstate = eStateOver) Then
      '-- Mouse in (over)
      If Not m_PictureHot Is Nothing Then
         Set tmppic = m_PictureHot
      Else
         Set tmppic = m_Picture
      End If
   Else
      '-- Mouse out (normal)
      Set tmppic = m_Picture
   End If

   ' --Adjust Picture Sizes
   PicH = ScaleX(tmppic.Height, vbHimetric, vbPixels)
   PicW = ScaleX(tmppic.Width, vbHimetric, vbPixels)
   
   ' --Get the drawing area of caption
   If m_DropDownSymbol <> ebsNone Or m_bDropDownSep Then
      If m_PictureAlign = epRightEdge Or m_PictureAlign = epRightOfCaption Then
         SetRect m_TextRect, 0, 0, lw - 24, lh
      Else
         SetRect m_TextRect, 0, 0, lw - 16, lh
      End If
   Else
      SetRect m_TextRect, 0, 0, lw - 8, lh
   End If

   ' --Calc rects for multiline
   If m_WindowsNT Then
      DrawTextW hdc, StrPtr(m_Caption), -1, m_TextRect, DT_CALCRECT Or DT_WORDBREAK Or IIf(m_bRTL, DT_RTLREADING, 0)
   Else
      DrawText hdc, m_Caption, -1, m_TextRect, DT_CALCRECT Or DT_WORDBREAK Or IIf(m_bRTL, DT_RTLREADING, 0)
   End If

   ' --Copy rect into temp var
   CopyRect lpRect, m_TextRect

   ' --Move the caption area according to Caption alignments
   Select Case m_CaptionAlign
      Case ecLeftAlign
         OffsetRect lpRect, 2, (lh - lpRect.Bottom) \ 2

      Case ecCenterAlign
         OffsetRect lpRect, (lw - lpRect.Right + PicW) \ 2, (lh - lpRect.Bottom) \ 2
         If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
            OffsetRect lpRect, -8, 0
         End If
         If m_PictureAlign = epBottomEdge Or m_PictureAlign = epBottomOfCaption Or m_PictureAlign = epTopOfCaption Or m_PictureAlign = epTopEdge Then
            OffsetRect lpRect, -(PicW \ 2), 0
         End If

      Case ecRightAlign
         OffsetRect lpRect, (lw - lpRect.Right - 4), (lh - lpRect.Bottom) \ 2

   End Select

   With lpRect

      If Not m_Picture Is Nothing Then
         Select Case m_PictureAlign
            Case epLeftEdge, epLeftOfCaption
               If m_CaptionAlign <> ecCenterAlign Then
                  If .Left < PicW + 4 Then
                     .Left = PicW + 4: .Right = .Right + PicW + 4
                  End If
               End If

            Case epRightEdge, epRightOfCaption
               If .Right > lw - PicW - 4 Then
                  .Right = lw - PicW - 4: .Left = .Left - PicW - 4
               End If
               If m_CaptionAlign = ecCenterAlign Then
                  OffsetRect lpRect, -12, 0
               End If

            Case epTopOfCaption, epTopEdge
               OffsetRect lpRect, 0, PicH \ 2

            Case epBottomOfCaption, epBottomEdge
               OffsetRect lpRect, 0, -PicH \ 2

            Case epBackGround
               If m_CaptionAlign = ecCenterAlign Then
                  OffsetRect lpRect, -16, 0
               End If
         End Select
      End If

      If m_CaptionAlign = ecRightAlign Then
         If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
            OffsetRect lpRect, -16, 0
         End If
      End If

      ' --For themed style, we are not able to draw borders
      ' --after drawing the caption. i mean the whole button is painted at once.
      If m_ButtonStyle = eWindowsTheme Then
         If .Left < 4 Then .Left = 4
         If .Right > ScaleWidth - 4 Then .Right = ScaleWidth - 4
         If .Top < 4 Then .Top = 4
         If .Bottom > ScaleHeight - 4 Then .Bottom = ScaleHeight - 4
      End If
   End With
   
   ' --Save the caption rect
   CopyRect m_TextRect, lpRect

   ' --Calculate Pictures positions once we have caption rects
   CalcPicRects
   
   ' --Calculate rects with the dropdown symbol
   If m_DropDownSymbol <> ebsNone Then
      ' --Drawing area for dropdown symbol  (the symbol is optional;)
      SetRect lpSignRect, lw - 15, lh / 2 - 7, lw, lh / 2 + 8
   End If

   If m_bDropDownSep Then
      If m_PictureAlign <> epRightEdge Or m_PictureAlign <> epRightOfCaption Then
         If m_TextRect.Right < ScaleWidth - 8 Then
            DrawLineApi lw - 16, 3, lw - 16, lh - 3, ShiftColor(GetPixel(hdc, 7, 7), -0.1)
            DrawLineApi lw - 15, 3, lw - 15, lh - 3, ShiftColor(GetPixel(hdc, 7, 7), 0.1)
         End If
      ElseIf m_PictureAlign = epRightEdge Or m_PictureAlign = epRightOfCaption Then
         DrawLineApi lw - 16, 3, lw - 16, lh - 3, ShiftColor(GetPixel(hdc, 7, 7), -0.1)
         DrawLineApi lw - 15, 3, lw - 15, lh - 3, ShiftColor(GetPixel(hdc, 7, 7), 0.1)
      End If
   End If

   ' --Some styles on down state donot change their text positions
   ' --See your XP and Vista buttons ;)
   If m_Buttonstate = eStateDown Then
      If m_ButtonStyle = e3DHover Or m_ButtonStyle = eFlat Or m_ButtonStyle = eFlatHover Or _
         m_ButtonStyle = eGelButton Or m_ButtonStyle = eOffice2003 _
         Or m_ButtonStyle = eXPToolbar Or m_ButtonStyle = eVistaToolbar Or m_ButtonStyle = eStandard Then
         OffsetRect m_TextRect, 1, 1
         OffsetRect m_PicRect, 1, 1
         OffsetRect lpSignRect, 1, 1
      End If
   End If
   
   ' --Draw Pictures
   If m_bPicPushOnHover And m_Buttonstate = eStateOver Then
      lShadowClr = TranslateColor(&HC0C0C0)
      DrawPicture m_PicRect, lShadowClr
      CopyRect pRect, m_PicRect
      OffsetRect pRect, -2, -2
      DrawPicture pRect
   Else
      DrawPicture m_PicRect
   End If

   If m_PictureShadow Then
      If Not (m_bPicPushOnHover And m_Buttonstate = eStateOver) Then
         DrawPicShadow
      End If
   End If
   
   ' --Text Effects
   If m_CaptionEffects <> eseNone Then
      DrawCaptionEffect
   End If

   ' --At Last, draw the Captions
   If m_bEnabled Then
      If m_Buttonstate = eStateOver Then
         DrawCaptionEx m_TextRect, TranslateColor(m_bColors.tForeColorOver), 0, 0
      Else
         DrawCaptionEx m_TextRect, TranslateColor(m_bColors.tForeColor), 0, 0
      End If
   Else
      DrawCaptionEx m_TextRect, GetSysColor(COLOR_GRAYTEXT), 0, 0
   End If

   If m_DropDownSymbol <> ebsNone Then

      If m_ButtonStyle = eStandard Or m_ButtonStyle = e3DHover Or m_ButtonStyle = eFlat Or m_ButtonStyle = eFlatHover Or m_ButtonStyle = eVistaToolbar Or m_ButtonStyle = eXPToolbar Then
         ' --move the symbol downwards for some button style on mouse down
         If m_Buttonstate = eStateDown Then
            OffsetRect lpSignRect, 1, 1
         End If
      End If
        
      If m_bEnabled Then
         UserControl.ForeColor = TranslateColor(m_bColors.tForeColor)
      Else
         UserControl.ForeColor = GetSysColor(COLOR_GRAYTEXT)
      End If
      DrawSymbol m_DropDownSymbol
   End If

End Sub

Private Sub CalcPicRects()

'****************************************************************************
' Calculate the rects for positioning pictures                              *
'****************************************************************************

   If m_Picture Is Nothing Then Exit Sub

   With m_PicRect

      If Trim(m_Caption) <> "" And m_PictureAlign <> epBackGround Then

         Select Case m_PictureAlign

            Case epLeftEdge
               .Left = 3
               .Top = (lh - PicH) \ 2
               If m_PicRect.Left < 0 Then
                  OffsetRect m_PicRect, PicW, 0
                  OffsetRect m_TextRect, PicW, 0
               End If

            Case epLeftOfCaption
               .Left = m_TextRect.Left - PicW - 4
               .Top = (lh - PicH) \ 2

            Case epRightEdge
               .Left = lw - PicW - 3
               .Top = (lh - PicH) \ 2
               ' --If picture overlaps text
               If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                  OffsetRect m_PicRect, -16, 0
               End If
               If .Left < m_TextRect.Right + 2 Then
                  .Left = m_TextRect.Right + 2
               End If

            Case epRightOfCaption
               .Left = m_TextRect.Right + 4
               .Top = (lh - PicH) \ 2
               If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                  OffsetRect m_PicRect, -16, 0
               End If
               ' --If picture overlaps text
               If .Left < m_TextRect.Right + 2 Then
                  .Left = m_TextRect.Right + 2
               End If

            Case epTopOfCaption
               .Left = (lw - PicW) \ 2
               .Top = m_TextRect.Top - PicH - 2
               If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                  OffsetRect m_PicRect, -8, 0
               End If

            Case epTopEdge
               .Left = (lw - PicW) \ 2
               .Top = 4
               If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                  OffsetRect m_PicRect, -8, 0
               End If

            Case epBottomOfCaption
               .Left = (lw - PicW) \ 2
               .Top = m_TextRect.Bottom + 2
               If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                  OffsetRect m_PicRect, -8, 0
               End If

            Case epBottomEdge
               .Left = (lw - PicW) \ 2
               .Top = lh - PicH - 4
               If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                  OffsetRect m_PicRect, -8, 0
               End If

         End Select
      Else
         .Left = (lw - PicW) \ 2
         .Top = (lh - PicH) \ 2
         If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
            OffsetRect m_PicRect, -8, 0
         End If
      End If

        ' --Set the height and width
        .Right = .Left + PicW
        .Bottom = .Top + PicH
       
   End With

End Sub

Private Sub DrawPicture(lpRect As RECT, Optional lBrushColor As Long = -1)

'****************************************************************************
' draw the picture by calling the TransBlt routines                         *
'****************************************************************************

   Dim tmpMaskColor     As Long

   ' --Draw picture
   If tmppic.Type = vbPicTypeIcon Then
      tmpMaskColor = TranslateColor(&HC0C0C0)
   Else
      tmpMaskColor = m_lMaskColor
   End If
   
   If Is32BitBMP(tmppic) Then
      TransBlt32 hdc, lpRect.Left, lpRect.Top, PicW, PicH, tmppic, lBrushColor
   Else
      TransBlt hdc, lpRect.Left, lpRect.Top, PicW, PicH, tmppic, tmpMaskColor, lBrushColor
   End If
   
End Sub

Private Sub DrawPicShadow()

'  Still not satisfied results for picture shadows

   Dim lShadowClr       As Long
   Dim lPixelClr        As Long
   Dim lpRect           As RECT

   If m_bPicPushOnHover And m_Buttonstate = eStateOver Then
      OffsetRect m_PicRect, -2, -2
   End If
     
   lShadowClr = BlendColors(TranslateColor(&H808080), TranslateColor(m_bColors.tBackColor))
   CopyRect lpRect, m_PicRect
    
   OffsetRect lpRect, 2, 2
   DrawPicture lpRect, ShiftColor(lShadowClr, 0.05)
   OffsetRect lpRect, -1, -1
   DrawPicture lpRect, ShiftColor(lShadowClr, -0.1)

   DrawPicture m_PicRect

End Sub

Private Sub DrawCaptionEffect()

'****************************************************************************
'* Draws the caption with/without unicode along with the special effects    *
'****************************************************************************

   Dim bColor           As Long                                  'BackColor

   bColor = TranslateColor(m_bColors.tBackColor)

   ' --Set new colors according to effects
   Select Case m_CaptionEffects
      Case eseEmbossed
         DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.14), -1, -1
      Case eseEngraved
         DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.14), 1, 1
      Case eseShadowed
         DrawCaptionEx m_TextRect, TranslateColor(&HC0C0C0), 1, 1
      Case eseOutline
         DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.1), 1, 1
         DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.1), 1, -1
         DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.1), -1, 1
         DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.1), -1, -1
      Case eseCover
         DrawCaptionEx m_TextRect, ShiftColor(bColor, -0.1), 1, 1
         DrawCaptionEx m_TextRect, ShiftColor(bColor, -0.1), 1, -1
         DrawCaptionEx m_TextRect, ShiftColor(bColor, -0.1), -1, 1
         DrawCaptionEx m_TextRect, ShiftColor(bColor, -0.1), -1, -1

   End Select

   If m_bEnabled Then
      DrawCaptionEx m_TextRect, TranslateColor(m_bColors.tForeColor), 0, 0
   Else
      DrawCaptionEx m_TextRect, GetSysColor(COLOR_GRAYTEXT), 0, 0
   End If

End Sub

Private Sub DrawCaptionEx(lpRect As RECT, lColor As Long, OffsetX As Long, OffsetY As Long)

   Dim tRect            As RECT
   Dim lOldForeColor    As Long

   ' --Get current forecolor
   lOldForeColor = GetTextColor(hdc)

   CopyRect tRect, lpRect
   OffsetRect tRect, OffsetX, OffsetY

   SetTextColor hdc, lColor

   If m_WindowsNT Then
      DrawTextW hdc, StrPtr(m_Caption), -1, tRect, DT_DRAWFLAG Or IIf(m_bRTL, DT_RTLREADING, 0)
   Else
      DrawText hdc, m_Caption, -1, tRect, DT_DRAWFLAG Or IIf(m_bRTL, DT_RTLREADING, 0)
   End If

   ' --Restore previous forecolor
   SetTextColor hdc, lOldForeColor

End Sub

Private Sub UncheckAllValues()

' --Many Thanks to Morgan Haueisen

  Dim objButton As Object
   ' Check all controls in parent
   For Each objButton In Parent.Controls
       ' Is it a jcbutton?
      If TypeOf objButton Is jcbutton Then
         ' Is the button in the same container?
         If objButton.Container.hWnd = UserControl.ContainerHwnd Then
            ' is the button type Option?
            If objButton.Mode = [ebmOptionButton] Then
               ' is it not this button
               If Not objButton.hWnd = UserControl.hWnd Then
                  objButton.Value = False
               End If
            End If
         End If
      End If
    Next objButton
    
End Sub

Private Sub SetAccessKey()

   Dim i                As Long

   UserControl.AccessKeys = vbNullString
   If Len(m_Caption) > 1 Then
      i = InStr(1, m_Caption, "&", vbTextCompare)
      If (i < Len(m_Caption)) And (i > 0) Then
         If Mid$(m_Caption, i + 1, 1) <> "&" Then
            AccessKeys = LCase$(Mid$(m_Caption, i + 1, 1))
         Else
            i = InStr(i + 2, m_Caption, "&", vbTextCompare)
            If Mid$(m_Caption, i + 1, 1) <> "&" Then
               AccessKeys = LCase$(Mid$(m_Caption, i + 1, 1))
            End If
         End If
      End If
   End If

End Sub

Private Sub DrawCorners(Color As Long)

'****************************************************************************
'* Draws four Corners of the button specified by Color                      *
'****************************************************************************

   lh = ScaleHeight
   lw = ScaleWidth

   SetPixel hdc, 1, 1, Color
   SetPixel hdc, 1, lh - 2, Color
   SetPixel hdc, lw - 2, 1, Color
   SetPixel hdc, lw - 2, lh - 2, Color

End Sub

Private Sub DrawStandardButton(ByVal vState As enumButtonStates)

   '****************************************************************************
   ' Draws  four different styles in one procedure                             *
   ' Makes reading the code difficult, but saves much space!! ;)               *
   '****************************************************************************

   Dim FocusRect        As RECT
   Dim tmpRect          As RECT

   lh = ScaleHeight
   lw = ScaleWidth
   SetRect m_ButtonRect, 0, 0, lw, lh

   If Not m_bEnabled Then
      ' --Draws raised edge border
      If m_ButtonStyle = eStandard Then
         DrawEdge hdc, m_ButtonRect, BDR_RAISED95, BF_RECT
      ElseIf m_ButtonStyle = eFlat Then
         DrawEdge hdc, m_ButtonRect, BDR_RAISEDINNER, BF_RECT
      End If
      DrawPicwithCaption
      Exit Sub
   End If

   If m_ButtonMode <> ebmCommandButton And m_bValue Then
      PaintRect ShiftColor(TranslateColor(m_bColors.tBackColor), 0.02), m_ButtonRect
      DrawPicwithCaption
      If m_ButtonStyle <> eFlatHover Then
         DrawEdge hdc, m_ButtonRect, BDR_SUNKEN95, BF_RECT
         If m_bShowFocus And m_bHasFocus And m_ButtonStyle = eStandard Then
            DrawRectangle 4, 4, lw - 7, lh - 7, TranslateColor(vbApplicationWorkspace)
         End If
      End If
      Exit Sub
   End If

   Select Case vState
      Case eStateNormal
         CreateRegion
         PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
         DrawPicwithCaption
         Select Case m_ButtonStyle
            Case eStandard
               DrawEdge hdc, m_ButtonRect, BDR_RAISED95, BF_RECT
            Case eFlat
               DrawEdge hdc, m_ButtonRect, BDR_RAISEDINNER, BF_RECT
         End Select
      Case eStateOver
         PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
         DrawPicwithCaption
         Select Case m_ButtonStyle
            Case eFlatHover, eFlat
               ' --Draws flat raised edge border
               DrawEdge hdc, m_ButtonRect, BDR_RAISEDINNER, BF_RECT
            Case Else
               ' --Draws 3d raised edge border
               DrawEdge hdc, m_ButtonRect, BDR_RAISED95, BF_RECT
         End Select

      Case eStateDown
         PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
         DrawPicwithCaption
         Select Case m_ButtonStyle
            Case eStandard
               DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(&H99A8AC)
               DrawRectangle 0, 0, lw, lh, TranslateColor(vbBlack)
            Case e3DHover
               DrawEdge hdc, m_ButtonRect, BDR_SUNKEN95, BF_RECT
            Case eFlatHover, eFlat
               ' --Draws flat pressed edge
               DrawRectangle 0, 0, lw, lh, TranslateColor(vbWhite)
               DrawRectangle 0, 0, lw + 1, lh + 1, TranslateColor(vbGrayText)
         End Select
   End Select

   ' --Button has focus but not downstate Or button is Default

   If m_bHasFocus Or m_bDefault Then
      On Error Resume Next
      If m_bShowFocus And Ambient.UserMode Then
         If m_ButtonStyle = e3DHover Or m_ButtonStyle = eStandard Then
            SetRect FocusRect, 4, 4, lw - 4, lh - 4
         Else
            SetRect FocusRect, 3, 3, lw - 3, lh - 3
         End If
         If m_bParentActive Then
            DrawFocusRect hdc, FocusRect
         End If
      End If
      If vState <> eStateDown And m_ButtonStyle = eStandard Then
         SetRect tmpRect, 0, 0, lw - 1, lh - 1
         DrawEdge hdc, tmpRect, BDR_RAISED95, BF_RECT
         DrawRectangle 0, 0, lw - 1, lh - 1, TranslateColor(vbApplicationWorkspace)
         DrawRectangle 0, 0, lw, lh, TranslateColor(vbBlack)
      End If
   End If

End Sub

Private Sub DrawXPToolbar(ByVal vState As enumButtonStates)

   Dim lpRect           As RECT
   Dim bColor           As Long

   lh = ScaleHeight
   lw = ScaleWidth
   bColor = TranslateColor(m_bColors.tBackColor)

   If vState = eStateDown Then
      m_bColors.tForeColor = TranslateColor(vbWhite)
   Else
      m_bColors.tForeColor = TranslateColor(vbButtonText)
   End If

   If m_ButtonMode <> ebmCommandButton And m_bValue Then
      If m_bIsDown Then vState = eStateDown
   End If

   If m_ButtonMode <> ebmCommandButton And m_bValue And vState <> eStateDown Then
      SetRect lpRect, 0, 0, lw, lh
      PaintRect TranslateColor(&HFEFEFE), lpRect
      m_bColors.tForeColor = TranslateColor(vbButtonText)
      DrawPicwithCaption
      DrawRectangle 0, 0, lw, lh, TranslateColor(&HAF987A)
      DrawCorners ShiftColor(TranslateColor(&HC1B3A0), -0.2)
      If vState = eStateOver Then
         DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&HEDF0F2)  'Right Line
         DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&HD8DEE4)   'Bottom
         DrawLineApi 1, lh - 3, lw - 1, lh - 3, TranslateColor(&HE8ECEF)  'Bottom
         DrawLineApi 1, lh - 4, lw - 1, lh - 4, TranslateColor(&HF8F9FA)   'Bottom
      End If
      ' --Necessary to redraw text & pictures 'coz we are painting usercontrol agaon
      Exit Sub
   End If

   Select Case vState
      Case eStateNormal
         CreateRegion
         PaintRect bColor, m_ButtonRect
         DrawPicwithCaption
      Case eStateOver
         DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(&HFDFEFE), TranslateColor(&HEEF4F4), gdVertical
         DrawGradientEx 0, lh / 2, lw, lh / 2, TranslateColor(&HEEF4F4), TranslateColor(&HEAF1F1), gdVertical
         DrawPicwithCaption
         DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&HE0E7EA) 'right line
         DrawLineApi lw - 3, 2, lw - 3, lh - 2, TranslateColor(&HEAF0F0)
         DrawLineApi 0, lh - 4, lw, lh - 4, TranslateColor(&HE5EDEE)    'Bottom
         DrawLineApi 0, lh - 3, lw, lh - 3, TranslateColor(&HD6E1E4)    'Bottom
         DrawLineApi 0, lh - 2, lw, lh - 2, TranslateColor(&HC6D2D7)    'Bottom
         DrawRectangle 0, 0, lw, lh, TranslateColor(&HC3CECE)
         DrawCorners ShiftColor(TranslateColor(&HC9D4D4), -0.05)
      Case eStateDown
         PaintRect TranslateColor(&HDDE4E5), m_ButtonRect                 'Paint with Darker color
         DrawPicwithCaption
         DrawLineApi 1, 1, lw - 2, 1, ShiftColor(TranslateColor(&HD1DADC), -0.02)          'Topmost Line
         DrawLineApi 1, 2, lw - 2, 2, ShiftColor(TranslateColor(&HDAE1E3), -0.02)          'A lighter top line
         DrawLineApi 1, lh - 3, lw - 2, lh - 3, ShiftColor(TranslateColor(&HDEE5E6), 0.02) 'Bottom Line
         DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(TranslateColor(&HE5EAEB), 0.02)
         DrawRectangle 0, 0, lw, lh, TranslateColor(&H929D9D)
         DrawCorners ShiftColor(TranslateColor(&HABB4B5), -0.2)
   End Select

End Sub

Private Sub DrawWinXPButton(ByVal vState As enumButtonStates)

'****************************************************************************
'* Windows XP Button                                                        *
'* Totally written from Scratch and coded by Me!!  hehe                     *
'****************************************************************************

   Dim lpRect           As RECT
   Dim bColor           As Long

   lh = ScaleHeight
   lw = ScaleWidth
   bColor = TranslateColor(m_bColors.tBackColor)
   SetRect m_ButtonRect, 0, 0, lw, lh

   If Not m_bEnabled Then
      CreateRegion
      PaintRect BlendColors(GetSysColor(COLOR_BTNFACE), ShiftColor(bColor, 0.1)), m_ButtonRect
      DrawPicwithCaption
      DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.1)
      DrawCorners ShiftColor(bColor, -0.1)
      Exit Sub
   End If

   Select Case vState

      Case eStateNormal
         CreateRegion
         Select Case m_lXPColor
            Case ecsBlue, ecsOliveGreen, ecsCustom
               ' --mimic the XP styles
               DrawGradientEx 0, 0, lw, lh, ShiftColor(bColor, 0.07), bColor, gdVertical
               DrawGradientEx 0, 0, lw, 4, ShiftColor(bColor, 0.1), ShiftColor(bColor, 0.08), gdVertical
               DrawPicwithCaption
               DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(bColor, -0.09) 'BottomMost line
               DrawLineApi 1, lh - 3, lw - 2, lh - 3, ShiftColor(bColor, -0.05) 'Bottom Line
               DrawLineApi 1, lh - 4, lw - 2, lh - 4, ShiftColor(bColor, -0.01) 'Bottom Line
               DrawLineApi lw - 2, 2, lw - 2, lh - 2, ShiftColor(bColor, -0.08) 'Right Line
               DrawLineApi 1, 1, 1, lh - 2, BlendColors(TranslateColor(vbWhite), (bColor)) 'Left Line
            Case ecsSilver
               ' --mimic the Silver XP style
               DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(bColor, 0.22), bColor, gdVertical
               DrawGradientEx 0, lh / 2, lw, lh / 2, ShiftColor(bColor, -0.01), ShiftColor(bColor, -0.15), gdVertical
               DrawPicwithCaption
               DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(vbWhite)  'Right Line
               DrawLineApi 1, 1, 1, lh - 2, TranslateColor(vbWhite)            'Left Line
         End Select

      Case eStateOver
         Select Case m_lXPColor
            Case ecsBlue, ecsOliveGreen, ecsCustom
               DrawGradientEx 0, 0, lw, lh, ShiftColor(bColor, 0.07), bColor, gdVertical
               DrawGradientEx 0, 0, lw, 4, ShiftColor(bColor, 0.1), ShiftColor(bColor, 0.08), gdVertical
               DrawPicwithCaption
            Case ecsSilver
               DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(bColor, 0.22), bColor, gdVertical
               DrawGradientEx 0, lh / 2, lw, lh / 2, ShiftColor(bColor, -0.01), ShiftColor(bColor, -0.15), gdVertical
               DrawPicwithCaption
         End Select
         ' --Draw the ORANGE border lines....
         DrawLineApi 1, 2, lw - 2, 2, TranslateColor(&H89D8FD)           'uppermost inner hover
         DrawLineApi 1, 1, lw - 2, 1, TranslateColor(&HCFF0FF)           'uppermost outer hover
         DrawLineApi 1, 1, 1, lh - 2, TranslateColor(&H49BDF9)           'Leftmost Line
         DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&H49BDF9) 'Rightmost Line
         DrawLineApi 2, 2, 2, lh - 3, TranslateColor(&H7AD2FC)           'Left Line
         DrawLineApi lw - 3, 3, lw - 3, lh - 3, TranslateColor(&H7AD2FC) 'Right Line
         DrawLineApi 2, lh - 3, lw - 2, lh - 3, TranslateColor(&H30B3F8) 'BottomMost Line
         DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&H97E5&)  'Bottom Line

      Case eStateDown
         Select Case m_lXPColor
            Case ecsBlue, ecsOliveGreen, ecsCustom
               PaintRect ShiftColor(bColor, -0.05), m_ButtonRect               'Paint with Darker color
               DrawPicwithCaption
               DrawLineApi 1, 1, lw - 2, 1, ShiftColor(bColor, -0.16)           'Topmost Line
               DrawLineApi 1, 2, lw - 2, 2, ShiftColor(bColor, -0.1)            'A lighter top line
               DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(bColor, 0.01)  'Bottom Line
               DrawLineApi 1, 1, 1, lh - 2, ShiftColor(bColor, -0.16)           'Leftmost Line
               DrawLineApi 2, 2, 2, lh - 2, ShiftColor(bColor, -0.1)            'Left1 Line
               DrawLineApi lw - 2, 2, lw - 2, lh - 2, ShiftColor(bColor, 0.04)  'Right Line
            Case ecsSilver
               DrawGradientEx 0, 0, lw, lh - 6, ShiftColor(bColor, -0.2), ShiftColor(bColor, 0.05), gdVertical
               DrawGradientEx 0, lh - 6, lw, lh - 1, ShiftColor(bColor, 0.08), TranslateColor(vbWhite), gdVertical
               DrawPicwithCaption
               DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite)
         End Select
   End Select

   If m_bParentActive Then
      If (m_bHasFocus Or m_bDefault) And (vState <> eStateDown And vState <> eStateOver) Then
         Select Case m_lXPColor
            Case ecsBlue, ecsCustom
               DrawLineApi 1, 2, lw - 2, 2, TranslateColor(&HF6D4BC)           'uppermost inner hover
               DrawLineApi 1, 1, lw - 2, 1, TranslateColor(&HFFE7CE)           'uppermost outer hover
               DrawLineApi 1, 1, 1, lh - 2, TranslateColor(&HE6AF8E)           'Leftmost Line
               DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&HE6AF8E) 'Rightmost Line
               DrawLineApi 2, 2, 2, lh - 3, TranslateColor(&HF4D1B8)           'Left Line
               DrawLineApi lw - 3, 3, lw - 3, lh - 3, TranslateColor(&HF4D1B8) 'Right Line
               DrawLineApi 2, lh - 3, lw - 2, lh - 3, TranslateColor(&HE4AD89) 'BottomMost Line
               DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&HEE8269) 'Bottom Line
            Case ecsOliveGreen
               DrawLineApi 1, 2, lw - 2, 2, TranslateColor(&H8FD1C2)           'uppermost inner hover
               DrawLineApi 1, 1, lw - 2, 1, TranslateColor(&H80CBB1)           'uppermost outer hover
               DrawLineApi 1, 1, 1, lh - 2, TranslateColor(&H68C8A0)           'Leftmost Line
               DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&H68C8A0) 'Rightmost Line
               DrawLineApi 2, 2, 2, lh - 3, TranslateColor(&H68C8A0)           'Left Line
               DrawLineApi lw - 3, 3, lw - 3, lh - 3, TranslateColor(&H68C8A0) 'Right Line
               DrawLineApi 2, lh - 3, lw - 2, lh - 3, TranslateColor(&H68C8A0) 'Bottom Line
               DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&H66A7A8) 'BottomMost Line
            Case ecsSilver
               DrawLineApi 1, 2, lw - 2, 2, TranslateColor(&HF6D4BC)           'uppermost inner hover
               DrawLineApi 1, 1, lw - 2, 1, TranslateColor(&HFFE7CE)           'uppermost outer hover
               DrawLineApi 1, 1, 1, lh - 2, TranslateColor(&HE6AF8E)           'Leftmost Line
               DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&HE6AF8E) 'Rightmost Line
               DrawLineApi 2, 2, 2, lh - 3, TranslateColor(vbWhite)            'Left Line
               DrawLineApi lw - 3, 3, lw - 3, lh - 3, TranslateColor(vbWhite) 'Right Line
               DrawLineApi 2, lh - 3, lw - 2, lh - 3, TranslateColor(&HE4AD89) 'BottomMost Line
               DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&HEE8269) 'Bottom Line
         End Select
      End If
   End If

   On Error Resume Next  'Some times error occurs that Client site not available
   If m_bParentActive Then 'I mean some times ;)
      If m_bShowFocus And m_bParentActive And (m_bHasFocus Or m_bDefault) Then  'show focusrect at runtime only
         SetRect lpRect, 2, 2, lw - 2, lh - 2     'I don't like this ugly focusrect!!
         DrawFocusRect hdc, lpRect
      End If
   End If

   Select Case m_lXPColor
      Case ecsBlue, ecsSilver, ecsCustom
         DrawRectangle 0, 0, lw, lh, TranslateColor(&H743C00)
         DrawCorners ShiftColor(TranslateColor(&H743C00), 0.3)
      Case ecsOliveGreen
         DrawRectangle 0, 0, lw, lh, RGB(55, 98, 6)
         DrawCorners ShiftColor(RGB(55, 98, 6), 0.3)
   End Select

End Sub

Private Sub DrawOfficeXP(ByVal vState As enumButtonStates)

   Dim lpRect           As RECT
   Dim pRect            As RECT
   Dim bColor           As Long
   Dim oColor           As Long
   Dim BorderColor      As Long

   lh = UserControl.ScaleHeight
   lw = UserControl.ScaleWidth

   bColor = TranslateColor(m_bColors.tBackColor)
   SetRect lpRect, 0, 0, lw, lh

   Select Case m_lXPColor
      Case ecsBlue
         oColor = TranslateColor(&HEED2C1)
         BorderColor = TranslateColor(&HC56A31)
      Case ecsSilver
         oColor = TranslateColor(&HE3DFE0)
         BorderColor = TranslateColor(&HBFB4B2)
      Case ecsOliveGreen
         oColor = TranslateColor(&HBAD6D4)
         BorderColor = TranslateColor(&H70A093)
      Case ecsCustom
         oColor = bColor
         BorderColor = ShiftColor(bColor, -0.12)
   End Select

   If m_ButtonMode <> ebmCommandButton And m_bValue Then
      PaintRect ShiftColor(oColor, -0.05), m_ButtonRect
      DrawRectangle 0, 0, lw, lh, BorderColor
      If m_bMouseInCtl Then
         PaintRect ShiftColor(oColor, -0.01), m_ButtonRect
         DrawRectangle 0, 0, lw, lh, BorderColor
      End If
      DrawPicwithCaption
      Exit Sub
   End If

   Select Case vState
      Case eStateNormal
         PaintRect bColor, lpRect
      Case eStateOver
         PaintRect ShiftColor(oColor, 0.03), lpRect
      Case eStateDown
         PaintRect ShiftColor(oColor, -0.08), lpRect
   End Select

   DrawPicwithCaption

   If m_Buttonstate <> eStateNormal Then
      DrawRectangle 0, 0, lw, lh, BorderColor
   End If

End Sub

Private Sub DrawInstallShieldButton(ByVal vState As enumButtonStates)

   '****************************************************************************
   '* I saw this style while installing JetAudio in my PC.                     *
   '* I liked it, so I implemented and gave it a name 'InstallShield'          *
   '* hehe .....
   '****************************************************************************

   Dim FocusRect        As RECT
   Dim lpRect           As RECT

   lh = ScaleHeight
   lw = ScaleWidth

   If Not m_bEnabled Then
      vState = eStateNormal                 'Simple draw normal state for Disabled
   End If

   Select Case vState
      Case eStateNormal
         CreateRegion
         SetRect m_ButtonRect, 0, 0, lw, lh 'Maybe have changed before!

         ' --Draw upper gradient
         DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(vbWhite), TranslateColor(m_bColors.tBackColor), gdVertical
         ' --Draw Bottom Gradient
         DrawGradientEx 0, lh / 2, lw, lh, TranslateColor(m_bColors.tBackColor), TranslateColor(m_bColors.tBackColor), gdVertical
         DrawPicwithCaption
         ' --Draw Inner White Border
         DrawRectangle 1, 1, lw - 2, lh, TranslateColor(vbWhite)
         ' --Draw Outer Rectangle
         DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.2)
         DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.25)
      Case eStateOver

         ' --Draw upper gradient
         DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(vbWhite), TranslateColor(m_bColors.tBackColor), gdVertical
         ' --Draw Bottom Gradient
         DrawGradientEx 0, lh / 2, lw, lh, TranslateColor(m_bColors.tBackColor), TranslateColor(m_bColors.tBackColor), gdVertical
         DrawPicwithCaption
         ' --Draw Inner White Border
         DrawRectangle 1, 1, lw - 2, lh, TranslateColor(vbWhite)
         ' --Draw Outer Rectangle
         DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.2)
         DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.25)
      Case eStateDown

         ' --draw upper gradient
         DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(vbWhite), ShiftColor(TranslateColor(m_bColors.tBackColor), -0.1), gdVertical
         ' --Draw Bottom Gradient
         DrawGradientEx 0, lh / 2, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.1), ShiftColor(TranslateColor(m_bColors.tBackColor), -0.05), gdVertical
         DrawPicwithCaption
         ' --Draw Inner White Border
         DrawRectangle 1, 1, lw - 2, lh, TranslateColor(vbWhite)
         ' --Draw Outer Rectangle
         DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.23)
         DrawCorners ShiftColor(TranslateColor(m_bColors.tBackColor), -0.1)
         DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.4)

   End Select

   DrawCorners ShiftColor(TranslateColor(m_bColors.tBackColor), 0.05)

   If m_bParentActive And m_bShowFocus And (m_bHasFocus Or m_bDefault) Then
      SetRect FocusRect, 3, 3, lw - 3, lh - 3
      DrawFocusRect hdc, FocusRect
   End If

End Sub

Private Sub DrawGelButton(ByVal vState As enumButtonStates)

   '****************************************************************************
   ' Draws a Gelbutton                                                         *
   '****************************************************************************

   Dim lpRect           As RECT                              'RECT to fill regions
   Dim bColor           As Long                              'Original backcolor

   lh = ScaleHeight
   lw = ScaleWidth

   bColor = TranslateColor(m_bColors.tBackColor)

   If Not m_bEnabled Then

      ' --Fill the button region with background color
      SetRect lpRect, 0, 0, lw, lh
      PaintRect bColor, lpRect

      ' --Make a shining Upper Light
      DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(bColor, TranslateColor(vbWhite)), 0.05), bColor, gdVertical
      DrawGradientEx 0, 6, lw, lh - 1, ShiftColor(bColor, -0.02), BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.08)), gdVertical

      DrawPicwithCaption
      DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.2)
      DrawCorners ShiftColor(bColor, -0.23)

      Exit Sub
   End If

   Select Case vState

      Case eStateNormal                                'Normal State

         CreateRegion

         ' --Fill the button region with background color
         SetRect lpRect, 0, 0, lw, lh
         PaintRect ShiftColor(bColor, -0.03), lpRect

         ' --Make a shining Upper Light
         DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(bColor, TranslateColor(vbWhite)), 0.1), bColor, gdVertical
         DrawGradientEx 0, 6, lw, lh - 1, ShiftColor(bColor, -0.05), BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.1)), gdVertical

         DrawPicwithCaption
         DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.33)

      Case eStateOver
         ' --Fill the button region with background color
         SetRect lpRect, 0, 0, lw, lh
         PaintRect ShiftColor(bColor, -0.03), lpRect

         ' --Make a shining Upper Light
         DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(bColor, TranslateColor(vbWhite)), 0.15), bColor, gdVertical
         DrawGradientEx 0, 6, lw, lh - 1, ShiftColor(bColor, -0.05), BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.2)), gdVertical

         DrawPicwithCaption
         DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.28)

      Case eStateDown

         ' --fill the button region with background color
         SetRect lpRect, 0, 0, lw, lh
         PaintRect ShiftColor(bColor, -0.03), lpRect

         ' --Make a shining Upper Light
         DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(bColor, TranslateColor(vbWhite)), 0.1), bColor, gdVertical
         DrawGradientEx 0, 6, lw, lh - 1, ShiftColor(bColor, -0.08), BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.05)), gdVertical

         DrawPicwithCaption
         DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.36)

   End Select

   DrawCorners ShiftColor(bColor, -0.36)

End Sub

Private Sub DrawVistaToolbarStyle(ByVal vState As enumButtonStates)

   Dim lpRect           As RECT
   Dim FocusRect        As RECT

   lh = ScaleHeight
   lw = ScaleWidth

   If Not m_bEnabled Then
      ' --Draw Disabled button
      PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
      DrawPicwithCaption
      DrawCorners TranslateColor(m_bColors.tBackColor)
      Exit Sub
   End If

   If vState = eStateNormal Then
      CreateRegion
      ' --Set the rect to fill back color
      SetRect lpRect, 0, 0, lw, lh
      ' --Simply fill the button with one color (No gradient effect here!!)
      PaintRect TranslateColor(m_bColors.tBackColor), lpRect
      DrawPicwithCaption
   ElseIf vState = eStateOver Then

      ' --Draws a gradient effect with the folowing colors
      DrawGradientEx 1, 1, lw - 2, lh - 2, TranslateColor(&HFDF9F1), TranslateColor(&HF8ECD0), gdVertical

      ' --Draws a gradient in half region to give a Light Effect
      DrawGradientEx 1, lh / 1.7, lw - 2, lh - 2, TranslateColor(&HF8ECD0), TranslateColor(&HF8ECD0), gdVertical

      DrawPicwithCaption

      ' --Draw outside borders
      DrawRectangle 0, 0, lw, lh, TranslateColor(&HCA9E61)
      DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite)

   ElseIf vState = eStateDown Then

      DrawGradientEx 1, 1, lw - 2, lh - 2, TranslateColor(&HF1DEB0), TranslateColor(&HF9F1DB), gdVertical

      DrawPicwithCaption
      ' --Draws outside borders
      DrawRectangle 0, 0, lw, lh, TranslateColor(&HCA9E61)
      DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite)

   End If

   If vState = eStateDown Or vState = eStateOver Then
      DrawCorners ShiftColor(TranslateColor(&HCA9E61), 0.3)
   End If

End Sub

Private Sub DrawVistaButton(ByVal vState As enumButtonStates)

   '*************************************************************************
   '* Draws a cool Vista Aero Style Button                                  *
   '* Use a light background color for best result                          *
   '*************************************************************************

   Dim lpRect           As RECT            'Used to set rect for drawing rectangles
   Dim Color1           As Long            'Shifted / Blended color
   Dim bColor           As Long            'Original back Color

   lh = ScaleHeight
   lw = ScaleWidth
   Color1 = ShiftColor(TranslateColor(m_bColors.tBackColor), 0.05)
   bColor = TranslateColor(m_bColors.tBackColor)

   If Not m_bEnabled Then
      ' --Draw the Disabled Button
      CreateRegion
      ' --Fill the button with disabled color
      SetRect lpRect, 0, 0, lw, lh
      PaintRect ShiftColor(bColor, 0.03), lpRect

      DrawPicwithCaption

      ' --Draws outside disabled color rectangle
      DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.25)
      DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(bColor, 0.25)
      DrawCorners ShiftColor(bColor, -0.03)
      Exit Sub
   End If

   Select Case vState

      Case eStateNormal

         CreateRegion

         ' --Draws a gradient in the full region
         DrawGradientEx 1, 1, lw - 1, lh, Color1, bColor, gdVertical

         ' --Draws a gradient in half region to give a glassy look
         DrawGradientEx 1, lh / 2, lw - 2, lh - 2, ShiftColor(bColor, -0.02), ShiftColor(bColor, -0.15), gdVertical

         DrawPicwithCaption

         ' --Draws border rectangle
         DrawRectangle 0, 0, lw, lh, TranslateColor(&H707070)   'outer
         DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite) 'inner

      Case eStateOver

         ' --Make gradient in the upper half region
         DrawGradientEx 1, 1, lw - 2, lh / 2, TranslateColor(&HFFF7E4), TranslateColor(&HFFF3DA), gdVertical

         ' --Draw gradient in half button downside to give a glass look
         DrawGradientEx 1, lh / 2, lw - 2, lh - 2, TranslateColor(&HFFE9C1), TranslateColor(&HFDE1AE), gdVertical

         ' --Draws left side gradient effects horizontal
         DrawGradientEx 1, 3, 5, lh / 2 - 2, TranslateColor(&HFFEECD), TranslateColor(&HFFF7E4), gdHorizontal    'Left
         DrawGradientEx 1, lh / 2, 5, lh - (lh / 2) - 1, TranslateColor(&HFAD68F), ShiftColor(TranslateColor(&HFDE1AC), 0.01), gdHorizontal   'Left

         ' --Draws right side gradient effects horizontal
         DrawGradientEx lw - 6, 3, 5, lh / 2 - 2, TranslateColor(&HFFF7E4), TranslateColor(&HFFEECD), gdHorizontal 'Right
         DrawGradientEx lw - 6, lh / 2, 5, lh - (lh / 2) - 1, ShiftColor(TranslateColor(&HFDE1AC), 0.01), TranslateColor(&HFAD68F), gdHorizontal 'Right

         DrawPicwithCaption
         ' --Draws border rectangle
         DrawRectangle 0, 0, lw, lh, TranslateColor(&HA77532)   'outer
         DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite) 'inner

      Case eStateDown

         ' --Draw a gradent in full region
         DrawGradientEx 1, 1, lw - 1, lh, TranslateColor(&HF6E4C2), TranslateColor(&HF6E4C2), gdVertical

         ' --Draw gradient in half button downside to give a glass look
         DrawGradientEx 1, lh / 2, lw - 2, lh - 2, TranslateColor(&HF0D29A), TranslateColor(&HF0D29A), gdVertical

         ' --Draws down rectangle

         DrawRectangle 0, 0, lw, lh, TranslateColor(&H5C411D)    '
         DrawLineApi 1, 1, lw - 1, 1, TranslateColor(&HB39C71)   '\Top Lines
         DrawLineApi 1, 2, lw - 1, 2, TranslateColor(&HD6C6A9)   '/
         DrawLineApi 1, 3, lw - 1, 3, TranslateColor(&HECD9B9)   '

         DrawLineApi 1, 1, 1, lh / 2 - 1, TranslateColor(&HCFB073)   'Left upper
         DrawLineApi 1, lh / 2, 1, lh - (lh / 2) - 1, TranslateColor(&HC5912B)   'Left Bottom

         ' --Draws left side gradient effects horizontal
         DrawGradientEx 1, 3, 5, lh / 2 - 2, ShiftColor(TranslateColor(&HE6C891), 0.02), ShiftColor(TranslateColor(&HF6E4C2), -0.01), gdHorizontal   'Left
         DrawGradientEx 1, lh / 2, 5, lh - (lh / 2) - 1, ShiftColor(TranslateColor(&HDCAB4E), 0.02), ShiftColor(TranslateColor(&HF0D29A), -0.01), gdHorizontal 'Left

         ' --Draws right side gradient effects horizontal
         DrawGradientEx lw - 6, 3, 5, lh / 2 - 2, ShiftColor(TranslateColor(&HF6E4C2), -0.01), ShiftColor(TranslateColor(&HE6C891), 0.02), gdHorizontal 'Right
         DrawGradientEx lw - 6, lh / 2, 5, lh - (lh / 2) - 1, ShiftColor(TranslateColor(&HF0D29A), -0.01), ShiftColor(TranslateColor(&HDCAB4E), 0.02), gdHorizontal 'Right
         DrawPicwithCaption

   End Select

   ' --Draw a focus rectangle if button has focus

   If m_bParentActive Then
      If (m_bHasFocus Or m_bDefault) And vState = eStateNormal Then
         ' --Draw darker outer rectangle
         DrawRectangle 0, 0, lw, lh, TranslateColor(&HA77532)
         ' --Draw light inner rectangle
         DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(&HFBD848)
      End If

      If (m_bShowFocus And m_bHasFocus) Then
         SetRect lpRect, 1.5, 1.5, lw - 2, lh - 2
         DrawFocusRect hdc, lpRect
      End If
   End If

   ' --Create four corners which will be common to all states
   DrawCorners TranslateColor(&HBE965F)

End Sub

Private Sub DrawOutlook2007(ByVal vState As enumButtonStates)

   Dim lpRect           As RECT
   Dim bColor           As Long

   lh = ScaleHeight
   lw = ScaleWidth
   bColor = TranslateColor(m_bColors.tBackColor)

   If m_ButtonMode <> ebmCommandButton And m_bValue Then
      DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&HA9D9FF), TranslateColor(&H6FC0FF), gdVertical
      DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H3FABFF), TranslateColor(&H75E1FF), gdVertical
      DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
      If m_bMouseInCtl Then
         DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&H58C1FF), TranslateColor(&H51AFFF), gdVertical
         DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H468FFF), TranslateColor(&H5FD3FF), gdVertical
         DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
      End If
      DrawPicwithCaption
      Exit Sub
   End If

   Select Case vState
      Case eStateNormal
         PaintRect bColor, m_ButtonRect
         DrawGradientEx 0, 0, lw, lh / 2.7, BlendColors(ShiftColor(bColor, 0.09), TranslateColor(vbWhite)), BlendColors(ShiftColor(bColor, 0.07), bColor), gdVertical
         DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), bColor, ShiftColor(bColor, 0.03), gdVertical
         DrawPicwithCaption
         DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
      Case eStateOver
         DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&HE1FFFF), TranslateColor(&HACEAFF), gdVertical
         DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H67D7FF), TranslateColor(&H99E4FF), gdVertical
         DrawPicwithCaption
         DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
      Case eStateDown
         DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&H58C1FF), TranslateColor(&H51AFFF), gdVertical
         DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H468FFF), TranslateColor(&H5FD3FF), gdVertical
         DrawPicwithCaption
         DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
   End Select

End Sub

Private Sub DrawOffice2003(ByVal vState As enumButtonStates)

   Dim lpRect           As RECT
   Dim bColor           As Long

   lh = UserControl.ScaleHeight
   lw = UserControl.ScaleWidth

   bColor = TranslateColor(m_bColors.tBackColor)
   SetRect m_ButtonRect, 0, 0, lw, lh

   If m_ButtonMode <> ebmCommandButton And m_bValue Then
      If m_bMouseInCtl Then
         DrawGradientEx 0, 0, lw, lh, TranslateColor(&H4E91FE), TranslateColor(&H8ED3FF), gdVertical
      Else
         DrawGradientEx 0, 0, lw, lh, TranslateColor(&H8CD5FF), TranslateColor(&H55ADFF), gdVertical
      End If
      DrawPicwithCaption
      DrawRectangle 0, 0, lw, lh, TranslateColor(&H800000)
      Exit Sub
   End If

   Select Case vState

      Case eStateNormal
         CreateRegion
         DrawGradientEx 0, 0, lw, lh / 2, BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.08)), bColor, gdVertical
         DrawGradientEx 0, lh / 2, lw, lh / 2 + 1, bColor, ShiftColor(bColor, -0.15), gdVertical
      Case eStateOver
         DrawGradientEx 0, 0, lw, lh, TranslateColor(&HCCF4FF), TranslateColor(&H91D0FF), gdVertical
      Case eStateDown
         DrawGradientEx 0, 0, lw, lh, TranslateColor(&H4E91FE), TranslateColor(&H8ED3FF), gdVertical
   End Select

   DrawPicwithCaption

   If m_Buttonstate <> eStateNormal Then
      DrawRectangle 0, 0, lw, lh, TranslateColor(&H800000)
   End If

End Sub

Private Sub WindowsThemeButton(ByVal vState As enumButtonStates)

   Dim tmpState         As Long

   UserControl.BackColor = GetSysColor(COLOR_BTNFACE)

   If Not m_bEnabled Then
      tmpState = 4
   DrawTheme "Button", 1, tmpState
   DrawPicwithCaption
   Exit Sub
   End If

   Select Case vState
      Case eStateNormal
         tmpState = 1
      Case eStateOver
         tmpState = 2
      Case eStateDown
         tmpState = 3
   End Select
   
   If m_Buttonstate = eStateNormal Then
      If (m_bHasFocus Or m_bDefault) And m_bParentActive Then
         tmpState = 5
      End If
   End If
      
   DrawTheme "Button", 1, tmpState
   DrawPicwithCaption

End Sub

Private Function DrawTheme(sClass As String, ByVal iPart As Long, ByVal vState As Long) As Boolean

   Dim hTheme           As Long
   Dim lResult          As Boolean
   Dim m_btnRect        As RECT
   Dim hRgn             As Long

   hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))
   If hTheme Then
      ' --Necessary for rounded buttons
      SetRect m_btnRect, m_ButtonRect.Left - 1, m_ButtonRect.Top - 1, m_ButtonRect.Right + 1, m_ButtonRect.Bottom + 2
      GetThemeBackgroundRegion hTheme, hdc, iPart, vState, m_btnRect, hRgn
      SetWindowRgn hWnd, hRgn, True
      ' --clean up
      DeleteObject hRgn
      ' --Draw the theme
      lResult = DrawThemeBackground(hTheme, hdc, iPart, vState, m_ButtonRect, m_ButtonRect)
      DrawTheme = lResult
   Else
      DrawTheme = False
   End If

End Function

Private Sub PaintRect(ByVal lColor As Long, lpRect As RECT)

   'Fills a region with specified color

   Dim hOldBrush        As Long
   Dim hBrush           As Long

   hBrush = CreateSolidBrush(lColor)
   hOldBrush = SelectObject(hdc, hBrush)

   FillRect hdc, lpRect, hBrush

   SelectObject hdc, hOldBrush
   DeleteObject hBrush

End Sub

Private Sub ShowPopupMenu()

   '* Shows a popupmenu
   '* Inspired from Noel Dacara's dcbutton

   Const TPM_BOTTOMALIGN As Long = &H20&

   Dim Menu             As VB.Menu
   Dim Align            As enumMenuAlign
   Dim Flags            As Long
   Dim DefaultMenu      As VB.Menu

   Dim x                As Long
   Dim Y                As Long

   Set Menu = DropDownMenu
   Align = MenuAlign
   Flags = MenuFlags
   Set DefaultMenu = DefaultMenu

   lh = ScaleHeight: lw = ScaleWidth

   m_bPopupInit = True

   ' --Set the drop down menu position
   Select Case Align
      Case edaBottom
         Y = lh

      Case edaLeft, edaBottomLeft
         MenuFlags = MenuFlags Or vbPopupMenuRightAlign
         If (MenuAlign = edaBottomLeft) Then
            Y = lh
         End If

      Case edaRight, edaBottomRight
         x = lw
         If (MenuAlign = edaBottomRight) Then
            Y = lh
         End If

      Case edaTop, edaTopRight, edaTopLeft
         MenuFlags = TPM_BOTTOMALIGN
         If (MenuAlign = edaTopRight) Then
            x = lw
         ElseIf (MenuAlign = edaTopLeft) Then
            MenuFlags = MenuFlags Or vbPopupMenuRightAlign
         End If

      Case Else
         m_bPopupInit = False

   End Select

   If (m_bPopupInit) Then

      ' /--Show the dropdown menu
      If (DefaultMenu Is Nothing) Then
         UserControl.PopupMenu DropDownMenu, MenuFlags, x, Y
      Else
         UserControl.PopupMenu DropDownMenu, MenuFlags, x, Y, DefaultMenu
      End If

      Dim lpPoint          As POINT
      GetCursorPos lpPoint

      If (WindowFromPoint(lpPoint.x, lpPoint.Y) = UserControl.hWnd) Then
         m_bPopupShown = True
      Else
         m_bIsDown = False
         m_bMouseInCtl = False
         m_bIsSpaceBarDown = False
         m_Buttonstate = eStateNormal
         m_bPopupShown = False
         m_bPopupInit = False
         RedrawButton
      End If
   End If

End Sub

Private Function ShiftColor(Color As Long, PercentInDecimal As Single) As Long

   '****************************************************************************
   '* This routine shifts a color value specified by PercentInDecimal          *
   '* Function inspired from DCbutton                                          *
   '* All Credits goes to Noel Dacara                                          *
   '* A Littlebit modified by me                                               *
   '****************************************************************************

   Dim r                As Long
   Dim g                As Long
   Dim B                As Long

   '  Add or remove a certain color quantity by how many percent.

   r = Color And 255
   g = (Color \ 256) And 255
   B = (Color \ 65536) And 255

   r = r + PercentInDecimal * 255       ' Percent should already
   g = g + PercentInDecimal * 255       ' be translated.
   B = B + PercentInDecimal * 255       ' Ex. 50% -> 50 / 100 = 0.5

   '  When overflow occurs, ....
   If (PercentInDecimal > 0) Then       ' RGB values must be between 0-255 only
      If (r > 255) Then r = 255
      If (g > 255) Then g = 255
      If (B > 255) Then B = 255
   Else
      If (r < 0) Then r = 0
      If (g < 0) Then g = 0
      If (B < 0) Then B = 0
   End If

   ShiftColor = r + 256& * g + 65536 * B ' Return shifted color value

End Function

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

   If m_bEnabled Then                           'Disabled?? get out!!
      If m_bIsSpaceBarDown Then
         m_bIsSpaceBarDown = False
         m_bIsDown = False
      End If
      If m_ButtonMode = ebmCheckBox Then       'Checkbox Mode?
         If KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub 'Checkboxes dont repond to Enter/Escape'
         m_bValue = Not m_bValue             'Change Value (Checked/Unchecked)
         If Not m_bValue Then                'If value unchecked then
            m_Buttonstate = eStateNormal     'Normal State
         End If
         RedrawButton
      ElseIf m_ButtonMode = ebmOptionButton Then
         If KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub 'Checkboxes dont repond to Enter/Escape'
         UncheckAllValues
         m_bValue = True
         RedrawButton
      End If
      DoEvents                               'To remove focus from other button and Do events before click event
      RaiseEvent Click                       'Now Raiseevent
   End If

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

   m_bDefault = Ambient.DisplayAsDefault
   If PropertyName = "DisplayAsDefault" Then
      RedrawButton
   End If

   If PropertyName = "BackColor" Then
      RedrawButton
   End If

End Sub

Private Sub UserControl_DblClick()

   If m_bHandPointer Then
      SetCursor m_lCursor
   End If

   If m_lDownButton = 1 Then                    'React to only Left button

      SetCapture (hWnd)                         'Preserve Hwnd on DoubleClick
      If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
      RedrawButton
      UserControl_MouseDown m_lDownButton, m_lDShift, m_lDX, m_lDY
      If Not m_bPopupEnabled Then
         RaiseEvent DblClick
      Else
         If Not m_bPopupShown Then
            ShowPopupMenu
         End If
      End If
   End If

End Sub

Private Sub UserControl_GotFocus()

   m_bHasFocus = True

End Sub


Private Sub UserControl_Hide()

   UserControl.Extender.ToolTipText = m_sTooltipText

End Sub

Private Sub UserControl_Initialize()

   Dim i                As Long
   Dim OS               As OSVERSIONINFO

   'Prebuid Lighten/Darken arrays
   For i = 0 To 255
      aLighten(i) = Lighten(i)
      aDarken(i) = Darken(i)
   Next
   
   ' --Get the operating system version for text drawing purposes.
   m_hMode = LoadLibraryA("shell32.dll")
   OS.dwOSVersionInfoSize = Len(OS)
   GetVersionEx OS
   m_WindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    
End Sub

Private Sub UserControl_InitProperties()

   'Initialize Properties for User Control
   'Called on designtime everytime a control is added

   m_ButtonStyle = eWindowsTheme   'As all the commercial buttons initialize with this them, ;)
   m_bShowFocus = True
   m_bEnabled = True
   m_Caption = Ambient.DisplayName
   UserControl.FontName = "Tahoma"
   Set mFont = UserControl.Font
   mFont_FontChanged ("")
   m_PictureOpacity = 255
   m_PicOpacityOnOver = 255
   m_PictureAlign = epLeftOfCaption
   m_bUseMaskColor = True
   m_lMaskColor = &HE0E0E0
   m_CaptionAlign = ecCenterAlign
   lh = UserControl.ScaleHeight
   lw = UserControl.ScaleWidth
   InitThemeColors
   SetThemeColors
   Refresh

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    Case 13                                    'Enter Key
        RaiseEvent Click
    Case 37, 38                                'Left and Up Arrows
        SendKeys "+{TAB}"                      'Button should transfer focus to other ctl
    Case 39, 40                                'Right and Down Arrows
        SendKeys "{TAB}"                       'Button should transfer focus to other ctl
    Case 32                                    'SpaceBar held down
        If Shift = 4 Then Exit Sub             'System Menu Should pop up
        If Not m_bIsDown Then
            m_bIsSpaceBarDown = True           'Set space bar as pressed
            
            If (m_ButtonMode = ebmCheckBox) Then 'Is CheckBoxMode??
                m_bValue = Not m_bValue         'Toggle Check Value
            ElseIf m_ButtonMode = ebmOptionButton Then
                UncheckAllValues                'Option Button Mode
                m_bValue = True                 'Pressed button Checked
            End If
            
            If m_Buttonstate <> eStateDown Then
                m_Buttonstate = eStateDown 'Button state should be down
                RedrawButton
            End If
        Else
            If m_bMouseInCtl Then
                If m_Buttonstate <> eStateDown Then
                    m_Buttonstate = eStateDown
                    RedrawButton
                End If
            Else
                If m_Buttonstate <> eStateNormal Then
                    m_Buttonstate = eStateNormal  'jump button from
                    RedrawButton                  'downstate - normal state
                End If                            'if mouse button is pressed
            End If
        End If                                'when spacebar being held

        If (Not GetCapture = UserControl.hWnd) Then
            ReleaseCapture
            SetCapture UserControl.hWnd     'No other processing until spacebar is released
        End If                              'Thanks to APIGuide
    
    Case Else
        If m_bIsSpaceBarDown Then
            m_bIsSpaceBarDown = False
            m_Buttonstate = eStateNormal
            RedrawButton
        End If
    End Select
        
    RaiseEvent KeyDown(KeyCode, Shift)


End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    
    ' --Simply raise the event =)
    RaiseEvent KeyPress(KeyAscii)
    
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeySpace Then
        
        ReleaseCapture                          'Now you can process further
                                                'as the spacebar is released
        If m_bMouseInCtl And m_bIsDown Then
            If m_Buttonstate <> eStateDown Then
                m_Buttonstate = eStateDown
                RedrawButton
            End If
        ElseIf m_bMouseInCtl Then                'If spacebar released over ctl
            If m_Buttonstate <> eStateOver Then
                m_Buttonstate = eStateOver 'Draw Hover State
                RedrawButton
            End If
            If Not m_bIsDown And m_bIsSpaceBarDown Then
                RaiseEvent Click
            End If
        Else                                         'If Spacebar released outside ctl
            If m_Buttonstate <> eStateNormal Then
                m_Buttonstate = eStateNormal
                RedrawButton
            End If
            If Not m_bIsDown And m_bIsSpaceBarDown Then
                RaiseEvent Click
            End If
        End If

        RaiseEvent KeyUp(KeyCode, Shift)
        m_bIsSpaceBarDown = False
        m_bIsDown = False
    End If

End Sub

Private Sub UserControl_LostFocus()

   m_bHasFocus = False                                 'No focus
   m_bIsDown = False                                   'No down state
   m_bIsSpaceBarDown = False                           'No spacebar held
   If Not m_bParentActive Then
      If m_Buttonstate <> eStateNormal Then
          m_Buttonstate = eStateNormal
      End If
   ElseIf m_bMouseInCtl Then
      If m_Buttonstate <> eStateOver Then
         m_Buttonstate = eStateOver
      End If
   Else
      If m_Buttonstate <> eStateNormal Then
         m_Buttonstate = eStateNormal
      End If
   End If
   RedrawButton

   If m_bDefault Then                                  'If default button,
      RedrawButton                                    'Show Focus
   End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

   m_lDownButton = Button                       'Button pressed for Dblclick
   m_lDX = x
   m_lDY = Y
   m_lDShift = Shift

   ' --Set HandPointer if any!
   If m_bHandPointer Then
      SetCursor m_lCursor
   End If

   If Button = vbLeftButton Or m_bPopupShown Then
      m_bHasFocus = True
      m_bIsDown = True

      If (Not m_bIsSpaceBarDown) Then
         If m_Buttonstate <> eStateDown Then
            m_Buttonstate = eStateDown
            RedrawButton
         End If
      End If

      If Not m_bPopupEnabled Then
         RaiseEvent MouseDown(Button, Shift, x, Y)
      Else
         If Not m_bPopupShown Then
            ShowPopupMenu
         End If
      End If
   End If

End Sub

Private Sub CreateToolTip()

'****************************************************************************
'* A very nice and flexible sub to create balloon tool tips
'* Author :- Fred.CPP
'* Added as requested by many users
'* Modified by me to support unicode
'* Thanks Alfredo ;)
'****************************************************************************

   Dim lpRect           As RECT
   Dim lWinStyle        As Long
   Dim lPtr             As Long
   Dim ttip             As TOOLINFO
   Dim ttipW            As TOOLINFOW
   Const CS_DROPSHADOW     As Long = &H20000
   Const GCL_STYLE         As Long = (-26)

   ' --Dont show tooltips if disabled
   If (Not m_bEnabled) Or m_bPopupShown Or m_Buttonstate = eStateDown Then Exit Sub

   ' --Destroy any previous tooltip
   If m_lttHwnd <> 0 Then
      DestroyWindow m_lttHwnd
   End If

   lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX

   ''create baloon style if desired
   If m_lTooltipType = TooltipBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
    
   If m_bttRTL Then
      m_lttHwnd = CreateWindowEx(WS_EX_LAYOUTRTL, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hWnd, 0&, App.hInstance, 0&)
   Else
      m_lttHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hWnd, 0&, App.hInstance, 0&)
   End If
   
   SetClassLong m_lttHwnd, GCL_STYLE, GetClassLong(m_lttHwnd, GCL_STYLE) Or CS_DROPSHADOW

   'make our tooltip window a topmost window
   ' This is creating some problems as noted by K-Zero
   'SetWindowPos m_lttHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE

   ''get the rect of the parent control
   GetClientRect UserControl.hWnd, lpRect

   If m_WindowsNT Then
      ' --set our tooltip info structure  for UNICODE SUPPORT >> WinNT
      With ttipW
         ' --if we want it centered, then set that flag
         If m_lttCentered Then
            .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
         Else
            .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
         End If

         ' --set the hwnd prop to our parent control's hwnd
         .lHwnd = UserControl.hWnd
         .lId = hWnd
         .lSize = Len(ttipW)
         .hInstance = App.hInstance
         .lpStrW = StrPtr(m_sTooltipText)
         .lpRect = lpRect
      End With
      ' --add the tooltip structure
      SendMessage m_lttHwnd, TTM_ADDTOOLW, 0&, ttipW
   Else
      ' --set our tooltip info structure for << WinNT
      With ttip
         ''if we want it centered, then set that flag
         If m_lttCentered Then
            .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
         Else
            .lFlags = TTF_SUBCLASS
         End If

         ' --set the hwnd prop to our parent control's hwnd
         .lHwnd = UserControl.hWnd
         .lId = hWnd
         .lSize = Len(ttip)
         .hInstance = App.hInstance
         .lpStr = m_sTooltipText
         .lpRect = lpRect
      End With
      ' --add the tooltip structure
      SendMessage m_lttHwnd, TTM_ADDTOOLA, 0&, ttip
   End If

   'if we want a title or we want an icon
   If m_sTooltiptitle <> vbNullString Or m_lToolTipIcon <> TTNoIcon Then
      If m_WindowsNT Then
         lPtr = StrPtr(m_sTooltiptitle)
         If lPtr Then
            SendMessage m_lttHwnd, TTM_SETTITLEW, m_lToolTipIcon, ByVal lPtr
         End If
      Else
         SendMessage m_lttHwnd, TTM_SETTITLE, CLng(m_lToolTipIcon), ByVal m_sTooltiptitle
      End If

   End If
   SendMessage m_lttHwnd, TTM_SETMAXTIPWIDTH, 0, 240 'for Multiline capability
   If m_lttBackColor <> Empty Then
      SendMessage m_lttHwnd, TTM_SETTIPBKCOLOR, TranslateColor(m_lttBackColor), 0&
   End If

End Sub

Private Sub InitThemeColors()

   Select Case m_ButtonStyle
      Case eStandard, eFlat, eVistaToolbar, eXPToolbar, eOfficeXP, eWindowsXP, eOutlook2007, eGelButton
         m_lXPColor = ecsBlue
      Case eInstallShield, eVistaAero
         m_lXPColor = ecsSilver
   End Select

End Sub

Private Sub SetThemeColors()

   'Sets a style colors to default colors when button initialized
   'or whenever you change the style of Button

   With m_bColors

      Select Case m_ButtonStyle

         Case eStandard, eFlat, eVistaToolbar, e3DHover, eFlatHover, eXPToolbar, eOfficeXP
            .tBackColor = GetSysColor(COLOR_BTNFACE)
         Case eWindowsXP
            Select Case m_lXPColor
               Case ecsBlue
                  .tBackColor = TranslateColor(&HE7EBEC)
               Case ecsOliveGreen
                  .tBackColor = TranslateColor(&HDBEEF3)
               Case ecsSilver
                  .tBackColor = TranslateColor(&HFCF1F0)
            End Select
         Case eOutlook2007, eGelButton
            Select Case m_lXPColor
               Case ecsBlue
                  .tBackColor = TranslateColor(&HFFD1AD)
               Case ecsOliveGreen
                  .tBackColor = TranslateColor(&HBAD6D4)
               Case ecsSilver
                  .tBackColor = TranslateColor(&HE3DFE0)
            End Select
            .tForeColor = TranslateColor(&H8B4215)
         Case eVistaAero
            Select Case m_lXPColor
               Case ecsBlue
                  .tBackColor = TranslateColor(&HFDECE0)
               Case ecsOliveGreen
                  .tBackColor = TranslateColor(&HDEEDE8)
               Case ecsSilver
                  .tBackColor = ShiftColor(TranslateColor(&HD4D4D4), 0.06)
            End Select
         Case eInstallShield
            Select Case m_lXPColor
               Case ecsBlue
                  .tBackColor = TranslateColor(&HFFD1AD)
               Case ecsOliveGreen
                  .tBackColor = TranslateColor(&HBAD6D4)
               Case ecsSilver
                  .tBackColor = TranslateColor(&HE1D6D5)
            End Select
         Case eOffice2003
            Select Case m_lXPColor
               Case ecsBlue
                  .tBackColor = TranslateColor(&HFCE1CA)
               Case ecsOliveGreen
                  .tBackColor = TranslateColor(&HBAD6D4)
               Case ecsSilver
                  .tBackColor = ShiftColor(TranslateColor(&HBA9EA0), 0.15)
            End Select
      End Select

      .tForeColor = TranslateColor(vbButtonText)
      If m_ButtonStyle = eFlat Or m_ButtonStyle = eInstallShield Or m_ButtonStyle = eStandard Then
         m_bShowFocus = True
      Else
         m_bShowFocus = False
      End If
      
      If m_ButtonStyle = eOfficeXP Then
         m_bPicPushOnHover = True
      End If

   End With

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

   Dim lp               As POINT
   
   GetCursorPos lp
   ' --Set hand pointer if any!
   If m_bHandPointer Then
      SetCursor m_lCursor
   End If
   
   If Not (WindowFromPoint(lp.x, lp.Y) = UserControl.hWnd) Then
      ' --Mouse yet not entered in the control
      m_bMouseInCtl = False
   Else
      m_bMouseInCtl = True
      ' --Check when the Mouse leaves the control
      TrackMouseLeave hWnd
      ' --Raise a MouseEnter event(it's Same as mouseMove)
      RaiseEvent MouseEnter
   End If
   
   ' --Proceed only if spacebar is not pressed
   If m_bIsSpaceBarDown Then Exit Sub
   
   ' --We are inside button
   If m_bMouseInCtl Then
   
      ' --Mouse button is pressed down
      If m_bIsDown Then
         If m_Buttonstate <> eStateDown Then
            m_Buttonstate = eStateDown
            RedrawButton
         End If
      Else
         ' --Button should be in hot state if user leaves the button
         ' --with mouse button pressed
         If m_Buttonstate <> eStateOver Then
            m_Buttonstate = eStateOver
            RedrawButton
            ' --Create Tooltip Here
            If m_Buttonstate <> eStateDown Then
               CreateToolTip
            End If
         End If
      End If
   
   Else
      If m_Buttonstate <> eStateNormal Then
         m_Buttonstate = eStateNormal
         RedrawButton
      End If
   End If

   'RaiseEvent MouseMove(Button, Shift, x, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

   If m_bHandPointer Then
      SetCursor m_lCursor
   End If

   ' --Popupmenu enabled
   If m_bPopupEnabled Then
      m_bIsDown = False
      m_bPopupShown = False
      m_Buttonstate = eStateNormal
      RedrawButton
      Exit Sub
   End If

   ' --React only to Left mouse button
   If Button = vbLeftButton Then
      '--Button released
      m_bIsDown = False
      ' --If button released in button area
      If (x > 0 And Y > 0) And (x < ScaleWidth And Y < ScaleHeight) Then

         ' --If check box mode
         If m_ButtonMode = ebmCheckBox Then
            m_bValue = Not m_bValue
            RedrawButton
            ' --If option button mode
         ElseIf m_ButtonMode = ebmOptionButton Then
            UncheckAllValues
            m_bValue = True
         End If

         ' --redraw Normal State
         m_Buttonstate = eStateNormal
         RedrawButton
         RaiseEvent Click
      End If
   End If

   RaiseEvent MouseUp(Button, Shift, x, Y)

End Sub

Private Sub UserControl_Resize()

   ' --At least, a checkbox will also need this much of size!!!!
   If Height < 220 Then Height = 220
   If Width < 220 Then Width = 220

   ' --On resize, create button region again
   CreateRegion
   RedrawButton                'then redraw

End Sub

Private Sub UserControl_Paint()

   ' --this routine typically called by Windows when another window covering
   '   this button is removed, or when the parent is moved/minimized/etc.

   RedrawButton

End Sub

'Load property values from storage

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      m_ButtonStyle = .ReadProperty("ButtonStyle", eFlat)
      m_bShowFocus = .ReadProperty("ShowFocusRect", False)
      Set mFont = .ReadProperty("Font", Ambient.Font)
      Set UserControl.Font = mFont
      m_bColors.tBackColor = .ReadProperty("BackColor", GetSysColor(COLOR_BTNFACE))
      m_bEnabled = .ReadProperty("Enabled", True)
      m_Caption = .ReadProperty("Caption", "jcbutton")
      m_bValue = .ReadProperty("Value", False)
      UserControl.MousePointer = .ReadProperty("MousePointer", 0)
      m_bHandPointer = .ReadProperty("HandPointer", False)
      Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
      Set m_Picture = .ReadProperty("PictureNormal", Nothing)
      Set m_PictureHot = .ReadProperty("PictureHot", Nothing)
      Set m_PictureDown = .ReadProperty("PictureDown", Nothing)
      m_PictureShadow = .ReadProperty("PictureShadow", False)
      m_PictureOpacity = .ReadProperty("PictureOpacity", 255)
      m_PicOpacityOnOver = .ReadProperty("PictureOpacityOnOver", 255)
      m_PicDisabledMode = .ReadProperty("DisabledPictureMode", edpBlended)
      m_bPicPushOnHover = .ReadProperty("PicturePushOnHover", False)
      m_PicEffectonOver = .ReadProperty("PictureEffectOnOver", epeLighter)
      m_PicEffectonDown = .ReadProperty("PictureEffectOnDown", epeDarker)
      m_lMaskColor = .ReadProperty("MaskColor", &HE0E0E0)
      m_bUseMaskColor = .ReadProperty("UseMaskColor", True)
      m_CaptionEffects = .ReadProperty("CaptionEffects", eseNone)
      m_ButtonMode = .ReadProperty("Mode", ebmCommandButton)
      m_PictureAlign = .ReadProperty("PictureAlign", epLeftOfCaption)
      m_CaptionAlign = .ReadProperty("CaptionAlign", ecCenterAlign)
      m_bColors.tForeColor = .ReadProperty("ForeColor", TranslateColor(vbButtonText))
      m_bColors.tForeColorOver = .ReadProperty("ForeColorHover", TranslateColor(vbButtonText))
      UserControl.ForeColor = m_bColors.tForeColor
      m_bDropDownSep = .ReadProperty("DropDownSeparator", False)
      m_sTooltiptitle = .ReadProperty("TooltipTitle", vbNullString)
      m_sTooltipText = .ReadProperty("ToolTip", vbNullString)
      m_lToolTipIcon = .ReadProperty("TooltipIcon", TTNoIcon)
      m_lTooltipType = .ReadProperty("TooltipType", TooltipStandard)
      m_lttBackColor = .ReadProperty("TooltipBackColor", TranslateColor(vbInfoBackground))
      m_bRTL = .ReadProperty("RightToLeft", False)
      m_bttRTL = .ReadProperty("RightToLeft", False)
      m_DropDownSymbol = .ReadProperty("DropDownSymbol", ebsNone)
      m_lXPColor = .ReadProperty("ColorScheme", ecsBlue)
      UserControl.Enabled = m_bEnabled
      SetAccessKey
      lh = UserControl.ScaleHeight
      lw = UserControl.ScaleWidth
      m_lParenthWnd = UserControl.Parent.hWnd
   End With

   If m_bHandPointer Then
      m_lCursor = LoadCursor(0, IDC_HAND)     'Load System Hand pointer
      m_bHandPointer = (Not m_lCursor = 0)
   End If

   UserControl_Resize

On Error GoTo h:
   If Ambient.UserMode Then                                                              'If we're not in design mode
      TrackUser32 = IsFunctionSupported("TrackMouseEvent", "User32")

      If Not TrackUser32 Then IsFunctionSupported "_TrackMouseEvent", "ComCtl32"

         'OS supports mouse leave so subclass for it
         With UserControl
            'Start subclassing the UserControl
            Subclass_Initialize .hWnd
            Subclass_Initialize m_lParenthWnd
            Subclass_AddMsg .hWnd, WM_MOUSELEAVE, MSG_AFTER
            Subclass_AddMsg .hWnd, WM_THEMECHANGED, MSG_AFTER
            If IsThemed Then
               Subclass_AddMsg .hWnd, WM_SYSCOLORCHANGE, MSG_AFTER
            End If
            On Error Resume Next
            If UserControl.Parent.MDIChild Then
               Call Subclass_AddMsg(m_lParenthWnd, WM_NCACTIVATE, MSG_AFTER)
            Else
               Call Subclass_AddMsg(m_lParenthWnd, WM_ACTIVATE, MSG_AFTER)
            End If
         End With

   End If

h:

End Sub

Private Sub UserControl_Show()

   UserControl.Extender.ToolTipText = vbNullString
   
End Sub

'A nice place to stop subclasser

Private Sub UserControl_Terminate()

On Error GoTo Crash:
   If m_lButtonRgn Then DeleteObject m_lButtonRgn      'Delete button region
   Set mFont = Nothing                                 'Clean up Font (StdFont)
   FreeLibrary m_hMode
   UnsetPopupMenu
   If Ambient.UserMode Then
      Subclass_Terminate
      Subclass_Terminate
   End If
Crash:

End Sub

'Write property values to storage

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "ButtonStyle", m_ButtonStyle, eFlat
      .WriteProperty "ShowFocusRect", m_bShowFocus, False
      .WriteProperty "Enabled", m_bEnabled, True
      .WriteProperty "Font", mFont, Ambient.Font
      .WriteProperty "BackColor", m_bColors.tBackColor, GetSysColor(COLOR_BTNFACE)
      .WriteProperty "Caption", m_Caption, "jcbutton1"
      .WriteProperty "ForeColor", m_bColors.tForeColor, TranslateColor(vbButtonText)
      .WriteProperty "ForeColorHover", m_bColors.tForeColorOver, TranslateColor(vbButtonText)
      .WriteProperty "Mode", m_ButtonMode, ebmCommandButton
      .WriteProperty "Value", m_bValue, False
      .WriteProperty "MousePointer", UserControl.MousePointer, 0
      .WriteProperty "HandPointer", m_bHandPointer, False
      .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
      .WriteProperty "PictureNormal", m_Picture, Nothing
      .WriteProperty "PictureHot", m_PictureHot, Nothing
      .WriteProperty "PictureDown", m_PictureDown, Nothing
      .WriteProperty "PictureAlign", m_PictureAlign, epLeftOfCaption
      .WriteProperty "PictureEffectOnOver", m_PicEffectonOver, epeLighter
      .WriteProperty "PictureEffectOnDown", m_PicEffectonDown, epeDarker
      .WriteProperty "PicturePushOnHover", m_bPicPushOnHover, False
      .WriteProperty "PictureShadow", m_PictureShadow, False
      .WriteProperty "PictureOpacity", m_PictureOpacity, 255
      .WriteProperty "PictureOpacityOnOver", m_PicOpacityOnOver, 255
      .WriteProperty "DisabledPictureMode", m_PicDisabledMode, edpBlended
      .WriteProperty "CaptionEffects", m_CaptionEffects, vbNullString
      .WriteProperty "UseMaskColor", m_bUseMaskColor, True
      .WriteProperty "MaskColor", m_lMaskColor, &HE0E0E0
      .WriteProperty "CaptionAlign", m_CaptionAlign, ecCenterAlign
      .WriteProperty "ToolTip", m_sTooltipText, vbNullString
      .WriteProperty "TooltipType", m_lTooltipType, TooltipStandard
      .WriteProperty "TooltipIcon", m_lToolTipIcon, TTNoIcon
      .WriteProperty "TooltipTitle", m_sTooltiptitle, vbNullString
      .WriteProperty "TooltipBackColor", m_lttBackColor, TranslateColor(vbInfoBackground)
      .WriteProperty "RightToLeft", m_bRTL, False
      .WriteProperty "DropDownSymbol", m_DropDownSymbol, ebsNone
      .WriteProperty "DropDownSeparator", m_bDropDownSep, False
      .WriteProperty "ColorScheme", m_lXPColor, ecsBlue
   End With

End Sub

Private Function Is32BitBMP(obj As Object) As Boolean
   Dim uBI              As BITMAP

   If obj.Type = vbPicTypeBitmap Then
      Call GetObject(obj.Handle, Len(uBI), uBI)
      Is32BitBMP = uBI.bmBitsPixel = 32
   End If
   
End Function

'Purpose: Returns True if DLL is present.
Private Function IsDLLPresent(ByVal sDLL As String) As Boolean

   On Error GoTo NotPresent
   Dim hLib             As Long
   hLib = LoadLibraryA(sDLL)
   If hLib <> 0 Then
      FreeLibrary hLib
      IsDLLPresent = True
   End If
NotPresent:

End Function

Private Property Get IsThemed() As Boolean

   On Error Resume Next
   Static m_bInit
   If HasUxTheme Then
      If Not (m_bInit) Then
         m_bIsThemed = IsAppThemed
         m_bInit = True
      End If
   End If
   IsThemed = m_bIsThemed
   
End Property

Private Property Get HasUxTheme() As Boolean

   Static m_bInit
   If Not (m_bInit) Then
      m_bHasUxTheme = IsDLLPresent("uxtheme.dll")
      m_bInit = True
   End If
   HasUxTheme = m_bHasUxTheme
   
End Property

'Determine if the passed function is supported
Private Function IsFunctionSupported(ByVal sFunction As String, ByVal sModule As String) As Boolean

Dim lngModule As Long

   lngModule = GetModuleHandle(sModule)
   
   If lngModule = 0 Then lngModule = LoadLibraryA(sModule)
   
   If lngModule Then
      IsFunctionSupported = GetProcAddress(lngModule, sFunction)
      FreeLibrary lngModule
   End If

End Function


'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

   Dim tme              As TRACKMOUSEEVENT_STRUCT

   If TrackUser32 Then
      With tme
         .cbSize = Len(tme)
         .dwFlags = TME_LEAVE
         .hwndTrack = lng_hWnd
      End With

      If TrackUser32 Then
         TrackMouseEvent tme
      Else
         TrackMouseEventComCtl tme
      End If
   End If

End Sub

'=========================================================================
'PUBLIC ROUTINES including subclassing & public button properties

' CREDITS: Paul Caton
'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub Subclass_WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lHwnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

   'Parameters:
   'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
   'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
   'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
   'hWnd     - The window handle
   'uMsg     - The message number
   'wParam   - Message related data
   'lParam   - Message related data
   'Notes:
   'If you really know what you're doing, it's possible to change the values of the
   'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
   'values get passed to the default handler.. and optionaly, the 'after' callback

   Static bMoving As Boolean

   Select Case uMsg

      Case WM_MOUSELEAVE

         m_bMouseInCtl = False
         If m_bPopupEnabled Then
            If m_bPopupInit Then
               m_bPopupInit = False
               m_bPopupShown = True
               Exit Sub
            Else
               m_bPopupShown = False
            End If
         End If

         If m_bIsSpaceBarDown Then Exit Sub
         If m_Buttonstate <> eStateNormal Then
            m_Buttonstate = eStateNormal
            RedrawButton
         End If
         RaiseEvent MouseLeave

      Case WM_NCACTIVATE, WM_ACTIVATE
         If wParam Then
            m_bParentActive = True
            If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
            If m_bDefault Then
               RedrawButton
            End If
            RedrawButton
         Else
            m_bIsDown = False
            m_bIsSpaceBarDown = False
            m_bHasFocus = False
            m_bParentActive = False
            If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
            RedrawButton
         End If

      Case WM_THEMECHANGED
         RedrawButton

      Case WM_SYSCOLORCHANGE
         RedrawButton
   End Select

End Sub

Public Sub SetPopupMenu(Menu As Object, Optional Align As enumMenuAlign, Optional Flags = 0, Optional DefaultMenu = Nothing)
Attribute SetPopupMenu.VB_Description = "Sets a dropdown menu to the button."

   If Not (Menu Is Nothing) Then
      If (TypeOf Menu Is VB.Menu) Then

         Set DropDownMenu = Menu
         MenuAlign = Align
         MenuFlags = Flags
         Set DefaultMenu = DefaultMenu
         m_bPopupEnabled = True
      End If
   End If

End Sub

Public Sub UnsetPopupMenu()
Attribute UnsetPopupMenu.VB_Description = "Unsets a popupmenu that was previously set for that button."

   ' --Free the popup menu
   Set DropDownMenu = Nothing
   Set DefaultMenu = Nothing
   m_bPopupEnabled = False
   m_bPopupShown = False

End Sub

Public Sub OpenWebsite(ByVal sAddress As String)
Attribute OpenWebsite.VB_Description = "Opens a website specified by URL"
    
    ShellExecute hWnd, "open", sAddress, vbNullString, vbNullString, 1
    
End Sub

Public Sub About()
Attribute About.VB_Description = "Displays information about the control and its author."
Attribute About.VB_UserMemId = -552

   MsgBox "JCButton v 1.02" & vbNewLine & _
      "Author: Juned S. Chhipa" & vbNewLine & _
      "Contact: juned.chhipa@yahoo.com" & vbNewLine & vbNewLine & _
      "Copyright � 2008-2009 Juned Chhipa. All rights reserved.", vbInformation + vbOKOnly, "About"

End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used for the button."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501

   BackColor = m_bColors.tBackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

   m_bColors.tBackColor = New_BackColor
   If m_ButtonStyle <> eOfficeXP Then
      m_lXPColor = ecsCustom
   End If
   RedrawButton
   PropertyChanged "BackColor"

End Property

Public Property Get ButtonStyle() As enumButtonStlyes
Attribute ButtonStyle.VB_Description = "Returns/sets a value to determine the style used to draw the button."
Attribute ButtonStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"

   ButtonStyle = m_ButtonStyle

End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As enumButtonStlyes)

   m_ButtonStyle = New_ButtonStyle
   InitThemeColors
   SetThemeColors          'Set colors
   CreateRegion            'Create Region Again
   RedrawButton            'Obviously, force redraw!!!
   PropertyChanged "ButtonStyle"

End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the button."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518

   Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

   m_Caption = New_Caption
   SetAccessKey
   RedrawButton
   PropertyChanged "Caption"

End Property

Public Property Get CaptionAlign() As enumCaptionAlign
Attribute CaptionAlign.VB_Description = "Returns/Sets the position of the Caption."
Attribute CaptionAlign.VB_ProcData.VB_Invoke_Property = ";Position"

   CaptionAlign = m_CaptionAlign

End Property

Public Property Let CaptionAlign(ByVal New_CaptionAlign As enumCaptionAlign)

   m_CaptionAlign = New_CaptionAlign
   RedrawButton
   PropertyChanged "CaptionAlign"

End Property

Public Property Get DropDownSymbol() As enumSymbol
Attribute DropDownSymbol.VB_Description = "Returns/Sets the Symbol to be used for displaying PopupMenu."
Attribute DropDownSymbol.VB_ProcData.VB_Invoke_Property = ";Appearance"

   DropDownSymbol = m_DropDownSymbol

End Property

Public Property Let DropDownSymbol(ByVal New_Align As enumSymbol)

   m_DropDownSymbol = New_Align
   RedrawButton
   PropertyChanged "DropDownSymbol"

End Property

Public Property Get DropDownSeparator() As Boolean
Attribute DropDownSeparator.VB_Description = "Returns/Sets the value whether to display DropDown Separator."
Attribute DropDownSeparator.VB_ProcData.VB_Invoke_Property = ";Appearance"

   DropDownSeparator = m_bDropDownSep

End Property

Public Property Let DropDownSeparator(ByVal New_Value As Boolean)

   m_bDropDownSep = New_Value
   RedrawButton
   PropertyChanged "DropDownSeparator"

End Property

Public Property Get DisabledPictureMode() As enumDisabledPicMode
Attribute DisabledPictureMode.VB_Description = "Returns/Sets the effect to be used for picture when button is disabled."
Attribute DisabledPictureMode.VB_ProcData.VB_Invoke_Property = ";Appearance"
    
   DisabledPictureMode = m_PicDisabledMode
    
End Property

Public Property Let DisabledPictureMode(ByVal New_mode As enumDisabledPicMode)
 
   m_PicDisabledMode = New_mode
   RedrawButton
   PropertyChanged "DisabledPictureMode"
   
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value to determine whether the button can respond to events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514

   Enabled = m_bEnabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

   m_bEnabled = New_Enabled
   UserControl.Enabled = m_bEnabled
   RedrawButton
   PropertyChanged "Enabled"

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns/sets the Font used to display text on the button."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512

   Set Font = mFont

End Property

Public Property Set Font(ByVal New_Font As StdFont)

   Set mFont = New_Font
   Refresh
   RedrawButton
   PropertyChanged "Font"
   Call mFont_FontChanged("")

End Property

Private Sub mFont_FontChanged(ByVal PropertyName As String)

   Set UserControl.Font = mFont
   Refresh
   RedrawButton
   PropertyChanged "Font"

End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the text color of the button caption."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513

   ForeColor = m_bColors.tForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)

   m_bColors.tForeColor = New_ForeColor
   UserControl.ForeColor = m_bColors.tForeColor
   RedrawButton
   PropertyChanged "ForeColor"

End Property

Public Property Get ForeColorHover() As OLE_COLOR
Attribute ForeColorHover.VB_Description = "Returns/sets the text color of the button caption when Mouse is over the control."
Attribute ForeColorHover.VB_ProcData.VB_Invoke_Property = ";Appearance"

   ForeColorHover = m_bColors.tForeColorOver

End Property

Public Property Let ForeColorHover(ByVal New_ForeColorHover As OLE_COLOR)

   m_bColors.tForeColorOver = New_ForeColorHover
   UserControl.ForeColor = m_bColors.tForeColorOver
   RedrawButton
   PropertyChanged "ForeColorHover"

End Property

Public Property Get HandPointer() As Boolean
Attribute HandPointer.VB_Description = "Returns/sets a value to determine whether the control uses the system's hand pointer as its cursor."
Attribute HandPointer.VB_ProcData.VB_Invoke_Property = ";Misc"

   HandPointer = m_bHandPointer

End Property

Public Property Let HandPointer(ByVal New_HandPointer As Boolean)

   m_bHandPointer = New_HandPointer
   If m_bHandPointer Then
      UserControl.MousePointer = 0
   End If
   RedrawButton
   PropertyChanged "HandPointer"

End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle that uniquely identifies the control."
Attribute hWnd.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute hWnd.VB_UserMemId = -515

   ' --Handle that uniquely identifies the control
   hWnd = UserControl.hWnd

End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets a color in a button's picture to be transparent."
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"

   MaskColor = m_lMaskColor

End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)

   m_lMaskColor = New_MaskColor
   RedrawButton
   PropertyChanged "MaskColor"

End Property

Public Property Get Mode() As enumButtonModes
Attribute Mode.VB_Description = "Returns/sets the type of control the button will observe."
Attribute Mode.VB_ProcData.VB_Invoke_Property = ";Behavior"

   Mode = m_ButtonMode

End Property

Public Property Let Mode(ByVal New_mode As enumButtonModes)

   m_ButtonMode = New_mode
   If m_ButtonMode = ebmCommandButton Then
      m_Buttonstate = eStateNormal        'Force Normal State for command buttons
   End If
   RedrawButton
   PropertyChanged "Value"
   PropertyChanged "Mode"

End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon for the button."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"

   Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal New_Icon As IPictureDisp)

   On Error Resume Next
   Set UserControl.MouseIcon = New_Icon
   If (New_Icon Is Nothing) Then
      UserControl.MousePointer = 0 ' vbDefault
   Else
      m_bHandPointer = False
      PropertyChanged "HandPointer"
      UserControl.MousePointer = 99 ' vbCustom
   End If
   PropertyChanged "MouseIcon"

End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when cursor over the button."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"

   MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal New_Cursor As MousePointerConstants)

   UserControl.MousePointer = New_Cursor
   PropertyChanged "MousePointer"

End Property

Public Property Get PictureNormal() As StdPicture
Attribute PictureNormal.VB_Description = "Returns/sets the picture displayed on a normal state button."
Attribute PictureNormal.VB_ProcData.VB_Invoke_Property = ";Appearance"

   Set PictureNormal = m_Picture

End Property

Public Property Set PictureNormal(ByVal New_Picture As StdPicture)

   Set m_Picture = New_Picture
   If Not New_Picture Is Nothing Then
      RedrawButton
      PropertyChanged "PictureNormal"
   Else
      UserControl_Resize
      Set m_PictureHot = Nothing
      Set m_PictureDown = Nothing
      PropertyChanged "PictureHot"
      PropertyChanged "PictureDown"
   End If

End Property

Public Property Get PictureHot() As StdPicture
Attribute PictureHot.VB_Description = "Returns/sets the picture displayed when the cursor is over the control."
Attribute PictureHot.VB_ProcData.VB_Invoke_Property = ";Appearance"

   Set PictureHot = m_PictureHot

End Property

Public Property Set PictureHot(ByVal New_Hot As StdPicture)

   If m_Picture Is Nothing Then
      Set m_Picture = New_Hot
      PropertyChanged "PictureNormal"
      Exit Property
   End If

   Set m_PictureHot = New_Hot
   PropertyChanged "PictureHot"
   RedrawButton

End Property

Public Property Get PictureDown() As StdPicture
Attribute PictureDown.VB_Description = "Returns/sets the picture displayed when the control is pressed down or in checked state."
Attribute PictureDown.VB_ProcData.VB_Invoke_Property = ";Appearance"

   Set PictureDown = m_PictureDown

End Property

Public Property Set PictureDown(ByVal New_Down As StdPicture)

   If m_Picture Is Nothing Then
      Set m_Picture = New_Down
      PropertyChanged "PictureNormal"
      Exit Property
   End If

   Set m_PictureDown = New_Down
   PropertyChanged "PictureDown"
   RedrawButton

End Property

Public Property Get PictureAlign() As enumPictureAlign
Attribute PictureAlign.VB_Description = "Returns/sets a value to determine where to draw the picture in the button."
Attribute PictureAlign.VB_ProcData.VB_Invoke_Property = ";Position"

   PictureAlign = m_PictureAlign

End Property

Public Property Let PictureAlign(ByVal New_PictureAlign As enumPictureAlign)

   m_PictureAlign = New_PictureAlign
   If Not m_Picture Is Nothing Then
      RedrawButton
   End If
   PropertyChanged "PictureAlign"

End Property

Public Property Get PictureShadow() As Boolean
Attribute PictureShadow.VB_Description = "Returns/Sets a value to determine whether to display Picture Shadow"
Attribute PictureShadow.VB_ProcData.VB_Invoke_Property = ";Appearance"

   PictureShadow = m_PictureShadow

End Property

Public Property Let PictureShadow(ByVal New_Shadow As Boolean)

   m_PictureShadow = New_Shadow
   RedrawButton
   PropertyChanged "PictureShadow"

End Property

Public Property Get PictureEffectOnOver() As enumPicEffect
Attribute PictureEffectOnOver.VB_Description = "Returns/Sets the Picture Effects to be applied when the mouseis over the control."
Attribute PictureEffectOnOver.VB_ProcData.VB_Invoke_Property = ";Appearance"

   PictureEffectOnOver = m_PicEffectonOver

End Property

Public Property Let PictureEffectOnOver(ByVal New_Effect As enumPicEffect)

   m_PicEffectonOver = New_Effect
   RedrawButton
   PropertyChanged "PictureEffectOnOver"

End Property

Public Property Get PictureEffectOnDown() As enumPicEffect
Attribute PictureEffectOnDown.VB_Description = "Returns/Sets the Picture Effects to be applied when the Button is pressed down."
Attribute PictureEffectOnDown.VB_ProcData.VB_Invoke_Property = ";Appearance"

   PictureEffectOnDown = m_PicEffectonDown

End Property

Public Property Let PictureEffectOnDown(ByVal New_Effect As enumPicEffect)

   m_PicEffectonDown = New_Effect
   RedrawButton
   PropertyChanged "PictureEffectOnDown"

End Property

Public Property Get PicturePushOnHover() As Boolean
Attribute PicturePushOnHover.VB_Description = "Returns/Sets a value to determine whether to Push picture when Mouse is over the control."
Attribute PicturePushOnHover.VB_ProcData.VB_Invoke_Property = ";Appearance"

   PicturePushOnHover = m_bPicPushOnHover

End Property

Public Property Let PicturePushOnHover(ByVal Value As Boolean)

   m_bPicPushOnHover = Value
   RedrawButton
   PropertyChanged "PicturePushOnHover"

End Property

Public Property Get PictureOpacity() As Byte
Attribute PictureOpacity.VB_Description = "Returns/Sets a byte value to control the Opacity of the Picture."
Attribute PictureOpacity.VB_ProcData.VB_Invoke_Property = ";Appearance"

    PictureOpacity = m_PictureOpacity
    
End Property

Public Property Let PictureOpacity(ByVal New_Opacity As Byte)

    m_PictureOpacity = New_Opacity
    RedrawButton
    PropertyChanged "PictureOpacity"
    
End Property

Public Property Get PictureOpacityOnOver() As Byte
Attribute PictureOpacityOnOver.VB_Description = "Returns/Sets a byte value to control the Opacity of the Picture when Mouse is over the button."
Attribute PictureOpacityOnOver.VB_ProcData.VB_Invoke_Property = ";Appearance"

    PictureOpacityOnOver = m_PicOpacityOnOver
    
End Property

Public Property Let PictureOpacityOnOver(ByVal New_Opacity As Byte)

    m_PicOpacityOnOver = New_Opacity
    RedrawButton
    PropertyChanged "PictureOpacityOnOver"
    
End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Returns/Sets a value to determine whether to display text in RTL mode."
Attribute RightToLeft.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute RightToLeft.VB_UserMemId = -611
    
    RightToLeft = m_bRTL
    
End Property

Public Property Let RightToLeft(ByVal Value As Boolean)
    
    m_bttRTL = Value
    m_bRTL = Value
    RedrawButton
    PropertyChanged "RightToLeft"
    
End Property

Public Property Get CaptionEffects() As enumCaptionEffects
Attribute CaptionEffects.VB_Description = "Returns/Sets the Special Effects apply to the caption."
Attribute CaptionEffects.VB_ProcData.VB_Invoke_Property = ";Appearance"

   CaptionEffects = m_CaptionEffects

End Property

Public Property Let CaptionEffects(ByVal New_Effects As enumCaptionEffects)

   m_CaptionEffects = New_Effects
   RedrawButton
   PropertyChanged "CaptionEffects"

End Property

Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_Description = "Returns/Sets a value to show Focusrect when the button has focus"
Attribute ShowFocusRect.VB_ProcData.VB_Invoke_Property = ";Appearance"

   ShowFocusRect = m_bShowFocus

End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)

   m_bShowFocus = New_ShowFocusRect
   PropertyChanged "ShowFocusRect"

End Property

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets a value to determine whether to use MaskColor to create transparent areas of the picture."
Attribute UseMaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"

   UseMaskColor = m_bUseMaskColor

End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)

   m_bUseMaskColor = New_UseMaskColor
   If Not m_Picture Is Nothing Then
      RedrawButton
   End If
   PropertyChanged "UseMaskColor"

End Property

Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Returns/sets the value or state of the button."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Behavior"

   Value = m_bValue

End Property

Public Property Let Value(ByVal New_Value As Boolean)

   If m_ButtonMode <> ebmCommandButton Then
      m_bValue = New_Value
      If Not m_bValue Then
         m_Buttonstate = eStateNormal
      End If
      RedrawButton
      PropertyChanged "Value"
   Else
      m_Buttonstate = eStateNormal
      RedrawButton
   End If

End Property

Public Property Get TooltipTitle() As String
Attribute TooltipTitle.VB_Description = "Returns/Sets the text displayed in the title of the tooltip."

   TooltipTitle = m_sTooltiptitle

End Property

Public Property Let TooltipTitle(ByVal New_title As String)

   m_sTooltiptitle = New_title
   RedrawButton
   PropertyChanged "TooltipTitle"

End Property

Public Property Get ToolTip() As String
Attribute ToolTip.VB_Description = "Returns/Sets the text displayed when mouse is paused over the control."
Attribute ToolTip.VB_ProcData.VB_Invoke_Property = ";Appearance"

   ToolTip = m_sTooltipText

End Property

Public Property Let ToolTip(ByVal New_Tooltip As String)

   m_sTooltipText = New_Tooltip
   RedrawButton
   PropertyChanged "ToolTip"

End Property

Public Property Get TooltipBackColor() As OLE_COLOR
Attribute TooltipBackColor.VB_Description = "Returns/Sets color to be displayed in the tooltip."
Attribute TooltipBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"

    TooltipBackColor = m_lttBackColor
    
End Property

Public Property Let TooltipBackColor(ByVal New_Color As OLE_COLOR)

    m_lttBackColor = New_Color
    RedrawButton
    PropertyChanged "TooltipBackcolor"
    
End Property

Public Property Let ToolTipIcon(lTooltipIcon As enumIconType)
Attribute ToolTipIcon.VB_Description = "Returns/Sets an icon value to be displayed in the tooltip."
Attribute ToolTipIcon.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"

   m_lToolTipIcon = lTooltipIcon
   RedrawButton
   PropertyChanged "TooltipIcon"

End Property

Public Property Get ToolTipIcon() As enumIconType

   ToolTipIcon = m_lToolTipIcon

End Property

Public Property Get ToolTipType() As enumTooltipStyle
Attribute ToolTipType.VB_Description = "Returns/Sets the Style of the tooltip."
Attribute ToolTipType.VB_ProcData.VB_Invoke_Property = ";Appearance"

   ToolTipType = m_lTooltipType

End Property

Public Property Let ToolTipType(ByVal lNewTTType As enumTooltipStyle)

   m_lTooltipType = lNewTTType
   RedrawButton
   PropertyChanged "ToolTipType"

End Property

Public Property Get ColorScheme() As enumXPThemeColors
Attribute ColorScheme.VB_Description = "Returns/Sets the ColorScheme to be used for the Background color."
Attribute ColorScheme.VB_ProcData.VB_Invoke_Property = ";Appearance"

   ColorScheme = m_lXPColor

End Property

Public Property Let ColorScheme(ByVal New_Color As enumXPThemeColors)

   m_lXPColor = New_Color
   SetThemeColors
   RedrawButton
   PropertyChanged "ColorScheme"

End Property

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Stop subclassing the passed window handle

Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long

   Subclass_AddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
   Debug.Assert Subclass_AddrFunc

End Function

Private Function Subclass_Index(ByVal lHwnd As Long, Optional ByVal bAdd As Boolean) As Long

   For Subclass_Index = UBound(SubclassData) To 0 Step -1
      If SubclassData(Subclass_Index).hWnd = lHwnd Then
         If Not bAdd Then Exit Function
         
      ElseIf SubclassData(Subclass_Index).hWnd = 0 Then
         If bAdd Then Exit Function
      End If
   Next 'Subclass_Index
   
   If Not bAdd Then Debug.Assert False

End Function

Private Function Subclass_InIDE() As Boolean

   Debug.Assert Subclass_SetTrue(Subclass_InIDE)

End Function

Private Function Subclass_Initialize(ByVal lHwnd As Long) As Long

Const CODE_LEN                  As Long = 200
Const GMEM_FIXED                As Long = 0
Const PATCH_01                  As Long = 18
Const PATCH_02                  As Long = 68
Const PATCH_03                  As Long = 78
Const PATCH_06                  As Long = 116
Const PATCH_07                  As Long = 121
Const PATCH_0A                  As Long = 186
Const FUNC_CWP                  As String = "CallWindowProcA"
Const FUNC_EBM                  As String = "EbMode"
Const FUNC_SWL                  As String = "SetWindowLongA"
Const MOD_USER                  As String = "User32"
Const MOD_VBA5                  As String = "vba5"
Const MOD_VBA6                  As String = "vba6"

Static bytBuffer(1 To CODE_LEN) As Byte
Static lngCWP                   As Long
Static lngEbMode                As Long
Static lngSWL                   As Long

Dim lngCount                    As Long
Dim lngIndex                    As Long
Dim strHex                      As String

   If bytBuffer(1) Then
      lngIndex = Subclass_Index(lHwnd, True)
      
      If lngIndex = -1 Then
         lngIndex = UBound(SubclassData()) + 1
         
         ReDim Preserve SubclassData(lngIndex) As SubClassDatatype
      End If
      
      Subclass_Initialize = lngIndex
      
   Else
      strHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
      
      For lngCount = 1 To CODE_LEN
         bytBuffer(lngCount) = Val("&H" & Left(strHex, 2))
         strHex = Mid(strHex, 3)
      Next 'lngCount
      
      If Subclass_InIDE Then
         bytBuffer(16) = &H90
         bytBuffer(17) = &H90
         lngEbMode = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)
         
         If lngEbMode = 0 Then lngEbMode = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)
      End If
      
      lngCWP = Subclass_AddrFunc(MOD_USER, FUNC_CWP)
      lngSWL = Subclass_AddrFunc(MOD_USER, FUNC_SWL)
      
      ReDim SubclassData(0) As SubClassDatatype
   End If
   
   With SubclassData(lngIndex)
      .hWnd = lHwnd
      .nAddrSclass = GlobalAlloc(GMEM_FIXED, CODE_LEN)
      .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSclass)
      
      Call CopyMemory(ByVal .nAddrSclass, bytBuffer(1), CODE_LEN)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_01, lngEbMode)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_02, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_03, lngSWL)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_06, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_07, lngCWP)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_0A, ObjPtr(Me))
   End With

End Function

Private Function Subclass_SetTrue(ByRef bValue As Boolean) As Boolean

   Subclass_SetTrue = True
   bValue = True

End Function

Private Sub Subclass_AddMsg(ByVal lHwnd As Long, ByVal uMsg As Long, Optional ByVal When As MsgWhen = MSG_AFTER)

   With SubclassData(Subclass_Index(lHwnd))
      If When And MSG_BEFORE Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass)
      If When And MSG_AFTER Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass)
   End With

End Sub

Private Sub Subclass_DelMsg(ByVal lHwnd As Long, ByVal uMsg As Long, Optional ByVal When As MsgWhen = MSG_AFTER)

   With SubclassData(Subclass_Index(lHwnd))
      If When And MSG_BEFORE Then Call Subclass_DoDelMsg(uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass)
      If When And MSG_AFTER Then Call Subclass_DoDelMsg(uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass)
   End With

End Sub

Private Sub Subclass_DoAddMsg(ByVal uMsg As Long, ByRef aMsgTabel() As Long, ByRef nMsgCount As Long, ByVal When As MsgWhen, ByVal nAddr As Long)

Const PATCH_04 As Long = 88
Const PATCH_08 As Long = 132

Dim lngEntry   As Long

   ReDim lngOffset(1) As Long
   
   If uMsg = ALL_MESSAGES Then
      nMsgCount = ALL_MESSAGES
      
   Else
      For lngEntry = 1 To nMsgCount - 1
         If aMsgTabel(lngEntry) = 0 Then
            aMsgTabel(lngEntry) = uMsg
            
            GoTo ExitSub
            
         ElseIf aMsgTabel(lngEntry) = uMsg Then
            GoTo ExitSub
         End If
      Next 'lngEntry
      
      nMsgCount = nMsgCount + 1
      
      ReDim Preserve aMsgTabel(1 To nMsgCount) As Long
      
      aMsgTabel(nMsgCount) = uMsg
   End If
   
   If When = MSG_BEFORE Then
      lngOffset(0) = PATCH_04
      lngOffset(1) = PATCH_05
      
   Else
      lngOffset(0) = PATCH_08
      lngOffset(1) = PATCH_09
   End If
   
   If uMsg <> ALL_MESSAGES Then Call Subclass_PatchVal(nAddr, lngOffset(0), VarPtr(aMsgTabel(1)))
   
   Call Subclass_PatchVal(nAddr, lngOffset(1), nMsgCount)
   
ExitSub:
   Erase lngOffset

End Sub

Private Sub Subclass_DoDelMsg(ByVal uMsg As Long, ByRef aMsgTabel() As Long, ByRef nMsgCount As Long, ByVal When As MsgWhen, ByVal nAddr As Long)

Dim lngEntry As Long

   If uMsg = ALL_MESSAGES Then
      nMsgCount = 0
      
      If When = MSG_BEFORE Then
         lngEntry = PATCH_05
         
      Else
         lngEntry = PATCH_09
      End If
      
      Call Subclass_PatchVal(nAddr, lngEntry, 0)
      
   Else
      For lngEntry = 1 To nMsgCount - 1
         If aMsgTabel(lngEntry) = uMsg Then
            aMsgTabel(lngEntry) = 0
            Exit For
         End If
      Next 'lngEntry
   End If

End Sub

Private Sub Subclass_PatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)

   Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)

End Sub

Private Sub Subclass_PatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)

   Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)

End Sub

Private Sub Subclass_Stop(ByVal lHwnd As Long)

   With SubclassData(Subclass_Index(lHwnd))
      SetWindowLongA .hWnd, GWL_WNDPROC, .nAddrOrig
      
      Call Subclass_PatchVal(.nAddrSclass, PATCH_05, 0)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_09, 0)
      
      GlobalFree .nAddrSclass
      .hWnd = 0
      .nMsgCountA = 0
      .nMsgCountB = 0
      Erase .aMsgTabelA, .aMsgTabelB
   End With

End Sub

Private Sub Subclass_Terminate()

Dim lngCount As Long

   For lngCount = UBound(SubclassData) To 0 Step -1
      If SubclassData(lngCount).hWnd Then Call Subclass_Stop(SubclassData(lngCount).hWnd)
   Next 'lngCount

End Sub

'---------------x---------------x--------------x--------------x-----------x---
' Oops! Control resulted Longer than expected!
' Lots of hours and lots of tedious work!   This is my first submission on PSC
' So if you want to vote for this, just do it ;)
' Comments are greatly appreciated...
' Enjoy!
'---------------x---------------x--------------x--------------x-----------x---
