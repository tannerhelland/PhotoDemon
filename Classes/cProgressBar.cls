VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'This class is ©2003 Steve McMahon, vbaccelerator.com

'This class was originally written by Steve McMahon.

'You may download the original version of this code from the following link (good as of June 2012):
' http://www.vbaccelerator.com/home/VB/Code/Controls/Progress_Bar/No_Progress_Bar_Class/article.asp

'To the best of my knowledge, this code is released under a CC-BY-1.0 license.  (Assumed from the footer text of vbaccelerator.com: "All contents of this web site are licensed under a Creative Commons Licence, except where otherwise noted.")
' You may access a complete copy of this license at the following link:
' http://creativecommons.org/licenses/by/1.0/

'Many thanks to Steve McMahon and vbaccelerator.com for this excellent progress bar class.


Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal HWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal HWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_SOFT = &H1000    ' For softer buttons
Private Const BF_FLAT = &H4000    ' For flat rather than 3D borders
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_SINGLELINE = &H20
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal HWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" _
   (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal lHDC As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pRect As RECT, pClipRect As RECT) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pBoundingRect As RECT, pContentRect As RECT) As Long
Private Declare Function DrawThemeText Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, ByVal pszText As Long, _
    ByVal iCharCount As Long, ByVal dwTextFlag As Long, _
    ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Private Declare Function DrawThemeIcon Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, pRect As RECT, _
    ByVal himl As Long, ByVal iImageIndex As Long) As Long
Private Declare Function DrawThemeEdge Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
   ByVal iStateId As Long, pDestRect As RECT, _
   ByVal uEdge As Long, ByVal uFlags As Long, _
   pContentRect As RECT) As Long
Private Enum THEMESIZE
    TS_MIN = 0             '// minimum size
    TS_TRUE = 1            '// size without stretching
    TS_DRAW = 2             ' // size that theme mgr will use to draw part
End Enum
Private Declare Function GetThemeInt Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, ByVal iPropId As Long, _
    piVal As Long) As Long
Private Const PROGRESSCHUNKSIZE = 2411
Private Const PROGRESSSPACESIZE = 2412

Private Const S_OK = 0

Public Enum EVPRGAppearanceConstants
   [EVPRGFlat]
   [EVPRG3DThin]
   [EVPRG3D]
End Enum
Public Enum EVPRGBorderStyleConstants
   [EVPRGNone]
   [EVPRGInset]
   [EVPRGRaised]
End Enum
Public Enum EVPRGPictureModeConstants
   [EVPRGStretch]
   [EVPRGTile]
End Enum
Public Enum EVPRGHorizontalTextAlignConstants
   [EVPRGLeft]
   [EVPRGcenter]
   [EVPRGRight]
End Enum
Public Enum EVPRGVerticalTextAlignConstants
   [EVPRGTop]
   [EVPRGVCenter]
   [EVPRGBottom]
End Enum

Private m_obj As Object
Private m_cMemDC As pcMemDC
Private m_picBack As pcMemDC
Private m_picBar As pcMemDC
Private m_eAppearance As EVPRGAppearanceConstants
Private m_eBorderStyle As EVPRGBorderStyleConstants
Private m_fnt As IFont
Private m_oBackColor As OLE_COLOR
Private m_oForeColor As OLE_COLOR
Private m_oBarColor As OLE_COLOR
Private m_oBarForeColor As OLE_COLOR
Private m_eBarPictureMode As EVPRGPictureModeConstants
Private m_eBackPictureMode As EVPRGPictureModeConstants
Private m_lMin As Long
Private m_lMax As Long
Private m_lValue As Long
Private m_eTextAlignX As EVPRGHorizontalTextAlignConstants
Private m_eTextAlignY As EVPRGVerticalTextAlignConstants
Private m_bShowText As Boolean
Private m_sText As String
Private m_bSegments As Boolean
Private m_bXpStyle As Boolean
Private m_sTag As String
Private m_tLastR As RECT
Private m_hWndLast As Long

Public Property Get Tag() As String
   Tag = m_sTag
End Property
Public Property Let Tag(ByVal sTag As String)
   m_sTag = sTag
End Property

Public Property Get Segments() As Boolean
   Segments = m_bSegments
End Property
Public Property Let Segments(ByVal bState As Boolean)
   m_bSegments = bState
   Draw
End Property

Public Property Get XpStyle() As Boolean
   XpStyle = m_bXpStyle
End Property
Public Property Let XpStyle(ByVal bState As Boolean)
   m_bXpStyle = bState
   Draw
End Property

Public Property Get Text() As String
   Text = m_sText
End Property

Public Property Let Text(ByVal sText As String)
   m_sText = sText
   Draw
End Property

Public Property Get TextAlignX() As EVPRGHorizontalTextAlignConstants
   TextAlignX = m_eTextAlignX
End Property
Public Property Let TextAlignX(ByVal eAlign As EVPRGHorizontalTextAlignConstants)
   m_eTextAlignX = eAlign
   Draw
End Property
Public Property Get TextAlignY() As EVPRGVerticalTextAlignConstants
   TextAlignY = m_eTextAlignY
End Property
Public Property Let TextAlignY(ByVal eAlign As EVPRGVerticalTextAlignConstants)
   m_eTextAlignY = eAlign
   Draw
End Property

Public Property Get ShowText() As Boolean
   ShowText = m_bShowText
End Property

Public Property Let ShowText(ByVal bState As Boolean)
   m_bShowText = bState
   Draw
End Property

Public Property Get Percent() As Double
Dim fPercent As Double
   fPercent = (m_lValue - m_lMin) / (m_lMax - m_lMin)
   If fPercent > 1# Then fPercent = 1#
   If fPercent < 0# Then fPercent = 0#
   Percent = fPercent * 100#
End Property

Public Property Get Min() As Long
   Min = m_lMin
End Property
Public Property Let Min(ByVal lMin As Long)
   m_lMin = lMin
   Draw
End Property
Public Property Get Max() As Long
   Max = m_lMax
End Property
Public Property Let Max(ByVal lMax As Long)
   m_lMax = lMax
   Draw
End Property
Public Property Get Value() As Long
   Value = m_lValue
End Property
Public Property Let Value(ByVal lValue As Long)
   m_lValue = lValue
   Draw
End Property

Public Property Get BorderStyle() As EVPRGBorderStyleConstants
   BorderStyle = m_eBorderStyle
End Property
Public Property Let BorderStyle(ByVal eStyle As EVPRGBorderStyleConstants)
   m_eBorderStyle = eStyle
   Draw
End Property
Public Property Get Appearance() As EVPRGAppearanceConstants
   Appearance = m_eAppearance
End Property
Public Property Let Appearance(ByVal eAppearance As EVPRGAppearanceConstants)
   m_eAppearance = eAppearance
   Draw
End Property
Public Property Get DrawObject() As Object
   Set DrawObject = m_obj
End Property
Public Property Let DrawObject(obj As Object)
   Set m_obj = obj
   Draw
End Property
Public Property Set DrawObject(obj As Object)
   Set m_obj = obj
   Draw
End Property

Public Sub Draw()
   If Not m_obj Is Nothing Then
      Dim tR As RECT
      GetClientRect m_obj.HWnd, tR
      DrawToDC m_obj.HWnd, m_obj.hdc, tR.Left, tR.Top, tR.Right, tR.Bottom
      Me.DrawObject.Picture = Me.DrawObject.Image
      Me.DrawObject.Refresh
      DoEvents
   Else
      ' Could use this to automatically paint the owner object
      ' when setting Min,Max,Value,Text etc
      If (m_hWndLast) Then
          'm_hWndLast , m_tLastR, 1
      End If
   End If
End Sub

Public Sub DrawToDC( _
      ByVal HWnd As Long, _
      ByVal hdc As Long, _
      ByVal lLeft As Long, ByVal lTop As Long, _
      ByVal lRight As Long, ByVal lBottom As Long _
   )
Dim lhDCU As Long
Dim lHDC As Long
Dim bMem As Boolean
Dim tR As RECT, tBR As RECT, tSR As RECT, tWR As RECT
Dim lWidth As Long, lHeight As Long
Dim lColor As Long
Dim hBr As Long
Dim hRgn As Long
Dim fPercent As Double
Dim bDrawText As Boolean
Dim hFntOld As Long
Dim iFnt As IFont
Dim i As Long
Dim lSegmentWidth As Long, lSegmentSpacing As Long
Dim bDrawnXpStyle As Boolean
Dim hTheme As Long
Dim hR As Long
Dim bDrawn As Boolean
   
   ' Reconstruct the rectangle:
   tR.Left = lLeft
   tR.Top = lTop
   tR.Right = lRight
   tR.Bottom = lBottom
   
   LSet m_tLastR = tR
   m_hWndLast = WindowFromDC(hdc)

   lWidth = Abs(tR.Right - tR.Left)
   lHeight = Abs(tR.Bottom - tR.Top)

   lhDCU = hdc
   lHDC = m_cMemDC.hdc(lWidth, lHeight)
   If lHDC = 0 Then
      lHDC = lhDCU
   Else
      bMem = True
      OffsetRect tR, -tR.Left, -tR.Top
   End If
   
   ' Draw background:
   If Not m_picBack Is Nothing Then
      If m_eBackPictureMode = EVPRGTile Then
         TileArea _
            lHDC, _
            0, 0, _
            lWidth, lHeight, _
            m_picBack.hdc(m_picBack.Width, m_picBack.Height), _
            m_picBack.Width, m_picBack.Height, _
            0, 0
      Else
         StretchBlt _
            lHDC, _
            0, 0, lWidth, lHeight, _
            m_picBack.hdc(m_picBack.Width, m_picBack.Height), _
            0, 0, _
            m_picBack.Width, m_picBack.Height, _
            vbSrcCopy
      End If
   Else
      If (m_bXpStyle) Then
         hTheme = OpenThemeData(HWnd, StrPtr("Progress"))
         If (hTheme <> 0) Then
            hR = GetThemeInt(hTheme, 0, 0, PROGRESSCHUNKSIZE, lSegmentWidth)
            If (hR = S_OK) Then
               hR = GetThemeInt(hTheme, 0, 0, PROGRESSSPACESIZE, lSegmentSpacing)
               If (hR = S_OK) Then
                  lSegmentWidth = lSegmentWidth + lSegmentSpacing
                  If (lWidth > lHeight) Then
                     hR = DrawThemeBackground(hTheme, lHDC, 1, 0, tR, tR)
                  Else
                     hR = DrawThemeBackground(hTheme, lHDC, 2, 0, tR, tR)
                  End If
                  If (hR = S_OK) Then
                     bDrawn = True
                  End If
               End If
            End If
         End If
      End If
      
      If Not (bDrawn) Then
         lColor = BackColor
         If lColor And &H80000000 Then
            hBr = GetSysColorBrush(lColor And &H1F&)
         Else
            hBr = CreateSolidBrush(lColor)
         End If
         FillRect lHDC, tR, hBr
         DeleteObject hBr
      End If
   End If
      
   If (m_bSegments) And Not (bDrawn) Then
      lSegmentWidth = 8
      lSegmentSpacing = 2
   End If

      
   LSet tWR = tR
   If m_eBorderStyle > EVPRGNone Then
      If bDrawn Then
         InflateRect tR, -1, -1
      Else
         If m_eAppearance = EVPRG3D Then
            InflateRect tR, -2, -2
         Else
            InflateRect tR, -1, -1
         End If
      End If
   End If
   
   If (m_bShowText) And Len(m_sText) > 0 Then
      bDrawText = True
   End If
   If (bDrawText) And Not (bDrawn) Then
      Set iFnt = Font
      hFntOld = SelectObject(lHDC, iFnt.hFont)
      SetBkMode lHDC, TRANSPARENT
      SetTextColor lHDC, TranslateColor(m_oForeColor)
      DrawText lHDC, " " & m_sText & " ", -1, tR, DT_SINGLELINE Or m_eTextAlignX Or m_eTextAlignY * 4
      SelectObject lHDC, hFntOld
   End If
   
   ' Draw bar:
   ' Get the bar rectangle:
   LSet tBR = tR
   fPercent = (m_lValue - m_lMin) / (m_lMax - m_lMin)
   If fPercent > 1# Then fPercent = 1#
   If fPercent < 0# Then fPercent = 0#
   If lWidth > lHeight Then
      tBR.Right = tR.Left + (tR.Right - tR.Left) * fPercent
      If (m_bSegments Or bDrawn) Then
         ' Quantise bar:
         tBR.Right = tBR.Right - ((tBR.Right - tBR.Left) Mod (lSegmentWidth + lSegmentSpacing))
         'Debug.Assert ((tBR.Right - tBR.Left) Mod (lSegmentWidth + lSegmentSpacing) = 0)
         If tBR.Right < tR.Left Then
            tBR.Right = tR.Left
         End If
      End If
   Else
      fPercent = 1# - fPercent
      tBR.Top = tR.Top + (tR.Bottom - tR.Top) * fPercent
      If (m_bSegments Or bDrawn) Then
         ' Quantise bar:
         tBR.Top = tBR.Top - ((tBR.Top - tBR.Bottom) Mod (lSegmentWidth + lSegmentSpacing))
         If tBR.Top > tR.Bottom Then
            tBR.Top = tR.Bottom
         End If
      End If
   End If
   
   If Not bDrawn Then
      hRgn = CreateRectRgnIndirect(tBR)
      SelectClipRgn lHDC, hRgn
   End If
   
   If Not m_picBar Is Nothing Then
      If m_eBackPictureMode = EVPRGTile Then
         TileArea _
            lHDC, _
            0, 0, _
            lWidth, lHeight, _
            m_picBar.hdc(m_picBar.Width, m_picBar.Height), _
            m_picBar.Width, m_picBar.Height, _
            0, 0
      Else
         StretchBlt _
            lHDC, _
            0, 0, lWidth, lHeight, _
            m_picBar.hdc(m_picBar.Width, m_picBar.Height), _
            0, 0, _
            m_picBar.Width, m_picBar.Height, _
            vbSrcCopy
      End If
   Else
      If bDrawn Then
         If (lWidth > lHeight) Then
            hR = DrawThemeBackground(hTheme, lHDC, 3, 0, tBR, tBR)
         Else
            hR = DrawThemeBackground(hTheme, lHDC, 4, 0, tBR, tBR)
         End If
      Else
         lColor = m_oBarColor
         If lColor And &H80000000 Then
            hBr = GetSysColorBrush(lColor And &H1F&)
         Else
            hBr = CreateSolidBrush(lColor)
         End If
         FillRect lHDC, tBR, hBr
         DeleteObject hBr
      End If
   End If
   
   If m_bSegments And Not bDrawn Then
      lColor = BackColor
      If lColor And &H80000000 Then
         hBr = GetSysColorBrush(lColor And &H1F&)
      Else
         hBr = CreateSolidBrush(lColor)
      End If
      LSet tSR = tR
      If lWidth > lHeight Then
         For i = tBR.Left + lSegmentWidth To tBR.Right Step lSegmentWidth + lSegmentSpacing
            tSR.Left = i
            tSR.Right = i + lSegmentSpacing
            FillRect lHDC, tSR, hBr
         Next i
      Else
         For i = tBR.Bottom To tBR.Top + lSegmentWidth Step -(lSegmentWidth + lSegmentSpacing)
            tSR.Top = i
            tSR.Bottom = i + lSegmentSpacing
            FillRect lHDC, tSR, hBr
         Next i
      End If
      DeleteObject hBr
   End If
      
   If bDrawText Then
      Set iFnt = Font
      hFntOld = SelectObject(lHDC, iFnt.hFont)
      If (bDrawn) Then
         Dim rcContent As RECT
         hR = GetThemeBackgroundContentRect(hTheme, _
                lHDC, 0, 0, tR, rcContent)
         hR = DrawThemeText(hTheme, lHDC, 0, 0, _
            StrPtr(m_sText), -1, _
            DT_SINGLELINE Or m_eTextAlignX Or m_eTextAlignY * 4, _
            0, rcContent)
      Else
         SetBkMode lHDC, TRANSPARENT
         SetTextColor lHDC, TranslateColor(m_oBarForeColor)
         DrawText lHDC, " " & m_sText & " ", -1, _
            tR, DT_SINGLELINE Or m_eTextAlignX Or m_eTextAlignY * 4
      End If
      SelectObject lHDC, hFntOld
   End If
      
   If Not bDrawn Then
      SelectClipRgn lHDC, 0
      DeleteObject hRgn
      
      ' Draw border:
      Select Case m_eBorderStyle
      Case EVPRGRaised
         Select Case m_eAppearance
         Case EVPRGFlat
            Border lHDC, EVPRGFlat, tWR, True
         Case EVPRG3DThin
            Border lHDC, EVPRG3DThin, tR, True
         Case EVPRG3D
            Border lHDC, EVPRG3D, tWR, True
         End Select
      Case EVPRGInset
         Select Case m_eAppearance
         Case EVPRGFlat
            Border lHDC, EVPRGFlat, tWR, False
         Case EVPRG3DThin
            Border lHDC, EVPRG3DThin, tWR, False
         Case EVPRG3D
            Border lHDC, EVPRG3D, tWR, False
         End Select
      End Select
   End If
   
   ' Swap memdc<->Screen
   If bMem Then
      m_cMemDC.Draw lhDCU, 0, 0, lWidth, lHeight, lLeft, lTop
   End If
   
   If (hTheme) Then
      CloseThemeData hTheme
   End If

End Sub
Private Sub Border( _
      ByVal lHDC As Long, _
      ByVal lStyle As Long, _
      ByRef tR As RECT, _
      ByVal bRaised As Boolean _
   )
Dim hPenDark As Long, hPenLight As Long, hPenBlack As Long
Dim hPenOld As Long
Dim tJunk As POINTAPI

   Select Case lStyle
   Case 0
      hPenBlack = CreatePen(0, 1, 0)
      hPenOld = SelectObject(lHDC, hPenBlack)
      MoveToEx lHDC, tR.Left, tR.Top, tJunk
      LineTo lHDC, tR.Right - 1, tR.Top
      LineTo lHDC, tR.Right - 1, tR.Bottom - 1
      LineTo lHDC, tR.Left, tR.Bottom - 1
      LineTo lHDC, tR.Left, tR.Top
      SelectObject lHDC, hPenOld
      DeleteObject hPenBlack
   Case 1
      hPenDark = CreatePen(0, 1, GetSysColor(vbButtonShadow And &H1F&))
      hPenLight = CreatePen(0, 1, GetSysColor(vb3DHighlight And &H1F&))
      If bRaised Then
         MoveToEx lHDC, tR.Left, tR.Bottom - 2, tJunk
         hPenOld = SelectObject(lHDC, hPenLight)
         LineTo lHDC, tR.Left, tR.Top
         LineTo lHDC, tR.Right - 1, tR.Top
         SelectObject lHDC, hPenOld
         MoveToEx lHDC, tR.Right - 1, tR.Top, tJunk
         hPenOld = SelectObject(lHDC, hPenDark)
         LineTo lHDC, tR.Right - 1, tR.Bottom - 1
         LineTo lHDC, tR.Left - 1, tR.Bottom - 1
         SelectObject lHDC, hPenOld
      Else
         MoveToEx lHDC, tR.Left, tR.Bottom - 1, tJunk
         hPenOld = SelectObject(lHDC, hPenDark)
         LineTo lHDC, tR.Left, tR.Top
         LineTo lHDC, tR.Right, tR.Top
         SelectObject lHDC, hPenOld
         MoveToEx lHDC, tR.Right - 1, tR.Top + 1, tJunk
         hPenOld = SelectObject(lHDC, hPenLight)
         LineTo lHDC, tR.Right - 1, tR.Bottom - 1
         LineTo lHDC, tR.Left, tR.Bottom - 1
         SelectObject lHDC, hPenOld
      End If
      DeleteObject hPenDark
      DeleteObject hPenLight
   Case 2
      If bRaised Then
         DrawEdge lHDC, tR, EDGE_RAISED, BF_RECT Or BF_SOFT
      Else
         DrawEdge lHDC, tR, EDGE_SUNKEN, BF_RECT Or BF_SOFT
      End If
   End Select
End Sub
      
Private Sub TileArea( _
        ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal lSrcDC As Long, _
        ByVal lBitmapW As Long, _
        ByVal lBitmapH As Long, _
        ByVal lSrcOffsetX As Long, _
        ByVal lSrcOffsetY As Long _
    )
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = ((x + lSrcOffsetX) Mod lBitmapW)
    lSrcStartY = ((y + lSrcOffsetY) Mod lBitmapH)
    lSrcStartWidth = (lBitmapW - lSrcStartX)
    lSrcStartHeight = (lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (y + Height)
        If (lDstY + lDstHeight) > (y + Height) Then
            lDstHeight = y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + Width)
            If (lDstX + lDstWidth) > (x + Width) Then
                lDstWidth = x + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDC, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = lBitmapH
    Loop
End Sub


Public Property Get BackColor() As OLE_COLOR
   BackColor = m_oBackColor
End Property
Public Property Let BackColor(oColor As OLE_COLOR)
   m_oBackColor = oColor
   Draw
End Property
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = m_oForeColor
End Property
Public Property Let ForeColor(oColor As OLE_COLOR)
   m_oForeColor = oColor
   Draw
End Property
Public Property Get Font() As IFont
   Set Font = m_fnt
End Property
Public Property Set Font(ByRef fnt As IFont)
   Set m_fnt = fnt
   Draw
End Property
Public Property Let Font(ByRef fnt As IFont)
   Set m_fnt = fnt
   Draw
End Property
Public Property Get BarColor() As OLE_COLOR
   BarColor = m_oBarColor
End Property
Public Property Let BarColor(oColor As OLE_COLOR)
   m_oBarColor = oColor
   Draw
End Property
Public Property Get BarForeColor() As OLE_COLOR
   BarForeColor = m_oBarForeColor
End Property
Public Property Let BarForeColor(oColor As OLE_COLOR)
   m_oBarForeColor = oColor
   Draw
End Property
Public Property Let BarPicture(pic As IPicture)
   pPicture pic, m_picBar
End Property
Public Property Set BarPicture(pic As IPicture)
   pPicture pic, m_picBar
End Property
Public Property Get BarPictureMode() As EVPRGPictureModeConstants
   BarPictureMode = m_eBarPictureMode
End Property
Public Property Let BarPictureMode(ByVal eMode As EVPRGPictureModeConstants)
   m_eBarPictureMode = eMode
   Draw
End Property
Public Property Get BackPictureMode() As EVPRGPictureModeConstants
   BackPictureMode = m_eBackPictureMode
End Property
Public Property Let BackPictureMode(ByVal eMode As EVPRGPictureModeConstants)
   m_eBackPictureMode = eMode
   Draw
End Property
Public Property Let Picture(pic As IPicture)
   pPicture pic, m_picBack
End Property
Public Property Set Picture(pic As IPicture)
   pPicture pic, m_picBack
End Property
Private Sub pPicture(pic As IPicture, memDC As pcMemDC)
Dim Handle As Long
   If pic Is Nothing Then
      Set memDC = Nothing
   Else
      Set memDC = New pcMemDC
      memDC.CreateFromPicture pic
   End If
   Draw
End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


Private Sub Class_Initialize()
   m_eAppearance = EVPRG3DThin
   m_eBorderStyle = EVPRGInset
   m_oBackColor = CLR_INVALID
   m_oBarColor = CLR_INVALID
   m_oBarForeColor = &HFFFFFF
   m_eBarPictureMode = EVPRGTile
   m_eBackPictureMode = EVPRGTile
   m_lMax = 100
   m_eTextAlignX = EVPRGcenter
   m_eTextAlignY = EVPRGVCenter
   Set m_cMemDC = New pcMemDC
   Dim fnt As New StdFont
   fnt.Name = "Tahoma"
   fnt.Size = 8
   Set Font = fnt
End Sub
