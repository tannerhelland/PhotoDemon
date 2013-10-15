VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Font Manager (and Renderer)
'Copyright ©2012-2013 by Tanner Helland
'Created: 31/May/13
'Last updated: 31/May/13
'Last update: initial build
'
'While VB6 actually provides pretty good font support, PhotoDemon has some unique needs (as usual).  Chief among these
' are detailed handling over font positioning, rendering to arbitrary DCs, and managing fonts in a more speed-friendly
' way to maximize peformance.
'
'This class provides persistent storage for a variety of font parameters.  To maximize performance, it breaks down font
' management into three steps:
' 1) Set font parameters.  This can be done via a variety of support functions.  All attributes (minus color and
'     alignment, which are set separately) should be set before attempting to create the font.
' 2) Create font.  Do not do this until all settings are specified, because font creation is computationally expensive.
' 3) Attach font to DC.  Font color and text alignment should be set prior to this action.
' 4) (Optional) Render font to attached DC.  At present, ClearType antialiasing is assumed, but in the future it may
'     become user-settable.
'
'Obviously, this class relies heavily on the API.  Functions are documented to the best of my knowledge and ability.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 32
End Type

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

'We only deal with ANSI characters at present
Private Const ANSI_CHARSET = 0

'Constants for font weight (bold)
Private Const FW_DONTCARE As Long = 0
Private Const FW_THIN As Long = 100
Private Const FW_EXTRALIGHT As Long = 200
Private Const FW_ULTRALIGHT As Long = 200
Private Const FW_LIGHT As Long = 300
Private Const FW_NORMAL As Long = 400
Private Const FW_REGULAR As Long = 400
Private Const FW_MEDIUM As Long = 500
Private Const FW_SEMIBOLD As Long = 600
Private Const FW_DEMIBOLD As Long = 600
Private Const FW_BOLD As Long = 700
Private Const FW_EXTRABOLD As Long = 800
Private Const FW_ULTRABOLD As Long = 800
Private Const FW_HEAVY As Long = 900
Private Const FW_BLACK As Long = 900

'Font quality
Private Const DEFAULT_QUALITY As Long = 0
Private Const DRAFT_QUALITY As Long = 1
Private Const PROOF_QUALITY As Long = 2
Private Const NONANTIALIASED_QUALITY As Long = 3
Private Const ANTIALIASED_QUALITY As Long = 4
Private Const CLEARTYPE_QUALITY As Byte = 5

'Text alignment
Private Const TA_LEFT = 0
Private Const TA_RIGHT = 2
Private Const TA_CENTER = 6

'Transparent font background
Private Const FONT_TRANSPARENT = &H3

'Formatting constants for DrawText
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000
Private Const DT_EDITCONTROL = &H2000
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Integer) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long

Private Declare Function ExtTextOutNoRect Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByVal fakeRectPointer As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpDx As Long) As Long
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByRef lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, ByRef lpDx As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'Used to measure the expected length of a string (single line only; line breaks do not work)
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, ByRef lpSize As POINTAPI) As Long

'Used to track the created font
Dim curFont As Long, oldFont As Long

'Used to track the DC we are rendering to
Dim curDC As Long

'When the class is initialized, this value is recorded to help us convert from points to ems
Dim curLogPixelsY As Long

'Current text alignment (as an API constant, not a VB one)
Dim curTextAlignment As Long

'Current font color
Dim fontColor As Long

'This LOGFONT object stores all current settings for the font
Dim currentLogFont As LOGFONT

'This function draws horizontally and vertically centered text onto the attached DC, and it does it without modifying any of the current
' font settings (including alignment).  It is designed for one-off renderings of centered text.
Public Sub drawCenteredText(ByRef srcString As String)

    'Because DCs are dimensionless abstractions, we need to retrieve the attached window and use it to calculate the output rect
    Dim winClientRect As winRect
    GetClientRect WindowFromDC(curDC), winClientRect
    
    'Use the client rectangle to determine a width and height midpoint for the target DC
    Dim halfDCWidth As Long, halfDCHeight As Long
    halfDCWidth = winClientRect.x2 \ 2
    halfDCHeight = winClientRect.y2 \ 2

    'Next, determine the height of the string we were passed
    Dim strHeightHalf As Long
    strHeight = getHeightOfString(srcString) \ 2
    
    'Back-up the current text alignment, then change it to horizontally aligned
    Dim backupAlignment As AlignmentConstants
    backupAlignment = curTextAlignment
    updateAlignment vbCenter
    
    'Render the text
    ExtTextOutNoRect curDC, halfDCWidth, halfDCHeight - strHeight, 0, ByVal 0&, srcString, Len(srcString), ByVal 0&
    
    'Restore the original text alignment
    updateAlignment backupAlignment

End Sub

'If the font has already been created and set to a DC, this function can be used to quickly render text (because no
' additional parameters have to be set)
Public Sub fastRenderText(ByVal dstX As Long, ByVal dstY As Long, ByRef srcString As String)
    ExtTextOutNoRect curDC, dstX, dstY, 0, ByVal 0&, srcString, Len(srcString), ByVal 0&
End Sub

'Select the current font and other text settings into the specified DC.  Ideally, this should only be done when absolutely
' necessary to keep speed high.  (Also, note the DC so we can do fast font rendering later.)
Public Sub attachToDC(ByVal targetDC As Long)
    
    'Set the font
    oldFont = SelectObject(targetDC, curFont)
    curDC = targetDC
    
    'Set other font parameters
    SetTextColor targetDC, fontColor
    SetTextAlign targetDC, curTextAlignment
    
    'Enable transparent font rendering
    SetBkMode targetDC, FONT_TRANSPARENT
    
End Sub

'If we want to detach this font from its current DC, use this call
Public Sub releaseFromDC()
    SelectObject curDC, oldFont
End Sub

'Create the font object on-demand
Public Sub createFontObject()

    'If a font has already been created, delete it now
    If curFont > 0 Then
        If curDC > 0 Then SelectObject curDC, oldFont
        DeleteObject curFont
    End If
    
    curFont = CreateFontIndirect(currentLogFont)

End Sub

Public Sub setFontBold(ByVal newBold As Boolean)
    If newBold Then
        currentLogFont.lfWeight = FW_BOLD
    Else
        currentLogFont.lfWeight = FW_NORMAL
    End If
End Sub

Public Sub setFontColor(ByVal newColor As Long)
    fontColor = newColor
End Sub

Public Sub setFontFace(ByVal newFace As String)
    currentLogFont.lfFaceName = newFace & Chr(0)
End Sub

Public Sub setFontItalic(ByVal newItalic As Boolean)
    currentLogFont.lfItalic = -1 * newItalic
End Sub

Public Sub setFontSize(ByVal newFontSize As Long)
    currentLogFont.lfHeight = convertFontSize(newFontSize)
End Sub

Private Function convertFontSize(ByVal sizeInPoints As Long) As Long
    convertFontSize = -1 * MulDiv(sizeInPoints, curLogPixelsY, 72)
End Function

Public Sub setFontStrikeout(ByVal newStrikeout As Boolean)
    currentLogFont.lfStrikeOut = -1 * newStrikeout
End Sub

Public Sub setFontUnderline(ByVal newUnderline As Boolean)
    currentLogFont.lfUnderline = -1 * newUnderline
End Sub

Public Sub setTextAlignment(ByVal newAlignment As AlignmentConstants)
    Select Case newAlignment
        Case vbLeftJustify
            curTextAlignment = TA_LEFT
        Case vbCenter
            curTextAlignment = TA_CENTER
        Case vbRightJustify
            curTextAlignment = TA_RIGHT
    End Select
End Sub

'If this class has already been attached to a DC, this function can be used to immediately update the alignment
Public Sub updateAlignment(ByVal newAlignment As AlignmentConstants)
    Select Case newAlignment
        Case vbLeftJustify
            curTextAlignment = TA_LEFT
        Case vbCenter
            curTextAlignment = TA_CENTER
        Case vbRightJustify
            curTextAlignment = TA_RIGHT
    End Select
    
    SetTextAlign curDC, curTextAlignment
    
End Sub

'If this class has already been attached to a DC, this function can be used to immediately update the color
Public Sub updateColor(ByVal newColor As Long)
    fontColor = newColor
    SetTextColor curDC, fontColor
End Sub

'Return the width (and below, height) of a string, in pixels.  This action REQUIRES that a DC be currently set for this
' object, as the DC's font settings are used to determine the string's dimensions.
Public Function getWidthOfString(ByVal srcString As String) As Long
    Dim txtSize As POINTAPI
    GetTextExtentPoint32 curDC, srcString, Len(srcString), txtSize
    getWidthOfString = txtSize.x
End Function

Public Function getHeightOfString(ByVal srcString As String) As Long
    Dim txtSize As POINTAPI
    GetTextExtentPoint32 curDC, srcString, Len(srcString), txtSize
    getHeightOfString = txtSize.y
End Function

'With the current font settings, find the height of a word-wrapped string.  This action REQUIRES that a DC be currently
' set for this object, as the DC's font settings are used to calculate the word-wrapped string's dimensions.
Public Function getHeightOfWordwrapString(ByVal srcString As String, ByVal strWidth As Long) As Long

    Dim tmpRect As RECT
    tmpRect.Left = 0
    tmpRect.Top = 0
    tmpRect.Bottom = 0
    tmpRect.Right = strWidth - 1
    
    DrawText curDC, srcString, Len(srcString), tmpRect, DT_CALCRECT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_NOPREFIX Or DT_TOP
    
    getHeightOfWordwrapString = tmpRect.Bottom

End Function

Private Sub Class_Initialize()
    
    'Determine the proper font point-to-em conversion based on the current screen DPI
    Dim TmpDC As Long
    TmpDC = GetDC(0)
    curLogPixelsY = GetDeviceCaps(TmpDC, LOGPIXELSY)
    ReleaseDC 0, TmpDC
    
    'When the class is initialized, set up a default PhotoDemon font.  Face is automatically set per
    ' the current OS, so external function do not need to set this manually.
    With currentLogFont
        .lfHeight = convertFontSize(11)
        .lfWidth = 0        'This will automatically set the width based on the height
        .lfEscapement = 0   'Used for rotation; not utilized at present
        .lfOrientation = 0  'An additional rotation parameter; not utilized at present
        .lfWeight = FW_REGULAR
        .lfItalic = False
        .lfUnderline = False
        .lfStrikeOut = False
        .lfCharSet = ANSI_CHARSET
        .lfOutPrecision = 0
        .lfClipPrecision = 0
        .lfQuality = CLEARTYPE_QUALITY
        .lfPitchAndFamily = 0
        .lfFaceName = g_InterfaceFont & Chr(0)
    End With
    
    curFont = 0
    curDC = 0
    
    curTextAlignment = vbLeftJustify
    fontColor = 0
    curTextAlignment = TA_LEFT
    
End Sub

Private Sub Class_Terminate()

    'If a font was created, delete it now
    If curFont > 0 Then
        SelectObject curDC, oldFont
        DeleteObject curFont
    End If
    
End Sub
