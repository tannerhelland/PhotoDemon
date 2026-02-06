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
'PhotoDemon Font Renderer
'Copyright 2013-2026 by Tanner Helland
'Created: 31/May/13
'Last updated: 25/August/17
'Last update: improve font creation performance
'
'While VB6 actually provides pretty good font support, PhotoDemon has unique needs (as usual).  In particular,
' we need detailed handling over font positioning, we often render to arbitrary DCs, and we need to manage
' fonts in specialized ways to maximize peformance while minimizing GDI resources.
'
'This class provides a (fairly) comprehensive wrapper for common GDI-based font operations.  In general,
' it separates font management into three steps:
' 1) Set font parameters.  This can be done via a variety of support functions.  For maximum performance,
'    all font attributes should be set *before* creating the font.
' 2) Create font.  Do not do this any more than you need to; font creation is expensive.
' 3) Attach font to a DC.  Font color and text alignment can be set prior to this action; this class is
'    smart enough to track these properties and assign them to attached DCs as necessary.
' 4) Render text to the attached DC (using this font, obviously).  GDI provides a number of different
'    text-rendering functions, and we wrap different ones in different ways.  If you aren't sure which
'    one is fastest (or best for your needs), profile several.  GDI performance is not always predictable.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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
    lfFaceName(0 To 63) As Byte
End Type

'GDI still takes a legacy charset value as input; this doesn't matter for our purposes, so pass a default value
Private Const DEFAULT_CHARSET As Long = 1

'Transparent font background
Private Const FONT_TRANSPARENT As Long = &H1

'Formatting constants for DrawText
Private Enum DT_Flags
    DT_TOP = &H0&
    DT_LEFT = &H0&
    DT_CENTER = &H1&
    DT_RIGHT = &H2&
    DT_VCENTER = &H4&
    DT_BOTTOM = &H8&
    DT_WORDBREAK = &H10&
    DT_SINGLELINE = &H20&
    DT_EXPANDTABS = &H40&
    DT_TABSTOP = &H80&
    DT_NOCLIP = &H100&
    DT_EXTERNALLEADING = &H200&
    DT_CALCRECT = &H400&
    DT_NOPREFIX = &H800&
    DT_INTERNAL = &H1000&
    DT_EDITCONTROL = &H2000&
    DT_PATH_ELLIPSIS = &H4000&
    DT_END_ELLIPSIS = &H8000&
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
End Enum

#If False Then
    Private Const DT_TOP = &H0&, DT_LEFT = &H0&, DT_CENTER = &H1&, DT_RIGHT = &H2&, DT_VCENTER = &H4&, DT_BOTTOM = &H8&, DT_WORDBREAK = &H10&, DT_SINGLELINE = &H20&, DT_EXPANDTABS = &H40&, DT_TABSTOP = &H80&, DT_NOCLIP = &H100&, DT_EXTERNALLEADING = &H200&, DT_CALCRECT = &H400&, DT_NOPREFIX = &H800&, DT_INTERNAL = &H1000&, DT_EDITCONTROL = &H2000&, DT_PATH_ELLIPSIS = &H4000&, DT_END_ELLIPSIS = &H8000&, DT_MODIFYSTRING = &H10000, DT_RTLREADING = &H20000, DT_WORD_ELLIPSIS = &H40000, DT_NOFULLWIDTHCHARBREAK = &H80000
#End If

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByVal ptrToLogFont As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As DT_Flags) As Long
Private Declare Function ExtTextOutNoRect Lib "gdi32" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByVal fakeRectPointer As Long, ByVal lpStrPointer As Long, ByVal nCount As Long, ByRef lpDx As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpStrPointer As Long, ByVal cbString As Long, ByRef lpSize As PointAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As GDI_TextAlign) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long

'Used to track the created font
Private m_hFont As Long, m_oldHFont As Long

'Used to track the DC we are rendering to
Private m_hDC As Long

'Current text alignment (as an API constant, not a VB one)
Private m_TextAlignment As Long

'Current font color
Private m_FontColor As Long

'This LOGFONT object stores all current settings for the font
Private m_currentLogFont As LOGFONT

'Certain settings have to be modified before placing them inside a LOGFONT struct.  We store local copies of these values in VB-friendly formats,
' which simplies the process of checking these properties prior to replacing them.
Private m_origFontFace As String
Private m_origFontSize As Single
Private m_origFontBold As Boolean
Private m_origFontItalic As Boolean

'If the user wants to call DrawText manually, they can use this function.  Note that hDC is *not* part of the call, as it's assumed
' the caller has already selected this object into the relevant DC!
Friend Function DrawTextWrapper(ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
    DrawText m_hDC, lpStr, nCount, lpRect, wFormat
End Function

'This function draws horizontally and vertically centered text onto the attached DC, and it does it without modifying any of the current
' font settings (including alignment).  It is designed for one-off renderings of centered text.
Friend Sub DrawCenteredText(ByRef srcString As String, Optional ByVal customWidth As Long = -1, Optional ByVal customHeight As Long = -1)

    'Because DCs are dimensionless abstractions, we need to retrieve the attached window and use it to calculate the output rect.
    ' (Note that this does not work when rendering to a windowless DC, like a DIB.)
    Dim winClientRect As winRect
    GetClientRect WindowFromDC(m_hDC), winClientRect
    
    'Use the client rectangle to determine a width and height midpoint for the target DC
    Dim halfDCWidth As Long, halfDCHeight As Long
    If (winClientRect.x2 = 0) Then
        halfDCWidth = customWidth \ 2
        halfDCHeight = customHeight \ 2
    Else
        halfDCWidth = winClientRect.x2 \ 2
        halfDCHeight = winClientRect.y2 \ 2
    End If
    
    'Create a bounding rect for the text
    Dim tmpRect As RECT
    
    tmpRect.Left = halfDCWidth
    tmpRect.Right = tmpRect.Left + GetWidthOfString(srcString)
    tmpRect.Top = halfDCHeight
    tmpRect.Bottom = tmpRect.Top + GetHeightOfWordwrapString(srcString, GetWidthOfString(srcString))
    
    'DrawText handles the actual rendering
    DrawText m_hDC, StrPtr(srcString), Len(srcString), tmpRect, DT_WORDBREAK Or DT_NOCLIP
    
End Sub

'This function draws horizontally and vertically centered text within a specified rect, and it does it without modifying any of the current
' font settings (including alignment).  It is designed for one-off renderings of centered text.
Friend Sub DrawCenteredTextToRect(ByRef srcString As String, ByRef dstRect As RECT, Optional ByVal centerVertically As Boolean = False)
    
    Dim dtFlags As Long
    dtFlags = DT_CENTER
    
    If centerVertically Then
        dtFlags = dtFlags Or DT_VCENTER Or DT_SINGLELINE
    Else
        dtFlags = dtFlags Or DT_WORDBREAK
    End If
    
    'Rendering this style of text is pretty damn easy
    DrawText m_hDC, StrPtr(srcString), Len(srcString), dstRect, dtFlags
    
End Sub

'If the font has already been created and set to a DC, this function can be used to quickly render text (because no
' additional parameters have to be set).  Note that it only works for left-aligned text, as no contraining rect is used.
Friend Sub FastRenderText(ByVal dstX As Long, ByVal dstY As Long, ByRef srcString As String)
    ExtTextOutNoRect m_hDC, dstX, dstY, 0, ByVal 0&, StrPtr(srcString), Len(srcString), ByVal 0&
End Sub

'A roughly identical function to fastRenderText, above, but with a constraining rect
Friend Sub FastRenderTextWithClipping(ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcString As String, Optional ByVal useEllipses As Boolean = True, Optional ByVal centerVertically As Boolean = False, Optional ByVal treatAmpAsControlChar As Boolean = True)
    
    'IMPORTANT NOTE!  Per MSDN, the following caveat is made regarding the source string pointer:
    '  "A pointer to the string that specifies the text to be drawn. If the nCount parameter is -1, the string must
    '   be null-terminated.  If uFormat includes DT_MODIFYSTRING, the function could add up to four additional characters
    '   to this string. The buffer containing the string should be large enough to accommodate these extra characters."
    '
    'Because the string may be modified to fit within the clipping boundaries, we need to extend the string array by four chars.
    Dim srcStringLen As Long
    srcStringLen = Len(srcString)
    
    If useEllipses Then srcString = srcString & String$(4, ChrW$(0))
    
    Dim tmpRect As RECT
    tmpRect.Left = dstX
    tmpRect.Top = dstY
    tmpRect.Right = dstX + dstWidth
    tmpRect.Bottom = dstY + dstHeight
    
    'If the text is going to extend past the target rect, draw it with ellipses.  Note that we deliberately misrepresent the
    ' actual length of the string in that case.  This is on purpose, to account for the DT_MODIFYSTRING flag above.  (If the
    ' string is too long to fit inside the clipping rect, DrawText may choose to extend the character count of the string.
    ' As such, we pass only the currently relevant length of the string, without our added null chars, and DrawText can use
    ' the extra buffer space as necessary.)
    Dim dtFlags As DT_Flags
    If centerVertically Then dtFlags = DT_VCENTER Else dtFlags = DT_TOP
    If (Not treatAmpAsControlChar) Then dtFlags = dtFlags Or DT_NOPREFIX
    
    Const ELLIPSES_METHOD As Long = DT_WORD_ELLIPSIS
    
    Select Case m_TextAlignment
        
        Case ta_LEFT
            If useEllipses Then
                dtFlags = dtFlags Or ELLIPSES_METHOD Or DT_MODIFYSTRING Or DT_LEFT
            Else
                dtFlags = dtFlags Or DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP
            End If
            
        Case ta_RIGHT
            If useEllipses Then
                dtFlags = dtFlags Or ELLIPSES_METHOD Or DT_MODIFYSTRING Or DT_RIGHT
            Else
                dtFlags = dtFlags Or DT_SINGLELINE Or DT_RIGHT Or DT_NOCLIP
            End If
            
        Case ta_CENTER
            If useEllipses Then
                dtFlags = dtFlags Or ELLIPSES_METHOD Or DT_MODIFYSTRING Or DT_CENTER
            Else
                dtFlags = dtFlags Or DT_SINGLELINE Or DT_CENTER Or DT_NOCLIP
            End If
        
    End Select
    
    DrawText m_hDC, StrPtr(srcString), srcStringLen, tmpRect, dtFlags
    
End Sub

'A roughly identical function to fastRenderText, above, but with a constraining rect and expansion to multiple lines as necessary
Friend Sub FastRenderMultilineTextWithClipping(ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcString As String, Optional ByVal centerVertically As Boolean = False, Optional ByVal preProcessAmps As Boolean = True)
    
    'Multiline text does not need ellipses, so this function is actually simpler than its single-line counterpart
    Dim srcStringLen As Long
    srcStringLen = Len(srcString)
    
    Dim tmpRect As RECT
    tmpRect.Left = dstX
    tmpRect.Top = dstY
    tmpRect.Right = dstX + dstWidth
    tmpRect.Bottom = dstY + dstHeight
    
    'When rendering multiline text, we want top-alignment by default (vertical centering is handled manually).
    ' We must also specify word-breaking, and the caller can control whether ampersands are treated as control chars.
    Dim dtFlags As DT_Flags
    dtFlags = DT_WORDBREAK Or DT_TOP
    If (Not preProcessAmps) Then dtFlags = dtFlags Or DT_NOPREFIX
    
    'Vertical centering isn't automatic when multiline text is used.  Instead, we need to manually calculate centering.
    If centerVertically Then
        Dim tmpHeight As Single
        tmpHeight = Me.GetHeightOfWordwrapString(srcString, dstWidth)
        tmpRect.Top = dstY + (dstHeight - tmpHeight) \ 2
    End If
    
    Select Case m_TextAlignment
        
        Case ta_LEFT
            dtFlags = dtFlags Or DT_LEFT
            
        Case ta_RIGHT
            dtFlags = dtFlags Or DT_RIGHT
            
        Case ta_CENTER
            dtFlags = dtFlags Or DT_CENTER
        
    End Select
    
    DrawText m_hDC, StrPtr(srcString), srcStringLen, tmpRect, dtFlags
    
End Sub

'A roughly identical function to fastRenderText, above, but with support for linebreaks
Friend Sub FastRenderMultilineText(ByVal dstX As Long, ByVal dstY As Long, ByRef srcString As String)
        
    'Find the bounding rect of the text
    Dim tmpRect As RECT
    
    'Modify the rect to account for the specified offset
    tmpRect.Left = dstX
    tmpRect.Top = dstY
    tmpRect.Right = dstX + GetWidthOfString(srcString)
    tmpRect.Bottom = dstY + GetHeightOfWordwrapString(srcString, GetWidthOfString(srcString))
    
    'If the text is going to extend past the target rect, draw it with ellipses
    DrawText m_hDC, StrPtr(srcString), Len(srcString), tmpRect, DT_LEFT Or DT_TOP Or DT_WORDBREAK
    
End Sub

'Select the current font and other text settings into the specified DC.  Ideally, this should only be done when absolutely
' necessary to keep speed high.  (Also, note the DC so we can do fast font rendering later.)
Friend Sub AttachToDC(ByVal targetDC As Long)
    
    If (targetDC = 0) Then Debug.Print "WARNING!  pdFont.AttachToDC() was called with targetDC = 0.  Fix this!"
    If (m_hDC <> 0) Then Debug.Print "WARNING!  pdFont.AttachToDC() was called, but this font is already selected into a DC!"
    If (targetDC = m_hDC) Then Debug.Print "WARNING!  pdFont.AttachToDC() was called, but the specified DC is already the current DC!"
    
    'Create the font, as necessary
    If (m_hFont = 0) Then Me.CreateFontObject
    
    'Set the font
    If (targetDC <> m_hDC) Then
        m_oldHFont = SelectObject(targetDC, m_hFont)
        m_hDC = targetDC
    End If
    
    'Set other font parameters
    SetTextColor targetDC, m_FontColor
    SetTextAlign targetDC, m_TextAlignment
    
    'Enable transparent font rendering
    SetBkMode targetDC, FONT_TRANSPARENT
    
End Sub

'If we want to detach this font from its current DC, use this call
Friend Sub ReleaseFromDC()
    If (m_hDC <> 0) Then
        SelectObject m_hDC, m_oldHFont
        m_hDC = 0
    End If
End Sub

'Create the font object on-demand.
' IMPORTANT NOTE: the "createWithEscapement" parameter is *not* well-tested.  PD uses it infrequently,
' in very specialized circumstances, and I can't guarantee that caches and other font interfaces will work
' correctly with it.  Generally speaking, rotated fonts should be cached specially.
Friend Function CreateFontObject(Optional ByVal createWithAngle As Long = 0) As Boolean

    'If a font has already been created, delete it now
    If (m_hDC <> 0) Then Me.ReleaseFromDC
    If (m_hFont <> 0) Then Me.DeleteCurrentFont
    
    If (createWithAngle <> 0) Then
        m_currentLogFont.lfEscapement = createWithAngle
        m_currentLogFont.lfOrientation = createWithAngle
    End If
    
    m_hFont = CreateFontIndirect(VarPtr(m_currentLogFont))
    CreateFontObject = (m_hFont <> 0)
    
    'Inside debug mode, update a global counter (for leak tracking)
    If CreateFontObject Then PDDebug.UpdateResourceTracker PDRT_hFont, 1
    
End Function

'Shortcut function to set all font properties at once
Friend Sub SetFontPropsAllAtOnce(ByRef newFace As String, ByVal newFontSize As Single, ByVal newBold As Boolean, ByVal newItalic As Boolean, ByVal newUnderline As Boolean)
    
    'First, cache all persistent settings
    m_origFontFace = newFace
    m_origFontSize = newFontSize
    m_origFontBold = newBold
    m_origFontItalic = newItalic
    
    'Next, translate these into actual LOGFONT properties
    
    'Make sure we have a trailing null char before copying the font name into place
    If (LenB(newFace) < 64) Then
        FillMemory VarPtr(m_currentLogFont.lfFaceName(0)), 64, 0
        CopyMemoryStrict VarPtr(m_currentLogFont.lfFaceName(0)), StrPtr(newFace), LenB(newFace)
    Else
        CopyMemoryStrict VarPtr(m_currentLogFont.lfFaceName(0)), StrPtr(newFace), 64
    End If
    
    m_currentLogFont.lfHeight = Fonts.ConvertToGDIFontSize(newFontSize)
    If newBold Then m_currentLogFont.lfWeight = fw_Bold Else m_currentLogFont.lfWeight = fw_Normal
    m_currentLogFont.lfItalic = -1 * newItalic
    m_currentLogFont.lfUnderline = -1 * newUnderline
    
End Sub

Friend Function GetFontBold() As Boolean
    GetFontBold = m_origFontBold
End Function

Friend Sub SetFontBold(ByVal newBold As Boolean)
    m_origFontBold = newBold
    If newBold Then m_currentLogFont.lfWeight = fw_Bold Else m_currentLogFont.lfWeight = fw_Normal
End Sub

'If we are currently selected into a DC, update the color immediately; otherwise, the color will be applied upon DC selection
Friend Function GetFontColor() As Long
    GetFontColor = m_FontColor
End Function

Friend Sub SetFontColor(ByVal newFontColor As Long)
    m_FontColor = newFontColor
    If (m_hDC <> 0) Then SetTextColor m_hDC, m_FontColor
End Sub

Friend Function GetFontFace() As String
    GetFontFace = m_origFontFace
End Function

Friend Sub SetFontFace(ByRef newFace As String)
    
    m_origFontFace = newFace
    
    'Make sure we have a trailing null char before copying the font name into place
    If (LenB(newFace) < 64) Then
        FillMemory VarPtr(m_currentLogFont.lfFaceName(0)), 64, 0
        CopyMemoryStrict VarPtr(m_currentLogFont.lfFaceName(0)), StrPtr(newFace), LenB(newFace)
    Else
        CopyMemoryStrict VarPtr(m_currentLogFont.lfFaceName(0)), StrPtr(newFace), 64
    End If
    
End Sub

Friend Function GetFontItalic() As Boolean
    GetFontItalic = m_origFontItalic
End Function

Friend Sub SetFontItalic(ByVal newItalic As Boolean)
    m_origFontItalic = newItalic
    m_currentLogFont.lfItalic = -1 * newItalic
End Sub

Friend Function GetFontSize() As Single
    GetFontSize = m_origFontSize
End Function

Friend Sub SetFontSize(ByVal newFontSize As Single)
    m_origFontSize = newFontSize
    m_currentLogFont.lfHeight = Fonts.ConvertToGDIFontSize(newFontSize)
End Sub

Friend Sub SetFontStrikeout(ByVal newStrikeout As Boolean)
    m_currentLogFont.lfStrikeOut = -1 * newStrikeout
End Sub

Friend Function GetFontStrikeout() As Boolean
    GetFontStrikeout = (m_currentLogFont.lfStrikeOut * -1)
End Function

Friend Sub SetFontUnderline(ByVal newUnderline As Boolean)
    m_currentLogFont.lfUnderline = -1 * newUnderline
End Sub

Friend Function GetFontUnderline() As Boolean
    GetFontUnderline = (m_currentLogFont.lfUnderline * -1)
End Function

Friend Sub SetTextAlignment(ByVal newAlignment As AlignmentConstants)
    
    Select Case newAlignment
        Case vbLeftJustify
            m_TextAlignment = ta_LEFT
        Case vbCenter
            m_TextAlignment = ta_CENTER
        Case vbRightJustify
            m_TextAlignment = ta_RIGHT
    End Select
    
    'All alignment settings are handled via flags to DrawText.  For DrawText flags to work, the alignment of the current DC must be
    ' set to TA_LEFT, TA_TOP, and TA_NOUPDATECP (see http://msdn.microsoft.com/en-us/library/dd162498.aspx)
    If (m_hDC <> 0) Then SetTextAlign m_hDC, ta_LEFT Or ta_TOP Or ta_NOUPDATECP
    
End Sub

'Set text alignment directly.  Note that this can fuck with DrawText - see SetTextAlignment, above, for details!
Friend Sub SetTextAlignmentEx(ByVal newAlignment As GDI_TextAlign)
    If (m_hDC <> 0) Then SetTextAlign m_hDC, newAlignment
End Sub

'Return the width (and below, height) of a string, in pixels.  This action REQUIRES that a DC be currently set for this
' object, as the DC's font settings are used to determine the string's dimensions.
Friend Function GetWidthOfString(ByRef srcString As String) As Long
    
    Dim txtSize As PointAPI
    
    'If we're not currently selected into a DC, silently plug in a temporary DC
    Dim targetDC As Long, tmpDCActive As Boolean
    If (m_hDC = 0) Then
        tmpDCActive = True
        targetDC = GDI.GetMemoryDC()
        AttachToDC targetDC
    Else
        targetDC = m_hDC
    End If
    
    'Retrieve the string's dimensions
    GetTextExtentPoint32 targetDC, StrPtr(srcString), Len(srcString), txtSize
    GetWidthOfString = txtSize.x
    
    'Release the temporary DC, as necessary
    If tmpDCActive Then
        ReleaseFromDC
        GDI.FreeMemoryDC targetDC
    End If
    
End Function

'With the current font settings, find the height of either a single-line string, or a word-wrapped string.
' For convenience, I've separated these into separate functions, as the underlying logic is actually quite different.
Friend Function GetHeightOfString(ByRef srcString As String) As Long
    
    'If we're not currently selected into a DC, silently plug in a temporary DC
    Dim targetDC As Long, tmpDCActive As Boolean
    If (m_hDC <> 0) Then
        targetDC = m_hDC
    Else
        tmpDCActive = True
        targetDC = GDI.GetMemoryDC()
        AttachToDC targetDC
    End If
    
    'Retrieve the string's dimensions
    Dim txtSize As PointAPI
    GetTextExtentPoint32 targetDC, StrPtr(srcString), Len(srcString), txtSize
    GetHeightOfString = txtSize.y
    
    'Release the temporary DC, as necessary
    If tmpDCActive Then
        ReleaseFromDC
        GDI.FreeMemoryDC targetDC
    End If
    
End Function

Friend Function GetHeightOfWordwrapString(ByRef srcString As String, ByVal strWidth As Long) As Long

    Dim tmpRect As RECT
    tmpRect.Left = 0
    tmpRect.Top = 0
    tmpRect.Bottom = 0
    tmpRect.Right = strWidth - 1
    
    'If we're not currently selected into a DC, silently plug in a temporary DC
    Dim targetDC As Long, tmpDCActive As Boolean
    If (m_hDC <> 0) Then
        targetDC = m_hDC
    Else
        tmpDCActive = True
        targetDC = GDI.GetMemoryDC()
        AttachToDC targetDC
    End If
    
    'Retrieve the string's dimensions
    DrawText targetDC, StrPtr(srcString), Len(srcString), tmpRect, DT_LEFT Or DT_CALCRECT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_NOPREFIX Or DT_TOP
    GetHeightOfWordwrapString = tmpRect.Bottom
    
    'Release the temporary DC, as necessary
    If tmpDCActive Then
        ReleaseFromDC
        GDI.FreeMemoryDC targetDC
    End If

End Function

'This is just a thin wrapper to DrawText+DT_CALCRECT.  Note the (many) caveats, per MSDN:
' "Determines the width and height of the rectangle. If there are multiple lines of text, DrawText uses the
'  width of the rectangle pointed to by the lpRect parameter and extends the base of the rectangle to bound
'  the last line of text. If the largest word is wider than the rectangle, the width is expanded. If the text
'  is less than the width of the rectangle, the width is reduced. If there is only one line of text, DrawText
'  modifies the right side of the rectangle so that it bounds the last character in the line. In either case,
'  DrawText returns the height of the formatted text but does not draw the text."
Friend Function GetBoundaryRectOfMultilineString(ByVal srcString As String, ByVal maxAllowedWidth As Long, ByRef dstRect As RectL, Optional ByVal turnOffPrefixHandling As Boolean = False) As Boolean

    Dim tmpRect As RECT
    tmpRect.Left = 0
    tmpRect.Top = 0
    tmpRect.Bottom = 0
    tmpRect.Right = maxAllowedWidth - 1
    
    'If we're not currently selected into a DC, silently plug in a temporary DC
    Dim targetDC As Long, tmpDCActive As Boolean
    If m_hDC = 0 Then
        tmpDCActive = True
        targetDC = GDI.GetMemoryDC()
        AttachToDC targetDC
    Else
        targetDC = m_hDC
    End If
    
    'Create the magical combination of DrawText flags
    Dim dtFlags As Long
    dtFlags = DT_LEFT Or DT_CALCRECT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_TOP
    If turnOffPrefixHandling Then dtFlags = dtFlags Or DT_NOPREFIX
    
    'Retrieve the string's dimensions
    GetBoundaryRectOfMultilineString = (DrawText(targetDC, StrPtr(srcString), Len(srcString), tmpRect, dtFlags) <> 0)
    
    With dstRect
        .Left = tmpRect.Left
        .Top = tmpRect.Top
        .Right = tmpRect.Right
        .Bottom = tmpRect.Bottom
    End With
    
    'Release the temporary DC, as necessary
    If tmpDCActive Then
        ReleaseFromDC
        GDI.FreeMemoryDC targetDC
    End If

End Function

Private Sub Class_Initialize()
    
    'When the class is initialized, set up a default PhotoDemon font.  Face is automatically set per
    ' the current OS, so external functions do not need to set this manually.
    With m_currentLogFont
    
        .lfHeight = Fonts.ConvertToGDIFontSize(11!)
        .lfWidth = 0        'This will automatically set the width based on the height
        
        .lfEscapement = 0   'Used for rotation; not utilized at present
        .lfOrientation = 0  'An additional rotation parameter; not utilized at present
        
        .lfWeight = fw_Regular
        
        .lfItalic = False
        .lfUnderline = False
        .lfStrikeOut = False
        
        .lfCharSet = DEFAULT_CHARSET
        
        .lfOutPrecision = 0
        .lfClipPrecision = 0
        .lfQuality = fq_ClearType
        
        'Ideally, we should allow the caller to specify pitch and family requests.  I may add this
        ' in the future, but for now we want to fallback to a variable width, non-serifed font if
        ' (for some reason) Segoe and Tahoma aren't available on the target machine.
        Const VARIABLE_PITCH As Long = 2
        Const FF_SWISS As Long = 32       'Variable stroke width + sans-serif
        .lfPitchAndFamily = VARIABLE_PITCH Or FF_SWISS
        
        If PDMain.IsProgramRunning() Then
            m_origFontFace = Fonts.GetUIFontName()
        Else
            m_origFontFace = "Tahoma"
        End If
        
        'Make sure we have a trailing null char before copying the font name into place
        If (LenB(m_origFontFace) < 64) Then
            FillMemory VarPtr(m_currentLogFont.lfFaceName(0)), 64, 0
            CopyMemoryStrict VarPtr(m_currentLogFont.lfFaceName(0)), StrPtr(m_origFontFace), LenB(m_origFontFace)
        Else
            CopyMemoryStrict VarPtr(m_currentLogFont.lfFaceName(0)), StrPtr(m_origFontFace), 64
        End If
        
    End With
    
    m_hFont = 0
    m_hDC = 0
        
    m_origFontBold = False
    m_origFontItalic = False
    m_origFontSize = 11
    
    m_FontColor = 0
    m_TextAlignment = ta_LEFT
        
End Sub

'Like most things, it's ideal if the caller frees resources when they no longer need them, but this class should successfully
' release all attached GDI objects regardless of caller behavior.
Private Sub Class_Terminate()
    If (m_hDC <> 0) Then Me.ReleaseFromDC
    DeleteCurrentFont
End Sub

Friend Function HasFontBeenCreated() As Boolean
    HasFontBeenCreated = (m_hFont <> 0)
End Function

Friend Function GetAttachedDC() As Long
    GetAttachedDC = m_hDC
End Function

Friend Function GetFontHandle() As Long
    GetFontHandle = m_hFont
End Function

Friend Sub DeleteCurrentFont()
    
    'If a font was created, delete it now
    If (m_hFont <> 0) Then
        
        'If the font is inside a DC, remove it (otherwise we can't delete it)
        If (m_hDC <> 0) Then
            PDDebug.LogAction "HEY!  Font was not removed from DC - please do that before deleting it!"
            Me.ReleaseFromDC
        End If
        
        'Inside debug mode, update a global counter (for leak tracking)
        If (DeleteObject(m_hFont) <> 0) Then
            PDDebug.UpdateResourceTracker PDRT_hFont, -1
            m_hFont = 0
        Else
            PDDebug.LogAction "WARNING! Could not delete font resource - is it selected into a DC??"
        End If
        
    End If
    
End Sub

'Quick and dirty function for checking font existence.  Does not use API at present, so fonts with Unicode names won't return correct results.
' I haven't done a check, but I expect this to be faster than enum-ing all available fonts and looking for a match...
Friend Function DoesFontExist(ByVal fontNameToCheck As String, Optional ByRef fontSubstitute As String, Optional ByRef nearestFontSize As Long = 8) As Boolean

    'I don't know if this function will throw errors, but better safe than sorry...
    On Error Resume Next
    
    'Create a temporary font object to test name support
    Dim tmpFont As StdFont
    Set tmpFont = New StdFont
    
    'Copy name and size (if any)
    tmpFont.Name = fontNameToCheck
    tmpFont.Size = nearestFontSize
    
    'See if the font object has the same name.  Windows will automatically substitute a new font if the requested one can't be found.
    DoesFontExist = Strings.StringsEqual(fontNameToCheck, tmpFont.Name, True)
    
    'Special check for vertical fonts
    If (Not DoesFontExist) And (InStr(1, fontNameToCheck, "@", vbBinaryCompare) <> 0) Then
        DoesFontExist = Strings.StringsEqual(Right$(fontNameToCheck, Len(fontNameToCheck) - 1), tmpFont.Name, True)
    End If
    
    'Not sure if helpful, but return the replacement font name and nearest supported size
    fontSubstitute = tmpFont.Name
    nearestFontSize = tmpFont.Size
    
End Function

'Given a desired width (in pixels) and a source string, search for the largest font size that allows said string to fit within that width.
' Word-wrapping is not used, by design, as it's assumed the caller already tried word-wrapping as a solution.
'
'RETURNS: largest font size that fits the string into the desired width.  (7 is the smallest size this function will return, FYI.)
Friend Function GetMaxFontSizeToFitStringWidth(ByRef srcString As String, ByVal desiredWidth As Long, Optional ByVal startingFontSize As Long = 10) As Single
    
    'Start by determining a targetDC.  A temporary DC will be created and used, as necessary.
    Dim targetDC As Long, backupDC As Long
    backupDC = m_hDC
    
    If (m_hDC = 0) Then
        targetDC = GDI.GetMemoryDC()
    Else
        ReleaseFromDC
        targetDC = backupDC
    End If
    
    'Make a backup of the current font size, if any
    Dim backupFontSize As Single
    backupFontSize = m_origFontSize
    
    'Create the font anew, using the default string size we were passed
    If (startingFontSize <> m_origFontSize) Or (m_hFont = 0) Then
        If (m_hFont <> 0) Then DeleteCurrentFont
        m_origFontSize = startingFontSize
        SetFontSize m_origFontSize
        CreateFontObject
    End If
    
    'Select the font into the target DC
    AttachToDC targetDC
    
    'Measure the font relative to the current control size
    Dim stringWidth As Long
    stringWidth = GetWidthOfString(srcString)
    
    'If the string does not fit within the control size, shrink the font accordingly.
    Do While (stringWidth > desiredWidth) And (m_origFontSize >= 7!)
        
        'Shrink the font size
        m_origFontSize = m_origFontSize - 0.5!
        
        'Recreate the font
        ReleaseFromDC
        DeleteCurrentFont
        Me.SetFontSize m_origFontSize
        Me.CreateFontObject
        Me.AttachToDC targetDC
        
        'Measure the new size
        stringWidth = GetWidthOfString(srcString)
        
    Loop
    
    'By this point, m_origFontSize represents the valid size.  Return it.
    GetMaxFontSizeToFitStringWidth = m_origFontSize
    
    'Before exiting, restore any parameters we changed
    ReleaseFromDC
    
    If (m_origFontSize <> backupFontSize) Then
        DeleteCurrentFont
        SetFontSize backupFontSize
        CreateFontObject
    End If
    
    'If we created a temporary DC, free it now
    If (targetDC <> backupDC) Then
        GDI.FreeMemoryDC targetDC
    
    'Conversely, if we were originally selected into a DC, restore that DC now
    Else
        If (backupDC <> 0) Then AttachToDC backupDC
    End If
    
End Function

'Given a desired width and height (in pixels) and a source string, search for the largest font size that allows said string to fit within
' that rect.  Word-wrapping will be used.
'
'RETURNS: largest font size that fits the string into the desired width.  (7 is the smallest size this function will return, FYI.)
Friend Function GetMaxFontSizeToFitWordWrap(ByRef srcString As String, ByVal desiredWidth As Long, ByVal desiredHeight As Long, Optional ByVal startingFontSize As Long = 10) As Single
    
    'Start by determining a targetDC.  A temporary DC will be created and used, as necessary.
    Dim targetDC As Long, backupDC As Long
    backupDC = m_hDC
    
    If (m_hDC = 0) Then
        targetDC = GDI.GetMemoryDC()
    Else
        ReleaseFromDC
        targetDC = backupDC
    End If
    
    'Make a backup of the current font size, if any
    Dim backupFontSize As Single
    backupFontSize = m_origFontSize
    
    'Create the font anew, using the default string size we were passed
    ReleaseFromDC
    m_origFontSize = startingFontSize
    SetFontSize m_origFontSize
    CreateFontObject
    AttachToDC targetDC
    
    'Measure the font relative to the current control size
    Dim stringHeight As Long
    stringHeight = GetHeightOfWordwrapString(srcString, desiredWidth)
    
    'If the string does not fit within the control size, shrink the font accordingly.
    Do While (stringHeight > desiredHeight) And (m_origFontSize >= 7!)
        
        'Shrink the font size
        m_origFontSize = m_origFontSize - 0.5!
        
        'Recreate the font
        ReleaseFromDC
        SetFontSize m_origFontSize
        CreateFontObject
        AttachToDC targetDC
        
        'Measure the new size
        stringHeight = GetHeightOfWordwrapString(srcString, desiredWidth)
        
    Loop
    
    'By this point, m_origFontSize represents the valid size.  Return it.
    GetMaxFontSizeToFitWordWrap = m_origFontSize
    
    'Before exiting, restore any parameters we changed
    ReleaseFromDC
    
    If (m_origFontSize <> backupFontSize) Then
        SetFontSize backupFontSize
        CreateFontObject
    End If
    
    'If we created a temporary DC, free it now
    If (targetDC <> backupDC) Then
        GDI.FreeMemoryDC targetDC
    
    'Conversely, if we were originally selected into a DC, restore that DC now
    Else
        If (backupDC <> 0) Then AttachToDC backupDC
    End If
    
End Function
