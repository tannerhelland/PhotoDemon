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
'Copyright ©2013-2014 by Tanner Helland
'Created: 31/May/13
'Last updated: 29/October/14
'Last update: many small improvements to make this class work better with the new pdLabel control.
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
'Obviously, this class relies heavily on WAPI.  Functions are documented to the best of my knowledge and ability.
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

'Previously we only suppored an ANSI charset, but now we want to default to the system value
Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1

'Pitch and family constants
Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2
Private Const FF_DONTCARE = 0    'Any family works
Private Const FF_ROMAN = 16      'Variable stroke width + serifs
Private Const FF_SWISS = 32      'Variable stroke width + sans-serif
Private Const FF_MODERN = 48     'Constant stroke width + serif AND sans-serif
Private Const FF_SCRIPT = 64
Private Const FF_DECORATIVE = 80

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

'Font precision
Private Const OUT_DEFAULT_PRECIS As Long = 0
Private Const OUT_CHARACTER_PRECIS As Long = 2
Private Const OUT_DEVICE_PRECIS As Long = 5

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

Private Const TA_TOP = 0
Private Const TA_BOTTOM = 8
Private Const TA_BASELINE = 24

Private Const TA_UPDATECP = 1
Private Const TA_NOUPDATECP = 0

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
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long

Private Declare Function ExtTextOutNoRect Lib "gdi32" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByVal fakeRectPointer As Long, ByVal lpStrPointer As Long, ByVal nCount As Long, ByRef lpDx As Long) As Long
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByRef lpRect As RECT, ByVal lpStrPointer As Long, ByVal nCount As Long, ByRef lpDx As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long

'Used to measure the expected length of a string (single line only; line breaks do not work)
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpStrPointer As Long, ByVal cbString As Long, ByRef lpSize As POINTAPI) As Long

'Used to track the created font
Private curFont As Long, oldFont As Long

'Used to track the DC we are rendering to
Private curDC As Long

'When the class is initialized, this value is recorded to help us convert from points to ems
Private curLogPixelsY As Long

'Current text alignment (as an API constant, not a VB one)
Private curTextAlignment As Long

'Current font color
Private fontColor As Long

'This LOGFONT object stores all current settings for the font
Private currentLogFont As LOGFONT

'Certain settings have to be modified before placing them inside a LOGFONT struct.  We store local copies of these values in VB-friendly formats,
' which simplies the process of checking these properties prior to replacing them.
Private m_origFontFace As String
Private m_origFontSize As Single
Private m_origFontBold As Boolean
Private m_origFontItalic As Boolean

'This function draws horizontally and vertically centered text onto the attached DC, and it does it without modifying any of the current
' font settings (including alignment).  It is designed for one-off renderings of centered text.
Public Sub drawCenteredText(ByRef srcString As String, Optional ByVal customWidth As Long = -1, Optional ByVal customHeight As Long = -1)

    'Check for multiline text first; it requires special handling
    Dim textIsMultiline As Boolean
    If InStr(1, srcString, vbCrLf, vbBinaryCompare) > 0 Then
        textIsMultiline = True
    Else
        textIsMultiline = False
    End If
    
    'Because DCs are dimensionless abstractions, we need to retrieve the attached window and use it to calculate the output rect.
    ' (Note that this does not work when rendering to a windowless DC, like a DIB.)
    Dim winClientRect As winRect
    GetClientRect WindowFromDC(curDC), winClientRect
    
    'Use the client rectangle to determine a width and height midpoint for the target DC
    Dim halfDCWidth As Long, halfDCHeight As Long
    If winClientRect.x2 = 0 Then
        halfDCWidth = customWidth \ 2
        halfDCHeight = customHeight \ 2
    Else
        halfDCWidth = winClientRect.x2 \ 2
        halfDCHeight = winClientRect.y2 \ 2
    End If
    
    'Next, determine the height of the string we were passed
    Dim strHeightHalf As Long
    If textIsMultiline Then
        strHeightHalf = getHeightOfWordwrapString(srcString, getWidthOfString(srcString)) \ 2
    Else
        strHeightHalf = getHeightOfString(srcString) \ 2
    End If
    
    'Back-up the current text alignment, then change it to horizontally aligned
    Dim backupAlignment As Long
    backupAlignment = curTextAlignment
    curTextAlignment = TA_CENTER
    
    'Render the text
    If textIsMultiline Then
    
        'Create a bounding rect for the text
        Dim tmpRect As RECT
        
        tmpRect.Left = halfDCWidth
        tmpRect.Right = tmpRect.Left + getWidthOfString(srcString)
        tmpRect.Top = halfDCHeight
        tmpRect.Bottom = tmpRect.Top + getHeightOfWordwrapString(srcString, getWidthOfString(srcString))
        
        'If the text is going to extend past the target rect, draw it with ellipses
        DrawText curDC, StrPtr(srcString), Len(srcString), tmpRect, DT_WORDBREAK Or DT_NOCLIP
    
    Else
        ExtTextOutNoRect curDC, halfDCWidth, halfDCHeight - strHeightHalf, 0, ByVal 0&, StrPtr(srcString), Len(srcString), ByVal 0&
    End If
    
    'Restore the original text alignment
    curTextAlignment = backupAlignment

End Sub

'This function draws horizontally and vertically centered text within a specified rect, and it does it without modifying any of the current
' font settings (including alignment).  It is designed for one-off renderings of centered text.
Friend Sub drawCenteredTextToRect(ByVal srcString As String, ByRef dstRect As RECT)
    
    'Rendering this style of text is pretty damn easy
    DrawText curDC, StrPtr(srcString), Len(srcString), dstRect, DT_CENTER Or DT_WORDBREAK
    
End Sub

'If the font has already been created and set to a DC, this function can be used to quickly render text (because no
' additional parameters have to be set)
Public Sub fastRenderText(ByVal dstX As Long, ByVal dstY As Long, ByRef srcString As String)
    ExtTextOutNoRect curDC, dstX, dstY, 0, ByVal 0&, StrPtr(srcString), Len(srcString), ByVal 0&
End Sub

'A roughly identical function to fastRenderText, above, but with a constraining rect
Public Sub fastRenderTextWithClipping(ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcString As String, Optional ByVal useEllipses As Boolean = True)
    
    'IMPORTANT NOTE!  Per MSDN, the following caveat is made regarding the source string pointer:
    '  "A pointer to the string that specifies the text to be drawn. If the nCount parameter is -1, the string must
    '   be null-terminated.  If uFormat includes DT_MODIFYSTRING, the function could add up to four additional characters
    '   to this string. The buffer containing the string should be large enough to accommodate these extra characters."
    '
    'Because the string may be modified to fit within the clipping boundaries, we need to extend the string array by four chars.
    Dim srcStringLen As Long
    srcStringLen = Len(srcString)
    If useEllipses Then srcString = srcString & ChrW(0) & ChrW(0) & ChrW(0) & ChrW(0)
    
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
    Select Case curTextAlignment
        
        Case TA_LEFT
            If useEllipses Then
                DrawText curDC, StrPtr(srcString), srcStringLen, tmpRect, DT_WORD_ELLIPSIS Or DT_MODIFYSTRING Or DT_LEFT Or DT_TOP
            Else
                DrawText curDC, StrPtr(srcString), srcStringLen, tmpRect, DT_SINGLELINE Or DT_LEFT Or DT_TOP Or DT_NOCLIP
            End If
            
        Case TA_RIGHT
            If useEllipses Then
                DrawText curDC, StrPtr(srcString), srcStringLen, tmpRect, DT_WORD_ELLIPSIS Or DT_MODIFYSTRING Or DT_RIGHT Or DT_TOP
            Else
                DrawText curDC, StrPtr(srcString), srcStringLen, tmpRect, DT_SINGLELINE Or DT_RIGHT Or DT_TOP Or DT_NOCLIP
            End If
            
        Case TA_CENTER
            If useEllipses Then
                DrawText curDC, StrPtr(srcString), srcStringLen, tmpRect, DT_WORD_ELLIPSIS Or DT_MODIFYSTRING Or DT_CENTER Or DT_TOP
            Else
                DrawText curDC, StrPtr(srcString), srcStringLen, tmpRect, DT_SINGLELINE Or DT_CENTER Or DT_TOP Or DT_NOCLIP
            End If
        
    End Select
    
End Sub

'A roughly identical function to fastRenderText, above, but with a constraining rect and expansion to multiple lines as necessary
Public Sub fastRenderMultilineTextWithClipping(ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcString As String)
    
    'Multiline text does not need ellipses, so this function is actually simpler than its single-line counterpart
    Dim srcStringLen As Long
    srcStringLen = Len(srcString)
    
    Dim tmpRect As RECT
    tmpRect.Left = dstX
    tmpRect.Top = dstY
    tmpRect.Right = dstX + dstWidth
    tmpRect.Bottom = dstY + dstHeight
    
    'If the text is going to extend past the target rect, draw it with ellipses.  Note that we deliberately misrepresent the
    ' actual length of the string here.  This is on purpose, to account for the DT_MODIFYSTRING flag above.  (If the string
    ' is too long to fit inside the clipping rect, DrawText may choose to extend the character count of the string.  As such
    ' we pass only the currently relevant length of the string, without our added null chars, and DrawText can use the extra
    ' buffer space as necessary.)
    Select Case curTextAlignment
        
        Case TA_LEFT
            DrawText curDC, StrPtr(srcString), srcStringLen, tmpRect, DT_WORDBREAK Or DT_LEFT Or DT_TOP
            
        Case TA_RIGHT
            DrawText curDC, StrPtr(srcString), srcStringLen, tmpRect, DT_WORDBREAK Or DT_RIGHT Or DT_TOP
            
        Case TA_CENTER
            DrawText curDC, StrPtr(srcString), srcStringLen, tmpRect, DT_WORDBREAK Or DT_CENTER Or DT_TOP
        
    End Select
    
End Sub

'A roughly identical function to fastRenderText, above, but with support for linebreaks
Public Sub fastRenderMultilineText(ByVal dstX As Long, ByVal dstY As Long, ByRef srcString As String)
        
    'Find the bounding rect of the text
    Dim tmpRect As RECT
    
    'Modify the rect to account for the specified offset
    tmpRect.Left = dstX
    tmpRect.Top = dstY
    tmpRect.Right = dstX + getWidthOfString(srcString)
    tmpRect.Bottom = dstY + getHeightOfWordwrapString(srcString, getWidthOfString(srcString))
    
    'If the text is going to extend past the target rect, draw it with ellipses
    DrawText curDC, StrPtr(srcString), Len(srcString), tmpRect, DT_LEFT Or DT_TOP Or DT_WORDBREAK
    
End Sub

'Select the current font and other text settings into the specified DC.  Ideally, this should only be done when absolutely
' necessary to keep speed high.  (Also, note the DC so we can do fast font rendering later.)
Public Sub attachToDC(ByVal targetDC As Long)
    
    'Set the font
    If targetDC <> curDC Then
        oldFont = SelectObject(targetDC, curFont)
        curDC = targetDC
    End If
    
    'Set other font parameters
    SetTextColor targetDC, fontColor
    SetTextAlign targetDC, curTextAlignment
    
    'Enable transparent font rendering
    SetBkMode targetDC, FONT_TRANSPARENT
    
End Sub

'If we want to detach this font from its current DC, use this call
Public Sub releaseFromDC()
    If curDC <> 0 Then
        SelectObject curDC, oldFont
        curDC = 0
    End If
End Sub

'Create the font object on-demand
Public Function createFontObject() As Boolean

    'If a font has already been created, delete it now
    deleteCurrentFont
    
    curFont = CreateFontIndirect(currentLogFont)
    
    If curFont = 0 Then createFontObject = False Else createFontObject = True
    
    'Inside debug mode, increase a global font counter.  This is extremely helpful for finding font leaks
    #If DEBUGMODE = 1 Then
        If createFontObject Then g_FontsCreated = g_FontsCreated + 1
    #End If

End Function

Public Function getFontBold() As Boolean
    getFontBold = m_origFontBold
End Function

Public Sub setFontBold(ByVal newBold As Boolean)
    m_origFontBold = newBold
    If newBold Then
        currentLogFont.lfWeight = FW_BOLD
    Else
        currentLogFont.lfWeight = FW_NORMAL
    End If
End Sub

Public Sub setFontColor(ByVal newColor As Long)
    
    fontColor = newColor
    
    'If we are currently selected into a DC, update the color immediately; otherwise, the color will be applied upon DC selection
    If curDC <> 0 Then SetTextColor curDC, fontColor
    
End Sub

Public Function getFontFace() As String
    getFontFace = m_origFontFace
End Function

Public Sub setFontFace(ByVal newFace As String)
    m_origFontFace = newFace
    currentLogFont.lfFaceName = newFace & Chr(0)
End Sub

Public Function getFontItalic() As Boolean
    getFontItalic = m_origFontItalic
End Function

Public Sub setFontItalic(ByVal newItalic As Boolean)
    m_origFontItalic = newItalic
    currentLogFont.lfItalic = -1 * newItalic
End Sub

Public Function getFontSize() As Single
    getFontSize = m_origFontSize
End Function

Public Sub setFontSize(ByVal newFontSize As Single)
    m_origFontSize = newFontSize
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
    
    'All alignment settings are handled via flags to DrawText.  For DrawText flags to work, the alignment of the current DC must be
    ' set to TA_LEFT, TA_TOP, and TA_NOUPDATECP (see http://msdn.microsoft.com/en-us/library/dd162498.aspx)
    If curDC <> 0 Then SetTextAlign curDC, TA_LEFT Or TA_TOP Or TA_NOUPDATECP
    
End Sub

'Return the width (and below, height) of a string, in pixels.  This action REQUIRES that a DC be currently set for this
' object, as the DC's font settings are used to determine the string's dimensions.
Public Function getWidthOfString(ByVal srcString As String) As Long
    Dim txtSize As POINTAPI
    If curDC = 0 Then
        Debug.Print "Can't return string width, as pdFont has not been selected into a DC!"
    Else
        GetTextExtentPoint32 curDC, StrPtr(srcString), Len(srcString), txtSize
        getWidthOfString = txtSize.x
    End If
End Function

Public Function getHeightOfString(ByVal srcString As String) As Long
    Dim txtSize As POINTAPI
    If curDC = 0 Then
        Debug.Print "Can't return string height, as pdFont has not been selected into a DC!"
    Else
        GetTextExtentPoint32 curDC, StrPtr(srcString), Len(srcString), txtSize
        getHeightOfString = txtSize.y
    End If
End Function

'With the current font settings, find the height of a word-wrapped string.  This action REQUIRES that a DC be currently
' set for this object, as the DC's font settings are used to calculate the word-wrapped string's dimensions.
Public Function getHeightOfWordwrapString(ByVal srcString As String, ByVal strWidth As Long) As Long

    Dim tmpRect As RECT
    tmpRect.Left = 0
    tmpRect.Top = 0
    tmpRect.Bottom = 0
    tmpRect.Right = strWidth - 1
    
    DrawText curDC, StrPtr(srcString), Len(srcString), tmpRect, DT_LEFT Or DT_CALCRECT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_NOPREFIX Or DT_TOP
    
    getHeightOfWordwrapString = tmpRect.Bottom

End Function

Private Sub Class_Initialize()
    
    'Determine the proper font point-to-em conversion based on the current screen DPI
    Dim TmpDC As Long
    TmpDC = GetDC(0)
    curLogPixelsY = GetDeviceCaps(TmpDC, LOGPIXELSY)
    ReleaseDC 0, TmpDC
    
    'When the class is initialized, set up a default PhotoDemon font.  Face is automatically set per
    ' the current OS, so external functions do not need to set this manually.
    With currentLogFont
    
        .lfHeight = convertFontSize(11)
        .lfWidth = 0        'This will automatically set the width based on the height
        
        .lfEscapement = 0   'Used for rotation; not utilized at present
        .lfOrientation = 0  'An additional rotation parameter; not utilized at present
        
        .lfWeight = FW_REGULAR
        
        .lfItalic = False
        .lfUnderline = False
        .lfStrikeOut = False
        
        .lfCharSet = DEFAULT_CHARSET
        
        .lfOutPrecision = 0
        .lfClipPrecision = 0
        .lfQuality = CLEARTYPE_QUALITY
        
        'Ideally, we should allow the caller to specify pitch and family requests.  I may add this in the future.
        ' For PD's purposes, however, we want to fallback to a variable width, non-serifed font if (for some reason)
        ' Segoe and Tahoma aren't available on the target machine.
        .lfPitchAndFamily = VARIABLE_PITCH Or FF_SWISS
        
        If g_UserModeFix Then
            m_origFontFace = g_InterfaceFont
        Else
            m_origFontFace = "Tahoma"
        End If
        .lfFaceName = m_origFontFace & Chr(0)
    End With
    
    curFont = 0
    curDC = 0
        
    m_origFontBold = False
    m_origFontItalic = False
    m_origFontSize = 11
    
    fontColor = 0
    curTextAlignment = TA_LEFT
    
End Sub

Private Sub Class_Terminate()

    'Free any remaining font resources
    deleteCurrentFont
    
End Sub

Public Function hasFontBeenCreated() As Boolean
    hasFontBeenCreated = (curFont <> 0)
End Function

Public Function getAttachedDC() As Long
    getAttachedDC = curDC
End Function

Public Function getFontHandle() As Long
    getFontHandle = curFont
End Function

Private Sub deleteCurrentFont()
    
    'If a font was created, delete it now
    If curFont <> 0 Then
        If curDC <> 0 Then SelectObject curDC, oldFont
        
        #If DEBUGMODE = 1 Then
            If DeleteObject(curFont) = 0 Then
                Debug.Print "WARNING! Could not delete font resource - is it selected into a DC??"
            Else
                g_FontsDestroyed = g_FontsDestroyed + 1
            End If
        #Else
            DeleteObject curFont
        #End If
    End If
    
End Sub

'Quick and dirty function for checking font existence.  Does not use API at present, so fonts with Unicode names won't return correct results.
' I haven't done a check, but I expect this to be faster than enum-ing all available fonts and looking for a match...
Public Function doesFontExist(ByVal fontNameToCheck As String, Optional ByRef fontSubstitute As String, Optional ByRef nearestFontSize As Long = 8) As Boolean

    'I don't know if this function will throw errors, but better safe than sorry...
    On Error Resume Next
    
    'Create a temporary font object to test name support
    Dim tmpFont As StdFont
    Set tmpFont = New StdFont
    
    'Copy name and size (if any)
    tmpFont.Name = fontNameToCheck
    tmpFont.Size = nearestFontSize
    
    'See if the font object has the same name.  Windows will automatically substitute a new font if the requested one can't be found.
    If StrComp(fontNameToCheck, tmpFont.Name, vbTextCompare) = 0 Then doesFontExist = True Else doesFontExist = False
    
    'Not sure if helpful, but return the replacement font name and nearest supported size
    fontSubstitute = tmpFont.Name
    nearestFontSize = tmpFont.Size
    
End Function
