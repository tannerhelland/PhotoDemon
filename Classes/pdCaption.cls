VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Caption" support class for user controls
'Copyright 2014-2026 by Tanner Helland
'Created: 27/August/15
'Last updated: 05/June/22
'Last update: remove check for spaces prior to enabling wordwrap; CJK languages (and likely others) can break
'             in plenty of non-space locations!  Thank you to @shishi2019 for catching and reporting this bug.
'
'As part of streamlining PD's many custom user controls, I've been adding "caption" support to each relevant UC.
' This spares us from having to use two controls for each instance (the UC itself, plus a label), simplifying layout
' code while also cutting down on processing time for stuff like theming loops, which need to iterate through all
' control instances on a form.
'
'Rather than stick a bunch of redundant translation and rendering code inside each UC, I've tried to wrap as much as
' possible within this class.  This class is simple; it basically just manages translations for a caption, and returns
' size and rendering info as relevant.  If a control already manages its own backbuffer (which most do), rendering the
' text requires only a single line of code!
'
'The initial implementation of this class was the pdSlider UC, so look there for detailed comments on how to
' make use of the class.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The user-specified size of the caption font.  We manage this for the UC, and return it as the specified "font size" if queried.
' All "resize-to-fit" font decisions are managed internally.
Private m_FontSizeCaption As Single, m_FontSizeCaptionFit As Single

'Current caption string (persistent within the IDE, but must be set at run-time for Unicode languages).  Note that m_CaptionEn
' is the ENGLISH CAPTION ONLY.  A translated caption will be stored in m_CaptionTranslated; the translated copy will be updated
' by any caption change, or by a call to UpdateAgainstCurrentTheme.
Private m_CaptionEn As String, m_CaptionTranslated As String

'Width and height of our parent container.  This is used to calculate automatic font sizing.
Private m_ContainerWidth As Single, m_ContainerHeight As Single

'Some controls may allow wordwrap when resizing to fit a target rect
Private m_WordWrapAllowed As Boolean

'This class will auto-fit the current caption to the available container area.  If we've checked the current caption + fontsize
' combination against a known container width, this will be set to TRUE.
Private m_FontSizeVerified As Boolean

'Other font settings.  I implement these as I need them, so some may be missing, by design.
Private m_FontBold As Boolean, m_FontItalic As Boolean, m_FontUnderline As Boolean

'Caption alignment defaults to left-alignment, but to someday support RTL languages properly, we may need to toggle this
' value at run-time, as part of the translation engine.  Note that this setting is not binding, because some controls
' will want to manually call the DrawCaptionCentered function (e.g. button strips, which do a bunch of custom text rendering).
Private m_Alignment As AlignmentConstants

'Normally, the caption's color is pulled from PD's visual themes engine.  However, parent controls can override the
' color if they must.
Private m_CaptionColor As Long, m_CaptionColorCustom As Long

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDCAPTION_COLOR_LIST
    [_First] = 0
    PDC_Caption = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'Get/set alignment.  Please note the caveats under m_Alignment's definition, above.
Friend Function GetAlignment() As AlignmentConstants
    GetAlignment = m_Alignment
End Function

Friend Function SetAlignment(ByVal newAlignment As AlignmentConstants) As Boolean
    If (newAlignment <> m_Alignment) Then
        m_Alignment = newAlignment
        SetAlignment = True
    End If
End Function

'Get the current ENGLISH caption.
Friend Function GetCaptionEn() As String
    GetCaptionEn = m_CaptionEn
End Function

'Get the current TRANSLATED caption.
Friend Function GetCaptionTranslated() As String
    GetCaptionTranslated = m_CaptionTranslated
End Function

'Set a new caption.  Note that this is a *function*, not a *sub*, and it will return TRUE if the property change requires
' a redraw.  (Previously, I would raise an event when a redraw was required, but that burdens the caller even more because
' a lot of redraw events are raised during initialization, which we don't really want.)  Anyway, the take-home message is that
' the caller needs to check the return of this function, and handle redraws accordingly.
Friend Function SetCaption(ByRef newCaption As String) As Boolean

    If Strings.StringsNotEqual(newCaption, m_CaptionEn, False) Then
        
        'Whenever the caption changes, we need to double-check that the caption + fontsize combination fits inside the control.
        m_FontSizeVerified = False
        m_CaptionEn = newCaption
        
        'During run-time, apply translations as necessary
        If PDMain.IsProgramRunning() Then
        
            'See if translations are necessary.
            Dim isTranslationActive As Boolean
                
            If (Not g_Language Is Nothing) Then
                isTranslationActive = g_Language.TranslationActive
            Else
                isTranslationActive = False
            End If
            
            'Update the translated caption accordingly
            If isTranslationActive Then
                m_CaptionTranslated = g_Language.TranslateMessage(m_CaptionEn)
            Else
                m_CaptionTranslated = m_CaptionEn
            End If
        
        Else
            m_CaptionTranslated = m_CaptionEn
        End If
        
        SetCaption = True
            
    Else
        SetCaption = False
    End If
    
End Function

'Simplified wrapper for checking if a caption is active
Friend Function IsCaptionActive() As Boolean
    IsCaptionActive = (LenB(m_CaptionEn) <> 0) And (m_FontSizeCaption <> 0)
End Function

'Assign a new caption color.  Remember that color changes do not require rebuilding the font.  Also, font color changes are persistent
' (by design), so they will not be overridden except by subsequent setCaptionColor calls.
Friend Sub SetCaptionColor(ByVal newCaptionColor As Long)
    m_CaptionColorCustom = newCaptionColor
End Sub

Friend Function GetFontBold() As Boolean
    GetFontBold = m_FontBold
End Function

Friend Function SetFontBold(ByVal newValue As Boolean) As Boolean
    If (m_FontBold <> newValue) Then
        m_FontBold = newValue
        m_FontSizeVerified = False
        SetFontBold = True
    Else
        SetFontBold = False
    End If
End Function

Friend Function GetFontItalic() As Boolean
    GetFontItalic = m_FontItalic
End Function

Friend Function SetFontItalic(ByVal newValue As Boolean) As Boolean
    If (m_FontItalic <> newValue) Then
        m_FontItalic = newValue
        m_FontSizeVerified = False
        SetFontItalic = True
    Else
        SetFontItalic = False
    End If
End Function

Friend Function GetFontUnderline() As Boolean
    GetFontUnderline = m_FontUnderline
End Function

Friend Function SetFontUnderline(ByVal newValue As Boolean) As Boolean
    If (m_FontUnderline <> newValue) Then
        m_FontUnderline = newValue
        m_FontSizeVerified = False
        SetFontUnderline = True
    Else
        SetFontUnderline = False
    End If
End Function

Friend Function GetFontSize() As Single
    GetFontSize = m_FontSizeCaption
End Function

'Set a new font size.  Note that this is a *function*, not a *sub*, and it will return TRUE if the property change requires
' a redraw.  (Previously, I would raise an event when a redraw was required, but that burdens the caller even more because
' a lot of redraw events are raised during initialization, which we don't really want.)  Anyway, the take-home message is that
' the caller needs to check the return of this function, and handle redraws accordingly.
Friend Function SetFontSize(ByVal newSize As Single) As Boolean
    
    If (newSize <> m_FontSizeCaption) Then
        m_FontSizeCaption = newSize
        m_FontSizeCaptionFit = m_FontSizeCaption
        m_FontSizeVerified = False
        SetFontSize = True
    Else
        SetFontSize = False
    End If
    
End Function

'Wordwrap can now be optionally enabled on-the-fly
Friend Sub SetWordWrapSupport(ByVal newSupportValue As Boolean)
    If (m_WordWrapAllowed <> newSupportValue) Then
        m_WordWrapAllowed = newSupportValue
        m_FontSizeVerified = False
    End If
End Sub

Private Sub Class_Initialize()
    
    'By default, we assume our parent control does not have an active caption.
    m_CaptionEn = 0
    m_CaptionTranslated = 0
    m_FontSizeCaption = 0
    m_FontSizeCaptionFit = m_FontSizeCaption
    
    m_ContainerWidth = 0
    m_ContainerHeight = 0
    m_WordWrapAllowed = False
    m_Alignment = vbLeftJustify
    
    m_FontSizeVerified = False
    m_FontBold = False
    m_FontItalic = False
    m_FontUnderline = False
    
    m_CaptionColor = -1
    m_CaptionColorCustom = -1
    
End Sub

'For automatic font-fitting to work, we have to know our parent container's width.  The parent needs to supply this value
' 1) prior to requesting any draw or measurement actions, and 2) whenever the control dimensions change.
Friend Sub SetControlSize(ByVal newWidth As Single, ByVal newHeight As Single)
    
    If (m_ContainerWidth <> newWidth) Or (m_ContainerHeight <> newHeight) Then
        m_ContainerWidth = newWidth
        m_ContainerHeight = newHeight
        m_FontSizeVerified = False
    End If
    
End Sub

'Returns the width of the current caption, in pixels, with autofit active.  No padding is applied.
' Note that this return *ignores the wordwrap setting.*  Wordwrap (obviously) affects the width of the string,
' but because the width of a wordwrap string is simply the container width - which the caller must specify -
' there's no point in returning it here.
Friend Function GetCaptionWidth() As Single
    
    'Make sure we can fit the current caption inside the control, when using the selected font size.  (If we can't,
    ' this function will automatically calculation a new font size, and stick it inside m_FontSizeCaptionFit.)
    If (Not m_FontSizeVerified) Then VerifyFontSize
    
    'Get a copy of the current font from the central font cache
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(m_FontSizeCaptionFit, m_FontBold, m_FontItalic, m_FontUnderline)
    
    GetCaptionWidth = tmpFont.GetWidthOfString(m_CaptionTranslated)
    
End Function

'Returns the width of the current caption, in pixels, *without considering autofit*.  No padding is applied.
Friend Function GetCaptionWidth_NoFit() As Single
    
    'Get a copy of the default font size from the central font cache
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(m_FontSizeCaption, m_FontBold, m_FontItalic, m_FontUnderline)
    
    GetCaptionWidth_NoFit = tmpFont.GetWidthOfString(m_CaptionTranslated)
    
End Function

'Returns the height of the current caption, in pixels.  No padding is applied.
' Note that this return varies depending on the presence of wordwrapping.  Wordwrap will (obviously) affect the return.
Friend Function GetCaptionHeight() As Single
    
    'Make sure we can fit the current caption inside the control, when using the selected font size.  (If we can't,
    ' this function will automatically calculation a new font size, and stick it inside m_FontSizeCaptionFit.)
    If (Not m_FontSizeVerified) Then VerifyFontSize
    
    'Get a copy of the current font from the central font cache
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(m_FontSizeCaptionFit, m_FontBold, m_FontItalic, m_FontUnderline)
    
    If m_WordWrapAllowed Then
        GetCaptionHeight = tmpFont.GetHeightOfWordwrapString(m_CaptionTranslated, m_ContainerWidth)
    Else
        GetCaptionHeight = tmpFont.GetHeightOfString(m_CaptionTranslated)
    End If
    
End Function

'Returns the height of the current caption, in pixels, *without considering autofit*.  No padding is applied.
' Note that this return varies depending on the presence of wordwrapping.  Wordwrap will (obviously) affect the return.
Friend Function GetCaptionHeight_NoFit() As Single
    
    'Get a copy of the current font from the central font cache
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(m_FontSizeCaption, m_FontBold, m_FontItalic, m_FontUnderline)
    
    If m_WordWrapAllowed Then
        GetCaptionHeight_NoFit = tmpFont.GetHeightOfWordwrapString(m_CaptionTranslated, m_ContainerWidth)
    Else
        GetCaptionHeight_NoFit = tmpFont.GetHeightOfString(m_CaptionTranslated)
    End If
    
End Function

'When all font and caption settings are ready, use this to actually render the caption onto a target DC.
' IMPORTANT NOTE: this command exists only to support controls that do their own custom text positioning.  It does not support
'                  text alignment (how could it, without a constraining rect?) so you must use DrawCaption_Clipped if you want
'                  alignment support.
Friend Sub DrawCaption(ByVal dstDC As Long, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal customColor As Long = -1)
    
    'Make sure we can fit the current caption inside the control, when using the selected font size.  (If we can't,
    ' this function will automatically calculation a new font size, and stick it inside m_FontSizeCaptionFit.)
    If (Not m_FontSizeVerified) Then VerifyFontSize
    
    'Failsafe check
    If (dstDC <> 0) Then
        
        'Get a copy of the current font from the central font cache
        Dim tmpFont As pdFont
        Set tmpFont = Fonts.GetMatchingUIFont(m_FontSizeCaptionFit, m_FontBold, m_FontItalic, m_FontUnderline)
        
        'Make sure colors have correctly retrieved from the active theme
        If (customColor <> -1) Then
            tmpFont.SetFontColor customColor
        Else
            If (m_Colors Is Nothing) Then InitializeColorCache
            If (m_CaptionColor = -1) Then m_CaptionColor = m_Colors.RetrieveColor(PDC_Caption)
            If (m_CaptionColorCustom <> -1) Then tmpFont.SetFontColor m_CaptionColorCustom Else tmpFont.SetFontColor m_CaptionColor
        End If
        
        tmpFont.AttachToDC dstDC
        
        If m_WordWrapAllowed Then
            tmpFont.FastRenderMultilineText dstX, dstY, m_CaptionTranslated
        Else
            tmpFont.FastRenderText dstX, dstY, m_CaptionTranslated
        End If
        
        tmpFont.ReleaseFromDC
        
    End If
    
End Sub

'When all font and caption settings are ready, use this to actually render the caption onto a target DC.
' Clipping to the destination area will occur automatically, with ellipses applied according to the
' useEllipses parameter.  Note that you can also request to use the original font size, instead of the
' auto-fit size.
Friend Sub DrawCaption_Clipped(ByVal dstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, Optional ByVal customColor As Long = -1, Optional ByVal useEllipses As Boolean = False, Optional ByVal useOriginalFontSize As Boolean = False, Optional ByVal centerVertically As Boolean = False)
    
    'Make sure we can fit the current caption inside the control, when using the selected font size.  (If we can't,
    ' this function will automatically calculation a new font size, and stick it inside m_FontSizeCaptionFit.)
    If (Not m_FontSizeVerified) Then VerifyFontSize
    
    'Failsafe check
    If (dstDC <> 0) Then
        
        'Get a copy of the current font from the central font cache
        Dim tmpFont As pdFont
        
        If useOriginalFontSize Then
            Set tmpFont = Fonts.GetMatchingUIFont(m_FontSizeCaption, m_FontBold, m_FontItalic, m_FontUnderline)
        Else
            Set tmpFont = Fonts.GetMatchingUIFont(m_FontSizeCaptionFit, m_FontBold, m_FontItalic, m_FontUnderline)
        End If
        
        If (customColor <> -1) Then
            tmpFont.SetFontColor customColor
        Else
            If (m_Colors Is Nothing) Then InitializeColorCache
            If (m_CaptionColor = -1) Then m_CaptionColor = m_Colors.RetrieveColor(PDC_Caption)
            If (m_CaptionColorCustom <> -1) Then tmpFont.SetFontColor m_CaptionColorCustom Else tmpFont.SetFontColor m_CaptionColor
        End If
        
        tmpFont.AttachToDC dstDC
        
        'NOTE: pdFont.SetTextAlignment does not actually set text alignment as a property inside the target DC.
        '      Instead, it uses DrawText flags to enforce alignment.  This leaves the target DC untouched,
        '      so previously set text flags won't be modified by calls to pdFont, by design.
        tmpFont.SetTextAlignment m_Alignment
        
        If m_WordWrapAllowed Then
            tmpFont.FastRenderMultilineTextWithClipping dstX, dstY, dstWidth, dstHeight, m_CaptionTranslated, centerVertically
        Else
            tmpFont.FastRenderTextWithClipping dstX, dstY, dstWidth, dstHeight, m_CaptionTranslated, useEllipses, centerVertically
        End If
        
        tmpFont.ReleaseFromDC
        
    End If
    
End Sub

'Same as drawCaption, but centered on an arbitrary rect.  This helper function primarily exists for controls that must do
' a bunch of non-standard text rendering (e.g. button strips, which individually fit each button caption).
Friend Sub DrawCaptionCentered(ByVal dstDC As Long, ByRef dstRect As RECT, Optional ByVal customColor As Long = -1)
    
    'Make sure we can fit the current caption inside the control, when using the selected font size.  (If we can't,
    ' this function will automatically calculation a new font size, and stick it inside m_FontSizeCaptionFit.)
    If (Not m_FontSizeVerified) Then VerifyFontSize
    
    'Failsafe check
    If (dstDC <> 0) Then
    
        'Get a copy of the current font from the central font cache
        Dim tmpFont As pdFont
        Set tmpFont = Fonts.GetMatchingUIFont(m_FontSizeCaptionFit, m_FontBold, m_FontItalic, m_FontUnderline)
        
        If (customColor <> -1) Then
            tmpFont.SetFontColor customColor
        Else
            If (m_Colors Is Nothing) Then InitializeColorCache
            If (m_CaptionColor = -1) Then m_CaptionColor = m_Colors.RetrieveColor(PDC_Caption)
            If (m_CaptionColorCustom <> -1) Then tmpFont.SetFontColor m_CaptionColorCustom Else tmpFont.SetFontColor m_CaptionColor
        End If
        
        tmpFont.AttachToDC dstDC
        tmpFont.SetTextAlignment vbLeftJustify
        
        If m_WordWrapAllowed Then
            
            'See if the string, as-is, fits within the target area
            Dim strWidth As Long
            strWidth = tmpFont.GetWidthOfString(m_CaptionTranslated)
            
            If (strWidth < (dstRect.Right - dstRect.Left)) Then
                tmpFont.DrawCenteredTextToRect m_CaptionTranslated, dstRect, True
            Else
            
                'Find the height of the string as it currently exists
                Dim strHeight As Long
                strHeight = tmpFont.GetHeightOfWordwrapString(m_CaptionTranslated, dstRect.Right - dstRect.Left)
                
                'Recalculate the rect top so that the text is centered
                Dim newRect As RECT
                newRect.Left = dstRect.Left
                newRect.Bottom = dstRect.Bottom
                newRect.Right = dstRect.Right
                newRect.Top = dstRect.Top + (CSng(CSng(dstRect.Bottom - dstRect.Top) / 2) - CSng(strHeight / 2))
                
                Const DT_CENTER = &H1, DT_WORDBREAK = &H10, DT_NOCLIP = &H100
                tmpFont.DrawTextWrapper StrPtr(m_CaptionTranslated), Len(m_CaptionTranslated), newRect, DT_WORDBREAK Or DT_NOCLIP Or DT_CENTER
                
            End If
            
        Else
            tmpFont.DrawCenteredTextToRect m_CaptionTranslated, dstRect, True
        End If
        
        tmpFont.ReleaseFromDC
        
    End If
    
End Sub

'Make sure the current caption, drawn with the current font, fits the available rendering area.  (Note that this function
' is pointless if we haven't been notified of a container size yet.)
Private Sub VerifyFontSize()
    
    'Our parent control is responsible for notifying us if its size changes.
    ' If it hasn't notified us yet, verification can't proceed.
    If (m_ContainerWidth > 0) Then
        
        'Shortcut if the caption is null (many toolbar buttons don't use captions)
        If (LenB(m_CaptionTranslated) = 0) Then
            m_FontSizeCaptionFit = m_FontSizeCaption
            m_FontSizeVerified = True
            Exit Sub
        End If
        
        Dim newFontSize As Single
        
        'There are two ways to verify font size, and they differ depending on whether the caller
        ' wants us to enable wordwrapping.
        If m_WordWrapAllowed Then
            newFontSize = Fonts.FindFontSizeWordWrap(m_CaptionTranslated, m_ContainerWidth, m_ContainerHeight, m_FontSizeCaption, m_FontBold, m_FontItalic, m_FontUnderline, True)
        Else
            newFontSize = Fonts.FindFontSizeSingleLine(m_CaptionTranslated, m_ContainerWidth, m_FontSizeCaption, m_FontBold, m_FontItalic, m_FontUnderline, True)
        End If
        
        'Note the size required to fit the text inside the target area.
        ' (On subsequent draw calls, we will silently switch to *this* font size.)
        m_FontSizeCaptionFit = newFontSize
        
        'Once we've verified the control's font size, we don't need to verify it again, unless...
        ' 1) the caption changes
        ' 2) the parent control size changes
        ' 3) our font settings change
        ' Those actions all reset m_FontSizeVerified to FALSE.
        m_FontSizeVerified = True
        
    Else
        m_FontSizeVerified = False
    End If

End Sub

'Normally, we initialize the color class at class load, but this class makes that tricky as it may be initialized
' before PD's theming engine (as FormMain contains controls that use this class).  As such, we manually construct
' the color class the first time we receive a draw request, which prevents us from even touching the theme engine
' if we're initialized but never used for drawing.
Private Sub InitializeColorCache()
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCAPTION_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "GenericPDControl", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    
    If (m_Colors Is Nothing) Then InitializeColorCache
    
    'Color list retrieval is pretty darn easy - just load each color one at a time, and leave the rest to the color class.
    ' It will build an internal hash table of the colors we request, which makes rendering much faster.
    m_Colors.LoadThemeColor PDC_Caption, "Caption", IDE_GRAY
    
End Sub

'If the theme or language is changed at run-time, the caller should call this function.  It will refresh all tooltips in
' its collection against the current language and theme, without requiring any special input or parameters.
Friend Sub UpdateAgainstCurrentTheme(Optional ByRef ctlName As String = vbNullString)
    
    Dim oldTranslatedText As String
    oldTranslatedText = m_CaptionTranslated
    
    'See if a new language has been selected
    If PDMain.IsProgramRunning() Then
    
        Dim isTranslationActive As Boolean
        If (Not g_Language Is Nothing) Then
            isTranslationActive = g_Language.TranslationActive
        Else
            isTranslationActive = False
        End If
        
        'Update the translated caption accordingly
        If isTranslationActive Then
            If (LenB(ctlName) > 0) Then
                m_CaptionTranslated = g_Language.TranslateMessage(m_CaptionEn, SPECIAL_TRANSLATION_OBJECT_PREFIX & ctlName)
            Else
                m_CaptionTranslated = g_Language.TranslateMessage(m_CaptionEn)
            End If
        Else
            m_CaptionTranslated = m_CaptionEn
        End If
    
    Else
        m_CaptionTranslated = m_CaptionEn
    End If
    
    'If our translated caption has changed, the new text may overflow its container, so we need to run size heuristics.
    m_FontSizeVerified = Strings.StringsEqual(oldTranslatedText, m_CaptionTranslated, False)
    
    'Theme changes may also affect caption coloring; update the relevant color now
    m_CaptionColor = -1
    UpdateColorList
    
End Sub
