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
'Copyright 2014-2015 by Tanner Helland
'Created: 27/August/15
'Last updated: 28/October/15
'Last update: rewrite the class to use a program-level cache of requested caption sizes.  This spares us from needing to
'              create tons of redundant font objects for rendering UI text.
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
'The initial implementation of this class was the sliderTextCombo UC, so look there for detailed comments on how to
' make use of the class.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user-specified size of the caption font.  We manage this for the UC, and return it as the specified "font size" if queried.
' All "resize-to-fit" font decisions are managed internally.
Private m_FontSizeCaption As Single, m_FontSizeCaptionFit As Single

'Current caption string (persistent within the IDE, but must be set at run-time for Unicode languages).  Note that m_CaptionEn
' is the ENGLISH CAPTION ONLY.  A translated caption will be stored in m_CaptionTranslated; the translated copy will be updated
' by any caption change, or by a call to UpdateAgainstCurrentTheme.
Private m_CaptionEn As String
Private m_CaptionTranslated As String

'Width and height of our parent container.  This is used to calculate automatic font sizing.
Private m_ContainerWidth As Single, m_ContainerHeight As Single

'Some controls may allow wordwrap when resizing to fit a target rect
Private m_WordWrapAllowed As Boolean

'This class will auto-fit the current caption to the available container area.  If we've checked the current caption + fontsize
' combination against a known container width, this will be set to TRUE.
Private m_FontSizeVerified As Boolean

'Current caption color
Private m_CaptionColor As Long

'Other font settings.  I implement these as I need them, so some may be missing, by design.
Private m_FontBold As Boolean

'Get the current ENGLISH caption.
Public Function GetCaptionEn() As String
    GetCaptionEn = m_CaptionEn
End Function

'Get the current TRANSLATED caption.
Public Function GetCaptionTranslated() As String
    GetCaptionTranslated = m_CaptionTranslated
End Function

'Set a new caption.  Note that this is a *function*, not a *sub*, and it will return TRUE if the property change requires
' a redraw.  (Previously, I would raise an event when a redraw was required, but that burdens the caller even more because
' a lot of redraw events are raised during initialization, which we don't really want.)  Anyway, the take-home message is that
' the caller needs to check the return of this function, and handle redraws accordingly.
Public Function SetCaption(ByVal newCaption As String) As Boolean

    If StrComp(newCaption, m_CaptionEn, vbBinaryCompare) <> 0 Then
        
        'Whenever the caption changes, we need to double-check that the caption + fontsize combination fits inside the control.
        m_FontSizeVerified = False
                
        m_CaptionEn = newCaption
        
        'During run-time, apply translations as necessary
        If g_IsProgramRunning Then
        
            'See if translations are necessary.
            Dim isTranslationActive As Boolean
                
            If Not (g_Language Is Nothing) Then
                If g_Language.translationActive Then
                    isTranslationActive = True
                Else
                    isTranslationActive = False
                End If
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
Public Function IsCaptionActive() As Boolean
    
    If (Len(m_CaptionEn) <> 0) And (m_FontSizeCaption <> 0) Then
        IsCaptionActive = True
    Else
        IsCaptionActive = False
    End If
    
End Function

'Assign a new caption color.  Remember that color changes do not require rebuilding the font.  Also, font color changes are persistent
' (by design), so they will not be overridden except by subsequent setCaptionColor calls.
Public Sub SetCaptionColor(ByVal newColor As Long)
    m_CaptionColor = newColor
End Sub

Public Function GetFontBold() As Boolean
    GetFontBold = m_FontBold
End Function

Public Function SetFontBold(ByVal newValue As Boolean) As Boolean
    If m_FontBold <> newValue Then
        m_FontBold = newValue
        m_FontSizeVerified = False
        SetFontBold = True
    Else
        SetFontBold = False
    End If
End Function

Public Function GetFontSize() As Single
    GetFontSize = m_FontSizeCaption
End Function

'Set a new font size.  Note that this is a *function*, not a *sub*, and it will return TRUE if the property change requires
' a redraw.  (Previously, I would raise an event when a redraw was required, but that burdens the caller even more because
' a lot of redraw events are raised during initialization, which we don't really want.)  Anyway, the take-home message is that
' the caller needs to check the return of this function, and handle redraws accordingly.
Public Function SetFontSize(ByVal newSize As Single) As Boolean
    
    If newSize <> m_FontSizeCaption Then
        m_FontSizeCaption = newSize
        m_FontSizeCaptionFit = m_FontSizeCaption
        m_FontSizeVerified = False
        SetFontSize = True
    Else
        SetFontSize = False
    End If
    
End Function

'Whenever relevant, the owner of this class can call this function to see if a redraw is needed.  The redraw flags will be
' reset whenever a successful draw request occurs.
Public Function IsRedrawNeeded() As Boolean
    If (Not m_FontSizeVerified) Then
        IsRedrawNeeded = True
    Else
        IsRedrawNeeded = False
    End If
End Function

'Wordwrap can now be optionally enabled on-the-fly
Public Sub SetWordWrapSupport(ByVal newSupportValue As Boolean)
    If m_WordWrapAllowed <> newSupportValue Then
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
    
    m_FontSizeVerified = False
    m_FontBold = False
    
    If g_IsProgramRunning And Not (g_Themer Is Nothing) Then
        m_CaptionColor = g_Themer.GetThemeColor(PDTC_TEXT_DEFAULT)
    Else
        m_CaptionColor = RGB(96, 96, 96)
    End If
    
End Sub

'For automatic font-fitting to work, we have to know our parent container's width.  The parent needs to supply this value
' 1) prior to requesting any draw or measurement actions, and 2) whenever the control dimensions change.
Public Sub SetControlSize(ByVal newWidth As Single, ByVal newHeight As Single)
    
    If (m_ContainerWidth <> newWidth) Or (m_ContainerHeight <> newHeight) Then
        m_ContainerWidth = newWidth
        m_ContainerHeight = newHeight
        m_FontSizeVerified = False
    End If
    
End Sub

'Returns the width of the current caption, in pixels.  No padding is applied.
' Note that this return *ignores the wordwrap setting.*  Wordwrap will (obviously) affect the width of the string, but because the width
' of a wordwrap string is simply the container width - which the caller must specify - there's no point in returning it here.
Public Function GetCaptionWidth() As Single
    
    'Make sure we can fit the current caption inside the control, when using the selected font size.  (If we can't,
    ' this function will automatically calculation a new font size, and stick it inside m_FontSizeCaptionFit.)
    If Not m_FontSizeVerified Then VerifyFontSize
    
    'Get a copy of the current font from the central font cache
    Dim tmpFont As pdFont
    Set tmpFont = Font_Management.GetMatchingUIFont(m_FontSizeCaptionFit, m_FontBold, False, False)
    
    GetCaptionWidth = tmpFont.GetWidthOfString(m_CaptionTranslated)
    
End Function

'Returns the height of the current caption, in pixels.  No padding is applied.
' Note that this return varies depending on the presence of wordwrapping.  Wordwrap will (obviously) affect the return.
Public Function GetCaptionHeight() As Single
    
    'Make sure we can fit the current caption inside the control, when using the selected font size.  (If we can't,
    ' this function will automatically calculation a new font size, and stick it inside m_FontSizeCaptionFit.)
    If Not m_FontSizeVerified Then VerifyFontSize
    
    'Get a copy of the current font from the central font cache
    Dim tmpFont As pdFont
    Set tmpFont = Font_Management.GetMatchingUIFont(m_FontSizeCaptionFit, m_FontBold, False, False)
    
    If m_WordWrapAllowed Then
        GetCaptionHeight = tmpFont.GetHeightOfWordwrapString(m_CaptionTranslated, m_ContainerWidth)
    Else
        GetCaptionHeight = tmpFont.GetHeightOfString(m_CaptionTranslated)
    End If
    
End Function

'When all font and caption settings are ready, use this to actually render the caption onto a target DC
Public Sub DrawCaption(ByVal dstDC As Long, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal customColor As Long = -1)
    
    'Make sure we can fit the current caption inside the control, when using the selected font size.  (If we can't,
    ' this function will automatically calculation a new font size, and stick it inside m_FontSizeCaptionFit.)
    If Not m_FontSizeVerified Then VerifyFontSize
    
    'Get a copy of the current font from the central font cache
    Dim tmpFont As pdFont
    Set tmpFont = Font_Management.GetMatchingUIFont(m_FontSizeCaptionFit, m_FontBold, False, False)
    
    If customColor <> -1 Then
        tmpFont.SetFontColor customColor
    Else
        tmpFont.SetFontColor m_CaptionColor
    End If
    
    tmpFont.AttachToDC dstDC
    
    If m_WordWrapAllowed Then
        tmpFont.FastRenderMultilineText dstX, dstY, m_CaptionTranslated
    Else
        tmpFont.FastRenderText dstX, dstY, m_CaptionTranslated
    End If
    
    tmpFont.ReleaseFromDC
    
End Sub

'Same as drawCaption, but centered on an arbitrary rect
Friend Sub DrawCaptionCentered(ByVal dstDC As Long, ByRef dstRect As RECT, Optional ByVal customColor As Long = -1)
    
    'Make sure we can fit the current caption inside the control, when using the selected font size.  (If we can't,
    ' this function will automatically calculation a new font size, and stick it inside m_FontSizeCaptionFit.)
    If Not m_FontSizeVerified Then VerifyFontSize
    
    'Get a copy of the current font from the central font cache
    Dim tmpFont As pdFont
    Set tmpFont = Font_Management.GetMatchingUIFont(m_FontSizeCaptionFit, m_FontBold, False, False)
    
    If customColor <> -1 Then
        tmpFont.SetFontColor customColor
    Else
        tmpFont.SetFontColor m_CaptionColor
    End If
    
    tmpFont.AttachToDC dstDC
    
    If m_WordWrapAllowed Then
        
        'See if the string, as-is, fits within the target area
        Dim strWidth As Long
        strWidth = tmpFont.GetWidthOfString(m_CaptionTranslated)
        
        If strWidth < (dstRect.Right - dstRect.Left) Then
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
    Set tmpFont = Nothing
    
End Sub

'Make sure the current caption, drawn with the current font, fits the available rendering area.  (Note that this function
' is pointless if we haven't been notified of a container size yet.)
Private Sub VerifyFontSize()
    
    'Our parent control is responsible for notifying us if its size changes.  If it hasn't notified us yet, verification can't proceed.
    If m_ContainerWidth > 0 Then
        
        Dim newFontSize As Single
        
        'There are two ways to verify font size, and they differ depending on whether the caller wants us to enable wordwrapping.
        If m_WordWrapAllowed Then
            newFontSize = Font_Management.FindFontSizeWordWrap(m_CaptionTranslated, m_ContainerWidth, m_ContainerHeight, m_FontSizeCaption, m_FontBold, False, False, True)
        Else
            newFontSize = Font_Management.FindFontSizeSingleLine(m_CaptionTranslated, m_ContainerWidth, m_FontSizeCaption, m_FontBold, False, False, True)
        End If
        
        'Note the size required to fit the text inside the target area.  On subsequent draw calls, we will silently switch
        ' to *this* font size.
        m_FontSizeCaptionFit = newFontSize
        
        'Once we've verified the control's font size, we don't need to verify it again, unless...
        ' 1) the caption changes
        ' 2) the parent control size changes
        ' 3) our font settings change
        ' Those actions all reset m_FontSizeVerified to FALSE.
        m_FontSizeVerified = True
        
    Else
        Debug.Print "WARNING!  pdCaption.verifyFontSize was called, but the container width is unknown!"
        m_FontSizeVerified = False
    End If

End Sub

'If the theme or language is changed at run-time, the caller should call this function.  It will refresh all tooltips in
' its collection against the current language and theme, without requiring any special input or parameters.
Public Sub UpdateAgainstCurrentTheme()
    
    Dim oldTranslatedText As String
    oldTranslatedText = m_CaptionTranslated
    
    'See if a new language has been selected
    If g_IsProgramRunning Then
    
        Dim isTranslationActive As Boolean
            
        If Not (g_Language Is Nothing) Then
            If g_Language.translationActive Then
                isTranslationActive = True
            Else
                isTranslationActive = False
            End If
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
    
    'If our translated caption has changed, the new text may overflow its container, so we need to run size heuristics.
    If StrComp(oldTranslatedText, m_CaptionTranslated, vbBinaryCompare) <> 0 Then m_FontSizeVerified = False
    
    'Theme changes may also affect caption coloring; update the relevant color now
    If g_IsProgramRunning Then
        m_CaptionColor = g_Themer.GetThemeColor(PDTC_TEXT_DEFAULT)
    Else
        m_CaptionColor = RGB(96, 96, 96)
    End If
        
End Sub
