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
'Last updated: 28/August/15
'Last update: migrate individual caption implementations into this singular class
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
Private m_FontSizeCaption As Single

'Current caption string (persistent within the IDE, but must be set at run-time for Unicode languages).  Note that m_CaptionEn
' is the ENGLISH CAPTION ONLY.  A translated caption will be stored in m_CaptionTranslated; the translated copy will be updated
' by any caption change, or by a call to UpdateAgainstCurrentTheme.
Private m_CaptionEn As String
Private m_CaptionTranslated As String

'When a caption is active, we use pdFont for actual text rendering duties.
Private m_TextRenderer As pdFont

'To optimize font handling, we only recreate the font when absolutely necessary.
Private m_FontObjectClean As Boolean

'Width and height of our parent container.  This is used to calculate automatic font sizing.
Private m_ContainerWidth As Single, m_ContainerHeight As Single

'Some controls may allow wordwrap when resizing to fit a target rect
Private m_WordWrapAllowed As Boolean

'This class will auto-fit the current caption to the available container area.  If we've established
Private m_FontSizeVerified As Boolean

'Current caption color
Private m_CaptionColor As Long

'Other font settings.  I implement these as I need them, so some may be missing, by design.
Private m_FontBold As Boolean

'Get the current ENGLISH caption.
Public Function getCaptionEn() As String
    getCaptionEn = m_CaptionEn
End Function

'Get the current TRANSLATED caption.
Public Function getCaptionTranslated() As String
    getCaptionTranslated = m_CaptionTranslated
End Function

'Set a new caption.  Note that this is a *function*, not a *sub*, and it will return TRUE if the property change requires
' a redraw.  (Previously, I would raise an event when a redraw was required, but that burdens the caller even more because
' a lot of redraw events are raised during initialization, which we don't really want.)  Anyway, the take-home message is that
' the caller needs to check the return of this function, and handle redraws accordingly.
Public Function setCaption(ByVal newCaption As String) As Boolean

    If StrComp(newCaption, m_CaptionEn, vbBinaryCompare) <> 0 Then
        
        m_FontObjectClean = False
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
        
        setCaption = True
            
    Else
        setCaption = False
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
Public Sub setCaptionColor(ByVal newColor As Long)
    m_CaptionColor = newColor
End Sub

Public Function getFontBold() As Boolean
    getFontBold = m_FontBold
End Function

Public Function setFontBold(ByVal newValue As Boolean) As Boolean
    If m_FontBold <> newValue Then
        m_FontBold = newValue
        m_FontObjectClean = False
        m_FontSizeVerified = False
        setFontBold = True
    Else
        setFontBold = False
    End If
End Function

Public Function getFontSize() As Single
    getFontSize = m_FontSizeCaption
End Function

'Set a new font size.  Note that this is a *function*, not a *sub*, and it will return TRUE if the property change requires
' a redraw.  (Previously, I would raise an event when a redraw was required, but that burdens the caller even more because
' a lot of redraw events are raised during initialization, which we don't really want.)  Anyway, the take-home message is that
' the caller needs to check the return of this function, and handle redraws accordingly.
Public Function setFontSize(ByVal newSize As Single) As Boolean
    
    If newSize <> m_FontSizeCaption Then
        m_FontSizeCaption = newSize
        m_FontObjectClean = False
        m_FontSizeVerified = False
        setFontSize = True
    Else
        setFontSize = False
    End If
    
End Function

'Whenever relevant, the owner of this class can call this function to see if a redraw is needed.  The redraw flags will be
' reset whenever a successful draw request occurs.
Public Function isRedrawNeeded() As Boolean
    If (Not m_FontObjectClean) Or (Not m_FontSizeVerified) Then
        isRedrawNeeded = True
    Else
        isRedrawNeeded = False
    End If
End Function

'Wordwrap can now be optionally enabled
Public Sub setWordWrapSupport(ByVal newSupportValue As Boolean)
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
    
    m_ContainerWidth = 0
    m_ContainerHeight = 0
    m_WordWrapAllowed = False
    
    m_FontObjectClean = False
    m_FontSizeVerified = False
    m_FontBold = False
    
    If g_IsProgramRunning And Not (g_Themer Is Nothing) Then
        m_CaptionColor = g_Themer.getThemeColor(PDTC_TEXT_DEFAULT)
    Else
        m_CaptionColor = RGB(96, 96, 96)
    End If
    
End Sub

'For automatic font-fitting to work, we have to know our parent container's width.  The parent needs to supply this value
' 1) prior to requesting any draw or measurement actions, and 2) whenever the control dimensions change.
Public Sub setControlSize(ByVal newWidth As Single, ByVal newHeight As Single)
    
    If (m_ContainerWidth <> newWidth) Or (m_ContainerHeight <> newHeight) Then
        m_ContainerWidth = newWidth
        m_ContainerHeight = newHeight
        m_FontSizeVerified = False
    End If
    
End Sub

'Returns the width of the current caption, in pixels.  No padding is applied.
' Note that this return *ignores the wordwrap setting.*  Wordwrap will (obviously) affect the width of the string, but because the width
' of a wordwrap string is simply the container width - which the caller must specify - there's no point in returning it here.
Public Function getCaptionWidth() As Single
    
    'Make sure the current font object reflects all active settings
    If Not m_FontObjectClean Then rebuildFont
    If Not m_FontSizeVerified Then verifyFontSize
    
    getCaptionWidth = m_TextRenderer.getWidthOfString(m_CaptionTranslated)
    
End Function

'Returns the height of the current caption, in pixels.  No padding is applied.
' Note that this return varies depending on the presence of wordwrapping.  Wordwrap will (obviously) affect the return.
Public Function getCaptionHeight() As Single
    
    'Make sure the current font object reflects all active settings
    If Not m_FontObjectClean Then rebuildFont
    If Not m_FontSizeVerified Then verifyFontSize
    
    If m_WordWrapAllowed Then
        getCaptionHeight = m_TextRenderer.getHeightOfWordwrapString(m_CaptionTranslated, m_ContainerWidth)
    Else
        getCaptionHeight = m_TextRenderer.getHeightOfString(m_CaptionTranslated)
    End If
    
End Function

'When all font and caption settings are ready, use this to actually render the caption onto a target DC
Public Sub drawCaption(ByVal dstDC As Long, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal customColor As Long = -1)
    
    'Make sure the current font object reflects all active settings
    If Not m_FontObjectClean Then rebuildFont
    If Not m_FontSizeVerified Then verifyFontSize
    
    If customColor <> -1 Then
        m_TextRenderer.setFontColor customColor
    Else
        m_TextRenderer.setFontColor m_CaptionColor
    End If
    
    m_TextRenderer.attachToDC dstDC
    
    If m_WordWrapAllowed Then
        m_TextRenderer.fastRenderMultilineText dstX, dstY, m_CaptionTranslated
    Else
        m_TextRenderer.fastRenderText dstX, dstY, m_CaptionTranslated
    End If
    
    m_TextRenderer.releaseFromDC
    
End Sub

'Same as drawCaption, but centered on an arbitrary rect
Friend Sub drawCaptionCentered(ByVal dstDC As Long, ByRef dstRect As RECT, Optional ByVal customColor As Long = -1)
    
    'Make sure the current font object reflects all active settings
    If Not m_FontObjectClean Then rebuildFont
    If Not m_FontSizeVerified Then verifyFontSize
    
    If customColor <> -1 Then
        m_TextRenderer.setFontColor customColor
    Else
        m_TextRenderer.setFontColor m_CaptionColor
    End If
    
    m_TextRenderer.attachToDC dstDC
    
    If m_WordWrapAllowed Then
        
        'See if the string, as-is, fits within the target area
        Dim strWidth As Long
        strWidth = m_TextRenderer.getWidthOfString(m_CaptionTranslated)
        
        If strWidth < (dstRect.Right - dstRect.Left) Then
            m_TextRenderer.drawCenteredTextToRect m_CaptionTranslated, dstRect, True
        Else
        
            'Find the height of the string as it currently exists
            Dim strHeight As Long
            strHeight = m_TextRenderer.getHeightOfWordwrapString(m_CaptionTranslated, dstRect.Right - dstRect.Left)
            
            'Recalculate the rect top so that the text is centered
            Dim newRect As RECT
            newRect.Left = dstRect.Left
            newRect.Bottom = dstRect.Bottom
            newRect.Right = dstRect.Right
            newRect.Top = dstRect.Top + (CSng(CSng(dstRect.Bottom - dstRect.Top) / 2) - CSng(strHeight / 2))
            
            Const DT_CENTER = &H1, DT_WORDBREAK = &H10, DT_NOCLIP = &H100
            m_TextRenderer.DrawTextWrapper StrPtr(m_CaptionTranslated), Len(m_CaptionTranslated), newRect, DT_WORDBREAK Or DT_NOCLIP Or DT_CENTER
            
        End If
        
    Else
        m_TextRenderer.drawCenteredTextToRect m_CaptionTranslated, dstRect, True
    End If
    
    m_TextRenderer.releaseFromDC
    
End Sub

'After font or theme changes, this function must be called to rebuild our font object.
Private Sub rebuildFont()
    
    'Create an initial instance, as necessary
    If m_TextRenderer Is Nothing Then
        
        Set m_TextRenderer = New pdFont
            
        'Prep the font using default settings
        With m_TextRenderer
            
            .setFontItalic False
            .setFontStrikeout False
            .setFontUnderline False
            
            .setTextAlignment vbLeftJustify
            
        End With
            
    'If an instance is already active, delete it
    Else
        If m_TextRenderer.hasFontBeenCreated Then m_TextRenderer.deleteCurrentFont
    End If
    
    'Next, set font properties that may change at run-time
    With m_TextRenderer
        If g_IsProgramRunning Then .setFontFace g_InterfaceFont
        .setFontSize m_FontSizeCaption
        .setFontBold m_FontBold
        .setFontColor m_CaptionColor
    End With
    
    'Create the new font
    m_TextRenderer.createFontObject
    
    'Because the font is newly created, we don't know if we need to shrink the text to fit the container.
    ' Verify the font size immediately if a container width has already been supplied.
    If m_ContainerWidth > 0 Then verifyFontSize
    
End Sub

'Make sure the current caption, drawn with the current font, fits the available rendering area.  This function is pointless if
' 1) the font hasn't been created yet, or 2) we haven't been notified of a container size.
'
'This sub will automatically trigger (1) as necessary, but it can't do a thing about (2)... besides wait for the caller
' to supply us with a valid width!
Private Sub verifyFontSize()

    'Make sure a font object is ready!
    If m_TextRenderer Is Nothing Then rebuildFont
    
    'If the current caption + font size will overflow control boundaries, resize the font automatically
    If m_ContainerWidth > 0 Then
        
        Dim newFontSize As Single
        
        'There are two ways to verify font size, and they differ depending on whether the caller wants us to enable wordwrapping.
        If m_WordWrapAllowed Then
            newFontSize = m_TextRenderer.getMaxFontSizeToFitWordWrap(m_CaptionTranslated, m_ContainerWidth, m_ContainerHeight, m_FontSizeCaption)
        Else
            newFontSize = m_TextRenderer.getMaxFontSizeToFitStringWidth(m_CaptionTranslated, m_ContainerWidth, m_FontSizeCaption)
        End If
        
        With m_TextRenderer
            If .getFontSize <> newFontSize Then
                .deleteCurrentFont
                .setFontSize newFontSize
                .setFontBold m_FontBold
                .createFontObject
            End If
        End With
        
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
    
    'If our translated caption has changed, our font object is potentially dirty and must be recreated.  (Because the new text
    ' may overflow its container, we need to run size heuristics.)  Mark the font as dirty.
    If StrComp(oldTranslatedText, m_CaptionTranslated, vbBinaryCompare) <> 0 Then m_FontObjectClean = False
    
    'If the font object is clean, we won't be recreating it, so instead, update the font color now
    If m_FontObjectClean And Not (m_TextRenderer Is Nothing) Then
    
        If g_IsProgramRunning Then
            m_CaptionColor = g_Themer.getThemeColor(PDTC_TEXT_DEFAULT)
        Else
            m_CaptionColor = RGB(96, 96, 96)
        End If
        
    End If
    
    'The only relevant theme setting for this class is the active font color, which doesn't actually require a font recreation.
    ' Instead, we will trust the user to redraw their control after a theme update, so there's no need to raise a redundant
    ' redraw event now.
        
End Sub
