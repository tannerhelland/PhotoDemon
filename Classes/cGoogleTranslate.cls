VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGoogleTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.  It may require modifications to work with other projects.
'
'This code was originally written by the talented Frank Donckers.  Many thanks to Frank for allowing me to use his
' code in PhotoDemon.  (Frank also provided PhotoDemon's original German, French, and Vlaams translations.)

'***************************************************************************
'Google Translate Interface
'Copyright 2013-2020 by Frank Donckers
'Created: 19/February/13
'Last updated: 12/December/17
'Last update: harden against IE initialization errors
'
'This class only works if IE is installed and functional.  Other browsers don't currently work.
'
'For Google Translate's Terms of Use, please visit http://www.google.com/policies/terms/
'
'Thank you to everyone who has contributed fixes for different locales, including ChenLin and Zhu JinYong.
'
'(Note from Tanner: this class relies on Google to provide its translation with very specific formatting.
' Google may change its code at any time, rendering this class ineffective.  Thus I cannot guarantee that
' this code will always work.  Sorry!)
'
'***************************************************************************

Option Explicit

'IE object used to pass our text to Google and capture the return.  Note that Internet Explorer *can* be
' uninstalled by determined users, so this function is not guaranteed to work.
Private m_IE As Object

'The IE object may mark itself as "ready" before it's actually ready, so various artificial pauses are applied.
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'This class supports translation between any two languages (assuming Google supports them, obviously).
' Specify which languages by calling the setSrcLanguage and setDstLanguage functions.
Private m_srcLanguage As String, m_dstLanguage As String

'If the ie object fails to retrieve a translation, it will sometimes return the previous page instead of the current one.
' To protect against this, we manually compare this translation to the previous one and request a new translation as necessary.
Private m_lastTranslation As String

'If we are able to successfully initialize an IE instance, this will be set to TRUE.  Do not attempt to
' interact with the m_IE object *unless* this is TRUE.
Private m_IEAvailable As Boolean

Public Sub SetSrcLanguage(ByRef srcLangID As String)
    m_srcLanguage = srcLangID
End Sub

Public Sub SetDstLanguage(ByRef dstLangID As String)
    m_dstLanguage = dstLangID
End Sub

Private Sub Class_Initialize()
    
    On Error GoTo IENotAvailable
    
    'By default, translate from English to English
    m_srcLanguage = "en"
    m_dstLanguage = "en"
    m_lastTranslation = vbNullString
    
    'Prepare an Internet Explorer scripting object
    If (Not m_IE Is Nothing) Then
        m_IE.quit
        Set m_IE = Nothing
    End If
    
    'IE is now uninstallable, so it is possible for this line to fail
    Set m_IE = CreateObject("InternetExplorer.application")
    
    If (Not m_IE Is Nothing) Then
        m_IE.Visible = False
        m_IEAvailable = True
    Else
        m_IEAvailable = False
        PDDebug.LogAction "WARNING!  clsGoogleTranslate failed to create an InternetExplorer.application instance."
    End If
    
    Exit Sub
    
IENotAvailable:
    
    m_IEAvailable = False
    PDDebug.LogAction "WARNING!  clsGoogleTranslate could not initialize an IE instance.  Last error was #" & Err.Number & ", " & Err.Description
    
End Sub

Private Sub Class_Terminate()
    
    On Error GoTo IEWasntInitialized
    
    If (Not m_IE Is Nothing) Then
        m_IE.quit
        Set m_IE = Nothing
    End If
    
IEWasntInitialized:
    m_IEAvailable = False
    
End Sub

'Given a string, return the Google Translate equivalent (per the set source and destination language codes).
' Optionally, the HTML of the translation can also be returned for further parsing.
Public Function GetGoogleTranslation(ByVal srcText As String, Optional ByRef copyOfHtml As String) As String
    
    On Error GoTo NoTranslationAvailable
    
    Dim origHTML As String, processedHTML As String
    origHTML = vbNullString
    processedHTML = vbNullString
    
    'If an IE object wasn't initialized successfully, this whole function is pointless
    Dim ieAvailable As Boolean
    ieAvailable = m_IEAvailable
    If ieAvailable Then ieAvailable = (Not m_IE Is Nothing)
    
    If ieAvailable Then
        
        'PhotoDemon uses %n to define custom parameters in text entries.  URLs also make use of the % identifier,
        ' so replace our % with something else before attempting to translate.
        If (InStr(srcText, "%") <> 0) Then srcText = Replace$(srcText, "%", "&#37;", , , vbBinaryCompare)
        
        'Some PhotoDemon phrases contain line breaks.  We can't pass these in a URL, so replace them with a custom ID.
        If (InStr(srcText, vbCrLf) <> 0) Then srcText = Replace$(srcText, vbCrLf, "VBCRLF")
        If (InStr(srcText, vbLf) <> 0) Then srcText = Replace$(srcText, vbLf, "VBCRLF")
        If (InStr(srcText, vbCr) <> 0) Then srcText = Replace$(srcText, vbCr, "VBCRLF")
        
        'Also replace some other problematic forbidden chars
        If (InStr(srcText, "\") <> 0) Then processedHTML = Replace$(srcText, "\", "&#92;", 1, , vbBinaryCompare)
        If (InStr(srcText, "/") <> 0) Then processedHTML = Replace$(srcText, "/", "&#47;", 1, , vbBinaryCompare)
        
        GetGoogleTranslation = vbNullString
        
        'If translation fails, we will try again twice.  (Could be changed to as many times as you want, but twice is sufficient IMO.)
        Dim numOfRetries As Long
        numOfRetries = 0
        
RetryTranslation:
        
        'Use the ie scripting object to access Google translate.  Note that special care must be taken for
        ' Chinese languages, if the user originates in China; we must use the .cn URL there or Google will fail.
        Dim useChineseEngine As Boolean: useChineseEngine = False
        useChineseEngine = (InStr(1, LCase$(m_dstLanguage), "zh", vbBinaryCompare) <> 0) Or (InStr(1, LCase$(m_dstLanguage), "cn", vbBinaryCompare) <> 0)
        
        Dim fullTargetURL As String
        
        If useChineseEngine Then
            fullTargetURL = "http://translate.google.cn/#" & m_srcLanguage & "/" & m_dstLanguage & "/" & srcText
        Else
            fullTargetURL = "http://translate.google.com/#" & m_srcLanguage & "/" & m_dstLanguage & "/" & srcText
        End If
        
        m_IE.navigate fullTargetURL
        
        'Loop until the ie object has received the translation from Google.  The amount of time this requires is
        ' obviously dependent on this PC's Internet connection.
        Do While m_IE.busy Or (Not m_IE.ReadyState = 4)
            
            'Briefly wait to prevent spamming the CPU with repeated loop calls
            Sleep 200
            
        Loop
            
        'Sometimes the ie object claims to be ready, despite the output result not fully being filled.
        ' In an attempt to remedy this, we'll wait just a tiny bit longer...
        Do
            Sleep 200
        Loop Until (m_IE.ReadyState = 4) And (Not m_IE.busy)
            
        'Copy the translation (including all HTML elements) into a VB string
        origHTML = m_IE.Document.getElementById("result_box").innerHTML
        
        'If this translation matches the previous one, try again if necessary.
        If (InStr(1, origHTML, m_lastTranslation, vbTextCompare) <> 0) Then
            If (numOfRetries < 2) Then
                numOfRetries = numOfRetries + 1
                GoTo RetryTranslation
            Else
                m_lastTranslation = origHTML
            End If
        Else
            m_lastTranslation = origHTML
        End If
        
        'Return any characters we removed before translation
        If (InStr(origHTML, "&#37;") <> 0) Then origHTML = Replace$(origHTML, "&#37;", "%", , , vbBinaryCompare)
        If (InStr(origHTML, "VBCRLF") <> 0) Then origHTML = Replace$(origHTML, "VBCRLF", vbCrLf)
        If (InStr(1, origHTML, "&#92;", vbBinaryCompare) <> 0) Then processedHTML = Replace$(origHTML, "&#92;", "\", 1, , vbBinaryCompare)
        If (InStr(1, origHTML, "&#47;", vbBinaryCompare) <> 0) Then processedHTML = Replace$(origHTML, "&#47;", "/", 1, , vbBinaryCompare)
        
        'Remove all irrelevant HTML tags
        processedHTML = StripHTMLTags(origHTML)
        
        'Finally, fix some weird quirks of the Google translate service.  For example, spaces are frequently
        ' added to translated text for no good reason - so remove these in an attempt to preserve the original
        ' text as much as possible.
        If (InStr(processedHTML, "&gt;") <> 0) Then processedHTML = Replace$(processedHTML, "&gt;", ">", 1, , vbTextCompare)
        If (InStr(processedHTML, "&lt;") <> 0) Then processedHTML = Replace$(processedHTML, "&lt;", "<", 1, , vbTextCompare)
        If (InStr(processedHTML, "&amp;") <> 0) Then processedHTML = Replace$(processedHTML, "&amp;", "&", 1, , vbTextCompare)
        If (InStr(processedHTML, vbTab) <> 0) Then processedHTML = Replace$(processedHTML, vbTab, vbNullString, 1, , vbTextCompare)
        If (InStr(processedHTML, "( ") <> 0) Then processedHTML = Replace$(processedHTML, "( ", "(", 1, , vbTextCompare)
        If (InStr(processedHTML, " )") <> 0) Then processedHTML = Replace$(processedHTML, " )", ")", 1, , vbTextCompare)
        If (InStr(processedHTML, " ,") <> 0) Then processedHTML = Replace$(processedHTML, " ,", ",", 1, , vbTextCompare)
        If (InStr(processedHTML, " :") <> 0) Then processedHTML = Replace$(processedHTML, " :", ":", 1, , vbTextCompare)
        If (InStr(processedHTML, " .") <> 0) Then processedHTML = Replace$(processedHTML, " .", ".", 1, , vbTextCompare)
        If (InStr(processedHTML, " \") <> 0) Then processedHTML = Replace$(processedHTML, " \ ", "\", 1, , vbTextCompare)
        If (InStr(processedHTML, " /") <> 0) Then processedHTML = Replace$(processedHTML, " / ", "/", 1, , vbTextCompare)
        If (InStr(processedHTML, "% ") <> 0) Then processedHTML = Replace$(processedHTML, "% ", " %", 1, , vbTextCompare)
        
        'Thanks to Zhu JY for pointing out that guillemet (<</>> used as single-char quotation marks)
        ' are not handled properly under some localizations; we'll use specific character codes instead.
        If (InStr(processedHTML, ChrW$(174)) <> 0) Then processedHTML = Replace$(processedHTML, ChrW$(174), "'", 1, , vbTextCompare)
        If (InStr(processedHTML, ChrW$(175)) <> 0) Then processedHTML = Replace$(processedHTML, ChrW$(175), "'", 1, , vbTextCompare)
    
    End If
    
NoTranslationAvailable:
    
    If (Err.Number <> 0) Then
        
        'An "object required" error means we could find the target tag; Google Translate has probably changed
        ' their HTML again, and we are bad and don't use their recommended API so we get what we deserve.
        If (Err.Number = 424) Then
            PDDebug.LogAction "Google Translate changed their page layout; the auto-translate function won't work for this language."
        Else
            PDDebug.LogAction "An unknown error occurred in clsGoogleTranslate: " & Err.Number & ", " & Err.Description
        End If
        
    End If
    
    'Return the translated data
    GetGoogleTranslation = processedHTML
    
    'Optionally, return the original translation HTML as well.
    copyOfHtml = origHTML

End Function

'Simple function to strip all HTML tags from a string.  Angle brackets that are part of content should
' (mostly) be properly preserved by this function.
Private Function StripHTMLTags(ByVal srcString As String) As String
    
    If (LenB(srcString) <> 0) Then
    
        Dim curPosition As Long
        curPosition = InStr(1, srcString, "<")
        
        Dim closePosition As Long, nextPosition As Long
        
        Do While (curPosition > 0)
            
            'curPosition points to the location of a "<" sign.  We need to find the matching close bracket, with the
            ' following condition: if we find another "<" sign before finding a ">" sign, this "<" sign is part of
            ' the document content - so do not remove it.
            
            'Find the close bracket
            closePosition = InStr(curPosition + 1, srcString, ">")
            
            'Find the next "<"
            nextPosition = InStr(curPosition + 1, srcString, "<")
            If (nextPosition = 0) Then nextPosition = Len(srcString) + 1
            
            If (closePosition < nextPosition) Then
            
                'Remove this tag
                srcString = Left$(srcString, curPosition - 1) & Right$(srcString, Len(srcString) - closePosition)
            
            End If
        
            'Find the next occurrence
            curPosition = InStr(curPosition, srcString, "<")
        
        Loop
    
    End If
    
    StripHTMLTags = srcString

End Function
