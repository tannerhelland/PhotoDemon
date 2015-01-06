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
' code in PhotoDemon.  (Frank also provided PhotoDemon's German, French, and Vlaams translations.)

'***************************************************************************
'Google Translate Interface
'Copyright 2013-2015 by Frank Donckers
'Created: 19/February/13
'Last updated: 10/September/13
'Last update: (by Tanner) moved all relevant code into this singular class, various minor modifications
'
'This class requires your project to have a reference to the Microsoft Scripting Library (scrrun.dll).
'
'For Google Translate's Terms of Use, please visit http://www.google.com/policies/terms/
'
'(Note from Tanner: this class relies on Google to provide its translation with very specific formatting.
' Google may change its code at any time, rendering this class ineffective.  Thus I cannot guarantee that
' this code will always work.  Sorry!)
'
'***************************************************************************

Option Explicit

'IE object used to pass our text to Google and capture the return
Private ie As Object

'The IE object may mark itself as "ready" before it's actually ready, so various artificial pauses are applied.
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'This class supports translation between any two languages (assuming Google supports them, obviously).
' Specify which languages by calling the setSrcLanguage and setDstLanguage functions.
Private srcLanguage As String, dstLanguage As String

'If the ie object fails to retrieve a translation, it will sometimes return the previous page instead of the current one.
' To protect against this, we manually compare this translation to the previous one and request a new translation as necessary.
Private lastTranslation As String

Public Sub setSrcLanguage(ByVal srcLangID As String)
    srcLanguage = srcLangID
End Sub

Public Sub setDstLanguage(ByVal dstLangID As String)
    dstLanguage = dstLangID
End Sub

Private Sub Class_Initialize()
    
    'Prepare an Internet Explorer scripting object
    If Not (ie Is Nothing) Then
        ie.quit
        Set ie = Nothing
    End If
    
    Set ie = CreateObject("InternetExplorer.application")
    ie.Visible = False
    
    'By default, translate from English to English
    srcLanguage = "en"
    dstLanguage = "en"
    
    lastTranslation = ""

End Sub

Private Sub Class_Terminate()
    ie.quit
    Set ie = Nothing
End Sub

'Given a string, return the Google Translate equivalent (per the set source and destination language codes).
' Optionally, the HTML of the translation can also be returned for further parsing.
Public Function getGoogleTranslation(ByVal srcText As String, Optional ByRef copyOfHtml As String) As String

    Dim origHTML As String, processedHTML As String
    origHTML = ""
    processedHTML = ""
    
    'PhotoDemon uses %n to define custom parameters in text entries.  URLs also make use of the % identifier,
    ' so replace our % with something else before attempting to translate.
    If InStr(srcText, "%") <> 0 Then srcText = Replace(srcText, "%", "&#37;", , , vbBinaryCompare)
    
    'Some PhotoDemon phrases contain line breaks.  We can't pass these in a URL, so replace them with a custom ID.
    If InStr(srcText, vbCrLf) <> 0 Then srcText = Replace(srcText, vbCrLf, "VBCRLF")
    If InStr(srcText, vbLf) <> 0 Then srcText = Replace(srcText, vbLf, "VBCRLF")
    If InStr(srcText, vbCr) <> 0 Then srcText = Replace(srcText, vbCr, "VBCRLF")
    
    getGoogleTranslation = ""
    
    'If translation fails, we will try again twice.  (Could be changed to as many times as you want, but twice is sufficient IMO.)
    Dim numOfRetries As Long
    numOfRetries = 0
    
RetryTranslation:
    
    'Use the ie scripting object to access Google translate
    ie.navigate "http://translate.google.com/#" & srcLanguage & "/" & dstLanguage & "/" & srcText
    
    'Loop until the ie object has received the translation from Google.  The amount of time this requires is
    ' obviously dependent on this PC's Internet connection.
    Do While ie.busy Or (Not ie.ReadyState = 4)
        
        'Briefly wait to prevent spamming the CPU with repeated loop calls
        Sleep 200
        
    Loop
        
    'Sometimes the ie object claims to be ready, despite the output result not fully being filled.
    ' In an attempt to remedy this, we'll wait just a tiny bit longer...
    Do
        Sleep 200
    Loop Until (ie.ReadyState = 4) And (Not ie.busy)
        
    'Copy the translation (including all HTML elements) into a VB string
    origHTML = ie.Document.getElementById("result_box").innerHTML
    ie.Document.Clear
    
    'If this translation matches the previous one, try again if necessary.
    If InStr(1, origHTML, lastTranslation, vbTextCompare) > 0 Then
        If numOfRetries < 2 Then
            numOfRetries = numOfRetries + 1
            GoTo RetryTranslation
        Else
            lastTranslation = origHTML
        End If
    Else
        lastTranslation = origHTML
    End If
    
    'Return any characters we removed before translation
    If InStr(origHTML, "&#37;") <> 0 Then origHTML = Replace(origHTML, "&#37;", "%", , , vbBinaryCompare)
    If InStr(origHTML, "VBCRLF") <> 0 Then origHTML = Replace(origHTML, "VBCRLF", vbCrLf)
    
    'Remove all irrelevant HTML tags
    processedHTML = stripHTMLTags(origHTML)
    
    'Finally, fix some weird quirks of the Google translate service.  For example, spaces are frequently
    ' added to translated text for no good reason - so remove these in an attempt to preserve the original
    ' text as much as possible.
    If InStr(processedHTML, "&gt;") <> 0 Then processedHTML = Replace(processedHTML, "&gt;", ">", 1, , vbTextCompare)
    If InStr(processedHTML, "&lt;") <> 0 Then processedHTML = Replace(processedHTML, "&lt;", "<", 1, , vbTextCompare)
    If InStr(processedHTML, "&amp;") <> 0 Then processedHTML = Replace(processedHTML, "&amp;", "&", 1, , vbTextCompare)
    If InStr(processedHTML, vbTab) <> 0 Then processedHTML = Replace(processedHTML, vbTab, "", 1, , vbTextCompare)
    If InStr(processedHTML, "( ") <> 0 Then processedHTML = Replace(processedHTML, "( ", "(", 1, , vbTextCompare)
    If InStr(processedHTML, " )") <> 0 Then processedHTML = Replace(processedHTML, " )", ")", 1, , vbTextCompare)
    If InStr(processedHTML, " ,") <> 0 Then processedHTML = Replace(processedHTML, " ,", ",", 1, , vbTextCompare)
    If InStr(processedHTML, " :") <> 0 Then processedHTML = Replace(processedHTML, " :", ":", 1, , vbTextCompare)
    If InStr(processedHTML, " .") <> 0 Then processedHTML = Replace(processedHTML, " .", ".", 1, , vbTextCompare)
    If InStr(processedHTML, " \") <> 0 Then processedHTML = Replace(processedHTML, " \ ", "\", 1, , vbTextCompare)
    If InStr(processedHTML, " /") <> 0 Then processedHTML = Replace(processedHTML, " / ", "/", 1, , vbTextCompare)
    If InStr(processedHTML, "% ") <> 0 Then processedHTML = Replace(processedHTML, "% ", " %", 1, , vbTextCompare)
    
    'Thanks to Zhu JY for pointing out that «/» are not handled properly under some localizations, so use their
    ' character codes instead.
    If InStr(processedHTML, ChrW$(174)) <> 0 Then processedHTML = Replace(processedHTML, ChrW$(174), "'", 1, , vbTextCompare)
    If InStr(processedHTML, ChrW$(175)) <> 0 Then processedHTML = Replace(processedHTML, ChrW$(175), "'", 1, , vbTextCompare)
    
    'Return the translated data
    getGoogleTranslation = processedHTML
    
    'Optionally, return the original translation HTML as well.
    copyOfHtml = origHTML

End Function

'Simple function to strip all HTML tags from a string.  Angle brackets that are part of content should
' (mostly) be properly preserved by this function.
Private Function stripHTMLTags(ByVal srcString As String) As String

    Dim curPosition As Long
    curPosition = InStr(1, srcString, "<")
    
    Dim closePosition As Long, nextPosition As Long
    
    Do While curPosition > 0
        
        'curPosition points to the location of a "<" sign.  We need to find the matching close bracket, with the
        ' following condition: if we find another "<" sign before finding a ">" sign, this "<" sign is part of
        ' the document content - so do not remove it.
        
        'Find the close bracket
        closePosition = InStr(curPosition + 1, srcString, ">")
        
        'Find the next "<"
        nextPosition = InStr(curPosition + 1, srcString, "<")
        If nextPosition = 0 Then nextPosition = Len(srcString) + 1
        
        If closePosition < nextPosition Then
        
            'Remove this tag
            srcString = Left$(srcString, curPosition - 1) & Right$(srcString, Len(srcString) - closePosition)
        
        End If
    
        'Find the next occurrence
        curPosition = InStr(curPosition, srcString, "<")
    
    Loop
    
    stripHTMLTags = srcString

End Function
