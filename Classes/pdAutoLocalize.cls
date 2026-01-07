VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdAutoLocalize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Automatic Localization Helper
'Copright 2013-2026 by Tanner Helland
'Created: 19/February/13
'Last updated: 27/June/22
'Last update: rewrite class against the official DeepL.com REST API; note that users must supply their own (free) DeepL API key
'TODO: investigate wqweto's https://github.com/wqweto/VbAsyncSocket as an XP-compatible solution
'      (vs the lazy WinHttpRequest approach I'm using currently)
'
'PhotoDemon's built-in Language Editor tool can interface with the DeepL web service to provide "suggested"
' translations for localizers.  This is especially helpful with long, complex, and/or overtly technical text.
'
'Use of this class requires a valid free DeepL API key.  You must set up a free DeepL account and supply
' the API key you receive to this class via the .SetAPIKey function.
'
'Besides its relatively high translation quality, DeepL was chosen for PhotoDemon localization because
' it explicitly releases copyright on all translated text.  From section 6 of their Terms and Conditions
' (retrieved June 2022):
'
'"6.6 DeepL does not assume any copyrights to the translations made by Customer using the Products.
' In the event that the translations made by Customer using the Products are deemed to be protected under
' copyright laws to the benefit of DeepL, DeepL grants to Customer, upon creation of such translations,
' all exclusive, transferable, sublicensable, worldwide perpetual rights to use the translations without
' limitation and for any existing or future types of use, including without limitation the right to modify
' the translations and to create derivative works."
'
'This is somewhat rare among free localization tools.  I am grateful to DeepL for this licensing decision.
'
'Thank you to everyone who has contributed fixes to this class to improve behavior across different locales,
' in particular ChenLin and Zhu JinYong.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit
Option Compare Binary

'Using a system WinHttpRequest object means this class will likely not work on XP due to lack of modern TLS support.
Private m_httpRequest As Object

'This class supports translation between any arbitrary languages, but obviously the online service you call
' (currently DeepL) needs to support both languages too!
'
'Specify source and destination languages by calling the setSrcLanguage and setDstLanguage functions.
Private m_srcLanguage As String, m_dstLanguage As String

'If we are able to successfully initialize a WinHTTP instance, this will be set to TRUE.  Do not attempt to
' interact with the m_httpRequest object *unless* this is TRUE.
Private m_ServiceAvailable As Boolean

'The user needs to supply an API key for the target translation service; without this, the class WILL NOT WORK
Private m_apiKey As String

Public Function GetAPIKey() As String
    GetAPIKey = m_apiKey
End Function

Public Sub SetAPIKey(ByRef srcKey As String)
    m_apiKey = srcKey
End Sub

Public Sub SetDstLanguage(ByRef dstLangID As String)
    m_dstLanguage = dstLangID
End Sub

Public Sub SetSrcLanguage(ByRef srcLangID As String)
    m_srcLanguage = srcLangID
End Sub

Private Sub Class_Initialize()
    
    On Error GoTo HttpRequestNotAvailable
    
    'By default, translate from English to English - we will update the destination language if/when the caller
    ' chooses a language file
    m_srcLanguage = "en"
    m_dstLanguage = "en"
    
    'Prepare an http request object and cache it at class level
    Set m_httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    m_ServiceAvailable = (Not m_httpRequest Is Nothing)
    If (Not m_ServiceAvailable) Then PDDebug.LogAction "WARNING!  clsAutoLocalize failed to create a WinHttpRequest instance."
    
    Exit Sub
    
HttpRequestNotAvailable:
    m_ServiceAvailable = False
    PDDebug.LogAction "WARNING!  clsAutoLocalize failed to initialize.  Last error was #" & Err.Number & ", " & Err.Description
    
End Sub

'Given a string, return a corresponding DeepL translation (per the class-level source and destination language codes).
' IMPORTANTLY: you must have already specified a valid free DeepL API key via the .SetAPIKey function.
' (Note that the "free" part is important - you'll need to slightly modify the URL, below, if to use a paid Pro key).
Public Function GetDeepLTranslation(ByVal srcText As String) As String
    
    'Anything involving the web can fail unexpectedly
    On Error GoTo NoTranslationAvailable
    
    'To ensure some PhotoDemon text formatting quirks are OK, we pre- and post-process text.  You may need to modify
    ' this for your own purposes.
    Dim returnText As String
    returnText = vbNullString
    
    'If an http request object wasn't initialized successfully, this whole function is pointless
    If m_ServiceAvailable Then
        
        'Manually escape some chars in the source text.  This list is not exhaustive - use UrlEscapeW for that:
        ' https://docs.microsoft.com/en-us/windows/win32/api/shlwapi/nf-shlwapi-urlescapew
        ' - ...but is intended to cover common behavior in PhotoDemon text.
        
        'PhotoDemon uses %n to define custom parameters in text entries.  URLs also make use of the % identifier,
        ' so replace our % with something else before attempting to translate.
        If (InStr(srcText, "%") <> 0) Then srcText = Replace$(srcText, "%", "%25")
        
        'Some PhotoDemon phrases contain line breaks.  We can't pass these in a URL, so replace 'em.
        If (InStr(srcText, vbCrLf) <> 0) Then srcText = Replace$(srcText, vbCrLf, "%0A")
        If (InStr(srcText, vbLf) <> 0) Then srcText = Replace$(srcText, vbLf, "%0A")
        If (InStr(srcText, vbCr) <> 0) Then srcText = Replace$(srcText, vbCr, "%0A")
        
        'Also replace some other problematic forbidden chars.  (This could likely be replaced
        ' with a call to InternetCanonicalizeUrlW() but I haven't studied it in-depth...)
        If (InStr(srcText, ":") <> 0) Then srcText = Replace$(srcText, ":", "%3A")
        If (InStr(srcText, "/") <> 0) Then srcText = Replace$(srcText, "/", "%2F")
        If (InStr(srcText, "\") <> 0) Then srcText = Replace$(srcText, "\", "%5C")
        If (InStr(srcText, "?") <> 0) Then srcText = Replace$(srcText, "?", "%3F")
        If (InStr(srcText, "#") <> 0) Then srcText = Replace$(srcText, "#", "%23")
        If (InStr(srcText, "[") <> 0) Then srcText = Replace$(srcText, "[", "%5B")
        If (InStr(srcText, "]") <> 0) Then srcText = Replace$(srcText, "]", "%5D")
        If (InStr(srcText, "@") <> 0) Then srcText = Replace$(srcText, "@", "%40")
        If (InStr(srcText, "!") <> 0) Then srcText = Replace$(srcText, "!", "%21")
        If (InStr(srcText, "$") <> 0) Then srcText = Replace$(srcText, "$", "%24")
        If (InStr(srcText, "&") <> 0) Then srcText = Replace$(srcText, "&", "%26")
        If (InStr(srcText, "'") <> 0) Then srcText = Replace$(srcText, "'", "%27")
        If (InStr(srcText, "(") <> 0) Then srcText = Replace$(srcText, "(", "%28")
        If (InStr(srcText, ")") <> 0) Then srcText = Replace$(srcText, ")", "%29")
        If (InStr(srcText, "*") <> 0) Then srcText = Replace$(srcText, "*", "%2A")
        If (InStr(srcText, "+") <> 0) Then srcText = Replace$(srcText, "+", "%2B")
        If (InStr(srcText, ",") <> 0) Then srcText = Replace$(srcText, ",", "%2C")
        If (InStr(srcText, ";") <> 0) Then srcText = Replace$(srcText, ";", "%3B")
        If (InStr(srcText, "=") <> 0) Then srcText = Replace$(srcText, "=", "%3D")
        If (InStr(srcText, " ") <> 0) Then srcText = Replace$(srcText, " ", "%20")
        If (InStr(srcText, """") <> 0) Then srcText = Replace$(srcText, """", "%22")
        
        'Return nothing on failure
        GetDeepLTranslation = vbNullString
        
        'Construct a full GET string against v2 of the DeepL API.
        ' (TODO: stop being lazy and solve a proper POST approach instead)
        Dim fullRequest As pdString
        Set fullRequest = New pdString
        
        fullRequest.Append "https://api-free.deepl.com/v2/translate?auth_key="
        fullRequest.Append Replace$(m_apiKey, ":", "%3A")
        fullRequest.Append "&text="
        fullRequest.Append srcText
        fullRequest.Append "&target_lang=" & m_dstLanguage
        fullRequest.Append "&source_lang=" & m_srcLanguage
        fullRequest.Append "&preserve_formatting=1"
        
        m_httpRequest.Open "GET", fullRequest.ToString, False
        
        'WinHTTP can fail randomly here, and that's OK: just note that we don't retry on a failure,
        ' by design.  (PD's UI just displays "translation failed", after which the user can freely try again.)
        m_httpRequest.Send
        
        'We need translations as UTF-8, so *DO NOT* use .ResponseText
        Dim responseBytes As Variant
        responseBytes = m_httpRequest.ResponseBody
        
        'VarPtr is screwy on variant array elements, so perform a (silly) copy into a strongly typed array,
        ' then use *that* to pass pointers
        If (UBound(responseBytes) >= 0) Then
        
            Dim tmpArray() As Byte
            tmpArray = responseBytes
            returnText = Strings.StringFromUTF8Ptr(VarPtr(tmpArray(LBound(tmpArray))), (UBound(tmpArray) - LBound(tmpArray) + 1))
            
            'The returned text is JSON.  A smart solution here would be full JSON parsing, but for now I am
            ' lazily grabbing just the known translation item.
            Dim idxStart As Long, idxEnd As Long
            Const TRNS_RETURN As String = """text"":""", TRNS_END As String = """}]}"
            idxStart = InStr(1, LCase$(returnText), TRNS_RETURN, vbBinaryCompare) + Len(TRNS_RETURN)
            idxEnd = InStrRev(returnText, TRNS_END, -1, vbBinaryCompare)
            
            If (idxEnd > idxStart) Then
                returnText = Mid$(returnText, idxStart, idxEnd - idxStart)
            Else
                If (InStr(1, returnText, "Value for 'target_lang' not supported.", vbTextCompare) <> 0) Then
                    returnText = "[this language is not supported by DeepL]"
                Else
                    returnText = vbNullString
                End If
            End If
            
            'Note that these post-processing steps are specific to DeepL.  Other services may have their own quirks.
            If (LenB(returnText) <> 0) Then
            
                If (InStr(returnText, "\n") <> 0) Then returnText = Replace$(returnText, "\n", vbCrLf)
                If (InStr(returnText, "\""") <> 0) Then returnText = Replace$(returnText, "\""", """")
                
                'Thank you to Zhu JY for pointing out that guillemet (<</>> used as single-char quotation marks)
                ' are not handled properly under some localizations; let's use specific character codes to cover these.
                If (InStr(returnText, ChrW$(174)) <> 0) Then returnText = Replace$(returnText, ChrW$(174), "'")
                If (InStr(returnText, ChrW$(175)) <> 0) Then returnText = Replace$(returnText, ChrW$(175), "'")
                
                'Finally, normalize the return string to form C (composed, the Windows default)
                returnText = Strings.StringNormalize(returnText)
                
            End If
            
        Else
            PDDebug.LogAction "Null return from server"
        End If
        
    Else
        PDDebug.LogAction "WinHttpRequest unavailable"
    End If
    
NoTranslationAvailable:
    
    If (Err.Number <> 0) Then PDDebug.LogAction "An unknown error occurred in pdAutoLocalize: " & Err.Number & ", " & Err.Description
    
    'Return the translated data, if any
    GetDeepLTranslation = returnText
    
End Function
