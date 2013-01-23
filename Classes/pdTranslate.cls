VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Central Language and Translation Handler
'Copyright ©2011-2012 by Frank Donckers and Tanner Helland
'Created: 12/21/12
'Last updated: 22/January/13
'Last update: began official implementation
'
'Thanks to the incredible work of Frank Donckers, PhotoDemon now has a working translation engine.
'
'Most of the crucial translation work happens inside a global object called "g_Language", which is an object
' of this class (pdTranslate).
'
'Better description forthcoming once I actually have some code in place...
'
'***************************************************************************

Option Explicit

'Language-related API calls
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

'Constants related to API language calls
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SISO639LANGNAME As Long = &H59   'ISO 639 abbreviated language name, e.g. "en" for "English"
Private Const LOCALE_SISO3166CTRYNAME As Long = &H5A   'ISO 4166 abbreviated country/region name, e.g. "US" for "United States"

'Which language file is currently active.  This may not be the same one as the INI file (for example, when the
' program is first run, PhotoDemon has to estimate which language file to use).

'Determine which language to use.  This can be done one of several ways:
' 1) If the user has already specified a language, use that one (obviously)
' 2) If the user has not specified a language, and this is not the first time they have run the program, default to English.
' 3) If this is the first time PhotoDemon has been run on this machine, estimate what language to present based on the
'     current system language.  If a matching language is available, present that.  If not, present English.
'     (In the future, the user will be presented with an actual "choose your language" dialog if the system language
'      is NOT set to English.)
Public Sub DetermineLanguage()

    'FOR TESTING ONLY!!!
    'SET FIRSTRUN TO TRUE TO TEST LOCALE CHECKING
    g_IsFirstRun = True

    'Is this the first time PhotoDemon has been run?  If it is, check the system language.
    If g_IsFirstRun Then
    
        Dim curLangID As String
        curLangID = getDefaultUserLanguage()
        
        'TODO: Compare the full language ID (language and region) against language files available in the
        '       Languages folder.  If an exact match is found, present that as the default program language.
        
        'TODO: If an exact match is not found, try comparing just the language part of the code.  If a match
        '       is found, present that as the default language.
        
        'TODO: If no match is found, present the language selection screen in English.
        
    End If

    'Before doing anything else, check to see if the user has already specified a language to use.
    '<Future INI stuff here>
    

End Sub

'ALWAYS CALL DETERMINE LANGUAGE BEFORE THIS STAGE!
'Once a PhotoDemon language has been determined - by checking a saved value in the INI or by reading the
' default user language - this function can be called to "apply" that language to the program.
'If the language is changed, this function must be called again to set the new language program-wide.
Public Sub ApplyLanguage()

    'DetermineLanguage() is responsible for actually choosing the language to be applied.

End Sub

'Check the current default user language, and return it as a standard language code, e.g. "en" or "en-GB" or "de-CH"
Private Function getDefaultUserLanguage() As String

    'First, retrieve the user's current language.
    Dim langReturn As Integer
    langReturn = GetUserDefaultLangID()
    
    'We now need to deconstruct that numerical, Microsoft-specific ID into a standard ISO locale ID.  Rather
    ' than build our own table, we can use Windows to do this for us.  The results may not always be perfect,
    ' but they should be "close enough" to estimate the best language to suggest.
    Dim abbrLang As String, abbrRegion As String
    
    'Get the ISO 639 abbreviated language name (e.g. "en" for "English")
    abbrLang = LCase(GetSpecificLocaleInfo(LOCALE_SISO639LANGNAME))
    
    'Get the ISO 4166 abbreviated country/region name (e.g. "US" for "United States")
    abbrRegion = LCase(GetSpecificLocaleInfo(LOCALE_SISO3166CTRYNAME))
    
    'Return the language codes
    getDefaultUserLanguage = abbrLang & "-" & abbrRegion

End Function

'Given a locale information constant, return a corresponding string
Private Function GetSpecificLocaleInfo(ByVal lInfo As Long) As String
    Dim sBuffer As String, sRet As String
    sBuffer = String$(256, 0)
    sRet = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, sBuffer, Len(sBuffer))
    If sRet > 0 Then
        GetSpecificLocaleInfo = Left$(sBuffer, sRet - 1)
    Else
        GetSpecificLocaleInfo = ""
    End If
End Function
