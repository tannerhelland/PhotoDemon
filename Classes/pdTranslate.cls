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

'Basic information about a language file is stored in this type of variable
Private Type pdLanguageFile
    filename As String
    langID As String
    langName As String
    langVersion As String
    langStatus As String
End Type

'Language-related API calls
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

'Constants related to API language calls
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SISO639LANGNAME As Long = &H59   'ISO 639 abbreviated language name, e.g. "en" for "English"
Private Const LOCALE_SISO3166CTRYNAME As Long = &H5A   'ISO 4166 abbreviated country/region name, e.g. "US" for "United States"

'Array of currently available languages
Private numOfLanguages As Long
Private availableLanguages() As pdLanguageFile

'Array ID of the current language.  If no language is selected (e.g. PhotoDemon is using its default text), this value is -1
Private curLanguage As Long

'Array ID of the default language (English - US)
Private defaultLanguage As Long

'If the user needs to be shown a language select screen, this will be set to true at some point in the initialization process
Private needToShowLanguageDialog As Boolean

'If the select language dialog needs to be shown, this value can be used to determine "why"
Private Const LANGREASON_USERINITIATED As Long = 0      ' 0) User-initiated dialog
Private Const LANGREASON_APPROXIMATEMATCH As Long = 1   ' 1) Approximate (but not exact) language match found
Private Const LANGREASON_NOMATCH As Long = 2            ' 2) No language match found
Private Const LANGREASON_FILEMISSING As Long = 3        ' 3) Requested language file couldn't be located
Private reasonForLanguageDialog As Long


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
    'g_IsFirstRun = True

    'Is this the first time PhotoDemon has been run?  If it is, check the system language.
    If g_IsFirstRun Then
    
        Dim curLangID As String
        curLangID = getDefaultUserLanguage()
                
        'Compare the full language ID (language and region) against language files available in the Languages folder.
        ' If an exact match is found, present that as the default program language.
        curLanguage = isExactLanguageMatch(curLangID)
        
        'If an exact match is not found, try comparing just the language part of the code.  If a match is found,
        ' present that as the default language - and note that we also need to show the language selection dialog.
        If curLanguage = -1 Then
            curLanguage = isApproximateLanguageMatch(curLangID)
            reasonForLanguageDialog = LANGREASON_APPROXIMATEMATCH
            needToShowLanguageDialog = True
        End If
        
        'TODO: If no match is found, present the language selection screen in English.
        If curLanguage = -1 Then
            curLanguage = defaultLanguage
            reasonForLanguageDialog = LANGREASON_NOMATCH
            needToShowLanguageDialog = True
        End If
        
    'If this is not the first run, retrieve the user's language preference from the INI file
    Else
    
        Dim curLanguageFile As String
        curLanguageFile = g_UserPreferences.GetPreference_String("General Preferences", "CurrentLanguageFile", "")
        
        'If no specific file has been named, use the default language
        If curLanguageFile = "" Then
            curLanguage = defaultLanguage
        
        'If a file HAS been named, make sure it still exists; if it doesn't, default to English (US) and show the language dialog
        Else
            If FileExist(curLanguageFile) Then
                curLanguage = getLangIndexFromFile(curLanguageFile)
            Else
                curLanguage = defaultLanguage
                reasonForLanguageDialog = LANGREASON_FILEMISSING
                needToShowLanguageDialog = True
            End If
        End If
    
    End If

End Sub

'Given a language filename, return the index in the availableLanguages() array
Private Function getLangIndexFromFile(ByVal srcFile As String) As Long
    
    Dim i As Long
    For i = 0 To numOfLanguages - 1
        If StrComp(availableLanguages(i).filename, g_UserPreferences.getLanguagePath & srcFile, vbTextCompare) = 0 Then
            getLangIndexFromFile = i
            Exit Function
        End If
    Next i
    
End Function

'Given a language ID (as a string), try to find an approximate match (just language, region doesn't matter) in the availableLanguages array
Private Function isApproximateLanguageMatch(ByRef srcLangID As String)

    Dim i As Long
    For i = 0 To numOfLanguages - 1
        If StrComp(Left(srcLangID, 2), Left(availableLanguages(i).langID, 2), vbTextCompare) = 0 Then
            isApproximateLanguageMatch = i
            'MsgBox "Approximate language match found: " & availableLanguages(i).langID
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, no match was found.  Return -1
    isApproximateLanguageMatch = -1

End Function

'Given a language ID (as a string), try to find an exact match (language AND region) in the availableLanguages array
Private Function isExactLanguageMatch(ByRef srcLangID As String)

    Dim i As Long
    For i = 0 To numOfLanguages - 1
        If StrComp(srcLangID, availableLanguages(i).langID, vbTextCompare) = 0 Then
            isExactLanguageMatch = i
            'MsgBox "Exact language match found: " & availableLanguages(i).langID
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, no match was found.  Return -1
    isExactLanguageMatch = -1

End Function

'ALWAYS CALL DETERMINE LANGUAGE BEFORE THIS STAGE!
'Once a PhotoDemon language has been determined, this function can be called to "apply" that language to the program.
' NOTE: if the language is changed, this function must be called again to set the new language program-wide.
Public Sub ApplyLanguage()

    'Mark the active language menu entry (and unmark all others)
    Dim i As Long
    For i = 0 To numOfLanguages - 1
        If curLanguage = i Then FormMain.mnuLanguages(i).Checked = True Else FormMain.mnuLanguages(i).Checked = False
    Next i

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

'Search the Languages folder, and make a list of all available languages
Public Sub CheckAvailableLanguages()
    
    numOfLanguages = 0
    ReDim Preserve availableLanguages(0 To 99) As pdLanguageFile
    
    'Scan the MRU icon directory for .xml files.  Ignore anything that isn't XML.
    Dim chkFile As String, curFileAsText As String
    chkFile = Dir(g_UserPreferences.getLanguagePath & "*.xml", vbNormal)
        
    Do While chkFile <> ""
                
        'Load the first 512 bytes of the file into a string (this is done to keep this routine fast)
        curFileAsText = getPartialFileAsString(g_UserPreferences.getLanguagePath & chkFile)
        
        'If the file has actual contents, investigate further.
        If Len(curFileAsText) > 0 Then
        
            'Is the .xml file a valid xml file?  Check the header to make sure.
            If verifyXMLHeader(curFileAsText) Then
            
                'If this is a valid language file, get the language information (if available).
                fillLanguageInfo curFileAsText, availableLanguages(numOfLanguages)
                availableLanguages(numOfLanguages).filename = chkFile
                numOfLanguages = numOfLanguages + 1
            
            End If
            
        End If
        
        'Retrieve the next file and repeat
        chkFile = Dir
    
    Loop
    
    'Add a dummy entry for PhotoDemon's default language (en-US)
    availableLanguages(numOfLanguages).filename = ""
    availableLanguages(numOfLanguages).langID = "en-US"
    availableLanguages(numOfLanguages).langName = "English (US) - Default"
    availableLanguages(numOfLanguages).langStatus = "Complete"
    availableLanguages(numOfLanguages).langVersion = App.Major & "." & App.Minor & "." & App.Revision
    numOfLanguages = numOfLanguages + 1
    
    'Resize the array to contain only the number of languages found
    ReDim Preserve availableLanguages(0 To numOfLanguages - 1) As pdLanguageFile
    
    'Sort the array alphabetically by language name (because we are going to fill the languages submenu with the entries)
    sortLanguageData
    
    'Populate the languages menu
    populateLanguageMenu
    
    'Mark the default language location, as other functions may need to know it
    Dim i As Long
    For i = 0 To numOfLanguages - 1
        If StrComp(availableLanguages(i).langName, "English (US) - Default", vbTextCompare) = 0 Then defaultLanguage = i
    Next i
    
    'TESTING: display the info we retrieved
    'Dim i As Long
    'For i = 0 To numOfLanguages - 1
    '    MsgBox availableLanguages(i).langID & vbCrLf & availableLanguages(i).langName & vbCrLf & availableLanguages(i).langStatus & vbCrLf & availableLanguages(i).langVersion
    'Next i
    
End Sub

'Populate the Tools -> Languages menu
Private Sub populateLanguageMenu()

    'Start by unloading all current language menu entries
    Dim i As Long
    If FormMain.mnuLanguages.Count > 1 Then
        For i = 0 To FormMain.mnuLanguages.Count - 1
            Unload FormMain.mnuLanguages(i)
        Next i
    End If
    
    'Load all the menu entries we'll be needing.
    If numOfLanguages > 1 Then
        For i = 1 To numOfLanguages - 1
            Load FormMain.mnuLanguages(i)
        Next i
    End If
    
    'Assign names to all the menu entries
    For i = 0 To numOfLanguages - 1
        FormMain.mnuLanguages(i).Caption = availableLanguages(i).langName
        FormMain.mnuLanguages(i).Checked = False
    Next i

End Sub

'Given a language file (or partial language file), extract the key language information and place it in the passed variable.
Private Sub fillLanguageInfo(ByRef fileString As String, ByRef targetLangHolder As pdLanguageFile)

    'First, get the language ID - this is the most important value, and technically the only REQUIRED one.
    targetLangHolder.langID = getTextBetweenTags(fileString, "langid")
    targetLangHolder.langName = getTextBetweenTags(fileString, "langname")
    targetLangHolder.langVersion = getTextBetweenTags(fileString, "langversion")
    targetLangHolder.langStatus = getTextBetweenTags(fileString, "langstatus")
    
    'If, for some reason, the language name is not available, copy the language ID into its place
    If targetLangHolder.langName = "" Then targetLangHolder.langName = UCase(targetLangHolder.langID)
    
End Sub

'Sort the availableLanguages array alphabetically, using language names as the sort parameter
Private Sub sortLanguageData()

    Dim i As Long, j As Long
    
    'Loop through all entries in the languages array, sorting them as we go
    For i = 0 To numOfLanguages - 1
        For j = 0 To numOfLanguages - 1
            
            'Compare two language names, and if one is less (e.g. earlier alphabetically) than the other, swap them
            If StrComp(availableLanguages(i).langName, availableLanguages(j).langName, vbTextCompare) = -1 Then
                SwapLanguageData availableLanguages(i), availableLanguages(j)
            End If
        Next j
    Next i

End Sub

'Swap the values of two language file headers
Private Sub SwapLanguageData(ByRef langOne As pdLanguageFile, ByRef langTwo As pdLanguageFile)
    Dim langTmp As pdLanguageFile
    CopyLanguageData langTmp, langOne
    CopyLanguageData langOne, langTwo
    CopyLanguageData langTwo, langTmp
End Sub

'Copy the contents of one language file header over another
Private Sub CopyLanguageData(ByRef dstLang As pdLanguageFile, ByRef srcLang As pdLanguageFile)
    dstLang.langID = srcLang.langID
    dstLang.langName = srcLang.langName
    dstLang.langStatus = srcLang.langStatus
    dstLang.langVersion = srcLang.langVersion
    dstLang.filename = srcLang.filename
End Sub

'Given a file (as a String) and a tag (without brackets), return the text between that tag.
' NOTE: this function will always return the first occurence of the specified tag, starting at the specified search position.
' If the tag is not found, a blank string will be returned.
Private Function getTextBetweenTags(ByRef fileText As String, ByRef fTag As String, Optional ByVal searchLocation As Long = 1, Optional ByRef whereTagFound As Long = -1) As String

    Dim tagStart As Long, tagEnd As Long
    tagStart = InStr(searchLocation, fileText, "<" & fTag & ">", vbTextCompare)

    'If the tag was found in the file, we also need to find the closing tag.
    If tagStart > 0 Then
    
        tagEnd = InStr(tagStart, fileText, "</" & fTag & ">", vbTextCompare)
        
        'If the closing tag exists, return everything between that and the opening tag
        If tagEnd > tagStart Then
            
            'Increment the tag start location by the length of the tag plus two (+1 for each bracket: <>)
            tagStart = tagStart + Len(fTag) + 2
            
            'If the user passed a long, they want to know where this tag was found - return the location just after the
            ' location where the closing tag was located.
            If whereTagFound <> -1 Then whereTagFound = tagEnd + Len(fTag) + 2
            getTextBetweenTags = Mid(fileText, tagStart, tagEnd - tagStart)
            
        Else
            getTextBetweenTags = "ERROR: specified tag wasn't properly closed!"
        End If
        
    Else
        getTextBetweenTags = ""
    End If

End Function

'Given an XML file (or sometimes, just the first 1024 bytes of an XML file), check to see if it has a valid XML header.
' Also, check to see if it's a PhotoDemon language file or just a random XML file.
Private Function verifyXMLHeader(ByRef fileContents As String) As Boolean
        
    'Check for "<?xml" in the file.  We don't care about encoding, etc - just that to keep things quick.
    If InStr(1, fileContents, "<?xml", vbTextCompare) > 0 Then
    
        'If an XML header is found, check to see if a <language> tag exists too.
        If InStr(1, fileContents, "<language>", vbTextCompare) > 0 Then
        
            'Next, ensure that a language id (tag <langid>) exists
            If InStr(1, fileContents, "<langid>", vbTextCompare) > 0 Then
            
                'And finally, make sure the file contains at least one <phrase> tag
                If InStr(1, fileContents, "<phrase>", vbTextCompare) > 0 Then
                    verifyXMLHeader = True
                Else
                    verifyXMLHeader = False
                End If
            
            Else
                verifyXMLHeader = False
            End If
        
        Else
            verifyXMLHeader = False
        End If
    
    Else
        verifyXMLHeader = False
    End If

End Function

'Retrieve only the first 1024 bytes of a file as a string.
Private Function getPartialFileAsString(ByVal fName As String) As String
        
    ' ensure that the file exists
    If FileExist(fName) Then
        
        Dim fileNum As Integer
        fileNum = FreeFile
        
        Dim fileLen As Long
    
        Open fName For Binary As #fileNum
            fileLen = LOF(fileNum)
            If fileLen > 1024 Then fileLen = 1024
            getPartialFileAsString = Space$(fileLen)
            Get #fileNum, , getPartialFileAsString
        Close #fileNum
    
    Else
        getPartialFileAsString = ""
    End If
    
End Function

'Retrieve an entire file and return it as a string.
Private Function getFileAsString(ByVal fName As String) As String
        
    ' ensure that the file exists
    If FileExist(fName) Then
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        Open fName For Binary As #fileNum
            getFileAsString = Space$(LOF(fileNum))
            Get #fileNum, , getFileAsString
        Close #fileNum
    
    Else
        getFileAsString = ""
    End If
    
End Function
