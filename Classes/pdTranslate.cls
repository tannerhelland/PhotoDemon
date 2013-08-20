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
'Copyright ©2012-2013 by Frank Donckers and Tanner Helland
'Created: 21/December/12
'Last updated: 23/April/13
'Last update: add translation support for text box tooltips
'
'Thanks to the incredible work of Frank Donckers, PhotoDemon now has a working translation engine.  Many
' thanks to Frank for taking the initiative on not only implementing the translation engine prototype, but
' also for taking the time to translate the entire PhotoDemon text collection into multiple languages.
' (This was a huge project, as PhotoDemon contains a LOT of text.)
'
'The latest revision of the translation engine operates on translation files stored as XML.  You can see
' (and freely edit) these files inside the /App/PhotoDemon/Languages/ subfolder of wherever PD is installed.
'
'Many features are on the roadmap for translation, including things like interactive editing and automatic
' updating of translation files, but this version of the engine focuses primarily on just getting the damn
' thing implemented across all the text in the program, including all controls, messages, message boxes,
' and more.
'
'Most of the crucial translation work in the program happens inside a global object called "g_Language",
' which is an object of this class (pdTranslate).  Most subs and functions should be self-explanatory.  To
' really get a feel for how the translation engine works, I suggest starting with the Loading module, and
' looking at the way this class is initialized and loaded.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Basic information about a language file is stored in this type of variable
Private Type pdLanguageFile
    FileName As String
    langID As String
    langName As String
    langVersion As String
    langStatus As String
End Type

'Language-related API calls
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
Private m_needToShowLanguageDialog As Boolean

'If the select language dialog needs to be shown, this value can be used to determine "why"
Private Const LANGREASON_USERINITIATED As Long = 0      ' 0) User-initiated dialog
Private Const LANGREASON_APPROXIMATEMATCH As Long = 1   ' 1) Approximate (but not exact) language match found
Private Const LANGREASON_NOMATCH As Long = 2            ' 2) No language match found
Private Const LANGREASON_FILEMISSING As Long = 3        ' 3) Requested language file couldn't be located
Private m_reasonForLanguageDialog As Long

'The full contents of the translation file will be loaded into this string.
Private m_translationFileContents As String

'Is the class ready to be used for translations?
Private m_isReady As Boolean

'The following boolean values are marked by preProcessString, then read by postProcessString
Private m_hasDotDotDot As Boolean
Private m_hasTrailingColon As Boolean

'Searching the main XML string is a time-consuming process.  Shortcuts can be activated if all text lies on the same form.
Dim m_shortcutActive As Boolean
Dim m_shortcutLocation As Long

'External functions can request the current language, which may then be used to request translations from
' 3rd-party plugins (specifically, ExifTool).
Public Function getCurrentLanguage() As String
    
    If translationActive Then
        getCurrentLanguage = Left$(availableLanguages(curLanguage).langID, 2)
    Else
        getCurrentLanguage = "en"
    End If
    
End Function

'To improve performance, the user can specify a "shortcut" - e.g. the name of the form being processed.  The translation engine
' will then attempt to search in the vicinity of that shortcut, which should help it find a match more quickly
Public Sub activateShortcut(ByVal shortcutText As String)

    m_shortcutLocation = InStr(1, m_translationFileContents, "(" & shortcutText & ")", vbTextCompare)
    'MsgBox m_shortcutLocation
    If m_shortcutLocation = 0 Then
        m_shortcutActive = False
        m_shortcutLocation = 1
    End If

End Sub

Public Sub deactivateShortcut()
    m_shortcutActive = False
    m_shortcutLocation = 1
End Sub

'External function for asking the translator if a language selection dialog is needed
Public Function isLanguageDialogNeeded(Optional ByRef reasonForDialog As Long = -1) As Boolean
    If reasonForDialog <> -1 Then reasonForDialog = m_reasonForLanguageDialog
    isLanguageDialogNeeded = m_needToShowLanguageDialog
End Function

'If the class has successfully initialized a translation file, this function will return TRUE
Public Function readyToTranslate() As Boolean
    readyToTranslate = m_isReady
End Function

'If the currently active language is different from the default language, this function will return TRUE
Public Function translationActive() As Boolean

    If curLanguage = defaultLanguage Then
        translationActive = False
    Else
        translationActive = True
    End If

End Function

'Given an index in the availableLanguages array, mark that as the new language (including saving it to the preferences file)
Public Sub activateNewLanguage(ByVal newLangIndex As Long)

    curLanguage = newLangIndex
    g_UserPreferences.SetPref_String "General Preferences", "CurrentLanguageFile", availableLanguages(newLangIndex).FileName
    ApplyLanguage

End Sub

'Determine which language to use.  This can be done one of several ways:
' 1) If the user has already specified a language, use that one (obviously)
' 2) If the user has not specified a language, and this is not the first time they have run the program, default to English.
' 3) If this is the first time PhotoDemon has been run on this machine, estimate what language to present based on the
'     current system language.  If a matching language is available, present that.  If not, present English.
'     (In the future, the user will be presented with an actual "choose your language" dialog if the system language
'      is NOT set to English.)
Public Sub DetermineLanguage()

    'FOR TESTING PURPOSES ONLY: uncomment the line below to trick the program into thinking it is being run for the first time.
    ' This will also initiate the language engine's analysis of the current system language, and it will do its best to match
    ' the system language to the best available language in PhotoDemon's language folder.
    ' This process only happens once, at first run - after that the user must initiate any language changes.
    'g_IsFirstRun = True

    'By default, do not display a language selection dialog
    m_needToShowLanguageDialog = False

    'Is this the first time PhotoDemon has been run?  If it is, check the system language - we'll use that to determine
    ' which available PhotoDemon language file is the best match for this user.
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
            m_reasonForLanguageDialog = LANGREASON_APPROXIMATEMATCH
            m_needToShowLanguageDialog = True
        End If
        
        'If we still can't find a match, present the language selection screen in English.
        If curLanguage = -1 Then
            curLanguage = defaultLanguage
            m_reasonForLanguageDialog = LANGREASON_NOMATCH
            m_needToShowLanguageDialog = True
        End If
        
    'If this is not the first run, retrieve the user's language preference from the preferences file
    Else
    
        Dim curLanguageFile As String
        curLanguageFile = g_UserPreferences.GetPref_String("General Preferences", "CurrentLanguageFile", "")
        
        'If no specific file has been named, use the default language
        If curLanguageFile = "" Then
            curLanguage = defaultLanguage
        
        'If a file HAS been named, make sure it still exists; if it doesn't, default to English (US) and show the language dialog
        Else
            If FileExist(g_UserPreferences.getLanguagePath & curLanguageFile) Then
                curLanguage = getLangIndexFromFile(curLanguageFile)
            Else
                curLanguage = defaultLanguage
                m_reasonForLanguageDialog = LANGREASON_FILEMISSING
                m_needToShowLanguageDialog = True
            End If
        End If
    
    End If
    
    m_isReady = True

End Sub

'Given a language filename, return the index in the availableLanguages() array
Private Function getLangIndexFromFile(ByVal srcFile As String) As Long
    
    Dim i As Long
    For i = 0 To numOfLanguages - 1
        If StrComp(availableLanguages(i).FileName, srcFile, vbTextCompare) = 0 Then
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
            'pdMsgBox "Approximate language match found: " & availableLanguages(i).langID
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
            'pdMsgBox "Exact language match found: " & availableLanguages(i).langID
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, no match was found.  Return -1
    isExactLanguageMatch = -1

End Function

'ALWAYS CALL DetermineLanguage BEFORE THIS STAGE!
'Once a PhotoDemon language has been selected (via DetermineLanguage), this function can be called to "apply" that language to the program.
' NOTE: if the language is changed, this function must be called again to set the new language program-wide.
Public Sub ApplyLanguage()

    'Mark the active language menu entry (and unmark all others)
    Dim i As Long
    For i = 0 To numOfLanguages - 1
        If curLanguage = i Then FormMain.mnuLanguages(i).Checked = True Else FormMain.mnuLanguages(i).Checked = False
    Next i
    
    'If the language is set to anything other than the default, load the contents of the language XML file into memory.
    If translationActive Then
        m_translationFileContents = ""
        m_translationFileContents = getFileAsString(g_UserPreferences.getLanguagePath & availableLanguages(curLanguage).FileName)
    End If

    'Redraw all text on the main form using the new language
    If FormMain.Visible Then
        FormMain.requestMakeFormPretty
        DrawMenuShortcuts
        ApplyAllMenuIcons
        
        'Because the Undo and Redo captions are dynamic, they must be recreated in the current language
        tInit tUndo, FormMain.MnuUndo.Enabled
        tInit tRedo, FormMain.MnuRedo.Enabled
        
        'Normally we would want to reset menu icons manually, but the tInit calls above call ResetMenuIcons for us.
        'ResetMenuIcons
        
        'If an image is loaded, we need to redraw the "size" caption on the left
        If (NumOfWindows > 0) Then
            If Not pdImages(CurrentImage) Is Nothing Then DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
        End If
        
    End If
        
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

'Add all the language files in a given folder.  The function will return the number of languages found.
' Optionally, a language offset count can be provided.  This is required for any folders besides the first, as the
' function will use that as the array index where it can place found language file data.
Private Function tallyLanguageFilesInFolder(ByVal srcFolder As String, Optional ByVal langOffset As Long = 0) As Long

    Dim nLanguages As Long
    nLanguages = 0
    
    'Scan the translation folder for .xml files.  Ignore anything that isn't XML.
    Dim chkFile As String, curFileAsText As String
    chkFile = Dir(srcFolder & "*.xml", vbNormal)
        
    Do While (chkFile <> "")
                
        'Load the first 1024 bytes of the file into a string (this is done to keep this routine fast)
        curFileAsText = getPartialFileAsString(srcFolder & chkFile)
        
        'If the file has actual contents, investigate further.
        If Len(curFileAsText) > 0 Then
        
            'Is the .xml file a valid xml file?  Check the header to make sure.
            If verifyXMLHeader(curFileAsText) Then
            
                'If this is a valid language file, get the language information (if available).
                fillLanguageInfo curFileAsText, availableLanguages(nLanguages + langOffset)
                availableLanguages(nLanguages + langOffset).FileName = chkFile
                nLanguages = nLanguages + 1
                
                'If, for some ungodly reason, the user has more than 100 languages available, dynamically resize the array to fit.
                If nLanguages + langOffset > UBound(availableLanguages) Then ReDim Preserve availableLanguages(0 To (nLanguages + langOffset) * 2) As pdLanguageFile
            
            End If
            
        End If
        
        'Retrieve the next file and repeat
        chkFile = Dir
    
    Loop
    
    tallyLanguageFilesInFolder = nLanguages

End Function

'Search the Languages folder, and make a list of all available languages
Public Sub CheckAvailableLanguages()
    
    numOfLanguages = 0
    ReDim Preserve availableLanguages(0 To 99) As pdLanguageFile
    
    'Find all available OFFICIAL language files
    numOfLanguages = tallyLanguageFilesInFolder(g_UserPreferences.getLanguagePath, 0)
    
    'Find all available USER language files
    numOfLanguages = numOfLanguages + tallyLanguageFilesInFolder(g_UserPreferences.getLanguagePath(True), numOfLanguages)
    
    'Add a dummy entry for PhotoDemon's default language (en-US)
    availableLanguages(numOfLanguages).FileName = ""
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
    
    'Mark the location in the array of the default language, as other functions may need to know it
    Dim i As Long
    For i = 0 To numOfLanguages - 1
        If StrComp(availableLanguages(i).langName, "English (US) - Default", vbTextCompare) = 0 Then defaultLanguage = i
    Next i
    
    'TESTING: display the info we retrieved
    'Dim i As Long
    'For i = 0 To numOfLanguages - 1
    '    pdMsgBox availableLanguages(i).langID & vbCrLf & availableLanguages(i).langName & vbCrLf & availableLanguages(i).langStatus & vbCrLf & availableLanguages(i).langVersion
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
    dstLang.FileName = srcLang.FileName
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
    
    Dim tmpFileContents As String
    
    'Ensure that the file exists before attempting to load it
    If FileExist(fName) Then
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        Open fName For Binary As #fileNum
            tmpFileContents = Space$(LOF(fileNum))
            Get #fileNum, , tmpFileContents
        Close #fileNum
    
        'Remove all tabs from the source file (which may have been added in by an XML editor, but are not relevant to the translation process)
        If InStr(1, tmpFileContents, vbTab) <> 0 Then tmpFileContents = Replace(tmpFileContents, vbTab, "")
        
        getFileAsString = tmpFileContents
    
    Else
        getFileAsString = ""
    End If
    
End Function

'All text to be translated should be passed here first.  This sub will perform a few automatic checks of the string, and it will record what
' it finds.  For example, if the text has "..." at the end, this sub will remove that.  Its counterpart - postProcessText() - will re-apply
' any necessary changes to the translation.
Private Sub preProcessText(ByRef srcString As String)

    '1) Trim the string
    srcString = Trim$(srcString)
    
    '2) Check for a trailing "..."
    If Right$(srcString, 3) = "..." Then
        m_hasDotDotDot = True
        srcString = Left$(srcString, Len(srcString) - 3)
    Else
        m_hasDotDotDot = False
    End If
    
    '3) Check for a trailing colon ":"
    If Right$(srcString, 1) = ":" Then
        m_hasTrailingColon = True
        srcString = Left$(srcString, Len(srcString) - 1)
    Else
        m_hasTrailingColon = False
    End If
    
End Sub

Private Sub postProcessText(ByRef srcString As String)

    '1) Trim the string
    srcString = Trim$(srcString)
    
    '2) If the string is now empty, return a blank string rather than attempting to reconstruct it
    If srcString <> "" Then
    
        '2) If the original had a trailing "...", restore it now
        If m_hasDotDotDot Then
            srcString = srcString & "..."
            m_hasDotDotDot = False
        End If
        
        '3) If the original had a trailing colon, restore it now
        If m_hasTrailingColon Then
            srcString = srcString & ":"
            m_hasTrailingColon = False
        End If
    End If

End Sub

'Given the translated caption of a message or control, return the original translation from the active translation file
Private Function getOriginalTagFromTranslation(ByVal curCaption As String) As String

    'Remove white space from the caption (if necessary, white space will be added back in after retrieving the translation from file)
    preProcessText curCaption
    curCaption = "<translation>" & curCaption & "</translation>"
    
    Dim phraseLocation As Long
    phraseLocation = getPhraseTagLocation(curCaption)
    
    'Make sure a phrase tag was found
    If phraseLocation > 0 Then
        
        'Retrieve the <translation> tag inside this phrase tag
        curCaption = getTextBetweenTags(m_translationFileContents, "original", phraseLocation)
        postProcessText curCaption
        getOriginalTagFromTranslation = curCaption
        
    Else
        getOriginalTagFromTranslation = ""
    End If

End Function

'Given the original caption of a message or control, return the matching translation from the active translation file
Private Function getTranslationTagFromCaption(ByVal origCaption As String) As String

    'Remove white space from the caption (if necessary, white space will be added back in after retrieving the translation from file)
    preProcessText origCaption
    origCaption = "<original>" & origCaption & "</original>"
    
    Dim phraseLocation As Long
    phraseLocation = getPhraseTagLocation(origCaption)
    
    'Make sure a phrase tag was found
    If phraseLocation > 0 Then
        
        'Retrieve the <translation> tag inside this phrase tag
        origCaption = getTextBetweenTags(m_translationFileContents, "translation", phraseLocation)
        postProcessText origCaption
        If origCaption <> "" Then
            getTranslationTagFromCaption = origCaption
        Else
            getTranslationTagFromCaption = ""
        End If
        
    Else
        getTranslationTagFromCaption = ""
    End If

End Function

'Given a string, return the location of the <phrase> tag enclosing said string
Private Function getPhraseTagLocation(ByRef srcString As String) As Long

    Dim sLocation As Long
    
    'If the user has activated shortcut searching, start from the position previously discovered
    If m_shortcutActive Then
        sLocation = InStr(m_shortcutLocation, m_translationFileContents, srcString, vbTextCompare)
        If sLocation = 0 Then sLocation = InStrRev(m_translationFileContents, srcString, m_shortcutLocation, vbTextCompare)
    'If shortcuts are inactive, or if they're active but the previous search failed, search starting from the start of the file
    Else
        sLocation = InStr(1, m_translationFileContents, srcString, vbTextCompare)
    End If
    
    'If the source string was found, work backward to find the phrase tag location
    If sLocation > 0 Then
        sLocation = InStrRev(m_translationFileContents, "<phrase>", sLocation, vbTextCompare)
        If sLocation > 0 Then
            getPhraseTagLocation = sLocation
        Else
            getPhraseTagLocation = 0
        End If
    Else
        getPhraseTagLocation = 0
    End If

End Function

'Given a message (basically, a string), return the translation of said string.
' If a translation cannot be found, return the original string.
Public Function TranslateMessage(ByRef srcMessage As String, ParamArray ExtraText() As Variant) As String
    
    Dim searchString As String, newString As String
    If translationActive Then
        searchString = srcMessage
        newString = getTranslationTagFromCaption(searchString)
    Else
        newString = srcMessage
    End If
    
    'If no translation was found, return the original message as the "translation".  This allows the program to still be usable
    ' even if translation text is missing.
    If newString = "" Then
        
        'If this string contains vbCrLf characters, replace them with plain vbLF characters and try again
        If translationActive And (InStr(1, searchString, vbCrLf) > 0) Then
            newString = getTranslationTagFromCaption(Replace$(searchString, vbCrLf, vbLf))
            If newString = "" Then newString = srcMessage
        Else
            newString = srcMessage
        End If
    End If
    
    'newString now contains the translated message (or the original text, if no translation was found).
    
    'The last step is to replace any optional parameters in the text
    If Not IsMissing(ExtraText) Then
    
        Dim i As Long
        For i = LBound(ExtraText) To UBound(ExtraText)
            newString = Replace$(newString, "%" & i + 1, ExtraText(i))
        Next i
    
    End If
    
    'Return the fully translated string
    TranslateMessage = newString
    
End Function

'Given a message (basically, a string), return the translation of said string.
' If a translation cannot be found, return the original string.
Public Function RestoreMessage(ByRef srcMessage As String) As String
    
    Dim searchString As String, newString As String
    searchString = srcMessage
    newString = getOriginalTagFromTranslation(searchString)
    
    'If a new string was found, apply it to the object's caption
    If newString <> "" Then RestoreMessage = newString Else RestoreMessage = srcMessage
    
End Function

'Given an object with a guaranteed .Caption property (command button, label, etc.), replace the caption with a translation from file
Public Sub translateObjectCaption(ByRef srcObject As Control)

    Dim searchString As String, newString As String
    
    If translationActive Then
        searchString = srcObject.Caption
        newString = getTranslationTagFromCaption(searchString)
    Else
        newString = srcObject.Caption
    End If
    
    'If a new string was found, apply it to the object's caption
    If newString <> "" Then
        
        'Check boxes, option buttons, and forms require a space at the front of their captions (it looks better)
        If (TypeOf srcObject Is CheckBox) Or (TypeOf srcObject Is OptionButton) Or (TypeOf srcObject Is Form) Then
            srcObject.Caption = " " & newString
        Else
            srcObject.Caption = newString
        End If
    End If

End Sub

'Given an object with a guaranteed .Caption property (command button, label, etc.), replace a translated caption with the original
Public Sub restoreObjectCaption(ByRef srcObject As Control)

    Dim searchString As String, newString As String
    searchString = srcObject.Caption
    newString = getOriginalTagFromTranslation(searchString)
    
    'If a new string was found, apply it to the object's caption
    If newString <> "" Then
        
        'Check boxes, option buttons, and forms require a space at the front of their captions (it looks better)
        If (TypeOf srcObject Is CheckBox) Or (TypeOf srcObject Is OptionButton) Or (TypeOf srcObject Is Form) Then
            srcObject.Caption = " " & newString
        Else
            srcObject.Caption = newString
        End If
    End If

End Sub

'Restore a form caption
Public Sub restoreFormCaption(ByRef srcForm As Form)

    Dim searchString As String, newString As String
    searchString = srcForm.Caption
    newString = getOriginalTagFromTranslation(searchString)
    
    'If a new string was found, apply it to the object's caption
    If newString <> "" Then srcForm.Caption = " " & newString
    
End Sub

'Translate a form caption
Public Sub translateFormCaption(ByRef srcForm As Form)

    Dim searchString As String, newString As String
    
    If translationActive Then
        searchString = srcForm.Caption
        newString = getTranslationTagFromCaption(searchString)
    Else
        newString = srcForm.Caption
    End If
    
    'If a new string was found, apply it to the object's caption
    If newString <> "" Then srcForm.Caption = " " & newString
    
End Sub

'Enumerate all translations on a form (and its controls), and replace them with their original English equivalents
Public Sub undoTranslations(ByRef srcForm As Form)

    'Start by returning the form's original caption
    restoreFormCaption srcForm
    
    Dim i As Long, strPos As Long
    Dim tmpString As String
    
    'Iterate through every control on the form, and restore original captions as necessary
    Dim eControl As Control
    For Each eControl In srcForm.Controls
    
        'If this is an object with a standard .Caption property, like a label or command button, replacing the text is easy.
        If (TypeOf eControl Is CommandButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is Label) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is jcbutton) Then
            If (Trim(eControl.Caption) <> "") And (Not IsNumeric(eControl.Caption)) Then g_Language.restoreObjectCaption eControl
            
            If (Trim(eControl.ToolTipText) <> "") Then eControl.ToolTipText = RestoreMessage(eControl.ToolTipText)
                        
            'jcbutton controls have multiple tooltip entries, including a tooltip title
            If (TypeOf eControl Is jcbutton) Then
                If (Trim(eControl.TooltipTitle) <> "") Then eControl.TooltipTitle = RestoreMessage(eControl.TooltipTitle)
                If (Trim(eControl.ToolTip) <> "") Then eControl.ToolTip = RestoreMessage(eControl.ToolTip)
            End If
            
        End If
        
        'Menus are a special case.  Some PhotoDemon menus have custom accelerators (keyboard shortcuts) that are added at run-time.
        ' We can detect these by checking for a vbTab character in the menu entry.  If this is found, strip it out before searching
        ' for the original text.
        If (TypeOf eControl Is Menu) Then
            tmpString = eControl.Caption
            strPos = InStr(1, tmpString, vbTab, vbBinaryCompare)
            
            'If this menu caption has a custom accelerator appended, strip it off
            If strPos > 0 Then tmpString = Left(tmpString, strPos - 1)
            eControl.Caption = RestoreMessage(tmpString)
            
        End If
        
        'If this is a list box or combo box, replacing the text is a bit more confusing
        If (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Then
            
            'Enumerate through each entry in the control
            For i = 0 To eControl.ListCount - 1
                If (Trim(eControl.List(i)) <> "") And (Not IsNumeric(eControl.List(i))) Then eControl.List(i) = RestoreMessage(eControl.List(i))
            Next i
            
            'Restore the tooltip as well
            If (Trim(eControl.ToolTipText) <> "") Then eControl.ToolTipText = RestoreMessage(eControl.ToolTipText)
            
        End If
    
    Next

End Sub

'Enumerate all captions on a form (and its controls), and replace them with the translated equivalents of the current language
Public Sub applyTranslations(ByRef srcForm As Form)

    If Not translationActive Then Exit Sub

    'Dim localTimeTaken As Long
    'localTimeTaken = Timer

    'Start by translating the form's caption
    translateFormCaption srcForm

    'Iterate through every control on the form, and restore original captions as necessary
    Dim i As Long, strPos As Long
    Dim tmpString As String
    Dim eControl As Control
    
    For Each eControl In srcForm.Controls
    
        'The goal is to apply translations to any object with a non-blank, non-numeric caption property
        
        'If this is an object with a standard .Caption property, like a label or command button, replacing the text is easy.
        If (TypeOf eControl Is CommandButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is Label) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is jcbutton) Or (TypeOf eControl Is smartOptionButton) Or (TypeOf eControl Is smartCheckBox) Then
            
            'TESTING ONLY
            'If srcForm.Name <> "FormMain" Then MsgBox eControl.Caption
            
            If (Trim(eControl.Caption) <> "") And (Not IsNumeric(eControl.Caption)) Then translateObjectCaption eControl
            If (Trim(eControl.ToolTipText) <> "") Then eControl.ToolTipText = TranslateMessage(eControl.ToolTipText)
                        
            'jcbutton controls have multiple tooltip entries, including a tooltip title
            If (TypeOf eControl Is jcbutton) Then
                If (Trim(eControl.TooltipTitle) <> "") Then eControl.TooltipTitle = TranslateMessage(eControl.TooltipTitle)
                If (Trim(eControl.ToolTip) <> "") Then eControl.ToolTip = TranslateMessage(eControl.ToolTip)
            End If
            
        End If
        
        'Picture boxes and text boxes can have tooltips.  Translate those too.
        If (TypeOf eControl Is PictureBox) Or (TypeOf eControl Is TextBox) Then
            If (Trim(eControl.ToolTipText) <> "") Then eControl.ToolTipText = TranslateMessage(eControl.ToolTipText)
        End If
        
        'Menus are a special case.  Some PhotoDemon menus have custom accelerators (keyboard shortcuts) that are added at run-time.
        ' We can detect these by checking for a vbTab character in the menu entry.  If this is found, strip it out before searching
        ' for a translation.
        If (TypeOf eControl Is Menu) Then
            tmpString = eControl.Caption
            strPos = InStr(1, tmpString, vbTab, vbBinaryCompare)
            
            'If this menu caption has a custom accelerator appended, strip it off
            If strPos > 0 Then tmpString = Left(tmpString, strPos - 1)
                
            eControl.Caption = TranslateMessage(tmpString)
            
        End If
            
        'If this is a list box or combo box, replacing the text is a bit more confusing
        If (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Then
            
            'Enumerate through each entry in the control
            For i = 0 To eControl.ListCount - 1
                If (Trim(eControl.List(i)) <> "") And (Not IsNumeric(eControl.List(i))) Then eControl.List(i) = TranslateMessage(eControl.List(i))
            Next i
            
            'Translate the tooltip as well
            If (Trim(eControl.ToolTipText) <> "") Then eControl.ToolTipText = TranslateMessage(eControl.ToolTipText)
            
        End If
        
        Dim curFontSize As Long
        
        'Labels receive a bit of extra attention.  Because translations can be longer than their original English text, a translated
        ' label may extend beyond the edge of its container.  If this happens, attempt to shrink the label's font size until it fits.
        If (TypeOf eControl Is Label) Then
            If (eControl.Left + eControl.Width > eControl.Container.ScaleWidth) And eControl.AutoSize Then
            
                'Try shrinking the font size until an acceptable width is found
                curFontSize = eControl.FontSize
                Do While ((eControl.Left + eControl.Width) > eControl.Container.ScaleWidth) And (curFontSize >= 8)
            
                    curFontSize = curFontSize - 1
                    eControl.FontSize = curFontSize
                    eControl.Refresh
            
                Loop
            
            End If
        End If
        
        'PhotoDemon's "smart" option and check box buttons receive the same autosize treatment
        If (TypeOf eControl Is smartOptionButton) Or (TypeOf eControl Is smartCheckBox) Then
            If (eControl.Left + eControl.Width > eControl.Container.ScaleWidth) Then
            
                'Try shrinking the font size until an acceptable width is found
                curFontSize = eControl.Font.Size
                Do While ((eControl.Left + eControl.Width) > eControl.Container.ScaleWidth) And (curFontSize >= 8)
                
                    curFontSize = curFontSize - 1
                    eControl.Font.Size = curFontSize
                
                Loop
            
            End If
        End If
        
    Next
    
    'If DISPLAY_TIMINGS Then MsgBox "Translation of " & srcForm.Name & " took " & (Timer - localTimeTaken) & " seconds."
    
End Sub

Private Sub Class_Initialize()
    
    'By default, shortcut searching is inactive
    m_shortcutActive = False
    m_shortcutLocation = 1
    
End Sub
