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
'Copyright 2012-2026 by Frank Donckers and Tanner Helland
'Created: 21/December/12
'Last updated: 28/September/23
'Last update: new "assign translation by object ID instead of caption" feature (see https://github.com/tannerhelland/PhotoDemon/issues/491)
'
'In 2012, Frank Donckers made a significant contribution to PhotoDemon: a prototype for a run-time
' localization engine.  This may not sound like much, but VB6 does not provide native support for
' anything like this, so this was a *huge* undertaking, especially for a volunteer working on a
' foreign codebase.  Thank you to Frank for kicking off what would become a multi-year initiative
' to provide the best multi-language experience ever achieved in VB6.  (Frank also provided the
' initial language files for three new localizations: Vlaams, Francais, and Deutsch.)
'
'Today, PhotoDemon's translation engine is a lean-and-mean little machine.  It provides full
' support for Unicode text both on the backend (e.g. Unicode filenames) and on the front-end
' (UI elements, user input).  This is achieved via a fully themable, UTF-8 compatible UI
' suite hand-built for the project. New localizations can be applied (or undone) *at* run-time and
' *in* real-time, which I consider an indispensable feature, especially because translations are
' supplied by volunteers and may not always be 100% up-to-date.
'
'Translation files are shipped as basic XML.  You can see (and freely edit) these files inside the
' /App/PhotoDemon/Languages/ subfolder wherever PD is installed.  If you do decide to edit PD's
' language files, I strongly suggest saving the new files inside the USER folder of the project,
' e.g. /Data/Languages/.  This will ensure they are not overwritten by software updates or future
' changes to the official language files that ship with the app.
'
'Most of the crucial translation work in the program happens inside a global object called
' "g_Language", which is an instance of this class (pdTranslate).  Most subs and functions here
' should be self-explanatory.  To really get a feel for how the translation engine works, I suggest
' starting with the Loading module, and looking at the way this class is initialized and loaded.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'NOTE: this class frequently uses a type called PDLanguageFile.  Because outside functions
' (like the Language Editor) also access this type, it is located in Public_Variables.

'This class offers detailed performance tracking, but it pollutes debug logs so please DISABLE in production builds
Private Const DETAILED_PERF_TRACKING As Boolean = False

'Language-related WAPI
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long

Private Const LOCALE_USER_DEFAULT As Long = &H400
Private Const LOCALE_SISO639LANGNAME As Long = &H59   'ISO 639 abbreviated language name, e.g. "en" for "English"
Private Const LOCALE_SISO3166CTRYNAME As Long = &H5A   'ISO 4166 abbreviated country/region name, e.g. "US" for "United States"

'Array of currently available language files.
Private m_numOfLanguages As Long
Private m_AvailableLanguages() As PDLanguageFile

'Array ID of the current language.
' If no language is selected (e.g. PhotoDemon is using en-US), this value is -1.
Private m_currentLanguage As Long

'Array ID of the default language (en-US).
Private m_defaultLanguage As Long

'If the user should see a language select screen, this will be set to TRUE during initialization
Private m_needToShowLanguageDialog As Boolean

'The full contents of the translation file are loaded into this string.  Queries against this string
' will always return the correct translation for a given en-US equivalent, but additional mechanisms
' may be used to avoid touching this original XML string (unless absolutely necessary).
Private m_XMLEngine As pdXML

'Is the class ready for use?  Early in the app load process, this may be FALSE.  We use this value
' internally to determine whether we can actually return localized strings.
Private m_isReady As Boolean

'The following boolean values are marked by PreProcessString(), then read by PostProcessString().
' (Things like trailing colons or ellipses are stripped before matching translations.)
Private m_hasEllipsis As Boolean
Private m_hasTrailingColon As Boolean

'Language files are checksummed to improve startup time.  This checksum class is not guaranteed
' to exist, so please review individual functions before attempting to use the crypto object.
' (Functions that require it should always check instantation before acccessing.)
Private m_Crypto As pdCrypto

'The language file cache gets stored as its own persistent XML file.  If we detect changes,
' we'll note that cache change state separately.
Private m_LangCacheXML As pdXML, m_LangCacheChanged As Boolean

'Performance is tracked by developer builds
Private m_NetTranslateTime As Double

'When a translation is successfully retrieved, we cache it in a fast string collection.
' Subsequent requests can query the table for greatly improved performance.
Private m_FoundTranslations As pdStringHash

'Some languages need translations that do not map 1:1 with English text.
' (For example, `Color > Invert` vs `Select > Invert` requires different translations in Chinese for `Invert`,
'  per https://github.com/tannerhelland/PhotoDemon/issues/491.)
' Not all languages require this cache, so it may contain 0 entries at run-time - which is fine!
Private m_ObjectTranslations As pdStringHash

'If a special object ID cache exists, this will be set to TRUE (otherwise FALSE, which allows us to skip
' additional queries into this table when looking up translations, for objects that pass object IDs)
Private m_ObjectTranslationsExist As Boolean

'In debug mode we track extra profiling data
Private m_CacheHits As Long, m_CacheMisses As Long

'Used for querying file attributes (at startup)
Private m_FileAttributesBasic As WIN32_FILE_ATTRIBUTES_BASIC

'External functions can request the a copy of all found languages.  The caller must provide
' a pdLanguageFile array; this function will fill it (with data retrieved by CheckAvailableLanguages()).
Friend Sub CopyListOfLanguages(ByRef dstListOfLanguages() As PDLanguageFile)
    
    ReDim dstListOfLanguages(0 To m_numOfLanguages - 1) As PDLanguageFile
    
    Dim i As Long
    For i = 0 To m_numOfLanguages - 1
        dstListOfLanguages(i) = m_AvailableLanguages(i)
    Next i
    
End Sub

'External functions can request the current language, which may then be used to request translations
' from 3rd-party plugins (ExifTool, at present).
Friend Function GetCurrentLanguage(Optional ByVal getLocaleToo As Boolean = False) As String
    If Me.TranslationActive Then
        If getLocaleToo Then
            GetCurrentLanguage = m_AvailableLanguages(m_currentLanguage).langID
        Else
            GetCurrentLanguage = Left$(m_AvailableLanguages(m_currentLanguage).langID, 2)
        End If
    Else
        If getLocaleToo Then GetCurrentLanguage = "en-US" Else GetCurrentLanguage = "en"
    End If
End Function

'This function will return PD's internal language index (m_currentLanguage).  This also serves as
' an index into the Language menu.
Friend Function GetCurrentLanguageIndex() As Long
    GetCurrentLanguageIndex = m_currentLanguage
End Function

'External function for asking the translation engine if a language selection dialog is needed
Friend Function IsLanguageDialogNeeded() As Boolean
    IsLanguageDialogNeeded = m_needToShowLanguageDialog
End Function

'If a translation file has been requested and successfully loaded, this function will return TRUE
Friend Function ReadyToTranslate() As Boolean
    If (Not m_XMLEngine Is Nothing) Then ReadyToTranslate = m_isReady Else ReadyToTranslate = False
End Function

'If the currently active language is different from the default language, this function will return TRUE.
' If this function returns TRUE, you do not need to query the translation engine at all - just use the
' default en-US text.
Friend Function TranslationActive() As Boolean
    TranslationActive = (m_currentLanguage <> m_defaultLanguage)
End Function

'Store the user's language preference
Friend Sub WriteLanguagePreferencesToFile()
    UserPrefs.SetPref_String "Language", "Current Language File", m_AvailableLanguages(m_currentLanguage).FileName
End Sub

'Given an index in the m_availableLanguages array, mark that as the new language
' (including saving it to the preferences file)
Friend Sub ActivateNewLanguage(ByVal newLangIndex As Long)
    m_currentLanguage = newLangIndex
    Interface.GenerateInterfaceID
    WriteLanguagePreferencesToFile
End Sub

'Determine which language to use.  This can be done one of several ways:
' 1) If the user has already specified a language, use it (obviously)
' 2) If the user has *not* specified a language, and this is *not* the first time they've run the program,
'     assume they want English.
' 3) If this is the user's first time running PhotoDemon, estimate what language to present based on the
'     current system language.  If a matching language is available, present that.  (If not, present English.)
Friend Sub DetermineLanguage()

    'FOR TESTING PURPOSES ONLY:
    ' uncomment the line below to trick the program into thinking it is being run for the first time.
    ' This will also initiate the language engine's analysis of the current system language,
    ' and it will do its best to match the system language to the best available language in PD's collection.
    ' This process only happens once, at first run - after that the user must manually initiate language changes.
    'g_IsFirstRun = True

    'By default, do not display a language selection dialog
    m_needToShowLanguageDialog = False

    'Is this the first time PhotoDemon has been run?  If it is, check the system language - we'll use that
    ' to determine which available PhotoDemon language file is the best match for this user.
    If g_IsFirstRun Then
    
        Dim curLangID As String
        curLangID = GetDefaultUserLanguage()        'Want to default to PD's stock translation?  Set this to "en-US"
        PDDebug.LogAction "Auto-detected preferred locale as " & curLangID
        
        'Compare the full language ID (language and region) against language files available in the Languages folder.
        ' If an exact match is found, present that as the default program language.
        m_currentLanguage = IsExactLanguageMatch(curLangID)
        
        'If an exact match is not found, try comparing just the language part of the code.  If a match is found,
        ' present that as the default language - and note that we also need to show the language selection dialog.
        If (m_currentLanguage < 0) Then
            m_currentLanguage = IsApproximateLanguageMatch(curLangID)
            m_needToShowLanguageDialog = True
        End If
        
        'If we still can't find a match, present the language selection screen in English.
        If (m_currentLanguage < 0) Then
            m_currentLanguage = m_defaultLanguage
            m_needToShowLanguageDialog = True
        End If
        
        PDDebug.LogAction "First-run language decision is " & m_AvailableLanguages(m_currentLanguage).langID
        If m_needToShowLanguageDialog Then PDDebug.LogAction "Language dialog will be displayed at end-of-load."
        
    'If this is not the first run, retrieve the user's language preference from the preferences file
    Else
    
        Dim curLangFile As String
        curLangFile = UserPrefs.GetPref_String("Language", "Current Language File", vbNullString)
        
        'If no specific file has been named, use the default language
        If (LenB(curLangFile) = 0) Then
            m_currentLanguage = m_defaultLanguage
        
        'If a file HAS been named, make sure it still exists; if it doesn't, default to English (US) and show the language dialog
        Else
            If Files.FileExists(curLangFile) Then
                m_currentLanguage = GetLangIndexFromFile(curLangFile)
            Else
                m_currentLanguage = m_defaultLanguage
                m_needToShowLanguageDialog = True
            End If
        End If
    
    End If
    
    m_isReady = True
    
End Sub

'Given a language filename, return the index in the m_availableLanguages() array
Private Function GetLangIndexFromFile(ByRef srcFile As String) As Long
    
    Dim i As Long
    For i = 0 To m_numOfLanguages - 1
        If Strings.StringsEqual(m_AvailableLanguages(i).FileName, srcFile, True) Then
            GetLangIndexFromFile = i
            Exit Function
        End If
    Next i
    
End Function

'Given a language ID (as a string), try to find an approximate match (just language, region doesn't matter) in the
' current language collection.
Private Function IsApproximateLanguageMatch(ByRef srcLangID As String) As Long

    IsApproximateLanguageMatch = -1
    
    'Why search the language collection backward?  PD supplies multiple translations for some languages, and the second
    ' of the two (alphabetically speaking) tends to be the better one.
    Dim i As Long
    For i = m_numOfLanguages - 1 To 0 Step -1
        If Strings.StringsEqual(Left$(srcLangID, 2), Left$(m_AvailableLanguages(i).langID, 2), True) Then
            IsApproximateLanguageMatch = i
            Exit Function
        End If
    Next i
    
End Function

'Given a language ID (as a string), try to find an exact match (language AND region) in the current language collection.
Private Function IsExactLanguageMatch(ByRef srcLangID As String) As Long

    IsExactLanguageMatch = -1
    
    'Why search the language collection backward?  PD supplies multiple translations for some languages, and the second
    ' of the two (alphabetically speaking) tends to be the better one.
    Dim i As Long
    For i = m_numOfLanguages - 1 To 0 Step -1
        If Strings.StringsEqual(srcLangID, m_AvailableLanguages(i).langID, True) Then
            IsExactLanguageMatch = i
            Exit Function
        End If
    Next i
    
End Function

'ALWAYS CALL DetermineLanguage BEFORE THIS STAGE!
'Once a PhotoDemon language has been selected (via DetermineLanguage), this function can be called to "apply" that language to the program.
' NOTE: if the language is changed, this function must be called again to set the new language program-wide.
Friend Sub ApplyLanguage(Optional ByVal redrawMainUI As Boolean = True, Optional ByVal useDoEvents As Boolean = False)
    
    'If the language is set to anything other than the default, load the contents of the language XML file into memory.
    If Me.TranslationActive Then
    
        Set m_XMLEngine = New pdXML
        m_XMLEngine.LoadXMLFile m_AvailableLanguages(m_currentLanguage).FileName
        
        'PhotoDemon uses stict binary text comparisons when matching translations.  This provides a large
        ' performance boost, and is necessary to produce correct captioning when the same word is used in
        ' multiple places with different cases (e.g. titlecase in menus vs lowercase in other UI elements).
        m_XMLEngine.SetTextCompareMode vbBinaryCompare
        
        'Reset the fast translation cache whenever the language changes
        If (m_FoundTranslations Is Nothing) Then Set m_FoundTranslations = New pdStringHash Else m_FoundTranslations.Reset
        
        'Reset the custom per-object translation cache too.
        ' (This cache is used in special situations, where English text does not map 1:1 to translated text.
        '  It is used to look-up translations by object-ID instead of English caption/text.)
        If (m_ObjectTranslations Is Nothing) Then Set m_ObjectTranslations = New pdStringHash Else m_ObjectTranslations.Reset
        
        'Build the special object-ID translation table (if one exists in the source file)
        m_ObjectTranslationsExist = False
        
        Dim posSpecialObjectTable As Long
        If m_XMLEngine.DoesTagExist("special-translations", vbNullString, vbNullString, posSpecialObjectTable) Then BuildSpecialObjectTable posSpecialObjectTable
        
    End If
    
    'Mark the active language menu entry (and unmark all others).  This is necessary even if a translation is *not* active,
    ' to ensure that the "English (US)" entry is checked correctly.
    Dim i As Long
    For i = 0 To m_numOfLanguages - 1
        FormMain.MnuLanguages(i).Checked = (m_currentLanguage = i)
    Next i
    
    'If any objects cache translations, reset them now.
    UserControls.ResetCommonTranslations
    
    'The caller may choose to suspend the main interface redraw.  (We do this when loading the program, for example.)
    If redrawMainUI Then Interface.RedrawEntireUI useDoEvents
    
    'Some other esoteric areas may need their translations live-updated.  (Use this block only
    ' for objects that do not have a persistent UI - those should have been updated *before* the
    ' big .RedrawEntireUI call, above.)
    ImageFormats.NotifyLanguageChanged
    
End Sub

Friend Function GetSystemUserCountry() As String
    GetSystemUserCountry = UCase$(GetSpecificLocaleInfo(LOCALE_SISO3166CTRYNAME))
End Function

Friend Function GetSystemUserLanguage() As String
    GetSystemUserLanguage = LCase$(GetSpecificLocaleInfo(LOCALE_SISO639LANGNAME))
End Function

'Check the current default user language, and return it as a standard language code, e.g. "en" or "en-GB" or "de-CH"
Private Function GetDefaultUserLanguage() As String

    'First, retrieve the user's current language.
    Dim langReturn As Integer
    langReturn = GetUserDefaultLangID()
    
    'We now need to deconstruct that numerical, Microsoft-specific ID into a standard ISO locale ID.  Rather
    ' than build our own table, we can use Windows to do this for us.  The results may not always be perfect,
    ' but they should be "close enough" to estimate the best language to suggest.
    Dim abbrLang As String, abbrRegion As String
    
    'Get the ISO 639 abbreviated language name (e.g. "en" for "English")
    abbrLang = Me.GetSystemUserLanguage()
    
    'Get the ISO 4166 abbreviated country/region name (e.g. "US" for "United States")
    abbrRegion = Me.GetSystemUserCountry()
    
    'Return the language codes
    GetDefaultUserLanguage = abbrLang & "-" & abbrRegion

End Function

'Given a locale information constant, return a corresponding string
Private Function GetSpecificLocaleInfo(ByVal lInfo As Long) As String
    Dim sBuffer As String, sRet As Long
    sBuffer = String$(256, 0)
    sRet = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, StrPtr(sBuffer), Len(sBuffer))
    If (sRet > 0) Then
        GetSpecificLocaleInfo = Left$(sBuffer, sRet - 1)
    Else
        GetSpecificLocaleInfo = vbNullString
    End If
End Function

'Add all the language files in a given folder.  The function will return the number of languages found.
' Optionally, a language offset count can be provided.  This is required for any folders besides the first, as the
' function will use that as the array index where it can place found language file data.
Private Function TallyLanguageFilesInFolder(ByRef srcFolder As String, Optional ByVal langOffset As Long = 0, Optional ByVal langTypeForThisFolder As String = vbNullString) As Long

    Dim nLanguages As Long
    nLanguages = langOffset
    
    'If this is the official language folder, we will mark those entries specially, as they are the only ones PD is allowed to auto-update.
    Dim isOfficialFolder As Boolean
    isOfficialFolder = Strings.StringsEqual(langTypeForThisFolder, "OFFICIAL", True)
    
    'An XML class will be used to parse and validate language files, as needed.  (This class may not be instantiated
    ' if the current language file list matches the list at last-run, because language values are cached.  Plan accordingly.)
    Dim tmpXMLEngine As pdXML
    
    'Scan the translation folder for .xml files.  Ignore anything that isn't XML.
    Dim listOfFiles As pdStringStack
    Set listOfFiles = New pdStringStack
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    Dim netTime1 As Double, netTime2 As Double, netTime3 As Double
    
    If Files.RetrieveAllFiles(srcFolder, listOfFiles, False, False, "xml") Then
        
        If DETAILED_PERF_TRACKING Then
            PDDebug.LogTiming "pdTranslate / TallyLanguageFilesInFolder / RetrieveAllFiles", VBHacks.GetTimerDifferenceNow(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'We need to checksum various files, so prep a crypto object in advance
        Set m_Crypto = New pdCrypto
        
        Dim chkFile As String, chkCachedValue As String, chkXMLName As String
        Dim fileValidationRequired As Boolean
        
        Do While listOfFiles.PopString(chkFile)
            
            fileValidationRequired = True
            
            'We now have two options for loading this file.  If this is the first time we've seen this language file,
            ' we need to validate it (to make sure it's a usable PD language file).  But if we've seen this file before,
            ' we should have checksummed it and stored the result inside the user's preferences file.  If checksums
            ' match, we don't need to load and validate the file this time.
            
            VBHacks.GetHighResTime startTime
            
            'Look for an existing checksum
            chkXMLName = m_LangCacheXML.GetXMLSafeTagName(Files.FileGetName(chkFile, True))
            chkCachedValue = Trim$(m_LangCacheXML.GetUniqueTag_String(chkXMLName))
            If (LenB(chkCachedValue) <> 0) Then
            
                'A checksum was found.  If it matches the file's current checksum, we can skip further validation.
                fileValidationRequired = Strings.StringsNotEqual(chkCachedValue, GetLangFileChecksum(chkFile), False)
                If fileValidationRequired Then PDDebug.LogAction "Note: a language file checksum didn't match (" & chkXMLName & ").  Manual validation will follow."
                
            Else
                PDDebug.LogAction "Note: no cache found for " & chkXMLName & ".  Validating language file manually..."
            End If
            
            netTime1 = netTime1 + VBHacks.GetTimerDifferenceNow(startTime)
            
            'If validation is required, we need to load the file from scratch and run some checks on it
            If fileValidationRequired Then
                
                VBHacks.GetHighResTime startTime
                
                If (tmpXMLEngine Is Nothing) Then
                    Set tmpXMLEngine = New pdXML
                    tmpXMLEngine.SetTextCompareMode vbBinaryCompare
                End If
                
                'Use PD's XML engine to load the file
                If tmpXMLEngine.LoadXMLFile(chkFile) Then
                
                    'Use the XML engine to validate this file, and to make sure it contains at least a language ID, name, and one (or more) translated phrase
                    If tmpXMLEngine.IsPDDataType("Translation") Then
                        
                        'If this is a valid language file, get the language information (if available).
                        FillLanguageInfo tmpXMLEngine, m_AvailableLanguages(nLanguages)
                        m_AvailableLanguages(nLanguages).FileName = chkFile
                        m_AvailableLanguages(nLanguages).LangType = langTypeForThisFolder
                        m_AvailableLanguages(nLanguages).IsOfficial = isOfficialFolder
                        
                        'Because this is a valid language file, we want to cache its key properties in the
                        ' central user preferences file.  This lets us skip validation on subsequent runs.
                        m_LangCacheXML.UpdateTag chkXMLName, GetLangFileChecksum(chkFile)
                        CacheLanguageInfo chkXMLName, m_AvailableLanguages(nLanguages)
                        
                        'If, for some ungodly reason, the user has more than 100 languages available, dynamically resize the array to fit.
                        nLanguages = nLanguages + 1
                        If (nLanguages > UBound(m_AvailableLanguages)) Then ReDim Preserve m_AvailableLanguages(0 To (nLanguages + 1) * 2) As PDLanguageFile
                        
                    End If
                    
                End If
                
                netTime2 = netTime2 + VBHacks.GetTimerDifferenceNow(startTime)
                
            'If file validation is *not* required, pull previously cached values from the preference file
            Else
                
                VBHacks.GetHighResTime startTime
                
                GetCachedLanguageInfo chkXMLName, m_AvailableLanguages(nLanguages)
                m_AvailableLanguages(nLanguages).FileName = chkFile
                m_AvailableLanguages(nLanguages).LangType = langTypeForThisFolder
                m_AvailableLanguages(nLanguages).IsOfficial = isOfficialFolder
                
                nLanguages = nLanguages + 1
                If (nLanguages > UBound(m_AvailableLanguages)) Then ReDim Preserve m_AvailableLanguages(0 To (nLanguages + 1) * 2) As PDLanguageFile
                
                netTime3 = netTime3 + VBHacks.GetTimerDifferenceNow(startTime)
            
            End If
            
        'Retrieve the next file and repeat
        Loop
        
    End If
    
    If DETAILED_PERF_TRACKING Then
        PDDebug.LogTiming "pdTranslate / TallyLanguageFilesInFolder / IsValidationRequired", netTime1
        PDDebug.LogTiming "pdTranslate / TallyLanguageFilesInFolder / ActualValidation", netTime2
        PDDebug.LogTiming "pdTranslate / TallyLanguageFilesInFolder / PastValidation", netTime3
    End If
    
    TallyLanguageFilesInFolder = nLanguages - langOffset

End Function

Private Function GetLangFileChecksum(ByRef srcFile As String) As String
    
    'Pull file attributes just once
    If Files.FileGetAttributesBasic(srcFile, m_FileAttributesBasic) Then
        
        'Mix together the filename, file last access date, and file size to arrive at a "good enough" unique string.
        Dim strHashSource As String
        strHashSource = srcFile & "|" & CStr(m_FileAttributesBasic.nFileSizeBig) & "|" & CStr(m_FileAttributesBasic.ftLastWriteTime)
        
        'A dedicated crypto class handles the actual hashing
        GetLangFileChecksum = m_Crypto.QuickHashString(strHashSource)
        
    End If
        
End Function

'Search the Languages folder, and make a list of all available languages
Friend Sub CheckAvailableLanguages()
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
        
    m_numOfLanguages = 0
    ReDim m_AvailableLanguages(0 To 99) As PDLanguageFile
    
    'Language file names and attributes are cached inside a dedicated language settings file.  This file lets
    ' us bypass a lot of obnoxious and time-consuming validation steps.
    Dim langCacheFilename As String
    langCacheFilename = UserPrefs.GetPresetPath() & "LangCache.xml"
    
    Dim cacheLoadedSuccessfully As Boolean: cacheLoadedSuccessfully = False
    
    Set m_LangCacheXML = New pdXML
    m_LangCacheXML.SetTextCompareMode vbBinaryCompare
    
    If Files.FileExists(langCacheFilename) Then cacheLoadedSuccessfully = m_LangCacheXML.LoadXMLFile(langCacheFilename)
    
    If DETAILED_PERF_TRACKING Then
        PDDebug.LogTiming "pdTranslate / load cache file", VBHacks.GetTimerDifferenceNow(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'If the cache exists, we won't write it back out to file unless it somehow changes during this session
    ' (from a new language file being discovered, for example).  If it does *not* exist, we'll create it anew
    ' and force it to write out to file at the end of this initiation.
    If (Not cacheLoadedSuccessfully) Then
    
        PDDebug.LogAction "Central language cache wasn't found.  Creating new cache now..."
        m_LangCacheXML.PrepareNewXML "PDLanguageCache"
        m_LangCacheChanged = True
        
        If DETAILED_PERF_TRACKING Then
            PDDebug.LogTiming "pdTranslate / create new cache", VBHacks.GetTimerDifferenceNow(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
    End If
    
    'Before proceeding, perform some automatic language file cleanup.  In-place upgrades
    ' (e.g. between nightly builds or beta/release versions) may lead to duplicate language files with
    ' similar names; we always want to remove the old ones, *if* they're in the official language folder.
    PerformLanguageFileCleanup
    
    If DETAILED_PERF_TRACKING Then
        PDDebug.LogTiming "pdTranslate / lang file clean-up", VBHacks.GetTimerDifferenceNow(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Find all available OFFICIAL language files
    m_numOfLanguages = TallyLanguageFilesInFolder(UserPrefs.GetLanguagePath(False), 0, "Official")
    
    If DETAILED_PERF_TRACKING Then
        PDDebug.LogTiming "pdTranslate / scan official languages", VBHacks.GetTimerDifferenceNow(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Find all available USER language files
    m_numOfLanguages = m_numOfLanguages + TallyLanguageFilesInFolder(UserPrefs.GetLanguagePath(True), m_numOfLanguages, "Unofficial")
    
    If DETAILED_PERF_TRACKING Then
        PDDebug.LogTiming "pdTranslate / scan user languages", VBHacks.GetTimerDifferenceNow(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Add a dummy entry for PhotoDemon's default language (en-US)
    m_AvailableLanguages(m_numOfLanguages).FileName = vbNullString
    m_AvailableLanguages(m_numOfLanguages).langID = "en-US"
    m_AvailableLanguages(m_numOfLanguages).LangName = "English (US)"
    m_AvailableLanguages(m_numOfLanguages).LangType = "Default"
    m_AvailableLanguages(m_numOfLanguages).LangStatus = "Complete"
    m_AvailableLanguages(m_numOfLanguages).langVersion = VBHacks.AppMajor_Safe() & "." & VBHacks.AppMinor_Safe() & "." & VBHacks.AppRevision_Safe()
    m_AvailableLanguages(m_numOfLanguages).IsOfficial = True
    m_numOfLanguages = m_numOfLanguages + 1
    
    'Resize the array to contain only the number of languages found
    ReDim Preserve m_AvailableLanguages(0 To m_numOfLanguages - 1) As PDLanguageFile
    
    'Sort the array alphabetically by language name (because we are going to fill the languages submenu with the entries)
    SortLanguageData
    
    'Populate the languages menu
    PopulateLanguageMenu
    
    'Mark the location in the array of the default language, as other functions may need to know it
    Dim i As Long
    For i = 0 To m_numOfLanguages - 1
        If Strings.StringsEqual(m_AvailableLanguages(i).LangName, "english (us)", True) Then
            m_defaultLanguage = i
            Exit For
        End If
    Next i
    
    'If the cached language data was modified during this session, write the updated cache out to file
    If m_LangCacheChanged Then m_LangCacheXML.WriteXMLToFile langCacheFilename
    Set m_LangCacheXML = Nothing
    
    If DETAILED_PERF_TRACKING Then
        PDDebug.LogTiming "pdTranslate / everything else", VBHacks.GetTimerDifferenceNow(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'TESTING: display the info we retrieved
    'Dim i As Long
    'For i = 0 To m_numOfLanguages - 1
    '    Debug.Print m_availableLanguages(i).langID & vbCrLf & m_availableLanguages(i).langName & vbCrLf & m_availableLanguages(i).langStatus & vbCrLf & m_availableLanguages(i).langVersion
    'Next i
    
End Sub

'Some languages need access to translations that do not map 1:1 with English text.
' (For example, `Color > Invert` vs `Select > Invert` requires different translations in Chinese for `Invert`,
'  per https://github.com/tannerhelland/PhotoDemon/issues/491.)
'
'To enable this type of mapping, translators are allowed to add a custom <special-translations> block to
' a language file, which maps text to an *object-ID* instead of an English equivalent.
'
'This function builds that mapping (if one exists) directly into a hash table, which reduces the impact of
' this feature on run-time translation lookup performance.
Private Sub BuildSpecialObjectTable(Optional ByVal posSpecialBlock As Long = 1)
    
    'Ensure the object table exists and is empty
    If (m_ObjectTranslations Is Nothing) Then Set m_ObjectTranslations = New pdStringHash Else m_ObjectTranslations.Reset
    
    'If the caller didn't pass an XML block location (in chars), retrieve that position now.
    Const SPECIAL_BLOCK_NAME As String = "special-translations"
    If (posSpecialBlock = 1) Then
        If Not m_XMLEngine.DoesTagExist(SPECIAL_BLOCK_NAME, vbNullString, vbNullString, posSpecialBlock) Then
            m_ObjectTranslationsExist = False
            Exit Sub
        End If
    End If
    
    'posSpecialBlock now points to the start of the <special-translations> block in the source file
    ' (if one exists).
    
    'Start by noting that object ID mappings are available for this language.
    m_ObjectTranslationsExist = True
    
    'We now want to retrieve the list of object mappings specified in the language file.
    Dim mappingBlockText As String
    mappingBlockText = Trim$(m_XMLEngine.GetUniqueTag_String(SPECIAL_BLOCK_NAME, vbNullString, posSpecialBlock))
    
    'Failsafe only
    If (LenB(mappingBlockText) = 0) Then
        m_ObjectTranslationsExist = False
        Exit Sub
    End If
    
    'This is a little convoluted, but we're now going to dump this sub-block into its own XML object.
    ' (This makes it faster to parse, especially if it's large.)
    Dim mappingXML As pdXML
    Set mappingXML = New pdXML
    mappingXML.PrepareNewXML "object-text-mapping"
    mappingXML.WriteTag SPECIAL_BLOCK_NAME, mappingBlockText, False, False
    
    'From the new XML object, retrieve a list of all object IDs with custom translations
    Const OBJECT_ID_TAG As String = "object-id", OBJECT_TRANSLATION As String = "translation"
    Dim posObjectNames() As Long, numObjectNames As Long
    If mappingXML.FindAllTagLocations(posObjectNames, OBJECT_ID_TAG) Then
        
        numObjectNames = UBound(posObjectNames) + 1
        If (numObjectNames > 0) Then
            
            'Iterate all found locations and pull the *actual* object IDs and corresponding translations
            Dim i As Long
            For i = 0 To numObjectNames - 1
                
                Dim objectID As String, translatedText As String
                objectID = mappingXML.GetUniqueTag_String(OBJECT_ID_TAG, vbNullString, posObjectNames(i))
                translatedText = mappingXML.GetUniqueTag_String(OBJECT_TRANSLATION, vbNullString, posObjectNames(i) + Len(objectID))
                
                'Add all non-zero-length translations and IDs to a fast hash table
                If (LenB(translatedText) > 0) Then m_ObjectTranslations.AddItem objectID, translatedText
                
            Next i
            
            m_ObjectTranslationsExist = (m_ObjectTranslations.GetNumOfItems() > 0)
            
        End If
        
    Else
        PDDebug.LogAction "WARNING: bad special object table in language file!"
    End If
    
End Sub

Private Sub PerformLanguageFileCleanup()
    
    'This step is pointless if we are in non-portable mode (because we don't have access to our app folder)
    If Not UserPrefs.IsReady Then Exit Sub
    If UserPrefs.IsNonPortableModeActive Then Exit Sub
    
    '****
    '6.6 > 7.0 RELEASE CLEANUP
    
    'In PD 7.0, language filenames were switched to all-English to simplify their distribution and update process.
    ' When updating from 6.6 to 7.0, this can lead to duplicate copies of some language files - one file using
    ' the old name, and the new (updated) one using the new name.
    
    'To prevent duplicate copies, we manually check for some old entries and remove them if their newer
    ' counterpart is found.
    
    Dim langPath As String
    langPath = UserPrefs.GetLanguagePath()
    
    Dim targetFilename As String, replaceFilename As String
    targetFilename = langPath & "Francais.xml"
    replaceFilename = langPath & "French.xml"
    
    If (Files.FileExists(targetFilename) And Files.FileExists(replaceFilename)) Then
        
        'Check filesize.  This uses magic numbers taken from the last 6.6 release.
        If (Files.FileLenW(targetFilename) = 424711) Then Files.FileDelete targetFilename
        
    End If
    
    'Repeat above steps for all other relevant language files
    targetFilename = langPath & "Espanol.xml"
    replaceFilename = langPath & "Spanish_(Mexico).xml"
    If Files.FileExists(targetFilename) And Files.FileExists(replaceFilename) Then If (Files.FileLenW(targetFilename) = 421189) Then Files.FileDelete targetFilename
    
    targetFilename = langPath & "Italiano.xml"
    replaceFilename = langPath & "Italian.xml"
    If Files.FileExists(targetFilename) And Files.FileExists(replaceFilename) Then If (Files.FileLenW(targetFilename) = 423659) Then Files.FileDelete targetFilename
    
    targetFilename = langPath & "Portuguese.xml"
    replaceFilename = langPath & "Portuguese_(Brazil).xml"
    If Files.FileExists(targetFilename) And Files.FileExists(replaceFilename) Then If (Files.FileLenW(targetFilename) = 416158) Then Files.FileDelete targetFilename
    
    targetFilename = langPath & "Deutsch_1.xml"
    replaceFilename = langPath & "German_1.xml"
    If Files.FileExists(targetFilename) And Files.FileExists(replaceFilename) Then If (Files.FileLenW(targetFilename) = 417364) Then Files.FileDelete targetFilename
    
    targetFilename = langPath & "Deutsch_2.xml"
    replaceFilename = langPath & "German_2.xml"
    If Files.FileExists(targetFilename) And Files.FileExists(replaceFilename) Then If (Files.FileLenW(targetFilename) = 418885) Then Files.FileDelete targetFilename
    
    'END 6.6 > 7.0 RELEASE CLEANUP
    '****
    
End Sub

'Populate the Tools -> Languages menu.  Note that most of this is now handled via the Menus module (which allows us to display Unicode names).
Private Sub PopulateLanguageMenu()

    'Start by unloading all current language menu entries
    Dim i As Long
    If (FormMain.MnuLanguages.Count > 1) Then
        For i = FormMain.MnuLanguages.Count - 1 To 0 Step -1
            Unload FormMain.MnuLanguages(i)
        Next i
    End If
    
    'Load all the menu entries we'll be needing.
    If (m_numOfLanguages > 1) Then
        For i = 1 To m_numOfLanguages - 1
            Load FormMain.MnuLanguages(i)
        Next i
    End If
    
    'Make sure all menus are un-checked
    For i = 0 To m_numOfLanguages - 1
        FormMain.MnuLanguages(i).Checked = False
    Next i
    
    'Use the menus module to update menu captions; this allows for Unicode text
    Menus.UpdateSpecialMenu_Language m_numOfLanguages, m_AvailableLanguages

End Sub

'Given a language file (or partial language file), extract the key language information and place it in the passed variable.
Private Sub FillLanguageInfo(ByRef srcXmlEngine As pdXML, ByRef targetLangHolder As PDLanguageFile)

    'First, get the language ID and name - these are the most important values, and technically the only REQUIRED ones.
    With targetLangHolder
        
        .langID = srcXmlEngine.GetUniqueTag_String("langid")
        .LangName = srcXmlEngine.GetUniqueTag_String("langname")
    
        'Version, status, and author information should also be present, but the file will still be loaded even if
        ' they don't exist
        .langVersion = srcXmlEngine.GetUniqueTag_String("langversion")
        .LangStatus = srcXmlEngine.GetUniqueTag_String("langstatus")
        .Author = srcXmlEngine.GetUniqueTag_String("author")
        
    End With
        
End Sub

Private Sub GetCachedLanguageInfo(ByRef srcFilename As String, ByRef dstLangHolder As PDLanguageFile)
    With dstLangHolder
        .langID = m_LangCacheXML.GetUniqueTag_String(srcFilename & "-langid")
        .LangName = m_LangCacheXML.GetUniqueTag_String(srcFilename & "-langname")
        .langVersion = m_LangCacheXML.GetUniqueTag_String(srcFilename & "-langversion")
        .LangStatus = m_LangCacheXML.GetUniqueTag_String(srcFilename & "-langstatus")
        .Author = m_LangCacheXML.GetUniqueTag_String(srcFilename & "-langauthor")
    End With
End Sub

'Cache key language information inside the user's preference file.  This allows for faster retrieval on future runs.
Private Sub CacheLanguageInfo(ByRef srcFilename As String, ByRef srcLangHolder As PDLanguageFile)
    With srcLangHolder
        m_LangCacheXML.UpdateTag srcFilename & "-langid", .langID
        m_LangCacheXML.UpdateTag srcFilename & "-langname", .LangName
        m_LangCacheXML.UpdateTag srcFilename & "-langversion", .langVersion
        m_LangCacheXML.UpdateTag srcFilename & "-langstatus", .LangStatus
        m_LangCacheXML.UpdateTag srcFilename & "-langauthor", .Author
    End With
    m_LangCacheChanged = True
End Sub

'Sort the m_availableLanguages() array alphabetically, using language names as the sort parameter
Private Sub SortLanguageData()

    Dim i As Long, j As Long
    Dim langTmp As PDLanguageFile
    
    'Loop through all entries in the languages array, sorting them as we go
    For i = 0 To m_numOfLanguages - 1
        For j = 0 To m_numOfLanguages - 1
            
            'Compare two language names, and if one is less (e.g. earlier alphabetically) than the other, swap them
            If (Strings.StrCompSort(m_AvailableLanguages(i).LangName, m_AvailableLanguages(j).LangName) < 0) Then
                langTmp = m_AvailableLanguages(i)
                m_AvailableLanguages(i) = m_AvailableLanguages(j)
                m_AvailableLanguages(j) = langTmp
            End If
            
        Next j
    Next i

End Sub

'External functions can use this retrieve the current pdLanguageFile contents for a given language ID.
' Note that this ONLY SUPPORTS OFFICIAL LANGUAGE FILES, by design.  User language files should not be modified by PD itself.
' Returns TRUE if a match is found; FALSE otherwise.
Friend Function GetPDLanguageFileObject(ByRef dstLanguageFileObject As PDLanguageFile, ByVal languageID As String) As Boolean
    
    languageID = Trim$(UCase$(languageID))
    
    'Find the matching language ID
    Dim i As Long
    For i = 0 To UBound(m_AvailableLanguages)
    
        If m_AvailableLanguages(i).IsOfficial Then
            
            If Strings.StringsEqual(languageID, Trim$(m_AvailableLanguages(i).langID), True) Then
                
                'It's a match!  Fill the target object and exit
                dstLanguageFileObject = m_AvailableLanguages(i)
                GetPDLanguageFileObject = True
                Exit Function
            
            End If
        
        End If
    
    Next i
    
    'If we made it here, no matching language was found
    GetPDLanguageFileObject = False

End Function

'All text to be translated should be passed here first.  This sub will perform a few automatic checks of the string, and it will record what
' it finds.  For example, if the text has "..." at the end, this sub will remove that.  Its counterpart - postProcessText() - will re-apply
' any necessary changes to the translation.
Private Sub PreProcessText(ByRef srcString As String)

    '1) Trim the string
    srcString = Trim$(srcString)
    
    '2) Check for a trailing "..."
    m_hasEllipsis = (Right$(srcString, 3) = "...")
    If m_hasEllipsis Then srcString = Left$(srcString, Len(srcString) - 3)
    
    '3) Check for a trailing colon ":"
    m_hasTrailingColon = (Right$(srcString, 1) = ":")
    If m_hasTrailingColon Then srcString = Left$(srcString, Len(srcString) - 1)
    
End Sub

Private Sub PostProcessText(ByRef srcString As String)

    '1) Trim the string
    srcString = Trim$(srcString)
    
    '2) If the string is now empty, return a blank string rather than attempting to reconstruct it
    If (LenB(srcString) <> 0) Then
    
        '2) If the original had a trailing "...", restore it now
        If m_hasEllipsis Then
            srcString = srcString & "..."
            m_hasEllipsis = False
        End If
        
        '3) If the original had a trailing colon, restore it now
        If m_hasTrailingColon Then
            srcString = srcString & ":"
            m_hasTrailingColon = False
        End If
    End If

End Sub

'Given the translated caption of a message or control, return the original translation from the active translation file
Private Function GetOriginalTagFromTranslation(ByVal curCaption As String) As String

    'If translations aren't active, return the requested caption as-is
    If Me.TranslationActive Then

        'Remove white space from the caption (if necessary, white space will be added back in after retrieving the translation from file)
        PreProcessText curCaption
        
        Dim phraseLocation As Long
        phraseLocation = m_XMLEngine.GetLocationOfParentTag("phrase", "translation", curCaption)
        
        'Make sure a phrase tag was found
        If (phraseLocation > 0) Then
            
            'Retrieve the <translation> tag inside this phrase tag
            curCaption = m_XMLEngine.GetUniqueTag_String("original", vbNullString, phraseLocation)
            
        End If
        
        'Apply any relevant post-processing, then return the text
        PostProcessText curCaption
        GetOriginalTagFromTranslation = curCaption
        
    Else
        GetOriginalTagFromTranslation = curCaption
    End If

End Function

'Given the original caption of a message or control, return the matching translation from the active translation file
Private Function GetTranslationTagFromCaption(ByVal origCaption As String) As String

    If (m_XMLEngine Is Nothing) Then
        GetTranslationTagFromCaption = origCaption
        Exit Function
    End If

    'Remove white space from the caption (if necessary, white space will be added back in after retrieving the translation from file)
    PreProcessText origCaption
    
    Dim phraseLocation As Long
    phraseLocation = m_XMLEngine.GetLocationOfParentTag("phrase", "original", origCaption, True)
    
    'Make sure a phrase tag was found
    If (phraseLocation > 0) Then
        
        'Retrieve the <translation> tag inside this phrase tag
        origCaption = m_XMLEngine.GetUniqueTag_String("translation", vbNullString, phraseLocation)
        PostProcessText origCaption
        If (LenB(origCaption) <> 0) Then GetTranslationTagFromCaption = origCaption Else GetTranslationTagFromCaption = vbNullString
        
    Else
        GetTranslationTagFromCaption = vbNullString
    End If

End Function

'Given a source message (an English string), return a translated version using the active language file.
' RETURNS: translated text if found, unmodified original en-US string if translated text is *not* found.
Friend Function TranslateMessage(ByRef srcMessage As String, ParamArray ExtraText() As Variant) As String
    
    'Translations are profiled to prevent perf regressions
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Parameters can pass run-time text as optional parameters
    Dim numOptionalParams As Long
    If (UBound(ExtraText) >= LBound(ExtraText)) Then numOptionalParams = (UBound(ExtraText) - LBound(ExtraText)) + 1
    
    'In version 9.2+, objects can now (optionally) identify themselves with a unique object ID.
    ' PhotoDemon can use this ID to assign special translations that do not map 1:1 to English text
    ' (based on caption alone.)
    
    'See https://github.com/tannerhelland/PhotoDemon/issues/491 for details on why this solution was necessary.
    Dim searchString As String, newString As String
    Dim customIdTranslationFound As Boolean
    
    'See if this optional feature is active for this language
    If m_ObjectTranslationsExist Then
        
        'See if the caller passed any optional params
        If (numOptionalParams >= 1) Then
            
            searchString = ExtraText(numOptionalParams - 1)
            
            'Look for passed object IDs.  (These IDs must always be passed as the *last* parameter, using the prefix `obj-`.)
            If (Len(searchString) > 4) Then
                If (Left$(searchString, Len(SPECIAL_TRANSLATION_OBJECT_PREFIX)) = SPECIAL_TRANSLATION_OBJECT_PREFIX) Then
                    
                    'This object passed a unique ID as its last parameter.  Retrieve the object name, and look for a
                    ' matching translation in the special object-ID-based translation table.
                    customIdTranslationFound = m_ObjectTranslations.GetItemByKey(Right$(searchString, Len(searchString) - Len(SPECIAL_TRANSLATION_OBJECT_PREFIX)), newString)
                    
                    'Regardless of success or failure, decrement the parameter count to note that an object ID was passed.
                    ' (We can use this knowledge to skip some parameter processing later in this function.)
                    numOptionalParams = numOptionalParams - 1
                    
                End If
            End If
            
        End If
    End If
    
    If Me.TranslationActive And (LenB(srcMessage) <> 0) Then
        
        'Skip translations that were already applied using special per-object-id rules
        If (Not customIdTranslationFound) Then
            
            newString = vbNullString
            searchString = vbNullString
            
            'Always start by querying the fast translation cache.  Previously accessed translations
            ' will be stored there.
            Dim fastTranslationWorked As Boolean
            If (m_FoundTranslations Is Nothing) Then
                Set m_FoundTranslations = New pdStringHash
            Else
                fastTranslationWorked = m_FoundTranslations.GetItemByKey(srcMessage, newString)
            End If
            
            If fastTranslationWorked Then m_CacheHits = m_CacheHits + 1 Else m_CacheMisses = m_CacheMisses + 1
            
            'If we didn't find the requested translation, pop into the full translation file to grab it.
            If (LenB(newString) = 0) Then
                searchString = srcMessage
                newString = GetTranslationTagFromCaption(searchString)
            End If
            
            'If no translation was found, return the original message as the "translation".
            ' (This makes the program usable if a translation file is incomplete.)
            If (LenB(newString) = 0) Then
                
                'If the source caption contains vbCrLf characters, replace them with plain vbLF characters and try again.
                ' (This solves issues when software - including 3rd-party Git clients - modifies line endings on
                '  localization files.)
                If (InStr(1, searchString, vbCrLf, vbBinaryCompare) <> 0) Then
                    
                    newString = GetTranslationTagFromCaption(Replace$(searchString, vbCrLf, vbLf, 1, -1, vbBinaryCompare))
                    
                    If (LenB(newString) = 0) Then
                        newString = srcMessage
                    Else
                    
                        'A match was found!  Convert vbLf to vbCrLf, to ensure proper rendering.
                        ' (Thanks to Frank Donckers for this fix!)
                        newString = Replace$(newString, vbLf, vbCrLf, 1, -1, vbBinaryCompare)
                    
                    End If
                    
                Else
                    newString = srcMessage
                End If
                
            End If
            
            'If the fast translation cache didn't contain this translation, add the manually retrieved value now.
            ' (This makes subsequent lookups very fast.)
            If (Not fastTranslationWorked) Then m_FoundTranslations.AddItem srcMessage, newString
            
        '/end "skip because we already applied a per-object-ID translation" check
        End If
            
    'If translations are not active, just use the existing string "as-is"
    Else
        newString = srcMessage
    End If
    
    'newString now contains the translated message (or the original text, if no translation was found).
    
    'The last step is to replace any optional parameters in the text
    If (numOptionalParams > 0) Then
        Dim i As Long
        For i = 0 To numOptionalParams - 1
            newString = Replace$(newString, "%" & CStr(i + 1), ExtraText(i), 1, -1, vbBinaryCompare)
        Next i
    End If
    
    'Return the fully translated string
    TranslateMessage = newString
    
    'Track net time spent in translation
    m_NetTranslateTime = m_NetTranslateTime + VBHacks.GetTimerDifferenceNow(startTime)
    
End Function

'Enumerate all translations on a form (and its controls), and replace them with their original English equivalents
Friend Sub UndoTranslations(ByRef srcForm As Form)
    
    'Unicode form captions require special help; see GetWindowCaptionW for details.
    Dim searchString As String, newString As String
    If (g_WindowManager Is Nothing) Then
        searchString = srcForm.Caption
    Else
        searchString = g_WindowManager.GetWindowCaptionW(srcForm.hWnd)
    End If
    
    newString = GetOriginalTagFromTranslation(searchString)
    
    'If a new string was found, apply it to the object's caption
    If (LenB(newString) <> 0) Then
        If (Not g_WindowManager Is Nothing) Then
            g_WindowManager.SetWindowCaptionW srcForm.hWnd, " " & newString
        Else
            srcForm.Caption = " " & newString
        End If
    End If
    
End Sub

'Apply any special handling to a given form.  At present, this only affects form captions; everything else is handled internally.
Friend Sub ApplyTranslations(ByRef srcForm As Form)
    
    'For en-US locales, we can entirely skip this step
    If Me.TranslationActive Then
    
        'At present, all we have to do is translate the form's caption
        If (LenB(srcForm.Caption) <> 0) Then
        
            Dim searchString As String, newString As String
            searchString = srcForm.Caption
            newString = GetTranslationTagFromCaption(searchString)
            
            'If a new string was found, apply it to the object's caption
            If (LenB(newString) <> 0) Then
                If (Not g_WindowManager Is Nothing) Then
                    g_WindowManager.SetWindowCaptionW srcForm.hWnd, " " & newString
                Else
                    srcForm.Caption = " " & newString
                End If
            End If
            
        End If
        
    End If
    
End Sub

'In developer builds, net translation engine time is tracked for perf reasons
Friend Function GetNetTranslationTime() As Double
    GetNetTranslationTime = m_NetTranslateTime
End Function

Friend Sub PrintAdditionalDebugInfo()
    
    'Since we moved to a hash table for retrieved translations, I'm interested in what percentage
    ' of retrievals hit that table vs the original XML file.
    If (m_CacheMisses <> 0) Then
        PDDebug.LogAction "Translation cache had " & m_CacheHits & " hits and " & m_CacheMisses & " misses (" & Format$(m_CacheHits / (m_CacheHits + m_CacheMisses), "0%") & ")"
    End If
    
    'If the fast translation cache exists, it may have additional debug info to report
    If (Not m_FoundTranslations Is Nothing) Then m_FoundTranslations.PrintClassDebugInfo
    
End Sub

