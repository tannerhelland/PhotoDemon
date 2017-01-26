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
'Copyright 2012-2017 by Frank Donckers and Tanner Helland
'Created: 21/December/12
'Last updated: 26/January/17
'Last update: use a dedicate language file settings cache to greatly improve startup performance
'
'Thanks to the incredible work of Frank Donckers, PhotoDemon now has a working translation engine.  Many
' thanks to Frank for taking the initiative on not only implementing the translation engine prototype, but
' also for taking the time to translate the entire PhotoDemon text collection into multiple languages.
' (This was a huge project, as PhotoDemon contains a LOT of text.)
'
'The latest revision of the translation engine operates on translation files stored as XML.  You can see
' (and freely edit) these files inside the /App/PhotoDemon/Languages/ subfolder of wherever PD is installed.
' If you do decide to edit these language files, I strongly suggest saving the new files inside the USER
' folder of the project, e.g. /Data/Languages/.  This will ensure they are not overwritten by software
' updates or future changes to the official language files included with PD.
'
'Most of the crucial translation work in the program happens inside a global object called "g_Language",
' which is an object of this class (pdTranslate).  Most subs and functions should be self-explanatory.  To
' really get a feel for how the translation engine works, I suggest starting with the Loading module, and
' looking at the way this class is initialized and loaded.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'NOTE: this class frequently uses a type called pdLanguageFile.  Because outside functions (like the Language Editor) also need
' access to this type, I have moved it to the Public_Variables module.

'Language-related API calls
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

'Constants related to API language calls
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SISO639LANGNAME As Long = &H59   'ISO 639 abbreviated language name, e.g. "en" for "English"
Private Const LOCALE_SISO3166CTRYNAME As Long = &H5A   'ISO 4166 abbreviated country/region name, e.g. "US" for "United States"

'Array of currently available languages
Private m_numOfLanguages As Long
Private m_availableLanguages() As pdLanguageFile

'Array ID of the current language.  If no language is selected (e.g. PhotoDemon is using its default text), this value is -1
Private m_currentLanguage As Long

'Array ID of the default language (English - US)
Private m_defaultLanguage As Long

'If the user needs to be shown a language select screen, this will be set to true at some point in the initialization process
Private m_needToShowLanguageDialog As Boolean

'If the select language dialog needs to be shown, this value can be used to determine "why"
Private Const LANGREASON_USERINITIATED As Long = 0      ' 0) User-initiated dialog
Private Const LANGREASON_APPROXIMATEMATCH As Long = 1   ' 1) Approximate (but not exact) language match found
Private Const LANGREASON_NOMATCH As Long = 2            ' 2) No language match found
Private Const LANGREASON_FILEMISSING As Long = 3        ' 3) Requested language file couldn't be located
Private m_reasonForLanguageDialog As Long

'The full contents of the translation file will be loaded into this string.
Private m_XMLEngine As pdXML

'Is the class ready to be used for translations?
Private m_isReady As Boolean

'The following boolean values are marked by preProcessString, then read by postProcessString
Private m_hasDotDotDot As Boolean
Private m_hasTrailingColon As Boolean

'If the active language file is updated at run-time, this value will be set to TRUE.  PD's internal processor calls the refreshAsRequired() function
' whenever a user-initiated action triggers, and refreshAsRequired() uses this value to know whether we need to re-translate the main form.
Private m_ActiveLanguageWasPatched As Boolean

'File interactions in this class are handled by pdFSO
Private m_FSO As pdFSO

'Language files are checksummed to improve startup time.  This class is not guaranteed to exist, so please review individual
' functions before attempting to use the crypto object.
Private m_Crypto As pdCrypto

'The language file cache is handled inside this XML object.
Private m_LangCacheXML As pdXML, m_LangCacheChanged As Boolean

Public Sub RefreshAsRequired()
    
    If m_ActiveLanguageWasPatched Then
        
        Debug.Print "Patching active language file now..."
        
        m_ActiveLanguageWasPatched = False
        UndoTranslations FormMain
        ApplyLanguage False
        
    End If
    
End Sub

'External functions can request the a copy of all found languages; they must provide an array of type pdLanguageFile, which this
' function will then fill with the data assembled from Checkm_availableLanguages().
Friend Sub CopyListOfLanguages(ByRef dstListOfLanguages() As pdLanguageFile)
    
    ReDim dstListOfLanguages(0 To m_numOfLanguages - 1) As pdLanguageFile
    
    'Copy all language data entries into the new array.
    Dim i As Long
    For i = 0 To m_numOfLanguages - 1
        dstListOfLanguages(i) = m_availableLanguages(i)
    Next i
    
End Sub

'External functions can request the current language, which may then be used to request translations from
' 3rd-party plugins (specifically, ExifTool).
Public Function GetCurrentLanguage() As String
    If TranslationActive Then
        GetCurrentLanguage = Left$(m_availableLanguages(m_currentLanguage).langID, 2)
    Else
        GetCurrentLanguage = "en"
    End If
End Function

'This function will return PD's internal language index (m_currentLanguage), which matches the language's index in the Language menu.
Public Function GetCurrentLanguageIndex() As Long
    GetCurrentLanguageIndex = m_currentLanguage
End Function

'External function for asking the translator if a language selection dialog is needed
Public Function IsLanguageDialogNeeded(Optional ByRef reasonForDialog As Long = -1) As Boolean
    If (reasonForDialog <> -1) Then reasonForDialog = m_reasonForLanguageDialog
    IsLanguageDialogNeeded = m_needToShowLanguageDialog
End Function

'If a translation file has been requested and successfully loaded, this function will return TRUE
Public Function ReadyToTranslate() As Boolean
    If (Not m_XMLEngine Is Nothing) Then
        ReadyToTranslate = m_isReady
    Else
        ReadyToTranslate = False
    End If
End Function

'If the currently active language is different from the default language, this function will return TRUE
Public Function TranslationActive() As Boolean
    TranslationActive = Not (m_currentLanguage = m_defaultLanguage)
End Function

Public Sub WriteLanguagePreferencesToFile()
    g_UserPreferences.SetPref_String "Language", "Current Language File", m_availableLanguages(m_currentLanguage).FileName
End Sub

'Given an index in the m_availableLanguages array, mark that as the new language (including saving it to the preferences file)
Public Sub ActivateNewLanguage(ByVal newLangIndex As Long, Optional ByVal useDoEvents As Boolean = False)
    m_currentLanguage = newLangIndex
    Interface.GenerateInterfaceID
    WriteLanguagePreferencesToFile
    ApplyLanguage useDoEvents
End Sub

'If the language file updater has updated one or more language files, it will call this function.  Any subsequent language update functions (such as loading a
' new XML file) happen here.
Public Sub NotifyHotPatchingComplete()

    'Figure out if the current language is affected by the patch
    If (m_currentLanguage <> m_defaultLanguage) Then
    
        'Retrieve the current language version
        Dim curLangVersion As String
        curLangVersion = m_XMLEngine.GetUniqueTag_String("langversion")
        
        'Retrieve the version of the XML file on disk; if an update has occurred, we can detect it by comparing version numbers.
        Dim tmpXML As pdXML
        Set tmpXML = New pdXML
        If tmpXML.LoadXMLFile(m_availableLanguages(m_currentLanguage).FileName) Then
            
            'Retrieve the version number from the potentially updated file
            Dim newLanguageVersion As String
            newLanguageVersion = tmpXML.GetUniqueTag_String("langversion")
            
            If (StrComp(curLangVersion, newLanguageVersion, vbBinaryCompare) <> 0) Then
            
                'The version numbers do not match.  Set a module-level refresh flag; PD's internal processor will a refresh function the next time a
                ' user-initiated action completes, and that function will check this module-level value to see if a refresh is required.
                m_ActiveLanguageWasPatched = True
            
            Else
                m_ActiveLanguageWasPatched = False
            End If
            
        Else
            m_ActiveLanguageWasPatched = False
        End If
    
    Else
        m_ActiveLanguageWasPatched = False
    End If

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
        curLangID = "en-US" 'GetDefaultUserLanguage()
                
        'Compare the full language ID (language and region) against language files available in the Languages folder.
        ' If an exact match is found, present that as the default program language.
        m_currentLanguage = IsExactLanguageMatch(curLangID)
        
        'If an exact match is not found, try comparing just the language part of the code.  If a match is found,
        ' present that as the default language - and note that we also need to show the language selection dialog.
        If m_currentLanguage = -1 Then
            m_currentLanguage = IsApproximateLanguageMatch(curLangID)
            m_reasonForLanguageDialog = LANGREASON_APPROXIMATEMATCH
            m_needToShowLanguageDialog = True
        End If
        
        'If we still can't find a match, present the language selection screen in English.
        If m_currentLanguage = -1 Then
            m_currentLanguage = m_defaultLanguage
            m_reasonForLanguageDialog = LANGREASON_NOMATCH
            m_needToShowLanguageDialog = True
        End If
        
    'If this is not the first run, retrieve the user's language preference from the preferences file
    Else
    
        Dim curLangFile As String
        curLangFile = g_UserPreferences.GetPref_String("Language", "Current Language File", "")
        
        'If no specific file has been named, use the default language
        If (Len(curLangFile) = 0) Then
            m_currentLanguage = m_defaultLanguage
        
        'If a file HAS been named, make sure it still exists; if it doesn't, default to English (US) and show the language dialog
        Else
            If m_FSO.FileExist(curLangFile) Then
                m_currentLanguage = GetLangIndexFromFile(curLangFile)
            Else
                m_currentLanguage = m_defaultLanguage
                m_reasonForLanguageDialog = LANGREASON_FILEMISSING
                m_needToShowLanguageDialog = True
            End If
        End If
    
    End If
    
    m_isReady = True
    
End Sub

'Given a language filename, return the index in the m_availableLanguages() array
Private Function GetLangIndexFromFile(ByVal srcFile As String) As Long
    
    Dim i As Long
    For i = 0 To m_numOfLanguages - 1
        If (StrComp(m_availableLanguages(i).FileName, srcFile, vbTextCompare) = 0) Then
            GetLangIndexFromFile = i
            Exit Function
        End If
    Next i
    
End Function

'Given a language ID (as a string), try to find an approximate match (just language, region doesn't matter) in the m_availableLanguages array
Private Function IsApproximateLanguageMatch(ByRef srcLangID As String)

    Dim i As Long
    For i = 0 To m_numOfLanguages - 1
        If (StrComp(Left$(srcLangID, 2), Left$(m_availableLanguages(i).langID, 2), vbTextCompare) = 0) Then
            IsApproximateLanguageMatch = i
            'pdMsgBox "Approximate language match found: " & m_availableLanguages(i).langID
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, no match was found.  Return -1
    IsApproximateLanguageMatch = -1

End Function

'Given a language ID (as a string), try to find an exact match (language AND region) in the m_availableLanguages array
Private Function IsExactLanguageMatch(ByRef srcLangID As String)

    Dim i As Long
    For i = 0 To m_numOfLanguages - 1
        If (StrComp(srcLangID, m_availableLanguages(i).langID, vbTextCompare) = 0) Then
            IsExactLanguageMatch = i
            'pdMsgBox "Exact language match found: " & m_availableLanguages(i).langID
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, no match was found.  Return -1
    IsExactLanguageMatch = -1

End Function

'ALWAYS CALL DetermineLanguage BEFORE THIS STAGE!
'Once a PhotoDemon language has been selected (via DetermineLanguage), this function can be called to "apply" that language to the program.
' NOTE: if the language is changed, this function must be called again to set the new language program-wide.
Public Sub ApplyLanguage(Optional ByVal useDoEvents As Boolean = False)

    'Mark the active language menu entry (and unmark all others)
    Dim i As Long
    For i = 0 To m_numOfLanguages - 1
        FormMain.mnuLanguages(i).Checked = CBool(m_currentLanguage = i)
    Next i
    
    'If the language is set to anything other than the default, load the contents of the language XML file into memory.
    If TranslationActive Then
        Set m_XMLEngine = New pdXML
        m_XMLEngine.LoadXMLFile m_availableLanguages(m_currentLanguage).FileName
        
        'New as of August '14 is use of binary text comparisons when matching translations.  This provides a large
        ' performance boost, and also results in correct captioning when the same word is used in multiple places
        ' with different cases (e.g. menus vs sliders)
        m_XMLEngine.SetTextCompareMode vbBinaryCompare
        
    End If
    
    Interface.RedrawEntireUI useDoEvents, True
    
End Sub

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
    abbrLang = LCase$(GetSpecificLocaleInfo(LOCALE_SISO639LANGNAME))
    
    'Get the ISO 4166 abbreviated country/region name (e.g. "US" for "United States")
    abbrRegion = LCase$(GetSpecificLocaleInfo(LOCALE_SISO3166CTRYNAME))
    
    'Return the language codes
    GetDefaultUserLanguage = abbrLang & "-" & abbrRegion

End Function

'Given a locale information constant, return a corresponding string
Private Function GetSpecificLocaleInfo(ByVal lInfo As Long) As String
    Dim sBuffer As String, sRet As String
    sBuffer = String$(256, 0)
    sRet = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, sBuffer, Len(sBuffer))
    If (sRet > 0) Then
        GetSpecificLocaleInfo = Left$(sBuffer, sRet - 1)
    Else
        GetSpecificLocaleInfo = ""
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
    isOfficialFolder = (StrComp(UCase(langTypeForThisFolder), "OFFICIAL", vbBinaryCompare) = 0)
    
    'An XML class will be used to parse and validate language files, as needed.  (This class may not be instantiated
    ' if the current language file list matches the list at last-run, because language values are cached.  Plan accordingly.)
    Dim tmpXMLEngine As pdXML
    
    'Scan the translation folder for .xml files.  Ignore anything that isn't XML.
    Dim listOfFiles As pdStringStack
    Set listOfFiles = New pdStringStack
    
    If m_FSO.RetrieveAllFiles(srcFolder, listOfFiles, False, False, "xml") Then
        
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
            
            'Look for an existing checksum
            chkXMLName = m_LangCacheXML.GetXMLSafeTagName(m_FSO.GetFilename(chkFile, True))
            chkCachedValue = Trim$(m_LangCacheXML.GetUniqueTag_String(chkXMLName))
            If (Len(chkCachedValue) <> 0) Then
            
                'A checksum was found.  If it matches the file's current checksum, we can skip further validation.
                fileValidationRequired = Not (StrComp(chkCachedValue, GetLangFileChecksum(chkFile), vbBinaryCompare) = 0)
                
                #If DEBUGMODE = 1 Then
                    If fileValidationRequired Then pdDebug.LogAction "Note: a language file checksum didn't match (" & chkXMLName & ").  Manual validation will follow."
                #End If
                
            End If
            
            'If validation is required, we need to load the file from scratch and run some checks on it
            If fileValidationRequired Then
                
                If (tmpXMLEngine Is Nothing) Then
                    Set tmpXMLEngine = New pdXML
                    tmpXMLEngine.SetTextCompareMode vbBinaryCompare
                End If
                
                'Use PD's XML engine to load the file
                If tmpXMLEngine.LoadXMLFile(chkFile) Then
                
                    'Use the XML engine to validate this file, and to make sure it contains at least a language ID, name, and one (or more) translated phrase
                    If tmpXMLEngine.IsPDDataType("Translation") Then
                        
                        'If this is a valid language file, get the language information (if available).
                        FillLanguageInfo tmpXMLEngine, m_availableLanguages(nLanguages)
                        m_availableLanguages(nLanguages).FileName = chkFile
                        m_availableLanguages(nLanguages).LangType = langTypeForThisFolder
                        m_availableLanguages(nLanguages).IsOfficial = isOfficialFolder
                        
                        'Because this is a valid language file, we want to cache its key properties in the
                        ' master user preferences file.  This lets us skip validation on subsequent runs.
                        m_LangCacheXML.UpdateTag chkXMLName, GetLangFileChecksum(chkFile)
                        CacheLanguageInfo chkXMLName, m_availableLanguages(nLanguages)
                        
                        'If, for some ungodly reason, the user has more than 100 languages available, dynamically resize the array to fit.
                        nLanguages = nLanguages + 1
                        If (nLanguages > UBound(m_availableLanguages)) Then ReDim Preserve m_availableLanguages(0 To (nLanguages + 1) * 2) As pdLanguageFile
                        
                    End If
                    
                End If
            
            'If file validation is *not* required, pull previously cached values from the preference file
            Else
            
                GetCachedLanguageInfo chkXMLName, m_availableLanguages(nLanguages)
                m_availableLanguages(nLanguages).FileName = chkFile
                m_availableLanguages(nLanguages).LangType = langTypeForThisFolder
                m_availableLanguages(nLanguages).IsOfficial = isOfficialFolder
                
                nLanguages = nLanguages + 1
                If (nLanguages > UBound(m_availableLanguages)) Then ReDim Preserve m_availableLanguages(0 To (nLanguages + 1) * 2) As pdLanguageFile
            
            End If
            
        'Retrieve the next file and repeat
        Loop
        
    End If
    
    TallyLanguageFilesInFolder = nLanguages - langOffset

End Function

Private Function GetLangFileChecksum(ByRef srcFile As String) As String

    'Mix together the filename, file last access date, and file size to arrive at a "good enough" unique string.
    Dim strHashSource As String
    strHashSource = srcFile & "|" & CStr(m_FSO.FileLenW(srcFile)) & "|" & CStr(m_FSO.GetFileTimeAsCurrency(srcFile, PDFT_WriteTime))
    
    'A dedicated crypto class handles the actual hashing
    GetLangFileChecksum = m_Crypto.QuickHashString(strHashSource, 64)
    
End Function

'Search the Languages folder, and make a list of all available languages
Public Sub CheckAvailableLanguages()
    
    m_numOfLanguages = 0
    ReDim m_availableLanguages(0 To 99) As pdLanguageFile
    
    'Language file names and attributes are cached inside a dedicated language settings file.  This file lets
    ' us bypass a lot of obnoxious and time-consuming validation steps.
    Dim langCacheFilename As String
    langCacheFilename = g_UserPreferences.GetLanguagePath(True) & "LangCache.xml"
    
    Dim cacheLoadedSuccessfully As Boolean: cacheLoadedSuccessfully = False
    
    Set m_LangCacheXML = New pdXML
    m_LangCacheXML.SetTextCompareMode vbBinaryCompare
    
    If m_FSO.FileExist(langCacheFilename) Then
        cacheLoadedSuccessfully = m_LangCacheXML.LoadXMLFile(langCacheFilename)
    End If
    
    'If the cache exists, we won't write it back out to file unless it somehow changes during this session
    ' (from a new language file being discovered, for example).  If it does *not* exist, we'll create it anew
    ' and force it to write out to file at the end of this initiation.
    If (Not cacheLoadedSuccessfully) Then
        m_LangCacheXML.PrepareNewXML "PDLanguageCache"
        m_LangCacheChanged = True
    End If
    
    'Find all available OFFICIAL language files
    m_numOfLanguages = TallyLanguageFilesInFolder(g_UserPreferences.GetLanguagePath(False), 0, "Official")
    
    'Find all available USER language files
    m_numOfLanguages = m_numOfLanguages + TallyLanguageFilesInFolder(g_UserPreferences.GetLanguagePath(True), m_numOfLanguages, "Unofficial")
    
    'Add a dummy entry for PhotoDemon's default language (en-US)
    m_availableLanguages(m_numOfLanguages).FileName = ""
    m_availableLanguages(m_numOfLanguages).langID = "en-US"
    m_availableLanguages(m_numOfLanguages).LangName = "English (US) - Default"
    m_availableLanguages(m_numOfLanguages).LangType = "Default"
    m_availableLanguages(m_numOfLanguages).LangStatus = "Complete"
    m_availableLanguages(m_numOfLanguages).langVersion = App.Major & "." & App.Minor & "." & App.Revision
    m_availableLanguages(m_numOfLanguages).IsOfficial = True
    m_numOfLanguages = m_numOfLanguages + 1
    
    'Resize the array to contain only the number of languages found
    ReDim Preserve m_availableLanguages(0 To m_numOfLanguages - 1) As pdLanguageFile
    
    'Sort the array alphabetically by language name (because we are going to fill the languages submenu with the entries)
    SortLanguageData
    
    'Populate the languages menu
    PopulateLanguageMenu
    
    'Mark the location in the array of the default language, as other functions may need to know it
    Dim i As Long
    For i = 0 To m_numOfLanguages - 1
        If (StrComp(UCase$(m_availableLanguages(i).LangName), "ENGLISH (US) - DEFAULT", vbBinaryCompare) = 0) Then
            m_defaultLanguage = i
            Exit For
        End If
    Next i
    
    'If the cached language data was modified during this session, write the updated cache out to file
    If m_LangCacheChanged Then m_LangCacheXML.WriteXMLToFile langCacheFilename
    Set m_LangCacheXML = Nothing
    
    'TESTING: display the info we retrieved
    'Dim i As Long
    'For i = 0 To m_numOfLanguages - 1
    '    pdMsgBox m_availableLanguages(i).langID & vbCrLf & m_availableLanguages(i).langName & vbCrLf & m_availableLanguages(i).langStatus & vbCrLf & m_availableLanguages(i).langVersion
    'Next i
    
End Sub

'Populate the Tools -> Languages menu
Private Sub PopulateLanguageMenu()

    'Start by unloading all current language menu entries
    Dim i As Long
    If (FormMain.mnuLanguages.Count > 1) Then
        For i = FormMain.mnuLanguages.Count - 1 To 0 Step -1
            Unload FormMain.mnuLanguages(i)
        Next i
    End If
    
    'Load all the menu entries we'll be needing.
    If (m_numOfLanguages > 1) Then
        For i = 1 To m_numOfLanguages - 1
            Load FormMain.mnuLanguages(i)
        Next i
    End If
    
    'Assign names to all the menu entries
    For i = 0 To m_numOfLanguages - 1
        FormMain.mnuLanguages(i).Caption = m_availableLanguages(i).LangName
        FormMain.mnuLanguages(i).Checked = False
    Next i

End Sub

'Given a language file (or partial language file), extract the key language information and place it in the passed variable.
Private Sub FillLanguageInfo(ByRef srcXmlEngine As pdXML, ByRef targetLangHolder As pdLanguageFile)

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

Private Sub GetCachedLanguageInfo(ByRef srcFilename As String, ByRef dstLangHolder As pdLanguageFile)
    With dstLangHolder
        .langID = m_LangCacheXML.GetUniqueTag_String(srcFilename & "-langid")
        .LangName = m_LangCacheXML.GetUniqueTag_String(srcFilename & "-langname")
        .langVersion = m_LangCacheXML.GetUniqueTag_String(srcFilename & "-langversion")
        .LangStatus = m_LangCacheXML.GetUniqueTag_String(srcFilename & "-langstatus")
        .Author = m_LangCacheXML.GetUniqueTag_String(srcFilename & "-langauthor")
    End With
End Sub

'Cache key language information inside the user's preference file.  This allows for faster retrieval on future runs.
Private Sub CacheLanguageInfo(ByRef srcFilename As String, ByRef srcLangHolder As pdLanguageFile)
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
    Dim langTmp As pdLanguageFile
    
    'Loop through all entries in the languages array, sorting them as we go
    For i = 0 To m_numOfLanguages - 1
        For j = 0 To m_numOfLanguages - 1
            
            'Compare two language names, and if one is less (e.g. earlier alphabetically) than the other, swap them
            If (StrComp(m_availableLanguages(i).LangName, m_availableLanguages(j).LangName, vbTextCompare) = -1) Then
                langTmp = m_availableLanguages(i)
                m_availableLanguages(i) = m_availableLanguages(j)
                m_availableLanguages(j) = langTmp
            End If
            
        Next j
    Next i

End Sub

'External functions can use this retrieve the current pdLanguageFile contents for a given language ID.
' Note that this ONLY SUPPORTS OFFICIAL LANGUAGE FILES, by design.  User language files should not be modified by PD itself.
' Returns TRUE if a match is found; FALSE otherwise.
Friend Function GetPDLanguageFileObject(ByRef dstLanguageFileObject As pdLanguageFile, ByVal languageID As String) As Boolean
    
    languageID = Trim$(UCase(languageID))
    
    'Find the matching language ID
    Dim i As Long
    For i = 0 To UBound(m_availableLanguages)
    
        If m_availableLanguages(i).IsOfficial Then
            
            If (StrComp(languageID, Trim$(UCase(m_availableLanguages(i).langID)), vbBinaryCompare) = 0) Then
                
                'It's a match!  Fill the target object and exit
                dstLanguageFileObject = m_availableLanguages(i)
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

Private Sub PostProcessText(ByRef srcString As String)

    '1) Trim the string
    srcString = Trim$(srcString)
    
    '2) If the string is now empty, return a blank string rather than attempting to reconstruct it
    If (Len(srcString) <> 0) Then
    
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
Private Function GetOriginalTagFromTranslation(ByVal curCaption As String) As String

    'If translations aren't active, return the requested caption as-is
    If TranslationActive Then

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
    phraseLocation = m_XMLEngine.GetLocationOfParentTag("phrase", "original", origCaption)
    
    'Make sure a phrase tag was found
    If (phraseLocation > 0) Then
        
        'Retrieve the <translation> tag inside this phrase tag
        origCaption = m_XMLEngine.GetUniqueTag_String("translation", vbNullString, phraseLocation)
        PostProcessText origCaption
        If (Len(origCaption) <> 0) Then
            GetTranslationTagFromCaption = origCaption
        Else
            GetTranslationTagFromCaption = vbNullString
        End If
        
    Else
        GetTranslationTagFromCaption = vbNullString
    End If

End Function

'Given a message (basically, a string), return the translation of said string.
' If a translation cannot be found, return the original string.
Public Function TranslateMessage(ByRef srcMessage As String, ParamArray ExtraText() As Variant) As String
    
    Dim searchString As String, newString As String
    If TranslationActive Then
        searchString = srcMessage
        newString = GetTranslationTagFromCaption(searchString)
    Else
        newString = srcMessage
    End If
    
    'If no translation was found, return the original message as the "translation".  This allows the program to still be usable
    ' even if translation text is missing.
    If (Len(newString) = 0) Then
        
        'If this string contains vbCrLf characters, replace them with plain vbLF characters and try again
        If TranslationActive And (InStr(1, searchString, vbCrLf) > 0) Then
            
            newString = GetTranslationTagFromCaption(Replace$(searchString, vbCrLf, vbLf))
            
            If (Len(newString) = 0) Then
                newString = srcMessage
            Else
            
                'Convert vbLf to vbCrLf, to ensure proper rendering.  Thanks to Frank Donckers for this fix!
                newString = Replace$(newString, vbLf, vbCrLf)
            
            End If
            
        Else
            newString = srcMessage
        End If
        
    End If
    
    'newString now contains the translated message (or the original text, if no translation was found).
    
    'The last step is to replace any optional parameters in the text
    If (UBound(ExtraText) >= LBound(ExtraText)) Then
    
        Dim i As Long
        For i = LBound(ExtraText) To UBound(ExtraText)
            newString = Replace$(newString, "%" & i + 1, ExtraText(i))
        Next i
    
    End If
    
    'Return the fully translated string
    TranslateMessage = newString
    
End Function

'Given a translated string, return the original English version of said string.
'
'Returns: ORIGINAL ENGLISH VERSION of string, if such a thing exists; ORIGINAL STRING otherwise.
Public Function RestoreMessage(ByRef srcMessage As String) As String
    
    Dim searchString As String, newString As String
    searchString = srcMessage
    newString = GetOriginalTagFromTranslation(searchString)
    
    'If a new string was found, apply it to the object's caption
    If (Len(newString) <> 0) Then RestoreMessage = newString Else RestoreMessage = srcMessage
    
End Function

'Given an object with a guaranteed .Caption property (command button, label, etc.), replace the caption with a translation from file
Public Sub TranslateObjectCaption(ByRef srcObject As Control)

    Dim searchString As String, newString As String
    
    If TranslationActive Then
        searchString = srcObject.Caption
        newString = GetTranslationTagFromCaption(searchString)
    Else
        newString = srcObject.Caption
    End If
    
    'If a new string was found, apply it to the object's caption
    If (Len(newString) <> 0) Then
        
        'Form captions look better with a prepended space
        If (TypeOf srcObject Is Form) Then
            srcObject.Caption = " " & newString
        Else
            srcObject.Caption = newString
        End If
    End If

End Sub

'Given an object with a guaranteed .Caption property (command button, label, etc.), replace a translated caption with the original
Public Sub RestoreObjectCaption(ByRef srcObject As Control)

    Dim searchString As String, newString As String
    searchString = srcObject.Caption
    newString = GetOriginalTagFromTranslation(searchString)
    
    'If a new string was found, apply it to the object's caption
    If (Len(newString) <> 0) Then
        
        'Form captions look better with a prepended space
        If (TypeOf srcObject Is Form) Then
            srcObject.Caption = " " & newString
        Else
            srcObject.Caption = newString
        End If
    End If

End Sub

'Restore a form caption
Public Sub RestoreFormCaption(ByRef srcForm As Form)

    Dim searchString As String, newString As String
    searchString = srcForm.Caption
    newString = GetOriginalTagFromTranslation(searchString)
    
    'If a new string was found, apply it to the object's caption
    If (Len(newString) <> 0) Then
        If (Not g_WindowManager Is Nothing) Then
            g_WindowManager.SetWindowCaptionW srcForm.hWnd, " " & newString
        Else
            srcForm.Caption = " " & newString
        End If
    End If
    
End Sub

'Translate a form caption
Public Sub TranslateFormCaption(ByRef srcForm As Form)
    
    If (Len(srcForm.Caption) <> 0) Then
    
        Dim searchString As String, newString As String
        
        If TranslationActive Then
            searchString = srcForm.Caption
            newString = GetTranslationTagFromCaption(searchString)
        Else
            newString = srcForm.Caption
        End If
        
        'If a new string was found, apply it to the object's caption
        If (Len(newString) <> 0) Then
            If (Not g_WindowManager Is Nothing) Then
                g_WindowManager.SetWindowCaptionW srcForm.hWnd, " " & newString
            Else
                srcForm.Caption = " " & newString
            End If
        End If
        
    End If
    
End Sub

'Enumerate all translations on a form (and its controls), and replace them with their original English equivalents
Public Sub UndoTranslations(ByRef srcForm As Form, Optional ByVal useDoEvents As Boolean = False)

    'This function only has to be applied if a translation is currently active.  If it isn't, simply exit.
    If (Not TranslationActive) Then Exit Sub

    'Start by returning the form's original caption
    RestoreFormCaption srcForm
    
    Dim i As Long, strPos As Long
    Dim tmpString As String
    
    'Iterate through every control on the form, and restore original captions as necessary
    Dim eControl As Control
    For Each eControl In srcForm.Controls
        
        'Menus are a special case.  Some PhotoDemon menus have custom accelerators (keyboard shortcuts) that are added at run-time.
        ' We can detect these by checking for a vbTab character in the menu entry.  If this is found, strip it out before searching
        ' for the original text.
        If (TypeOf eControl Is Menu) Then
            tmpString = eControl.Caption
            strPos = InStr(1, tmpString, vbTab, vbBinaryCompare)
            
            'If this menu caption has a custom accelerator appended, strip it off
            If (strPos > 0) Then tmpString = Left$(tmpString, strPos - 1)
            eControl.Caption = RestoreMessage(tmpString)
            
        End If
        
        'Optionally, DoEvents can be called after each change.  This slows the translation process, but it allows external progress
        ' bars to be automatically refreshed.
        If useDoEvents Then DoEvents
    
    Next

End Sub

'Enumerate all captions on a form (and its controls), and replace them with the translated equivalents of the current language
Public Sub ApplyTranslations(ByRef srcForm As Form, Optional ByVal useDoEvents As Boolean = False)
    
    'For en-US locales, we can entirely skip this step
    If (Not TranslationActive) Then Exit Sub
    
    'Start by translating the form's caption
    TranslateFormCaption srcForm

    'Iterate through every control on the form, and restore original captions as necessary
    Dim i As Long, strPos As Long
    Dim tmpString As String
    Dim eControl As Control
    
    For Each eControl In srcForm.Controls
        
        'PD's custom-built controls self-manage translations.  Historically, we asked them to update against the current theme here, but because
        ' the only place this function is called is within the ApplyThemeAndTranslations function (which handles master theming/translation requests
        ' on a per-dialog basis), it's redundant to request a second translation update!
        '
        'As such, PD's custom controls no longer require special attention in this function.
        '
        'Instead, the only controls we must deal with here are standard VB objects with non-blank, non-numeric
        ' caption properties.  PD only has a few of these left, thankfully.
        
        'Menus are a special case.  PhotoDemon menus use custom accelerators (keyboard shortcuts) that are added
        ' at run-time.  We can detect these by checking for a vbTab character in the menu entry.  If a tab is found,
        ' strip it out before searching for a translation.
        If (TypeOf eControl Is Menu) Then
            tmpString = eControl.Caption
            strPos = InStr(1, tmpString, vbTab, vbBinaryCompare)
            
            'If this menu caption has a custom accelerator appended, strip it off
            If (strPos > 0) Then tmpString = Left$(tmpString, strPos - 1)
            eControl.Caption = TranslateMessage(tmpString)
            
            'Note that we do not re-append the accelerator here.  This is handled by a subsequent function
            ' (required because menu icons must be reapplied as well).
            
        End If
        
        'Optionally, DoEvents can be called after each change.  This slows the process, but it allows external progress
        ' bars to be automatically refreshed.  (We use this on the main form, while the language is being changed.)
        If useDoEvents Then DoEvents
        
    Next
    
End Sub

Private Sub Class_Initialize()
    Set m_FSO = New pdFSO
End Sub

Private Sub Class_Terminate()
    Set m_FSO = Nothing
End Sub
