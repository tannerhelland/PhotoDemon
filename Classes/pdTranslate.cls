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
'Copyright 2012-2015 by Frank Donckers and Tanner Helland
'Created: 21/December/12
'Last updated: 02/February/15
'Last update: integrate hot-patching language support
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
Private xmlEngine As pdXML

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

Public Sub refreshAsRequired()
    
    If m_ActiveLanguageWasPatched Then
        
        Debug.Print "Patching active language file now..."
        
        m_ActiveLanguageWasPatched = False
        undoTranslations FormMain
        ApplyLanguage False
        
    End If
    
End Sub

'External functions can request the a copy of all found languages; they must provide an array of type pdLanguageFile, which this
' function will then fill with the data assembled from CheckAvailableLanguages().
Friend Sub copyListOfLanguages(ByRef dstListOfLanguages() As pdLanguageFile)
    
    ReDim dstListOfLanguages(0 To numOfLanguages - 1) As pdLanguageFile
    
    'Copy all language data entries into the new array.
    Dim i As Long
    For i = 0 To numOfLanguages - 1
        dstListOfLanguages(i) = availableLanguages(i)
    Next i
    
End Sub

'External functions can request the current language, which may then be used to request translations from
' 3rd-party plugins (specifically, ExifTool).
Public Function getCurrentLanguage() As String
    
    If translationActive Then
        getCurrentLanguage = Left$(availableLanguages(curLanguage).langID, 2)
    Else
        getCurrentLanguage = "en"
    End If
    
End Function

'This function will return PD's internal language index (curLanguage), which matches the language's index in the Language menu.
Public Function getCurrentLanguageIndex() As Long
    getCurrentLanguageIndex = curLanguage
End Function

'External function for asking the translator if a language selection dialog is needed
Public Function isLanguageDialogNeeded(Optional ByRef reasonForDialog As Long = -1) As Boolean
    If reasonForDialog <> -1 Then reasonForDialog = m_reasonForLanguageDialog
    isLanguageDialogNeeded = m_needToShowLanguageDialog
End Function

'If a translation file has been requested and successfully loaded, this function will return TRUE
Public Function readyToTranslate() As Boolean
    
    If Not (xmlEngine Is Nothing) Then
        readyToTranslate = m_isReady
    Else
        readyToTranslate = False
    End If
    
End Function

'If the currently active language is different from the default language, this function will return TRUE
Public Function translationActive() As Boolean

    If curLanguage = defaultLanguage Then
        translationActive = False
    Else
        translationActive = True
    End If

End Function

Public Sub writeLanguagePreferencesToFile()
    g_UserPreferences.SetPref_String "Language", "Current Language File", availableLanguages(curLanguage).FileName
End Sub

'Given an index in the availableLanguages array, mark that as the new language (including saving it to the preferences file)
Public Sub activateNewLanguage(ByVal newLangIndex As Long, Optional ByVal useDoEvents As Boolean = False)

    curLanguage = newLangIndex
    writeLanguagePreferencesToFile
    ApplyLanguage useDoEvents

End Sub

'If the language file updater has updated one or more language files, it will call this function.  Any subsequent language update functions (such as loading a
' new XML file) happen here.
Public Sub notifyHotPatchingComplete()

    'Figure out if the current language is affected by the patch
    If curLanguage <> defaultLanguage Then
    
        'Retrieve the current language version
        Dim curLanguageVersion As String
        curLanguageVersion = xmlEngine.getUniqueTag_String("langversion")
        
        'Retrieve the version of the XML file on disk; if an update has occurred, we can detect it by comparing version numbers.
        Dim tmpXML As pdXML
        Set tmpXML = New pdXML
        If tmpXML.loadXMLFile(availableLanguages(curLanguage).FileName) Then
            
            'Retrieve the version number from the potentially updated file
            Dim newLanguageVersion As String
            newLanguageVersion = tmpXML.getUniqueTag_String("langversion")
            
            If StrComp(curLanguageVersion, newLanguageVersion, vbBinaryCompare) <> 0 Then
            
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
        curLangID = "en-US" 'getDefaultUserLanguage()
                
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
        curLanguageFile = g_UserPreferences.GetPref_String("Language", "Current Language File", "")
        
        'If no specific file has been named, use the default language
        If Len(curLanguageFile) = 0 Then
            curLanguage = defaultLanguage
        
        'If a file HAS been named, make sure it still exists; if it doesn't, default to English (US) and show the language dialog
        Else
            If m_FSO.FileExist(curLanguageFile) Then
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
Public Sub ApplyLanguage(Optional ByVal useDoEvents As Boolean = False)

    'Mark the active language menu entry (and unmark all others)
    Dim i As Long
    For i = 0 To numOfLanguages - 1
        If curLanguage = i Then FormMain.mnuLanguages(i).Checked = True Else FormMain.mnuLanguages(i).Checked = False
    Next i
    
    'If the language is set to anything other than the default, load the contents of the language XML file into memory.
    If translationActive Then
        Set xmlEngine = New pdXML
        xmlEngine.loadXMLFile availableLanguages(curLanguage).FileName
        
        'New as of August '14 is use of binary text comparisons.  This provides a huge performance boost, while also rendering
        ' more faithful text where case is concerned.
        xmlEngine.setTextCompareMode vbBinaryCompare
        
    End If

    'Redraw all text on the main form using the new language
    If FormMain.Visible Then
    
        FormMain.UpdateAgainstCurrentTheme useDoEvents
        DrawAccelerators
        applyAllMenuIcons True
        
        'Resync the interface to redraw any remaining text and/or buttons
        SyncInterfaceToCurrentImage
        
        'Redraw any/all toolbars as well
        toolbar_Toolbox.UpdateAgainstCurrentTheme
        toolbar_Options.UpdateAgainstCurrentTheme
        toolbar_Layers.UpdateAgainstCurrentTheme
        toolbar_ImageTabs.UpdateAgainstCurrentTheme
        
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
Private Function tallyLanguageFilesInFolder(ByVal srcFolder As String, Optional ByVal langOffset As Long = 0, Optional ByVal langTypeForThisFolder As String = "") As Long

    Dim nLanguages As Long
    nLanguages = langOffset
    
    'If this is the official language folder, we will mark those entries specially, as they are the only ones PD is allowed to auto-update.
    Dim isOfficialFolder As Boolean
    If StrComp(UCase(langTypeForThisFolder), "OFFICIAL", vbBinaryCompare) = 0 Then
        isOfficialFolder = True
    Else
        isOfficialFolder = False
    End If
    
    'An XML class will be used to parse the various language files
    Dim tmpXMLEngine As pdXML
    Set tmpXMLEngine = New pdXML
    
    'A pdFSO class is required for Unicode support
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'Scan the translation folder for .xml files.  Ignore anything that isn't XML.
    Dim listOfFiles As pdStringStack
    Set listOfFiles = New pdStringStack
    
    If cFile.RetrieveAllFiles(srcFolder, listOfFiles, False, False, "xml") Then
        
        Dim chkFile As String
        
        Do While listOfFiles.PopString(chkFile)
            
            'Use PD's XML engine to load the file
            If tmpXMLEngine.loadXMLFile(chkFile) Then
            
                'Use the XML engine to validate this file, and to make sure it contains at least a language ID, name, and one (or more) translated phrase
                If tmpXMLEngine.isPDDataType("Translation") And tmpXMLEngine.validateLoadedXMLData("langid", "langname", "phrase") Then
                    
                    'If this is a valid language file, get the language information (if available).
                    fillLanguageInfo tmpXMLEngine, availableLanguages(nLanguages)
                    availableLanguages(nLanguages).FileName = chkFile
                    availableLanguages(nLanguages).langType = langTypeForThisFolder
                    availableLanguages(nLanguages).isOfficial = isOfficialFolder
                    nLanguages = nLanguages + 1
                    
                    'If, for some ungodly reason, the user has more than 100 languages available, dynamically resize the array to fit.
                    If nLanguages > UBound(availableLanguages) Then ReDim Preserve availableLanguages(0 To (nLanguages + 1) * 2) As pdLanguageFile
                
                End If
                
            End If
            
        'Retrieve the next file and repeat
        Loop
        
    End If
    
    tallyLanguageFilesInFolder = nLanguages - langOffset

End Function

'Search the Languages folder, and make a list of all available languages
Public Sub CheckAvailableLanguages()
    
    numOfLanguages = 0
    ReDim Preserve availableLanguages(0 To 99) As pdLanguageFile
    
    'Find all available OFFICIAL language files
    numOfLanguages = tallyLanguageFilesInFolder(g_UserPreferences.getLanguagePath(False), 0, "Official")
    
    'Find all available USER language files
    numOfLanguages = numOfLanguages + tallyLanguageFilesInFolder(g_UserPreferences.getLanguagePath(True), numOfLanguages, "Unofficial")
    
    'Add a dummy entry for PhotoDemon's default language (en-US)
    availableLanguages(numOfLanguages).FileName = ""
    availableLanguages(numOfLanguages).langID = "en-US"
    availableLanguages(numOfLanguages).langName = "English (US) - Default"
    availableLanguages(numOfLanguages).langType = "Default"
    availableLanguages(numOfLanguages).langStatus = "Complete"
    
    'If the MASTER.XML language file is available, retrieve versioning information directly from it.
    Dim masterEnFile As String
    masterEnFile = g_UserPreferences.getLanguagePath() & "Master\MASTER.xml"
        
    If m_FSO.FileExist(masterEnFile) Then
        
        'Retrieve the version directly from file
        Dim tmpXML As pdXML
        Set tmpXML = New pdXML
        If tmpXML.loadXMLFile(masterEnFile) Then
            
            Dim curMasterVersion As String
            curMasterVersion = tmpXML.getUniqueTag_String("langversion")
            availableLanguages(numOfLanguages).langVersion = curMasterVersion
            
        Else
            availableLanguages(numOfLanguages).langVersion = App.Major & "." & App.Minor & "." & App.Revision
        End If
        
    Else
        availableLanguages(numOfLanguages).langVersion = App.Major & "." & App.Minor & "." & App.Revision
    End If
    
    availableLanguages(numOfLanguages).isOfficial = True
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
        If StrComp(UCase$(availableLanguages(i).langName), "ENGLISH (US) - DEFAULT", vbBinaryCompare) = 0 Then defaultLanguage = i
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
Private Sub fillLanguageInfo(ByRef srcXMLEngine As pdXML, ByRef targetLangHolder As pdLanguageFile)

    'First, get the language ID and name - these are the most important values, and technically the only REQUIRED ones.
    With targetLangHolder
        
        .langID = srcXMLEngine.getUniqueTag_String("langid")
        .langName = srcXMLEngine.getUniqueTag_String("langname")
    
    'Version, status, and author information should also be present, but the file will still be loaded even if they don't exist
        .langVersion = srcXMLEngine.getUniqueTag_String("langversion")
        .langStatus = srcXMLEngine.getUniqueTag_String("langstatus")
        .Author = srcXMLEngine.getUniqueTag_String("author")
        
    End With
        
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
    langTmp = langOne
    langOne = langTwo
    langTwo = langTmp
End Sub

'External functions can use this retrieve the current pdLanguageFile contents for a given language ID.
' Note that this ONLY SUPPORTS OFFICIAL LANGUAGE FILES, by design.  User language files should not be modified by PD itself.
' Returns TRUE if a match is found; FALSE otherwise.
Friend Function getPDLanguageFileObject(ByRef dstLanguageFileObject As pdLanguageFile, ByVal languageID As String) As Boolean
    
    languageID = Trim$(UCase(languageID))
    
    'Find the matching language ID
    Dim i As Long
    For i = 0 To UBound(availableLanguages)
    
        If availableLanguages(i).isOfficial Then
            
            If StrComp(languageID, Trim$(UCase(availableLanguages(i).langID)), vbBinaryCompare) = 0 Then
                
                'It's a match!  Fill the target object and exit
                dstLanguageFileObject = availableLanguages(i)
                getPDLanguageFileObject = True
                Exit Function
            
            End If
        
        End If
    
    Next i
    
    'If we made it here, no matching language was found
    getPDLanguageFileObject = False

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

    'If translations aren't active, return the requested caption as-is
    If translationActive Then

        'Remove white space from the caption (if necessary, white space will be added back in after retrieving the translation from file)
        preProcessText curCaption
        
        Dim phraseLocation As Long
        phraseLocation = xmlEngine.getLocationOfParentTag("phrase", "translation", curCaption)
        
        'Make sure a phrase tag was found
        If phraseLocation > 0 Then
            
            'Retrieve the <translation> tag inside this phrase tag
            curCaption = xmlEngine.getUniqueTag_String("original", "", phraseLocation)
            
        End If
        
        'Apply any relevant post-processing, then return the text
        postProcessText curCaption
        getOriginalTagFromTranslation = curCaption
        
    Else
        getOriginalTagFromTranslation = curCaption
    End If

End Function

'Given the original caption of a message or control, return the matching translation from the active translation file
Private Function getTranslationTagFromCaption(ByVal origCaption As String) As String

    If xmlEngine Is Nothing Then
        getTranslationTagFromCaption = origCaption
        Exit Function
    End If

    'Remove white space from the caption (if necessary, white space will be added back in after retrieving the translation from file)
    preProcessText origCaption
    
    Dim phraseLocation As Long
    phraseLocation = xmlEngine.getLocationOfParentTag("phrase", "original", origCaption)
    
    'Make sure a phrase tag was found
    If phraseLocation > 0 Then
        
        'Retrieve the <translation> tag inside this phrase tag
        origCaption = xmlEngine.getUniqueTag_String("translation", "", phraseLocation)
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
    If Len(newString) = 0 Then
        
        'If this string contains vbCrLf characters, replace them with plain vbLF characters and try again
        If translationActive And (InStr(1, searchString, vbCrLf) > 0) Then
            
            newString = getTranslationTagFromCaption(Replace$(searchString, vbCrLf, vbLf))
            
            If Len(newString) = 0 Then
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
    If UBound(ExtraText) >= LBound(ExtraText) Then
    
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
    If Len(newString) <> 0 Then
        If Not (g_WindowManager Is Nothing) Then
            g_WindowManager.SetWindowCaptionW srcForm.hWnd, " " & newString
        Else
            srcForm.Caption = " " & newString
        End If
    End If
    
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
    If Len(newString) <> 0 Then
        If Not (g_WindowManager Is Nothing) Then
            g_WindowManager.SetWindowCaptionW srcForm.hWnd, " " & newString
        Else
            srcForm.Caption = " " & newString
        End If
    End If
    
End Sub

'Enumerate all translations on a form (and its controls), and replace them with their original English equivalents
Public Sub undoTranslations(ByRef srcForm As Form, Optional ByVal useDoEvents As Boolean = False)

    'This function only has to be applied if a translation is currently active.  If it isn't, simply exit.
    If Not translationActive Then Exit Sub

    'Start by returning the form's original caption
    restoreFormCaption srcForm
    
    Dim i As Long, strPos As Long
    Dim tmpString As String
    
    'Iterate through every control on the form, and restore original captions as necessary
    Dim eControl As Control
    For Each eControl In srcForm.Controls
    
        'If this is an object with a standard .Caption property, like a label or command button, replacing the text is easy.
        If (TypeOf eControl Is CommandButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is Label) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is pdLabel) Then
            If (Trim(eControl.Caption) <> "") And (Not IsNumeric(eControl.Caption)) Then g_Language.restoreObjectCaption eControl
            If (Trim(eControl.ToolTipText) <> "") Then eControl.ToolTipText = RestoreMessage(eControl.ToolTipText)
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
        
        'Optionally, DoEvents can be called after each change.  This slows the translation process, but it allows external progress
        ' bars to be automatically refreshed.
        If useDoEvents Then DoEvents
    
    Next

End Sub

'Enumerate all captions on a form (and its controls), and replace them with the translated equivalents of the current language
Public Sub applyTranslations(ByRef srcForm As Form, Optional ByVal useDoEvents As Boolean = False)

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
        If (TypeOf eControl Is CommandButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is Label) Or (TypeOf eControl Is OptionButton) Then
            
            'TESTING ONLY
            'If srcForm.Name <> "FormMain" Then MsgBox eControl.Caption
            
            If (Len(Trim(eControl.Caption)) <> 0) And (Not IsNumeric(eControl.Caption)) Then translateObjectCaption eControl
            If (Len(Trim(eControl.ToolTipText)) <> 0) Then eControl.ToolTipText = TranslateMessage(eControl.ToolTipText)
            
        End If
                        
        'Picture boxes and text boxes can have tooltips.  Translate those too.
        If (TypeOf eControl Is PictureBox) Or (TypeOf eControl Is TextBox) Then
            If (Len(Trim$(eControl.ToolTipText)) <> 0) Then eControl.ToolTipText = TranslateMessage(eControl.ToolTipText)
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
            
            'Enumerate through each entry in the control (if any are present!)
            If eControl.ListCount > 0 Then
            
                For i = 0 To eControl.ListCount - 1
                    If (Len(Trim$(eControl.List(i))) <> 0) And (Not IsNumeric(eControl.List(i))) Then eControl.List(i) = TranslateMessage(eControl.List(i))
                Next i
                
            End If
            
            'Translate the tooltip as well
            If (Len(Trim$(eControl.ToolTipText)) <> 0) Then eControl.ToolTipText = TranslateMessage(eControl.ToolTipText)
            
        End If
        
        'PD's custom-built controls self-manage translations.  Historically, we asked them to update against the current theme here, but because
        ' the only place this function is called is within the makeFormPretty function (which handles master theming/translation requests
        ' on a per-dialog basis), it's redundant to request a second translation update!
        '
        'As such, PD's custom controls no longer require special attention in this function.
        
        Dim curFontSize As Long
        
        'Old common control labels receive a bit of extra attention.  Because translations can be longer than their original English text,
        ' a translated label may extend beyond the edge of its container.  If this happens, attempt to shrink the label's font size until
        ' it fits.  (This is only necessary until all labels are switched over to the new pdLabel object.)
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
        
        'Optionally, DoEvents can be called after each change.  This slows the process, but it allows external progress
        ' bars to be automatically refreshed.
        If useDoEvents Then DoEvents
        
    Next
    
    'If g_DisplayTimingReports Then MsgBox "Translation of " & srcForm.Name & " took " & (Timer - localTimeTaken) & " seconds."
    
End Sub

Private Sub Class_Initialize()
    Set m_FSO = New pdFSO
End Sub

Private Sub Class_Terminate()
    Set m_FSO = Nothing
End Sub
