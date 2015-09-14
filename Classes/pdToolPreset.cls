VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdToolPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Individual Tool Preset handler
'Copyright 2013-2015 by Tanner Helland
'Created: 14/August/13
'Last updated: 07/March/15
'Last update: split some parts of preset management out of the command bar and into this standalone class
'Dependencies: pdXML class (for parsing previously saved preset files)
'              pdFSO class (for saving/loading preset files)
'              pdStringStack class (for returning lists of presets)
'
'Supporting presets for every possible tool is an unpleasant task.  This class is designed to make it easier.
'
'Each command bar contains an instance of this class.  All preset management is forwarded to us, and in return, we supply
' preset names and/or values as the command bar needs them.
'
'This class also handles all preset file/save operations.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Each preset supports the following options
Private Type pdPreset
    PresetNameEn As String          'The preset's name, UNTRANSLATED
    PresetNameTranslated As String  'This preset's name, TRANSLATED
    SpecialPresetID As Long         'If this preset is a special, internal PD preset (like Default Settings or Last-Used Settings),
                                    ' this value will be non-zero
    PresetXML As String             'The preset's XML chunk.  This does not include an encoding header or pdDataType; it's just the
                                    ' string containing the preset's various tags.  The parent commandBar object constructs this.
End Type

'Preset collection for this tool
Private m_NumOfPresets As Long
Private m_ListOfPresets() As pdPreset

'The name and (optional) description of this tool.  These values are stored in the preset file, and can be used for further validation.
Private m_ToolName As String, m_ToolDescription As String

'PD stores presets as XML, so we need an XML parser to deal with save/load nuances
Private m_XMLEngine As pdXML

'Since we can also save/load files, we need a pdFSO instance
Private m_File As pdFSO

'Location of the preset file.  If this isn't set, this class is useless.
Private m_PresetFilePath As String

'To improve performance, this class asks the caller to specify a preset name prior to batch retrieval of all preset values.
' This string is set to the preset's XML-safe name, as set by the beginPresetRetrieval function.
Private m_ActivePresetName As String


'The caller must specify a preset path prior to any other interactions.
' Note that this function ALSO LOADS THE PRESET FILE.  (It seemed redundant to force separate handling for that step.)
Public Sub setPresetFilePath(ByRef srcPresetFilePath As String, ByRef toolName As String, Optional ByRef toolDescription As String = "")
    
    m_PresetFilePath = srcPresetFilePath
    m_ToolName = toolName
    m_ToolDescription = toolDescription
    
    If m_File.FileExist(srcPresetFilePath) Then
            
        'Attempt to load and validate the relevant preset file; if we can't, create a new, blank XML object
        If (Not m_XMLEngine.loadXMLFile(srcPresetFilePath)) Or Not (m_XMLEngine.validateLoadedXMLData("toolName")) Then
            debugNote "This tool's preset file may be corrupted.  A new preset file has been created."
            resetXMLData
        End If
        
    Else
        resetXMLData
    End If
    
End Sub

'After the caller is done working with presets, they can call this function to write the current preset collection to file.
Public Sub writePresetFile()
    m_XMLEngine.writeXMLToFile m_PresetFilePath
End Sub

'Given a pdStringStack object, fill it with a list of all available preset names.
'
'RETURNS: number of presets found; 0 if no presets were found, which also happens if the preset file does not exist
Public Function getListOfPresets(ByRef dstStringStack As pdStringStack) As Long
    
    'As a convenience to the caller, initialize the string stack for them
    Set dstStringStack = New pdStringStack
    
    'The XML engine will do most the heavy lifting for this task.  We pass it a String array, and it fills it with
    ' all values corresponding to the given tag name and attribute.
    Dim allPresets() As String
    If m_XMLEngine.findAllAttributeValues(allPresets, "presetEntry", "id") Then
        
        Dim i As Long
        For i = 0 To UBound(allPresets)
            dstStringStack.AddString m_XMLEngine.getUniqueTag_String("fullPresetName", , , "presetEntry", "id", allPresets(i))
        Next i
        
        getListOfPresets = UBound(allPresets) + 1
    
    Else
        getListOfPresets = 0
    End If
    
End Function

'See if a given preset exists in the preset file.
'
'Returns TRUE if the requested preset is found; FALSE otherwise.
Public Function doesPresetExist(ByVal presetName As String) As Boolean

    'Users can name presets however they'd like, but when looking up preset data, we only use XML-friendly tag names
    ' (e.g. no spaces, etc).
    Dim xmlSafePresetName As String
    xmlSafePresetName = m_XMLEngine.getXMLSafeTagName(Trim$(presetName))
    
    'The XML engine handles the actual preset search
    doesPresetExist = m_XMLEngine.doesTagExist("presetEntry", "id", xmlSafePresetName)

End Function

'Prior to retrieving preset values, the caller MUST CALL THIS SUB.  It initializes all relevant XML settings, greatly improving
' performance during preset retrieval.
Public Sub beginPresetRead(ByVal presetName As String)
    
    'Users can name presets however they'd like, but when looking up preset data, we only use XML-friendly tag names
    ' (e.g. no spaces, etc).
    m_ActivePresetName = m_XMLEngine.getXMLSafeTagName(Trim$(presetName))
    
End Sub

'Given two things: an object name and a ByRef destination string, retrieve the preset value for said object,
' for the currently active preset (as set by beginPresetRetrieval), and store it in the destination string.
'
'Returns TRUE if object name exists for preset name; FALSE otherwise.
Public Function readPresetValue(ByVal objectName As String, ByRef dstObjectValue As String) As Boolean
    
    'Perform a failsafe check to make sure preset retrieval has been initialized
    If Len(m_ActivePresetName) <> 0 Then
    
        'Retrieve said value, and return TRUE if a non-zero length string is found
        dstObjectValue = m_XMLEngine.getUniqueTag_String(objectName, "", , "presetEntry", "id", m_ActivePresetName)
        readPresetValue = CBool(Len(dstObjectValue) <> 0)
        
    Else
        
        debugNote "You can't retrieve presets without first calling beginPresetRead!"
        readPresetValue = False
        
    End If
    
End Function

'Counterpart to beginPresetRead, above.  No variables or returns are necessary, as it simply clears the m_ActivePresetName string.
Public Sub endPresetRead()
    m_ActivePresetName = ""
End Sub

'Prior to writing preset values, the caller MUST CALL THIS SUB.  It initializes all relevant XML settings, greatly improving
' performance during preset writes.
Public Sub beginPresetWrite(ByVal presetName As String)
    
    'Users can name presets however they'd like, but when looking up preset data, we only use XML-friendly tag names
    ' (e.g. no spaces, etc).
    m_ActivePresetName = m_XMLEngine.getXMLSafeTagName(Trim$(presetName))
    
    'When writing presets, we need to make sure the parent set of preset tags exists.  If it does, no problem - we'll overwrite
    ' the existing settings.  But if it does NOT exist, we need to create it now.
    If Not m_XMLEngine.doesTagExist("presetEntry", "id", m_ActivePresetName) Then
        
        'Presets use a pretty simple set of tags
        m_XMLEngine.writeTagWithAttribute "presetEntry", "id", m_ActivePresetName, "", True
        
            'Note that we also write the preset name here; this is the exact string we are handed, spaces and all
            m_XMLEngine.writeTag "fullPresetName", presetName
            
        m_XMLEngine.closeTag "presetEntry"
        m_XMLEngine.writeBlankLine
        
    End If
    
End Sub

'Given two things: an object name and a object value, write the two as a standard preset pair into the
' master XML object.
Public Sub writePresetValue(ByRef objectName As String, ByRef objectValue As String)
    
    'Perform a failsafe check to make sure preset writing has been initialized
    If Len(m_ActivePresetName) <> 0 Then
        
        'Store the new value inside the master XML object
        m_XMLEngine.updateTag objectName, objectValue, "presetEntry", "id", m_ActivePresetName
    
    'If preset writing hasn't been initialized, there's nothing we can do
    Else
        debugNote "You can't write presets without first calling beginPresetWrite!"
    End If
    
End Sub

'Counterpart to beginPresetWrite, above.  No variables or returns are necessary, as it simply clears the m_ActivePresetName string.
Public Sub endPresetWrite()
    m_ActivePresetName = ""
End Sub

'Reset the XML engine for this session.  Note that the XML object SHOULD ALREADY BE INSTANTIATED before calling this function.
Private Function resetXMLData()

    m_XMLEngine.prepareNewXML "Tool preset"
    m_XMLEngine.writeBlankLine
    m_XMLEngine.writeTag "toolName", m_ToolName
    m_XMLEngine.writeTag "toolDescription", m_ToolDescription
    m_XMLEngine.writeBlankLine
    m_XMLEngine.writeComment "Everything past this point is tool preset data.  Presets are sorted in the order they were created."
    m_XMLEngine.writeBlankLine

End Function

Private Sub Class_Initialize()
    m_NumOfPresets = 0
    ReDim m_ListOfPresets(0) As pdPreset
    
    Set m_XMLEngine = New pdXML
    Set m_File = New pdFSO
End Sub

Private Sub debugNote(ByVal debugMessage As String)
    #If DEBUGMODE = 1 Then
        If Len(debugMessage) <> 0 Then pdDebug.LogAction "(pdToolPreset debug note) : " & debugMessage
    #End If
End Sub
