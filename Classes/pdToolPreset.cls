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
'Copyright 2013-2018 by Tanner Helland
'Created: 14/August/13
'Last updated: 07/March/15
'Last update: split some parts of preset management out of the command bar and into this standalone class
'Dependencies: pdXML class (for parsing previously saved preset files)
'              pdFSO class (for saving/loading preset files)
'              pdStringStack class (for returning lists of presets)
'
'Supporting presets for every possible tool is an unpleasant task.  This class is designed to make it easier.
'
'Each command bar contains an instance of this class.  All preset management is forwarded to us, and in return,
' we supply preset names and/or values as the command bar needs them.
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
                                    ' string containing the preset's various tags.  The parent pdCommandBar object constructs this.
End Type

'Preset collection for this tool
Private m_NumOfPresets As Long
Private m_ListOfPresets() As pdPreset

'The name and (optional) description of this tool.  These values are stored in the preset file, and can be used for further validation.
Private m_ToolName As String, m_ToolDescription As String

'PD stores presets as XML, so we need an XML parser to deal with save/load nuances
Private m_XMLEngine As pdXML

'Location of the preset file.  If this isn't set, this class is useless.
Private m_PresetFilePath As String

'To improve performance, this class asks the caller to specify a preset name prior to batch retrieval of all preset values.
' This string is set to the preset's XML-safe name, as set by the beginPresetRetrieval function.
Private m_ActivePresetName As String

'The caller must specify a preset path prior to any other interactions.
' Note that this function ALSO LOADS THE PRESET FILE.  (It seemed redundant to force separate handling for that step.)
Friend Sub SetPresetFilePath(ByRef srcPresetFilePath As String, ByRef toolName As String, Optional ByRef toolDescription As String = vbNullString)
    
    m_PresetFilePath = srcPresetFilePath
    m_ToolName = toolName
    m_ToolDescription = toolDescription
    
    If Files.FileExists(srcPresetFilePath) Then
            
        'Attempt to load and validate the relevant preset file; if we can't, create a new, blank XML object
        If (Not m_XMLEngine.LoadXMLFile(srcPresetFilePath)) Or Not (m_XMLEngine.ValidateLoadedXMLData("toolName")) Then
            DebugNote "This tool's preset file may be corrupted.  A new preset file has been created."
            ResetXMLData
        End If
        
    Else
        ResetXMLData
    End If
    
End Sub

'After the caller is done working with presets, they can call this function to write the current preset collection to file.
Friend Sub WritePresetFile()
    m_XMLEngine.WriteXMLToFile m_PresetFilePath
End Sub

'Given a pdStringStack object, fill it with a list of all available preset names.
'
'RETURNS: number of presets found; 0 if no presets were found, which also happens if the preset file does not exist
Friend Function GetListOfPresets(ByRef dstStringStack As pdStringStack) As Long
    
    'As a convenience to the caller, initialize the string stack for them
    Set dstStringStack = New pdStringStack
    
    'The XML engine will do most the heavy lifting for this task.  We pass it a String array, and it fills it with
    ' all values corresponding to the given tag name and attribute.
    Dim allPresets() As String
    If m_XMLEngine.FindAllAttributeValues(allPresets, "presetEntry", "id") Then
        
        Dim i As Long
        For i = 0 To UBound(allPresets)
            dstStringStack.AddString m_XMLEngine.GetUniqueTag_String("fullPresetName", , , "presetEntry", "id", allPresets(i))
        Next i
        
        GetListOfPresets = UBound(allPresets) + 1
    
    Else
        GetListOfPresets = 0
    End If
    
End Function

'See if a given preset exists in the preset file.
'
'Returns TRUE if the requested preset is found; FALSE otherwise.
Friend Function DoesPresetExist(ByVal presetName As String) As Boolean

    'Users can name presets however they'd like, but when looking up preset data, we only use XML-friendly tag names
    ' (e.g. no spaces, etc).
    Dim xmlSafePresetName As String
    xmlSafePresetName = m_XMLEngine.GetXMLSafeTagName(Trim$(presetName))
    
    'The XML engine handles the actual preset search
    DoesPresetExist = m_XMLEngine.DoesTagExist("presetEntry", "id", xmlSafePresetName)

End Function

'Prior to retrieving preset values, the caller MUST CALL THIS SUB.  It initializes all relevant XML settings, greatly improving
' performance during preset retrieval.
Friend Sub BeginPresetRead(ByVal presetName As String)
    
    'Users can name presets however they'd like, but when looking up preset data, we only use XML-friendly tag names
    ' (e.g. no spaces, etc).
    m_ActivePresetName = m_XMLEngine.GetXMLSafeTagName(Trim$(presetName))
    
End Sub

'Given two things: an object name and a ByRef destination string, retrieve the preset value for said object,
' for the currently active preset (as set by beginPresetRetrieval), and store it in the destination string.
'
'Returns TRUE if object name exists for preset name; FALSE otherwise.
Friend Function ReadPresetValue(ByVal objectName As String, ByRef dstObjectValue As String) As Boolean
    
    'Perform a failsafe check to make sure preset retrieval has been initialized
    If (LenB(m_ActivePresetName) <> 0) Then
    
        'Retrieve said value, and return TRUE if a non-zero length string is found
        dstObjectValue = m_XMLEngine.GetUniqueTag_String(objectName, vbNullString, , "presetEntry", "id", m_ActivePresetName)
        ReadPresetValue = (LenB(dstObjectValue) <> 0)
        
    Else
        
        DebugNote "You can't retrieve presets without first calling beginPresetRead!"
        ReadPresetValue = False
        
    End If
    
End Function

'Counterpart to beginPresetRead, above.  No variables or returns are necessary, as it simply clears the m_ActivePresetName string.
Friend Sub EndPresetRead()
    m_ActivePresetName = vbNullString
End Sub

'Prior to writing preset values, the caller MUST CALL THIS SUB.  It initializes all relevant XML settings, greatly improving
' performance during preset writes.
Friend Sub BeginPresetWrite(ByVal presetName As String)
    
    'Users can name presets however they'd like, but when looking up preset data, we only use XML-friendly tag names
    ' (e.g. no spaces, etc).
    m_ActivePresetName = m_XMLEngine.GetXMLSafeTagName(Trim$(presetName))
    
    'When writing presets, we need to make sure the parent set of preset tags exists.  If it does, no problem - we'll overwrite
    ' the existing settings.  But if it does NOT exist, we need to create it now.
    If (Not m_XMLEngine.DoesTagExist("presetEntry", "id", m_ActivePresetName)) Then
        
        'Presets use a pretty simple set of tags
        m_XMLEngine.WriteTagWithAttribute "presetEntry", "id", m_ActivePresetName, vbNullString, True
        
            'Note that we also write the preset name here; this is the exact string we are handed, spaces and all
            m_XMLEngine.WriteTag "fullPresetName", presetName
            
        m_XMLEngine.CloseTag "presetEntry"
        m_XMLEngine.WriteBlankLine
        
    End If
    
End Sub

'Given two things: an object name and a object value, write the two as a standard preset pair into the
' master XML object.
Friend Sub WritePresetValue(ByRef objectName As String, ByRef objectValue As String)
    
    'Perform a failsafe check to make sure preset writing has been initialized
    If (LenB(m_ActivePresetName) <> 0) Then
        
        'Store the new value inside the master XML object
        m_XMLEngine.UpdateTag objectName, objectValue, "presetEntry", "id", m_ActivePresetName
    
    'If preset writing hasn't been initialized, there's nothing we can do
    Else
        DebugNote "You can't write presets without first calling beginPresetWrite!"
    End If
    
End Sub

'Counterpart to beginPresetWrite, above.  No variables or returns are necessary, as it simply clears the m_ActivePresetName string.
Friend Sub EndPresetWrite()
    m_ActivePresetName = vbNullString
End Sub

'Reset the XML engine for this session.  Note that the XML object SHOULD ALREADY BE INSTANTIATED before calling this function.
Private Function ResetXMLData()

    m_XMLEngine.PrepareNewXML "Tool preset"
    m_XMLEngine.WriteBlankLine
    m_XMLEngine.WriteTag "toolName", m_ToolName
    m_XMLEngine.WriteTag "toolDescription", m_ToolDescription
    m_XMLEngine.WriteBlankLine
    m_XMLEngine.WriteComment "Everything past this point is tool preset data.  Presets are sorted in the order they were created."
    m_XMLEngine.WriteBlankLine

End Function

Private Sub Class_Initialize()
    
    m_NumOfPresets = 0
    ReDim m_ListOfPresets(0) As pdPreset
    
    Set m_XMLEngine = New pdXML
    
End Sub

Private Sub DebugNote(ByVal debugMessage As String)
    #If DEBUGMODE = 1 Then
        If (Len(debugMessage) <> 0) Then pdDebug.LogAction "(pdToolPreset debug note) : " & debugMessage
    #End If
End Sub
