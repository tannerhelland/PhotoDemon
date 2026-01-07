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
'Copyright 2013-2026 by Tanner Helland
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Each preset supports the following options
Private Type pdPreset
    presetName As String            'This preset's name, all characters allowed
    PresetNameXML As String         'The preset's name, made XML-safe.  This value is used to actually store the
                                    ' preset inside the XML file.
    SpecialPresetID As Long         'If this preset is a special, internal PD preset (like Default Settings or
                                    ' Last-Used Settings), this value will be non-zero.
    PresetXML As String             'The preset's XML chunk.  This does not include an encoding header or pdDataType;
                                    ' it's just the string containing the preset's various tags.  The parent
                                    ' pdCommandBar object constructs this.
End Type

'Preset collection for this tool
Private m_NumOfPresets As Long, m_ListOfPresets() As pdPreset

'When the "edit preset" window opens, we make a quick backup of all existing preset values.
' If the user cancels the dialog, we'll restore everything back to its original state.
Private m_BackupNumOfPresets As Long, m_BackupListOfPresets() As pdPreset

'The name and (optional) description of this tool.  These values are stored in the preset file, and can be used for further validation.
Private m_ToolName As String, m_ToolDescription As String

'PD stores presets as XML, so we need an XML parser to deal with save/load nuances
Private m_XMLEngine As pdXML

'Location of the preset file.  If this isn't set, this class is useless.
Private m_PresetFilePath As String

'To improve performance, this class asks the caller to specify a preset name prior to batch retrieval of all preset values.
' This string is set to the preset's XML-safe name, as set by the BeginPresetRetrieval function.
Private m_ActivePresetName As String, m_ActivePresetIndex As Long

'The caller must specify a preset path prior to any other interactions.
' Note that this function ALSO LOADS THE PRESET FILE.  (It seemed redundant to force separate handling for that step.)
Friend Sub SetPresetFilePath(ByRef srcPresetFilePath As String, ByRef toolName As String, Optional ByRef toolDescription As String = vbNullString)
    
    m_PresetFilePath = srcPresetFilePath
    m_ToolName = toolName
    m_ToolDescription = toolDescription
    
    'Attempt to load and validate the relevant preset file; if we can't, create a new, blank XML object
    If Files.FileExists(srcPresetFilePath) Then
        If (Not m_XMLEngine.LoadXMLFile(srcPresetFilePath)) Or Not (m_XMLEngine.ValidateLoadedXMLData("toolName")) Then
            DebugNote "This tool's preset file may be corrupted.  A new preset file has been created."
            ResetXMLData
        End If
    Else
        ResetXMLData
    End If
    
    'After loading (or resetting) this tool's XML, retrieve all presets in the file and cache them in a more easily
    ' accessible 1D array.
    
    'Start by pulling all preset names into a string stack
    Dim allPresetNames() As String
    If m_XMLEngine.FindAllAttributeValues(allPresetNames, "presetEntry", "id") Then
        
        m_NumOfPresets = UBound(allPresetNames) + 1
        ReDim m_ListOfPresets(0 To m_NumOfPresets - 1) As pdPreset
        
        Dim i As Long
        For i = 0 To m_NumOfPresets - 1
            With m_ListOfPresets(i)
                
                .PresetNameXML = allPresetNames(i)
                .presetName = m_XMLEngine.GetNonUniqueTag_String("fullPresetName", "presetEntry", "id", .PresetNameXML)
                .PresetXML = Replace$(m_XMLEngine.GetUniqueTag_StringEx("presetEntry", "id", .PresetNameXML), vbCrLf, vbNullString)
                
            End With
        Next i
        
    'No presets exist in this file
    Else
        m_NumOfPresets = 0
    End If
    
End Sub

'After the caller is done working with presets, they can call this function to write the current preset collection to file.
Friend Sub WritePresetFile()
    
    m_XMLEngine.PrepareNewXML "Tool preset"
    m_XMLEngine.WriteBlankLine
    m_XMLEngine.WriteTag "toolName", m_ToolName
    m_XMLEngine.WriteTag "toolDescription", m_ToolDescription
    m_XMLEngine.WriteBlankLine
    m_XMLEngine.WriteComment "Everything past this point is tool preset data."
    m_XMLEngine.WriteBlankLine
    
    If (m_NumOfPresets > 0) Then
        
        'Before writing preset data to file, we want to make sure that the "last-used settings" preset is written
        ' out *first*.  We want it to always appear at the top of the preset list, and to keep dialog initialization
        ' snappy, we don't want to perform sorting at load-time.
        Dim i As Long, lastUsedIndex As Long
        lastUsedIndex = GetPresetIndex("last-used settings")
        
        If (lastUsedIndex >= 0) Then
            m_XMLEngine.WriteTagWithAttribute "presetEntry", "id", m_ListOfPresets(lastUsedIndex).PresetNameXML, vbCrLf & m_ListOfPresets(lastUsedIndex).PresetXML & vbCrLf
            m_XMLEngine.WriteBlankLine
        End If
        
        For i = 0 To m_NumOfPresets - 1
            If (i <> lastUsedIndex) Then
                m_XMLEngine.WriteTagWithAttribute "presetEntry", "id", m_ListOfPresets(i).PresetNameXML, vbCrLf & m_ListOfPresets(i).PresetXML & vbCrLf
                m_XMLEngine.WriteBlankLine
            End If
        Next i
    
    End If
    
    m_XMLEngine.WriteXMLToFile m_PresetFilePath
    
End Sub

'Get the name of the last-created preset, if any
Friend Function GetActivePresetName() As String
    GetActivePresetName = m_ActivePresetName
End Function

'Note that no new presets have been created during this session
Friend Sub ClearActivePresetName()
    m_ActivePresetName = vbNullString
End Sub

'Given a pdStringStack object, fill it with a list of all available preset names.
' *IMPORTANTLY*, this function must be called *after* SetPresetFilePath(), as it will load and parse the XML file.
'
'RETURNS: number of presets found.  Note that 0 is a possible return, and it will be used for misc error cases
'         (e.g. if the preset file does not exist).
Friend Function GetListOfPresets(ByRef dstStringStack As pdStringStack) As Long
    
    'As a convenience to the caller, initialize the string stack for them
    If (dstStringStack Is Nothing) Then Set dstStringStack = New pdStringStack Else dstStringStack.ResetStack
    
    If (m_NumOfPresets > 0) Then
        Dim i As Long
        For i = 0 To m_NumOfPresets - 1
            dstStringStack.AddString m_ListOfPresets(i).presetName
        Next i
    End If
    
    GetListOfPresets = dstStringStack.GetNumOfStrings
    
End Function

'Delete a given preset (identified by actual name - *not* XML-safe name) from the preset list.
'RETURNS: TRUE if the preset was successfully deleted; FALSE otherwise
Friend Function DeletePreset(ByRef srcPresetName As String) As Boolean

    Dim presetIndex As Long
    presetIndex = GetPresetIndex(srcPresetName)
    DeletePreset = (presetIndex >= 0)
    
    If DeletePreset Then
        
        'Remove the selected preset by shifting all presets above it downward
        If (presetIndex < (m_NumOfPresets - 1)) Then
            
            Dim i As Long
            For i = presetIndex To m_NumOfPresets - 2
                m_ListOfPresets(i) = m_ListOfPresets(i + 1)
            Next i
            
        End If
        
        m_NumOfPresets = m_NumOfPresets - 1
        
    End If
    
End Function

'Shift a given preset one position up or down (as specified by the moveUp parameter).
'RETURNS: TRUE if the move was successful; FALSE otherwise.
Friend Function MovePreset(ByRef srcPresetName As String, ByVal moveUp As Boolean) As Boolean

    Dim presetIndex As Long
    presetIndex = GetPresetIndex(srcPresetName)
    MovePreset = (presetIndex >= 0)
    
    If MovePreset Then
        
        'Addtional checks for up vs down
        If moveUp Then
            MovePreset = (presetIndex > 0)
            If MovePreset Then
                Dim tmpPreset As pdPreset
                tmpPreset = m_ListOfPresets(presetIndex - 1)
                m_ListOfPresets(presetIndex - 1) = m_ListOfPresets(presetIndex)
                m_ListOfPresets(presetIndex) = tmpPreset
            End If
        Else
            MovePreset = (presetIndex < m_NumOfPresets - 1)
            If MovePreset Then
                'Dim tmpPreset As pdPreset
                tmpPreset = m_ListOfPresets(presetIndex + 1)
                m_ListOfPresets(presetIndex + 1) = m_ListOfPresets(presetIndex)
                m_ListOfPresets(presetIndex) = tmpPreset
            End If
        End If
        
    End If

End Function

'See if a given preset exists in the preset file.
'
'Returns TRUE if the requested preset is found; FALSE otherwise.
Friend Function DoesPresetExist(ByRef srcPresetName As String) As Boolean
    DoesPresetExist = (GetPresetIndex(srcPresetName) >= 0)
End Function

'Given a preset name, return that preset's index in the preset array.
' RETURNS: some index >= 0, if the preset exists; -1 if the preset does *not* exist
Private Function GetPresetIndex(ByVal srcPresetName As String) As Long
    
    'Abandon the search prematurely if no presets existed in the original XML file
    If (m_NumOfPresets = 0) Then
        GetPresetIndex = -1
        Exit Function
    End If
    
    'Users can name presets however they'd like, but when looking up preset data, we only use XML-friendly
    ' tag names (e.g. no spaces, etc).  Note that internal PD presets (like "last-used settings") may use
    ' a translated name under certain circumstances; we check these, as well.
    srcPresetName = Trim$(srcPresetName)
    Dim srcPresetName2 As String
    If (Not g_Language Is Nothing) Then srcPresetName2 = g_Language.TranslateMessage(srcPresetName) Else srcPresetName2 = srcPresetName
    
    GetPresetIndex = -1
    
    Dim i As Long
    For i = 0 To m_NumOfPresets - 1
        If Strings.StringsEqual(srcPresetName, m_ListOfPresets(i).presetName) Then
            GetPresetIndex = i
            Exit For
        Else
            If Strings.StringsEqual(srcPresetName2, m_ListOfPresets(i).presetName) Then
                GetPresetIndex = i
                Exit For
            End If
        End If
    Next i
    
End Function

'Prior to modifying this preset list, I suggest calling this function.  It will create an internal backup
' of all current XML settings.  These can be restored using the matching function, below.
Friend Sub BackupPresetsInternally()

    m_BackupNumOfPresets = m_NumOfPresets
    Erase m_BackupListOfPresets
    
    If (m_NumOfPresets > 0) Then
        ReDim m_BackupListOfPresets(0 To m_NumOfPresets - 1) As pdPreset
        Dim i As Long
        For i = 0 To m_NumOfPresets - 1
            m_BackupListOfPresets(i) = m_ListOfPresets(i)
        Next i
    End If
    
End Sub

'See the comments for BackupPresetsInternally(), above; this function is its sister.
Friend Sub RestoreBackedUpPresets()
    
    m_NumOfPresets = m_BackupNumOfPresets
    Erase m_ListOfPresets
    
    If (m_NumOfPresets > 0) Then
        ReDim m_ListOfPresets(0 To m_NumOfPresets - 1) As pdPreset
        Dim i As Long
        For i = 0 To m_NumOfPresets - 1
            m_ListOfPresets(i) = m_BackupListOfPresets(i)
        Next i
    End If
    
End Sub

'Return the XML param string for a given preset.
' RETURNS: null string if the source preset doesn't exist.
Friend Function GetPresetXML(ByRef srcPresetName As String) As String
    
    m_ActivePresetIndex = GetPresetIndex(srcPresetName)
    If (m_ActivePresetIndex >= 0) Then
        GetPresetXML = m_ListOfPresets(m_ActivePresetIndex).PresetXML
    Else
        GetPresetXML = vbNullString
    End If
    
End Function

Friend Sub AddPreset(ByRef srcPresetName As String, ByRef newPresetXML As String)

    'First, see if this preset already exists in our list
    Dim presetIndex As Long
    presetIndex = GetPresetIndex(srcPresetName)
    
    'If this preset already exists, we can just overwrite its existing value in-place
    If (presetIndex >= 0) Then
        m_ListOfPresets(presetIndex).PresetXML = newPresetXML
    
    'If this preset does *not* exist, we need to add it to the end of the list
    Else
    
        ReDim Preserve m_ListOfPresets(0 To m_NumOfPresets) As pdPreset
        presetIndex = m_NumOfPresets
        m_NumOfPresets = m_NumOfPresets + 1
        
        With m_ListOfPresets(presetIndex)
            .presetName = srcPresetName
            .PresetNameXML = m_XMLEngine.GetXMLSafeTagName(srcPresetName)
            .PresetXML = newPresetXML
        End With
    
    End If
    
End Sub

'Reset the XML engine for this session.  Note that the XML object SHOULD ALREADY BE INSTANTIATED before calling this function.
Private Sub ResetXMLData()

    m_XMLEngine.PrepareNewXML "Tool preset"
    m_XMLEngine.WriteBlankLine
    m_XMLEngine.WriteTag "toolName", m_ToolName
    m_XMLEngine.WriteTag "toolDescription", m_ToolDescription
    m_XMLEngine.WriteBlankLine
    m_XMLEngine.WriteComment "Everything past this point is tool preset data.  Presets are sorted in the order they were created."
    m_XMLEngine.WriteBlankLine

End Sub

Private Sub Class_Initialize()
    
    m_NumOfPresets = 0
    ReDim m_ListOfPresets(0) As pdPreset
    
    Set m_XMLEngine = New pdXML
    m_XMLEngine.SetTextCompareMode vbBinaryCompare
    
End Sub

Private Sub DebugNote(ByRef debugMessage As String)
    If (LenB(debugMessage) <> 0) Then PDDebug.LogAction "(pdToolPreset debug note) : " & debugMessage
End Sub
