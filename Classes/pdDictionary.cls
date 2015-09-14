VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Dictionary Class
'Copyright 2014-2015 by Tanner Helland
'Created: 16/October/14
'Last updated: 14/September/15
'Last update: minor bugfixes
'
'Per its name, this class provides a simple key/value approach to storing data.
'
'Why not use Scripting.Dictionary instead?  Scripting.Dictionary is great, but it's much heavier than PD requires,
' and it's a strict dictionary (meaning there is no notion of order, so you must use a For...Each approach to iterate
' its contents).  As PD dictionary collections tend to be small - <10 items - and we typically use dictionaries to
' store variable-length data lists that must eventually be written to file, it's often preferable to know both the
' number of items in a dictionary, and to also have control over their order, so we can write consistent files.
'
'I also prefer that wrapper functions explicitly state data types - e.g. GetDouble() or GetString() - to reduce
' the potential for Variant-related trouble, especially now that PD is used across so many locales.  Similarly, if a
' requested entry cannot be found in the dictionary, I like to allow functions to specify their own default value,
' which will be returned if an entry cannot be located.
'
'Thus this class.  (If Scripting.Dictionary did ever become a preferable alternative, it would be easy to modify this
' class to wrap it, so I'll be rewriting all existing Scripting.Dictionary occurrences to use this instead.)
'
'As far as implementation details go, the only really relevant one to outside users is that this class treats keys
' as case-inspecific.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Type dictEntry
    dKey As String
    dItem As Variant
    wasDeleted As Boolean
End Type

Private m_DictEntries() As dictEntry
Private m_NumOfEntries As Long
Private Const INIT_DICT_SIZE = 16

'Number of entries in the current dictionary; 1-based
Public Function getNumOfEntries() As Long
    getNumOfEntries = m_NumOfEntries
End Function

'Retrieve a particular entry by index.  Helpful when iterating the dictionary.
Public Function getValueByIndex(ByVal entryIndex As Long) As Variant
    
    If (entryIndex >= 0) And (entryIndex < m_NumOfEntries) Then
        getValueByIndex = m_DictEntries(entryIndex).dItem
    Else
        Debug.Print "Invalid dictionary entry request!"
    End If
    
End Function

'Retrieve a particular entry key by index.  Helpful when iterating the dictionary.
Public Function getKeyByIndex(ByVal entryIndex As Long) As String
    
    If (entryIndex >= 0) And (entryIndex < m_NumOfEntries) Then
        getKeyByIndex = m_DictEntries(entryIndex).dKey
    Else
        Debug.Print "Invalid dictionary entry request!"
    End If
    
End Function

'A single function is used to add or update dictionary entries.  If the specified key does not exist, it is added as a
' new entry.  If it does exist, it is simply updated.
'
'The function will return the index of the added location.  Indicies are valid for the life of the pdDictionary object,
' assuming that you do not forcibly remove the key/data pair!
Public Function AddEntry(ByVal entryKey As String, ByVal entryData As Variant) As Long

    'Keys are case-insensitive
    entryKey = LCase(entryKey)
    
    'Look for the requested key in the array
    If m_NumOfEntries > 0 Then
    
        Dim keyFound As Long
        keyFound = -1
        
        Dim i As Long
        For i = 0 To m_NumOfEntries - 1
            
            'Compare the requested key with this one.
            If StrComp(entryKey, m_DictEntries(i).dKey, vbBinaryCompare) = 0 Then
                
                'It's a match!  Make sure this key isn't scheduled for deletion.
                ' (For performance reasons, this class does delete entries immediately.  It just marks them for deletion.)
                If Not m_DictEntries(i).wasDeleted Then
                    keyFound = i
                    Exit For
                End If
                
                'Note: if this key was marked for deletion, we ignore it completely.
                
            End If
        
        Next i
        
        'If the key already exists in our collection, simply update it.
        If (keyFound >= 0) Then
        
            m_DictEntries(keyFound).dItem = entryData
            AddEntry = keyFound
        
        'If the key does not exist in our collection, add it anew, resizing the dictionary as necessary
        Else
        
            If m_NumOfEntries > UBound(m_DictEntries) Then ReDim Preserve m_DictEntries(0 To m_NumOfEntries * 2 - 1) As dictEntry
            
            With m_DictEntries(m_NumOfEntries)
                .dKey = entryKey
                .dItem = entryData
            End With
            
            AddEntry = m_NumOfEntries
            m_NumOfEntries = m_NumOfEntries + 1
            
        End If
    
    'For performance reasons, the first entry is handled specially.  (A number of places in PD may store only one item
    ' in a dictionary.)
    Else
        
        m_NumOfEntries = 1
                
        With m_DictEntries(0)
            .dKey = entryKey
            .dItem = entryData
            .wasDeleted = False
        End With
        
        AddEntry = 0
        
    End If
    
End Function

'Delete a dictionary entry.  For performance reasons, this doesn't actually delete the item.  It just marks it for deletion.
' (If I ever hit a point in the future where I'm deleting a ton of entries, I'll add a trim function to actually trim the
' dictionary, but for now, deleted entries remain in the table.)
'
'Returns: TRUE if item was found and deleted; FALSE otherwise
Public Function DeleteEntry(ByVal entryKey As String) As Boolean
    
    Dim keyIndex As Long
    keyIndex = getIndexOfEntry(entryKey)
    
    If keyIndex >= 0 Then
        m_DictEntries(keyIndex).wasDeleted = True
        DeleteEntry = True
    Else
        DeleteEntry = False
    End If
    
End Function

'The following set of functions are used to retrieve type-specific copies of data in the dictionary.  If an entry cannot be
' located, the defaultValue param will be supplied instead.
Public Function getEntry_Boolean(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As Boolean = False) As Boolean

    Dim keyIndex As Long
    keyIndex = getIndexOfEntry(entryKey)
    
    If keyIndex >= 0 Then
        getEntry_Boolean = CBool(m_DictEntries(keyIndex).dItem)
    Else
        getEntry_Boolean = defaultValueIfMissing
    End If
    
End Function

Public Function getEntry_Byte(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As Byte = 0) As Byte

    Dim keyIndex As Long
    keyIndex = getIndexOfEntry(entryKey)
    
    If keyIndex >= 0 Then
        getEntry_Byte = CByte(m_DictEntries(keyIndex).dItem)
    Else
        getEntry_Byte = defaultValueIfMissing
    End If

End Function

Public Function getEntry_Long(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As Long = 0) As Long

    Dim keyIndex As Long
    keyIndex = getIndexOfEntry(entryKey)
    
    If keyIndex >= 0 Then
        getEntry_Long = CLng(m_DictEntries(keyIndex).dItem)
    Else
        getEntry_Long = defaultValueIfMissing
    End If
    
End Function

Public Function getEntry_Double(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As Double = 0#) As Double

    Dim keyIndex As Long
    keyIndex = getIndexOfEntry(entryKey)
    
    If keyIndex >= 0 Then
        getEntry_Double = CDbl(m_DictEntries(keyIndex).dItem)
    Else
        getEntry_Double = defaultValueIfMissing
    End If

End Function

Public Function getEntry_String(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As String = "", Optional ByVal assumeLocaleInvariantNumber As Boolean = False) As String

    Dim keyIndex As Long
    keyIndex = getIndexOfEntry(entryKey)
    
    If keyIndex >= 0 Then
        
        'Some callers may use this function to return a numeric value as a String, e.g. prior to creating a param string.
        ' They can use the assumeLocaleInvariantNumber parameter to notify us of this, and we will translate the key
        ' at this point to a safe, locale-invariant string representation.
        If assumeLocaleInvariantNumber Then
        
            'If the string representation of this key can be coerced into a numeric value, use a (rather ugly) series
            ' of transforms to ensure that the string representation of the number *never* varies by locale.  This is
            ' important as the original string may be locale-specific (especially if it originated from a text box),
            ' but we only want to use locale-invariant versions internally.
            Dim testString As String
            testString = m_DictEntries(keyIndex).dItem
            
            If Text_Support.IsNumberLocaleUnaware(testString) Then
                getEntry_String = Trim$(Str(Val(testString)))
            Else
                getEntry_String = testString
            End If
        
        Else
            getEntry_String = m_DictEntries(keyIndex).dItem
        End If
    Else
        getEntry_String = defaultValueIfMissing
    End If

End Function

Public Function getEntry_Variant(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As Variant = 0) As Variant

    Dim keyIndex As Long
    keyIndex = getIndexOfEntry(entryKey)
    
    If keyIndex >= 0 Then
        getEntry_Variant = m_DictEntries(keyIndex).dItem
    Else
        getEntry_Variant = defaultValueIfMissing
    End If

End Function

'This function can be used to test for the existence of a given key in the dictionary.
Public Function doesKeyExist(ByVal entryKey As String) As Boolean

    If getIndexOfEntry(entryKey) >= 0 Then
        doesKeyExist = True
    Else
        doesKeyExist = False
    End If

End Function

'This helper function is used to retrieve dictionary indices, using the requested key.  If a key cannot be found,
' this function will return -1.
Private Function getIndexOfEntry(ByRef entryKey As String) As Long

    'Keys are treated as case-insensitive
    entryKey = LCase(entryKey)

    'Search the dictionary array, looking for a match
    Dim i As Long
    For i = 0 To m_NumOfEntries - 1
        
        If StrComp(entryKey, m_DictEntries(i).dKey, vbBinaryCompare) = 0 Then
            
            'If we find a match, make sure this entry isn't scheduled for deletion
            If Not m_DictEntries(i).wasDeleted Then
                getIndexOfEntry = i
                Exit Function
            End If
            
            'Note: if an entry *was* scheduled for deletion at some point, we simply ignore it
            
        End If
        
    Next i
    
    'If we made it all the way here, the requested key was not found.  Return -1.
    getIndexOfEntry = -1

End Function

Private Sub Class_Initialize()
    
    'Always start with an initialized array
    ReDim m_DictEntries(0 To INIT_DICT_SIZE - 1) As dictEntry
    m_NumOfEntries = 0
        
End Sub
