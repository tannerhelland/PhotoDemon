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
'Copyright 2014-2026 by Tanner Helland
'Created: 16/October/14
'Last updated: 06/February/16
'Last update: minor bugfixes
'
'Per its name, this class provides a simple key/value approach to storing data.
'
'Why not use Scripting.Dictionary instead?  Scripting.Dictionary is great, but it's much heavier
' than PD requires, and it's a strict dictionary (meaning there is no notion of order, so you must
' use a For...Each approach to iterate its contents).  As PD dictionary collections tend to be small
' - < 10 items - and we typically use dictionaries to store variable-length data lists that must
' eventually be written to file, it's often preferable to know both the number of items in the
' dictionary, and to also have control over their order, so we can write consistent files.
'
'I also prefer that wrapper functions explicitly state data types - e.g. GetDouble() or
' GetString() - to reduce the potential for Variant-related trouble, especially now that PD is used
' across so many locales.  Similarly, if a requested entry cannot be found in the dictionary, I like
' to allow functions to specify their own default value, which will be returned if an entry cannot
' be located.
'
'Thus this class.  I have not made any attempts to optimize it thus far, as its main use-case is
' for very short, transient lists.
'
'As far as implementation details go, the only really relevant one to outside users is that this
' class treats keys as case-specific (e.g. "key" and "KEY" would be treated as separate entries).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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
Private Const INIT_DICT_SIZE As Long = 16

'If entries have been removed from the dictionary, but the dictionary has not been trimmed, we will auto-initiate a trim
' under certain circumstances.
Private m_TrimNeeded As Boolean

'Number of entries in the current dictionary; 1-based.
' Note that this function will initiate a dictionary trim if items have been removed, as the number of entries will be incorrect otherwise.
Friend Function GetNumOfEntries() As Long
    If m_TrimNeeded Then Me.TrimDeletedEntries
    GetNumOfEntries = m_NumOfEntries
End Function

'Retrieve a particular entry by index.  Helpful when iterating the dictionary.
Friend Function GetValueByIndex(ByVal entryIndex As Long) As Variant
    
    If (entryIndex >= 0) And (entryIndex < m_NumOfEntries) Then
        GetValueByIndex = m_DictEntries(entryIndex).dItem
    Else
        Debug.Print "Invalid dictionary entry request!"
    End If
    
End Function

'Retrieve a particular entry key by index.  Helpful when iterating the dictionary.
Friend Function GetKeyByIndex(ByVal entryIndex As Long) As String
    
    If (entryIndex >= 0) And (entryIndex < m_NumOfEntries) Then
        GetKeyByIndex = m_DictEntries(entryIndex).dKey
    Else
        Debug.Print "Invalid dictionary entry request!"
    End If
    
End Function

'A single function is used to add or update dictionary entries.  If the specified key does not exist, it is added as a
' new entry.  If it does exist, it is simply updated.
'
'The function will return the index of the added location.  Indicies are valid for the life of the pdDictionary object,
' assuming that you do not forcibly remove the key/data pair!
Friend Function AddEntry(ByVal entryKey As String, ByVal entryData As Variant) As Long

    'Keys are case-insensitive
    entryKey = LCase$(entryKey)
    
    'Look for the requested key in the array
    If (m_NumOfEntries > 0) Then
    
        Dim keyFound As Long
        keyFound = -1
        
        Dim i As Long
        For i = 0 To m_NumOfEntries - 1
            
            'Compare the requested key with this one.
            If Strings.StringsEqual(entryKey, m_DictEntries(i).dKey, False) Then
                
                'It's a match!  If this key was scheduled for deletion, un-schedule it prior to returning.
                If m_DictEntries(i).wasDeleted Then m_DictEntries(i).wasDeleted = False
                keyFound = i
                Exit For
                
            End If
        
        Next i
        
        'If the key already exists in our collection, simply update it.
        If (keyFound >= 0) Then
        
            m_DictEntries(keyFound).dItem = entryData
            AddEntry = keyFound
        
        'If the key does not exist in our collection, add it anew, resizing the dictionary as necessary
        Else
        
            If (m_NumOfEntries > UBound(m_DictEntries)) Then ReDim Preserve m_DictEntries(0 To m_NumOfEntries * 2 - 1) As dictEntry
            
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
' (You can force a manual trim by calling TrimDeletedEntries(), and some internal functions may call that as well to ensure
'  proper dictionary behavior.)
'
'Returns: TRUE if item was found and deleted; FALSE otherwise
Friend Function DeleteEntry(ByVal entryKey As String) As Boolean
    
    Dim keyIndex As Long
    keyIndex = GetIndexOfEntry(entryKey)
    
    If (keyIndex >= 0) Then
        m_DictEntries(keyIndex).wasDeleted = True
        DeleteEntry = True
        m_TrimNeeded = True
    Else
        DeleteEntry = False
    End If
    
End Function

'Deleting entries does not actually remove them from the dictionary.  It just flags them for future removal.
' To physically remove those entries, this function must be called.
' (Note that the dictionary is not resized to match the new, potentially lower dictionary size, by design.)
Friend Function TrimDeletedEntries() As Boolean
    
    Dim numValidEntries As Long
    
    If (m_NumOfEntries > 0) Then
    
        Dim i As Long
        For i = 0 To m_NumOfEntries - 1
        
            If (numValidEntries <> i) Then
                m_DictEntries(numValidEntries) = m_DictEntries(i)
            End If
            
            If (Not m_DictEntries(i).wasDeleted) Then numValidEntries = numValidEntries + 1
            
        Next i
    
    End If
    
    m_NumOfEntries = numValidEntries
    m_TrimNeeded = False
    
End Function

'The following set of functions are used to retrieve type-specific copies of data in the dictionary.  If an entry cannot be
' located, the defaultValue param will be supplied instead.
Friend Function GetEntry_Boolean(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As Boolean = False) As Boolean

    Dim keyIndex As Long
    keyIndex = GetIndexOfEntry(entryKey)
    
    If (keyIndex >= 0) Then
        GetEntry_Boolean = CBool(m_DictEntries(keyIndex).dItem)
    Else
        GetEntry_Boolean = defaultValueIfMissing
    End If
    
End Function

Friend Function GetEntry_Byte(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As Byte = 0) As Byte

    Dim keyIndex As Long
    keyIndex = GetIndexOfEntry(entryKey)
    
    If (keyIndex >= 0) Then
        GetEntry_Byte = CByte(m_DictEntries(keyIndex).dItem)
    Else
        GetEntry_Byte = defaultValueIfMissing
    End If

End Function

Friend Function GetEntry_Long(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As Long = 0) As Long

    Dim keyIndex As Long
    keyIndex = GetIndexOfEntry(entryKey)
    
    If (keyIndex >= 0) Then
        GetEntry_Long = CLng(m_DictEntries(keyIndex).dItem)
    Else
        GetEntry_Long = defaultValueIfMissing
    End If
    
End Function

Friend Function GetEntry_Double(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As Double = 0#) As Double

    Dim keyIndex As Long
    keyIndex = GetIndexOfEntry(entryKey)
    
    If (keyIndex >= 0) Then
        GetEntry_Double = CDbl(m_DictEntries(keyIndex).dItem)
    Else
        GetEntry_Double = defaultValueIfMissing
    End If

End Function

Friend Function GetEntry_String(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As String = vbNullString, Optional ByVal assumeLocaleInvariantNumber As Boolean = False) As String

    Dim keyIndex As Long
    keyIndex = GetIndexOfEntry(entryKey)
    
    If (keyIndex >= 0) Then
        
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
            
            If TextSupport.IsNumberLocaleUnaware(testString) Then
                GetEntry_String = Trim$(Str$(Val(testString)))
            Else
                GetEntry_String = testString
            End If
        
        Else
            GetEntry_String = m_DictEntries(keyIndex).dItem
        End If
    Else
        GetEntry_String = defaultValueIfMissing
    End If

End Function

Friend Function GetEntry_Variant(ByVal entryKey As String, Optional ByVal defaultValueIfMissing As Variant = 0) As Variant

    Dim keyIndex As Long
    keyIndex = GetIndexOfEntry(entryKey)
    
    If (keyIndex >= 0) Then
        GetEntry_Variant = m_DictEntries(keyIndex).dItem
    Else
        GetEntry_Variant = defaultValueIfMissing
    End If

End Function

'This function can be used to test for the existence of a given key in the dictionary.
Friend Function DoesKeyExist(ByVal entryKey As String) As Boolean
    DoesKeyExist = (GetIndexOfEntry(entryKey) >= 0)
End Function

'This helper function is used to retrieve dictionary indices, using the requested key.
' If a key cannot be found, this function will return -1.
Private Function GetIndexOfEntry(ByRef entryKey As String) As Long
    
    'Search the dictionary array, looking for a match
    Dim i As Long
    For i = 0 To m_NumOfEntries - 1
        
        'Keys are treated as case-insensitive
        If Strings.StringsEqual(entryKey, m_DictEntries(i).dKey, True) Then
            
            'If we find a match, make sure this entry isn't scheduled for deletion
            If (Not m_DictEntries(i).wasDeleted) Then
                GetIndexOfEntry = i
                Exit Function
            End If
            
            'Note: if an entry *was* scheduled for deletion at some point, we simply ignore it
            
        End If
        
    Next i
    
    'If we made it all the way here, the requested key was not found.  Return -1.
    GetIndexOfEntry = -1

End Function

Private Sub Class_Initialize()
    
    'Always start with an initialized array
    ReDim m_DictEntries(0 To INIT_DICT_SIZE - 1) As dictEntry
    m_NumOfEntries = 0
        
End Sub
