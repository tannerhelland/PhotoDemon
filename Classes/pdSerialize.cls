VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdSerialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Internal Parameter Handler v2
'Copyright 2013-2026 by Tanner Helland
'Created: 25/March/13
'Last updated: 28/February/20
'Last update: make .AddParam() smarter - we know numeric datatypes are already XML-safe, so don't bother
'             manually delimiting them.
'
'PhotoDemon has unique needs regarding parameter passing.  Because it allows the user to record all
' actions as part of macros, it also requires a way to standardize parameter lists (regardless of
' count or type), and this standardization also needs to be easy to read/write from files.
'
'When first developing a solution to this, the biggest complication was finding a single system that
' worked with all possible function parameters.  Some actions, like "Invert Image", require no
' additional information.  Others, like "Curves", require a huge amount of custom data.  Various
' functions utilize conceivable every type of value (bool, int, float, custom types, etc), and if a
' macro's being recorded, all those specialized parameters need to be tracked and saved to file.
'
'The easiest way to handle that kind of variety from within VB is to use a string.  This allows any
' amount - and type - of custom data to be recorded, stored, and easily transferred between functions,
' provided that string serializers exist for all data types (obviously).
'
'Individual functions therefore rely on this class to create and parse parameter strings for them.
'
'In summer 2015, I moved from a pipe-delimited parameter system (where parameters were stored/retrieved
' by ordinal) to an order-agnostic XML system, where parameters are retrieved by name.  This class uses
' its own lightweight, custom-built XML parser, which is designed purely against the needs of this class.
' (For example, it doesn't make any attempt to construct a DOM - it just retrieves tags in order as fast
' as it possibly can.)  Specifically, the micro-XML parser used here requires these restrictions:
' 1) All comparisons are case-sensitive.  If you change case in PD's source, you will invalidate old
'    parameters, by design.
' 2) All parameter names must be unique.
' 3) Parameter names do not support attribute tags.  Instead, use different parameter names.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit
Option Compare Binary

'Current parameter string, stored in a pdString object (which supports fast string-building).
Private m_ParamString As pdString

'To improve performance, we search for tags starting from the last-retrieved tag position.  In 99% of cases, PD functions
' retrieve parameters in-order, so we can save time by moving through the XML sequentially.
Private m_SearchStart As Long

Friend Function GetParamString() As String
    GetParamString = m_ParamString.ToString()
End Function

'If you obtain a parameter string from elsewhere (e.g. a PDI file), you can simple pass it as-is to this function
Friend Sub SetParamString(ByRef srcString As String)
    m_ParamString.Reset srcString
End Sub

Friend Sub Reset(Optional ByVal paramStringVersion As Double = 1#)
    m_ParamString.Reset "<version>" & Trim$(Str$(paramStringVersion)) & "</version>"
    m_SearchStart = 1
End Sub

Private Sub Class_Initialize()
    
    Set m_ParamString = New pdString
    
    'For now, this class does not add any extra XML around the parameters themselves.
    ' It will, however, create a generic "version" tag at version 1.0.  Functions are free to override this as necessary.
    Me.Reset
    
End Sub

'If a function changes its parameter structure, it can use this "created-by-default" version tag
' to adjust handling as necessary.
Friend Function GetParamVersion() As Double
    
    Dim versionString As String
    If GetParamValue("version", versionString, True, 1) Then
        GetParamVersion = CDblCustom(versionString)
    Else
        GetParamVersion = 0#
    End If
    
End Function

Friend Sub SetParamVersion(Optional ByVal newVersion As Double = 1#)
    Me.UpdateParam "version", newVersion, True
End Sub

'Simple check to see if a parameter exists
Friend Function DoesParamExist(ByVal paramName As String, Optional ByVal nameGuaranteedXMLSafe As Boolean = False) As Boolean
    
    'Make the parameter name XML-safe
    If (Not nameGuaranteedXMLSafe) Then paramName = GetXMLSafeName(paramName)
    
    Dim paramPos As Long
    paramPos = m_ParamString.StrStr("<" & paramName & ">")
    DoesParamExist = (paramPos <> 0)
    
    'If this parameter *does* exist, update our search pointer to point at it (as it's likely the caller will try to retrieve
    ' it immediately following this check!)
    If DoesParamExist Then m_SearchStart = paramPos
    
End Function

'Retrieve various type-specific parameters.  Note that these ultimately wrap getParamValue; they simply cast the result explicitly.
Friend Function GetBool(ByRef paramName As String, Optional ByVal defaultReturn As Boolean = False, Optional ByVal nameGuaranteedXMLSafe As Boolean = False) As Boolean
    Dim paramValue As String
    If GetParamValue(paramName, paramValue, nameGuaranteedXMLSafe) Then
        GetBool = CBool(Trim$(paramValue))
    Else
        GetBool = defaultReturn
    End If
End Function

Friend Function GetByte(ByRef paramName As String, Optional ByVal defaultReturn As Byte = 0, Optional ByVal nameGuaranteedXMLSafe As Boolean = False) As Byte
    Dim paramValue As String
    If GetParamValue(paramName, paramValue, nameGuaranteedXMLSafe) Then
        GetByte = CByte(Trim$(paramValue))
    Else
        GetByte = defaultReturn
    End If
End Function

Friend Function GetInteger(ByRef paramName As String, Optional ByVal defaultReturn As Integer = 0, Optional ByVal nameGuaranteedXMLSafe As Boolean = False) As Integer
    Dim paramValue As String
    If GetParamValue(paramName, paramValue, nameGuaranteedXMLSafe) Then
        GetInteger = CInt(CDblCustom(Trim$(paramValue)))
    Else
        GetInteger = defaultReturn
    End If
End Function

Friend Function GetLong(ByRef paramName As String, Optional ByVal defaultReturn As Long = 0, Optional ByVal nameGuaranteedXMLSafe As Boolean = False) As Long
    Dim paramValue As String
    If GetParamValue(paramName, paramValue, nameGuaranteedXMLSafe) Then
        GetLong = Int(CDblCustom(Trim$(paramValue)) + 0.5)
    Else
        GetLong = defaultReturn
    End If
End Function

Friend Function GetSingle(ByRef paramName As String, Optional ByVal defaultReturn As Single = 0!, Optional ByVal nameGuaranteedXMLSafe As Boolean = False) As Single
    Dim paramValue As String
    If GetParamValue(paramName, paramValue, nameGuaranteedXMLSafe) Then
        GetSingle = CSngCustom(Trim$(paramValue))
    Else
        GetSingle = defaultReturn
    End If
End Function

Friend Function GetDouble(ByRef paramName As String, Optional ByVal defaultReturn As Double = 0#, Optional ByVal nameGuaranteedXMLSafe As Boolean = False) As Double
    Dim paramValue As String
    If GetParamValue(paramName, paramValue, nameGuaranteedXMLSafe) Then
        GetDouble = CDblCustom(Trim$(paramValue))
    Else
        GetDouble = defaultReturn
    End If
End Function

Friend Function GetString(ByRef paramName As String, Optional ByVal defaultReturn As String = vbNullString, Optional ByVal nameGuaranteedXMLSafe As Boolean = False) As String
    If (Not GetParamValue(paramName, GetString, nameGuaranteedXMLSafe)) Then GetString = defaultReturn
End Function

'If you want to combine "DoesParamExist" and "GetString", use this function
Friend Function GetStringEx(ByRef paramName As String, ByRef dstParamValue As String, Optional ByVal nameGuaranteedXMLSafe As Boolean = False, Optional ByVal searchStart As Long = -1) As Boolean
    GetStringEx = GetParamValue(paramName, dstParamValue, nameGuaranteedXMLSafe, searchStart)
End Function

Friend Function GetVariant(ByRef paramName As String, Optional ByVal defaultReturn As Variant = Empty) As Variant
    Dim paramValue As String
    If GetParamValue(paramName, paramValue) Then
        GetVariant = CVar(paramValue)
    Else
        GetVariant = defaultReturn
    End If
End Function

'Given a parameter name, fill a user-supplied string with the parameter value.
' Returns TRUE if parameter exists; FALSE otherwise.
Private Function GetParamValue(ByVal paramName As String, ByRef dstString As String, Optional ByVal nameGuaranteedXMLSafe As Boolean = False, Optional ByVal searchStart As Long = -1) As Boolean
    
    GetParamValue = False
    
    'Make the parameter name XML-safe
    If (Not nameGuaranteedXMLSafe) Then paramName = GetXMLSafeName(paramName)
    
    'If the caller supplied a custom search position, start searching there (under the assumption
    ' that they know something about the param string that we do not).
    '
    'Otherwise, start searching from wherever we found our previous tag.  This *greatly* accelerates
    ' parse time because PD tags are pretty much always read in sequential order.
    Dim tagStart As Long, tagEnd As Long
    If (searchStart < 1) Then
        If (m_SearchStart > m_ParamString.GetLength()) Then m_SearchStart = 1
    Else
        m_SearchStart = searchStart
    End If
    tagStart = m_ParamString.StrStr("<" & paramName & ">", m_SearchStart)
    
    'If we didn't find anything, and we didn't start our search at the front of the string, try a full-string search now
    If (tagStart = 0) And (m_SearchStart > 1) Then
        m_SearchStart = 1
        tagStart = m_ParamString.StrStr("<" & paramName & ">", m_SearchStart)
    End If
    
    'If the opening tag was found, we also need to find the closing tag.
    If (tagStart > 0) Then
    
        'Increment the tag start location by the length of the tag plus two (+1 for each bracket: <>)
        tagStart = tagStart + Len(paramName) + 2
            
        tagEnd = m_ParamString.StrStr("</" & paramName & ">", tagStart)
        
        'If the closing tag exists, return everything between that and the opening tag
        If (tagEnd > tagStart) Then
            
            'Advance our "search start" position to the end of this tag
            m_SearchStart = tagEnd + Len(paramName) + 3
            
            'Extract and return the parameter at this location
            dstString = m_ParamString.MidW(tagStart, tagEnd - tagStart)
            dstString = UnDelimitParamValue(dstString)
            GetParamValue = True
            
        Else
            
            'If the closing tag exists, but it *equals* tag start, it means this tag exists but is empty (which is fine)
            If (tagEnd = tagStart) Then
                dstString = vbNullString
            
            'If tag end is less than tag start, it means there's an XML parsing problem
            Else
                PDDebug.LogAction "pdSerialize.GetParamValue failed.  Param string follows."
                PDDebug.LogAction m_ParamString.ToString()
                dstString = vbNullString
            End If
            
        End If
    
    'When we look for a tag and don't find it, reset our search start position
    Else
        m_SearchStart = 1
        dstString = vbNullString
    End If

End Function

'Blindly add a parameter to the central string.
' No special checks (e.g. duplicates) are applied; use UpdateParam() if you need those.
Friend Function AddParam(ByVal paramName As String, ByVal paramValue As Variant, Optional ByVal nameGuaranteedXMLSafe As Boolean = False, Optional ByVal valueGuaranteedXMLSafe As Boolean = False) As Boolean
    
    On Error GoTo AddParamFailed
    
    'Convert the parameter value into a string.  We handle this manually to minimize the chance of locale issues.
    Dim strParamValue As String
    
    If (VarType(paramValue) = vbString) Then
        strParamValue = paramValue
    ElseIf (VarType(paramValue) = vbByte) Then
        strParamValue = Trim$(Str$(paramValue))
        valueGuaranteedXMLSafe = True
    ElseIf (VarType(paramValue) = vbInteger) Then
        strParamValue = Trim$(Str$(paramValue))
        valueGuaranteedXMLSafe = True
    ElseIf (VarType(paramValue) = vbLong) Then
        strParamValue = Trim$(Str$(paramValue))
        valueGuaranteedXMLSafe = True
    ElseIf (VarType(paramValue) = vbSingle) Then
        strParamValue = Trim$(Str$(paramValue))
        valueGuaranteedXMLSafe = True
    ElseIf (VarType(paramValue) = vbDouble) Then
        strParamValue = Trim$(Str$(paramValue))
        valueGuaranteedXMLSafe = True
    ElseIf (VarType(paramValue) = vbBoolean) Then
        strParamValue = Trim$(Str$(paramValue))
        valueGuaranteedXMLSafe = True
    ElseIf (VarType(paramValue) = vbDecimal) Then
        strParamValue = Trim$(Str$(paramValue))
        valueGuaranteedXMLSafe = True
    ElseIf (VarType(paramValue) = vbCurrency) Then
        strParamValue = Trim$(Str$(paramValue))
    ElseIf (VarType(paramValue) = vbNull) Then
        strParamValue = Trim$(Str$(0))
        valueGuaranteedXMLSafe = True
    ElseIf (VarType(paramValue) = vbDate) Then
        strParamValue = Format$(paramValue, "yyyy-mm-dd h:mm:ss", vbSunday, vbFirstJan1)
    
    'Pray for a correct implicit cast result
    Else
        strParamValue = paramValue
    End If
    
    'Make the parameter name and value XML-safe
    If (LenB(paramName) <> 0) And (Not nameGuaranteedXMLSafe) Then paramName = Me.GetXMLSafeName(paramName)
    If (LenB(strParamValue) <> 0) And (Not valueGuaranteedXMLSafe) Then strParamValue = DelimitParamValue(strParamValue)
    
    'Tack the parameter name and value we were passed onto the central string
    m_ParamString.XMLAppend paramName, strParamValue
    
    AddParam = True
    Exit Function
    
AddParamFailed:
    AddParam = False
    PDDebug.LogAction "WARNING!  pdSerialize failed to add param named " & paramName
    
End Function

'Want to add multiple parameters at once?  Use this function, and call it using the order "Name", Value, "Name 2", Value 2, etc.
Friend Function AddParamList(ParamArray allParams() As Variant) As Boolean
    
    On Error GoTo MultiParamFailure
    
    If (UBound(allParams) >= LBound(allParams)) Then
    
        Dim tmpName As String, tmpValue As Variant
        
        Dim i As Long
        For i = LBound(allParams) To UBound(allParams) Step 2
            
            'Parameters must be passed in a strict name/value order.  An odd number of parameters will cause crashes.
            tmpName = allParams(i)
            
            If (i + 1) <= UBound(allParams) Then
                tmpValue = allParams(i + 1)
            Else
                Err.Raise 9
            End If
            
            'Add this key/value pair to the current running param string
            Me.AddParam tmpName, tmpValue
            
        Next i
    
    End If
    
    AddParamList = True
    
    Exit Function
    
MultiParamFailure:
    PDDebug.LogAction "WARNING!  Parse failure in pdSerialize.AddParamList()."
    AddParamList = False

End Function

'Add an existing XML string to this one.  The incoming string is not modified in any way, so please make sure it contains
' valid XML!
Friend Sub AddXMLString(ByRef srcString As String)
    m_ParamString.Append srcString
End Sub

'Update a given parameter.  If the parameter is not found, it will be added to the string. (Create-if-missing behavior can be toggled.)
Friend Function UpdateParam(ByVal paramName As String, ByVal paramValue As String, Optional ByVal createIfMissing As Boolean = True) As Boolean
    
    'Make the parameter name and value XML-safe
    paramName = GetXMLSafeName(paramName)
    paramValue = DelimitParamValue(paramValue)
    
    'pdString can perform the update for us; note that it will return FALSE if the parameter does not exist
    UpdateParam = m_ParamString.XMLUpdateTagValue(paramName, paramValue)
    If (Not UpdateParam) And createIfMissing Then
    
        'The requested tag doesn't exist.  Append it instead of updating it.
        m_ParamString.XMLAppend paramName, paramValue
        UpdateParam = True
        
    End If
    
End Function

'Given a string, replace any characters that are not allowed with underscores;
' this is used as a failsafe when adding new parameters to the central string.
Friend Function GetXMLSafeName(ByRef srcString As String) As String
    
    'Remove any incidental white space before processing
    GetXMLSafeName = Trim$(srcString)
    
    'Create a string of valid numerical characters, based on the XML spec at http://www.w3.org/TR/1998/REC-xml-19980210.html#sec-common-syn
    Const validChars As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-_:"
    Const underscoreChar As String = "_"
    
    'Loop through the source string and replace any invalid characters with underscore
    Dim i As Long, newString As pdString
    Set newString = New pdString
    
    For i = 1 To Len(GetXMLSafeName)
        If (InStr(1, validChars, Mid$(GetXMLSafeName, i, 1)) = 0) Then
            newString.Append underscoreChar
        Else
            newString.Append Mid$(GetXMLSafeName, i, 1)
        End If
    Next i
    
    GetXMLSafeName = newString.ToString()
    
End Function

'Given a parameter value string, make it XML-safe (e.g. replace "&", "<", ">" with HTML equivalents).
' Note that we don't currently delimit other chars because PD uses only a subset of XML functionality, by design.
Private Function DelimitParamValue(ByVal srcString As String) As String
    DelimitParamValue = srcString
    If (InStr(1, srcString, "&") <> 0) Then DelimitParamValue = Replace$(DelimitParamValue, "&", "&amp;")
    If (InStr(1, srcString, "<") <> 0) Then DelimitParamValue = Replace$(DelimitParamValue, "<", "&lt;")
    If (InStr(1, srcString, ">") <> 0) Then DelimitParamValue = Replace$(DelimitParamValue, ">", "&gt;")
End Function

'When un-delimiting param values, it's important to do it in the *opposite* order in which we initially
' delimited them.  (This ensures that ampersands in escaped entities are restored correctly.)
Private Function UnDelimitParamValue(ByRef srcString As String) As String
    UnDelimitParamValue = srcString
    If (InStr(1, srcString, "&gt;") <> 0) Then UnDelimitParamValue = Replace$(UnDelimitParamValue, "&gt;", ">")
    If (InStr(1, srcString, "&lt;") <> 0) Then UnDelimitParamValue = Replace$(UnDelimitParamValue, "&lt;", "<")
    If (InStr(1, srcString, "&amp;") <> 0) Then UnDelimitParamValue = Replace$(UnDelimitParamValue, "&amp;", "&")
End Function

'A custom CDbl function that accepts both commas and decimals as a separator; this
' is relevant when moving floating-point data, represented as strings, between locales.
' (PhotoDemon no longer uses local-dependent string conversions, but ancient pre-Unicode-support
' versions did, so this exists primarily for backward-compatibility.)
Private Function CDblCustom(ByRef srcString As String) As Double
    If (InStr(1, srcString, ",") = 0) Then
        If TextSupport.IsNumberLocaleUnaware(srcString) Then CDblCustom = Val(srcString) Else CDblCustom = 0#
    Else
        CDblCustom = FixLocaleDependentNumber(srcString)
    End If
End Function

Private Function CSngCustom(ByRef srcString As String) As Single
    If (InStr(1, srcString, ",") = 0) Then
        If TextSupport.IsNumberLocaleUnaware(srcString) Then CSngCustom = Val(srcString) Else CSngCustom = 0!
    Else
        CSngCustom = FixLocaleDependentNumber(srcString)
    End If
End Function

Private Function FixLocaleDependentNumber(ByVal srcString As String) As Double
    srcString = Replace$(srcString, ",", ".")
    If TextSupport.IsNumberLocaleUnaware(srcString) Then FixLocaleDependentNumber = Val(srcString) Else FixLocaleDependentNumber = 0#
End Function
