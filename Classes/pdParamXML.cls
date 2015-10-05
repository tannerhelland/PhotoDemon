VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdParamXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Internal Parameter Handler v2
'Copyright 2013-2015 by Tanner Helland
'Created: 25/March/13
'Last updated: 05/October/15
'Last update: new helper function for adding a bunch of parameters at once
'
'PhotoDemon has unique needs regarding parameter passing.  Because it allows the user to record all actions as part of
' macros, the program needs a way to not only trade around parameters, but also store them in a file.
'
'The problem occurs when trying to develop a single system that works with all possible function parameters.  Some actions,
' like "Invert Image", require no additional information.  Others, like "Curves", require a huge amount of custom data.
' Various functions utilize conceivable every type of value (bool, int, float, custom types, etc), and if a macro is
' recording a function, all those specialized parameters need to be tracked and written to file.
'
'The easiest way to handle that kind of variety from within VB is to use a string.  This allows any amount - and type -
' of custom parameter data to be recorded, stored, and easily transferred between functions.
'
'Individual functions therefore rely on this class to create and parse parameter strings for them.
'
'In summer 2015, I moved from a pipe-delimited parameter system (where parameters were listed in order, and retrieved by order)
' to an order-agnostic XML system, where parameters are retrieved by name.  This class uses its own XML parsing code, which
' includes a few greatly stripped-down (and speed-optimized) basic XML functions from the main pdXML class.  This micro-XML
' parsing code introduces a few restrictions:
' 1) All comparisons are case-sensitive.  If you change case in PD's source, you will invalidate old parameters, by design.
' 2) All parameter names must be unique.
' 3) Parameter names do not support attribute tags.  Come up with different parameter names, instead.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit
Option Compare Binary

'Current parameter string.
Private m_ParamString As String

Public Function getParamString() As String
    getParamString = m_ParamString
End Function

'If you obtain a parameter string from elsewhere (e.g. a PDI file), you can simple pass it as-is to this function
Public Sub setParamString(ByVal srcString As String)
    m_ParamString = srcString
End Sub

Public Sub Reset(Optional ByVal paramStringVersion As Double = 1#)
    m_ParamString = "<version>" & Trim$(Str(paramStringVersion)) & "</version>"
End Sub

Private Sub Class_Initialize()
    
    'For now, this class does not add any extra XML around the parameters themselves.
    ' It will, however, create a generic "version" tag at version 1.0.  Functions are free to override this as necessary.
    Me.Reset
    
End Sub

'If a function changes drastically enough, it is free to use the "created-by-default" version tag to adjust handling as necessary.
Public Function getParamVersion() As Double
    
    Dim versionString As String
    If getParamValue("version", versionString) Then
        getParamVersion = CDblCustom(versionString)
    Else
        getParamVersion = 0
    End If
    
End Function

Public Sub setParamVersion(Optional ByVal newVersion As Double = 1#)
    updateParam "version", newVersion, True
End Sub

'Simple check to see if a parameter exists
Public Function doesParamExist(ByRef paramName As String) As Boolean
    
    'Make the parameter name XML-safe
    paramName = getXMLSafeName(paramName)
    
    If InStr(1, m_ParamString, "<" & paramName & ">") Then
        doesParamExist = True
    Else
        doesParamExist = False
    End If
    
End Function

'Retrieve various type-specific parameters.  Note that these ultimately wrap getParamValue; they simply cast the result explicitly.
Public Function GetBool(ByVal paramName As String, Optional ByVal defaultReturn As Boolean = False) As Boolean
    
    Dim paramValue As String
    If getParamValue(paramName, paramValue) Then
        GetBool = CBool(Trim$(paramValue))
    Else
        GetBool = defaultReturn
    End If

End Function

Public Function GetByte(ByVal paramName As String, Optional ByVal defaultReturn As Byte = 0) As Byte

    Dim paramValue As String
    If getParamValue(paramName, paramValue) Then
        GetByte = CByte(Trim$(paramValue))
    Else
        GetByte = defaultReturn
    End If

End Function

Public Function GetInteger(ByVal paramName As String, Optional ByVal defaultReturn As Integer = 0) As Integer

    Dim paramValue As String
    If getParamValue(paramName, paramValue) Then
        GetInteger = CInt(CDblCustom(Trim$(paramValue)))
    Else
        GetInteger = defaultReturn
    End If

End Function

Public Function GetLong(ByVal paramName As String, Optional ByVal defaultReturn As Long = 0) As Long

    Dim paramValue As String
    If getParamValue(paramName, paramValue) Then
        GetLong = CLng(CDblCustom(Trim$(paramValue)))
    Else
        GetLong = defaultReturn
    End If

End Function

Public Function GetSingle(ByVal paramName As String, Optional ByVal defaultReturn As Single = 0#) As Single

    Dim paramValue As String
    If getParamValue(paramName, paramValue) Then
        GetSingle = CDblCustom(Trim$(paramValue))
    Else
        GetSingle = defaultReturn
    End If

End Function

Public Function GetDouble(ByVal paramName As String, Optional ByVal defaultReturn As Double = 0#) As Double

    Dim paramValue As String
    If getParamValue(paramName, paramValue) Then
        GetDouble = CDblCustom(Trim$(paramValue))
    Else
        GetDouble = defaultReturn
    End If

End Function

Public Function GetString(ByVal paramName As String, Optional ByVal defaultReturn As String = "") As String

    Dim paramValue As String
    If getParamValue(paramName, paramValue) Then
        GetString = paramValue
    Else
        GetString = defaultReturn
    End If

End Function

Public Function GetVariant(ByVal paramName As String, Optional ByVal defaultReturn As Variant = Empty) As Variant

    Dim paramValue As String
    If getParamValue(paramName, paramValue) Then
        GetVariant = CVar(paramValue)
    Else
        GetVariant = defaultReturn
    End If

End Function

'Given a parameter name, fill a user-supplied string with the parameter value.
' Returns TRUE if parameter exists; FALSE otherwise.
Private Function getParamValue(ByVal paramName As String, ByRef dstString As String) As Boolean
    
    'Make the parameter name XML-safe
    paramName = getXMLSafeName(paramName)
    
    Dim tagStart As Long, tagEnd As Long
    tagStart = InStr(1, m_ParamString, "<" & paramName & ">")
    
    'If the opening tag was found, we also need to find the closing tag.
    If tagStart > 0 Then
    
        tagEnd = InStr(tagStart, m_ParamString, "</" & paramName & ">")
        
        'If the closing tag exists, return everything between that and the opening tag
        If tagEnd > tagStart Then
            
            'Increment the tag start location by the length of the tag plus two (+1 for each bracket: <>)
            tagStart = tagStart + Len(paramName) + 2
            
            If tagEnd > tagStart Then
                dstString = Mid$(m_ParamString, tagStart, tagEnd - tagStart)
                dstString = unDelimitParamValue(dstString)
                getParamValue = True
            Else
                dstString = ""
                getParamValue = False
            End If
            
        Else
            Debug.Print "WARNING: requested parameter (" & paramName & ") wasn't properly closed!"
            dstString = ""
            getParamValue = False
        End If
        
    Else
        dstString = ""
        getParamValue = False
    End If

End Function

'Blindly add a parameter to the master string.  No special checks (e.g. duplicates) are applied; use updateParam if you need those.
Public Function addParam(ByVal paramName As String, ByVal paramValue As Variant) As Boolean
    
    'Convert the parameter value into a string.  We handle this manually to minimize the chance of locale issues.
    Dim strParamValue As String
    
    If VarType(paramValue) = vbByte Then
        strParamValue = Trim$(Str(paramValue))
    ElseIf VarType(paramValue) = vbInteger Then
        strParamValue = Trim$(Str(paramValue))
    ElseIf VarType(paramValue) = vbLong Then
        strParamValue = Trim$(Str(paramValue))
    ElseIf VarType(paramValue) = vbSingle Then
        strParamValue = Trim$(Str(paramValue))
    ElseIf VarType(paramValue) = vbDouble Then
        strParamValue = Trim$(Str(paramValue))
    ElseIf VarType(paramValue) = vbDecimal Then
        strParamValue = Trim$(Str(paramValue))
    ElseIf VarType(paramValue) = vbCurrency Then
        strParamValue = Trim$(Str(paramValue))
    ElseIf VarType(paramValue) = vbNull Then
        strParamValue = Trim$(Str(0))
    ElseIf VarType(paramValue) = vbBoolean Then
        strParamValue = Trim$(Str(paramValue))
    ElseIf VarType(paramValue) = vbString Then
        strParamValue = paramValue
    ElseIf VarType(paramValue) = vbDate Then
        strParamValue = Format(paramValue, "yyyy-mm-dd h:mm:ss", vbSunday, vbFirstJan1)
    
    'Pray for a correct implicit cast result
    Else
        strParamValue = paramValue
    End If
    
    'Make the parameter name and value XML-safe
    paramName = getXMLSafeName(paramName)
    strParamValue = delimitParamValue(strParamValue)
    
    'Build a string with the parameter name and value we were passed
    Dim newParamEntry As String
    newParamEntry = vbCrLf & "<" & paramName & ">" & strParamValue & "</" & paramName & ">"
            
    'Tack it onto the master string
    m_ParamString = m_ParamString & newParamEntry
    
    addParam = True
    
End Function

'Want to add multiple parameters at once?  Use this function, and call it using the order "Name", Value, "Name 2", Value 2...
Public Function addMultipleParams(ParamArray allParams() As Variant) As Boolean
    
    On Error GoTo multiParamFailure
    
    If UBound(allParams) >= LBound(allParams) Then
    
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
            Me.addParam tmpName, tmpValue
            
        Next i
    
    End If
    
    addMultipleParams = True
    
    Exit Function
    
multiParamFailure:
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  Parse failure in pdParamXML.addMultipleParams()."
    #End If
    
    addMultipleParams = False

End Function

'Update a given parameter.  If the parameter is not found, it will be added to the string. (Create-if-missing behavior can be toggled.)
Public Function updateParam(ByVal paramName As String, ByVal paramValue As String, Optional ByVal createIfMissing As Boolean = True) As Boolean
    
    'Make the parameter name and value XML-safe
    paramName = getXMLSafeName(paramName)
    paramValue = delimitParamValue(paramValue)
    
    'See if the parameter already exists
    Dim paramLocation As Long
    paramLocation = InStr(1, m_ParamString, "<" & paramName & ">")
        
    Dim topHalf As String, bottomHalf As String
        
    'If the parameter already exists, just update its value in-place.
    If (paramLocation > 0) Then
    
        'Split the XML file into two halves: the half before the relevant tag, and the half after
        Dim paramCloseLocation As Long
        paramCloseLocation = InStr(paramLocation, m_ParamString, "</" & paramName & ">")
        splitStringIn2 m_ParamString, paramCloseLocation - 1, topHalf, bottomHalf
        
        'The "topHalf" string now includes everything before the closing tag.  Chop it off at the end of the start tag (e.g. after
        ' the closing bracket), add the new contents, then add the bottom half of the original XML string.
        m_ParamString = Left$(topHalf, paramLocation + Len(paramName) + 1) & paramValue & bottomHalf
        
        updateParam = True
    
    'The parameter does not exist; add it contingent on createIfMissing
    Else
    
        If createIfMissing Then
            
            'Build a string with the parameter name and value we were passed
            Dim newParamEntry As String
            newParamEntry = vbCrLf & "<" & paramName & ">" & paramValue & "</" & paramName & ">"
            
            'Reassemble the primary string
            m_ParamString = m_ParamString & newParamEntry
            
            updateParam = True
            
        Else
            updateParam = False
        End If
        
    End If

End Function

'Given a string and a position, split it into two strings at that position
Private Function splitStringIn2(ByRef srcString As String, ByVal splitPosition As Long, ByRef dstFirstHalf As String, ByRef dstSecondHalf As String)
    dstFirstHalf = Left$(srcString, splitPosition)
    dstSecondHalf = Right$(srcString, Len(srcString) - splitPosition)
End Function

'Given a string, replace any characters that are not allowed with underscores; this is used as a failsafe when adding
' new parameters to the master string.
Private Function getXMLSafeName(ByRef srcString As String) As String
    
    'Remove any incidental white space before processing
    getXMLSafeName = Trim(srcString)
    
    'Create a string of valid numerical characters, based on the XML spec at http://www.w3.org/TR/1998/REC-xml-19980210.html#sec-common-syn
    Dim validChars As String
    validChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-_:"
    
    'Loop through the source string and replace any invalid characters with underscore
    Dim i As Long
    
    For i = 1 To Len(getXMLSafeName)
        If InStr(validChars, Mid$(getXMLSafeName, i, 1)) = 0 Then
            getXMLSafeName = Left$(getXMLSafeName, i - 1) & "_" & Right$(getXMLSafeName, Len(getXMLSafeName) - i)
        End If
    Next i
    
End Function

'Given a parameter value string, make it XML-safe (e.g. replace "<" and ">" with HTML equivalents).
' Note that we don't currently delimit other chars because PD uses only a subset of XML functionality, by design.
Private Function delimitParamValue(ByVal srcString As String) As String

    If InStr(1, srcString, "<") Then srcString = Replace(srcString, "<", "&lt;")
    If InStr(1, srcString, ">") Then srcString = Replace(srcString, ">", "&gt;")
    
    delimitParamValue = srcString

End Function

Private Function unDelimitParamValue(ByVal srcString As String) As String
        
    If InStr(1, srcString, "&lt;") Then srcString = Replace(srcString, "&lt;", "<")
    If InStr(1, srcString, "&gt;") Then srcString = Replace(srcString, "&gt;", ">")
    
    unDelimitParamValue = srcString
    
End Function

'A custom CDbl function that accepts both commas and decimals as a separator; this is important when moving floating-point data,
' represented as strings, between locales.
Private Function CDblCustom(ByVal srcString As String) As Double

    'Replace commas with periods
    If InStr(1, srcString, ",") > 0 Then srcString = Replace(srcString, ",", ".")
    
    'We can now use Val() to convert to Double
    If IsNumberLocaleUnaware(srcString) Then
        CDblCustom = Val(srcString)
    Else
        CDblCustom = 0
    End If

End Function
