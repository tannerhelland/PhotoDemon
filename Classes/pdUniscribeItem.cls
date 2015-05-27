VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdUniscribeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Uniscribe Item Container
'Copyright 2015-2015 by Tanner Helland
'Created: 20/May/15
'Last updated: 20/May/15
'Last update: start migrating run-specific information into this class, to make it easier to handle complicated text layouts.
'
'This class is just a small container for holding individual Uniscribe item data.  We use a class instead of a struct because
' item coverage is highly variable; some items may hold a single glyph, while others hold 100+ glyphs.  It's impossible to
' predict run complexity until Uniscribe analyzes a string, then breaks it into individually shapeable items.
'
'We'll also need this functionality in the future to build a custom Rich-Textbox-type UC, so it makes sense to take full
' advantage of the functionality now.
'
'Relevant MSDN page for all things Uniscribe:
' https://msdn.microsoft.com/en-us/library/windows/desktop/dd374091%28v=vs.85%29.aspx
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Before proceeding further, it's worth explaining a few very confusing terms, which Uniscribe uses to describe the various
' units of text information it handles.  (Source: https://msdn.microsoft.com/en-us/library/windows/desktop/dd374094%28v=vs.85%29.aspx)

'ITEM:
' - ITEMS have a single script and direction (LTR or RTL).
' - ScriptItemize analyzes a *paragraph* and returns it as an array of ITEMS.  This means that Uniscribe *does not* handle text
'    larger than a paragraph, e.g. text that contains hard line breaks.  If a string contains hard line breaks (e.g. Cr/Lf chars),
'    you must subdivide it into paragraphs, then process each paragraph individually.
' - Simple text - e.g. text with a uniform style - may be both an ITEM and a RUN, but in general, you should not assume any equality
'    between these units.  This is true for even simple text, as Uniscribe prefers to break ITEMS on certain punctuation boundaries,
'    even if the script and direction do not change (and they thus would work fine as a RUN).
' - ITEMS can contain characters of multiple styles (font, size, color).  Note that this statement is very misleading, because
'    Uniscribe cannot automatically address or handle ITEMS consisting of multiple styles.  Rather, this means that ITEMS returned
'    by Uniscribe *can* be manually subdivided into smaller ITEMS of varying styles.  However, *YOU* must be the one to perform that
'    subdivision, as Uniscribe's itemizing function has no way for the caller to specify custom ITEM boundaries.
' - ITEM and RUN information are combined to create a full paragraph.

'RUN:
' - RUNS are contiguous passages of text with a uniform style (font, size, and color).
' - Multiple scripts (e.g. "languages" or "character sets", colloquially) are allowed in a RUN.
' - Unlike an ITEM, a RUN can contain both LTR and RTL content.

'These definitions are not hard and fast, and MSDN often uses one where it probably means the other.  I describe them here so
' you'll have compassion if I accidentally do the same.  ;)

'In general, I would prefer to use the term ITEM to describe a section of text with a uniform style, script, and direction.
' Multiple ITEMS can then be combined into a RUN, meaning a RUN always consists of one or more ITEMS.

'That definition makes much more sense to me, versus the gobbledygook used by Uniscribe's docs.

'-------------

'Now, onto the bits of per-item information this class stores.

'ScriptItemize is the first Uniscribe function called.  It divides a paragraph into individual items.  Each item has a uniform
' direction and script.  As an example, a sentence like "Hello [Arabic text] World" would consist of three items:
' "Hello"
' "[Arabic Text]"
' "World"

'"Hello" and "World" would be marked as LTR, while "[Arabic Text]" would be marked as RTL.  This is important, because we ultimately
' need to reorder RTL glyphs - post-shaping - into LTR order, to simplify the glyph rendering process.

'Anyway, this class represents a single item, meaning a single SCRIPT_ITEM struct as returned by ScriptItemize.
Private m_ScriptItem As SCRIPT_ITEM

'Upon generating our SCRIPT_ITEM copy, ScriptItemize also has knowledge of which chunk of the string our SCRIPT_ITEM represents.
' It will supply this substring via the get/setSubstring functions.  (Note that we also cache the length of this substring, as it's
' used for a lot of other calculations, like array bounds.
Private m_Substring As String
Private m_SubstringLength As Long

'The relative start and end position of this substring, relative to the original string passed to pdUniscribe.
' NOTE: in keeping with VB string measurements, these are 1-based, not 0-based.
Private m_RelativeStartPos As Long, m_RelativeEndPos As Long

'For each item created by ScriptItemize, ScriptLayout generates a corresponding Visual/Logical mapping order.  We don't store
' that order here (pdUniscribe manages it internally), but I mention this as a reminder that the order of these class instances
' inside pdUnicode's master array may not match their on-screen order, and you need to be aware of that if mapping from visual
' position to logical position (e.g. for caret or cursor behavior).  You can't simply rely on the internal order provided by
' pdUnicode, as it may not match your intended purpose.

'For each item, ScriptShape generates a cache list and returns three pieces of data:

' 1) an index list for mapping between characters and glyphs.  (This has one entry for every character in the string.)
Private m_LogicalClusterCache() As Integer

' 2) a list of glyphs.  The size of this list is totally independent of characters, and must be addressed separately.
Private m_NumOfGlyphs As Long
Private m_GlyphCache() As Integer

' 3) an array of glyph attributes.  This has the same size as the glyph cache in (2).
Private m_VisualAttributesCache() As SCRIPT_VISATTR

'For each item, ScriptPlace generates three pieces of data:

' 1) An advance width cache (the distance between neighboring glyphs, basically)
Private m_AdvanceWidthCache() As Long

' 2) a glyph offset cache (the displacement of a glyph within a cell, typically used for diacritics with variable per-glyph positioning)
Private m_GlyphOffsetCache() As GOFFSET

' 3) a single ABC measurement for the ENTIRE item.  (This is helpful when positioning the starts and ends of lines, to avoid glyph overhang.)
Private m_ItemABCWidth As ABC

'For each item, ScriptBreak generates one piece of data:

' 1) A logical attributes cache, with one entry FOR EACH CHARACTER (NOT GLYPH)!
Private m_CharAttributesCache() As SCRIPT_LOGATTR

'... because we only care about glyphs when rendering, we manually convert this data into a per-glyph list of attributes we need
' while rendering.  Irrelevant parameters are ignored.
Private m_GlyphSoftBreakCache() As Boolean
Private m_GlyphWhitespaceCache() As Boolean
Private m_GlyphZeroWidthCache() As Boolean

'ByRef + CopyMemory is confusing, especially when trading bare pointers (as I am prone to doing), so I prefer to use a ByVal-only
' version of RtlMoveMemory, and forcibly supply pointers to avoid the risk of messing things up.
Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)

Private Const DEFAULT_INITIAL_CACHE_SIZE As Long = 16

'Because VB doesn't allow a class to expose public UDTs (without a typelib, anyway), all custom structs use get/set functions.
' CopyMemory is used for a performance boost, but note that this doesn't work for variable-sized structs.
' (Those must be handled specially.)
Friend Sub getScriptItem(ByRef dstScriptItem As SCRIPT_ITEM)
    CopyMemoryStrict VarPtr(dstScriptItem), VarPtr(m_ScriptItem), Len(m_ScriptItem)
End Sub

Friend Sub setScriptItem(ByRef srcScriptItem As SCRIPT_ITEM)
    CopyMemoryStrict VarPtr(m_ScriptItem), VarPtr(srcScriptItem), Len(m_ScriptItem)
End Sub

'Passing bare pointers is cheaper than copying structs (or sub-structs), but please do not use these functions outside of pdUniscribe,
' and please do not cache their values.  This class may change memory allocations at any time, so you must always call these functions
' if you need a pointer to something.
Friend Function getScriptItemPointer() As Long
    getScriptItemPointer = VarPtr(m_ScriptItem)
End Function

Friend Function getScriptItemAnalysisPointer() As Long
    getScriptItemAnalysisPointer = VarPtr(m_ScriptItem.analysis)
End Function

'Substring is the only function that directly returns a value (since it uses an intrinsic VB type)
Public Function getSubstring() As String
    getSubstring = m_Substring
End Function

'When setting the substring, we set two other parameters as well: the relative start and end position of this substring, relative to
' the whole string passed to pdUniscribe.  This is helpful for identifying the chunk of the main string that this item represents.
'
'IMPORTANT NOTE: StartPos and EndPos are in *VB STRING COORDINATES*, meaning they are 1-based, not 0-based!
Public Sub setSubstring(ByVal newSubstring As String, ByVal relativeStartPos As Long, ByVal relativeEndPos As Long)
    m_Substring = newSubstring
    m_SubstringLength = Len(newSubstring)
    m_RelativeStartPos = relativeStartPos
    m_RelativeEndPos = relativeEndPos
End Sub

Public Function getRelativeStartPos() As Long
    getRelativeStartPos = m_RelativeStartPos
End Function

Public Function getRelativeEndPos() As Long
    getRelativeEndPos = m_RelativeEndPos
End Function

'ScriptShape generates three arrays for each item:
' 1) Logical cluster list (for mapping between characters and glyphs; same length as the substring)
' 2) Glyph list (actual indices into the font file for the generated glyphs; length is NOT the same as the substring)
' 3) Visual attributes list (extra display data for each glyph; same length as the glyph list)
'
'To improve performance, we copy all three of these simultaneously, at the end of ScriptShape processing.
Friend Sub setShapingData(ByRef logicalClusterCache() As Integer, ByVal numOfGlyphs As Long, ByRef glyphCache() As Integer, ByRef visualAttributesCache() As SCRIPT_VISATTR)
    
    'The logical cluster cache has the same length as our substring, so handle it separately
    If UBound(m_LogicalClusterCache) <> m_SubstringLength - 1 Then
        ReDim m_LogicalClusterCache(0 To m_SubstringLength - 1) As Integer
    End If
    
    CopyMemoryStrict VarPtr(m_LogicalClusterCache(0)), VarPtr(logicalClusterCache(0)), m_SubstringLength * 2
    
    'The glyph and visual attribute caches are the same size, so we can handle them together
    m_NumOfGlyphs = numOfGlyphs
    
    'Note that it is technically possible for a substring to not result in any actual glyphs.  (e.g. a bunch of zero-width control chars)
    ' This is still a valid return, and if it happens, we can't attempt to copy zero bytes of data.
    If m_NumOfGlyphs > 0 Then
    
        If UBound(m_GlyphCache) <> m_NumOfGlyphs - 1 Then
            ReDim m_GlyphCache(0 To m_NumOfGlyphs - 1) As Integer
            ReDim m_VisualAttributesCache(0 To m_NumOfGlyphs - 1) As SCRIPT_VISATTR
        End If
        
        CopyMemoryStrict VarPtr(m_GlyphCache(0)), VarPtr(glyphCache(0)), 2 * m_NumOfGlyphs
        CopyMemoryStrict VarPtr(m_VisualAttributesCache(0)), VarPtr(visualAttributesCache(0)), Len(m_VisualAttributesCache(0)) * m_NumOfGlyphs
        
    End If
    
End Sub

'If this item consists entirely of missing glyphs, we need to completely disable shaping.
Friend Sub disableAllShaping()
    
    'Shaping is controlled by the lowest 10 bits of our script analysis structure.
    m_ScriptItem.analysis.fBitFields1 = 0
    
    Dim tmpLong As Long
    tmpLong = m_ScriptItem.analysis.fBitFields2
    
    tmpLong = tmpLong And (Not 3&)
    m_ScriptItem.analysis.fBitFields2 = tmpLong And 255
    
End Sub

'While we can set all of ScriptShape's generated information at once, retrieving it is more helpful if we do it on a per-item basis.
Friend Function getNumOfGlyphs() As Long
    getNumOfGlyphs = m_NumOfGlyphs
End Function

'In some cases, it's easier (and much more lightweight!) to pass bare pointers to our various large structs and caches.
' Obviously, you MUST not cache these between requests, as this class may change its memory allocation at any time.
Friend Function getPointerToGlyphCache() As Long
    getPointerToGlyphCache = VarPtr(m_GlyphCache(0))
End Function

Friend Function getPointerToVisualAttributesCache() As Long
    getPointerToVisualAttributesCache = VarPtr(m_VisualAttributesCache(0))
End Function

'Similarly, sometimes we just want to grab a single entry at a given position.  For performance reasons, none of these wrappers check bounds.
' It's assumed that the caller is iterating using the result of getNumOfGlyphs(), so bounds won't be a problem.
Friend Function getGlyphCacheAtPosition(ByVal gPos As Long) As Long
    
    'Glyphs are returned by Uniscribe as 16-bit values.  Obviously, VB treats these values as if they are signed.
    ' To work around this, we forcibly deal with negative integer values.
    getGlyphCacheAtPosition = (m_GlyphCache(gPos) And &HFFFF&)
    
End Function

Friend Function getAdvanceWidthAtPosition(ByVal gPos As Long) As Long
    getAdvanceWidthAtPosition = m_AdvanceWidthCache(gPos)
End Function

Friend Sub copyGlyphOffsetToPointer(ByVal dstPointer As Long, ByVal gPos As Long)
    CopyMemoryStrict dstPointer, VarPtr(m_GlyphOffsetCache(gPos)), Len(m_GlyphOffsetCache(gPos))
End Sub

'ScriptPlace generates two arrays and one custom struct for each item:
' 1) Advance width list (distance between neighboring glyphs, basically)
' 2) Glyph offset list (displacement of a glyph within its cell; this is primarily used for diacritics with variable per-glyph positioning)
' 3) ABC measurement for the ENTIRE item.
'
'To improve performance, we copy all three of these simultaneously, at the end of ScriptPlace processing.
Friend Sub setPlacementData(ByRef advanceWidthCache() As Long, ByRef glyphOffsetCache() As GOFFSET, ByRef itemABCWidth As ABC)
    
    'The advance width and glyph offset caches are the same size, so we can handle them together
    If m_NumOfGlyphs > 0 Then
    
        If UBound(m_GlyphOffsetCache) <> m_NumOfGlyphs - 1 Then
            ReDim m_GlyphOffsetCache(0 To m_NumOfGlyphs - 1) As GOFFSET
            ReDim m_AdvanceWidthCache(0 To m_NumOfGlyphs - 1) As Long
        End If
        
        CopyMemoryStrict VarPtr(m_GlyphOffsetCache(0)), VarPtr(glyphOffsetCache(0)), 4 * m_NumOfGlyphs
        CopyMemoryStrict VarPtr(m_AdvanceWidthCache(0)), VarPtr(advanceWidthCache(0)), Len(m_AdvanceWidthCache(0)) * m_NumOfGlyphs
        
    End If
    
    m_ItemABCWidth = itemABCWidth
    
End Sub

'ScriptBreak generates a per-character array of logical attributes.  This is a pain, as the renderer only cares about glyphs, not chars,
' so we have to manually convert this array into something usable.
Friend Sub setCharacterAttributeCache(ByRef charAttributesCache() As SCRIPT_LOGATTR)
    
    If m_SubstringLength > 0 Then
    
        'First, make an internal copy of the data
        If UBound(m_CharAttributesCache) < m_SubstringLength Then ReDim m_CharAttributesCache(0 To m_SubstringLength) As SCRIPT_LOGATTR
        CopyMemoryStrict VarPtr(m_CharAttributesCache(0)), VarPtr(charAttributesCache(0)), m_SubstringLength
        
        'Now, we need to convert the per-character list of attributes to a per-glyph attribute.  This is a pain, as characters and
        ' glyphs do not necessarily have a 1:1 correlation.
        
        'Start by prepping our glyph softbreak cache
        If UBound(m_GlyphSoftBreakCache) < m_NumOfGlyphs Then
            ReDim m_GlyphSoftBreakCache(0 To m_NumOfGlyphs) As Boolean
            ReDim m_GlyphWhitespaceCache(0 To m_NumOfGlyphs) As Boolean
            ReDim m_GlyphZeroWidthCache(0 To m_NumOfGlyphs) As Boolean
        End If
        
        'Forcibly set all soft break and whitespace markers to false, as we initialize these arrays lazily
        Dim i As Long
        For i = 0 To m_NumOfGlyphs
            m_GlyphSoftBreakCache(i) = False
            m_GlyphWhitespaceCache(i) = False
            m_GlyphZeroWidthCache(i) = False
        Next i
                
        Dim tmpLong As Long, tmpIndex As Long
        
        'Because our break information is per-character, we'll iterate the string by character (not glyph).  If any character in a
        ' glyph is soft-breakable, we mark the ENTIRE glyph as breakable.
        For i = 0 To m_SubstringLength - 1
        
            'Find the glyph that corresponds to this character
            tmpIndex = m_LogicalClusterCache(i)
            tmpLong = charAttributesCache(i).fBitFields
            
            'Like most horribly obnoxious Uniscribe structs, data is embedded at the bit level.  Retrieve the flag bit for soft breaks.
            If getBitFlag(0, tmpLong) Then
                
                'Mark this glyph as soft-breakable.  This means that - if necessary - a line can be broken IMMEDIATELY BEFORE this glyph.
                If tmpIndex >= 0 And tmpIndex < m_NumOfGlyphs Then
                    m_GlyphSoftBreakCache(tmpIndex) = True
                Else
                    Debug.Print "WARNING!  m_LogicalClusterCache() returned a bogus value.  Please investigate!"
                End If
                
            End If
            
            '...and the same thing, but for white space
            If getBitFlag(1, tmpLong) Then
                
                'Mark this glyph as soft-breakable.  This means that - if necessary - a line can be broken IMMEDIATELY BEFORE this glyph.
                If tmpIndex >= 0 And tmpIndex < m_NumOfGlyphs Then
                    m_GlyphWhitespaceCache(tmpIndex) = True
                Else
                    Debug.Print "WARNING!  m_LogicalClusterCache() returned a bogus value.  Please investigate!"
                End If
                
            End If
                                
        Next i
        
        'The zero-width marker is different.  It is stored on a per-glyph basis, so we iterate over each glyph directly.
        For i = 0 To m_NumOfGlyphs - 1
        
            'Find the glyph that corresponds to this character
            tmpLong = m_VisualAttributesCache(i).fBitFields1
        
            'Retrieve the zero-width bit
            If getBitFlag(6, tmpLong) Then
                
                'Mark this glyph as soft-breakable.  This means that - if necessary - a line can be broken IMMEDIATELY BEFORE this glyph.
                m_GlyphZeroWidthCache(i) = True
                
            End If
            
        Next i
        
    End If
        
End Sub

Friend Function getSoftBreakAtPosition(ByVal gPos As Long) As Boolean
    getSoftBreakAtPosition = m_GlyphSoftBreakCache(gPos)
End Function

Friend Function getWhitespaceAtPosition(ByVal gPos As Long) As Boolean
    getWhitespaceAtPosition = m_GlyphWhitespaceCache(gPos)
End Function

Friend Function getZeroWidthAtPosition(ByVal gPos As Long) As Boolean
    getZeroWidthAtPosition = m_GlyphZeroWidthCache(gPos)
End Function

Friend Function getCorrespondingCharIndex(ByVal gPos As Long) As Long

    'Search the logical cluster cache until we find a glyph position that matches.  Note that the relationship between characters
    ' and glyphs is *not* 1:1, so one glyph may match to multiple characters!  This function will only return the first hit.
    ' (In the future, I suppose we could return an LBound and UBound if relevant, but for our current tasks, this is good enough.)
    Dim i As Long
    For i = 0 To UBound(m_LogicalClusterCache)
    
        If m_LogicalClusterCache(i) = gPos Then
            getCorrespondingCharIndex = m_RelativeStartPos + i
            Exit For
        End If
    
    Next i
    
End Function

'Retrieve an arbitrary 1-bit position [0-31] in a Long-type value.  Position 0 is the LEAST-SIGNIFICANT BIT, and Position 31
' is the SIGN BIT for a standard VB Long.
'
' Inputs:
'  1) position of the flag, which must be in the range [0, 31]
'  2) The Long from which you want the bit retrieved
Private Function getBitFlag(ByVal flagPosition As Long, ByVal srcLong As Long) As Boolean

    If (flagPosition >= 0) And (flagPosition <= 31) Then
    
        'Create a bitmask, using flagPosition to determine bit offset
        Dim longMask As Long
        If flagPosition < 31 Then longMask = 2 ^ flagPosition Else longMask = &H80000000
        getBitFlag = CBool(CLng(srcLong And longMask) <> 0)
        
        'If you're curious about how this function works (and what it returns), you can uncomment the line below
        'Debug.Print "getBitFlag: " & srcLong & " : " & longMask & " (" & CStr(getBitFlag) & ")"
        
    Else
        Debug.Print "C'mon - you know Longs only have 32 bits.  Flag positions must be on the range [0, 31]!"
    End If

End Function

Private Sub Class_Initialize()
    
    'Size all arrays to a nice minimum size, to reduce churn if the user is only working with small strings
    ReDim m_LogicalClusterCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Integer
    ReDim m_GlyphCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Integer
    ReDim m_VisualAttributesCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As SCRIPT_VISATTR
    ReDim m_AdvanceWidthCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Long
    ReDim m_GlyphOffsetCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As GOFFSET
    ReDim m_CharAttributesCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As SCRIPT_LOGATTR
    ReDim m_GlyphSoftBreakCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Boolean
    ReDim m_GlyphWhitespaceCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Boolean
    ReDim m_GlyphZeroWidthCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Boolean
        
End Sub
