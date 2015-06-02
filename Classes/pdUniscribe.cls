VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdUniscribe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Uniscribe API Interface
'Copyright 2015-2015 by Tanner Helland
'Created: 14/May/15
'Last updated: 19/May/15
'Last update: migrate everything from a module to a class; there's enough useful code in here that we may want to
' use it multiple places throughout PD, and because so much data has to be cached on a per-DC and per-string basis,
' a module won't work.  (I may also need to subclass some system messages, to account for OS-level RTL settings.)
'
'Relevant MSDN page for all things Uniscribe:
' https://msdn.microsoft.com/en-us/library/windows/desktop/dd374091%28v=vs.85%29.aspx
'
'Many thanks to Michael Kaplan for his endless work in demystifying Windows text handling.  Of particular value to
' this module is his personal blog - http://www.siao2.com/ - which I referenced liberally during the assembly of this
' class.  His blog includes an article with some Uniscribe-related VB declarations from his (now out of print) book
' on VB internationalization, but I am deliberately avoiding linking it as many of the declarations are missing or
' simply incorrect.  I *strongly* recommend referring to MSDN directly if you plan on working with Uniscribe.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'All Uniscribe APIs relevant to PD.  Note that we do not use the simplified ScriptString wrappers; they are not powerful enough
' for the kinds of tasks we tackle.
Private Declare Function ScriptApplyDigitSubstitution Lib "usp10" (psds As SCRIPT_DIGITSUBSTITUTE, psc As SCRIPT_CONTROL, pss As SCRIPT_STATE) As Long
Private Declare Function ScriptApplyLogicalWidth Lib "usp10" (piDx() As Long, ByVal cChars As Long, ByVal cGlyphs As Long, pwLogClust() As Integer, psva As SCRIPT_VISATTR, piAdvance() As Long, psa As SCRIPT_ANALYSIS, pABC As ABC, piJustify As Long) As Long
Private Declare Function ScriptBreak Lib "usp10" (ByVal ptrTopwcChars As Long, ByVal cChars As Long, ByVal ptrTopsa As Long, ByVal ptrTopsla As Long) As Long
Private Declare Function ScriptCPtoX Lib "usp10" (ByVal iCP As Long, ByVal fTrailing As Long, ByVal cChars As Long, ByVal cGlyphs As Long, pwLogClust As Integer, psva As SCRIPT_VISATTR, piAdvance As Long, psa As SCRIPT_ANALYSIS, piX As Long) As Long
Private Declare Function ScriptCacheGetHeight Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, tmHeight As Long) As Long
Private Declare Function ScriptFreeCache Lib "usp10" (psc As SCRIPT_CACHE) As Long
Private Declare Function ScriptGetCMap Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, ByVal pwcInChars As Long, ByVal cChars As Long, ByVal dwFlags As SCRIPT_GET_CMAP_FLAGS, pwOutGlyphs() As Integer) As Long
Private Declare Function ScriptGetFontProperties Lib "usp10" (ByVal srcDC As Long, ByVal ptrToSCpsc As Long, ByVal ptrTosfp As Long) As Long
Private Declare Function ScriptGetGlyphABCWidth Lib "usp10" (ByVal srcDC As Long, ByVal ptrTopscScriptCache As Long, ByVal wGlyph As Integer, ByVal ptrToABC As Long) As Long
Private Declare Function ScriptGetLogicalWidths Lib "usp10" (psa As SCRIPT_ANALYSIS, ByVal cChars As Long, ByVal cGlyphs As Long, piGlyphWidth() As Long, pwLogClust() As Integer, psva As SCRIPT_VISATTR, piDx As Long) As Long
Private Declare Function ScriptGetProperties Lib "usp10" (ppSp As SCRIPT_PROPERTIES, piNumScripts As Long) As Long
Private Declare Function ScriptIsComplex Lib "usp10" (ByVal pwcInChars As Long, ByVal cInChars As Long, ByVal dwFlags As SCRIPT_IS_COMPLEX_FLAGS) As Long
Private Declare Function ScriptItemize Lib "usp10" (ByVal pwcInChars As Long, ByVal cInChars As Long, ByVal cMaxItems As Long, ByVal ptrToScriptControl As Long, ByVal ptrToScriptState As Long, ByVal ptrToPItems As Long, ByRef pcItems As Long) As Long
Private Declare Function ScriptJustify Lib "usp10" (psva As SCRIPT_VISATTR, piAdvance() As Long, ByVal cGlyphs As Long, ByVal iDx As Long, ByVal iMinKashida As Long, piJustify() As Long) As Long
Private Declare Function ScriptLayout Lib "usp10" (ByVal cRuns As Long, ByVal ptrToPBLevel As Long, ByVal ptrToPIVisualToLogical As Long, ByVal ptrToPILogicalToVisual As Long) As Long
Private Declare Function ScriptPlace Lib "usp10" (ByVal srcDC As Long, ByRef psc As SCRIPT_CACHE, ByVal ptrToIntpwGlyphs As Long, ByVal cGlyphs As Long, ByVal ptrToSVpsva As Long, ByVal ptrTopsa As Long, ByVal ptrToLngpiAdvance As Long, ByVal ptrTopGoffset As Long, ByRef pABC As ABC) As Long
Private Declare Function ScriptRecordDigitSubstitution Lib "usp10" (ByVal Locale As Long, psds As SCRIPT_DIGITSUBSTITUTE) As Long
Private Declare Function ScriptShape Lib "usp10" (ByVal srcDC As Long, ByVal ptrToSCpsc As Long, ByVal pwcChars As Long, ByVal cChars As Long, ByVal cMaxGlyphs As Long, ByVal ptrToSApas As Long, ByVal ptrToIntpwOutGlyphs As Long, ByVal ptrToIntpwLogClust As Long, ByVal ptrToSVpsva As Long, ByRef pcGlyphs As Long) As Long
Private Declare Function ScriptTextOut Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, ByVal x As Long, ByVal y As Long, ByVal fuOptions As Long, lprc As RECTL, psa As SCRIPT_ANALYSIS, ByVal pwcReserved As Long, ByVal iReserved As Long, pwGlyphs() As Integer, ByVal cGlyphs As Long, piAdvance() As Long, piJustify As Any, pGoffset As GOFFSET) As Long
Private Declare Function ScriptXtoCP Lib "usp10" (ByVal iX As Long, ByVal cChars As Long, ByVal cGlyphs As Long, pwLogClust() As Integer, psva As SCRIPT_VISATTR, piAdvance() As Long, psa As SCRIPT_ANALYSIS, piCP As Long, piTrailing As Long) As Long

'Some Uniscribe features weren't added until Vista; these enable all kinds of neat behavior, but we must call them conditionally
Private Declare Function ScriptItemizeOpenType Lib "usp10" (ByVal pwcInChars As Long, ByVal cInChars As Long, ByVal cMaxItems As Long, ByVal ptrToScriptControl As Long, ByVal ptrToScriptState As Long, ByVal ptrToPItems As Long, ByVal ptrToPScriptTags As Long, ByRef pcItems As Long) As Long
Private Declare Function ScriptShapeOpenType Lib "usp10" (ByVal srcDC As Long, ByVal ptrToSCpsc As Long, ByVal ptrToSApas As Long, ByVal otTagScript As Long, ByVal otTagLang As Long, ByVal ptrToRcRangeChars As Long, ByVal ptrToRpRangeProperties As Long, ByVal numOfOTRanges As Long, ByVal pwcChars As Long, ByVal cChars As Long, ByVal cMaxGlyphs As Long, ByVal ptrToIntpwLogClust As Long, ByVal ptrToOutputCharProps As Long, ByVal ptrToPWOutGlyphs As Long, ByVal ptrToSVpsva As Long, ByRef pcGlyphs As Long) As Long


'For relevant type declarations, look in the Uniscribe_API_Decs module.

'ScriptItemize allows you to specify a language ID you want to use for itemizing.  We default to the current user setting.
Private Const LANG_USER_DEFAULT As Integer = &H400&

'Current string associated with our cache; if this doesn't change, we can skip most Uniscribe processing.  Note that this string will ultimately
' be subdivided into smaller strings, each of which is associated with an individual item.
Private m_CurrentString As String

'SCRIPT_ITEM cache generated by Step 1
Private m_ScriptItemCacheOK As Boolean
Private m_ScriptItemsCache() As SCRIPT_ITEM

'Visual order of the runs in m_ScriptItemsCache()
Private m_VisualToLogicalOrder() As Long

'Logical order of the runs in m_ScriptItemsCache()
Private m_LogicalToVisualOrder() As Long

'Opaque SCRIPT_CACHE handle.  This is allocated by the system on the initial call to ScriptShape.  We must free this handle
' when we are done with it.
Private m_ScriptCache As SCRIPT_CACHE

'Collection of Uniscribe items, which as a whole form a collective "run".  At present, these exactly mirror what we are handed
' by various stages of Uniscribe analysis, but in the future, it may be possible to split and/or merge these to implement our
' own per-character style support.
Private m_Items() As pdUniscribeItem
Private m_NumOfItems As Long

'Glyph cache.  This is generated by step 3, and is crucial for rendering as it contains the actual glyph indices from the current font!
Private m_GlyphCache() As Integer
Private m_NumOfGlyphs As Long

'Logical cluster cache.  This is generated by step 3, and per MSDN, "the value of each element is the offset from the first glyph in the
' run to the first glyph in the cluster containing the corresponding character."  Basically, this provides a mapping from character to
' glyph.  As such, one logical cluster entry exists for each character (not glyph!)
Private m_LogicalClusterCache() As Integer

'Glyph visual attributes cache.  This is generated by step 3, and one visual attribute entry is present for each glyph (not char!)
Private m_VisualAttributesCache() As SCRIPT_VISATTR

'Advance width cache.  This is generated by step 4, and one advance width is present for each glyph (not char!).  These values are the
' offsets, in pixels, from one glyph to the next.
Private m_AdvanceWidthCache() As Long

'Glyph offset cache.  This is generated by step 4, and one offset is present for each glyph (not char!).  This struct is only filled
' for combining glyphs - for example, for an "A" with an accent "`" over it, some fonts may only provide a plain "A" glyph, and a plain
' "`" accent glyph.  This offset tells us where to move the "`" accent glyph over the A, and the same "`" accent over a different glyph
' may use a different offset (e.g. for a lower-case "a").
Private m_GlyphOffsetCache() As GOFFSET

'On the first instantiation of a Uniscribe cache, we also retrieve some generic font information via ScriptGetFontProperties.  This lets
' us know things like which glyph to use for missing glyphs, which is crucial when intermixing scripts in a single string.
Private m_ScriptFontProperties As SCRIPT_FONTPROPERTIES

'ScriptBreak generates an array of SCRIPT_LOGATTR structures, one per Unicode codepoint in the original string.
Private m_CharAttributesCache() As SCRIPT_LOGATTR

'If the incoming string contains hard line-breaks, this will be set to TRUE.  Hard line-breaks require extra processing, as Uniscribe
' provides no native support for them.
Private m_StringHasLinebreaks As Boolean
Private m_NumOfLineBreaks As Long
Private m_LineBreakPositions() As Long

Private Const DEFAULT_INITIAL_CACHE_SIZE As Long = 16

'Set an arbitrary 1-bit position [0-31] in a Long-type value to either 1 or 0.  Position 0 is the LEAST-SIGNIFICANT BIT, and Position 31
' is the SIGN BIT for a standard VB Long.
'
' Inputs:
'  1) position of the flag, which must be in the range [0, 31]
'  2) value of the flag, TRUE for 1, FALSE for 0
'  3) The Long-type value where you want the flag placed
Private Sub setBitFlag(ByVal flagPosition As Long, ByVal flagValue As Boolean, ByRef copyOfDstLong As Long)

    If (flagPosition >= 0) And (flagPosition <= 31) Then
    
        'Create a bitmask, using flagPosition to determine bit offset
        Dim longMask As Long
        If flagPosition < 31 Then longMask = 2 ^ flagPosition Else longMask = &H80000000
        
        'Write a 1 flag
        If flagValue Then
        
            'Blend the flag with the source Long using OR, which will force the target bit to 1 while preserving the rest of the Long's contents
            copyOfDstLong = copyOfDstLong Or longMask
        
        'Write a 0 flag
        Else
        
            'Blend the INVERSE flag with the source Long, which will force the target bit to 0 while preserving the rest of the Long's contents
            copyOfDstLong = copyOfDstLong And Not longMask
        
        End If
        
    Else
        Debug.Print "C'mon - you know Longs only have 32 bits.  Flag positions must be on the range [0, 31]!"
    End If

End Sub

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

'Free any/all previously allocated cache structures
Public Function freeUniscribeCaches()
    If m_ScriptCache.p <> 0 Then ScriptFreeCache m_ScriptCache
End Function

'Uniscribe caches a number of font-specific values (https://msdn.microsoft.com/en-us/library/windows/desktop/dd317726%28v=vs.85%29.aspx),
' which can greatly improve performance during glyph retrieval and placement.  It maintains its own reference counts, but we can improve
' cache behavior by manually freeing our style cache whenever the associated font object is modified.
Public Sub updateFont()

    'Free our current cache copy.  Uniscribe will automatically generate a new cache when a DC with the new font is passed to
    ' step 3 or 4, below.
    freeUniscribeCaches
    
    'Reset our script property tracker, so it can be refreshed on the next call
    m_ScriptFontProperties.cBytes = 0
    
End Sub

'Step1_ScriptItemize is the first step in processing a Uniscribe string.  It generates the crucial SCRIPT_ITEM array used for pretty much
' every subsequent Uniscribe interaction.  This array is cached internally, in m_ScriptItemsCache().
'
'Returns TRUE if successful; FALSE otherwise.
Public Function Step1_ScriptItemize(ByRef srcString As String) As Boolean
    
    m_ScriptItemCacheOK = False
    
    Dim i As Long
    
    'Make a deep copy of the source string
    m_CurrentString = srcString
    
    'Uniscribe does not handle 0-length strings, so make sure at least one character is present
    If Len(m_CurrentString) < 1 Then m_CurrentString = " "
    
    'While here, check for linebreaks.  If any exist, we need to normalize them to vbLf, make a note of their positions, then remove
    ' them from the string.  (Uniscribe doesn't handle linebreaks.)
    m_StringHasLinebreaks = False
    
    If InStr(1, m_CurrentString, vbCrLf, vbBinaryCompare) <> 0 Then
        m_StringHasLinebreaks = True
        m_CurrentString = Replace$(m_CurrentString, vbCrLf, vbLf, , , vbBinaryCompare)
        
    ElseIf InStr(1, m_CurrentString, vbCr, vbBinaryCompare) <> 0 Then
        m_StringHasLinebreaks = True
        m_CurrentString = Replace$(m_CurrentString, vbCr, vbLf, , , vbBinaryCompare)
        
    ElseIf InStr(1, m_CurrentString, vbLf, vbBinaryCompare) <> 0 Then
        m_StringHasLinebreaks = True
    End If
    
    'If linebreaks were found, apply relevant pre-processing now.
    If m_StringHasLinebreaks Then
        
        Dim startPos As Long
        startPos = 1
        
        If m_NumOfLineBreaks = 0 Then ReDim m_LineBreakPositions(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Long
        
        m_NumOfLineBreaks = 0
        
        'Find the first linefeed in the string
        Dim lfPos As Long
        lfPos = InStr(1, m_CurrentString, vbLf, vbBinaryCompare)
        
        'Normalize everything to vbLf
        Do While lfPos <> 0
        
            m_LineBreakPositions(m_NumOfLineBreaks) = lfPos
            m_NumOfLineBreaks = m_NumOfLineBreaks + 1
            
            If m_NumOfLineBreaks > UBound(m_LineBreakPositions) Then ReDim Preserve m_LineBreakPositions(0 To m_NumOfLineBreaks * 2 - 1) As Long
            
            lfPos = InStr(lfPos + 1, m_CurrentString, vbLf, vbBinaryCompare)
        
        Loop
        
        'For debugging purposes, you can choose to list calculated line break character positions:
        'For i = 0 To m_NumOfLineBreaks - 1
        '    Debug.Print "ScriptItemize reports linebreak at char # " & m_LineBreakPositions(i)
        'Next i
        
        'Replace all linebreaks with zero-width space chars.  These force Uniscribe to break characters at these positions, and they are
        ' easy to detect later, when it's time to return a glyph array to the actual renderer.
        m_CurrentString = Replace$(m_CurrentString, vbLf, ChrW(8203), , , vbBinaryCompare)
        
    End If
    
    
    'Values we need to manually set:
    ' pwcInChars [In]: Pointer to a Unicode string to itemize.
    ' cInChars [In]: Number of characters in pwcInChars to itemize.
    ' cMaxItems [In]: Maximum number of SCRIPT_ITEM structures defining items to process.
    ' psControl [in, optional]: Pointer to a SCRIPT_CONTROL structure indicating the type of itemization to perform.
    '                           Alternatively, the application can set this parameter to NULL if no SCRIPT_CONTROL properties are needed.
    ' psState [in, optional]: Pointer to a SCRIPT_STATE structure indicating the initial bidirectional algorithm state.
    '                         Alternatively, the application can set this parameter to NULL if the script state is not needed.
    ' pItems [out]: Pointer to a buffer in which the function retrieves SCRIPT_ITEM structures representing the items that have been processed.
    '               The buffer should be (cMaxItems + 1) * sizeof(SCRIPT_ITEM) bytes in length. It is invalid to call this function with a
    '               buffer to hold less than two SCRIPT_ITEM structures. The function always adds a terminal item to the item analysis array
    '               so that the length of the item with zero-based index "i" is always available as:
    '               pItems[i+1].iCharPos - pItems[i].iCharPos;
    ' pcItems [out]: Pointer to the number of SCRIPT_ITEM structures processed.
    
    'Determine the maximum number of SCRIPT_ITEM structures to process.  This is an arbitrary value; MSDN says "The function returns E_OUTOFMEMORY
    ' if the value of cMaxItems is insufficient. As in all error cases, no items are fully processed and no part of the output array contains
    ' defined values. If the function returns E_OUTOFMEMORY, the application can call it again with a larger pItems buffer."
    '
    'Because PD doesn't work with particularly large strings (e.g. we're not MS Office), we err on the side of safety and use
    ' a very large buffer.  This may still be insufficient, because Uniscribe sometimes requires a much larger buffer than it even requires
    ' (e.g. it demands a buffer of 10 but returns it filled with 3 items), but short of gaining psychic powers, there's not much more we
    ' can do.
    '
    'Also, note that the buffer itself must be one larger than the number of SCRIPT_ITEM values passed, and it can never be less than two.
    Dim numScriptItems As Long
    numScriptItems = Len(m_CurrentString) * 2
    
    If UBound(m_ScriptItemsCache) < numScriptItems Then ReDim m_ScriptItemsCache(0 To numScriptItems) As SCRIPT_ITEM
    
    'SCRIPT_CONTROL and SCRIPT_STATE primarily deal with extremely technical details of localization and glyph runs.  Most of the settings
    ' they contrl are not really relevant to our usage in PD, because they involve low-level instructions for how to handle embedded data.
    ' Whenever possible, we prefer to let Uniscribe use its own internal heuristics to determine things like RTL runs inside LTR paragraphs,
    ' rather than arbitrarily forcing everything to a particular direction or script algorithm.
    '
    'For more details, see:
    ' SCRIPT_CONTROL: https://msdn.microsoft.com/en-us/library/windows/desktop/dd368800%28v=vs.85%29.aspx
    ' SCRIPT_STATE: https://msdn.microsoft.com/en-us/library/windows/desktop/dd374043%28v=vs.85%29.aspx
    Dim dummyScriptControl As SCRIPT_CONTROL
    Dim dummyScriptState As SCRIPT_STATE
    
    'Tell Uniscribe to use the default user language as the basis for its heuristics.  Note that this still allows things like LTR inside
    ' RTL runs, but for ambiguous data, we assume behavior corresponding to the current user default locale.
    dummyScriptControl.uDefaultLanguage = LANG_USER_DEFAULT
    
    'The one other flag we set is fMergeNeutralItems.  This tells the shaping engine to keep contiguous scripts together, rather than
    ' separating them at punctuation marks (like commas and strings).  This improves performance by reducing the number of separate calls
    ' to subsequent functions like ScriptShape, ScriptPlace, etc., if the user is writing in a single script (e.g. just normal en-US text).
    dummyScriptControl.fBitFields2 = 1
    
    Dim numItemsFilled As Long
    
    Dim unsReturn As HRESULT
    unsReturn = ScriptItemize(StrPtr(m_CurrentString), Len(m_CurrentString), numScriptItems, VarPtr(dummyScriptControl), VarPtr(dummyScriptState), VarPtr(m_ScriptItemsCache(0)), numItemsFilled)
    
    'Account for potential out of memory errors
    Do While unsReturn = E_OUTOFMEMORY
        
        'Double the allotted cache size and try again
        numScriptItems = numScriptItems * 2
        ReDim m_ScriptItemsCache(0 To numScriptItems - 1) As SCRIPT_ITEM
        unsReturn = ScriptItemize(StrPtr(m_CurrentString), Len(m_CurrentString), numScriptItems, VarPtr(dummyScriptControl), VarPtr(dummyScriptState), VarPtr(m_ScriptItemsCache(0)), numItemsFilled)
        
    Loop
    
    'If ScriptItemize still failed, it was not due to memory errors
    If unsReturn = S_OK Then
        
        Debug.Print "Uniscribe itemized the string correctly.  (" & numItemsFilled & " SCRIPT_ITEM structs were filled.)"
        
        'Technically, we could trim our cache to its relevant size here (remembering that we MUST leave a spare entry at the end), but throughout
        ' this class I allow cache sizes to float.  Redimming them saves a tiny amount of memory at great cost to churn and performance, so instead,
        ' we simply rely on the returned size to let us know what chunk of the cache is relevant.
        'ReDim Preserve m_ScriptItemsCache(0 To numItemsFilled) As SCRIPT_ITEM
        
        'We are now going to mirror the contents of the item cache into our pdUniscribeItem array.  While this might seem like
        ' unnecessary duplication, that array is important for storing other Uniscribe data assembled on a per-item basis.
        m_NumOfItems = numItemsFilled
        If UBound(m_Items) < numItemsFilled Then ReDim m_Items(0 To numItemsFilled) As pdUniscribeItem
        
        Dim endPos As Long
        
        For i = 0 To numItemsFilled
            Set m_Items(i) = New pdUniscribeItem
            m_Items(i).setScriptItem m_ScriptItemsCache(i)
            
            'While we're here, it's also helpful to copy the relevant portion of the target string into each object.
            '
            '(Uniscribe has broken our input string according to a variety of factors (punctuation, control chars, scripts, etc),
            ' so each item's SCRIPT_ITEM settings must be applied only to a portion of the full string we were handed.)
            '
            'Note also that we don't process this for the last item, as we'd get an OOB error from the "index + 1" access
            If i < numItemsFilled Then
            
                startPos = m_ScriptItemsCache(i).iCharPos
                endPos = m_ScriptItemsCache(i + 1).iCharPos
                m_Items(i).setSubstring Mid$(m_CurrentString, startPos + 1, endPos - startPos), startPos + 1, endPos
                
                'Curious about how the items are broken up by Uniscribe?  Uncomment this to see substrings and positions.
                'Debug.Print m_Items(i).getSubstring, startPos + 1, endPos
                
            End If
            
        Next i
        
        'Itemizing is now complete
        Step1_ScriptItemize = True
        
    Else
        Debug.Print "WARNING!  ScriptItemize failed with code " & unsReturn & ".  Please investigate."
        Step1_ScriptItemize = False
    End If
    
    'Cache the success/failure value, so subsequent calls don't error out
    m_ScriptItemCacheOK = Step1_ScriptItemize

End Function

'Step2_ScriptLayout is the second step in processing a Uniscribe string.  (Actually, there could be another step before this, where we
' manually break the results of Step1 into more fine-grained runs, accounting for differences in font.  PD doesn't support this right now
' so I don't provide a wrapper for that.)  Step 2 takes the "runs" generated by Step 1, and converts them from logical input order to
' visual output order.  This is crucial when intermixing LTR and RTL text, as the order in which characters are entered may be totally
' different from the order in which they are displayed.
'
'Note that this function takes no inputs; it relies entirely on the output of Step 1 for its behavior.
'
'Returns TRUE if successful; FALSE otherwise.  (A failure at step 1 that was not dealt with by the caller will cause this function to
' return FALSE, as it relies on the output of Step 1 to generate a correct layout.)
Public Function Step2_ScriptLayout() As Boolean
    
    'If a SCRIPT_ITEM cache does not exist, exit now
    If Not m_ScriptItemCacheOK Then
        Debug.Print "WARNING!  Step 1 failed, so Step 2 cannot proceed!"
        Step2_ScriptLayout = False
        Exit Function
    End If
    
    'Values we need to manually set:
    ' cRuns [In]: Number of runs to process.
    ' pbLevel [In]: Pointer to an array, of length indicated by cRuns, containing run embedding levels. Embedding levels for all runs on
    '               the line must be included, ordered logically.
    ' piVisualToLogical [out, optional]: Pointer to an array, of length indicated by cRuns, in which this function retrieves the run
    ' embedding levels reordered to visual order. The first array element represents the run to display at the far left, and subsequent
    ' entries should be displayed progressing from left to right. The function sets this parameter to NULL if there is no output.
    ' piLogicalToVisual [out, optional]: Pointer to an array, of length indicated by cRuns, in which this function retrieves the visual
    ' run positions. The first array element is the relative visual position where the first logical run should be displayed, the leftmost
    ' display position being 0. The function sets this parameter to NULL if there is no output.
    
    'The number of runs to process is simply the length of the SCRIPT_ITEM cache from step 1
    Dim numOfRuns As Long
    numOfRuns = m_NumOfItems + 1
    
    'Run embedding levels were automatically calculated by Step 1.  These values are buried deep within the item cache, unfortunately,
    ' so we need to retrieve them and store them in their own array.
    '
    'Also, I'm not sure if this structure needs to be DWORD-aligned, but it doesn't hurt to enforce it
    Dim elBound As Long
    elBound = (numOfRuns + 3) And &HFFFFFFFC
    
    Dim embeddingLevels() As Byte
    ReDim embeddingLevels(0 To elBound - 1) As Byte
    
    Dim tmpPBLevel As Long, extractedPBLevel As Long, tmpFlag As Boolean
    
    Dim i As Long
    For i = 0 To numOfRuns - 1
        
        'Retrieving the embedded bidi level is ugly, to put it kindly.  The top 5 bits of the embedded SCRIPT_STATE value for each run
        ' contain LTR vs RTL information.  There is no good way to retrieve this data in VB, so we are lazy and just do it manually.
        '
        'Start by copying the byte into a Long, which makes retrieval easier.
        tmpPBLevel = m_ScriptItemsCache(i).analysis.s.fBitFields1
        
        'Technically the bitfield contains 5 bits for potential values on the range [0, 31].  But there are really only values on the
        ' range [0, 3], basically every combination of LTR and RTL text embedded within each other.  As such, we only need to retrieve
        ' the bottom 2 of 5 bytes.
        tmpFlag = getBitFlag(0, tmpPBLevel)
        If tmpFlag Then extractedPBLevel = 1 Else extractedPBLevel = 0
        
        tmpFlag = getBitFlag(1, tmpPBLevel)
        If tmpFlag Then extractedPBLevel = extractedPBLevel + 2
        
        'Curious about the bidi levels we've extracted?  Display our analysis using this line of code:
        'Debug.Print i & ": RTL marked as " & extractedPBLevel
        
        'As an additional confirmation, you can read the RTL bit from the SCRIPT_ANALYSIS portion of the item cache.
        ' If the bidi level is marked as 1, this value should also be 1/TRUE, which provides a nice secondary confirmation.
        'Debug.Print "RTL? " & CStr(getBitFlag(2, m_ScriptItemsCache(i).analysis.fBitFields2))
        
        'Store the calculated value in our embeddingLevels() result array
        embeddingLevels(i) = extractedPBLevel
        
    Next i
    
    'With our bidi levels set, we can now determine the visual order of runs
    
    'Prep conversion arrays
    ReDim m_VisualToLogicalOrder(0 To numOfRuns - 1) As Long
    ReDim m_LogicalToVisualOrder(0 To numOfRuns - 1) As Long
    
    Dim ptrToVLO As Long, ptrToLVO As Long
    ptrToVLO = VarPtr(m_VisualToLogicalOrder(0))
    ptrToLVO = VarPtr(m_LogicalToVisualOrder(0))
    
    'Retrieve the final layout order of the runs
    Dim unsReturn As HRESULT
    unsReturn = ScriptLayout(numOfRuns, VarPtr(embeddingLevels(0)), ptrToVLO, ptrToLVO)
    
    'Again, debugging output can be helpful:
    'For i = 0 To numOfRuns - 1
    '    Debug.Print i & ":" & m_VisualToLogicalOrder(i) & ":" & m_LogicalToVisualOrder(i)
    'Next i
    
    If unsReturn = S_OK Then
        
        'In the future, we may need to look at reordering our m_Items() array to match the order returned by this function.
        ' I haven't tested complex enough test to know for sure, but if that becomes necessary, this section of code is where
        ' you'd want to reorder them accordingly.
        
        Debug.Print "Uniscribe laid out the string correctly."
        Step2_ScriptLayout = True
        
    Else
        Debug.Print "WARNING!  ScriptLayout failed with code " & unsReturn & ".  Please investigate."
        Step2_ScriptLayout = False
    End If
    
End Function

'Step3_ScriptShape is the third step in processing a Uniscribe string (and arguably the most important!).  ScriptShape is the far more
' powerful Uniscribe version of GDI's GetCharacterPlacement function.  Basically, it operates on a single item, and does all the messy
' behind-the-scene work to convert character values into glyph indices of the current font.  As such, it is the first step to require
' a DC, and you must (obviously) have the relevant font selected into the DC *prior* to calling this!
'
'Because this function only operates on a single SCRIPT_ITEM, it must iterate through the contents of our m_Items() array.
' The results of each pass are stored inside the respective m_Items object, so for any subsequent steps, you *must* rely on the
' contents of those objects instead of any function-specific arrays or returns.
'
'NOTE: in the future, I would very much like to have two codepaths here: one that uses the old ScriptShape function (required by XP),
' and a new one that uses ScriptShapeOpenType.  This would allow us to take full advantage of OpenType's many awesome features -
' but right now, I mostly just want to get the damn thing working.
'
'ANOTHER NOTE: if ScriptShape returns USP_E_SCRIPT_NOT_IN_FONT, it means this item contains characters that the current font doesn't
' support.  If we want to implement font fallback (and I do, it's just crazy cumbersome), this is where we should check for failure,
' as this step is the first one to use font-specific information for its decision-making.  Depending on our cleverness, we could do
' something like check the Unicode ranges involved, and test a few fallback fonts accordingly - but this would require heavy interaction
' with the caller, as only they have knowledge of font style settings.
'
'Returns TRUE if successful; FALSE otherwise.  (A failure at step 1 or 2 that was not dealt with by the caller will cause this function to
' return FALSE, as it relies on the output of Steps 1 and 2 to generate correct shapes.)
Public Function Step3_ScriptShape(ByRef srcDC As Long) As Boolean
    
    'If a SCRIPT_ITEM cache does not exist, exit now
    If Not m_ScriptItemCacheOK Then
        Debug.Print "WARNING!  Step 1 or 2 failed, so Step 3 cannot proceed!"
        Step3_ScriptShape = False
        Exit Function
    End If
        
    'Values we must generate prior to calling the API:
    ' hDC [In]: Handle to the device context. For more information, see Caching.
    ' psc [in, out]: Pointer to a SCRIPT_CACHE structure identifying the script cache.
    ' pwcChars [In]: Pointer to an array of Unicode characters defining the run.
    ' cChars [In]: Number of characters in the Unicode run.
    ' cMaxGlyphs [In]: Maximum number of glyphs to generate, and the length of pwOutGlyphs. A reasonable value is (1.5 * cChars + 16),
    '                  but this value might be insufficient in some circumstances.
    ' psa [in, out]: Pointer to the SCRIPT_ANALYSIS structure for the run, containing the results from an earlier call to ScriptItemize.
    ' pwOutGlyphs [out]: Pointer to a buffer in which this function retrieves an array of glyphs with size as indicated by cMaxGlyphs.
    ' pwLogClust [out]: Pointer to a buffer in which this function retrieves an array of logical cluster information. Each array element
    '                   corresponds to a character in the array of Unicode characters; therefore this array has the number of elements
    '                   indicated by cChars. The value of each element is the offset from the first glyph in the run to the first glyph
    '                   in the cluster containing the corresponding character. Note that, when the fRTL member is set to TRUE in the
    '                   SCRIPT_ANALYSIS structure, the elements decrease as the array is read.
    ' psva [out]: Pointer to a buffer in which this function retrieves an array of SCRIPT_VISATTR structures containing visual attribute
    '             information. Since each glyph has only one visual attribute, this array has the number of elements indicated by cMaxGlyphs.
    ' pcGlyphs [out]: Pointer to the location in which this function retrieves the number of glyphs indicated in pwOutGlyphs.
    
    'IMPORTANT NOTE: this function operates on *a single item at a time*.  This means we can't use module-level values like
    ' m_CurrentString - instead, we must iterate through each m_Item() object and process it individually.
    
    'This function returns success only if all items are successfully processed.  One or more failures results in a FALSE return,
    ' even if some items were processed successfully.  In the future, the caller might use this for font fallback purposes.
    Dim functionSuccess As Boolean
    functionSuccess = True
    
    Dim lSubstring As Long, sSubstring As String
    Dim cMaxGlyphs As Long
    Dim unsReturn As HRESULT
    
    'Start iterating through our item collection
    Dim i As Long, j As Long
    For i = 0 To m_NumOfItems - 1
    
        'A number of calculations are relative to the length of this item's substring.  Cache that value now.
        sSubstring = m_Items(i).getSubstring
        lSubstring = Len(sSubstring)
    
        'Determine an intial size for this item's glyph cache.  I use the MSDN-recommended formula, which is filled with magic numbers.
        cMaxGlyphs = 1.5 * CDbl(lSubstring) + 16
        
        'Prep the logical cluster cache, which has the same length as cChars.  This cache gives us a way to map between characters
        ' and glyphs, albeit confusingly.  Basically, this value is the offset between the first glyph in the item, and the glyph
        ' containing the corresponding character.
        ReDim m_LogicalClusterCache(0 To Len(m_CurrentString) - 1) As Integer
        
        'We now start a retrieval loop; despite our best efforts, ScriptShape may still require a larger buffer to return all its data,
        ' so we loop until something other than E_OUTOFMEMORY is returned.
        Do
            
            'Prep all glyph-specific caches to be the same size as cMaxGlyphs
            If UBound(m_GlyphCache) < cMaxGlyphs - 1 Then
                ReDim m_GlyphCache(0 To cMaxGlyphs - 1) As Integer
                ReDim m_VisualAttributesCache(0 To cMaxGlyphs - 1) As SCRIPT_VISATTR
            End If
            
            unsReturn = ScriptShape(srcDC, VarPtr(m_ScriptCache), StrPtr(sSubstring), lSubstring, cMaxGlyphs, m_Items(i).getScriptItemAnalysisPointer, VarPtr(m_GlyphCache(0)), VarPtr(m_LogicalClusterCache(0)), VarPtr(m_VisualAttributesCache(0)), m_NumOfGlyphs)
            
            'Because cMaxGlyphs initialization is a guessing game, ScriptShape will return an out of memory code if the buffer proves insufficient.
            ' If this happens, we double the size of the buffer and try again.
            If unsReturn = E_OUTOFMEMORY Then cMaxGlyphs = cMaxGlyphs * 2
            
        Loop While unsReturn = E_OUTOFMEMORY
                
        'Return success/failure
        If unsReturn = S_OK Then
        
            'm_NumOfGlyphs contains the number of glyphs generated by ScriptShape.  Forward the results to this Item object; it will extract
            ' only the entries it requires, so we don't have to worry about trimming array bounds.
            m_Items(i).setShapingData m_LogicalClusterCache, m_NumOfGlyphs, m_GlyphCache, m_VisualAttributesCache
        
        Else
        
            functionSuccess = False
            
            'In the future, we may be able to deal with some failure states
            If unsReturn = USP_E_SCRIPT_NOT_IN_FONT Then
            
                'If our SCRIPT_FONTPROPERTIES cache is outdated, retrieve a new set of font properties.  This tells us which glyphs to use
                ' for missing characters.
                If m_ScriptFontProperties.cBytes = 0 Then
                    m_ScriptFontProperties.cBytes = Len(m_ScriptFontProperties)
                    ScriptGetFontProperties srcDC, VarPtr(m_ScriptCache), VarPtr(m_ScriptFontProperties)
                End If
                
                'We are now going to disable shaping and call ScriptShape again; this will remove any character-specific measurements, and instead
                ' fill the glyph array with "missing glyph" markers.
                m_Items(i).disableAllShaping
                ScriptShape srcDC, VarPtr(m_ScriptCache), StrPtr(sSubstring), lSubstring, cMaxGlyphs, m_Items(i).getScriptItemAnalysisPointer, VarPtr(m_GlyphCache(0)), VarPtr(m_LogicalClusterCache(0)), VarPtr(m_VisualAttributesCache(0)), m_NumOfGlyphs
                
                'Forward the updated results to this Item object
                m_Items(i).setShapingData m_LogicalClusterCache, m_NumOfGlyphs, m_GlyphCache, m_VisualAttributesCache
                
                Debug.Print "WARNING!  ScriptShape could not find some glyphs in this font.  Recommend font fallback."
                
            Else
                Debug.Print "WARNING!  ScriptShape returned failure code #" & Hex$(unsReturn) & " on item # " & i & ".  Please investigate."
            End If
            
        End If
    
    'Repeat all the above steps for the next item
    Next i
    
    
    If functionSuccess Then Debug.Print "Uniscribe shaped the string correctly."
    Step3_ScriptShape = functionSuccess
    
    
End Function

'Step4_ScriptPlace is the fourth step in processing a Uniscribe string.  It could technically be called Step 3.5, as it operates
' directly on the results of ScriptShape; in fact, I may merge the two functions in the future.
'
'While ScriptShape converts characters to glyphs, ScriptPlace calculates positioning for all glyphs.  Once again, the choice of
' font is crucial, so this step also requires a DC (and you must have the relevant font selected into the DC *prior* to calling this!)
'
'Because this function only operates on a single SCRIPT_ITEM, it must iterate through the contents of our m_Items() array.
' The results of each pass are stored inside the respective m_Items object, so for any subsequent steps, you *must* rely on the
' contents of those objects instead of any function-specific arrays or returns.
'
'NOTE: in the future, I would very much like to have two codepaths here: one that uses the old ScriptPlace function, and a new one that
' uses ScriptPlaceOpenType.  This would allow us to take full advantage of OpenType's many awesome features - but right now, I mostly
' just want to get the damn thing working.
'
'Returns TRUE if successful; FALSE otherwise.  (A failure at step 1 or 2 that was not dealt with by the caller will cause this function to
' return FALSE, as it relies on the output of Steps 1 and 2 to generate correct glyph placement.)
Public Function Step4_ScriptPlace(ByRef srcDC As Long) As Boolean

    'If a SCRIPT_ITEM cache does not exist, exit now
    If Not m_ScriptItemCacheOK Then
        Debug.Print "WARNING!  Step 1 or 2 failed, so Step 4 cannot proceed!"
        Step4_ScriptPlace = False
        Exit Function
    End If
    
    'Values we must generate prior to calling the API:
    ' hDC [In]: Handle to the device context. For more information, see Caching.
    ' psc [in, out]: Pointer to a SCRIPT_CACHE structure identifying the script cache.
    ' pwGlyphs [In]: Pointer to a glyph buffer obtained from an earlier call to the ScriptShape function.
    ' cGlyphs [In]: Count of glyphs in the glyph buffer.
    ' psva [In]: Pointer to an array of SCRIPT_VISATTR structures indicating visual attributes.
    ' psa [in, out]: Pointer to a SCRIPT_ANALYSIS structure. On input, this structure is obtained from a previous call to ScriptItemize.
    '                On output, this structure contains values retrieved by ScriptPlace.
    ' piAdvance [out]: Pointer to an array in which this function retrieves advance width information.
    ' pGoffset [out]: Optional. Pointer to an array of GOFFSET structures in which this function retrieves the x and y offsets of
    '                 combining glyphs. This array must be of length indicated by cGlyphs.
    ' pABC [out]: Pointer to an ABC structure in which this function retrieves the ABC width for the *entire run*.
    
    'IMPORTANT NOTE: this function operates on *a single item at a time*.  This means we can't use module-level values like
    ' m_CurrentString - instead, we must iterate through each m_Item() object and process it individually.
    
    'This function returns success only if all items are successfully processed.  One or more failures results in a FALSE return,
    ' even if some items were processed successfully.  In the future, the caller might use this for font fallback purposes.
    Dim functionSuccess As Boolean
    functionSuccess = True
    
    'Prep a single ABC struct; ScriptPlace fills this with the ABC width for each individual entire item AS A WHOLE.  ABC widths for individual
    ' characters would need to be retrieved separately, using custom code and ScriptGetGlyphABCWidth.
    Dim tmpABC As ABC
    Dim numOfGlyphs As Long
    Dim unsReturn As HRESULT
    
    'Start iterating through our item collection
    Dim i As Long
    For i = 0 To m_NumOfItems - 1
    
        'Prep the offset and advance width caches.  Note that we only resize these as necessary; we know the relevant number of glyphs, so an
        ' over-large buffer isn't a problem (and it prevents us from wasting cycles reallocating memory).
        numOfGlyphs = m_Items(i).getNumOfGlyphs
        
        If numOfGlyphs > 0 Then
        
            If UBound(m_GlyphOffsetCache) < numOfGlyphs Then
                
                'Technically these only need to be resized to numOfGlyphs - 1, FYI
                ReDim m_GlyphOffsetCache(0 To numOfGlyphs) As GOFFSET
                ReDim m_AdvanceWidthCache(0 To numOfGlyphs) As Long
            
            End If
            
            'This function retrieves a bunch of bare pointers from the relevant m_Items() object.  This is generally a bad idea in VB,
            ' but it saves us having to make expensive copies of large structs.
            With m_Items(i)
                unsReturn = ScriptPlace(srcDC, m_ScriptCache, .getPointerToGlyphCache, numOfGlyphs, .getPointerToVisualAttributesCache, .getScriptItemAnalysisPointer, VarPtr(m_AdvanceWidthCache(0)), VarPtr(m_GlyphOffsetCache(0)), tmpABC)
            End With
            
            'If the return was successful, copy all retrieved data into the item container
            If unsReturn = S_OK Then
                m_Items(i).setPlacementData m_AdvanceWidthCache, m_GlyphOffsetCache, tmpABC
            Else
            
                functionSuccess = False
                
                'In the future, we may be able to deal with some failure states
                If unsReturn = USP_E_SCRIPT_NOT_IN_FONT Then
                    m_Items(i).setPlacementData m_AdvanceWidthCache, m_GlyphOffsetCache, tmpABC
                    Debug.Print "WARNING!  ScriptPlace could not find some glyphs in this font.  Recommend font fallback."
                Else
                    Debug.Print "WARNING!  ScriptPlace returned failure code #" & Hex$(unsReturn) & " on item # " & i & ".  Please investigate."
                End If
                
            End If
    
        End If
    
    Next i
    
    'Return success/failure
    If functionSuccess Then Debug.Print "Uniscribe placed the string correctly."
    Step4_ScriptPlace = functionSuccess

End Function

'Step5_ScriptBreak is the fifth step in processing a Uniscribe string.  It could technically happen earlier in the process, as it only
' calculates *potential* break positions, but I find it helpful to do this after glyphs have been generated, so we can map the break points
' directly to the list of generated glyphs.
'
'ScriptBreak does not require a DC or a string.  It operates directly on the information generated by previous runs.
'
'Because this function only operates on a single SCRIPT_ITEM, it must iterate through the contents of our m_Items() array.
' The results of each pass are stored inside the respective m_Items object, so for any subsequent steps, you *must* rely on the
' contents of those objects instead of any function-specific arrays or returns.
'
'Returns TRUE if successful; FALSE otherwise.  (A failure at step 1 or 2 that was not dealt with by the caller will cause this function to
' return FALSE, as it relies on the output of Steps 1 and 2 to generate correct breaks.)
Public Function Step5_ScriptBreak() As Boolean

    'If a SCRIPT_ITEM cache does not exist, exit now
    If Not m_ScriptItemCacheOK Then
        Debug.Print "WARNING!  Step 1 or 2 failed, so Step 5 cannot proceed!"
        Step5_ScriptBreak = False
        Exit Function
    End If
    
    'Values we must generate prior to calling the API:
    ' pwcChars [In]: Pointer to the Unicode characters to process.
    ' cChars [In]: Number of Unicode characters to process.
    ' psa [In]: Pointer to the SCRIPT_ANALYSIS structure obtained from an earlier call to ScriptItemize.
    ' psla [out]: Pointer to a buffer in which this function retrieves the character attributes as a SCRIPT_LOGATTR structure.
    
    'IMPORTANT NOTE: this function operates on *a single item at a time*.  This means we can't use module-level values like
    ' m_CurrentString - instead, we must iterate through each m_Item() object and process it individually.
    
    'This function returns success only if all items are successfully processed.  One or more failures results in a FALSE return,
    ' even if some items were processed successfully.  In the future, the caller might use this for font fallback purposes.
    Dim functionSuccess As Boolean
    functionSuccess = True
    
    Dim lSubstring As Long, sSubstring As String
    Dim unsReturn As HRESULT
    
    'Start iterating through our item collection
    Dim i As Long
    For i = 0 To m_NumOfItems - 1
        
        'Retrieve local copies of the relevant substring and length
        sSubstring = m_Items(i).getSubstring
        lSubstring = Len(sSubstring)
        
        If lSubstring > 0 Then
            
            'Like other caches in this class, PD doesn't trim them, and it only expands them as necessary.  This is important for performance.
            If UBound(m_CharAttributesCache) < lSubstring Then ReDim m_CharAttributesCache(0 To lSubstring) As SCRIPT_LOGATTR
            
            unsReturn = ScriptBreak(StrPtr(sSubstring), lSubstring, m_Items(i).getScriptItemAnalysisPointer, VarPtr(m_CharAttributesCache(0)))
            
            'If the return was successful, copy all retrieved data into the item container
            If unsReturn = S_OK Then
                m_Items(i).setCharacterAttributeCache m_CharAttributesCache
            Else
                
                functionSuccess = False
                
                'Because ScriptBreak operates is a pass/fail API, we don't have to worry about crap like guessing buffer sizes.
                ' Failure is always unacceptable (and shouldn't happen unless the string contains only gibberish).
                Debug.Print "WARNING!  ScriptBreak returned failure code #" & Hex$(unsReturn) & " on item # " & i & ".  Please investigate."
                
            End If
                        
        End If
    
    Next i
    
    'Return success/failure
    If functionSuccess Then Debug.Print "Uniscribe calculated word breaks correctly."
    Step5_ScriptBreak = functionSuccess

End Function

'Retrieve a copy of the currently calculated glyph cache, in a custom PD format.
'
'If a source DC is passed, the function will automatically fill the ABC width of each glyph prior to returning it.
'
' Returns a value >= 0 indicating the number of glyphs in the returned struct.  0 indicates failure or an empty string.
Friend Function getCopyOfGlyphCache(ByRef dstGlyphArray() As pdGlyphUniscribe, Optional ByVal srcDC As Long = 0) As Long
    
    'Before doing anything else, we need to calculate a total number of glyphs in our current collection.
    Dim totalGlyphCount As Long
    totalGlyphCount = 0
    
    Dim i As Long, j As Long
    For i = 0 To m_NumOfItems - 1
        totalGlyphCount = totalGlyphCount + m_Items(i).getNumOfGlyphs
    Next i
    
    'If the buffer is not large enough to hold our full glyph list, forcibly enlarge it, and leave some dead space for
    ' future entries, to reduce the need for resizes in the future.
    If UBound(dstGlyphArray) < totalGlyphCount - 1 Then
        ReDim dstGlyphArray(0 To totalGlyphCount * 2 - 1) As pdGlyphUniscribe
    End If
    
    'Next, iterate all items and copy the relevant glyph information from each item into the destination array
    Dim netGlyphOffset As Long, numOfGlyphsAtStartOfLine As Long
    netGlyphOffset = 0
    
    'DEBUG ONLY!
    'For i = 0 To m_NumOfItems - 1
    '    For j = 0 To m_Items(i).getNumOfGlyphs - 1
    '        Debug.Print "Post-Uniscribe linebreak: " & i & ", " & j & ", " & m_Items(i).getCorrespondingCharIndex(j)
    '        netGlyphOffset = netGlyphOffset + 1
    '    Next j
    'Next i
    '
    'netGlyphOffset = 0
    
    'END DEBUG ONLY!
    
    For i = 0 To m_NumOfItems - 1
    
        For j = 0 To m_Items(i).getNumOfGlyphs - 1
            
            'See if this is a zero-width glyph.  We can reduce render time by leaving out zero-width glyphs that don't convey
            ' useful information to the renderer.
            If m_Items(i).getZeroWidthAtPosition(j) Then
                
                'Debug.Print "zero-width: " & netGlyphOffset & ", char-index: " & m_Items(i).getCorrespondingCharIndex(j) & ", start/end: " & m_Items(i).getRelativeStartPos & "," & m_Items(i).getRelativeEndPos
                
                'Skip further processing if linebreaks weren't found in the original string
                If m_StringHasLinebreaks Then
                    
                    'See if the character position associated with this glyph was marked as a line-break character.
                    If isCharPositionLinebreak(m_Items(i).getCorrespondingCharIndex(j)) Then
                        
                        'Fill this glyph with dummy information, while taking care to manually mark it as a hard line break
                        With dstGlyphArray(netGlyphOffset)
                            .glyphIndex = 0
                            .AdvanceWidth = 0
                            .isZeroWidth = True
                            .isSoftBreak = True
                            .isWhiteSpace = True
                            .isHardLineBreak = True
                            .GlyphOffset.du = 0
                            .GlyphOffset.dv = 0
                            .ABCWidth.abcA = 0
                            .ABCWidth.abcB = 0
                            .ABCWidth.abcC = 0
                            .isFirstGlyphOnLine = False
                            .isLastGlyphOnLine = False
                        End With
                        
                        netGlyphOffset = netGlyphOffset + 1
                        
                    End If
                
                End If
            
            'This is not a zero-width glyph.  Treat it normally.
            Else
            
                With dstGlyphArray(netGlyphOffset)
                    .glyphIndex = m_Items(i).getGlyphCacheAtPosition(j)
                    .AdvanceWidth = m_Items(i).getAdvanceWidthAtPosition(j)
                    .isZeroWidth = False
                    .isSoftBreak = m_Items(i).getSoftBreakAtPosition(j)
                    .isWhiteSpace = m_Items(i).getWhitespaceAtPosition(j)
                    .isHardLineBreak = False
                    .isFirstGlyphOnLine = False
                    .isLastGlyphOnLine = False
                    m_Items(i).copyGlyphOffsetToPointer VarPtr(.GlyphOffset), j
                    
                    'While here, we also collect the ABC width of each glyph
                    If srcDC <> 0 Then ScriptGetGlyphABCWidth srcDC, VarPtr(m_ScriptCache), .glyphIndex, VarPtr(.ABCWidth)
                    
                End With
                
                'Debug word breaks and/or whitespace with the following code:
                'If dstGlyphArray(netGlyphOffset).isSoftBreak Then Debug.Print "word break at glyph #" & netGlyphOffset
                'If dstGlyphArray(netGlyphOffset).isWhiteSpace Then Debug.Print "whitespace at glyph #" & netGlyphOffset
                
                netGlyphOffset = netGlyphOffset + 1
            
            End If
            
        Next j
    
    Next i
    
    getCopyOfGlyphCache = netGlyphOffset
    
End Function

'Sometimes it's helpful to take a close look at the current glyph cache.  This function will provide basic per-glyph info for the base run.
Public Sub printUniscribeDebugInfo()
    
    Dim tmpGlyphCache() As pdGlyphUniscribe
    Dim numOfGlyphs As Long
    numOfGlyphs = getCopyOfGlyphCache(tmpGlyphCache)
    
    Debug.Print "-- Glyph data returned by Uniscribe --"
    
    If numOfGlyphs > 0 Then
    
        Dim i As Long
        For i = 0 To numOfGlyphs - 1
            With tmpGlyphCache(i)
                Debug.Print i & ": " & .glyphIndex & " (" & .AdvanceWidth & ")  (" & .GlyphOffset.du & ", " & .GlyphOffset.dv & ")"
            End With
        Next i
        
    End If
    
    Debug.Print "-- End of glyph data  --"
    
End Sub

'See if a given character position was marked as a hard linebreak prior to processing
Private Function isCharPositionLinebreak(ByVal charPos As Long) As Boolean
    
    isCharPositionLinebreak = False
    
    Dim i As Long
    For i = 0 To m_NumOfLineBreaks - 1
        If charPos = m_LineBreakPositions(i) Then
            isCharPositionLinebreak = True
            Exit For
        End If
    Next i

End Function

Private Sub Class_Initialize()
    
    'Initialize all caches to small initial sizes; this prevents memory churn if the user is working with nice, small text layers
    ReDim m_ScriptItemsCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As SCRIPT_ITEM
    ReDim m_VisualToLogicalOrder(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Long
    ReDim m_LogicalToVisualOrder(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Long
    ReDim m_Items(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As pdUniscribeItem
    ReDim m_GlyphCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Integer
    ReDim m_LogicalClusterCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Integer
    ReDim m_VisualAttributesCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As SCRIPT_VISATTR
    ReDim m_AdvanceWidthCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As Long
    ReDim m_GlyphOffsetCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As GOFFSET
    ReDim m_CharAttributesCache(0 To DEFAULT_INITIAL_CACHE_SIZE - 1) As SCRIPT_LOGATTR
        
End Sub

Private Sub Class_Terminate()
    
    'Free any Uniscribe-specific caches we have created
    freeUniscribeCaches
    
End Sub
