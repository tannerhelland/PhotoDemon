VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFontCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Font Collection Manager
'Copyright 2014-2026 by Tanner Helland
'Created: 12/May/15
'Last updated: 29/October/15
'Last update: allow the caller to request their own font matching mode.  Some font caches in the program use a uniform font size,
'              but different font faces (e.g. font dropdowns).  Others use the same font face, but different font sizes and styles
'              (e.g. the central UI font cache).  This class can now handle these different situations elegantly.
'
'PD's font selection dropdowns must manage a (potentially) huge gallery of fonts.  It is not conducive to cache every font we use,
' so instead, we use a simple "round robin" approach where we keep a rotating cache of fonts, and free them in the order they
' were created if more fonts are required.
'
'This class wraps pdFont for specific font creation/destruction and rendering duties.
'
'Obviously, this class relies heavily on WAPI.  Functions are documented to the best of my knowledge and ability.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Different parts of PD have different caching requirements.  Sometimes, they need to cache fonts that differ only by name
' (but not by size).  Other times, fonts may have matching names but different sizes or styles.  To accelerate the
' cumbersome process of font-matching, the caller can specify a caching mode.
Public Enum FONT_CACHE_MODE
    FCM_NameOnly = 0
    FCM_SizeOnly = 1
    FCM_NameAndSize = 2
    FCM_SizeAndStyle = 3
    FCM_NameAndSizeAndStyle = 4
End Enum
Private m_CurrentCacheMode As FONT_CACHE_MODE

'Size of the font collection.  This is the maximum number of fonts this class is allowed to cache.  For scrollable
' UI elements with variable font faces (e.g. font dropdowns), this should be at least 3x the size of the dropped list;
' that gives us some breathing room to cache new fonts during mousewheel events.
Private Const DEFAULT_CACHE_SIZE As Long = 24
Private m_CacheSize As Long

'Pointer to the current font index.  This will rotate around to 0 after the cache is filled.
Private m_FontPointer As Long

'Once the cache has been completely filled, this will be set to TRUE.  A full cache starts using font access times
' to determine which fonts are dumped from the cache.
Private m_CacheIsFull As Boolean

'Actual font collection
Private m_FontCollection() As pdFont

'If/when the font collection hits its maximum size, we clear out the oldest fonts first when making room for
' new fonts.  (This is most relevant for font dropdowns, which consume an enormous amount of font objects.)
Private m_FontAccessTimes() As Currency

'This class can (optionally) manage extended font properties as well, like which scripts are supported by a given font.
' To activate this behavior, pass TRUE to the setExtendedPropertyCaching sub.
Private m_ExtendedPropertiesActive As Boolean
Private m_ExtendedProperties() As PD_FONT_PROPERTY

'Add a font to the cache.
' Returns: value >= 0, indicating the index of said font in the cache.  (This index can be subsequently used to retrieve the
'           actual handle or object.)
Friend Function AddFontToCache(ByRef fontName As String, ByVal FontSize As Single, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False) As Long
    
    'First, see if this font already exists in the cache
    Dim fontIndex As Long
    fontIndex = DoesFontExist(fontName, FontSize, isBold, isItalic, isUnderline)
    
    'If this font already exists in our collection, don't recreate it; instead, return its current index.
    ' IMPORTANT NOTE: the matching criteria used by DoesFontExist defaults to font-name matching only.
    '                  Call SetCacheMode() if you want to match fonts by additional criteria.
    If (fontIndex >= 0) Then
    
        AddFontToCache = fontIndex
        
        'Also, update the font's access time.  (Oldest fonts are dumped from the cache first, on the assumption
        ' that they are needed less than new fonts.)
        m_FontAccessTimes(fontIndex) = OS.GetSystemTimeAsCurrency()
        
    Else
        
        'This font does not already exist, meaning we need to create it anew.
        
        'If the cache is full, we need to figure out which font to erase.
        If m_CacheIsFull Then
            
            'Search the font collection and look for the largest delta between the current time, and each font's creation time.
            Dim curTime As Currency, thisDelta As Currency, maxDelta As Currency, maxIndex As Long
            curTime = OS.GetSystemTimeAsCurrency()
            
            Dim i As Long
            For i = 0 To m_CacheSize - 1
                thisDelta = curTime - m_FontAccessTimes(i)
                If (thisDelta > maxDelta) Then
                    maxIndex = i
                    maxDelta = thisDelta
                End If
            Next i
            
            'Kill the oldest font
            m_FontPointer = maxIndex
        
        'If the cache is not yet full, the font index pointer will already point at the best available spot.
        End If
        
        'Create a new font object as necessary
        If (m_FontCollection(m_FontPointer) Is Nothing) Then
            Set m_FontCollection(m_FontPointer) = New pdFont
        Else
            m_FontCollection(m_FontPointer).DeleteCurrentFont
        End If
        
        'Initialize all the required font properties
        m_FontCollection(m_FontPointer).SetFontPropsAllAtOnce fontName, FontSize, isBold, isItalic, isUnderline
        
        'Create the font object
        m_FontCollection(m_FontPointer).CreateFontObject
        
        'Update this font's access time
        m_FontAccessTimes(m_FontPointer) = OS.GetSystemTimeAsCurrency()
        
        'If extended font caching is active, retrieve those values now
        If m_ExtendedPropertiesActive Then Uniscribe.GetScriptsSupportedByFont fontName, m_ExtendedProperties(m_FontPointer)
        
        'Return this index
        AddFontToCache = m_FontPointer
        
        'Increment the font pointer, and cycle back to zero as necessary
        If (Not m_CacheIsFull) Then
            m_FontPointer = m_FontPointer + 1
            If (m_FontPointer >= UBound(m_FontCollection)) Then
                m_FontPointer = 0
                m_CacheIsFull = True
            End If
        End If
        
    End If

End Function

'See if a given set of font properties exists in the current cache.  By default, only font face and size are currently matched.
' (In the future, it might be nice to expose an option for "comprehensive searching", which attempts to match all of a font's
' style attributes, too - bold/italic/underline, etc...?)
'
'Returns a value >= 0 if the font exists; the exact value is the index of the font in the collection.
'Returns -1 if the font does not exist.
Private Function DoesFontExist(ByRef fontName As String, Optional ByVal srcFontSize As Single = 0!, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False) As Long
    
    '-1 means the requested font does not exist.  If the requested font *does* exist, this will be set to a non-zero value.
    DoesFontExist = -1
    
    Dim i As Long
    For i = 0 To m_CacheSize - 1
    
        If (Not m_FontCollection(i) Is Nothing) Then
            
            'How we match fonts depends on the current cache mode
            Select Case m_CurrentCacheMode
                
                'Size and style is the most common check in PD, because UI fonts all share the same name
                ' (so we can skip name checks completely, improving performance).
                Case FCM_SizeAndStyle
                    If (srcFontSize = m_FontCollection(i).GetFontSize) Then
                        If (isBold = m_FontCollection(i).GetFontBold) Then
                            If (isItalic = m_FontCollection(i).GetFontItalic) Then
                                If (isUnderline = m_FontCollection(i).GetFontUnderline) Then
                                    DoesFontExist = i
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                    
                Case FCM_NameOnly
                    If Strings.StringsEqual(fontName, m_FontCollection(i).GetFontFace, False) Then
                        DoesFontExist = i
                        Exit For
                    End If
                
                Case FCM_SizeOnly
                    If (srcFontSize = m_FontCollection(i).GetFontSize) Then
                        DoesFontExist = i
                        Exit For
                    End If
                
                Case FCM_NameAndSize
                    If Strings.StringsEqual(fontName, m_FontCollection(i).GetFontFace, False) Then
                        If (srcFontSize = m_FontCollection(i).GetFontSize) Then
                            DoesFontExist = i
                            Exit For
                        End If
                    End If
                    
                Case FCM_NameAndSizeAndStyle
                    If Strings.StringsEqual(fontName, m_FontCollection(i).GetFontFace, False) Then
                        If (srcFontSize = m_FontCollection(i).GetFontSize) Then
                            If (isBold = m_FontCollection(i).GetFontBold) Then
                                If (isItalic = m_FontCollection(i).GetFontItalic) Then
                                    If (isUnderline = m_FontCollection(i).GetFontUnderline) Then
                                        DoesFontExist = i
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
            
            End Select
            
        End If
        
    Next i
    
End Function

'Given an index into the collection (returned by AddFontToCache, presumably), return the corresponding GDI font handle of that
' font object.
'
'IMPORTANT NOTE!  This function does *not* bound check the passed fontIndex, for performance reasons.  It is impossible to pass
' an invalid value if you use the results of AddFontToCache, above, so please do not modify that value after it's been returned.
Friend Function GetFontHandleByPosition(ByVal fontIndex As Long) As Long
    GetFontHandleByPosition = m_FontCollection(fontIndex).GetFontHandle
End Function

'Given an index into the collection (returned by AddFontToCache, presumably), return the corresponding pdFont reference of that
' font object.
'
'IMPORTANT NOTE!  This function does *not* bound check the passed fontIndex, for performance reasons.  It is impossible to pass
' an invalid value if you use the results of AddFontToCache, above, so please do not modify that value after it's been returned.
Friend Function GetFontObjectByPosition(ByVal fontIndex As Long) As pdFont
    Set GetFontObjectByPosition = m_FontCollection(fontIndex)
End Function

'Given an index into the collection (returned by AddFontToCache, presumably), return the corresponding extended font properties
' of that font object.
'
'IMPORTANT NOTE!  This function does *not* bound check the passed fontIndex, for performance reasons.  It is impossible to pass
' an invalid value if you use the results of AddFontToCache, above, so please do not modify that value after it's been returned.
Friend Function GetFontPropertiesByPosition(ByVal fontIndex As Long, ByRef dstProperties As PD_FONT_PROPERTY) As Boolean
    If m_ExtendedPropertiesActive Then
        dstProperties = m_ExtendedProperties(fontIndex)
        GetFontPropertiesByPosition = True
    Else
        GetFontPropertiesByPosition = False
    End If
End Function

'Manually reset the cache
Friend Sub ResetCache()
    
    Dim i As Long
    For i = LBound(m_FontCollection) To UBound(m_FontCollection)
        If (Not m_FontCollection(i) Is Nothing) Then m_FontCollection(i).DeleteCurrentFont
    Next i
    
    m_FontPointer = 0
    ReDim m_FontCollection(0 To m_CacheSize - 1) As pdFont
    ReDim m_ExtendedProperties(0 To m_CacheSize - 1) As PD_FONT_PROPERTY
    ReDim m_FontAccessTimes(0 To m_CacheSize - 1) As Currency
    
    m_CacheIsFull = False
    
End Sub

Friend Sub SetCacheMode(Optional ByVal newMode As FONT_CACHE_MODE = FCM_NameOnly)
    m_CurrentCacheMode = newMode
End Sub

'This function is completely optional, but if you have a known cache requirement, feel free to use of it.
' (Also, note that a large size doesn't hurt you unless you actually fill the cache completely.  Fonts are only created as-needed.)
Friend Sub SetCacheSize(Optional ByVal newSize As Long = DEFAULT_CACHE_SIZE)
    
    m_CacheSize = newSize
    ReDim Preserve m_FontCollection(0 To newSize - 1) As pdFont
    ReDim Preserve m_ExtendedProperties(0 To newSize - 1) As PD_FONT_PROPERTY
    ReDim Preserve m_FontAccessTimes(0 To newSize - 1) As Currency
    
End Sub

'Activate (or deactivate) extended font property caching.  Returns TRUE if successful; note that the function will fail on XP.
Friend Function SetExtendedPropertyCaching(ByVal newSetting As Boolean) As Boolean
    
    'Extended properties are only available on Vista+
    m_ExtendedPropertiesActive = newSetting
    If OS.IsVistaOrLater And m_ExtendedPropertiesActive Then
        
        'If any fonts are already loaded, cache them now
        Dim i As Long
        For i = 0 To UBound(m_FontCollection)
            If Not m_FontCollection(i) Is Nothing Then
                Uniscribe.GetScriptsSupportedByFont m_FontCollection(i).GetFontFace, m_ExtendedProperties(i)
            End If
        Next i
        
    Else
        
        If m_ExtendedPropertiesActive Then
            Debug.Print "WARNING!  Extended font properties are only available on Vista or later."
            m_ExtendedPropertiesActive = False
        End If
        
    End If
    
End Function

Private Sub Class_Initialize()
    
    'Initialize a default cache
    Me.SetCacheSize
    
    'By default, extended properties are not available
    m_ExtendedPropertiesActive = False
    
    'By default, fonts are matched only by name (and not by size or style)
    m_CurrentCacheMode = FCM_NameOnly
    
    'By default, the cache is not full (obviously!)
    m_CacheIsFull = False
    
End Sub

Private Sub Class_Terminate()
    
    If PDMain.IsProgramRunning() Then
    
        Dim i As Long
        For i = LBound(m_FontCollection) To UBound(m_FontCollection)
            If (Not m_FontCollection(i) Is Nothing) Then
                m_FontCollection(i).DeleteCurrentFont
                Set m_FontCollection(i) = Nothing
            End If
        Next i
        
    End If
    
End Sub
