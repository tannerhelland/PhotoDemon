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
'Copyright 2014-2015 by Tanner Helland
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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

'Size of the font collection.  This is the maximum number of fonts this class is allowed to cache.
Private Const DEFAULT_CACHE_SIZE As Long = 24
Private m_CacheSize As Long

'Pointer to the current font index.  This will rotate around to 0 after the cache is filled.
Private m_FontPointer As Long

'Actual collection
Private m_FontCollection() As pdFont

'This class can (optionally) manage extended font properties as well, like which scripts are supported by a given font.
' To activate this behavior, pass TRUE to the setExtendedPropertyCaching sub.
Private m_ExtendedPropertiesActive As Boolean
Private m_ExtendedProperties() As PD_FONT_PROPERTY

'Add a font to the cache.
' Returns: value >= 0, indicating the index of said font in the cache.  (This index can be subsequently used to retrieve the
'           actual handle or object.)
Public Function AddFontToCache(ByRef fontName As String, ByVal FontSize As Single, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False) As Long
    
    'First, see if this font already exists in the cache
    Dim fontIndex As Long
    fontIndex = DoesFontExist(fontName, FontSize, isBold, isItalic, isUnderline)
    
    'If this font already exists in our collection, don't recreate it; instead, return its current index.
    ' IMPORTANT NOTE: the matching criteria used by DoesFontExist defaults to fontname matches only.
    '                  Call SetCacheMode() if you want to match fonts by alternate criteria.
    If fontIndex >= 0 Then
    
        AddFontToCache = fontIndex
        
    Else
    
        'Create a new font object as necessary
        If m_FontCollection(m_FontPointer) Is Nothing Then
            Set m_FontCollection(m_FontPointer) = New pdFont
        Else
            m_FontCollection(m_FontPointer).DeleteCurrentFont
        End If
        
        'Initialize all the required font properties
        With m_FontCollection(m_FontPointer)
            .SetFontFace fontName
            .SetFontSize FontSize
            .SetFontBold isBold
            .SetFontItalic isItalic
            .SetFontUnderline isUnderline
        End With
    
        'Create the font object
        m_FontCollection(m_FontPointer).CreateFontObject
        
        'If extended font caching is active, retrieve those values now
        If m_ExtendedPropertiesActive Then Uniscribe_Interface.getScriptsSupportedByFont fontName, m_ExtendedProperties(m_FontPointer)
        
        'Return this index
        AddFontToCache = m_FontPointer
        
        'Increment the font pointer, and cycle back to zero as necessary
        m_FontPointer = m_FontPointer + 1
        If m_FontPointer >= UBound(m_FontCollection) Then m_FontPointer = 0
        
    End If

End Function

'See if a given set of font properties exists in the current cache.  By default, only font face and size are currently matched.
' (In the future, it might be nice to expose an option for "comprehensive searching", which attempts to match all of a font's
' style attributes, too - bold/italic/underline, etc...?)
'
'Returns a value >= 0 if the font exists; the exact value is the index of the font in the collection.
'Returns -1 if the font does not exist.
Private Function DoesFontExist(ByRef fontName As String, Optional ByVal FontSize As Single = 0#, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False) As Long
    
    '-1 means the requested font does not exist.  If the requested font *does* exist, this will be set to a non-zero value.
    DoesFontExist = -1
    
    Dim i As Long
    For i = 0 To m_CacheSize - 1
    
        If Not m_FontCollection(i) Is Nothing Then
            
            With m_FontCollection(i)
            
                'How we match fonts depends on the current cache mode
                Select Case m_CurrentCacheMode
                
                    Case FCM_NameOnly
                        If StrComp(fontName, .GetFontFace, vbBinaryCompare) = 0 Then
                            DoesFontExist = i
                            Exit For
                        End If
                    
                    Case FCM_SizeOnly
                        If (FontSize = .GetFontSize) Then
                            DoesFontExist = i
                            Exit For
                        End If
                    
                    Case FCM_NameAndSize
                        If StrComp(fontName, .GetFontFace, vbBinaryCompare) = 0 Then
                            If (FontSize = m_FontCollection(i).GetFontSize) Then
                                DoesFontExist = i
                                Exit For
                            End If
                        End If
                    
                    Case FCM_SizeAndStyle
                        If (FontSize = .GetFontSize) Then
                            If (isBold = .GetFontBold) Then
                                If (isItalic = .GetFontItalic) Then
                                    If (isUnderline = .GetFontUnderline) Then
                                        DoesFontExist = i
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                        
                    Case FCM_NameAndSizeAndStyle
                        If StrComp(fontName, .GetFontFace, vbBinaryCompare) = 0 Then
                            If (FontSize = m_FontCollection(i).GetFontSize) Then
                                If (isBold = .GetFontBold) Then
                                    If (isItalic = .GetFontItalic) Then
                                        If (isUnderline = .GetFontUnderline) Then
                                            DoesFontExist = i
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                
                End Select
            
            End With
            
        End If
        
    Next i
        
End Function

'Given an index into the collection (returned by AddFontToCache, presumably), return the corresponding GDI font handle of that
' font object.
'
'IMPORTANT NOTE!  This function does *not* bound check the passed fontIndex, for performance reasons.  It is impossible to pass
' an invalid value if you use the results of AddFontToCache, above, so please do not modify that value after it's been returned.
Public Function GetFontHandleByPosition(ByVal fontIndex As Long) As Long
    GetFontHandleByPosition = m_FontCollection(fontIndex).GetFontHandle
End Function

'Given an index into the collection (returned by AddFontToCache, presumably), return the corresponding pdFont reference of that
' font object.
'
'IMPORTANT NOTE!  This function does *not* bound check the passed fontIndex, for performance reasons.  It is impossible to pass
' an invalid value if you use the results of AddFontToCache, above, so please do not modify that value after it's been returned.
Public Function GetFontObjectByPosition(ByVal fontIndex As Long) As pdFont
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
Public Sub ResetCache()
    
    Dim i As Long
    For i = LBound(m_FontCollection) To UBound(m_FontCollection)
        If Not (m_FontCollection(i) Is Nothing) Then m_FontCollection(i).DeleteCurrentFont
    Next i
    
    m_FontPointer = 0
    ReDim m_FontCollection(0 To m_CacheSize - 1) As pdFont
    ReDim m_ExtendedProperties(0 To m_CacheSize - 1) As PD_FONT_PROPERTY
    
End Sub

Public Sub SetCacheMode(Optional ByVal newMode As FONT_CACHE_MODE = FCM_NameOnly)
    m_CurrentCacheMode = newMode
End Sub

'This function is completely optional, but if you have a known cache requirement, feel free to use of it.
' (Also, note that a large size doesn't hurt you unless you actually fill the cache completely.  Fonts are only created as-needed.)
Public Sub SetCacheSize(Optional ByVal newSize As Long = DEFAULT_CACHE_SIZE)
    
    m_CacheSize = newSize
    ReDim Preserve m_FontCollection(0 To newSize - 1) As pdFont
    ReDim Preserve m_ExtendedProperties(0 To newSize - 1) As PD_FONT_PROPERTY
    
End Sub

'Activate (or deactivate) extended font property caching.  Returns TRUE if successful; note that the function will fail on XP.
Public Function SetExtendedPropertyCaching(ByVal newSetting As Boolean) As Boolean
    
    'Extended properties are only available on Vista+
    Dim cSysInfo As pdSystemInfo
    Set cSysInfo = New pdSystemInfo
    
    m_ExtendedPropertiesActive = newSetting
    
    If cSysInfo.IsOSVistaOrLater And m_ExtendedPropertiesActive Then
        
        'If any fonts are already loaded, cache them now
        Dim i As Long
        For i = 0 To UBound(m_FontCollection)
            If Not m_FontCollection(i) Is Nothing Then
                Uniscribe_Interface.getScriptsSupportedByFont m_FontCollection(i).GetFontFace, m_ExtendedProperties(i)
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
    SetCacheSize
    
    'By default, extended properties are not available
    m_ExtendedPropertiesActive = False
    
    'By default, fonts are matched only by name (and not by size or style)
    m_CurrentCacheMode = FCM_NameOnly
    
End Sub

Private Sub Class_Terminate()
    
    Dim i As Long
    For i = LBound(m_FontCollection) To UBound(m_FontCollection)
        If Not (m_FontCollection(i) Is Nothing) Then
            m_FontCollection(i).DeleteCurrentFont
            Set m_FontCollection(i) = Nothing
        End If
    Next i
    
End Sub
