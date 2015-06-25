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
'Last updated: 12/May/15
'Last update: initial build
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

'Size of the font collection.  This is the maximum number of fonts this class is allowed to cache.
Private Const DEFAULT_CACHE_SIZE As Long = 24
Private m_CacheSize As Long

'Pointer to the current font index.  This will rotate around to 0 after the cache is filled.
Private m_FontPointer As Long

'Actual collection
Private fontCollection() As pdFont

'This class can (optionally) manage extended font properties as well, like which scripts are supported by a given font.
' To activate this behavior, pass TRUE to the setExtendedPropertyCaching sub.
Private m_ExtendedPropertiesActive As Boolean
Private m_ExtendedProperties() As PD_FONT_PROPERTY

'Add a font to the cache.
' Returns: value >= 0, indicating the index of said font in the cache.  This can be used to retrieve the font handle, which is faster
' than accessing fonts by name.
Public Function addFontToCache(ByRef fontName As String, ByVal fontSize As Single, Optional ByVal fontBold As Boolean = False, Optional ByVal fontItalic As Boolean = False, Optional ByVal fontUnderline As Boolean = False) As Long
    
    'First, see if this font already exists in the cache
    Dim fontIndex As Long
    fontIndex = doesFontExist(fontName)
    
    'This font already exists!  Don't recreate it; instead, return its current index.
    ' (IMPORTANT NOTE!  At present, the collection only attempts to match font face.  Size and other characteristics are NOT matched.
    '  This is by design to improve performance, and you will have to manually change this to cache the same font with different
    '  size or bold/italic/underline properties.)
    If fontIndex >= 0 Then
    
        addFontToCache = fontIndex
        
    Else
    
        'Create a new font object as necessary
        If fontCollection(m_FontPointer) Is Nothing Then
            Set fontCollection(m_FontPointer) = New pdFont
        Else
            fontCollection(m_FontPointer).deleteCurrentFont
        End If
        
        'Initialize all the required font properties
        With fontCollection(m_FontPointer)
            .setFontFace fontName
            .setFontSize fontSize
            .setFontBold fontBold
            .setFontItalic fontItalic
            .setFontUnderline fontUnderline
        End With
    
        'Create the font object
        With fontCollection(m_FontPointer)
            .createFontObject
        End With
        
        'If extended font caching is active, retrieve those values now
        If m_ExtendedPropertiesActive Then Uniscribe_Interface.getScriptsSupportedByFont fontName, m_ExtendedProperties(m_FontPointer)
        
        'Return this index
        addFontToCache = m_FontPointer
        
        'Increment the font pointer, and cycle back to zero as necessary
        m_FontPointer = m_FontPointer + 1
        If m_FontPointer >= UBound(fontCollection) Then m_FontPointer = 0
        
    End If

End Function

'See if a given set of font properties exists in the current cache.  By default, only font face is currently matched.
' (In the future, it might be nice to expose an option for "comprehensive searching", which attempts to match all of a font's
' attributes - size, bold/italic/underline, etc...?)
'
'Returns a value >= 0 if the font exists; the exact value is the index of the font in the collection.
'Returns -1 if the font does not exist.
Private Function doesFontExist(ByRef fontName As String) As Long
    
    Dim i As Long
    For i = 0 To m_CacheSize - 1
    
        If Not fontCollection(i) Is Nothing Then
            If StrComp(fontName, fontCollection(i).getFontFace, vbBinaryCompare) = 0 Then
                doesFontExist = i
                Exit Function
            End If
        End If
        
    Next i
    
    'If we made it here, the font does not exist
    doesFontExist = -1
    
End Function

'Given an index into the collection (returned by addFontToCache, presumably), return the corresponding GDI font handle of that
' font object.
'
'IMPORTANT NOTE!  This function does *not* bound check the passed fontIndex, for performance reasons.  It is impossible to pass
' an invalid value if you use the results of addFontToCache, above, so please do not modify that value after it's been returned.
Public Function getFontHandleByPosition(ByVal fontIndex As Long) As Long
    getFontHandleByPosition = fontCollection(fontIndex).getFontHandle
End Function

'Given an index into the collection (returned by addFontToCache, presumably), return the corresponding extended font properties
' of that font object.
'
'IMPORTANT NOTE!  This function does *not* bound check the passed fontIndex, for performance reasons.  It is impossible to pass
' an invalid value if you use the results of addFontToCache, above, so please do not modify that value after it's been returned.
Friend Function getFontPropertiesByPosition(ByVal fontIndex As Long, ByRef dstProperties As PD_FONT_PROPERTY) As Boolean
    
    If m_ExtendedPropertiesActive Then
        
        dstProperties = m_ExtendedProperties(fontIndex)
        getFontPropertiesByPosition = True
        
    Else
        getFontPropertiesByPosition = False
    End If
    
End Function

'Manually reset the cache
Public Sub resetCache()

    m_FontPointer = 0
    ReDim fontCollection(0 To m_CacheSize - 1) As pdFont
    ReDim m_ExtendedProperties(0 To m_CacheSize - 1) As PD_FONT_PROPERTY
    
End Sub

'This function is completely optional, but if you have a known cache requirement, feel free to make use of it
Public Sub setCacheSize(Optional ByVal newSize As Long = DEFAULT_CACHE_SIZE)
    
    m_CacheSize = newSize
    ReDim Preserve fontCollection(0 To newSize - 1) As pdFont
    ReDim Preserve m_ExtendedProperties(0 To newSize - 1) As PD_FONT_PROPERTY
    
End Sub

'Activate (or deactivate) extended font property caching.  Returns TRUE if successful; note that the function will fail on XP.
Public Function setExtendedPropertyCaching(ByVal newSetting As Boolean) As Boolean
    
    'Extended properties are only available on Vista+
    Dim cSysInfo As pdSystemInfo
    Set cSysInfo = New pdSystemInfo
    
    m_ExtendedPropertiesActive = newSetting
    
    If cSysInfo.isOSVistaOrLater And m_ExtendedPropertiesActive Then
        
        'If any fonts are already loaded, cache them now
        Dim i As Long
        For i = 0 To UBound(fontCollection)
            If Not fontCollection(i) Is Nothing Then
                Uniscribe_Interface.getScriptsSupportedByFont fontCollection(i).getFontFace, m_ExtendedProperties(i)
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
    setCacheSize
    
    'By default, extended properties are not available
    m_ExtendedPropertiesActive = False
    
End Sub
