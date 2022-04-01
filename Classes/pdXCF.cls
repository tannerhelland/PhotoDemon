VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdXCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon XCF (GIMP native file format) Container and Parser
'Copyright 2022-2022 by Tanner Helland
'Created: 30/March/22
'Last updated: 30/March/22
'Last update: initial build
'
'This class (and its associated pdXCF- child classes) handle GIMP XCF parsing duties.  It is custom-built
' for PhotoDemon, with an emphasis on performance and proper color-management of all imported data.
'
'As of 2022 this class is under heavy construction, so please be patient with feature support.
'
'As with all 3rd-party XCF engines, GIMP has many features that don't have direct analogs in PhotoDemon.
' Many of these features are still extracted by this class, but they will not "appear" in the final image.
' My ongoing goal is to expand support in this class as various GIMP features are implemented in PD itself.
'
'Unless otherwise noted, all code in this class is my original work.  I've based my work off the
' "official" XCF spec at this URL (link good as of March 2022):
' https://gitlab.gnome.org/GNOME/gimp/-/blob/master/devel-docs/specifications/xcf.txt
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output;
' this will dump a variety of diagnostic information to the debug log.
Private Const XCF_DEBUG_VERBOSE As Boolean = True

'XCF-specific enums are split between this class and the child pdXCFLayer class
Private Enum xcf_ColorMode
    xcf_RGB = 0
    xcf_Grayscale = 1
    xcf_Indexed = 2
End Enum

#If False Then
    Private Const xcf_RGB = 0, xcf_Grayscale = 1, xcf_Indexed = 2
#End If

Private Enum xcf_Precision
    xcf_08bitIntLinear = 0
    xcf_08bitIntGamma = 1
    xcf_16bitIntLinear = 2
    xcf_16bitIntGamma = 3
    xcf_32bitIntLinear = 4
    xcf_32bitIntGamma = 5
    xcf_16bitFltLinear = 6
    xcf_16bitFltGamma = 7
    xcf_32bitFltLinear = 8
    xcf_32bitFltGamma = 9
    xcf_64bitFltLinear = 10
    xcf_64bitFltGamma = 11
End Enum

#If False Then
    Private Const xcf_08bitIntLinear = 0, xcf_08bitIntGamma = 0, xcf_16bitIntLinear = 0, xcf_16bitIntGamma = 0, xcf_32bitIntLinear = 0, xcf_32bitIntGamma = 0, xcf_16bitFltLinear = 0, xcf_16bitFltGamma = 0, xcf_32bitFltLinear = 0, xcf_32bitFltGamma = 0, xcf_64bitFltLinear = 0, xcf_64bitFltGamma = 0
#End If

Private Enum xcf_PropertyID
    xcf_PROP_END = 0
    xcf_PROP_COLORMAP = 1
    xcf_PROP_ACTIVE_LAYER = 2
    xcf_PROP_ACTIVE_CHANNEL = 3
    xcf_PROP_SELECTION = 4
    xcf_PROP_FLOATING_SELECTION = 5
    xcf_PROP_OPACITY = 6
    xcf_PROP_BLEND_MODE = 7
    xcf_PROP_VISIBLE = 8
    xcf_PROP_LINKED = 9
    xcf_PROP_LOCK_ALPHA = 10
    xcf_PROP_APPLY_MASK = 11
    xcf_PROP_EDIT_MASK = 12
    xcf_PROP_SHOW_MASK = 13
    xcf_PROP_SHOW_MASKED = 14
    xcf_PROP_OFFSETS = 15
    xcf_PROP_COLOR = 16
    xcf_PROP_COMPRESSION = 17
    xcf_PROP_GUIDES = 18
    xcf_PROP_RESOLUTION = 19
    xcf_PROP_TATTOO = 20
    xcf_PROP_PARASITES = 21
    xcf_PROP_UNIT = 22
    xcf_PROP_PATHS = 23
    xcf_PROP_USER_UNIT = 24
    xcf_PROP_VECTORS = 25
    xcf_PROP_TEXT_LAYER_FLAGS = 26
    xcf_PROP_LOCK_CONTENT = 28
    xcf_PROP_GROUP_ITEM = 29
    xcf_PROP_ITEM_PATH = 30
    xcf_PROP_GROUP_ITEM_FLAGS = 31
    xcf_PROP_LOCK_POSITION = 32
    xcf_PROP_FLOAT_OPACITY = 33
    xcf_PROP_COLOR_TAG = 34
    xcf_PROP_COMPOSITE_MODE = 35
    xcf_PROP_COMPOSITE_SPACE = 36
    xcf_PROP_BLEND_SPACE = 37
    xcf_PROP_FLOAT_COLOR = 38
    xcf_PROP_SAMPLE_POINTS = 39
    xcf_PROP_ITEM_SET = 40
    xcf_PROP_ITEM_SET_ITEM = 41
    xcf_PROP_LOCK_VISIBILITY = 42
End Enum

#If False Then
    Private Const xcf_PROP_END = 0, xcf_PROP_COLORMAP = 1, xcf_PROP_ACTIVE_LAYER = 2, xcf_PROP_ACTIVE_CHANNEL = 3, xcf_PROP_SELECTION = 4, xcf_PROP_FLOATING_SELECTION = 5, xcf_PROP_OPACITY = 6, xcf_PROP_BLEND_MODE = 7, xcf_PROP_VISIBLE = 8, xcf_PROP_LINKED = 9
    Private Const xcf_PROP_LOCK_ALPHA = 10, xcf_PROP_APPLY_MASK = 11, xcf_PROP_EDIT_MASK = 12, xcf_PROP_SHOW_MASK = 13, xcf_PROP_SHOW_MASKED = 14, xcf_PROP_OFFSETS = 15, xcf_PROP_COLOR = 16, xcf_PROP_COMPRESSION = 17, xcf_PROP_GUIDES = 18, xcf_PROP_RESOLUTION = 19
    Private Const xcf_PROP_TATTOO = 20, xcf_PROP_PARASITES = 21, xcf_PROP_UNIT = 22, xcf_PROP_PATHS = 23, xcf_PROP_USER_UNIT = 24, xcf_PROP_VECTORS = 25, xcf_PROP_TEXT_LAYER_FLAGS = 26, xcf_PROP_LOCK_CONTENT = 28, xcf_PROP_GROUP_ITEM = 29
    Private Const xcf_PROP_ITEM_PATH = 30, xcf_PROP_GROUP_ITEM_FLAGS = 31, xcf_PROP_LOCK_POSITION = 32, xcf_PROP_FLOAT_OPACITY = 33, xcf_PROP_COLOR_TAG = 34, xcf_PROP_COMPOSITE_MODE = 35, xcf_PROP_COMPOSITE_SPACE = 36, xcf_PROP_BLEND_SPACE = 37, xcf_PROP_FLOAT_COLOR = 38, xcf_PROP_SAMPLE_POINTS = 39
    Private Const xcf_PROP_ITEM_SET = 40, xcf_PROP_ITEM_SET_ITEM = 41, xcf_PROP_LOCK_VISIBILITY = 42
#End If

Private Type xcf_Property
    propID As xcf_PropertyID
    propSize As Long
    propData() As Byte
End Type

'Image-wide properties include things like a global palette, tile compression IDs, etc
Private m_numImageProperties As Long, m_ImageProperties() As xcf_Property

'Canvas width/height, color model, and precision all come from the file header
Private m_ImageWidth As Long, m_ImageHeight As Long
Private m_imageColorMode As xcf_ColorMode, m_imagePrecision As xcf_Precision

'XCF version is critical to correct parsing; different versions have different fields in different places!
Private m_xcfVersion As Long

'Byte-by-byte access is provided, as always, by a pdStream instance
Private m_Stream As pdStream

'Validate a source filename as XCF format.  Validation *does* touch the file -
' we must validate a "magic number" in the header.
Friend Function IsFileXCF(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True, Optional ByVal onSuccessLeaveStreamOpen As Boolean = False) As Boolean
    
    Dim potentiallyXCF As Boolean
    potentiallyXCF = Files.FileExists(srcFilename)
    If potentiallyXCF Then potentiallyXCF = (Files.FileLenW(srcFilename) > 26)
    
    'Check extension up front, if requested.
    If (potentiallyXCF And requireValidFileExtension) Then
        potentiallyXCF = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "xcf", True)
    End If
    
    'Proceed with deeper validation as necessary
    If potentiallyXCF Then
        
        'Attempt to load the file
        Set m_Stream = New pdStream
        If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename) Then
            
            'The first 9 bytes of an XCF file must be the ASCII values "gimp xcf "
            potentiallyXCF = (m_Stream.ReadString_ASCII(9) = "gimp xcf ")
            If (potentiallyXCF And XCF_DEBUG_VERBOSE) Then PDDebug.LogAction "Valid XCF file found"
            
        End If
        
    End If
    
    IsFileXCF = potentiallyXCF
    If (Not IsFileXCF) Or (Not onSuccessLeaveStreamOpen) Then Set m_Stream = Nothing
    
End Function

'Validate and load a candidate QOI file
Friend Function LoadXCF_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "LoadXCF_FromFile"
    LoadXCF_FromFile = False
    
    'Validate the file
    If Me.IsFileXCF(srcFile, False, True) Then
        
        'Validation only checks the first 9-bytes of the file for a magic ASCII string.
        ' If we're still here, that validation string passed. (We can still reject the file
        ' if the header contains invalid members.)
        LoadXCF_FromFile = Import_Stage1_ParseHeader(srcFile, dstImage, dstDIB)
        If (Not LoadXCF_FromFile) Then Exit Function
        
        'Still here?  Header looks okay.  Time to proceed with image properties!
        LoadXCF_FromFile = Import_Stage2_LoadProps(srcFile, dstImage, dstDIB)
        
    'No penalty on failed validation; exit immediately
    Else
        Exit Function
    End If

End Function

'Import stages follow.
' IMPORTANT: all Import-prefixed stages must be called in succession, and they must only be called
' from a parent LoadXCF_ function(s).  (These functions rely on correct stream alignment from previous steps,
' and they *will break* if called any other way.

'Import step 2: load all image "properties".  After this step, the stream pointer will be aligned with the
' start of the layer segment.
Private Function Import_Stage2_LoadProps(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "Import_Stage2_LoadProps"
    Import_Stage2_LoadProps = False
    
    'Properties are basically a list of chunks with numeric IDs.  Each chunk self-describes
    ' its own length (although some historic XCF files may not report size correctly - I'm not sure
    ' how to address this and the spec is unclear, so this is TODO pending further testing)/
    
    'In this step, we basically just want to iterate all properties and copy them into a local collection.
    Dim curPropID As Long, curPropSize As Long
    curPropID = m_Stream.ReadLong_BE()
    
    'Start with an arbitrary collection size
    Const INIT_PROP_LIST_COUNT As Long = 16
    ReDim m_ImageProperties(0 To INIT_PROP_LIST_COUNT - 1) As xcf_Property
    
    Do While (curPropID <> 0)
        
        'Ensure we have room for this property, then store the new ID
        If (m_numImageProperties > UBound(m_ImageProperties)) Then ReDim Preserve m_ImageProperties(0 To m_numImageProperties * 2 - 1) As xcf_Property
        m_ImageProperties(m_numImageProperties).propID = curPropID
        
        'After ID comes property length
        curPropSize = m_Stream.ReadLong_BE()
        m_ImageProperties(m_numImageProperties).propSize = curPropSize
        
        'Special handling for legacy properties with potentially invalid payload sizes could go here
        
        'After ID comes payload
        If (curPropSize > 0) Then m_Stream.ReadBytes m_ImageProperties(m_numImageProperties).propData, curPropSize, True
        
        'Debug info (please disable in production)
        If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "Found image property: " & GetPropertyName(curPropID) & " (" & curPropSize & " bytes)"
        
        'Increment total image property count
        m_numImageProperties = m_numImageProperties + 1
        
        'Read next property ID
        curPropID = m_Stream.ReadLong_BE()
        
    Loop
    
End Function

'Import step 1: after "magic number" validation, parse the short XCF header and validate core members
' like image width/height
Private Function Import_Stage1_ParseHeader(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "Import_Stage1_ParseHeader"
    Import_Stage1_ParseHeader = False

    'Magic number is followed by version (as a string).  We remap this to a long for convenience.
    Dim strVersion As String
    strVersion = LCase$(m_Stream.ReadString_ASCII(4))
    
    If (strVersion = "file") Then
        m_xcfVersion = 0
    Else
        m_xcfVersion = CLng(Right$(strVersion, 3))
    End If
    
    If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "XCF version is " & m_xcfVersion
    
    '1-byte null-padding follows
    m_Stream.ReadByte
    
    'Canvas width/height
    m_ImageWidth = m_Stream.ReadLong_BE
    m_ImageHeight = m_Stream.ReadLong_BE
    
    'Color model
    m_imageColorMode = m_Stream.ReadLong_BE
    
    'If file is v4 or later, a 4-byte precision value follows next.  (For earlier versions,
    ' this field is omitted and a precision of "8-bit gamma integer" is assumed.)
    If (m_xcfVersion >= 4) Then
        
        Dim tmpInt As Long
        tmpInt = m_Stream.ReadLong_BE
        
        'This field is mapped weirdly; see the spec for details
        Select Case tmpInt
            
            Case 100
                m_imagePrecision = xcf_08bitIntLinear
            Case 150
                m_imagePrecision = xcf_08bitIntGamma
            Case 200
                m_imagePrecision = xcf_16bitIntLinear
            Case 250
                m_imagePrecision = xcf_16bitIntGamma
            Case 300
                m_imagePrecision = xcf_32bitIntLinear
            Case 350
                m_imagePrecision = xcf_32bitIntGamma
            Case 500
                m_imagePrecision = xcf_16bitFltLinear
            Case 550
                m_imagePrecision = xcf_16bitFltGamma
            Case 600
                m_imagePrecision = xcf_32bitFltLinear
            Case 650
                m_imagePrecision = xcf_32bitFltGamma
            Case 700
                m_imagePrecision = xcf_64bitFltLinear
            Case 750
                m_imagePrecision = xcf_64bitFltGamma
            Case Else
                InternalError FUNC_NAME, "unknown image precision: " & tmpInt
                Exit Function
        End Select
        
    Else
        m_imagePrecision = xcf_08bitIntGamma
    End If
    
    'And that's it for the header!  Do a few basic validations before continuing.
    If (m_ImageWidth <= 0) Or (m_ImageHeight <= 0) Then
        InternalError FUNC_NAME, "bad image width/height"
        Exit Function
        
    ElseIf (m_ImageWidth >= 65535) Or (m_ImageHeight >= 65535) Then
        InternalError FUNC_NAME, "image width/height too big"
        Exit Function
        
    ElseIf (m_imageColorMode < 0) Or (m_imageColorMode > 2) Then
        InternalError FUNC_NAME, "bad color mode: " & m_imageColorMode
        Exit Function
    End If
    
    'Still here?  Header looks okay!
    If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "XCF is " & m_ImageWidth & "x" & m_ImageHeight & " with color mode " & m_imageColorMode & " and precision " & m_imagePrecision
    Import_Stage1_ParseHeader = True
    
End Function

Private Function GetPropertyName(ByVal propID As xcf_PropertyID) As String
    Select Case propID
        Case xcf_PROP_END
            GetPropertyName = "end"
        Case xcf_PROP_COLORMAP
            GetPropertyName = "color map"
        Case xcf_PROP_ACTIVE_LAYER
            GetPropertyName = "active layer"
        Case xcf_PROP_ACTIVE_CHANNEL
            GetPropertyName = "active channel"
        Case xcf_PROP_SELECTION
            GetPropertyName = "selection"
        Case xcf_PROP_FLOATING_SELECTION
            GetPropertyName = "floating selection"
        Case xcf_PROP_OPACITY
            GetPropertyName = "opacity"
        Case xcf_PROP_BLEND_MODE
            GetPropertyName = "blend mode"
        Case xcf_PROP_VISIBLE
            GetPropertyName = "visibility"
        Case xcf_PROP_LINKED
            GetPropertyName = "linked"
        Case xcf_PROP_LOCK_ALPHA
            GetPropertyName = "lock alpha"
        Case xcf_PROP_APPLY_MASK
            GetPropertyName = "apply mask"
        Case xcf_PROP_EDIT_MASK
            GetPropertyName = "edit mask"
        Case xcf_PROP_SHOW_MASK
            GetPropertyName = "show mask"
        Case xcf_PROP_SHOW_MASKED
            GetPropertyName = "show masked"
        Case xcf_PROP_OFFSETS
            GetPropertyName = "offsets"
        Case xcf_PROP_COLOR
            GetPropertyName = "color"
        Case xcf_PROP_COMPRESSION
            GetPropertyName = "compression"
        Case xcf_PROP_GUIDES
            GetPropertyName = "guides"
        Case xcf_PROP_RESOLUTION
            GetPropertyName = "resolution"
        Case xcf_PROP_TATTOO
            GetPropertyName = "tattoo"
        Case xcf_PROP_PARASITES
            GetPropertyName = "parasite"
        Case xcf_PROP_UNIT
            GetPropertyName = "unit"
        Case xcf_PROP_PATHS
            GetPropertyName = "paths"
        Case xcf_PROP_USER_UNIT
            GetPropertyName = "user unit"
        Case xcf_PROP_VECTORS
            GetPropertyName = "vectors"
        Case xcf_PROP_TEXT_LAYER_FLAGS
            GetPropertyName = "text layer flags"
        Case xcf_PROP_LOCK_CONTENT
            GetPropertyName = "lock content"
        Case xcf_PROP_GROUP_ITEM
            GetPropertyName = "group item"
        Case xcf_PROP_ITEM_PATH
            GetPropertyName = "item path"
        Case xcf_PROP_GROUP_ITEM_FLAGS
            GetPropertyName = "item flags"
        Case xcf_PROP_LOCK_POSITION
            GetPropertyName = "lock position"
        Case xcf_PROP_FLOAT_OPACITY
            GetPropertyName = "float opacity"
        Case xcf_PROP_COLOR_TAG
            GetPropertyName = "color tag"
        Case xcf_PROP_COMPOSITE_MODE
            GetPropertyName = "composite mode"
        Case xcf_PROP_COMPOSITE_SPACE
            GetPropertyName = "composite space"
        Case xcf_PROP_BLEND_SPACE
            GetPropertyName = "blend space"
        Case xcf_PROP_FLOAT_COLOR
            GetPropertyName = "float color"
        Case xcf_PROP_SAMPLE_POINTS
            GetPropertyName = "sample points"
        Case xcf_PROP_ITEM_SET
            GetPropertyName = "item set"
        Case xcf_PROP_ITEM_SET_ITEM
            GetPropertyName = "item set item"
        Case xcf_PROP_LOCK_VISIBILITY
            GetPropertyName = "lock visibility"
    End Select
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdXCF." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdXCF." & funcName & "() reported an error: " & errDescription
    End If
End Sub
