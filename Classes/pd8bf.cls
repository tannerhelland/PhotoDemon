VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd8bf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'8bf Plugin Instance Handler
'Copyright 2025-2025 by Tanner Helland
'Created: 17/December/25
'Last updated: 17/December/25
'Last update: initial build (as part of a larger project to replace pspihost with our own 8bf implementation)
'
'8bf files are 3rd-party Adobe Photoshop plugins that implement one or more "filters".  These are
' basically DLL files with special interfaces for communicating with a parent Photoshop instance.
'
'We attempt to support these plugins in PhotoDemon, with PD standing in for Photoshop as the
' "host" of the plugins.
'
'This class manages a single 8bf instance.  When the Effects > 8bf dialog is initiated, PD assembles information
' on all installed 8bfs (this is necessary to create a list with plugin category/name) and an array of these
' classes is instantiated to manage plugin data.  Only one plugin at a time is ever *running*, however.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Toggle verbose debug logging.  Set to FALSE in production builds.
Private Const DEBUG_VERBOSE As Boolean = True

'Local copy of PiPL properties
Private m_numProperties As Long, m_Properties() As PIProperty

'Some properties are frequently accessed; we cache them at initialization time
Private m_Name As String, m_Category As String, m_VersionMajor As Long, m_VersionMinor As Long
Private m_EntryPoint As String

'Given a list of raw PiPL properties (and the target file from which they came), initialize this class.
' Returns TRUE if the property list is valid and compatible with PD; FALSE otherwise.  The plugin should
' not be exposed to the user if FALSE is returned.
Friend Function Initialize8bf_FromFile(ByRef srcFile As String, ByVal numProperties As Long, ByRef arrayOfProperties() As PIProperty) As Boolean
    
    Initialize8bf_FromFile = False
    
    'Ensure at least 2 properties were retrieved (technically we need more than this,
    ' but at a *minimum* we need a name and category to display)
    If (numProperties < 2) Then
        If DEBUG_VERBOSE Then PDDebug.LogAction "WARNING: insufficient properties in " & srcFile
        Exit Function
    End If
    
    'Make a local copy of all properties
    m_numProperties = numProperties
    ReDim m_Properties(0 To numProperties - 1) As PIProperty
    
    Dim i As Long
    For i = 0 To m_numProperties - 1
        m_Properties(i) = arrayOfProperties(i)
    Next i
    
    'Further processing is made much simpler by a stream object
    Dim lenPascalString As Byte, tmpLong As Long
    Dim kindOK As Boolean, versionOK As Boolean
    
    Dim cStream As pdStream
    Set cStream = New pdStream
    
    'Next, we need to classify and validate key properties.
    For i = 0 To m_numProperties - 1
        
        Select Case m_Properties(i).propertyKey
        
            'Type or kind of plug-in; key value is 'kind'.
            Case "kind"
                If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_Properties(i).propertyLength, VarPtr(m_Properties(i).pPropertyData(0))) Then
                    
                    tmpLong = cStream.ReadLong_BE()
                    Dim filterKind As String
                    filterKind = Strings.StringFromCharPtr(VarPtr(tmpLong), False, 4, True)
                    
                    'Filter plugins must explicitly be marked as "8BFM"
                    Const PS_FILTER_KIND_ID As String = "8BFM"
                    kindOK = Strings.StringsEqual(PS_FILTER_KIND_ID, filterKind, False)
                    If (Not kindOK) Then
                        If DEBUG_VERBOSE Then PDDebug.LogAction "WARNING: " & srcFile & " is wrong plugin kind (" & filterKind & ")"
                        Exit For
                    End If
                    
                End If
                
            'Code-specific properties.  We only need x86 here; others are included for curiosity only
            Case "m68k"
                'deprecated: 'm68k' 68k code descriptor.
            
            Case "pwpc", "ppcb"
                'deprecated: CodePowerPC property in the PiPL resource file.
            
            Case "mach"
                ''mach' PowerPC Mach-O code descriptor.
            
            Case "mi32"
                ''mi32' Intel 32 Mach-O code descriptor.
            
            Case "mi64"
                ''mi64' Intel 64 Mach-O code descriptor.
            
            Case "frag"
                ''frag' PowerPC fragment descriptor.
            
            Case "frgc"
                ''frgc' PowerPC CARBON fragment descriptor.
        
            Case "fx86"
                'deprecated: 'fx86' Win16 Intel code descriptor.
            
            ''8664' Win32 Intel code descriptor.
            Case "8664"
                '// Name of entrypoint ("main", "ENTRYPOINT", "", etc.):
                'unsigned char fEntryName[1];
                
            ''wx86' Win32 Intel code descriptor.
            Case "wx86"
                '// Name of entrypoint ("main", "ENTRYPOINT", "", etc.):
                'unsigned char fEntryName[1];
                If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_Properties(i).propertyLength, VarPtr(m_Properties(i).pPropertyData(0))) Then
                    m_EntryPoint = Trim$(Strings.TrimNull(cStream.ReadString_ASCII(m_Properties(i).propertyLength)))
                    Debug.Print "entry point: " & m_EntryPoint
                End If
                
            
            'Next are non-code properties
            
            '// 'vers' \<int32\> Major\<int16\>.Minor\<int16\> version number:
            Case "vers"
                If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_Properties(i).propertyLength, VarPtr(m_Properties(i).pPropertyData(0))) Then
                    m_VersionMinor = cStream.ReadIntUnsigned()
                    m_VersionMajor = cStream.ReadIntUnsigned()
                    Debug.Print "version: " & m_VersionMajor & "." & m_VersionMinor
                End If
                
                'TODO: validate versionOK here
            
            ''prty' \<int32\> Load order priority:
            'Also used to control the order in which items with the same name
            ' show up in menus. Lower numbers (including negative ones) load first.
            ' If NULL, the default is zero.
            Case "prty"
            
            ''cmpt' Component/Version ids:
            Case "cmpt"
            
            ''mode' \<FlagSet\> Image modes supported flags. (bitmask):
            ' The data for the property has type \c FlagSet; a variable length bitstring
            ' in which the first member is represented by the most significant bit of the
            ' first byte, the eighth member is in the least significant bit of the first byte, etc.
            ' The length of the set is in the property's length field. <br><br>
            ' Adobe Photoshop, has 18 modes, see @ref ImageModes. <br><br>
            ' This property determines whether your plug-in will be active (black) or inactive (gray)
            ' in Photoshop’s menus based on the current document’s image mode.
            Case "mode"
                'TODO: validate this against whichever modes I'm willing to support
            
            ''enbl' \<CString\> Enabling expression:
            Case "enbl"
            
            ''wnsc' plug in wants scrap
            Case "wnsc"
            
            ''coco' plug in has Cocoa, Objective-C code and should never be unloaded
            Case "coco"
            
            ''fbaw' plug in wants in the File Browser menu
            Case "fbaw"
            
            ''nabo' plug in does not want an about box menu entry
            Case "nabo"
            
            'lFGm' plug in should be hidden/shown by the Preference for legacy Filter Gallery menu entries
            Case "lFGm"
            
            ''flly' plug in filter layer properties
            Case "flly"
            
            ''host' \<PIType\> giving host required if any:
            Case "host"
            
            ''catg' \<PString\> Category name that appears on top level menu:
            Case "catg"
                If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_Properties(i).propertyLength, VarPtr(m_Properties(i).pPropertyData(0))) Then
                    lenPascalString = cStream.ReadByte()
                    m_Category = cStream.ReadString_ASCII(PDMath.Min2Int(lenPascalString, m_Properties(i).propertyLength))
                    Debug.Print "category: " & m_Category
                End If
                
            ''zcat' \<PString\> Category name that appears on top level menu:
            Case "zcat"
                'zstrings can hypothetically be used for localized strings, but I have not found a plugin that
                ' uses this (ones from Photoshop *do* but they're the exception that proves the rule in this case)
                
            ''name' \<PString\> Menu name:
            Case "name"
                If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_Properties(i).propertyLength, VarPtr(m_Properties(i).pPropertyData(0))) Then
                    lenPascalString = cStream.ReadByte()
                    m_Name = cStream.ReadString_ASCII(PDMath.Min2Int(lenPascalString, m_Properties(i).propertyLength))
                    Debug.Print "name: " & m_Name
                End If
                
            ''znam' \<PString\> Menu name:
            Case "znam"
                'zstrings can hypothetically be used for localized strings, but I have not found a plugin that
                ' uses this (ones from Photoshop *do* but they're the exception that proves the rule in this case)
            
            ''prog' \<PString\> Progress text:
            Case "prog"
            
            ''zpro' \<PString\> Progress text:
            Case "zpro"
            
            ''pnme' \<CString\> color picker ID:
            Case "pnme"
            
            ''piLU' plug-in supports load/unload selectors:
            Case "piLU"
            
            ''piPF' plug-in supports preferences dialog selectors:
            Case "piPF"
            
            ''feen' plug-in is not loaded unless this feature is enabled
            Case "feen"
            
            ''posd' plug-in is position dependent
            Case "posd"
            
        End Select
        
    Next i
    
End Function
