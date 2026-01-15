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
'Last updated: 15/January/26
'Last update: initial release to public; only plugin enumeration and "About" dialog is implemented here (for now);
'             full 8bf execution still requires pspihost
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
Private Const DEBUG_VERBOSE As Boolean = False

'Local copy of PiPL properties
Private m_numProperties As Long, m_Properties() As PIProperty

'Some properties are frequently accessed; we cache them at initialization time
Private m_Filename As String
Private m_Name As String, m_Category As String, m_VersionMajor As Long, m_VersionMinor As Long
Private m_EntryPoint As String, m_pEntryPoint As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to a maximum relevant size.
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

'Error codes come from PIGeneral.h (Adobe plugin SDK).
' Note that these are returned via *short int* not *long int*
Private Enum PS_8bfErrorCodes

    ps_Success = 0
    
    '//-------------------------------------------------------------------------------
    '//  Defines -- General plug-in error codes.
    '//-------------------------------------------------------------------------------

    '/** Indicates that the plug-in requires services not provided by this host. */
    errPlugInHostInsufficient = -30900
    
    '/** Indicates that a requested property could not be found. */
    errPlugInPropertyUndefined = -30901
    
    '/**
    '* Indicates that an image processing operation does not support colBytes values
    '* other than 1. Message is "The plug-in requires services not provided by this host."
    '*/
    errHostDoesNotSupportColStep = -30902

    '/**
    '* For the Color Services callback, during the operation plugIncolorServicesSamplePoint,
    '* indicates that the is not in the image window.
    '*/
    errInvalidSamplePoint = -30903
             
    '/**
    '* When a plug-in returns with \c result=errReportString, the
    '* host reports the \c errorString value from the plug-in's parameter block.
    '*/
    errReportString = -30904
    
    '//-------------------------------------------------------------------------------
    '//  Defines -- Channel Port error codes.
    '//-------------------------------------------------------------------------------
    
    errUnknownPort = -30910                        '/**< Unknown Channel Port. */
    errUnsupportedRowBits = -30911                 '/**< Host does not support value provided in
                                                   '            @ref PixelMemoryDesc::rowBits. */
    
    errUnsupportedColBits = -30912                 '/**< Host does not support value provided in
                                                   '            @ref PixelMemoryDesc::colBits. */
    
    errUnsupportedBitOffset = -30913               '/**< Host does not support value provided in
                                                   '            @ref PixelMemoryDesc::bitOffset. */
    
    errUnsupportedDepth = -30914                   '/**< Host does not support value provided in
                                                   '            @ref PixelMemoryDesc::depth. */
    
    errUnsupportedDepthConversion = -30915         '/**< Host does not provide conversion for the depth
                                                   '            provided in @ref PixelMemoryDesc::depth. */
    

    '/* Error return values. The plug-in module may also return standard Macintosh
    '   operating system error codes, or report its own errors, in which case it
    '   can return any positive integer. */
    filterBadParameters = -30100  '/**< The filter module had a problem with the filter module interface. */
    filterBadMode = -30101        '/**< The filter module does not support \<mode\> images. */
    
End Enum

'Photoshop types follow
'/** Acquires a function suite. Loads the suite if necessary,
'        and increments its reference count. For example:
'    @code
'SPErr error;
'SPBasicSuite *sBasic = message->d.basic;
'AIRandomSuite *sRandom;
'sBasic->AcquireSuite( kAIRandomSuite, kAIRandomVersion, &sRandom );
'    @endcode
'            @param name The suite name.
'            @param version The suite version number.
'            @param suite [out] A buffer in which to return the suite pointer.
'            @see \c #SPSuitesSuite::AcquireSuite()
'        */
Private Type SPBasicSuite
    
    'SPAPI SPErr (*AcquireSuite)( const char *name, int32 version, const void **suite );
    '/** Decrements the reference count of a suite and unloads it when the
    '    reference count reaches 0.
    '        @param name The suite name.
    '        @param version The suite version number.
    '    */
    pAcquireSuite As Long
    
    'SPAPI SPErr (*ReleaseSuite)( const char *name, int32 version );
    '/** Compares two strings for equality.
    '        @param token1 The first null-terminated string.
    '        @param token2 The second null-terminated string.
    '        @return True if the strings are the same, false otherwise.
    '    */
    pReleaseSuite As Long
    
    'SPAPI SPBoolean (*IsEqual)( const char *token1, const char *token2 );
    '/** Allocates a block of memory.
    '        @param size The number of bytes.
    '        @param block [out] A buffer in which to return the block pointer.
    '        @see \c #SPBlocksSuite::AllocateBlock()
    '    */
    pIsEqual As Long
    
    'SPAPI SPErr (*AllocateBlock)( size_t size, void **block );
    '/** Frees a block of memory allocated with \c #AllocateBlock().
    '        @param block The block pointer.
    '        @see \c #SPBlocksSuite::FreeBlock()
    '    */
    pAllocateBlock As Long
    
    'SPAPI SPErr (*FreeBlock)( void *block );
    '/** Reallocates a block previously allocated with \c #AllocateBlock().
    '    Increases the size without changing the location, if possible.
    '        @param block The block pointer.
    '        @param newSize The new number of bytes.
    '        @param newblock [out] A buffer in which to return the new block pointer.
    '        @see \c #SPBlocksSuite::ReallocateBlock()
    '    */
    pFreeBlock As Long
    
    'SPAPI SPErr (*ReallocateBlock)( void *block, size_t newSize, void **newblock );
    '/** A function pointer for unloaded suites. This is a protective measure
    '    against other plug-ins that may mistakenly use the suite after they have
    '    released it.
    '
    '    A plug-in that exports a suite should unload the suite's procedure pointers
    '    when it is unloaded, and restore them when the plug-in is reloaded.
    '    \li On unload, replace the suite's procedure pointers
    '        with the address of this function.
    '    \li On reload, restore the suite's procedure
    '        pointers with the updated addresses of their functions.
    '
    '    For example:
    '@code
    '    SPErr UnloadSuite( MySuite *mySuite, SPAccessMessage *message ) {
    '        mySuite->functionA = (void *) message->d.basic->Undefined;
    '        mySuite->functionB = (void *) message->d.basic->Undefined;
    '    }
    '
    '
    '    SPErr ReloadSuite( MySuite *mySuite, SPAccessMessage *message ) {
    '        mySuite->functionA = functionA;
    '        mySuite->functionB = functionB;
    '    }
    '@endcode
    '    */
    pReallocateBlock As Long
    
    'SPAPI SPErr (*Undefined)( void );
    pSpErr As Long

End Type

Private Type AboutRecord
    
    'void *      platformData;       /**< Platform specific information. See PlatformData in PITypes.h. */
    pPlatformDataHwnd As Long   'pointer to platform-specific data; on Windows, this is an hWnd of the owner window
    
    'SPBasicSuite *sSPBasic;         /**< PICA basic suite */
    pSPBasicSuite As Long
    
    'void        *plugInRef;         /**< Plug-in reference used by PICA */
    pPluginRef As Long
    
    'char        reserved [244];     /**< Reserved bytes. Set to zero */
    pReserved(0 To 243) As Byte
    
End Type

Private Enum PS_FilterSelector
    fs_About = 0
    fs_Parameters = 1
    fs_Prepare = 2
    fs_Start = 3
    fs_Continue = 4
    fs_Finish = 5
End Enum

#If False Then
    Private Const fs_About = 0, fs_Parameters = 1, fs_Prepare = 2, fs_Start = 3, fs_Continue = 4, fs_Finish = 5
#End If

'If the underlying library is loaded (via LoadLibrary), this handle will be non-0.  Note that LoadLibrary itself
' refcounts Load/Free calls, so we don't need to worry about multiple classes loading the same DLL.  (And really,
' FreeLibrary is mostly done to satisfy my OCD - the OS auto-frees loaded libraries when the base .exe terminates.)
' My bigger concern is memory usage if the user touches a ton of 8bf files in a single session; in that case,
' especially given that safe memory management is *not* guaranteed by a lot of 8bf files - unloading plugins
' post-usage can provide non-trivial stability and resource-reclamation benefits.
Private m_hLib As Long

'Filename of the underlying plugin.  Note that this is *not* unique per-plugin, as a single plugin file
' can expose multiple plugin interfaces.
Friend Function GetFilename() As String
    GetFilename = m_Filename
End Function

'Category of the underlying plugin.
Friend Function Get8bfCategory() As String
    Get8bfCategory = m_Category
End Function

'Name of the underlying plugin.
Friend Function Get8bfName() As String
    Get8bfName = m_Name
End Function

'Used internally to sort category/name entries alphabetically (first by category, second by plugin name)
Friend Function Get8bfSortKey() As String
    Get8bfSortKey = m_Category & "-" & m_Name
End Function

'Given a list of raw PiPL properties (and the target file from which they came), initialize this class.
' Returns TRUE if the property list is valid and compatible with PD; FALSE otherwise.  The plugin should
' not be exposed to the user if FALSE is returned.
Friend Function Initialize8bf_FromFile(ByRef srcFile As String, ByVal numProperties As Long, ByRef arrayOfProperties() As PIProperty) As Boolean
    
    Initialize8bf_FromFile = False
    m_Filename = srcFile
    
    'If this class already holds a plugin instance, free it
    UnloadPlugin
    
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
    Dim kindOK As Boolean, entryPointOK As Long
    
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
                    If DEBUG_VERBOSE Then PDDebug.LogAction "entry point: " & m_EntryPoint
                End If
                
                'The target DLL isn't actually loaded right now; we'll formally validate the entry point
                ' IFF the user actually interacts with this plugin
                entryPointOK = (LenB(m_EntryPoint) <> 0)
                
            'Next are non-code properties
            
            '// 'vers' \<int32\> Major\<int16\>.Minor\<int16\> version number:
            Case "vers"
                If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_Properties(i).propertyLength, VarPtr(m_Properties(i).pPropertyData(0))) Then
                    m_VersionMinor = cStream.ReadIntUnsigned()
                    m_VersionMajor = cStream.ReadIntUnsigned()
                    If DEBUG_VERBOSE Then PDDebug.LogAction "version: " & m_VersionMajor & "." & m_VersionMinor
                End If
                
                'TODO: validate versionOK here??
            
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
                    If DEBUG_VERBOSE Then PDDebug.LogAction "category: " & m_Category
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
                    If DEBUG_VERBOSE Then PDDebug.LogAction "name: " & m_Name
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
    
    'This plugin can be exposed to the user if...
    ' 1) It's actually an 8bf plugin
    ' 2) It exposes a valid entrypoint
    ' 3) It has a name
    ' 4) It has a category
    Initialize8bf_FromFile = kindOK And entryPointOK And (LenB(m_Name) <> 0) And (LenB(m_Category) <> 0)
    
End Function

'Returns TRUE if the plugin has a valid About dialog interface (as far as we can tell)
Friend Function IsAboutDialogPossiblyAvailable() As Boolean
    
    'It's impossible to know if an About dialog is available in advance.  Instead, you have to call the
    ' main entry point with the "About dialog ID", then check the return value.
    ' Of course, a lot of plugins *don't* show an About dialog and then return "success", so YMMV.
    IsAboutDialogPossiblyAvailable = EnsureLibLoaded()
    
End Function

'Show the About dialog associated with this plugin.  Returns TRUE if (as far as we know)
' the dialog was accessed successfully.
Friend Function ShowAboutDialog(ByVal srcHostHWnd As Long) As Boolean
    
    ShowAboutDialog = False
    
    'Ensure we have an entry point to use for the About dialog display request
    If (Not EnsureLibLoaded) Then Exit Function
    If (Not Me.IsAboutDialogPossiblyAvailable()) Then Exit Function
    
    'Failsafe only; previous step should have validated this
    If (m_pEntryPoint = 0) Then Exit Function
    
    'Prep the basic structs required for About display.  (These are very simple.)
    Dim tmpAboutRecord As AboutRecord
    tmpAboutRecord.pPlatformDataHwnd = VarPtr(srcHostHWnd)
    
    'Everything else in this struct can left be uninitialized/0
    
    'Results are returned via pointer to a 2-byte int
    Dim adobeReturn As Integer
    adobeReturn = 0
    
    'Plugin data for the caller (?)
    Dim pPluginData As Long
    pPluginData = 0
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Showing About dialog for " & m_Filename
    CallCDeclW m_pEntryPoint, vbEmpty, fs_About, VarPtr(tmpAboutRecord), VarPtr(pPluginData), VarPtr(adobeReturn)
    ShowAboutDialog = (adobeReturn = ps_Success)
    If DEBUG_VERBOSE And (Not ShowAboutDialog) Then PDDebug.LogAction "entry point return: " & adobeReturn
    
End Function

'Before touching any function calls in the 8bf file, call this function to ensure LoadLibrary has been called.
Private Function EnsureLibLoaded() As Boolean
    
    Const FUNC_NAME As String = "EnsureLibLoaded"
    EnsureLibLoaded = False
    
    If (m_hLib <> 0) Then
        EnsureLibLoaded = True
        Exit Function
    End If
    
    'If we're still here, the library has not been loaded yet.  Attempt to load it.
    ' Note that m_Filename *MUST* be set before calling this function (i.e. you must have called
    ' Initialize8bf_FromFile() and the plugin must have passed initial validation).
    If (LenB(m_Filename) = 0) Then
        InternalError FUNC_NAME, "no filename"
        Exit Function
    End If
    
    'Previously we checked bitness here, but I've since offloaded that to the Plugin_8bf module
    ' which already loads the DLL and "walks" it to retrieve properties - it's easier to check bitness there,
    ' and simply *not* forward any wrong-bitness libraries here in the first place.
    
    'Load the DLL
    m_hLib = VBHacks.LoadLib(m_Filename)
    If (m_hLib <> 0) Then
        
        'Attempt to retrieve the entry point self-defined by the plugin
        m_pEntryPoint = GetProcAddress(m_hLib, m_EntryPoint)
        EnsureLibLoaded = (m_pEntryPoint <> 0)
        If DEBUG_VERBOSE And (Not EnsureLibLoaded) Then InternalError FUNC_NAME, "GetProcAddress failed"
        
    Else
        If DEBUG_VERBOSE Then InternalError FUNC_NAME, "LoadLibrary failed; last DLL error: " & Err.LastDllError
        Exit Function
    End If
    
End Function

'Unloading the plugin doesn't just erase local properties - it also releases the associated library handle.
' (This is important as some plugins self-manage memory allocations, which can be *huge* depending on image size.)
Friend Function UnloadPlugin() As Boolean
    
    m_numProperties = 0
    Erase m_Properties
    
    'The most important step of unloading a plugin is freeing the underlying library
    If (m_hLib <> 0) Then VBHacks.FreeLib m_hLib
    UnloadPlugin = (m_hLib = 0)
    
End Function

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal pEntryPoint As Long, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

    Dim i As Long, vTemp() As Variant, hResult As Long
    
    Dim numParams As Long
    If (UBound(pa) < LBound(pa)) Then numParams = 0 Else numParams = UBound(pa) + 1
    
    If IsMissing(pa) Then
        ReDim vTemp(0) As Variant
    Else
        vTemp = pa 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
    End If
    
    For i = 0 To numParams - 1
        If VarType(pa(i)) = vbString Then vTemp(i) = StrPtr(pa(i))
        m_vType(i) = VarType(vTemp(i))
        m_vPtr(i) = VarPtr(vTemp(i))
    Next i
    
    Const CC_CDECL As Long = 1
    hResult = DispCallFunc(0, pEntryPoint, CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    
End Function

Private Sub Class_Initialize()

    'Initialize all module-level arrays used for DispCallFunc interop
    ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
    ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
        
End Sub

Private Sub InternalError(ByRef errFuncName As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    PDDebug.LogAction "WARNING!  Problem in pd8bf." & errFuncName & ": " & errDescription
    If (errNum <> 0) Then PDDebug.LogAction "  (If it helps, an error number was also reported: #" & errNum & ")"
End Sub

Private Sub Class_Terminate()
    UnloadPlugin
End Sub
