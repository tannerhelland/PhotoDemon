VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "pdPackage" Interface (e.g. Zip-like archive handler)
'Copyright ©2013-2014 by Tanner Helland
'Created: 05/April/13
'Last updated: 05/April/14
'Last update: initial build
'
'This class provides an interface for creating and reading "pdPackage" files.  pdPackages are zip-like archive files
' that contain one or more "nodes" (e.g. pieces of data), compressed or uncompressed, in a VB-friendly structure.
'
'Though I have created pdPackages specifically for use with PhotoDemon, this class should be easily usable by others
' with only minor modifications.  Note that an explicit path to a STDCALL (sometimes called WAPI) variant of zlib.dll
' is required, with the expected filename of "zlibwapi.dll".  If your STDCALL copy is called "zLib.dll", you must
' rewrite the few zLib-specific API declarations at the top of the file to match.
'
'While pdPackages have many similarities to ZIP files, THEY ARE NOT ACTUAL ZIP FILES, and this class cannot read or
' write actual ZIP files.  pdPackages are, by design, much simpler than ZIP files, and their structure and layout
' is much more friendly for VB coders.
'
'A few key pdPackage features include:
'
' 1) Data agnosticism, e.g. everything is treated as byte arrays.
' 2) Front-loaded header.  ZIP files place the header at the tail of the archive, but pdPackages place the header
'     at the head.  This is not ideal for files that must be repeatedly edited, but it allows for much faster archive
'     reading and writing, especially when accessing only a single file.
' 3) Fixed-width directory entries.  This allows the entire archive directory to be read in a single operation,
'     rather than manually parsing variable-width directory entries until all have been located.
' 4) Support for zLib-style compression on a per-node basis.
' 5) Support for CRC32 (adler-style) validation of each individual node.
' 6) Support for two data entries per node, typically a header byte array and an actual data byte array.  These two structs
'     don't need to be used (one or the other or neither is just fine), but I find it very helpful to be store two pieces of
'     data per node (typically so I can read key data from a node without extracting its full contents).
' 7) Per-node compression options, meaning you can extract a single file without having to decompress the entire
'     archive, and you can customize compression for each given file in a package.
'
'Here are a few things to note if you are using this class in your own projects:
'
' 1) At present, pdPackage files are not easily editable.  Once one is created, it cannot easily have nodes added or erased.
'     There's nothing in the spec that prevents this, but I just haven't written edit functions because I have no reason to
'     do so in PhotoDemon.
'
' 2) As noted above, zLib is required for compression.  The spec has been written so that you can easily add your own
'     compression (or encryption) functions if you want, and of course you don't have to compress nodes if you don't want to.
'
' 3) Up to 2GB of data is theoretically supported, but you won't be able to reach that amount from within VB.  For
'     performance reasons, this pdPacker class creates the full archive in RAM before writing it to file.  This makes it
'     very fast, but ill-suited to extraordinarily large archive sizes.
'
' 4) When reading pdPackage files, the full file contents will be cached in memory.  Do not reload the pdPackage file unless
'     you absolutely have to!  Once it is cached, you can access individual nodes without reloading the entire file.
'
'Input on the format and this class interface are always welcome, as are code fixes/changes/improvements.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This constant should be updated whenever the core assembly/disassembly code is modified.  This value is embedded in all
' pdPackage files written by this class, as a failsafe against potential ABI breakage in the future.  The lowest expected
' value for this constant is 64, representative of PhotoDemon 6.4, when the format was first implemented.
Private Const THIS_PDPACKAGE_VERSION As Long = 64

'The first four bytes of a valid pdPackage file must always be &h5044504B ("PDPK" in ASCII)
Private Const PDP_UNIVERSAL_IDENTIFIER As Long = &H5044504B

'Each pdPackage file has a short file header.  This header is separate from the node directory, and contains all information
' necessary to prepare a node directory array
Private Type PDP_HEADER
    PDP_ID As Long                          'pdPackage Identifier; must always be &h5044504B ("PDPK" in ASCII)
    PDP_SubID As Long                       'pdPackage sub-identifier.  This can be used by callers to define a specific type of pdPackage.  (PD uses this to denote PDI files.)
    PDP_Version As Long                     'Version number of the pdPackage class used to write this file.
    NodeCount As Long                       'Number of data nodes in this package, 1-based (e.g. if there is one node in the archive, this value will be 1).
    NodeStructSize As Long                  'Size of an individual node struct.  This can be used as a failsafe check against the PDP_Version, above.
    DirectoryChunkSize As Long              'Size of the full node directory structure, including all node directory entries.
    DirectoryFlags(0 To 3) As Long          'User-defined flags for the directory chunk.  These are unused at present, but could eventually be used to note directory compression, encryption, etc.
    DataChunkSize As Long                   'Size of the data chunk of the archive.  This could be inferred by calculating EOF - (End of Directory), but it's easier to simply note it here.
    DataFlags(0 To 3) As Long               'User-defined flags for the data chunk.  These are unused at present, but could eventually be used to note a second compression pass over the full data chunk, encryption, etc.
    Reserved As Long                        'Reserved for future use; no relevance at present.
End Type

'Immediately following the PDP_HEADER is the directory chunk, which is comprised of NodeCount individual PDP_NODE structs.
' These structs are small and flexible, and *they have a fixed size*, meaning they can be read into a fixed-width array
' in a single pass.
Private Type PDP_NODE

    NodeName(0 To 31) As Byte               '32 characters are allotted for each node name (in single-byte Windows-1252 format)
    NodeID As Long                          'Alternatively, calling functions can specify an optional 4-byte numerical ID.  Nodes can be read by 32-char name or 4-byte ID.
    OptionalNodeType As Long                'Calling functions can also assign each node a TYPE if they want; this value has no meaning to this class.
    NodeFlags(0 To 3) As Long               '16 bytes of node-specific flags are allowed.  At present, these are unused.
    
    'One of the unique features of pdPackages is that each node is allotted two entries in the data chunk.  These entries don't have
    ' to be used; in fact, neither has to be used, but they can be helpful for reading node-specific information without having to
    ' decode the entire node contents.
    NodeHeaderOffset As Long                'Absolute offset of this node's header in the data chunk, STARTING FROM THE START OF THE DATA CHUNK, not the start of the file!
    NodeHeaderCompSize As Long              'Compressed size of this node's header.  (This is the size the node's header occupies in the data chunk, even if the header is uncompressed!)
    NodeHeaderOrigSize As Long              'Original size of this node's header.  (If this value is the same as NodeHeaderCompSize, the node header was stored uncompressed.)
    NodeDataOffset As Long                  'Absolute offset of this node's data in the data chunk, STARTING FROM THE START OF THE DATA CHUNK, not the start of the file!
    NodeDataCompSize As Long                'Compressed size of this node's data.  (This is the size the node's data occupies in the data chunk, even if the data is uncompressed!)
    NodeDataOrigSize As Long                'Original size of this node's data.  (If this value is the same as NodeHeaderCompSize, the node data was stored uncompressed.)
    
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'zLib compression and CRC validation
Private Declare Function compress Lib "zlibwapi.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlibwapi.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function zlibVersion Lib "zlibwapi.dll" () As Long

'API calls for explicitly calling dlls.  This allows us to build DLL paths at runtime, and it also allows
' us to call any DLL we like without first passing them through regsvr32.
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'If zLib has been successfully initialized, this will be set to TRUE.  If we initialized the DLL within this class, we will also
' have a zLibToken value that must be freed at termination time.
Private m_ZLibAvailable As Boolean, m_ZLibToken As Long

'When writing new pdPackage files, these variables will hold the package contents as they are being assembled in memory.
Private m_FileHeader As PDP_HEADER
Private m_NodeDirectory() As PDP_NODE
Private m_DataChunk() As Byte

Private m_NumOfNodes As Long
Private m_CurDataOffset As Long, m_CurDataSize As Long


'pdPackage operations are roughly divided into two groups:
' - GET operations, for retrieving data from existing pdPackage files
' - SET operations, for creating new pdPackage files
'
'At present, these two types of operations do not interact, meaning you cannot use SET operations to modify a pdPackage
' you have loaded using GET operations.  Packages must be created and written to file in one fell swoop, and if you want
' to read a pdPackage, I strongly recommend creating a dedicated class for just that file (due to the way pdPacker caches
' file contents).


'Before creating a pdPackage file, you must call this function once.  It preps all internal structures in anticipation of
' data loading.  If you know the number of data nodes you will be writing, you can mention it in advance, which makes the
' directory assembly process much faster (because we don't have to ReDim Preserve the directory when we run out of space).
Public Sub prepareNewPackage(Optional ByVal numOfDataNodes As Long = 0)

    'Reset all module-level storage structs related to writing a new pdPackage file
    
    'Prepare the directory contents
    m_NumOfNodes = numOfDataNodes
    ReDim m_NodeDirectory(0 To m_NumOfNodes) As PDP_NODE
    
    'Prepare the data chunk
    m_CurDataOffset = 0
    m_CurDataSize = 0
    ReDim m_DataChunk(0 To 1023) As Byte

End Sub

Public Sub addNode()

End Sub


'If you want to use compression functions, you must provide the class with a path to a STDCALL (WAPI) copy of zLib, including
' "/zlibwapi.dll" at the end of the path.  This class assumes a file named "zlibwapi.dll"; if you are using one called zLib.dll,
' you must rewrite the zLib API function declarations to match.
'
'Optionally, if the caller wants to handle LoadLibrary/FreeLibrary by itself, it is welcome to do so.  (If you plan on using
' multiple pdPackage instances within a single application, it makes more sense to have the caller handle this.)  Simply pass
' TRUE as the second parameter and this class will assume you've already used LoadLibrary successfully.
Public Function init_ZLib(ByVal pathToZLibDLL As String, Optional ByVal leaveInitToCaller As Boolean = False) As Boolean

    If leaveInitToCaller Then
        m_ZLibAvailable = True
    
    'The caller wants us to handle zLib initialization.  Attempt to load the zLib library path we were passed.
    Else
    
        If Len(pathToZLibDLL) > 0 Then
        
            'If the library was successfully loaded, a non-zero value will be returned.
            m_ZLibToken = LoadLibrary(pathToZLibDLL)
            
            If m_ZLibToken <> 0 Then
                m_ZLibAvailable = True
            Else
                m_ZLibAvailable = False
            End If
            
        Else
            m_ZLibAvailable = False
        End If
        
    End If
    
    'This function will return TRUE if it considers zLib to be available, and FALSE if it does not.
    init_ZLib = m_ZLibAvailable

End Function

Private Sub Class_Initialize()

    'Reset all module-level variables
    m_ZLibAvailable = False
    m_ZLibToken = 0
    
End Sub

Private Sub Class_Terminate()

    'If we loaded zLib, free it now
    If m_ZLibToken > 0 Then FreeLibrary m_ZLibToken

End Sub
