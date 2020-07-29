VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPackager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "pdPackage" v2.0 Interface (e.g. Zip-like archive handler)
'Copyright 2014-2018 by Tanner Helland
'Created: 05/April/14
'Last updated: 23/September/17
'Last update: revert to writing original, uncompressed data if compression doesn't reduce data size
'Dependencies: - pdStream class from photodemon.org (used to easily read/write memory and file buffers)
'              - pdFSO class from photodemon.org (Unicode-friendly file operations)
'              - pdStringStack class from photodemon.org (optimized handling of large string collections)
'              - VBHacks module from photodemon.org (optimized workarounds for lacking VB functionality)
'              - Compression module and any attached plugin modules (e.g. Plugin_zlib, Plugin_lz4) from photodemon.org
'                 (if you want compression support)
'
'This class provides an interface for creating and reading "pdPackage" files.  pdPackages are zip-like archive files
' that contain one or more "nodes" (e.g. pieces of data), compressed or uncompressed, in a VB-friendly structure.
'
'pdPackager v2.0 is not backwards compatible with pdPackage v1.0.  It is an entirely separate format, with many
' improvements over v1.0.
'
'Though I have created pdPackages specifically for use with PhotoDemon, this class should be easily usable by others
' with only minor modifications.  Note that compression support requires an explicit path to an STDCALL (sometimes
' called WAPI) variant of either zlib, zstandard (zstd), or lz4/lz4-hc, with the expected filename of "zlibwapi.dll",
' "libzstd.dll", or "liblz4.dll", respectively.  pdPackage can also be used without compression, though its benefits
' are greatly reduced.  (And obviously, you will not be able to open compressed archives from other sources.)
'
'While pdPackages have many similarities to ZIP files, THEY ARE NOT ACTUAL ZIP FILES, and this class cannot read or
' write actual ZIP files.  pdPackages are, by design, much simpler than ZIP files, and their structure and layout
' plays better with VB's inherent limitations.
'
'Key pdPackage features include:
'
' 1) Data agnosticism, e.g. everything is treated as byte arrays, so you can store whatever data you'd like.
' 2) Very small front-loaded header, and rear-loaded directory.  A small front-loaded header allows for quick file
'     validation, but like zip files, the full directory exists at the header of the archive.  This allows us to
'     write pdPackages "as we go", which greatly improves both memory usage and performance.
' 3) Fixed-width directory entries.  This allows the entire archive directory to be read in a single operation,
'     rather than manually parsing variable-width directory entries until all have been located.
' 4) Support for compression on a per-node basis.  zLib and zstd are both supported, with zstd being recommended
'     whenever possible (due to better performance and compression ratios).
' 5) Support for two data "entries" per node, typically a header chunk and an actual data chunk.  These two structs
'     don't need to be used (one or the other or neither is just fine), but I greatly appreciate the simplicity of
'     storing two pieces of data per node.  For example, when using pdPackage's automatic file and folder compression
'     features, the "header" node stores the original filename for each file, while the "data" node stores the file
'     itself.  This makes it easy to iterate filenames without touching full file contents.
' 6) A per-node access system, including compression, so that you can easily extract a single node without needing to
'     decompress the entire archive.  Also, this allows you to use different compression settings for each node
'     (e.g. not everything needs to be compressed - only the bits you find relevant).
'
'Here are a few things to note if you want to use this class in your own projects:
'
' 1) At present, pdPackage files are not easily editable.  Once created, there is no interface for adding new nodes,
'     erasing existing nodes, or modifying any of the pdPackage settings (compression, etc).  There's nothing in the format
'     design that prevents these actions, but I haven't written edit functions because I have no reason to do so in PhotoDemon.
' 2) As noted above, external libraries (zLib, zstd, and/or lz4/lz4-hc) are required for compression.  pdPackages are designed
'     in a way that makes it easy to use any compression functions (or other modification functions, e.g. encryption) of your
'     choosing, or to ignore compression entirely if you don't require it.  That said, if you want to use the class without a
'     compression library, but you intend to work with pdPackage files from other sources (which may use compression), you need
'     to make use of the GetPackageFlag() function and the accompanying PDPF2_(library-name)_REQUIRED flags.  These will
'     identify files your software cannot process without compression libraries.
' 3) Up to 2GB of data is theoretically supported, but you won't be able to reach that amount from within VB.  The
'     rear-loaded directory format makes it possible to support larger file sizes, but I have not yet done the work to
'     implement this.  Sorry.  (Patches welcome!)
' 4) When reading pdPackage files, relevant file bits will only be loaded as-necessary.  As such, you cannot delete the
'     original file until all interactions are complete.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This constant should be updated whenever the core assembly/disassembly code is modified.  This value is embedded in all
' pdPackage files written by this class, as a failsafe against potential ABI breakage in the future.
' - The current expected value is 70.  Versions 66 and previous require the original pdPackage v1.0 class.  This class
'   *does not load pdPackage versions prior to v70*, by design.
' - The constants between 67 and 69, inclusive, have no meaning and are considered invalid.
' - The lowest supported value for this constant is 70, representative of PhotoDemon 7.0, when the PDI format was
'   overhauled to support a bunch of new features (like zstd compression).
Private Const THIS_PDPACKAGE_VERSION As Long = 70

'The first four bytes of a valid pdPackage file must always be &H4B504450 ("PDPK" in ASCII, if viewed in a text editor
' - note the little-endian declaration here).  Note that this constant is shared with the old pdPackage v1.0 format.
Private Const PDP_UNIVERSAL_IDENTIFIER As Long = &H4B504450

'Because pdPackage is awesome, it can store some complicated data structures automatically, with no work on the caller's part.
' Structures like this are identified with special internal names, which simplifies the extraction process.
Private Const PDP_DATA_FILE As String = "PDPFILE"
Private Const PDP_DATA_FILEANDPATH As String = "PDPFILE_AND_PATH"

Private Enum PDP_SPECIAL_NODE_TYPE
    PDP_SN_Filename = 0
    PDP_SN_FilenameAndRelativePath = 1
End Enum

#If False Then
    Private Const PDP_SN_Filename = 0, PDP_SN_FilenameAndRelativePath = 1
#End If

'The number of flags PD supports is relatively thin at present.  This is by design, as I don't want to use them for trivial shit.
' Note that in an effort to keep flags as simple as possible, they are referred to by bit ordinal position [0, 31] rather than raw hex value.
' Also note that at present, the same flags are used for all possible flag locations (directory, data, and individual nodes).  Future flags
' don't have to be used this way, but at present, each location shares enough properties to warrant use of the same flag values.

'FYI: flags can be stored in the directory and data chunks, as well as individual nodes.  When checking flags (e.g. after a load operation),
' the caller can specify which flag location they want checked.
Public Enum PDP_FLAG_LOCATION
    PDP_LOCATION_ANY = -1
    PDP_LOCATION_DIRECTORY = 0
    PDP_LOCATION_DATA = 1
    PDP_LOCATION_INDIVIDUALNODE = 2
End Enum

#If False Then
    Const PDP_LOCATION_ANY = -1, PDP_LOCATION_DIRECTORY = 0, PDP_LOCATION_DATA = 1, PDP_LOCATION_INDIVIDUALNODE = 2
#End If

'Each pdPackage file has a short file header.  This header is separate from the directory chunk, and it contains all information
' necessary to prepare a node directory array.
Private Type PDP_HEADER
    PDP_ID As Long                          'pdPackage Identifier; must always be &H4B504450 ("PDPK" in ASCII)
    
    PDP_SubID As Long                       'pdPackage sub-identifier.  This can be used by callers to define a specific type of pdPackage.
                                            ' (For example, PD uses this to denote PDI files specifically.)  This value has no intrinsic meaning
                                            ' to this class.
    PDP_Version As Long                     'Version number of the pdPackage class used to write this file.  All parsing behavior is based on this
                                            ' value, and it is automatically set according to the THIS_PDPACKAGE_VERSION constant at the top of
                                            ' this class.  (Note that you *must* check this version number manually if you want to support the old
                                            ' pdPackage 1.0 format.  This class does not handle versions prior to 70.)
    NodeCount As Long                       'Number of data nodes in this package, 1-based (e.g. if there is one node in the archive, this value will
                                            ' be set to 1).  Cannot currently be zero, FYI.
    NodeStructSize As Long                  'Size of an individual node struct.  This can be used as a failsafe check against PDP_Version, above.
    
    DirectoryChunkPosition As Long          'Position of the first byte in the directory byte stream, stored as an absolute value from position 0
                                            ' (e.g. this value will *always* be larger than the size of this header struct, and the size of this
                                            ' package's data chunk).
    DirectoryChunkSizePacked As Long        'Size of the directory chunk bytestream, as stored in the file.  If the directory chunk is compressed,
                                            ' this value will be different from DirectoryChunkSizeFinal, below.
    DirectoryChunkSizeOriginal As Long      'Size of the full node directory structure, including all node directory entries, when uncompressed.
                                            ' You can compare this value to DirectoryChunkSizeStored to determine if compression was applied.
    DirectoryFlags(0 To 2) As Long          'User-defined flags for the directory chunk.  See the PDP_FLAGS_V2 enum for details.
    
    DataChunkPosition As Long               'Position of the first byte in the data byte stream, stored as an absolute value from position 0
                                            ' (e.g. this value will *always* be larger than the size of this header struct).
    DataChunkSizePacked As Long             'Size of the data chunk bytestream, as stored in this package.  This could be inferred by calculating
                                            ' the difference between the end of the file header and the start of the directory chunk, but it's
                                            ' easier to simply note it right inside the header.
    DataChunkSizeOriginal As Long           'Size of the data chunk bytestream post-decompression.  At present, this will always be identical to
                                            ' the DataChunkSizeStored value, as compression is not currently supported for the entire data chunk.
    DataFlags(0 To 2) As Long               'User-defined flags for the data chunk.  See the PDP_FLAGS_V2 enum for details.
    
    Reserved As Long                        'Reserved for future use; no relevance at present.
    
End Type

'The number of flags PD supports is relatively thin at present.  This is by design, as I don't want to use them for trivial shit.
' Note that in an effort to keep flags as simple as possible, they are referred to by bit ordinal position [0, 31] rather than raw hex value.

'Also note that there are two sets of flags: one set for the file header (which describes the file as a whole, e.g. the ZLibRequired flag means
' "something in this file requires ZLib to be present), and a second set for individual nodes (which describe *just* that node,
' e.g. the ZlibRequiredHeader flag means "this node's header chunk requires ZLib for decompression").

'FYI: flags can be stored in the directory and data chunks, as well as individual nodes.  When checking flags (e.g. after a load operation),
' the caller can specify which flag location they want checked.
Public Enum PDP_FLAG_SOURCE
    PDP_FS_Any = -1
    PDP_FS_Directory = 0
    PDP_FS_Data = 1
    PDP_FS_IndividualNode = 2
End Enum

#If False Then
    Private Const PDP_FS_Any = -1, PDP_FS_Directory = 0, PDP_FS_Data = 1, PDP_FS_IndividualNode = 2
#End If

'The list of currently supported flags is as follows:
Public Enum PDP_HEADER_FLAGS_V2
    PDP_HF2_FileCollection = 0 'This file contains zip-like file+folder entries created from the AutoAddNode(s) function(s).
    PDP_HF2_ZlibRequired = 1   'There are one or more compressed entries in this chunk, so zLib will be required to read it
    PDP_HF2_ZstdRequired = 2   'There are one or more zstd-compressed entries in this chunk
    PDP_HF2_Lz4Required = 3    'There are one or more lz4-compressed entries in this chunk.  Note that this flag doubles for both LZ4 and LZ4-HC, as they use identical decompressors.
End Enum

#If False Then
    Private Const PDP_HF2_ZlibRequired = 0, PDP_HF2_FileCollection = 1, PDP_HF2_ZstdRequired = 2, PDP_HF2_Lz4Required = 3
#End If

Public Enum PDP_NODE_FLAGS_V2
    PDP_NF2_FileNode = 0       'This node contains a file added via the shortcut AutoAddNodeFromFile() function.
    PDP_NF2_ZlibRequiredH = 1  'There are one or more compressed entries in this chunk, so zLib will be required to read it
    PDP_NF2_ZstdRequiredH = 2  'There are one or more zstd-compressed entries in this chunk
    PDP_NF2_Lz4RequiredH = 3   'There are one or more lz4-compressed entries in this chunk.  Note that this flag doubles for both LZ4 and LZ4-HC, as they use identical decompressors.
    PDP_NF2_ZlibRequired = 4   'There are one or more compressed entries in this chunk, so zLib will be required to read it
    PDP_NF2_ZstdRequired = 5   'There are one or more zstd-compressed entries in this chunk
    PDP_NF2_Lz4Required = 6    'There are one or more lz4-compressed entries in this chunk.  Note that this flag doubles for both LZ4 and LZ4-HC, as they use identical decompressors.
End Enum

#If False Then
    Private Const PDP_NF2_FileNode = 0, PDP_NF2_ZlibRequiredH = 1, PDP_NF2_ZstdRequiredH = 2, PDP_NF2_Lz4RequiredH = 3, PDP_NF2_ZlibRequired = 4, PDP_NF2_ZstdRequired = 5, PDP_NF2_Lz4Required = 6
#End If

'Immediately following the PDP_HEADER is the data chunk, and following the data chunk is the data directory.  The directory
' is comprised of FileHeader.NodeCount individual PDP_NODE structs.  These structs are small and flexible, and *they have a
' fixed size*, meaning they can be read into a fixed-width array in a single pass.
Private Type PDP_NODE

    nodeName As String * 32                 'Name of the node, as a standard VB String (DBCS).  It is thus 64 bytes long.
                                            ' Note that node name has no special meaning to this class; it is up to the caller to make
                                            ' the name meaningful.
    NodeID As Long                          'Alternatively, calling functions can specify an optional 4-byte numerical ID.  Nodes can be
                                            ' read by 32-char name or 4-byte ID; this decision is up to the caller.
    OptionalNodeType As Long                'Calling functions can also assign each node a 4-byte TYPE identifier if they want.  This value
                                            ' has no intrinsic meaning to this class.
    NodeFlags(0 To 3) As Long               '16 bytes of node-specific flags are allowed.  At present, these are mostly unused.
    
    'One of the unique features of pdPackages is that each node is allotted two entries in the data chunk.  These entries don't have
    ' to be used; in fact, neither has to be used, but they can be helpful for reading node-specific information without having to
    ' decode the entire node contents.  (For example, when using the automatic file compression features, this class stores filenames
    ' inside the header entry, and the file contents inside the data entry.  This lets us extract the list of files very efficiently.)
    
    NodeHeaderOffset As Currency            'Absolute offset of this node's header in the data chunk, STARTING FROM THE START OF THE DATA CHUNK, not the start of the file!
    NodeHeaderPackedSize As Currency        'Packed size of this node's header chunk.  (This is the size the node's header array occupies in the pdPackage data chunk.)
    NodeHeaderOriginalSize As Currency      'Original size of this node's header chunk.  (If this value is the same as NodeHeaderPackedSize, the node header was stored uncompressed.)
    
    NodeDataOffset As Currency              'Absolute offset of this node's data in the data chunk, STARTING FROM THE START OF THE DATA CHUNK, not the start of the file!
    NodeDataPackedSize As Currency          'Packed size of this node's data chunk.  (This is the size the node's data array occupies in the pdPackage data chunk.)
    NodeDataOriginalSize As Currency        'Original size of this node's data chunk.  (If this value is the same as NodeHeaderPackedSize, the node data was stored uncompressed.)
    
End Type

'When writing new pdPackage files, these variables hold the package's contents as they are being assembled.  When loading a pdPackage file,
' these structs will be filled after the file header is successfully verified.
Private m_FileHeader As PDP_HEADER
Private m_NodeDirectory() As PDP_NODE
Private m_numOfNodes As Long

'The actual data chunk of the pdPackage is assembled using a pdStream instance.  It greatly simplifies the process of
' assembling a 1D byte array from discrete individual chunks.
Private m_DataBuffer As pdStream

'As of v2, pdPackages can be wrapped directly around source files, which greatly reduces memory usage compared to assembling
' everything in-memory.  The access mode of the current package instance (including read/write access) is stored here.
Private m_StreamMode As PD_STREAM_MODE, m_StreamAccess As PD_STREAM_ACCESS

'Certain modes may require us to use temporary buffers for compression/decompression.  We try to use this as little as possible,
' but sometimes it's inevitable.
Private m_CompressionBuffer() As Byte

'File operations are made easier by using the pdFSO class, which wraps a bunch of Unicode-friendly file APIs
Private m_File As pdFSO


'pdPackage operations are roughly divided into two groups:
'
' - GET operations, for retrieving data from existing pdPackage files.  GET operations include:
'        readPackageFromFile, getNodeInfo, getNodeDataByID/Index/Name, getPackageFlag
'
' - SET operations, for creating new pdPackage files.  SET operations include:
'        prepareNewPackage, addNode, addNodeData, writePackageToFile
'
'At present, these two types of operations do not interact reliably, meaning you cannot use SET operations to modify a
' pdPackage you have loaded using GET operations.  Packages must be created and written to file in one fell swoop, and
' if you want to read a pdPackage, I strongly recommend creating a dedicated class for just that file (due to the way
' this class caches file contents).


'Before creating a pdPackage file, you must call this function once.  It preps all internal structures in anticipation of
' data loading.  If you know the number of data nodes you will be writing, you can mention it in advance, which makes the
' directory assembly process much faster (because we don't have to ReDim Preserve the directory when we run out of space).
' Similarly, if you have some notion of how large the data chunk will be, you can supply it in advance.  There is generally
' no penalty to over-estimating space, as the buffer will be trimmed before writing it out to file.
Friend Sub PrepareNewPackage(Optional ByVal numOfDataNodes As Long = 0, Optional ByVal optPackageID As Long = 0, Optional ByVal estimatedDataChunkSize As Long = 0, Optional ByVal streamMode As PD_STREAM_MODE = PD_SM_MemoryBacked, Optional ByVal backingFilename As String = vbNullString)

    'Reset all module-level storage structs related to writing a new pdPackage file
    
    'Start by preparing the file header.  This will be updated before being written out to file, but we can set certain
    ' items in advance.
    With m_FileHeader
        .PDP_ID = PDP_UNIVERSAL_IDENTIFIER
        .PDP_SubID = optPackageID
        .PDP_Version = THIS_PDPACKAGE_VERSION
        .NodeCount = numOfDataNodes
        
        'Retrieve the size of a node struct.  This value should never change, but never is a long time, and this gives us
        ' a failsafe against things like mismatched PDP version numbers.  It also makes it easier for external load functions
        ' to know how to size the directory structure array.
        Dim tmpNode As PDP_NODE
        .NodeStructSize = LenB(tmpNode)
        
        'DirectoryChunkSize can be assumed from the numOfDataNodes, but note that it will be verified again before the data is
        ' actually written out to file.
        .DirectoryChunkSizeOriginal = .NodeStructSize * numOfDataNodes
        .DirectoryChunkSizePacked = 0
        
        'DirectoryFlags() and DataFlags() are set by separate functions.  For now, assume a value of 0 for all flags.
        Dim i As Long
        For i = 0 To UBound(.DirectoryFlags)
            .DirectoryFlags(i) = 0
            .DataFlags(i) = 0
        Next i
        
        'DataChunkSize will remain unknown until all nodes have been added.
        .DataChunkSizeOriginal = 0
        .DataChunkSizePacked = 0
        
        '4 bytes are reserved at the end as a "just in case".  For now, they should always be 0.
        .Reserved = 0
        
    End With
    
    'Resize the directory array to the number of supplied nodes; note that this step is optional; if no node count is supplied,
    ' the AddNode() function will automatically increment itself as necessary.  (We also cover the case where this array already
    ' exists from prior writes; we can skip initialization in that case, to save a bit of memory thrashing.)
    If (m_numOfNodes > 0) Then
        If (m_numOfNodes < numOfDataNodes) Then ReDim m_NodeDirectory(0 To numOfDataNodes) As PDP_NODE
    Else
        ReDim m_NodeDirectory(0 To numOfDataNodes) As PDP_NODE
    End If
    
    m_numOfNodes = 0
    
    'Prepare the data buffer.  Note that chunk size is largely irrelevant for our purposes; when handed a byte array that exceeds
    ' the size of the default chunk size, the buffer class is smart enough to extend the buffer by the size of that byte array.
    ' If writing a lot of small data, however, the chunk size becomes important; try to make it large enough to require minimal
    ' ReDim Preserve requests throughout the pdPackage assembly process.
    If (m_DataBuffer Is Nothing) Then
        Set m_DataBuffer = New pdStream
    Else
        If (m_StreamMode <> streamMode) Then m_DataBuffer.StopStream True
    End If
    
    m_StreamMode = streamMode
    
    If (streamMode = PD_SM_MemoryBacked) Then
        m_DataBuffer.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, , estimatedDataChunkSize, , , True
    
    'In file-backed mode, we'll actually open the file right away, and write our nodes to it "as we go"
    ElseIf (streamMode = PD_SM_FileBacked) Then
        m_DataBuffer.StartStream PD_SM_FileBacked, PD_SA_ReadWrite, backingFilename, estimatedDataChunkSize, LenB(m_FileHeader), OptimizeSequentialAccess
    End If

End Sub

'Add a new node to this pdPackage instance.  Note that this function DOES NOT actually add the node's data arrays to the main
' data buffer - those are done in a subsequent step, by the user, as necessary.  (It's a little pesky to separate node additions
' into multiple steps, but this allows for more fine-grained control over node addition, without overwhelming any function with
' a monstrous list of parameters.)
'
'Importantly, this function returns the index of the added node, which external functions must then use to supply this node's
' actual data arrays.
Friend Function AddNode(Optional ByVal thisNodeName As String = vbNullString, Optional ByVal thisNodeID As Long = 0, Optional ByVal thisNodeType As Long = 0) As Long

    'Increment our active node count
    m_numOfNodes = m_numOfNodes + 1
    
    'Start by making sure our node directory is large enough to hold this new node (if the caller supplied a node count up front,
    ' this step is overkill).  If the directory is too small, enlarge it.
    If (UBound(m_NodeDirectory) < (m_numOfNodes - 1)) Then
        
        'If the array has never been resized before, give it a nice starting size of 16 entries
        If (UBound(m_NodeDirectory) = 0) Then
            ReDim Preserve m_NodeDirectory(0 To 15) As PDP_NODE
        
        'If the directory has been resized before, double its size now.  (The directory will automatically be shrunk to minimum
        ' size when it's written out to file, so don't worry about excess space being allocated.)
        Else
            ReDim Preserve m_NodeDirectory(0 To UBound(m_NodeDirectory) * 2 + 1) As PDP_NODE
        End If
        
    End If
    
    'Copy the supplied values into the node directory.  Note that all three ID types are optional, but hopefully the user has
    ' been smart enough to make use of at least one of them!
    With m_NodeDirectory(m_numOfNodes - 1)
        .nodeName = thisNodeName
        .NodeID = thisNodeID
        .OptionalNodeType = thisNodeType
    End With
    
    'Return the index of this node, which the caller will use to supply this node's data (in a separate step).
    AddNode = m_numOfNodes - 1
    
End Function

'Add data for a given node.  Four required params, three optional params.  The function will return TRUE if successful.
'
' The four required params are:
' 1) Index of the target node.  This is the value returned by AddNode(), above.
' 2) Destination chunk for this data.  Remember that each node in a pdPackage supports TWO possible data arrays, which can be
'    utilized however the caller pleases.  (Typically, one is used for lightweight header data, while the other is used for
'    heavy content data.)
' 3, 4) Pointer and size of the byte array containing the data the caller wants written.  This class doesn't care how the
'    pointer is obtained or assembled, so long as the SIZE IS BYTE-ACCURATE.
'
' The three optional params are:
' 1) Whether to compress the data before writing it.  If zLib is unavailable, this param has no meaning, as data will always
'    be written without compression.
' 2) Compression level.  The meaning of this varies depending on the compression engine in use.  zLib allows a value from
'    0 (uncompressed) to 9 (best possible compression).  zstd allows a value from 1 (fast, leaner compression) to 22 (slow,
'    very intense compression).  The default value of -1 means "use default compression", and PD will choose an appropriate
'    value for the given library.
Friend Function AddNodeDataFromPointer(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByVal dataPointer As Long, ByVal dataLength As Long, Optional ByVal compressData As PD_CompressionEngine = PD_CE_Zstd, Optional ByVal compressionLevel As Long = -1) As Boolean

    'Start by validating the node index we were passed.  If it's out of range, exit immediately.
    If (nodeIndex < 0) Or (nodeIndex > m_numOfNodes - 1) Then
        InternalErrorMsg "Node index out of range - try again with a valid node index!"
        AddNodeDataFromPointer = False
        Exit Function
    End If
    
    'If compression is requested, make sure a matching compression engine is available
    If ((compressData = PD_CE_ZLib) And (Not Compression.IsCompressionEngineAvailable(PD_CE_ZLib))) Then compressData = PD_CE_NoCompression
    If ((compressData = PD_CE_Zstd) And (Not Compression.IsCompressionEngineAvailable(PD_CE_Zstd))) Then compressData = PD_CE_NoCompression
    If (((compressData = PD_CE_Lz4) Or (compressData = PD_CE_Lz4HC)) And (Not Compression.IsCompressionEngineAvailable(PD_CE_Lz4))) Then compressData = PD_CE_NoCompression
    
    'If compression is requested, update the "zLib required" flag for both the file as a whole, and this specific node.
    If (compressData <> PD_CE_NoCompression) Then SetCompressionFlag nodeIndex, useHeaderBuffer, compressData, True, True
    
    'Update the pre-compression values for this data chunk.
    If useHeaderBuffer Then
        m_NodeDirectory(nodeIndex).NodeHeaderOriginalSize = dataLength
    Else
        m_NodeDirectory(nodeIndex).NodeDataOriginalSize = dataLength
    End If
    
    'Mark the offset for this data chunk.
    If useHeaderBuffer Then
        m_NodeDirectory(nodeIndex).NodeHeaderOffset = m_DataBuffer.GetPosition
    Else
        m_NodeDirectory(nodeIndex).NodeDataOffset = m_DataBuffer.GetPosition
    End If
    
    'If we are not using compression, we can write the node data as-is.  Note that header or data doesn't matter here,
    ' because the data is simply added to the pdPackages data chunk.  Order of writing is irrelevant.
    If (compressData = PD_CE_NoCompression) Then
    
        m_DataBuffer.WriteBytesFromPointer dataPointer, dataLength
        
        If useHeaderBuffer Then
            m_NodeDirectory(nodeIndex).NodeHeaderPackedSize = dataLength
        Else
            m_NodeDirectory(nodeIndex).NodeDataPackedSize = dataLength
        End If
        
    'Data compression was requested, meaning we have a bit of extra work to do.
    Else
        
        Dim compressedSize As Long, compressionIsAnImprovement As Boolean
        
        'Because this code is time-sensitive in PD, I like to track and report performance
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime
        
        'copyPtrToArray will return TRUE if the data was compressed; FALSE if it was not.  There is no error state for the function.
        compressionIsAnImprovement = CopyPtrToArray(dataPointer, dataLength, m_CompressionBuffer, compressedSize, compressData, compressionLevel, False)
        
        'Perform a failsafe check to ensure that the compressed size is smaller than the uncompressed size
        If compressionIsAnImprovement Then compressionIsAnImprovement = (compressedSize < dataLength)
        
        'If compression was successful (and the end result is smaller than the original), point the source data pointer
        ' at the compressed data.
        Dim srcDataPointer As Long
        If compressionIsAnImprovement Then
            srcDataPointer = VarPtr(m_CompressionBuffer(0))
            'Debug.Print "Node data compressed in " & VBHacks.GetTimeDiffNowAsString(startTime) & "; size reduction is " & Format$(100# - (100# * (CDbl(compressedSize) / CDbl(dataLength))), "#0.0") & "% smaller (" & dataLength & " to " & compressedSize & " bytes, via " & Compression.GetCompressorName(compressData) & ")."
        
        'If compression failed - or if it resulted in a larger data size - disregard the compression results, and instead
        ' write out an uncompressed data chunk.
        Else
            
            'Debug.Print "Node data showed no benefit from compression; writing uncompressed chunk instead (" & dataPointer & ", " & dataLength & ", " & compressedSize & ")"
            SetCompressionFlag nodeIndex, useHeaderBuffer, compressData, False, False
            compressedSize = dataLength
            
            'Set the source data pointer to the original pointer, *not* the compressed data
            srcDataPointer = dataPointer
            
        End If
        
        'Copy the appropriate data source (compressed or original data, based on the above checks) into our primary
        ' buffer object, then update the node size flags to match
        m_DataBuffer.WriteBytesFromPointer srcDataPointer, compressedSize
        
        If useHeaderBuffer Then
            m_NodeDirectory(nodeIndex).NodeHeaderPackedSize = compressedSize
        Else
            m_NodeDirectory(nodeIndex).NodeDataPackedSize = compressedSize
        End If
        
    End If
    
    'This chunk was added successfully!  Return TRUE and exit.
    AddNodeDataFromPointer = True
    
End Function

Private Sub SetCompressionFlag(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByVal compressData As PD_CompressionEngine, ByVal newState As Boolean, Optional ByVal updateFileHeaderToo As Boolean = True)

    If (compressData = PD_CE_Zstd) Then
        If updateFileHeaderToo Then SetBitFlag PDP_HF2_ZstdRequired, newState, m_FileHeader.DataFlags(0)
        If useHeaderBuffer Then
            SetBitFlag PDP_NF2_ZstdRequiredH, newState, m_NodeDirectory(nodeIndex).NodeFlags(0)
        Else
            SetBitFlag PDP_NF2_ZstdRequired, newState, m_NodeDirectory(nodeIndex).NodeFlags(0)
        End If
    ElseIf (compressData = PD_CE_ZLib) Then
        If updateFileHeaderToo Then SetBitFlag PDP_HF2_ZlibRequired, newState, m_FileHeader.DataFlags(0)
        If useHeaderBuffer Then
            SetBitFlag PDP_NF2_ZlibRequiredH, newState, m_NodeDirectory(nodeIndex).NodeFlags(0)
        Else
            SetBitFlag PDP_NF2_ZlibRequired, newState, m_NodeDirectory(nodeIndex).NodeFlags(0)
        End If
    ElseIf ((compressData = PD_CE_Lz4) Or (compressData = PD_CE_Lz4HC)) Then
        If updateFileHeaderToo Then SetBitFlag PDP_HF2_Lz4Required, newState, m_FileHeader.DataFlags(0)
        If useHeaderBuffer Then
            SetBitFlag PDP_NF2_Lz4RequiredH, newState, m_NodeDirectory(nodeIndex).NodeFlags(0)
        Else
            SetBitFlag PDP_NF2_Lz4Required, newState, m_NodeDirectory(nodeIndex).NodeFlags(0)
        End If
    End If
    
End Sub

'Thin wrapper to addNodeDataFromPointer, above; if possible, use that function directly to avoid any unnecessary copying of arrays
Friend Function AddNodeDataFromByteArray(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByRef DataBytes() As Byte, Optional ByVal compressData As PD_CompressionEngine = PD_CE_Zstd, Optional ByVal compressionLevel As Long = -1) As Boolean
    AddNodeDataFromByteArray = AddNodeDataFromPointer(nodeIndex, useHeaderBuffer, VarPtr(DataBytes(0)), UBound(DataBytes) + 1, compressData, compressionLevel)
End Function

'Thin wrapper for addNodeDataFromPointer, above, but allows the user to supply a string.
Friend Function AddNodeDataFromString(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByRef srcDataString As String, Optional ByVal compressData As PD_CompressionEngine = PD_CE_Zstd, Optional ByVal compressionLevel As Long = -1) As Boolean
    If (LenB(srcDataString) <> 0) Then
        AddNodeDataFromString = AddNodeDataFromPointer(nodeIndex, useHeaderBuffer, StrPtr(srcDataString), Len(srcDataString) * 2, compressData, compressionLevel)
    End If
End Function

'Shortcut function to both create a new node, and copy a file's contents directly into that node, without any intermediate work on the caller's part.
' Note that only the filename and the file's contents are added; things like file attributes or security descriptors are *not*, by design.
'
'At present, default compression settings are used for the added files (e.g. zstd compression is always applied at the default setting).
'
'This function now supports relative paths, e.g. "\Subfolder\Filename.txt".  To make use of this feature, you *must* pass the desired
' RELATIVE PATH AND FILENAME as the optional "storeUsingThisRelativePathAndFilename" parameter.  (e.g. continuing with the above example,
' you would pass the full absolute path as pathToFile, "C:\User\Subfolder\Filename.txt", then pass "\Subfolder\Filename.txt" as
' storeUsingThisRelativePathAndFilename.)  At extraction time, pdPackage will create any required subfolders for you as part of the extraction process.
'
'If you do make use of the relative path feature, note that you *MUST SUPPLY THE FULL RELATIVE PATH* at extraction time, if extracting nodes
' individually.  This is required as a pdPackage may contain something like "\Subfolder1\Filename.txt" and "\Subfolder2\Filename.txt", and if
' extracting by filename, it can't differentiate between these two unless you provide the relative path too.
'
'Returns: TRUE if successful.
Friend Function AutoAddNodeFromFile(ByVal pathToFile As String, Optional ByVal OptionalNodeType As Long = 0, Optional ByVal storeUsingThisRelativePathAndFilename As String = vbNullString) As Boolean

    'Attempt to load the file into a byte array
    Dim fileContents() As Byte
    If m_File.FileLoadAsByteArray(pathToFile, fileContents) Then
        
        'Make sure the file loaded successfully
        If UBound(fileContents) >= LBound(fileContents) Then
            
            'Create a blank node
            Dim nodeIndex As Long
            
            'Files with relative paths in their names are specially flagged, because we need to reconstruct their folder hierarchy
            ' when extracting them.
            Dim filenameIncludesRelativePath As Boolean
            filenameIncludesRelativePath = (LenB(storeUsingThisRelativePathAndFilename) > 0)
            
            If filenameIncludesRelativePath Then
                nodeIndex = AddNode(PDP_DATA_FILEANDPATH, , OptionalNodeType)
            Else
                nodeIndex = AddNode(PDP_DATA_FILE, , OptionalNodeType)
            End If
            
            'Write the filename out to the first node.  If the caller specifically specified that the filename includes a relative path,
            ' we will write the whole filename, untouched
            If filenameIncludesRelativePath Then
                AutoAddNodeFromFile = AddNodeDataFromString(nodeIndex, True, storeUsingThisRelativePathAndFilename, False)
            Else
                AutoAddNodeFromFile = AddNodeDataFromString(nodeIndex, True, m_File.FileGetName(pathToFile), False)
            End If
            
            'Write the file contents out to the second node
            AutoAddNodeFromFile = AutoAddNodeFromFile And AddNodeDataFromByteArray(nodeIndex, False, fileContents, PD_CE_Zstd, Compression.GetMaxCompressionLevel(PD_CE_Zstd))
            
            'Mark a special flag bit in this node, to identify it as an auto-added file entry.  This simplifies the extraction process later,
            ' and also provides a failsafe against the user adding their own nodes called "PDPFILE".
            SetBitFlag PDP_HF2_FileCollection, True, m_FileHeader.DataFlags(0)
            SetBitFlag PDP_NF2_FileNode, True, m_NodeDirectory(nodeIndex).NodeFlags(0)
            
        End If
    
    Else
        InternalErrorMsg "WARNING! pdPackage.AutoAddNodeFromFile could not load " & pathToFile & ". Abandoning request."
        AutoAddNodeFromFile = False
    End If

End Function

'Shortcut function to create an arbitrary number of nodes, using some source folder as the reference.
'
'This function is basically a thin wrapper to pdFSO.RetrieveAllFiles and AutoAddNodeFromFile.  You could achieve the same results by using
' those two functions yourselves.
'
'The required srcFolder must be an absolute path, including drive letter (e.g. "C:\ThisFolder\SubFolder").
'
'For convenience, every file added via this request can be assigned an optional node type.  This same node type can be used when extracting
' files, if you want to add many folder groups to a single pdPackage instance, but still retain the ability to access them individually
' at extraction time.
'
'Subfolder recursion is assumed, but it can be switched off via the optional recurseSubfolders parameter.
'
'When subfolder recursion is in use, each file automatically has a relative path automatically generated for it.  By default, this path is
' constructed relative to the srcFolder parameter.  If this behavior is not desirable, an alternate base folder can be specified via the
' useCustomRelativeFolderBase string.  This is preferable if you are adding multiple folders via this call, but you want all added files
' extracted relative to some parent folder.  For example, when updating PD, I manually add the "PhotoDemon\App" subfolder to the update package.
' By default, all added files will be constructed relative to the \App subfolder, so "C:\PhotoDemon\App\Languages\Deutsch.xml" becomes
' "\Languages\Deutsch.xml".  At extraction time, this file would be erroneously extracted to "C:\New PD Folder\Languages\Deutsch.xml", because
' the \App folder was the original relative base.  To override this behavior, I manually specify a different base folder at creation time,
' e.g. "C:\PhotoDemon\".  This causes all added files to receive the relative path "App\Languages\Deutsch.xml", making extraction much simpler.
'
'If you do not want relative paths used whatsoever, set the optional preserveFolders parameter to FALSE.  If you do this, note that you cannot
' change your mind at extraction time, as relative paths will not physically exist inside the package.
'
'Default compression settings are assumed for all added files.
'
'Finally, whitelisting or blacklisting behavior is available via the optional onlyAllowTheseExtensions and doNotAllowTheseExtensions parameters.
' These two parameters are passed, untouched, to an underlying pdFSO instance, so please refer to pdFSO for details on how to use these.
'
'Returns: TRUE if all files are added successfully, FALSE otherwise.  FALSE should not occur unless read access to target folder(s) is restricted,
' or if memory runs out because the constructed package exceeds available memory.
Friend Function AutoAddNodesFromFolder(ByRef srcFolder As String, Optional ByVal OptionalNodeType As Long = 0, Optional ByVal recurseSubfolders As Boolean = True, Optional ByVal preserveFolders As Boolean = True, Optional ByVal useCustomRelativeFolderBase As String = vbNullString, Optional ByVal onlyAllowTheseExtensions As String = vbNullString, Optional ByVal doNotAllowTheseExtensions As String = vbNullString) As Boolean

    'As usual, our lovely pdFSO instance (m_File) is going to handle most the heavy lifting.  It will return the list of files to be added
    ' in a pdStringStack object, which we can then iterate at our leisure.
    Dim filesToAdd As pdStringStack
    Set filesToAdd = New pdStringStack
    
    'Retrieve all files in the specified folder
    If m_File.RetrieveAllFiles(srcFolder, filesToAdd, recurseSubfolders, False, onlyAllowTheseExtensions, doNotAllowTheseExtensions) Then
        
        Dim nodeAdditionSuccess As Boolean
        nodeAdditionSuccess = True
        
        'filesToAdd now contains absolute paths to all files we need to add.  Pop each file in turn, adding it as we go.
        Dim curFile As String, relativeFilePath As String
        
        Do While filesToAdd.PopString(curFile)
            
            'Start by constructing a relative path for this string.  (pdFSO could have done this for us automatically, but we need the full
            ' paths for node addition, and relative paths may vary if the user supplies custom path options.)
            '
            'First, check for folder preservation.
            If preserveFolders Then
                
                'Folder preservation is active.  Find the appropriate relative base, per the caller's supplied path(s).
                If (LenB(useCustomRelativeFolderBase) <> 0) Then
                    relativeFilePath = m_File.GenerateRelativePath(useCustomRelativeFolderBase, curFile)
                Else
                    relativeFilePath = m_File.GenerateRelativePath(srcFolder, curFile)
                End If
                
            'Folder preservation is inactive, so do not specify any relative paths whatsoever
            Else
                relativeFilePath = vbNullString
            End If
            
            'Add this node
            nodeAdditionSuccess = nodeAdditionSuccess And AutoAddNodeFromFile(curFile, OptionalNodeType, relativeFilePath)
            
        Loop
        
        AutoAddNodesFromFolder = nodeAdditionSuccess
        
        'Display some debug info if one or more nodes failed to add
        If (Not nodeAdditionSuccess) Then
            If recurseSubfolders Then
                InternalErrorMsg "WARNING! pdPackager.AutoAddNodesFromFolder() failed to add one or more nodes.  Do you have read access to all subfolders??"
            Else
                InternalErrorMsg "WARNING! pdPackager.AutoAddNodesFromFolder() failed to add one or more nodes.  Please investigate."
            End If
        End If
        
    Else
        InternalErrorMsg "WARNING! pdPackager.AutoAddNodesFromFolder() could not retrieve a valid list of files.  Do you have read access to the folder??"
        AutoAddNodesFromFolder = False
    End If

End Function

'Shortcut function to create an arbitrary number of nodes, using a pdStringStack full of files and/or folders as the reference.
' (Obviously, it's up to the caller to populate the string stack.)
'
'This function is basically a thin wrapper to AutoAddNodeFromFile and AutoAddNodesFromFolder.  You could achieve identical results by using
' those two functions yourselves.
'
'All files and folders in the in the pdStringStack object must be absolute paths, including drive letter (e.g. "C:\ThisFolder\SubFolder" or
' "C:\ThisFolder\SubFolder\file.txt")
'
'For convenience, every file and/or folder added via this request can be assigned an optional node type.  This same node type can be used
' when extracting files, if you want to add multiple file and/or folder groups to a single pdPackage instance, but still retain the ability
' to access them individually at extraction time.
'
'Subfolder recursion is assumed for any folders in the stack, but this behavior can be switched off via the optional recurseSubfolders parameter.
'
'Because the incoming pdStringStack can potentially contain a huge range of possible files and folders, this function is unique in forcing you
' to specify your own custom base folder (relativeFolderBase) against which relative paths will be contructed.  If this value is *not* supplied,
' relative paths will not be constructed - e.g. all discovered files will be added WITHOUT folder information attached - so plan accordingly.
'
'Whitelist and blacklist options are disabled for performance reasons.  It is assumed that the caller already made use of these, if desired,
' when constructing the incoming string stack.
'
'Default compression settings are assumed for all added files.
'
'Returns: TRUE if all files are added successfully, FALSE otherwise.  FALSE should not occur unless read access to target file(s) and/or folder(s)
' is restricted, or if memory runs out because the constructed package exceeds available memory.
Friend Function AutoAddNodesFromStringStack(ByRef srcStringStack As pdStringStack, Optional ByVal relativeFolderBase As String = vbNullString, Optional ByVal OptionalNodeType As Long = 0, Optional ByVal recurseSubfolders As Boolean = True) As Boolean
    
    Dim autoAddSuccess As Boolean
    autoAddSuccess = True
    
    'Prior to starting, mark whether relative folders are in use.  If they are not, we can bypass some steps within the main Do loop.
    Dim relativeFoldersInUse As Boolean
    relativeFoldersInUse = (LenB(relativeFolderBase) <> 0)
    
    Dim relativePath As String
    
    'This function relies on autoAddNodesFromFolder and autoAddNodeFromFile to do all the heavy lifting.  All we do is pop items off the
    ' string stack, and forward them to one of those two functions.
    Dim curFile As String
    
    Do While srcStringStack.PopString(curFile)
    
        'See if this is a file or a folder
        If m_File.FileExists(curFile) Then
        
            'This is a file.  Add it via the file function
            If relativeFoldersInUse Then
            
                'Construct a relative path
                relativePath = m_File.GenerateRelativePath(relativeFolderBase, curFile)
                Debug.Print "Compressing " & curFile & "..."
                autoAddSuccess = autoAddSuccess And AutoAddNodeFromFile(curFile, OptionalNodeType, relativePath)
                
            Else
                autoAddSuccess = autoAddSuccess And AutoAddNodeFromFile(curFile, OptionalNodeType)
            End If
        
        ElseIf m_File.PathExists(curFile, False) Then
        
            'This is a folder.  Add it via the folder function.
            If relativeFoldersInUse Then
                autoAddSuccess = autoAddSuccess And AutoAddNodesFromFolder(curFile, OptionalNodeType, recurseSubfolders, True, relativeFolderBase)
            Else
                autoAddSuccess = autoAddSuccess And AutoAddNodesFromFolder(curFile, OptionalNodeType, recurseSubfolders, False)
            End If
        
        Else
            InternalErrorMsg "WARNING!  pdPackager.AutoAddNodesFromStringStack() was handed something that isn't a file or folder (" & curFile & ")."
        End If
    
    Loop
    
    AutoAddNodesFromStringStack = autoAddSuccess
    
End Function

'When all nodes have been successfully added, the user can finally write out their data to file.  This function will return TRUE
' if successful, FALSE if unsuccessful.  (Note that it assumes the caller has done some basic validation on the file path, like
' obtaining permission from the user to overwrite an existing file.)
'
'The only required parameter is the destination filename.
'
'Two optional parameters exist for this function:
' - secondPassDirectoryCompression: compress the directory chunk before writing it out to file.
' - thisIsATempFile: special WAPI flags related to temp files will be set.  Note that this flag does not affect file correctness in
'                    any way; it simply specifies hints related to file caching.
Friend Function WritePackageToFile(ByVal dstFilename As String, Optional ByVal secondPassDirectoryCompression As PD_CompressionEngine = PD_CE_NoCompression, Optional ByVal thisIsATempFile As Boolean = False, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    'If memory is tight, you may choose to free any temporary compression buffers here, as they are no longer needed.
    ' (In PD, we keep them around as we may reuse a pdPackage instance for multiple writes.)
    'Erase m_CompressionBuffer
    
    'Start by updating the file header.  Most of this will have been done when the class was initialized, but some values
    ' can't be set until all nodes have been added.
    With m_FileHeader
    
        'Update the final node count
        .NodeCount = m_numOfNodes
                
        'Update the size of the directory chunk
        .DirectoryChunkSizeOriginal = .NodeStructSize * m_numOfNodes
        
        'We could trim the data buffer here, but it hurts performance and doesn't gain us anything - so instead, just read the
        ' current stream size and write it into the file header.
        .DataChunkSizeOriginal = m_DataBuffer.GetStreamSize
        
    End With
    
    'Before writing the file, we have a few other potential header items we need to construct.
    
    'Create the node directory byte array now.  (VB writes UDTs to file in a non-standard, non-obvious way, so we need to do this
    ' regardless of second-pass directory compression.)
    Dim rawDirectoryBuffer() As Byte
    Dim originalSize As Long, compressedSize As Long
    originalSize = m_FileHeader.DirectoryChunkSizeOriginal
    
    'Use the copyPtrToArray function to copy the contents of m_NodeDirectory into rawDirectoryBuffer().  The function will return TRUE
    ' if it compresses the data; false, otherwise.
    If CopyPtrToArray(VarPtr(m_NodeDirectory(0)), originalSize, rawDirectoryBuffer, compressedSize, secondPassDirectoryCompression, compressionLevel, True) Then
        m_FileHeader.DirectoryChunkSizePacked = compressedSize
        
        If ((secondPassDirectoryCompression = PD_CE_Zstd) And (originalSize <> compressedSize)) Then
            SetBitFlag PDP_HF2_ZstdRequired, True, m_FileHeader.DirectoryFlags(0)
        ElseIf ((secondPassDirectoryCompression = PD_CE_ZLib) And (originalSize <> compressedSize)) Then
            SetBitFlag PDP_HF2_ZlibRequired, True, m_FileHeader.DirectoryFlags(0)
        ElseIf (((secondPassDirectoryCompression = PD_CE_Lz4) Or (secondPassDirectoryCompression = PD_CE_Lz4HC)) And (originalSize <> compressedSize)) Then
            SetBitFlag PDP_HF2_Lz4Required, True, m_FileHeader.DirectoryFlags(0)
        End If
        
    Else
        m_FileHeader.DirectoryChunkSizePacked = originalSize
    End If
    
    'Now, we would typically repeat the above steps for the data chunk.  HOWEVER: the data chunk does not current support
    ' second-pass compression (where the entire data chunk is compressed AGAIN, as-is), so we know its packed size is identical
    ' to its original size.
    m_FileHeader.DataChunkSizePacked = m_FileHeader.DataChunkSizeOriginal
    
    'With all that data known, we can now fill-in the correct offsets inside the file header
    m_FileHeader.DataChunkPosition = LenB(m_FileHeader)
    m_FileHeader.DirectoryChunkPosition = m_FileHeader.DataChunkPosition + m_FileHeader.DataChunkSizePacked
    
    'If this is a memory-backed stream, kill the destination file if it already exists (because we will be writing the entire
    ' thing from scratch).
    If (m_StreamMode = PD_SM_MemoryBacked) Then
        If m_File.FileExists(dstFilename) Then m_File.FileDelete dstFilename
    
    'If this is a file-backed stream, the node data has already been written to file.  We just need to add the file header
    ' and the dictionary.  Close the stream.
    Else
        m_DataBuffer.StopStream
    End If
    
    'Retrieve a writable handle
    Dim hFile As Long, createSuccess As Boolean
    
    If (m_StreamMode = PD_SM_MemoryBacked) Then
        If thisIsATempFile Then
            createSuccess = m_File.FileCreateHandle(dstFilename, hFile, True, True, OptimizeTempFile Or OptimizeSequentialAccess)
        Else
            createSuccess = m_File.FileCreateHandle(dstFilename, hFile, True, True, OptimizeSequentialAccess)
        End If
    ElseIf (m_StreamMode = PD_SM_FileBacked) Then
        createSuccess = m_File.FileCreateAppendHandle(dstFilename, hFile)
    End If
    
    If createSuccess Then
    
        Dim mmapHandle As Long, mmapPtr As Long, totalWriteSize As Long, mmapOffset As Long
        
        'Once again, split handling based on the stream type.  For a memory-backed stream, we have to write out the entire
        ' data chunk, but for a file-backed stream, we have much less data to write.
        If (m_StreamMode = PD_SM_MemoryBacked) Then
        
            'First, attempt to write out the package using a memory-mapped file.  This can be faster (comparable to
            ' an asynchronous write) vs raw WriteData calls.
            totalWriteSize = LenB(m_FileHeader) + m_FileHeader.DataChunkSizePacked + m_FileHeader.DirectoryChunkSizePacked
            
            Dim useMemMapMode As Boolean
            useMemMapMode = m_File.FileConvertHandleToMMPtr(hFile, mmapHandle, mmapPtr, totalWriteSize)
            
            If useMemMapMode Then
    
                'Write out the file data using good ol' CopyMemoryStrict!
                
                mmapOffset = LenB(m_FileHeader)
                
                'File header...
                CopyMemoryStrict mmapPtr, VarPtr(m_FileHeader), mmapOffset
                
                'Data buffer...
                CopyMemoryStrict mmapPtr + mmapOffset, m_DataBuffer.Peek_PointerOnly(0), m_FileHeader.DataChunkSizePacked
                
                'Directory at the end...
                CopyMemoryStrict mmapPtr + mmapOffset + m_FileHeader.DataChunkSizePacked, VarPtr(rawDirectoryBuffer(0)), m_FileHeader.DirectoryChunkSizePacked
                
                'Release the mapped view and allow the system to write the file at its leisure
                m_File.ReleaseMMHandle mmapHandle, mmapPtr
    
            'If we failed to allocate sufficient memory for a memory-mapped write, fall back to default write techniques
            Else
        
                'Writing out the three crucial pieces of data: the header, the data, the directory
                With m_File
                    .FileWriteData hFile, VarPtr(m_FileHeader), LenB(m_FileHeader)
                    .FileWriteData hFile, m_DataBuffer.Peek_PointerOnly(0), m_FileHeader.DataChunkSizePacked
                    .FileWriteData hFile, VarPtr(rawDirectoryBuffer(0)), m_FileHeader.DirectoryChunkSizePacked
                End With
                
            End If
            
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
        
            'Write the header at the start of the file
            With m_File
                .FileMovePointer hFile, 0, FILE_BEGIN
                .FileWriteData hFile, VarPtr(m_FileHeader), LenB(m_FileHeader)
                
                'Skip to the end of the file, then write the directory
                .FileMovePointer hFile, LenB(m_FileHeader) + m_FileHeader.DataChunkSizePacked, FILE_BEGIN
                .FileWriteData hFile, VarPtr(rawDirectoryBuffer(0)), m_FileHeader.DirectoryChunkSizePacked
            End With
        
        End If
        
        m_File.FileCloseHandle hFile
        
        WritePackageToFile = True
    
    Else
        InternalErrorMsg "WARNING!  pdPackager.WritePackageToFile failed to create a valid destination file handle.  Package was not written."
        WritePackageToFile = False
    End If
        
    'TODO!  Add basic error handling.

End Function

'Copy one array into another, with variable compression.  This is used in many places throughout this class; basically any action that
' potentially supports compression should use this function to transfer data between arrays, even when compression is not in use.
' (This function abstracts away all the messy CopyMemoryStrict bits, keeping the rest of the class clean.)
'
'If compression *is* in use, this function will use the requested engine to compress the array accordingly, and it will fill the passed
' dstSize Long with the size (1-based) of the compressed array.  Explicit size values are required for both source and destination arrays,
' which allows both the caller and this function to avoid the need for precisely dimensioned arrays.  (Any time we can skip an unnecessary
' ReDim Preserve, we save precious processing time.)
'
'This function is designed to be error-proof.  If something goes wrong during the compression stage, it will do a straight copy from
' the source to the destination.  Similarly, if the compressed size is larger than the uncompressed size, it will also return an
' uncompressed copy of the data.  The returned BOOLEAN reports whether or not compression was actually used, NOT whether or not the
' function was successful.  (It is assumed the function is always successful, because if compression fails, the default uncompressed
' array will still be returned, and the caller can proceed normally with that data.)
Private Function CopyPtrToArray(ByVal srcPointer As Long, ByVal srcSize As Long, ByRef dstArray() As Byte, ByRef dstSize As Long, Optional ByVal useCompression As PD_CompressionEngine = PD_CE_NoCompression, Optional ByVal compressionLevel As Long = -1, Optional ByVal trimDestinationArray As Boolean = False) As Boolean
    CopyPtrToArray = Compression.CompressPtrToDstArray(dstArray, dstSize, srcPointer, srcSize, useCompression, compressionLevel, , trimDestinationArray)
End Function

'Close the current package.  Any used resources will be immediately freed, so make sure you're finished with the package data
' (e.g. by writing it out to file or something) before closing it!
Friend Sub ClosePackage(Optional ByVal retainInternalBuffers As Boolean = False)
    If (Not retainInternalBuffers) Then Erase m_CompressionBuffer
    If (Not m_DataBuffer Is Nothing) Then m_DataBuffer.StopStream (m_StreamMode <> PD_SM_MemoryBacked) Or (Not retainInternalBuffers)
End Sub

'Load a pdPackage file into memory.
'
'NOTE: the caller is expected to handle detailed testing of the supplied path (e.g. read access, etc).  This function performs
'    minimal error checking.
'
'If the file can be successfully loaded and parsed, this function returns TRUE.
'
'Parameters include:
' 1) Source file.  (Again, callers must do their own validation on this path.)
' 2) Optionally, a sub-type value you want to validate.  If this value is not found in bytes 4-7 of the file header, the file
'     will be rejected.  (If you didn't request a specific sub-type at pdPackage creation time, don't enable this check!)
' 3) Mode for the pdStream backer.  MEMORY mode means the entire file will be immediately read into memory and stored there.
'     For small files, this is fine, and will likely result in better node retrieval performance.  For large files, however,
'     you probably just want the stream to be created in FILE mode.  This will only pull data from the file as it is requested,
'     and you *must* keep the file alive and accessible for the duration of this pdPackage instance.
' 4) Access rights for the pdStream backer.  In READ-ONLY mode, you will not be able to append anything to this pdPackage file.
'     However, memory-mapping will be used to load nodes into file, resulting in even better performance.  This is the default
'     mode, and at present, the only supported mode.
Friend Function ReadPackageFromFile(ByVal srcFilename As String, Optional ByVal subTypeValidator As Long = 0, Optional ByVal streamMode As PD_STREAM_MODE = PD_SM_MemoryBacked, Optional ByVal streamAccess As PD_STREAM_ACCESS = PD_SA_ReadOnly, Optional ByVal optimizeAccess As PD_FILE_ACCESS_OPTIMIZE = OptimizeRandomAccess) As Boolean

    On Error GoTo StopPackageFileRead
    
    'If second-pass compression was used (e.g. the entire file was compressed *again* after all nodes were added), we need two levels
    ' of temporary arrays to handle the decompression process.
    Dim tmpCompressionBuffer() As Byte
    
    'Before doing anything else, make sure the file exists
    If (Not m_File.FileExists(srcFilename)) Then
        InternalErrorMsg "Requested pdPackage file doesn't exist.  Validate all paths before sending them to the pdPackage class!"
        GoTo StopPackageFileRead
    End If
    
    'Open the file.  Note that we play some games with the "optimizeAccess" parameter here.  If the caller wants this file
    ' to be memory-backed, optimizeAccess really doesn't matter (because we're just gonna copy the whole file into memory).
    ' If, however, the caller wants file-backed mode (where the package is left on-disk, and only accessed as nodes are retrieved),
    ' then we're gonna silently modify the optimizeAccess to start as Random Access (because we need to jump to the end of the
    ' file to grab the directory), but after the initial file handle is closed, we'll respect the caller's request inside the
    ' pdStream object (which grabs the actual nodes).
    Dim initialOptimizeAccess As PD_FILE_ACCESS_OPTIMIZE
    If (streamMode = PD_SM_MemoryBacked) Then
        initialOptimizeAccess = optimizeAccess
    ElseIf (streamMode = PD_SM_FileBacked) Then
        initialOptimizeAccess = OptimizeRandomAccess
    End If
    
    Dim hFile As Long
    If m_File.FileCreateHandle(srcFilename, hFile, True, False, initialOptimizeAccess) Then
        
        'Attempt to load the file header.  If the file is invalid, this may potentially be gibberish, but we won't know until we try!
        m_File.FileReadData hFile, VarPtr(m_FileHeader), LenB(m_FileHeader)
        
        'Validate the file signature.  Valid pdPackage files must have their first 4 bytes set to "PDPK" in ASCII.
        If (m_FileHeader.PDP_ID <> PDP_UNIVERSAL_IDENTIFIER) Then
            InternalErrorMsg "File doesn't have valid pdPackage header.  Abandoning load."
            GoTo StopPackageFileRead
        End If
        
        'If the caller requested a sub-type validation, perform that now
        If (subTypeValidator <> 0) Then
            If (m_FileHeader.PDP_SubID <> subTypeValidator) Then
                InternalErrorMsg "File doesn't match requested sub-type value.  Abandoning load."
                GoTo StopPackageFileRead
            End If
        End If
        
        'Finally, check the PDP version.  This check is currently meaningless, as this class only supports the current pdPackage version.
        If (m_FileHeader.PDP_Version < THIS_PDPACKAGE_VERSION) Then
            InternalErrorMsg "Legacy file format could not be decoded."
            GoTo StopPackageFileRead
            
        'This is the current PDP version.  Continue the normal loading process.
        Else
            
            'Make sure the struct size is correct; if it isn't, all subsequent parsing will fail.
            Dim tmpNode As PDP_NODE
            If (m_FileHeader.NodeStructSize <> LenB(tmpNode)) Then
                InternalErrorMsg "Node struct size in header is invalid.  This file looks to be corrupt!  Abandoning load."
                GoTo StopPackageFileRead
            End If
            
            'Next, it's time to retrieve the directory.  The size of the directory varies, based on whether it received
            ' second-pass compression, but the uncompressed result will *always* end up inside this m_CompressionBuffer() array/
            If (Not VBHacks.IsArrayInitialized(m_CompressionBuffer)) Then
                ReDim m_CompressionBuffer(0 To m_FileHeader.DirectoryChunkSizeOriginal - 1) As Byte
            Else
                If (UBound(m_CompressionBuffer) < m_FileHeader.DirectoryChunkSizeOriginal - 1) Then ReDim m_CompressionBuffer(0 To m_FileHeader.DirectoryChunkSizeOriginal - 1) As Byte
            End If
            
            'If no compression was applied, read the bytes directly into our buffer
            If (m_FileHeader.DirectoryChunkSizePacked = m_FileHeader.DirectoryChunkSizeOriginal) Then
                m_File.FileMovePointer hFile, m_FileHeader.DirectoryChunkPosition, FILE_BEGIN
                m_File.FileReadData hFile, VarPtr(m_CompressionBuffer(0)), m_FileHeader.DirectoryChunkSizePacked
            
            'If the packed and original sizes do *not* match, the directory is compressed.  Attempt to decompress it now.
            Else
                
                'Regardless of compression method, pull the raw bytes from file
                ReDim tmpCompressionBuffer(0 To m_FileHeader.DirectoryChunkSizePacked - 1) As Byte
                m_File.FileReadData hFile, VarPtr(tmpCompressionBuffer(0)), m_FileHeader.DirectoryChunkSizePacked, m_FileHeader.DirectoryChunkPosition, FILE_BEGIN
                
                Dim decompressionResult As Boolean: decompressionResult = False
                
                'Figure out which decompression library was used for the original compression.
                ' (Note that if the required decompression library isn't available, we're screwed; loading is impossible.)
                Dim cmpEngine As PD_CompressionEngine: cmpEngine = PD_CE_NoCompression
                If GetBitFlag(PDP_HF2_ZstdRequired, m_FileHeader.DirectoryFlags(0)) Then
                    cmpEngine = PD_CE_Zstd
                ElseIf GetBitFlag(PDP_HF2_ZlibRequired, m_FileHeader.DirectoryFlags(0)) Then
                    cmpEngine = PD_CE_ZLib
                ElseIf GetBitFlag(PDP_HF2_Lz4Required, m_FileHeader.DirectoryFlags(0)) Then
                    cmpEngine = PD_CE_Lz4
                End If
                
                If Compression.IsCompressionEngineAvailable(cmpEngine) Then
                
                    decompressionResult = Compression.DecompressPtrToPtr(VarPtr(m_CompressionBuffer(0)), m_FileHeader.DirectoryChunkSizeOriginal, VarPtr(tmpCompressionBuffer(0)), m_FileHeader.DirectoryChunkSizePacked, cmpEngine)
                
                    'If the decompression fails, it may be due to a bad compression encoding (e.g. the data is compressed using a system other
                    ' than the one flagged in the file.)  Try again using a different compression method.
                    If (Not decompressionResult) Then
                    
                        InternalErrorMsg "WARNING!  pdPackager.ReadPackageFromFile failed to decode directory; trying again with alternate decompressors..."
                        
                        Dim i As PD_CompressionEngine
                        For i = PD_CE_ZLib To PD_CE_Lz4
                            decompressionResult = Compression.DecompressPtrToPtr(VarPtr(m_CompressionBuffer(0)), m_FileHeader.DirectoryChunkSizeOriginal, VarPtr(tmpCompressionBuffer(0)), m_FileHeader.DirectoryChunkSizePacked, i)
                            If decompressionResult Then
                                InternalErrorMsg "Directory may be recoverable; compressor #" & i & " worked.  Proceeding with load..."
                                Exit For
                            End If
                        Next i
                        
                        'If all decompression attempts failed, we're shit out of luck
                        If (Not decompressionResult) Then
                            InternalErrorMsg "Sorry, but no decompressor worked.  Something's definitely wrong with this package."
                        End If
                        
                    End If
                    
                Else
                    InternalErrorMsg "WARNING!  Required decompression engine missing; pdPackage directory cannot be read."
                End If
                
                'If decompression failed, we have to abandon the load process, as we can't even access the original file's directory.
                If (Not decompressionResult) Then
                    InternalErrorMsg "WARNING!  Could not decompress pdPackage directory.  Possible explanations include missing decompression library or file corruption."
                    GoTo StopPackageFileRead
                End If
                    
            End If
            
            'Free the compressed copy of the directory (if any)
            Erase tmpCompressionBuffer
            
            'Use information from the file header to prepare the directory array, then copy the raw buffer into it.
            Dim newNodeCount As Long
            newNodeCount = m_FileHeader.NodeCount
            
            If (newNodeCount > m_numOfNodes) Then ReDim m_NodeDirectory(0 To newNodeCount - 1) As PDP_NODE
            m_numOfNodes = newNodeCount
            CopyMemoryStrict VarPtr(m_NodeDirectory(0)), VarPtr(m_CompressionBuffer(0)), m_FileHeader.DirectoryChunkSizeOriginal
            
            'If memory is tight, you might choose to free the raw directory buffer here (as it's no longer needed).  In PD,
            ' we skip this step as we may reuse the compression buffer for future read/write operations.
            'Erase m_CompressionBuffer
            
            'The package's directory is now available in-memory.  Before proceeding, make a note of the requested stream mode
            ' and access, which affect how we pull data out of this file.
            m_StreamMode = streamMode
            m_StreamAccess = streamAccess
            
            'Finally, retrieve the data chunk.  Note that we basically just wrap a pdStream object around the data, and it handles
            ' everything from there.
            If (m_DataBuffer Is Nothing) Then Set m_DataBuffer = New pdStream
            If (m_FileHeader.DataChunkSizePacked = m_FileHeader.DataChunkSizeOriginal) Then
                
                'For a memory-backed stream, immediately load the full data chunk into memory.
                If (streamMode = PD_SM_MemoryBacked) Then
                    
                    m_DataBuffer.StartStream PD_SM_MemoryBacked, streamAccess, , , , , True
                    m_DataBuffer.EnsureBufferSpaceAvailable m_FileHeader.DataChunkSizeOriginal, True
                    m_File.FileReadData hFile, m_DataBuffer.Peek_PointerOnly(0), m_FileHeader.DataChunkSizeOriginal, m_FileHeader.DataChunkPosition, FILE_BEGIN
                    
                    'Reset the buffer's internal data pointer (instead of letting it auto-position itself at the end of the buffer.)
                    m_DataBuffer.SetSizeExternally m_FileHeader.DataChunkSizeOriginal
                    m_DataBuffer.SetPosition 0
                    
                    'Close the file handle (because we don't need it anymore - everything's copied into memory!)
                    If (hFile <> 0) Then m_File.FileCloseHandle hFile
                
                'For a file-backed stream, simply wrap the pdStream object around the file, and set a new "zero" point inside
                ' the stream (since we don't want to overwrite the file header, which sits at a fixed size and position).
                ElseIf (streamMode = PD_SM_FileBacked) Then
                
                    'Start by closing our copy of the file handle
                    If (hFile <> 0) Then m_File.FileCloseHandle hFile
                    
                    'Open the file again, but this time, use pdStream to do it.  Note that the stream object will automatically
                    ' subtract the header size from all position requests we pass to it, so we don't have to worry about special
                    ' handling in other parts of this class.
                    m_DataBuffer.StartStream PD_SM_FileBacked, streamAccess, srcFilename, , LenB(m_FileHeader), optimizeAccess
                
                End If
                
            Else
                InternalErrorMsg "Data chunk sizes don't match.  Package cannot be read."
                GoTo StopPackageFileRead
            End If
            
        End If
        
        'File load complete!
        ReadPackageFromFile = True
        
    Else
        ReadPackageFromFile = False
        InternalErrorMsg "WARNING!  pdPackager.readPackageFromFile failed to create a valid handle for " & srcFilename & ".  Read abandoned."
    End If
    
    Exit Function
    
StopPackageFileRead:
    
    InternalErrorMsg "An internal error occurred in the pdPackager.ReadPackageFromFile function."
    If (hFile <> 0) Then m_File.FileCloseHandle hFile
    ReadPackageFromFile = False
    Exit Function

End Function

'Load a pdPackage file from some arbitrary pointer (+length, obviously).  If validated successfully, the source memory will be
' copied locally, meaning you can free the source immediately after this function successfully returns.
'
'If a pdPackage can be successfully parsed from the pointer, this function returns TRUE.
'
'Parameters include:
' 1, 2) Pointer + length.  Length is non-optional, and it *will* be validated against the pdPackage's internal pointers -
'     so make certain it is correct.
' 3) Optionally, a sub-type value you want to validate.  If this value is not found in bytes 4-7 of the file header, the file
'     will be rejected.  (If you didn't request a specific sub-type at pdPackage creation time, don't enable this check!)
' 4) Access rights for the pdStream backer.  READ-ONLY mode is the default mode, and at present, the only supported mode.
Friend Function ReadPackageFromMemory(ByVal srcPointer As Long, ByVal srcLength As Long, Optional ByVal subTypeValidator As Long = 0, Optional ByVal streamAccess As PD_STREAM_ACCESS = PD_SA_ReadOnly) As Boolean

    On Error GoTo StopPackageMemoryRead
    
    Dim rawBuffer() As Byte, tmpCompressionBuffer() As Byte
    
    'Attempt to load a file header.  If the source memory is invalid, this may potentially be gibberish, but we won't know until we try!
    If (srcLength > LenB(m_FileHeader)) Then
        CopyMemoryStrict VarPtr(m_FileHeader), srcPointer, LenB(m_FileHeader)
    Else
        InternalErrorMsg "Bad memory length specified (smaller than pdPackage header).  Abandoning load."
        GoTo StopPackageMemoryRead
    End If
    
    'Validate the file signature.  Valid pdPackage files must have their first 4 bytes set to "PDPK" in ASCII.
    If (m_FileHeader.PDP_ID <> PDP_UNIVERSAL_IDENTIFIER) Then
        InternalErrorMsg "Memory doesn't have valid pdPackage header.  Abandoning load."
        GoTo StopPackageMemoryRead
    End If
        
    'If the caller requested a sub-type validation, perform that now
    If (subTypeValidator <> 0) Then
        If (m_FileHeader.PDP_SubID <> subTypeValidator) Then
            InternalErrorMsg "Memory doesn't match requested sub-type value.  Abandoning load."
            GoTo StopPackageMemoryRead
        End If
    End If
        
    'Finally, check the PDP version.  This check is currently meaningless, as this class only supports the current pdPackage version.
    If (m_FileHeader.PDP_Version < THIS_PDPACKAGE_VERSION) Then
        InternalErrorMsg "Legacy file format could not be decoded."
        GoTo StopPackageMemoryRead
        
    'This is the current PDP version.  Continue the normal loading process.
    Else
        
        'Make sure the struct size is correct; if it isn't, all subsequent parsing will fail.
        Dim tmpNode As PDP_NODE
        If (m_FileHeader.NodeStructSize <> LenB(tmpNode)) Then
            InternalErrorMsg "Node struct size in header is invalid.  This memory looks to be invalid!  Abandoning load."
            GoTo StopPackageMemoryRead
        End If
        
        'Next, it's time to retrieve the directory.  The size of the directory varies, based on whether it received
        ' second-pass compression, but the uncompressed result will *always* end up inside this rawBuffer() array.
        ReDim rawBuffer(0 To m_FileHeader.DirectoryChunkSizeOriginal - 1) As Byte
        
        'Failsafe memory length check.  (Note that the pdPackage directory sits at the *end* of the file, so this length
        ' check is the most important one!)
        If (srcLength >= m_FileHeader.DirectoryChunkPosition + m_FileHeader.DirectoryChunkSizePacked) Then
        
            'If no compression was applied, read the bytes directly into our buffer
            If (m_FileHeader.DirectoryChunkSizePacked = m_FileHeader.DirectoryChunkSizeOriginal) Then
                CopyMemoryStrict VarPtr(rawBuffer(0)), srcPointer + m_FileHeader.DirectoryChunkPosition, m_FileHeader.DirectoryChunkSizePacked
                
            'If the packed and original sizes do *not* match, the directory is compressed.  Attempt to decompress it now.
            Else
                
                'Regardless of compression method, pull the raw bytes from file
                ReDim tmpCompressionBuffer(0 To m_FileHeader.DirectoryChunkSizePacked - 1) As Byte
                CopyMemoryStrict VarPtr(tmpCompressionBuffer(0)), srcPointer + m_FileHeader.DirectoryChunkPosition, m_FileHeader.DirectoryChunkSizePacked
                
                Dim decompressionResult As Boolean: decompressionResult = False
                
                'Figure out which decompression library was used for the original compression.
                ' (Note that if the required decompression library isn't available, we're screwed; loading is impossible.)
                Dim cmpEngine As PD_CompressionEngine: cmpEngine = PD_CE_NoCompression
                If GetBitFlag(PDP_HF2_ZstdRequired, m_FileHeader.DirectoryFlags(0)) Then
                    cmpEngine = PD_CE_Zstd
                ElseIf GetBitFlag(PDP_HF2_ZlibRequired, m_FileHeader.DirectoryFlags(0)) Then
                    cmpEngine = PD_CE_ZLib
                ElseIf GetBitFlag(PDP_HF2_Lz4Required, m_FileHeader.DirectoryFlags(0)) Then
                    cmpEngine = PD_CE_Lz4
                End If
                
                If Compression.IsCompressionEngineAvailable(cmpEngine) Then
                
                    decompressionResult = Compression.DecompressPtrToPtr(VarPtr(rawBuffer(0)), m_FileHeader.DirectoryChunkSizeOriginal, VarPtr(tmpCompressionBuffer(0)), m_FileHeader.DirectoryChunkSizePacked, cmpEngine)
                
                    'If the decompression fails, it may be due to a bad compression encoding (e.g. the data is compressed using a system other
                    ' than the one flagged in the file.)  Try again using a different compression method.
                    If (Not decompressionResult) Then
                    
                        InternalErrorMsg "WARNING!  pdPackager.ReadPackageFromFile failed to decode directory; trying again with alternate decompressors..."
                        
                        Dim i As PD_CompressionEngine
                        For i = PD_CE_ZLib To PD_CE_Lz4
                            decompressionResult = Compression.DecompressPtrToPtr(VarPtr(rawBuffer(0)), m_FileHeader.DirectoryChunkSizeOriginal, VarPtr(tmpCompressionBuffer(0)), m_FileHeader.DirectoryChunkSizePacked, i)
                            If decompressionResult Then
                                InternalErrorMsg "Directory may be recoverable; compressor #" & i & " worked.  Proceeding with load..."
                                Exit For
                            End If
                        Next i
                        
                        'If all decompression attempts failed, we're shit out of luck
                        If (Not decompressionResult) Then
                            InternalErrorMsg "Sorry, but no decompressor worked.  Something's definitely wrong with this package."
                        End If
                        
                    End If
                    
                Else
                    InternalErrorMsg "WARNING!  Required decompression engine missing; pdPackage directory cannot be read."
                End If
                
                'If decompression failed, we have to abandon the load process, as we can't even access the original file's directory.
                If (Not decompressionResult) Then
                    InternalErrorMsg "WARNING!  Could not decompress pdPackage directory.  Possible explanations include missing decompression library or file corruption."
                    GoTo StopPackageMemoryRead
                End If
                    
            End If
            
        'Source length is invalid (too small)
        Else
            InternalErrorMsg "WARNING!  Could not retrieve pdPackage directory; supplied mem length too small."
            GoTo StopPackageMemoryRead
        End If
        
        'Free the compressed copy of the directory (if any)
        Erase tmpCompressionBuffer
            
        'Use information from the file header to prepare the directory array, then copy the raw buffer into it.
        m_numOfNodes = m_FileHeader.NodeCount
        ReDim m_NodeDirectory(0 To m_numOfNodes - 1) As PDP_NODE
        CopyMemoryStrict VarPtr(m_NodeDirectory(0)), VarPtr(rawBuffer(0)), m_FileHeader.DirectoryChunkSizeOriginal
        
        'Free the raw directory buffer
        Erase rawBuffer
        
        'The package's directory is now available in-memory.  Before proceeding, make a note of the stream mode and access,
        ' which affect how we pull data out of this file.  (Note that stream mode is hard-coded to MEMORY because we are
        ' obviously loading the data from memory, not from file!)
        m_StreamMode = PD_SM_MemoryBacked
        m_StreamAccess = streamAccess
        
        'Finally, retrieve the data chunk.  Note that we basically just wrap a pdStream object around the data, and it handles
        ' everything from there.
        If (m_DataBuffer Is Nothing) Then Set m_DataBuffer = New pdStream Else m_DataBuffer.StopStream False
        If (m_FileHeader.DataChunkSizePacked = m_FileHeader.DataChunkSizeOriginal) Then
            
            'For a memory-backed stream, immediately load the full data chunk into memory.
            m_DataBuffer.StartStream PD_SM_MemoryBacked, m_StreamAccess
            m_DataBuffer.EnsureBufferSpaceAvailable m_FileHeader.DataChunkSizeOriginal, True
            CopyMemoryStrict m_DataBuffer.Peek_PointerOnly(0), srcPointer + m_FileHeader.DataChunkPosition, m_FileHeader.DataChunkSizeOriginal
            
            'Reset the buffer's internal data pointer (instead of letting it auto-position itself at the end of the buffer.)
            m_DataBuffer.SetSizeExternally m_FileHeader.DataChunkSizeOriginal
            m_DataBuffer.SetPosition 0
                
        Else
            InternalErrorMsg "Data chunk sizes don't match.  Package cannot be read."
            GoTo StopPackageMemoryRead
        End If
        
    End If
    
    'Memory-based load complete!
    ReadPackageFromMemory = True
    
    
    Exit Function
    
StopPackageMemoryRead:
    
    InternalErrorMsg "An internal error occurred in the pdPackager.ReadPackageFromMemory function."
    ReadPackageFromMemory = False
    
End Function

'Returns the version of a loaded pdPackage.  Return value is meaningless if no package is loaded.
Friend Function GetPDPackageVersion() As Long
    GetPDPackageVersion = m_FileHeader.PDP_Version
End Function

'Returns the current number of nodes in the package.  This is primarily designed to be used right after loading a pdPackage file;
' after this number is obtained, the caller can iterate through individual nodes, extracting data as they go.
Friend Function GetNumOfNodes() As Long
    GetNumOfNodes = m_numOfNodes
End Function

'Given a node index, return relevant header data for that node.  This function will fail if the node index is out of bounds.
' There is generally no need to use this function, as the direct node data access functions will handle all this automatically.
' But I guess it's here if you need it.
Friend Function GetNodeInfo(ByVal nodeIndex As Long, Optional ByRef dstNodeName As String, Optional ByRef dstNodeID As Long, Optional ByRef dstOptionalNodeType As Long) As Boolean

    'Start by validating the node index we were passed.  If it's out of range, exit immediately.
    If (nodeIndex < 0) Or (nodeIndex > m_numOfNodes - 1) Then
        InternalErrorMsg "Node index out of range - try again with a valid node index!"
        GetNodeInfo = False
        Exit Function
    End If
    
    'Fill the supplied variables with the corresponding data for this node
    With m_NodeDirectory(nodeIndex)
        dstNodeName = .nodeName
        dstNodeID = .NodeID
        dstOptionalNodeType = .OptionalNodeType
    End With
    
    GetNodeInfo = True

End Function

'Shortcut function for extracting a node's contents directly to file.  This function requires the data to have been added via the dedicated
' AutoAddNodeFromFile() function, above, as that function automatically populates all necessary fields for file extraction.
'
'In the future, additional functions will be added for extracting all files in the pdPackage, but for this function, a source filename
' (WITH EXTENSION) is required.  This filename is case-sensitive.  If the desired file was originally added using a relative path, YOU MUST
' INCLUDE THE RELATIVE PATH in the filename sent to this function.  This is required as a pdPackage may contain something like
' "\Subfolder1\Filename.txt" and "\Subfolder2\Filename.txt", and if extracting by filename, it can't differentiate between these two unless
' you provide the relative path.
'
'If the requested file was originally added with a relative path (e.g. "/Subfolder/Filename.txt") then two things should be noted:
' 1) Again, the full original filename PLUS PATH must be supplied as srcFilename
' 2) When writing out the file, any relative folders will be forcibly created, as necessary.  If they cannot be created, the file write will
'    (obviously) fail.
' 3) If you don't want to reconstruct the file's relative path, simply supply your own custom path+filename in dstPath, then set the
'    optional dstPathIncludesFilename to TRUE.  This will cause pdPackage to ignore the original relative path and filename information.
'
'If you don't want to mess with relative paths, you have two choices:
' 1) Simply supply bare filenames to the autoAddNodeFromFile function in the first place (duh).
' 2) Set the optional ignoreRelativePathIfPresent parameter to TRUE.  Note that this parameter does not change the requirement for passing a
'    relative path in as srcFilename; all it does is strip the relative path(s) prior to writing the file.
'
'As a convenience, this function also provides built-in file rename functionality.  If the dstPathIncludesFilename value is set to TRUE, the
' file's original name (AND ANY RELATIVE PATHS, if included), will be ignored in favor of the supplied dstPath string.  Obviously, if using
' this feature, make sure dstPath includes the filename and extension, e.g. dstPath = "C:\newFilename.ext"
'
'Returns TRUE if extraction was successful.  If a relative path must be constructed, a TRUE return means the path was successfully constructed.
Friend Function AutoExtractSingleFile(ByVal dstPath As String, ByVal srcFilename As String, Optional ByVal dstPathIncludesFilename As Boolean = False, Optional ByVal OptionalNodeType As Long = 0, Optional ByVal ignoreRelativePathIfPresent As Boolean = False) As Boolean

    'Node headers use a fixed-length 32 character string for storing names.  As such, we pad the file node type identifier to 32 chars.
    Dim nodeName32_1 As String * 32, nodeName32_2 As String * 32
    nodeName32_1 = PDP_DATA_FILE
    nodeName32_2 = PDP_DATA_FILEANDPATH
    
    'If this node contains a relative path, we'll detect it automatically
    Dim thisNodeType As PDP_SPECIAL_NODE_TYPE
    
    'Cast the source filename to an array.  This allows us to make direct byte array comparisons without coercing the byte arrays into strings.
    Dim filenameArray() As Byte
    ReDim filenameArray(0 To Len(srcFilename) * 2 - 1) As Byte
    CopyMemoryStrict VarPtr(filenameArray(0)), StrPtr(srcFilename), UBound(filenameArray) + 1
    
    'Start searching the node array for matches
    Dim retBytes() As Byte
    Dim i As Long, nodeIndex As Long, nodeFound As Boolean
    nodeIndex = -1
    nodeFound = False
    
    For i = 0 To UBound(m_NodeDirectory)
    
        'Look for two types of node names: bare filenames, and filenames with relative paths attached.
        ' (pdPackage handles both types automatically)
        If Strings.StringsEqual(nodeName32_1, m_NodeDirectory(i).nodeName, False) Then
            nodeFound = True
            thisNodeType = PDP_SN_Filename
        ElseIf Strings.StringsEqual(nodeName32_2, m_NodeDirectory(i).nodeName, False) Then
            nodeFound = True
            thisNodeType = PDP_SN_FilenameAndRelativePath
        Else
            nodeFound = False
        End If
        
        If nodeFound Then
        
            'This node appears to be a file node.  As a failsafe, check for the special file-type node flag.
            If GetBitFlag(PDP_NF2_FileNode, m_NodeDirectory(i).NodeFlags(0)) Then
            
                'This is a valid file-type node.  If the user supplied an optionalNodeType, check it now.
                nodeFound = True
                
                If (OptionalNodeType <> 0) Then
                    nodeFound = (m_NodeDirectory(i).OptionalNodeType = OptionalNodeType)
                End If
                
                'If the optional node type check succeeded, the last thing we need to do is check the actual filename of this node.
                If nodeFound Then
                    
                    'Extract the header array, which will contain the filename of this file
                    If GetNodeDataByIndex(i, True, retBytes) Then
                    
                        'If this byte array equals the source filename, we have a match (DING DING DING).
                        If CheckArrayEquality(filenameArray, retBytes) Then
                            nodeFound = True
                            nodeIndex = i
                            Exit For
                        Else
                            nodeFound = False
                        End If
                    
                    Else
                        nodeFound = False
                    End If
                    
                End If
                
            Else
                nodeFound = False
            End If
            
        End If
    
    Next i
    
    'If a matching node was found, use getNodeDataByIndex() to retrieve its data
    If (nodeIndex >= 0) And nodeFound Then
        
        'As a convenience to the caller, a rename can be applied at extraction time by supplying a destination path that includes a target filename,
        ' and notifying us via the dstPathIncludesFilename flag.  If that flag is NOT present, we create the full destination filename by appending
        ' the original filename+extension to the specified path.
        Dim fullDestinationPath As String
        
        If dstPathIncludesFilename Then
            fullDestinationPath = dstPath
        Else
            
            'Make sure the destination path has a trailing backslash
            dstPath = m_File.PathAddBackslash(dstPath)
            
            Select Case thisNodeType
            
                'This path is just a bare filename.  Append it to the destination path and carry on
                Case PDP_SN_Filename
                    fullDestinationPath = dstPath & srcFilename
                
                'This path includes a relative path.  Additional work is required.
                Case PDP_SN_FilenameAndRelativePath
                
                    'The user wants us to ignore the relative path and just write the damn file - easy enough!
                    If ignoreRelativePathIfPresent Then
                        fullDestinationPath = dstPath & m_File.FileGetName(srcFilename)
                    
                    'Relative path(s) must be preserved.
                    Else
                    
                        'Start by assembling the desired final filename, with all relative folders attached.
                        
                        'First, strip any preceding slashes in the relative path
                        If Strings.StringsEqual(Left$(srcFilename, 1), "\", False) Then srcFilename = Right$(srcFilename, Len(srcFilename) - 1)
                        
                        'Concatenate the two paths
                        Dim dstPathComplete As String
                        dstPathComplete = dstPath & m_File.FileGetPath(srcFilename)
                        
                        'Attempt to create the folder.  (If the folder already exists, this will return TRUE, so no harm.)
                        If m_File.PathCreate(dstPathComplete, True) Then
                        
                            'The full target path exists.  Append the original filename to it, then proceed with the extraction
                            fullDestinationPath = m_File.PathAddBackslash(dstPathComplete) & m_File.FileGetName(srcFilename)
                        
                        'The path could not be constructed.  Set the full destination path to a blank string, which will cause the function
                        ' to terminate prematurely.
                        Else
                            fullDestinationPath = vbNullString
                        End If
                    
                    End If
                    
            End Select
            
        End If
        
        'Make sure the destination path was successfully constructed, including any relative path data
        If (LenB(fullDestinationPath) > 0) Then
        
            'Retrieve this node's contents
            If GetNodeDataByIndex(nodeIndex, False, retBytes()) Then
                If m_File.FileExists(fullDestinationPath) Then m_File.FileDelete fullDestinationPath
                m_File.FileCreateFromByteArray retBytes, fullDestinationPath, True
                AutoExtractSingleFile = True
            Else
                AutoExtractSingleFile = False
            End If
            
        Else
            InternalErrorMsg "WARNING! pdPackage.AutoExtractSingleFile() was unable to construct the output path for " & srcFilename & "."
            InternalErrorMsg "WARNING! As such, the file was not written.  (But one or more folders may have been constructed as part of the extraction.)"
            AutoExtractSingleFile = False
        End If
        
    Else
        AutoExtractSingleFile = False
    End If

End Function

'Shortcut function to extract all files embedded into this pdPackage instance.  This function will extract all files to the specified
' destination path, using whatever relative path settings were established at creation time.  Subfolders will automatically be created
' as necessary.
'
'If you do not want relative paths used, set the optional ignoreRelativePaths parameter to TRUE.  This will cause all files to be
' dumped directly into the destination path.  (Also remember: if relative paths were explicitly ignored at creation time, this function
' cannot recreate them, as they don't exist inside the pdPackage data stream.)
'
'If one or more OptionalNodeType values were used at creation time, you may pass such a value to this function.  This will result in
' extraction of only the file(s) that match that node type value.  (This makes it possible to package multiple file groups into one
' package, then extract them as separate groups later.)
'
'Unlike other extraction functions, this function will automatically make use of checksums if they were embedded at creation time.  This
' behavior cannot be overridden, by design, as it's an important failsafe against package corruption, damage during transit, or tampering.
'
'Returns: TRUE if all files are extracted successfully, FALSE otherwise.  FALSE should not occur unless write access to the target folder
' is restricted.
Friend Function AutoExtractAllFiles(ByVal dstPath As String, Optional ByVal OptionalNodeType As Long = 0, Optional ByVal ignoreRelativePaths As Boolean = False) As Boolean
    
    'Enforce a trailing slash on dstPath, as we'll potentially be appending a lot of filenames to it
    dstPath = m_File.PathAddBackslash(dstPath)
    
    'Node headers use a fixed-length 32 character string for storing names.  As such, we pad the two valid file node type identifiers to
    ' 32 chars in advance.
    Dim nodeName32_1 As String * 32, nodeName32_2 As String * 32
    nodeName32_1 = PDP_DATA_FILE
    nodeName32_2 = PDP_DATA_FILEANDPATH
    
    'Nodes with relative paths are detected and handled automatically
    Dim thisNodeType As PDP_SPECIAL_NODE_TYPE
    
    'By default, we assume SUCCESS.  This value is && with the results of each individual file extraction, so it will return FALSE if any
    ' individual file fails.  We don't halt extraction on a single failure; instead, we try to complete as many extractions as we can.
    Dim allNodesSuccessful As Boolean
    allNodesSuccessful = True
    
    'Start searching the node array for any valid file entries.
    Dim retBytes() As Byte, strFilename As String
    Dim i As Long, nodeIndex As Long, nodeFound As Boolean
    nodeIndex = -1
    nodeFound = False
    
    For i = 0 To UBound(m_NodeDirectory)
    
        'Look for two types of node names: bare filenames, and filenames with relative paths attached.
        ' (pdPackage handles both types automatically)
        If Strings.StringsEqual(nodeName32_1, m_NodeDirectory(i).nodeName, False) Then
            nodeFound = True
            thisNodeType = PDP_SN_Filename
        ElseIf Strings.StringsEqual(nodeName32_2, m_NodeDirectory(i).nodeName, False) Then
            nodeFound = True
            thisNodeType = PDP_SN_FilenameAndRelativePath
        Else
            nodeFound = False
        End If
        
        If nodeFound Then
        
            'This node appears to be a file node.  As a failsafe, check for the special file-type node flag.
            If GetBitFlag(PDP_NF2_FileNode, m_NodeDirectory(i).NodeFlags(0)) Then
            
                'This is a valid file-type node.  If the user supplied an optionalNodeType, check it now.
                nodeFound = True
                
                If (OptionalNodeType <> 0) Then
                    nodeFound = (m_NodeDirectory(i).OptionalNodeType = OptionalNodeType)
                End If
                
                'If the optional node type check succeeded, this is a valid file.  Prepare to extract it.
                If nodeFound Then
                    
                    'Extract the header array, which contain the file's name and any relative path data
                    If GetNodeDataByIndex(i, True, retBytes) Then
                        
                        'Cast the returned bytes directly into a string; this preserves Unicode characters.
                        strFilename = Space$((UBound(retBytes) + 1) \ 2)
                        CopyMemoryStrict StrPtr(strFilename), VarPtr(retBytes(0)), UBound(retBytes) + 1
                        
                        'Now we need to construct a full destination path + filename.  This can vary by several criteria.
                        Dim fullDestinationPath As String
                        
                        'If the caller wants relative paths ignored, extract just the filename from the original path, and append it to the destination path.
                        If ignoreRelativePaths Then
                            fullDestinationPath = dstPath & m_File.FileGetName(strFilename)
                        
                        'The caller wants to make use of relative paths.
                        Else
                            
                            Select Case thisNodeType
                            
                                'This path is just a bare filename.  Append it to the destination path and carry on.
                                Case PDP_SN_Filename
                                    fullDestinationPath = dstPath & strFilename
                                
                                'This path includes a relative path.  Additional work is required to make sure the destination folder hierarchy exists.
                                Case PDP_SN_FilenameAndRelativePath
                                
                                    'Start by assembling the desired final filename, with all relative folders attached.
                                    
                                    'First, strip any preceding slashes in the relative path
                                    If Strings.StringsEqual(Left$(strFilename, 1), "\", False) Then strFilename = Right$(strFilename, Len(strFilename) - 1)
                                    
                                    'Concatenate the two paths
                                    Dim dstPathComplete As String
                                    dstPathComplete = dstPath & m_File.FileGetPath(strFilename)
                                    
                                    'Attempt to create the folder hierarchy.  (If the folder already exists, this will return TRUE, so no harm.)
                                    If m_File.PathCreate(dstPathComplete, True) Then
                                    
                                        'The full target path exists.  Append the original filename to it, then proceed with the extraction
                                        fullDestinationPath = m_File.PathAddBackslash(dstPathComplete) & m_File.FileGetName(strFilename)
                                        
                                    'The path could not be constructed.  Set the full destination path to a blank string, which will cause the function
                                    ' to terminate prematurely.
                                    Else
                                        fullDestinationPath = vbNullString
                                    End If
                                    
                            End Select
                        
                        End If
                        
                        'Make sure the destination path was successfully constructed, including any relative path data.
                        ' (If constructing the relative path failed, the destination path string will have been erased.)
                        If (LenB(fullDestinationPath) <> 0) Then
                            Debug.Print "working on " & fullDestinationPath
                            'Retrieve this node's contents
                            If GetNodeDataByIndex(i, False, retBytes()) Then
                            
                                'Dump the node's contents to file
                                If m_File.FileExists(fullDestinationPath) Then m_File.FileDelete fullDestinationPath
                                m_File.FileCreateFromByteArray retBytes, fullDestinationPath
                                
                            Else
                                InternalErrorMsg "WARNING! pdPackage.AutoExtractAllFiles() could not extract file contents from an otherwise valid node.  Extraction abandoned."
                                allNodesSuccessful = False
                            End If
                            
                        Else
                            InternalErrorMsg "WARNING! pdPackage.AutoExtractAllFiles() was unable to construct the output path for " & strFilename & "."
                            InternalErrorMsg "WARNING! As such, the file was not written.  (But one or more folders may have been constructed as part of the extraction.)"
                            allNodesSuccessful = False
                        End If
                        
                    Else
                        allNodesSuccessful = False
                        InternalErrorMsg "WARNING! pdPackage.AutoExtractAllFiles() could not extract a filename from an otherwise valid node.  Extraction abandoned."
                    End If
                    
                'This set of 3 EndIf's looks confusing, but they simply terminate all file node detection criteria blocks.
                ' No action needs to be taken if any node detection checks fail; that just means this node isn't part of
                ' the current extraction request.
                End If
            End If
        End If
    
    Next i
    
    'Report success.  allNodesSuccessful will be TRUE IFF there were no failures or errors during extraction.
    AutoExtractAllFiles = allNodesSuccessful

End Function

'Are two byte arrays equal?
Private Function CheckArrayEquality(ByRef srcArray1() As Byte, ByRef srcArray2() As Byte) As Boolean
    
    'Check bounds first
    If (UBound(srcArray1) <> UBound(srcArray2)) Then
        CheckArrayEquality = False
    
    'Array bounds are equal, so an actual per-element test is required
    Else
        CheckArrayEquality = VBHacks.MemCmp(VarPtr(srcArray1(0)), VarPtr(srcArray2(0)), UBound(srcArray1) + 1)
    End If

End Function

'Node data access functions.  These functions fill a byte array with the requested data, and they will return TRUE if successful.
' (FALSE generally only occurs if the requested node cannot be found in the source data.)
'
'If useHeaderBuffer is TRUE, the node's header buffer will be retrieved.  Otherwise, the node's data buffer will be used.
' Also note that decompression is handled automatically, depending on the contents of the source file.
'
'For an explanation of the trimDstArray and dstFinalSize parameters, please see GetNodeDataByIndex(), below (which this function
' merely wraps).
Friend Function GetNodeIndexFromName(ByRef targetNodeName As String, Optional ByVal compareMode As VbCompareMethod = vbTextCompare) As Long
    
    'Node headers use a fixed-length 32 character string for storing names.  As such, pad the supplied node name to
    ' 32 chars as necessary.
    Dim nodeName32 As String * 32
    nodeName32 = targetNodeName
    
    'Search the node array for a matching name
    GetNodeIndexFromName = -1
    
    If (m_numOfNodes > 0) Then
    
        Dim caseInsensitiveMatching As Boolean
        caseInsensitiveMatching = (compareMode = vbTextCompare)
        
        Dim i As Long
        For i = 0 To UBound(m_NodeDirectory)
            If Strings.StringsEqual(nodeName32, m_NodeDirectory(i).nodeName, caseInsensitiveMatching) Then
                GetNodeIndexFromName = i
                Exit For
            End If
        Next i
        
    End If
    
End Function

Friend Function GetNodeIndexFromID(ByVal targetNodeID As Long) As Long

    'Search the node array for a matching ID.  Return -1 if no match is found,
    Dim i As Long, nodeIndex As Long
    nodeIndex = -1
    
    If (m_numOfNodes > 0) Then
        For i = 0 To UBound(m_NodeDirectory)
            If (targetNodeID = m_NodeDirectory(i).NodeID) Then
                nodeIndex = i
                Exit For
            End If
        Next i
    End If
    
    GetNodeIndexFromID = nodeIndex
    
End Function

Friend Function GetNodeDataByName(ByRef targetNodeName As String, ByVal useHeaderBuffer As Boolean, ByRef dstArray() As Byte, Optional ByVal trimDstArray As Boolean = True, Optional ByRef dstFinalSize As Long = 0) As Boolean

    Dim nodeIndex As Long
    nodeIndex = GetNodeIndexFromName(targetNodeName)
    
    'If a matching node was found, use GetNodeDataByIndex() to retrieve its data
    If (nodeIndex >= 0) Then
        GetNodeDataByName = GetNodeDataByIndex(nodeIndex, useHeaderBuffer, dstArray(), trimDstArray, dstFinalSize)
    Else
        GetNodeDataByName = False
    End If

End Function

Friend Function GetNodeDataByName_String(ByRef targetNodeName As String, ByVal useHeaderBuffer As Boolean, ByRef dstString As String) As Boolean

    Dim nodeIndex As Long
    nodeIndex = GetNodeIndexFromName(targetNodeName)
    
    'If a matching node was found, use GetNodeDataByIndex() to retrieve its data
    If (nodeIndex >= 0) Then
        GetNodeDataByName_String = GetNodeDataByIndex_String(nodeIndex, useHeaderBuffer, dstString)
    Else
        GetNodeDataByName_String = False
    End If

End Function

Friend Function GetNodeDataByName_UnsafeDstPointer(ByRef targetNodeName As String, ByVal useHeaderBuffer As Boolean, ByVal dstPointer As Long) As Boolean

    Dim nodeIndex As Long
    nodeIndex = GetNodeIndexFromName(targetNodeName)
    
    'If a matching node was found, use GetNodeDataByIndex() to retrieve its data
    If (nodeIndex >= 0) Then
        GetNodeDataByName_UnsafeDstPointer = GetNodeDataByIndex_UnsafeDstPointer(nodeIndex, useHeaderBuffer, dstPointer)
    Else
        GetNodeDataByName_UnsafeDstPointer = False
    End If

End Function

Friend Function GetNodeDataByID(ByVal targetNodeID As Long, ByVal useHeaderBuffer As Boolean, ByRef dstArray() As Byte, Optional ByVal trimDstArray As Boolean = True, Optional ByRef dstFinalSize As Long = 0) As Boolean

    'Search the node array for a matching ID
    Dim nodeIndex As Long
    nodeIndex = GetNodeIndexFromID(targetNodeID)
        
    'If a matching node was found, use GetNodeDataByIndex() to retrieve its data
    If (nodeIndex >= 0) Then
        GetNodeDataByID = GetNodeDataByIndex(nodeIndex, useHeaderBuffer, dstArray(), trimDstArray, dstFinalSize)
    Else
        GetNodeDataByID = False
    End If
    
End Function

Friend Function GetNodeDataByID_UnsafeDstPointer(ByVal targetNodeID As Long, ByVal useHeaderBuffer As Boolean, ByVal dstPointer As Long) As Boolean

    'Search the node array for a matching ID
    Dim nodeIndex As Long
    nodeIndex = GetNodeIndexFromID(targetNodeID)
        
    'If a matching node was found, use GetNodeDataByIndex() to retrieve its data
    If (nodeIndex >= 0) Then
        GetNodeDataByID_UnsafeDstPointer = GetNodeDataByIndex_UnsafeDstPointer(nodeIndex, useHeaderBuffer, dstPointer)
    Else
        GetNodeDataByID_UnsafeDstPointer = False
    End If
        
End Function

'GetNodeHeaderByName and GetNodeHeaderByID both wrap this function.  (Those functions simply calculate the relevant index for
' their requested node, then return the results of this function.)
'
'Given a node index, and a target location (header buffer or data buffer), fill a destination array with the contents of that
' data buffer.  Decompression is handled automatically, depending on the contents of the node.
'
'By default, trimDstArray is TRUE, which tells this function to size the destination array to *exactly* match the size of the
' read data.  This is not a great solution for performance, as it means a unique destination array will likely be allocated
' on each call to this function.  Instead, it is suggested that you set trimDstArray to FALSE, then read the final size
' actually returned by checking the optional dstFinalSize parameter.  This allows you to reuse destination arrays wherever
' possible, reducing allocations and greatly improving read performance.
Friend Function GetNodeDataByIndex(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByRef dstArray() As Byte, Optional ByVal trimDstArray As Boolean = True, Optional ByRef dstFinalSize As Long = 0) As Boolean

    'Retrieve the offset, packed size, and original (unpacked) size of the target data
    Dim dataOffset As Long, dataPackedSize As Long, dataOriginalSize As Long
    
    If useHeaderBuffer Then
        dataOffset = m_NodeDirectory(nodeIndex).NodeHeaderOffset
        dataPackedSize = m_NodeDirectory(nodeIndex).NodeHeaderPackedSize
        dataOriginalSize = m_NodeDirectory(nodeIndex).NodeHeaderOriginalSize
    Else
        dataOffset = m_NodeDirectory(nodeIndex).NodeDataOffset
        dataPackedSize = m_NodeDirectory(nodeIndex).NodeDataPackedSize
        dataOriginalSize = m_NodeDirectory(nodeIndex).NodeDataOriginalSize
    End If
    
    'Move the data buffer position to the relevant offset for the requested chunk
    m_DataBuffer.SetPosition dataOffset
    
    'Before retrieving the data, we need to ensure the destination buffer is large enough to hold the results.
    ' (For performance reasons, we ReDim the destination buffer *only* if absolutely necessary.)
    Dim needToSizeArray As Boolean: needToSizeArray = False
    If VBHacks.IsArrayInitialized(dstArray) Then
        
        'If the user wants the destination array trimmed, we can still check to see if it's already the
        ' ideal size; while the odds of this are low, the potential performance gains are large, so it's
        ' worth the trouble.
        If trimDstArray Then
            If ((UBound(dstArray) - LBound(dstArray) + 1) <> dataOriginalSize) Then needToSizeArray = True
        
        'Meanwhile, if the caller doesn't care about a perfectly trimmed array, we still need to ensure
        ' sufficient size is available for the decompression operation.
        Else
            If ((UBound(dstArray) - LBound(dstArray) + 1) < dataOriginalSize) Then needToSizeArray = True
        End If
        
    Else
        needToSizeArray = True
    End If
    
    If needToSizeArray Then ReDim dstArray(0 To dataOriginalSize - 1) As Byte
    dstFinalSize = dataOriginalSize
    
    'If the original size and packed size are different, the data must be compressed.
    If (dataPackedSize <> dataOriginalSize) Then
        
        'Now that the destination array is appropriately sized, we can offload further loading to the unsafe pointer-only function
        GetNodeDataByIndex = Me.GetNodeDataByIndex_UnsafeDstPointer(nodeIndex, useHeaderBuffer, VarPtr(dstArray(LBound(dstArray))))
        
    'If the data is uncompressed, make a direct copy of the corresponding chunk of the master buffer
    Else
        If (dataPackedSize <> 0) Then
            m_DataBuffer.ReadBytes dstArray, dataPackedSize, trimDstArray
            dstFinalSize = dataOriginalSize
        Else
            InternalErrorMsg "File node is uncompressed, but size is listed as zero.  Load impossible; abandoning attempt."
            GetNodeDataByIndex = False
            Exit Function
        End If
    End If
    
    'If we made it all the way here, the node has been successfully loaded!
    GetNodeDataByIndex = True

End Function

'WARNING!  This function is for expert users only.  Do not use it unless you fully understand the repercussions described below.
'
'Given a node index and a target buffer region (header or data), blindly copy the contents of that data buffer to a pointer supplied by the user.
' Decompression is handled automatically, depending on the contents of the node.
'
'What makes this function different is that it completely relies on the caller to create a correctly sized buffer for the extracted data.
' For common cases, this is simply not possible, as the caller does not have enough knowledge necessary to prep such a buffer.  Instead, it relies
' on pdPackager to construct the destination buffer for it, using the size values stored inside the pdPackage instance.
'
'PhotoDemon is unique in this regard, as it can correctly calculate the size of raster layer pixel buffers (because it knows the associated
' width, height, stride and color depth of the buffer).  Because PD has to create a matching DIB in advance, a destination buffer of perfect size
' is already implicitly allocated, so there is no need for a temporary VB array to shuttle the data around.
'
'That said, note that PD only uses this function sparingly, when it is *absolutely certain of its safety*.  You would be wise to do the same.
' An incorrectly allocated destination buffer will cause mass destruction (and a guaranteed hard crash), so do not use this function ignorantly.
Friend Function GetNodeDataByIndex_UnsafeDstPointer(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByVal dstPointer As Long) As Boolean

    'Retrieve the offset, packed size, and original (unpacked) size of the target data
    Dim dataOffset As Long, dataPackedSize As Long, dataOriginalSize As Long
    
    If useHeaderBuffer Then
        dataOffset = m_NodeDirectory(nodeIndex).NodeHeaderOffset
        dataPackedSize = m_NodeDirectory(nodeIndex).NodeHeaderPackedSize
        dataOriginalSize = m_NodeDirectory(nodeIndex).NodeHeaderOriginalSize
    Else
        dataOffset = m_NodeDirectory(nodeIndex).NodeDataOffset
        dataPackedSize = m_NodeDirectory(nodeIndex).NodeDataPackedSize
        dataOriginalSize = m_NodeDirectory(nodeIndex).NodeDataOriginalSize
    End If
    
    'Move the data buffer position to the relevant offset for the requested chunk
    m_DataBuffer.SetPosition dataOffset
    
    'If the original size and packed size are different, assume the data is compressed.
    If (dataPackedSize <> dataOriginalSize) Then
        
        Dim decompressionResult As Boolean: decompressionResult = False
        
        'Figure out which decompression library was used for the original compression.
        ' (Note that if the required decompression library isn't available, we're screwed; loading is impossible.)
        Dim cmpEngine As PD_CompressionEngine: cmpEngine = PD_CE_Zstd
        If useHeaderBuffer Then
            If GetBitFlag(PDP_NF2_ZstdRequiredH, m_NodeDirectory(nodeIndex).NodeFlags(0)) Then
                cmpEngine = PD_CE_Zstd
            ElseIf GetBitFlag(PDP_NF2_ZlibRequiredH, m_NodeDirectory(nodeIndex).NodeFlags(0)) Then
                cmpEngine = PD_CE_ZLib
            ElseIf GetBitFlag(PDP_NF2_Lz4RequiredH, m_NodeDirectory(nodeIndex).NodeFlags(0)) Then
                cmpEngine = PD_CE_Lz4
            End If
        Else
            If GetBitFlag(PDP_NF2_ZstdRequired, m_NodeDirectory(nodeIndex).NodeFlags(0)) Then
                cmpEngine = PD_CE_Zstd
            ElseIf GetBitFlag(PDP_NF2_ZlibRequired, m_NodeDirectory(nodeIndex).NodeFlags(0)) Then
                cmpEngine = PD_CE_ZLib
            ElseIf GetBitFlag(PDP_NF2_Lz4Required, m_NodeDirectory(nodeIndex).NodeFlags(0)) Then
                cmpEngine = PD_CE_Lz4
            End If
        End If
        
        If Compression.IsCompressionEngineAvailable(cmpEngine) Then
            
            Dim i As PD_CompressionEngine
            
            'I'm not happy about this, but for the moment, we have to split handling here.  In file-backed mode, we can't currently
            ' read file data from blind pointers - instead, we need to load the compressed data into a temporary buffer, then decompress
            ' it from there.  This is fixable once a memory-mapped backer is in-place.
            
            'Memory mode doesn't require a temporary buffer; just decompress immediately
            If (m_StreamMode = PD_SM_MemoryBacked) Then
            
                decompressionResult = Compression.DecompressPtrToPtr(dstPointer, dataOriginalSize, m_DataBuffer.ReadBytes_PointerOnly(dataPackedSize), dataPackedSize, cmpEngine)
                
                'If the decompression fails, it may be due to a bad compression encoding (e.g. the data is compressed using a system other
                ' than the one flagged in the file.)  Try again using a different compression method.
                If (Not decompressionResult) Then
                
                    InternalErrorMsg "WARNING!  pdPackager.GetNodeDataByIndex_UnsafeDstPointer failed to decode node; trying again with alternate decompressors..."
                    
                    For i = PD_CE_ZLib To PD_CE_Lz4
                        decompressionResult = Compression.DecompressPtrToPtr(dstPointer, dataOriginalSize, m_DataBuffer.ReadBytes_PointerOnly(dataPackedSize), dataPackedSize, i)
                        If decompressionResult Then
                            InternalErrorMsg "Node may be recoverable; compressor #" & i & " worked.  Proceeding with load..."
                            Exit For
                        End If
                    Next i
                    
                    'If all decompression attempts failed, we're shit out of luck
                    If (Not decompressionResult) Then
                        InternalErrorMsg "Sorry, but no decompressor worked.  Something's definitely wrong with this node."
                    End If
                    
                End If
                
            ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            
                'In the future, read-only mode will allow us to use memory-mapping to retrieve the data.  Read/write mode will require
                ' a temporary buffer, however, as we don't have pointer-level access to the source file.
                
                'Prep the temporary compression buffer
                If VBHacks.IsArrayInitialized(m_CompressionBuffer) Then
                    If (UBound(m_CompressionBuffer) < (dataPackedSize - 1)) Then ReDim m_CompressionBuffer(0 To dataPackedSize - 1) As Byte
                Else
                    ReDim m_CompressionBuffer(0 To dataPackedSize - 1) As Byte
                End If
                
                'Retrieve the compressed data, then perform immediate decompression
                If (m_DataBuffer.ReadBytes(m_CompressionBuffer, dataPackedSize, False) <> 0) Then
                    
                    decompressionResult = Compression.DecompressPtrToPtr(dstPointer, dataOriginalSize, VarPtr(m_CompressionBuffer(0)), dataPackedSize, cmpEngine)
                    
                    'If the decompression fails, it may be due to a bad compression encoding (e.g. the data is compressed using a system other
                    ' than the one flagged in the file.)  Try again using a different compression method.
                    If (Not decompressionResult) Then
                    
                        InternalErrorMsg "WARNING!  pdPackager.GetNodeDataByIndex_UnsafeDstPointer failed to decode node; trying again with alternate decompressors..."
                        
                        For i = PD_CE_ZLib To PD_CE_Lz4
                            decompressionResult = Compression.DecompressPtrToPtr(dstPointer, dataOriginalSize, VarPtr(m_CompressionBuffer(0)), dataPackedSize, cmpEngine)
                            If decompressionResult Then
                                InternalErrorMsg "Node may be recoverable; compressor #" & i & " worked.  Proceeding with load..."
                                Exit For
                            End If
                        Next i
                        
                        'If all decompression attempts failed, we're shit out of luck
                        If (Not decompressionResult) Then
                            InternalErrorMsg "Sorry, but no decompressor worked.  Something's definitely wrong with this node."
                        End If
                        
                    End If
                    
                End If
            
            End If
        End If
        
        'If decompression failed, we have to abandon the load process, as we can't even access the original file's directory.
        If (Not decompressionResult) Then
            InternalErrorMsg "WARNING!  Could not decompress pdPackage node #" & nodeIndex & ".  Possible explanations include missing decompression library or file corruption."
            GetNodeDataByIndex_UnsafeDstPointer = False
            Exit Function
        End If
        
    'If the data is uncompressed, make a direct copy of the corresponding chunk of the master buffer
    Else
        
        'As with compressed mode, a switch is required depending on the stream backing mode
        If (m_StreamMode = PD_SM_MemoryBacked) Then
            CopyMemoryStrict dstPointer, m_DataBuffer.ReadBytes_PointerOnly(dataPackedSize), dataPackedSize
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            m_DataBuffer.ReadBytesToBarePointer dstPointer, dataPackedSize
        End If
        
    End If
    
    'If we made it all the way here, the node has been successfully loaded!
    GetNodeDataByIndex_UnsafeDstPointer = True

End Function

'Shortcut function for casting a node to a destination string.  If you're extracting a bunch of nodes, this is not the most
' performance-friendly solution (as new strings are created on every read), so it may be worthwhile to use your own string and
' only resize it as necessary.
Friend Function GetNodeDataByIndex_String(ByVal targetNodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByRef dstString As String) As Boolean
    
    Dim retBytes() As Byte, retSize As Long
    If Me.GetNodeDataByIndex(targetNodeIndex, useHeaderBuffer, retBytes, False, retSize) Then
        dstString = Space$(retSize \ 2)
        CopyMemory ByVal StrPtr(dstString), ByVal VarPtr(retBytes(0)), retSize
        GetNodeDataByIndex_String = True
    Else
        GetNodeDataByIndex_String = False
    End If
    
End Function

'Set an arbitrary 1-bit position [0-31] in a Long-type value to either 1 or 0.
' Inputs:
'  1) position of the flag, which must be in the range [0, 31]
'  2) value of the flag, TRUE for 1, FALSE for 0
'  3) The Long-type value where you want the flag placed
Private Sub SetBitFlag(ByVal flagPosition As Long, ByVal flagValue As Boolean, ByRef copyOfDstLong As Long)

    If (flagPosition >= 0) And (flagPosition <= 31) Then
    
        'Create a bitmask, using flagPosition to determine bit offset
        Dim longMask As Long
        If (flagPosition < 31) Then longMask = 2 ^ flagPosition Else longMask = &H80000000
        
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
        InternalErrorMsg "C'mon - you know Longs only have 32 bits.  Flag positions must be on the range [0, 31]!"
    End If

End Sub

'Retrieve an arbitrary 1-bit position [0-31] in a Long-type value.
' Inputs:
'  1) position of the flag, which must be in the range [0, 31]
'  2) The Long from which you want the bit retrieved
Private Function GetBitFlag(ByVal flagPosition As Long, ByVal srcLong As Long) As Boolean

    If (flagPosition >= 0) And (flagPosition <= 31) Then
    
        'Create a bitmask, using flagPosition to determine bit offset
        Dim longMask As Long
        If (flagPosition < 31) Then longMask = 2 ^ flagPosition Else longMask = &H80000000
        
        GetBitFlag = ((srcLong And longMask) <> 0)
        
    Else
        InternalErrorMsg "C'mon - you know Longs only have 32 bits.  Flag positions must be on the range [0, 31]!"
    End If

End Function

'Once a package has been loaded from file, you can use this function to check various flags in the file
Friend Function GetPackageFlag(ByVal flagToCheck As Byte, Optional ByVal checkLocation As PDP_FLAG_SOURCE = PDP_LOCATION_ANY, Optional ByVal nodeIndex As Long = 0) As Boolean

    'Check the relevant flag location
    Select Case checkLocation
    
        'Any location
        Case PDP_FS_Any
        
            'Check directory and data flags first
            GetPackageFlag = GetBitFlag(flagToCheck, m_FileHeader.DirectoryFlags(0)) Or GetBitFlag(flagToCheck, m_FileHeader.DataFlags(0))
            
            'If the flag is still false, scan individual nodes, looking for a hit
            If Not GetPackageFlag Then
            
                Dim i As Long
                For i = 0 To m_numOfNodes - 1
                    GetPackageFlag = GetPackageFlag Or GetBitFlag(flagToCheck, m_NodeDirectory(i).NodeFlags(0))
                    If GetPackageFlag Then Exit For
                Next i
            
            End If
        
        'Directory chunk flags
        Case PDP_FS_Directory
            GetPackageFlag = GetBitFlag(flagToCheck, m_FileHeader.DirectoryFlags(0))
        
        'Data chunk flags
        Case PDP_FS_Data
            GetPackageFlag = GetBitFlag(flagToCheck, m_FileHeader.DataFlags(0))
        
        'Specific node flags; if used, we'd better hope the user supplied the right node index!
        Case PDP_FS_IndividualNode
            GetPackageFlag = GetBitFlag(flagToCheck, m_NodeDirectory(nodeIndex).NodeFlags(0))
    
    End Select

End Function

Private Sub Class_Initialize()
    
    'Prepare a module-level pdFSO instance, which makes file operations much easier
    Set m_File = New pdFSO
    
End Sub

Private Sub Class_Terminate()
    
    'Release our data stream buffer
    If Not (m_DataBuffer Is Nothing) Then
        m_DataBuffer.StopStream
        Set m_DataBuffer = Nothing
        m_numOfNodes = 0
    End If
    
    'Release our pdFSO instance
    Set m_File = Nothing

End Sub

'If something unexpected occurs, and the user needs to know about it (e.g. it's not just debugging data), pdPackage will raise a message
' using this function.  PhotoDemon uses a central "Message" function to handle messages like this, but you can change it to Debug.Print instead.
Private Sub InternalErrorMsg(ByVal msgString As String, ParamArray ExtraText() As Variant)
        
    If (UBound(ExtraText) >= LBound(ExtraText)) Then
        Debug.Print msgString & " (extra parameters were supplied)"
    Else
        Debug.Print msgString
    End If
    
End Sub
