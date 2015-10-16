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
'PhotoDemon "pdPackage" Interface (e.g. Zip-like archive handler)
'Copyright 2014-2015 by Tanner Helland
'Created: 05/April/14
'Last updated: 06/February/15
'Last update: expand zip-like interface, with help from the new pdFSO class
'Dependencies: clsBasicBuffer class by dilettante (used to read/write byte array buffers more easily)
'              pdFSO class from photodemon.org (Unicode-friendly file operations)
'              pdStringStack class from photodemon.org (optimized handling of large string collections)
'
'This class provides an interface for creating and reading "pdPackage" files.  pdPackages are zip-like archive files
' that contain one or more "nodes" (e.g. pieces of data), compressed or uncompressed, in a VB-friendly structure.
'
'Though I have created pdPackages specifically for use with PhotoDemon, this class should be easily usable by others
' with only minor modifications.  Note that an explicit path to an STDCALL (sometimes called WAPI) variant of zlib.dll
' is required, with the expected filename of "zlibwapi.dll".  If your STDCALL copy is called "zLib.dll", you must
' rewrite the zLib API declarations at the top of the class to match.
'
'While pdPackages have many similarities to ZIP files, THEY ARE NOT ACTUAL ZIP FILES, and this class cannot read or
' write actual ZIP files.  pdPackages are, by design, much simpler than ZIP files, and their structure and layout
' plays better with VB's inherent limitations.
'
'Key pdPackage features include:
'
' 1) Data agnosticism, e.g. everything is treated as byte arrays, so you can store whatever data you'd like.
' 2) Front-loaded header.  ZIP files place the header at the tail of the archive, but pdPackages place the header
'     at the head.  This is not ideal for files that must be repeatedly edited, but it allows for much faster header
'     retrieval and parsing, which is a common use-case in PD.
' 3) Fixed-width directory entries.  This allows the entire archive directory to be read in a single operation,
'     rather than manually parsing variable-width directory entries until all have been located.
' 4) Support for zLib compression on a per-node basis, or for the entire directory and/or data chunks after they have
'     been assembled (so-called "second pass" compression).
' 5) Support for checksum (Adler32) validation of each individual node.
' 6) Support for two data entries per node, typically a header chunk and an actual data chunk.  These two structs
'     don't need to be used (one or the other or neither is just fine), but I find it very helpful to provide a simple
'     interface for storing two pieces of data per node; as an example, PD uses this feature to read key header data from
'     a layer node without extracting the full layer DIB contents (which tend to be enormous).
' 7) Per-node access system, including compression, so that you can easily extract a single node without having to
'     decompress the entire archive.  Also, this allows you to use different checksum and compression settings for each
'     node (e.g. not everything has to be compressed and/or checksummed - only the bits you find relevant).
'
'Here are a few things to note if you want to use this class in your own projects:
'
' 1) At present, pdPackage files are not easily editable.  Once created, there is no interface for adding new nodes,
'     erasing existing nodes, or modifying any of the pdPackage settings (compression, etc).  There's nothing in the format
'     design that prevents these actions, but I haven't written edit functions because I have no reason to do so in PhotoDemon.
'
' 2) As noted above, zLib is required for compression.  pdPackages are designed in a way that makes it easy to use any
'     compression (or other modification functions, e.g. encryption) functions of your choosing, or to ignore compression
'     entirely if you don't require it.  That said, if you want to use the class without zLib, but you intend to work with
'     pdPackage files from other sources (that may use compression), you need to make use of the getPackageFlag() function and
'     the accompanying PDP_FLAG_ZLIB_REQUIRED flag, to detect files that your software will not be able to process.
'
' 3) Up to 2GB of data is theoretically supported, but you won't be able to reach that amount from within VB.  For
'     performance reasons, this class creates the full archive in RAM before writing it to file.  This makes it very fast,
'     but ill-suited to extraordinarily large archive sizes.  (This is by design; sorry!)
'
' 4) When reading pdPackage files, the full original file contents will be cached in memory.  This allows you to load an
'     archive, then delete the original file without penalty.  Similarly, if you need access to an archive for an extended
'     period of time, this allows you to load it just once, then perform any operations from the in-memory copy alone.
'
' 5) This class has clsBasicBuffer as a dependency.  Thank you to dilettante for this minimal buffer class, which greatly
'     simplifies the process of reading/writing byte streams.  Note that I have modified the required clsBasicBuffer class
'     to operate directly on pointers (as it's much faster when working with large data chunks), so you cannot drop in
'     dilettante's original class and expect it to work.  You *must* use the modified version that comes with PhotoDemon.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This constant should be updated whenever the core assembly/disassembly code is modified.  This value is embedded in all
' pdPackage files written by this class, as a failsafe against potential ABI breakage in the future.
' - The current expected value is 66.  Version 66 added Unicode-friendly string handling, rather than the old StrConv methods.
' - Version 65 slightly modified the format of the directory chunk, by using a straight byte array instead of the actual UDT struct
'   used in previous versions (which VB custom-formats).  This was a necessary change prior to enabling second-pass compression of
'   the directory and data chunks.
' - The lowest expected value for this constant is 64, representative of PhotoDemon 6.4, when the PDI format was first implemented.
' - v64, 65, and 66 packages are still supported by this class, but it will only write v66 packages.
Private Const THIS_PDPACKAGE_VERSION As Long = 66
Private Const OLD_PDPACKAGE_VERSION_AUTUMN_2014 As Long = 65
Private Const OLD_PDPACKAGE_VERSION_SUMMER_2014 As Long = 64

'The first four bytes of a valid pdPackage file must always be &H4B504450 ("PDPK" in ASCII, if viewed in a text editor
' - note the little-endian declaration here).
Private Const PDP_UNIVERSAL_IDENTIFIER As Long = &H4B504450

'Because pdPackage is awesome, it can store some complicated data structures automatically, with no work on the caller's part.
' Structures like this are identified with special internal names, which simplifies the extraction process.
Private Const PDP_DATA_FILE As String = "PDPFILE"
Private Const PDP_DATA_FILEANDPATH As String = "PDPFILE_AND_PATH"

Private Enum PDP_SPECIAL_NODE_TYPE
    PDPSNT_FILENAME = 0
    PDPSNT_FILENAME_AND_RELATIVE_PATH = 1
End Enum

#If False Then
    Private Const PDPSNT_FILENAME = 0, PDPSNT_FILENAME_AND_RELATIVE_PATH = 1
#End If

'Each pdPackage file has a short file header.  This header is separate from the directory chunk, and it contains all information
' necessary to prepare a node directory array.
Private Type PDP_HEADER
    PDP_ID As Long                          'pdPackage Identifier; must always be &H4B504450 ("PDPK" in ASCII)
    
    PDP_SubID As Long                       'pdPackage sub-identifier.  This can be used by callers to define a specific type of pdPackage.
                                            ' (For example, PD uses this to denote PDI files specifically.)  This value has no intrinsic meaning.
    PDP_Version As Long                     'Version number of the pdPackage class used to write this file.  All parsing behavior is based on this value,
                                            ' and it is automatically set according to the THIS_PDPACKAGE_VERSION constant at the top of this class.
    NodeCount As Long                       'Number of data nodes in this package, 1-based (e.g. if there is one node in the archive, this value will
                                            ' be set to 1).  Cannot currently be zero, FYI.
    NodeStructSize As Long                  'Size of an individual node struct.  This can be used as a failsafe check against PDP_Version, above.
                                            ' Note that as of PDP version 65, this must be calculated using LenB() instead of Len().
    DirectoryChunkSize As Long              'Size of the full node directory structure, including all node directory entries.
    
    DirectoryFlags(0 To 2) As Long          'User-defined flags for the directory chunk.  See the PDP_FLAGS enum for details.
    
    DirectoryChunkSizeCompressed As Long    'If compression has been applied to the directory chunk, this value will be non-zero.  Use this to decompress
                                            ' the directory prior to copying it into the directory array.
    DataChunkSize As Long                   'Size of the data chunk of the archive.  This could be inferred by calculating EOF - (End of Directory Chunk),
                                            ' but it's easier to simply note it right inside the header.
    DataFlags(0 To 2) As Long               'User-defined flags for the data chunk.  See the PDP_FLAGS enum for details.
    
    DataChunkSizeCompressed As Long         'If compression has been applied to the data chunk, this value will be non-zero.  Use this to decompress
                                            ' the data chunk prior to copying it into the buffer class.
    Reserved As Long                        'Reserved for future use; no relevance at present.
    
End Type

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

' The list of currently supported flags is as follows:
Public Enum PDP_FLAGS
    PDP_FLAG_ZLIB_REQUIRED = 0    'There are one or more compressed entries in this chunk, so zLib will be required to read it
    PDP_FLAG_FILE_NODE = 1        'This node contains a file added via the shortcut autoAddNodeFromFile() function.
End Enum

#If False Then
    Private Const PDP_FLAG_ZLIB_REQUIRED = 0
    Private Const PDP_FLAG_FILE_NODE = 1
#End If

'Immediately following the PDP_HEADER is the directory chunk, which is comprised of FileHeader.NodeCount individual
' PDP_NODE structs.  These structs are small and flexible, and *they have a fixed size*, meaning they can be read into
' a fixed-width array in a single pass.
Private Type PDP_NODE

    NodeName As String * 32                 'Name of the node, as a standard VB String (DBCS)
    NodeID As Long                          'Alternatively, calling functions can specify an optional 4-byte numerical ID.  Nodes can be read by 32-char name or 4-byte ID.
    OptionalNodeType As Long                'Calling functions can also assign each node a TYPE if they want; this value has no meaning to the pdPackager class.
    NodeFlags(0 To 3) As Long               '16 bytes of node-specific flags are allowed.  At present, these are mostly unused.
    
    'One of the unique features of pdPackages is that each node is allotted two entries in the data chunk.  These entries don't have
    ' to be used; in fact, neither has to be used, but they can be helpful for reading node-specific information without having to
    ' decode the entire node contents.
    ' (Also, if I ever get around to implementing a file/folder wrapper for this class, the header chunk will be used to store
    '  relative path locations for each file in the package!)
    NodeHeaderOffset As Long                'Absolute offset of this node's header in the data chunk, STARTING FROM THE START OF THE DATA CHUNK, not the start of the file!
    NodeHeaderPackedSize As Long            'Packed size of this node's header.  (This is the size the node's header array occupies in the pdPackage data chunk.)
    NodeHeaderOriginalSize As Long          'Original size of this node's header.  (If this value is the same as NodeHeaderPackedSize, the node header was stored uncompressed.)
    NodeHeaderAdler32 As Long               'Adler32 checksum of the UNCOMPRESSED header bytes.  This can be used to verify the correctness of the data post-decompression.
    
    NodeDataOffset As Long                  'Absolute offset of this node's data in the data chunk, STARTING FROM THE START OF THE DATA CHUNK, not the start of the file!
    NodeDataPackedSize As Long              'Packed size of this node's data.  (This is the size the node's data array occupies in the pdPackage data chunk.)
    NodeDataOriginalSize As Long            'Original size of this node's data.  (If this value is the same as NodeHeaderPackedSize, the node data was stored uncompressed.)
    NodeDataAdler32 As Long                 'Adler32 checksum of the UNCOMPRESSED data bytes.  This can be used to verify the correctness of the data post-decompression.
    
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dstPointer As Long, ByVal srcPointer As Long, ByVal numOfBytes As Long)

'zLib compression and checksum validation
Private Declare Function compress Lib "zlibwapi.dll" (ByRef dstBuffer As Any, ByRef dstLen As Any, ByRef srcBuffer As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "zlibwapi.dll" (ByRef dstBuffer As Any, ByRef dstLen As Any, ByRef srcBuffer As Any, ByVal srcLen As Long, ByVal cmpLevel As Long) As Long
Private Declare Function uncompress Lib "zlibwapi.dll" (ByRef dstBuffer As Any, ByRef dstLen As Any, ByRef srcBuffer As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress_safe Lib "zlibwapi.dll" Alias "uncompress" (ByVal dstBufferPointer As Long, ByVal dstLenPointer As Long, ByVal srcBufferPointer As Long, ByVal srcLenNotAPointer As Long) As Long
Private Declare Function adler32 Lib "zlibwapi.dll" (ByVal adlerStart As Long, ByRef srcBuffer As Any, ByVal srcLength As Long) As Long
Private Declare Function crc32 Lib "zlibwapi.dll" (ByVal crcStart As Long, ByRef srcBuffer As Any, ByVal srcLength As Long) As Long
Private Declare Function zlibVersion Lib "zlibwapi.dll" () As Long

'API calls for explicitly calling dlls.  This allows us to build DLL paths at runtime, and it also allows
' us to call any DLL we like without first passing them through regsvr32.
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'If zLib has been successfully initialized, this will be set to TRUE.  If we initialized the DLL within this class, we will also
' have a zLibToken value; this value must be used at termination time to free our zLib library handle.
Private m_ZLibAvailable As Boolean, m_ZLibToken As Long

'When writing new pdPackage files, these variables will hold the package's contents as they are being assembled.
Private m_FileHeader As PDP_HEADER
Private m_NodeDirectory() As PDP_NODE
Private m_NumOfNodes As Long

'The actual data chunk of the pdPackage will be assembled using a clsBasicBuffer instance.  It greatly simplifies the process of
' assembling a 1D byte array from discrete individual chunks.
Private m_DataBuffer As clsBasicBuffer

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
Public Sub prepareNewPackage(Optional ByVal numOfDataNodes As Long = 0, Optional ByVal optPackageID As Long = 0, Optional ByVal estimatedDataChunkSize As Long = 0)

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
        .DirectoryChunkSize = .NodeStructSize * numOfDataNodes
        .DirectoryChunkSizeCompressed = 0
        
        'DirectoryFlags() and DataFlags() are set by separate functions.  For now, assume a value of 0 for all flags.
        Dim i As Long
        For i = 0 To UBound(.DirectoryFlags)
            .DirectoryFlags(i) = 0
            .DataFlags(i) = 0
        Next i
        
        'DataChunkSize will remain unknown until all nodes have been added.
        .DataChunkSize = 0
        .DataChunkSizeCompressed = 0
        
        '4 bytes are reserved at the end as a "just in case".  For now, they should always be 0.
        .Reserved = 0
        
    End With
    
    'Resize the directory array to the number of supplied nodes; note that this step is optional; if no node count is supplied,
    ' the addNode() function will automatically increment itself as necessary.
    m_NumOfNodes = 0
    ReDim m_NodeDirectory(0 To numOfDataNodes) As PDP_NODE
        
    'Prepare the data buffer.  Note that chunk size is largely irrelevant for our purposes; when handed a byte array that exceeds
    ' the size of the default chunk size, the buffer class is smart enough to extend the buffer by the size of that byte array.
    ' If writing a lot of small data, however, the chunk size becomes important; try to make it large enough that you will only
    ' require a handful of ReDim Preserve statements throughout the pdPackage assembly process.
    Set m_DataBuffer = New clsBasicBuffer
    m_DataBuffer.ChunkSize = 32768
    m_DataBuffer.OpenBuf estimatedDataChunkSize

End Sub

'Add a new node to this pdPackage instance.  Note that this function DOES NOT actually add the node's data arrays to the main
' data buffer - those are done in a subsequent step, by the user, as necessary.  (It's a little pesky to separate node additions
' into multiple steps, but this allows for more fine-grained control over node addition, without overwhelming any function with
' a monstrous list of parameters.)
'
'Importantly, this function returns the index of the added node, which external functions must then use to supply this node's
' actual data arrays.
Public Function addNode(Optional ByVal thisNodeName As String = "", Optional ByVal thisNodeID As Long = 0, Optional ByVal thisNodeType As Long = 0) As Long

    'Increment our active node count
    m_NumOfNodes = m_NumOfNodes + 1
    
    'Start by making sure our node directory is large enough to hold this new node (if the caller supplied a node count up front,
    ' this step is overkill).  If the directory is too small, enlarge it.
    If UBound(m_NodeDirectory) < (m_NumOfNodes - 1) Then
        
        'If the array has never been resized before, give it a nice starting size of 16 entries
        If UBound(m_NodeDirectory) = 0 Then
            ReDim Preserve m_NodeDirectory(0 To 15) As PDP_NODE
        
        'If the directory has been resized before, double its size now.  (The directory will automatically be shrunk to minimum
        ' size when it's written out to file, so don't worry about excess space being allocated.)
        Else
            ReDim Preserve m_NodeDirectory(0 To UBound(m_NodeDirectory) * 2 + 1) As PDP_NODE
        End If
        
    End If
    
    'Copy the supplied values into the node directory.  Note that all three ID types are optional, but hopefully the user has
    ' been smart enough to make use of at least one of them!
    With m_NodeDirectory(m_NumOfNodes - 1)
        .NodeName = thisNodeName
        .NodeID = thisNodeID
        .OptionalNodeType = thisNodeType
    End With
    
    'Return the index of this node, which the caller will use to supply this node's data (in a separate step).
    addNode = m_NumOfNodes - 1
    
End Function

'Add data for a given node.  Four required params, three optional params.  The function will return TRUE if successful.
'
' The four required params are:
' 1) Index of the target node.  This is the value returned by addNode(), above.
' 2) Destination chunk for this data.  Remember that each node in a pdPackage supports TWO possible data arrays, which can be
'    utilized however the caller pleases.  (Typically, one is used for lightweight header data, while the other is used for
'    heavy content data.)
' 3, 4) Pointer and size of the byte array containing the data the caller wants written.  This class doesn't care how the
'    pointer is obtained or assembled, so long as the SIZE IS BYTE-ACCURATE.
'
' The three optional params are:
' 1) Whether to compress the data before writing it.  If zLib is unavailable, this param has no meaning, as data will always
'    be written without compression.
' 2) Compression level.  Per the zLib spec, this is a value from 0 (uncompressed) to 9 (best possible compression).  Note that
'    a value of -1 means "use default compression level", and a value of 1 means "compress the data, but do it as quickly as
'    possible".  I don't recommend requesting compression, then using compression level 0 - that would be stupid!
' 3) Whether to calculate a checksum for the uncompressed data bytes.  This is completely optional.  If a checksum is found in
'    the file, pdPackage will automatically process it at load-time, and reject the file if checksums do not match.  Note that
'    checksumming the data increases write time non-trivially.
Public Function addNodeDataFromPointer(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByVal dataPointer As Long, ByVal dataLength As Long, Optional ByVal compressData As Boolean = True, Optional ByVal compressionLevel As Long = -1, Optional ByVal requestChecksum As Boolean = False) As Boolean

    'Start by validating the node index we were passed.  If it's out of range, exit immediately.
    If nodeIndex < 0 Or nodeIndex > m_NumOfNodes - 1 Then
        Debug.Print "Node index out of range - try again with a valid node index!"
        addNodeDataFromPointer = False
        Exit Function
    End If
    
    'If compression is requested, update the "zLib required" flag for both the file as a whole, and this specific node.
    If compressData Then
    
        'Overall data chunk flag
        setBitFlag PDP_FLAG_ZLIB_REQUIRED, True, m_FileHeader.DataFlags(0)
        
        'This node's flag
        setBitFlag PDP_FLAG_ZLIB_REQUIRED, True, m_NodeDirectory(nodeIndex).NodeFlags(0)
        
    End If
    
    'Update the pre-compression values for this data chunk.
    If useHeaderBuffer Then
        m_NodeDirectory(nodeIndex).NodeHeaderOriginalSize = dataLength
    Else
        m_NodeDirectory(nodeIndex).NodeDataOriginalSize = dataLength
    End If
    
    'Mark the offset for this data chunk.
    If useHeaderBuffer Then
        m_NodeDirectory(nodeIndex).NodeHeaderOffset = m_DataBuffer.Position
    Else
        m_NodeDirectory(nodeIndex).NodeDataOffset = m_DataBuffer.Position
    End If
    
    'If we are not using compression, we can write the node data as-is.  Note that header or data doesn't matter here,
    ' because the data is simply added to the pdPackages data chunk.  Order of writing is irrelevant.
    If (Not compressData) Or (Not m_ZLibAvailable) Then
    
        m_DataBuffer.WriteBytesFromPointer dataPointer, dataLength
        
        If useHeaderBuffer Then
            m_NodeDirectory(nodeIndex).NodeHeaderPackedSize = dataLength
        Else
            m_NodeDirectory(nodeIndex).NodeDataPackedSize = dataLength
        End If
        
    'Data compression was requested, meaning we have a bit of extra work to do.
    Else
        
        Dim compressedData() As Byte
        Dim compressedSize As Long
        
        'copyPtrToArray will return TRUE if the data was compressed; FALSE if it was not.  There is no error state for the function.
        If copyPtrToArray(dataPointer, dataLength, compressedData, compressedSize, True, compressionLevel, False) Then
            
            'For fun, supply some debug info on the compression.
            Debug.Print "Node data compressed successfully; compressed data is " & Format(CStr(100 - (100 * (CDbl(compressedSize) / CDbl(dataLength)))), "#0.00") & "% smaller (" & dataLength & " to " & compressedSize & " bytes)."
            
        Else
            
            'Something went horribly wrong.  Write the uncompressed original data to the buffer, and provide a debug warning.
            Debug.Print "zLib compression failed; writing original uncompressed data to buffer instead..."
        
        End If
        
        'Copy the contents of the compressedData() array into the buffer; note that even if zLib failed, copyPtrToArray will fill compressedData()
        ' with an uncompressed copy of the source bytes (and compressedSize with the source buffer size), so we can use the same code regardless.
        m_DataBuffer.WriteBytesFromPointer VarPtr(compressedData(0)), compressedSize
        
        If useHeaderBuffer Then
            m_NodeDirectory(nodeIndex).NodeHeaderPackedSize = compressedSize
        Else
            m_NodeDirectory(nodeIndex).NodeDataPackedSize = compressedSize
        End If
        
    End If
    
    'Next, checksum the file (if requested).
    If requestChecksum And m_ZLibAvailable Then
    
        Dim zLibAdlerSum As Long
        zLibAdlerSum = checkSumArbitraryData(dataPointer, dataLength)
        
        If useHeaderBuffer Then
            m_NodeDirectory(nodeIndex).NodeHeaderAdler32 = zLibAdlerSum
        Else
            m_NodeDirectory(nodeIndex).NodeDataAdler32 = zLibAdlerSum
        End If
    
    End If
    
    'This chunk was added successfully!  Return TRUE and exit.
    addNodeDataFromPointer = True
    
End Function

'Thin wrapper to addNodeDataFromPointer, above; if possible, use that function directly to avoid any unnecessary copying of arrays
Public Function addNodeDataFromByteArray(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByRef DataBytes() As Byte, Optional ByVal compressData As Boolean = True, Optional ByVal compressionLevel As Long = -1, Optional ByVal requestChecksum As Boolean = False) As Boolean

    addNodeDataFromByteArray = addNodeDataFromPointer(nodeIndex, useHeaderBuffer, VarPtr(DataBytes(0)), UBound(DataBytes) + 1, compressData, compressionLevel, requestChecksum)
    
End Function

'Thin wrapper for addNodeDataFromPointer, above, but allows the user to supply a string.
Public Function addNodeDataFromString(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByRef srcDataString As String, Optional ByVal compressData As Boolean = True, Optional ByVal compressionLevel As Long = -1, Optional ByVal requestChecksum As Boolean = False) As Boolean
    
    'Forward the string pointer to the actual addNodeData function, and return its result.
    If Len(srcDataString) <> 0 Then
        addNodeDataFromString = addNodeDataFromPointer(nodeIndex, useHeaderBuffer, StrPtr(srcDataString), Len(srcDataString) * 2, compressData, compressionLevel, requestChecksum)
    End If

End Function

'Shortcut function to both create a new node, and copy a file's contents directly into that node, without any intermediate work on the caller's part.
' Note that only the filename and the file's contents are added; things like file attributes or security descriptors are *not*, by design.
'
'Default compression settings are used for the added files (e.g. compression is always applied, at zLib's recommended setting).
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
Public Function autoAddNodeFromFile(ByVal pathToFile As String, Optional ByVal OptionalNodeType As Long = 0, Optional ByVal addChecksumVerification As Boolean = True, Optional ByVal storeUsingThisRelativePathAndFilename As String = "") As Boolean

    'Attempt to load the file into a byte array
    Dim fileContents() As Byte
    If m_File.LoadFileAsByteArray(pathToFile, fileContents) Then
        
        'Make sure the file loaded successfully
        If UBound(fileContents) >= LBound(fileContents) Then
            
            'Create a blank node
            Dim nodeIndex As Long
            
            'Files with relative paths in their names are specially flagged, because we need to reconstruct their folder hierarchy
            ' when extracting them.
            Dim filenameIncludesRelativePath As Boolean
            filenameIncludesRelativePath = (Len(storeUsingThisRelativePathAndFilename) > 0)
            
            If filenameIncludesRelativePath Then
                nodeIndex = addNode(PDP_DATA_FILEANDPATH, , OptionalNodeType)
            Else
                nodeIndex = addNode(PDP_DATA_FILE, , OptionalNodeType)
            End If
            
            'Write the filename out to the first node.  If the caller specifically specified that the filename includes a relative path,
            ' we will write the whole filename, untouched
            If filenameIncludesRelativePath Then
                autoAddNodeFromFile = addNodeDataFromString(nodeIndex, True, storeUsingThisRelativePathAndFilename, False)
            Else
                autoAddNodeFromFile = addNodeDataFromString(nodeIndex, True, m_File.GetFilename(pathToFile), False)
            End If
            
            'Write the file contents out to the second node
            autoAddNodeFromFile = autoAddNodeFromFile And addNodeDataFromByteArray(nodeIndex, False, fileContents, , , addChecksumVerification)
            
            'Mark a special flag bit in this node, to identify it as an auto-added file entry.  This simplifies the extraction process later,
            ' and also provides a failsafe against the user adding their own nodes called "PDPFILE".
            setBitFlag PDP_FLAG_FILE_NODE, True, m_NodeDirectory(nodeIndex).NodeFlags(0)
            
        End If
    
    Else
        Debug.Print "WARNING! pdPackage.autoAddNodeFromFile could not load " & pathToFile & ". Abandoning request."
        autoAddNodeFromFile = False
    End If

End Function

'Shortcut function to create an arbitrary number of nodes, using some source folder as the reference.
'
'This function is basically a thin wrapper to pdFSO.retrieveAllFiles and autoAddNodeFromFile.  You could achieve the same results by using
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
'Default compression settings are assumed for all added files.  Per-file checksum verifications can be toggled, but it's strongly recommended
' to leave it at its default setting (ON), as a failsafe against package corruption.
'
'Finally, whitelisting or blacklisting behavior is available via the optional onlyAllowTheseExtensions and doNotAllowTheseExtensions parameters.
' These two parameters are passed, untouched, to an underlying pdFSO instance, so please refer to pdFSO for details on how to use these.
'
'Returns: TRUE if all files are added successfully, FALSE otherwise.  FALSE should not occur unless read access to target folder(s) is restricted,
' or if memory runs out because the constructed package exceeds available memory.
Public Function autoAddNodesFromFolder(ByRef srcFolder As String, Optional ByVal OptionalNodeType As Long = 0, Optional ByVal recurseSubfolders As Boolean = True, Optional ByVal preserveFolders As Boolean = True, Optional ByVal useCustomRelativeFolderBase As String = "", Optional ByVal addChecksumVerification As Boolean = True, Optional ByVal onlyAllowTheseExtensions As String = "", Optional ByVal doNotAllowTheseExtensions As String = "") As Boolean

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
                If Len(useCustomRelativeFolderBase) <> 0 Then
                    relativeFilePath = m_File.GenerateRelativePath(useCustomRelativeFolderBase, curFile)
                Else
                    relativeFilePath = m_File.GenerateRelativePath(srcFolder, curFile)
                End If
                
            'Folder preservation is inactive, so do not specify any relative paths whatsoever
            Else
                relativeFilePath = ""
            End If
            
            'Add this node
            nodeAdditionSuccess = nodeAdditionSuccess And autoAddNodeFromFile(curFile, OptionalNodeType, addChecksumVerification, relativeFilePath)
            
        Loop
        
        autoAddNodesFromFolder = nodeAdditionSuccess
        
        'Display some debug info if one or more nodes failed to add
        If Not nodeAdditionSuccess Then
            
            If recurseSubfolders Then
                Debug.Print "WARNING! pdPackager.autoAddNodesFromFolder() failed to add one or more nodes.  Do you have read access to all subfolders??"
            Else
                Debug.Print "WARNING! pdPackager.autoAddNodesFromFolder() failed to add one or more nodes.  Please investigate."
            End If
        
        End If
        
    Else
        
        Debug.Print "WARNING! pdPackager.autoAddNodesFromFolder() could not retrieve a valid list of files.  Do you have read access to the folder??"
        autoAddNodesFromFolder = False
        
    End If

End Function

'Shortcut function to create an arbitrary number of nodes, using a pdStringStack full of files and/or folders as the reference.
' (Obviously, it's up to the caller to populate the string stack.)
'
'This function is basically a thin wrapper to autoAddNodeFromFile and autoAddNodesFromFolder.  You could achieve identical results by using
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
'Default compression settings are assumed for all added files.  Per-file checksum verifications can be toggled, but it's strongly recommended
' to leave it at its default setting (ON), as a failsafe against package corruption.
'
'Returns: TRUE if all files are added successfully, FALSE otherwise.  FALSE should not occur unless read access to target file(s) and/or folder(s)
' is restricted, or if memory runs out because the constructed package exceeds available memory.
Public Function autoAddNodesFromStringStack(ByRef srcStringStack As pdStringStack, Optional ByVal relativeFolderBase As String = "", Optional ByVal OptionalNodeType As Long = 0, Optional ByVal recurseSubfolders As Boolean = True, Optional ByVal addChecksumVerification As Boolean = True) As Boolean
    
    Dim autoAddSuccess As Boolean
    autoAddSuccess = True
    
    'Prior to starting, mark whether relative folders are in use.  If they are not, we can bypass some steps within the main Do loop.
    Dim relativeFoldersInUse As Boolean
    relativeFoldersInUse = (Len(relativeFolderBase) <> 0)
    
    Dim relativePath As String
    
    'This function relies on autoAddNodesFromFolder and autoAddNodeFromFile to do all the heavy lifting.  All we do is pop items off the
    ' string stack, and forward them to one of those two functions.
    Dim curFile As String
    
    Do While srcStringStack.PopString(curFile)
    
        'See if this is a file or a folder
        If m_File.FileExist(curFile) Then
        
            'This is a file.  Add it via the file function
            If relativeFoldersInUse Then
            
                'Construct a relative path
                relativePath = m_File.GenerateRelativePath(relativeFolderBase, curFile)
                autoAddSuccess = autoAddSuccess And autoAddNodeFromFile(curFile, OptionalNodeType, addChecksumVerification, relativePath)
                
            Else
                autoAddSuccess = autoAddSuccess And autoAddNodeFromFile(curFile, OptionalNodeType, addChecksumVerification)
            End If
        
        ElseIf m_File.FolderExist(curFile, False) Then
        
            'This is a folder.  Add it via the folder function.
            If relativeFoldersInUse Then
                autoAddSuccess = autoAddSuccess And autoAddNodesFromFolder(curFile, OptionalNodeType, recurseSubfolders, True, relativeFolderBase, addChecksumVerification)
            Else
                autoAddSuccess = autoAddSuccess And autoAddNodesFromFolder(curFile, OptionalNodeType, recurseSubfolders, False, , addChecksumVerification)
            End If
        
        Else
            Debug.Print "WARNING!  pdPackager.autoAddNodesFromStringStack() was handed something that isn't a file or folder (" & curFile & ")."
        End If
    
    Loop
    
    autoAddNodesFromStringStack = autoAddSuccess
    
End Function

'When all nodes have been successfully added, the user can finally write out their data to file.  This function will return TRUE
' if successful, FALSE if unsuccessful.  (Note that it assumes the caller has done some basic validation on the file path, like
' obtaining permission from the user to overwrite an existing file.)
Public Function writePackageToFile(ByVal dstFilename As String, Optional ByVal secondPassDirectoryCompression As Boolean = False, Optional ByVal secondPassDataCompression As Boolean = False, Optional ByVal thisIsATempFile As Boolean = False) As Boolean
    
    'Start by updating the file header.  Most of this will have been done when the class was initialized, but some values
    ' can't be set until all nodes have been added.
    With m_FileHeader
    
        'Update the final node count
        .NodeCount = m_NumOfNodes
                
        'Update the size of the directory chunk
        .DirectoryChunkSize = .NodeStructSize * m_NumOfNodes
        
        'Trim the data buffer, and update the size of the data chunk using the exact size of the final buffer.
        m_DataBuffer.ShrinkBuf
        .DataChunkSize = m_DataBuffer.Size + 1
        
    End With
    
    'Before writing the file, we have a few other potential header items we need to construct.
    
    'Create the node directory byte array now.  (VB writes UDTs to file in a non-standard, non-obvious way, so we need to do this
    ' regardless of second-pass directory compression.)
    Dim rawDirectoryBuffer() As Byte, rawDataBuffer() As Byte
    Dim preCompressionBuffer() As Byte
    Dim originalSize As Long, compressedSize As Long
    originalSize = m_FileHeader.DirectoryChunkSize
    
    'Use the copyPtrToArray function to copy the contents of m_NodeDirectory into rawDirectoryBuffer().  The function will return TRUE
    ' if it compresses the data; false, otherwise.
    If copyPtrToArray(VarPtr(m_NodeDirectory(0)), originalSize, rawDirectoryBuffer, compressedSize, secondPassDirectoryCompression, , True) Then
        m_FileHeader.DirectoryChunkSizeCompressed = compressedSize
        setBitFlag PDP_FLAG_ZLIB_REQUIRED, True, m_FileHeader.DirectoryFlags(0)
    Else
        m_FileHeader.DirectoryChunkSizeCompressed = 0
    End If
    
    'Repeat the above steps for the data chunk
    originalSize = m_FileHeader.DataChunkSize
    If copyPtrToArray(m_DataBuffer.CurrentBytesPointer, originalSize, rawDataBuffer, compressedSize, secondPassDataCompression, , True) Then
        m_FileHeader.DataChunkSizeCompressed = compressedSize
        setBitFlag PDP_FLAG_ZLIB_REQUIRED, True, m_FileHeader.DataFlags(0)
    Else
        m_FileHeader.DataChunkSizeCompressed = 0
    End If
    
    'Kill the destination file if it already exists
    If m_File.FileExist(dstFilename) Then m_File.KillFile dstFilename
    
    'Retrieve a writable handle
    Dim hFile As Long, createSuccess As Boolean
    
    If thisIsATempFile Then
        createSuccess = m_File.CreateFileHandle(dstFilename, hFile, True, True, OptimizeTempFile)
    Else
        createSuccess = m_File.CreateFileHandle(dstFilename, hFile, True, True, OptimizeSequentialAccess)
    End If
    
    If createSuccess Then
    
        'Writing out the three crucial pieces of data: the header, the directory, and the data
        With m_File
            .WriteDataToFile hFile, VarPtr(m_FileHeader), LenB(m_FileHeader)
            .WriteDataToFile hFile, VarPtr(rawDirectoryBuffer(0)), UBound(rawDirectoryBuffer) + 1
            .WriteDataToFile hFile, VarPtr(rawDataBuffer(0)), UBound(rawDataBuffer) + 1
        End With
        
        m_File.CloseFileHandle hFile
        
        writePackageToFile = True
    
    Else
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdPackager.writePackageToFile failed to create a valid destination file handle.  Package was not written."
        #End If
        
        writePackageToFile = False
    
    End If
        
    'TODO!  Add basic error handling.

End Function

'Copy one array into another, with variable compression.  This is used in many places throughout this class; basically any action that
' potentially supports compression should use this function to transfer data between arrays, even when compression is not in use.
' (This function abstracts away all the messy CopyMemory bits, keeping the rest of the class clean.)
'
'If compression *is* in use, this function will use zLib to compress the array accordingly, and it will fill the passed dstSize Long
' with the size (1-based) of the compressed array.  Explicit size values are required for both source and destination arrays, which
' allows both the caller and this function to avoid the need for precisely dimensioned arrays.  (Any time we can skip an unnecessary
' ReDim Preserve, we save ourselves precious processing time.)
'
'This function is designed to be error-proof.  If something goes wrong during the compression stage, it will do a straight copy from
' the source to the destination.  Similarly, if the compressed size is larger than the uncompressed size, it will also return an
' uncompressed copy of the data.  The returned BOOLEAN reports whether or not compression was actually used, NOT whether or not the
' function was successful.  (It is assumed the function is always successful, because if zLib fails, the default uncompressed
' array will still be returned, and the caller can proceed normally with that data.)
Private Function copyPtrToArray(ByVal srcPointer As Long, ByVal srcSize As Long, ByRef dstArray() As Byte, ByRef dstSize As Long, Optional ByVal useCompression As Boolean = False, Optional ByVal compressionLevel As Long = -1, Optional ByVal trimDestinationArray As Boolean = False) As Boolean

    'If compression is active, extra steps are (obviously) required.
    If useCompression Then
    
        'Start by preparing the destination array to receive the zLib-compressed data.  The zLib creators recommend a buffer 1% larger than
        ' the input buffer, plus 12 extra bytes for a header.  This is the "worst possible scenario" for zLib compression (e.g. if the
        ' input buffer is 100% uncompressable.)
        dstSize = srcSize + (CDbl(srcSize) * 0.01) + 12
        ReDim dstArray(0 To dstSize - 1) As Byte
        
        'Request compression from zLib.
        Dim zLibResult As Long
        If compressionLevel >= 0 Then
            zLibResult = compress2(dstArray(0), dstSize, ByVal srcPointer, srcSize, compressionLevel)
        Else
            zLibResult = compress(dstArray(0), dstSize, ByVal srcPointer, srcSize)
        End If
        
        'Make sure zLib compressed the data successfully.
        If zLibResult = 0 Then
        
            'ReDim the destination array, if requested
            If trimDestinationArray Then ReDim Preserve dstArray(0 To dstSize - 1) As Byte
            copyPtrToArray = True
        
        'zLib failed; write an uncompressed directory instead.
        Else
        
            dstSize = srcSize
            ReDim dstArray(0 To dstSize - 1) As Byte
            CopyMemory VarPtr(dstArray(0)), srcPointer, srcSize
            copyPtrToArray = False
        
        End If
        
    Else
    
        dstSize = srcSize
        ReDim dstArray(0 To dstSize - 1) As Byte
        CopyMemory VarPtr(dstArray(0)), srcPointer, srcSize
        copyPtrToArray = False
    
    End If
    
End Function

'This wrapper for copyPtrToArray, above, that allows the user to specify an array reference instead of a pointer
Private Function copyArray(ByRef srcArray() As Byte, ByVal srcSize As Long, ByRef dstArray() As Byte, ByRef dstSize As Long, Optional ByVal useCompression As Boolean = False, Optional ByVal compressionLevel As Long = -1, Optional ByVal trimDestinationArray As Boolean = False) As Boolean
    copyArray = copyPtrToArray(VarPtr(srcArray(0)), srcSize, dstArray, dstSize, useCompression, compressionLevel, trimDestinationArray)
End Function

'Load a pdPackage file into memory.  A few things to note:
' - For performance reasons, this class caches the entire file contents in RAM.  It does this so that individual nodes can be
'    quickly extracted without touching the hard drive.  If your particular pdPackage file is enormous, make sure you have enough
'    memory to cache it!
' - The caller is expected to handle detailed testing of the supplied path (e.g. read access, etc).  This function does only
'    minimal error checking.
'
'If the file can be successfully loaded, this function returns TRUE.
'
'Parameters include:
' 1) Source file.  (Again, callers must do their own validation on this path.)
' 2) Optionally, a sub-type value you want to validate.  If this value is not found in bytes 4-7 of the file header, the file
'     will be rejected.  (If you didn't request a specific sub-type at pdPackage creation time, don't enable this check!)
Public Function readPackageFromFile(ByVal srcFilename As String, Optional ByVal subTypeValidator As Long = 0) As Boolean

    On Error GoTo StopPackageFileRead
    
    Dim rawBuffer() As Byte, tmpCompressionBuffer() As Byte
    Dim tmpLength As Long
    Dim zLibResult As Long
    
    'Before doing anything else, make sure the file exists
    If Not m_File.FileExist(srcFilename) Then
        Debug.Print "Requested pdPackage file doesn't exist.  Validate all paths before sending them to the pdPackage class!"
        GoTo StopPackageFileRead
    End If
    
    'Open the file
    Dim hFile As Long
    If m_File.CreateFileHandle(srcFilename, hFile, True, False, OptimizeSequentialAccess) Then
        
        'Attempt to load the file header.  If the file is invalid, this may potentially be gibberish, but we won't know until we try!
        m_File.ReadDataFromFile hFile, VarPtr(m_FileHeader), LenB(m_FileHeader)
        
        'Validate the file signature.  Valid pdPackage files must have their first 4 bytes set to "PDPK" in ASCII.
        If (m_FileHeader.PDP_ID <> PDP_UNIVERSAL_IDENTIFIER) Then
            raiseInternalMessage "File doesn't have valid pdPackage header.  Abandoning load."
            GoTo StopPackageFileRead
        End If
        
        'If the caller requested a sub-type validation, perform that now
        If subTypeValidator <> 0 Then
            
            If (m_FileHeader.PDP_SubID <> subTypeValidator) Then
                raiseInternalMessage "File doesn't match requested sub-type value.  Abandoning load."
                GoTo StopPackageFileRead
            End If
        
        End If
        
        'Finally, check the PDP version.  If it's older than the current version, loading will be handled by one of our legacy load functions.
        If m_FileHeader.PDP_Version < THIS_PDPACKAGE_VERSION Then
            
            'Close our handle prior to continuing
            m_File.CloseFileHandle hFile
            
            'Let the legacy load function have a go at the data
            If Not loadLegacyPDPackage(srcFilename) Then
                raiseInternalMessage "Legacy file format could not be decoded."
                GoTo StopPackageFileRead
            End If
        
        'This is the current PDP version.  Continue the normal loading process.
        Else
            
            'Make sure the struct size is correct; if it isn't, all subsequent parsing will fail.
            Dim tmpNode As PDP_NODE
            If m_FileHeader.NodeStructSize <> LenB(tmpNode) Then
                raiseInternalMessage "Node struct size in header is invalid.  This file looks to be corrupt!  Abandoning load."
                GoTo StopPackageFileRead
            End If
            
            'Next, it's time to retrieve the directory.  The size of the directory varies, based on whether it received second-pass compression.
            If m_FileHeader.DirectoryChunkSizeCompressed = 0 Then
            
                'No compression applied, so we can retrieve the node directory as-is.
                ReDim rawBuffer(0 To m_FileHeader.DirectoryChunkSize - 1) As Byte
                m_File.ReadDataFromFile hFile, VarPtr(rawBuffer(0)), m_FileHeader.DirectoryChunkSize
                
            Else
                
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Second-pass compression was applied.  Decompressing node directory now..."
                #End If
                
                'Second-pass compression was applied.  Make sure zLib is available.  (If it isn't, the user is screwed.)
                If m_ZLibAvailable Then
                
                    'Resize the destination buffer to the original size of the data
                    ReDim rawBuffer(0 To m_FileHeader.DirectoryChunkSize - 1) As Byte
                    
                    'Resize the source buffer to the compressed directory size, and retrieve the compressed directory.
                    ReDim tmpCompressionBuffer(0 To m_FileHeader.DirectoryChunkSizeCompressed - 1) As Byte
                    m_File.ReadDataFromFile hFile, VarPtr(tmpCompressionBuffer(0)), m_FileHeader.DirectoryChunkSizeCompressed
                    
                    'Use zLib to decompress the data
                    tmpLength = m_FileHeader.DirectoryChunkSize
                    zLibResult = uncompress_safe(VarPtr(rawBuffer(0)), VarPtr(tmpLength), VarPtr(tmpCompressionBuffer(0)), m_FileHeader.DirectoryChunkSizeCompressed)
                    
                    'Make sure the decompression size is what we expected
                    If tmpLength <> m_FileHeader.DirectoryChunkSize Then
                        raiseInternalMessage "WARNING!  Decompressed directory chunk size shows possible CRC error (" & CStr(tmpLength) & " vs " & CStr(m_FileHeader.DirectoryChunkSize) & ")"
                    End If
                    
                    'Free the temporary compression buffer
                    'Erase tmpCompressionBuffer
                    
                    If zLibResult <> 0 Then
                        
                        raiseInternalMessage "Directory is compressed, but zLib failed to decompress it.  Error code returned: %1", zLibResult
                        GoTo StopPackageFileRead
                    
                    End If
                    
                Else
                
                    raiseInternalMessage "Directory chunk is compressed, but zLib is missing.  Decompression impossible; abandoning load."
                    GoTo StopPackageFileRead
                
                End If
            
            End If
            
            'Use information from the file header to prepare the directory array, then copy the raw buffer into it.
            m_NumOfNodes = m_FileHeader.NodeCount
            ReDim m_NodeDirectory(0 To m_NumOfNodes - 1) As PDP_NODE
            CopyMemory VarPtr(m_NodeDirectory(0)), VarPtr(rawBuffer(0)), m_FileHeader.DirectoryChunkSize
            
            'Now, retrieve the data chunk, and apply the same decompression test that we did for the directory chunk.
            If m_FileHeader.DataChunkSizeCompressed = 0 Then
            
                'No compression applied, so we can retrieve the node directory as-is.
                ReDim rawBuffer(0 To m_FileHeader.DataChunkSize - 1) As Byte
                m_File.ReadDataFromFile hFile, VarPtr(rawBuffer(0)), m_FileHeader.DataChunkSize
                
            Else
                
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Second-pass compression was applied.  Decompressing node data now..."
                #End If
                
                'Second-pass compression was applied.  Make sure zLib is available.  (If it isn't, the user is screwed.)
                If m_ZLibAvailable Then
                
                    'Resize the destination buffer to the original size of the data
                    ReDim rawBuffer(0 To m_FileHeader.DataChunkSize - 1) As Byte
                    
                    'Resize the source buffer to the compressed data chunk size, and retrieve the compressed chunk.
                    ReDim tmpCompressionBuffer(0 To m_FileHeader.DataChunkSizeCompressed - 1) As Byte
                    m_File.ReadDataFromFile hFile, VarPtr(tmpCompressionBuffer(0)), m_FileHeader.DataChunkSizeCompressed
                    
                    'Use zLib to decompress the data
                    tmpLength = m_FileHeader.DataChunkSize
                    zLibResult = uncompress_safe(VarPtr(rawBuffer(0)), VarPtr(tmpLength), VarPtr(tmpCompressionBuffer(0)), m_FileHeader.DataChunkSizeCompressed)
                    
                    'Make sure the decompression size is what we expected
                    If tmpLength <> m_FileHeader.DataChunkSize Then
                        raiseInternalMessage "WARNING!  Decompressed directory chunk size shows possible CRC error (" & CStr(tmpLength) & " vs " & CStr(m_FileHeader.DataChunkSize) & ")"
                    End If
                    
                    'Free our temporary buffer
                    'Erase tmpCompressionBuffer
                    
                    If zLibResult <> 0 Then
                        
                        raiseInternalMessage "Data chunk is compressed, but zLib failed to decompress it.  Error code returned: %1", zLibResult
                        GoTo StopPackageFileRead
                    
                    End If
                    
                Else
                
                    raiseInternalMessage "Data chunk is compressed, but zLib is missing.  Decompression impossible; abandoning load."
                    GoTo StopPackageFileRead
                
                End If
            
            End If
            
            'Pass the data chunk to the m_DataBuffer class (which we'll use for further parsing).
            Set m_DataBuffer = New clsBasicBuffer
            m_DataBuffer.WriteBytes rawBuffer, UBound(rawBuffer) + 1
            m_DataBuffer.Position = 0
            
        End If
        
        'Close the handle when finished, obviously!
        If hFile <> 0 Then m_File.CloseFileHandle hFile
        
        'File load complete!
        readPackageFromFile = True
        
    Else
        readPackageFromFile = False
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdPackager.readPackageFromFile failed to create a valid handle for " & srcFilename & ".  Read abandoned."
        #End If
    End If
    
    Exit Function
    
StopPackageFileRead:
    
    Debug.Print "An error occurred in the pdPackager.readPackageFromFile function.  Additional data should have been supplied by the Message() function."
    If hFile <> 0 Then m_File.CloseFileHandle hFile
    readPackageFromFile = False
    Exit Function

End Function

'Any PDPackage versions prior to the current version (THIS_PDPACKAGE_VERSION) are loaded via this function.  This function does not use
' Unicode-aware parsing code, FYI!
Private Function loadLegacyPDPackage(ByRef srcFile As String) As Boolean
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Legacy pdPackage format detected.  Attempting decoding now..."
    #End If
    
    'Open the file
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open srcFile For Binary As #fileNum
    
        'Load the full file header.
        Get #fileNum, 1, m_FileHeader
        
        'Separate handling based on the file header, which has already been stored to m_FileHeader
        Select Case m_FileHeader.PDP_Version
        
            'The original pdPackage format is very simple.  I had not intended to replace it, but at the time I created it, I didn't
            ' realize that VB stored UDTs to file in a non-standard, non-obvious way that differs from their in-memory representation.
            Case OLD_PDPACKAGE_VERSION_SUMMER_2014
                
                'Make sure the struct size matches; if it doesn't, all subsequent parsing will fail
                Dim tmpNode As PDP_NODE
                If m_FileHeader.NodeStructSize = Len(tmpNode) Then
                    
                    'Use information from the file header to prepare the directory array.
                    m_NumOfNodes = m_FileHeader.NodeCount
                    ReDim m_NodeDirectory(0 To m_NumOfNodes - 1) As PDP_NODE
                    
                    'Retrieve the node directory
                    Get #fileNum, , m_NodeDirectory
                    
                    'Finally, retrieve the data chunk, and pass it to the m_DataBuffer class (which we'll use for further parsing).
                    Dim rawDataBuffer() As Byte
                    ReDim rawDataBuffer(0 To m_FileHeader.DataChunkSize - 1) As Byte
                    Get #fileNum, , rawDataBuffer
                    
                    Set m_DataBuffer = New clsBasicBuffer
                    m_DataBuffer.OpenBuf
                    m_DataBuffer.WriteBytes rawDataBuffer
                    
                    'Mark the load as successful
                    loadLegacyPDPackage = True
                    
                'Struct size mismatch - file is probably not PDI type.
                Else
                    raiseInternalMessage "Node struct size in header is invalid.  This file looks to be corrupt!  Abandoning load."
                    loadLegacyPDPackage = False
                End If
                
            Case Else
                loadLegacyPDPackage = False
            
        End Select
        
    'Close the file
    Close #fileNum

End Function

'Returns the version of a loaded pdPackage.  Return value is meaningless if no package is loaded.
Public Function getPDPackageVersion() As Long
    getPDPackageVersion = m_FileHeader.PDP_Version
End Function

'Returns the current number of nodes in the package.  This is primarily designed to be used right after loading a pdPackage file;
' after this number is obtained, the caller can iterate through individual nodes, extracting data as they go.
Public Function getNumOfNodes() As Long
    getNumOfNodes = m_NumOfNodes
End Function

'Given a node index, return relevant header data for that node.  This function will fail if the node index is out of bounds.
' There is generally no need to use this function, as the direct node data access functions will handle all this automatically.
' But I guess it's here if you need it.
Public Function getNodeInfo(ByVal nodeIndex As Long, Optional ByRef dstNodeName As String, Optional ByRef dstNodeID As Long, Optional ByRef dstOptionalNodeType As Long) As Boolean

    'Start by validating the node index we were passed.  If it's out of range, exit immediately.
    If nodeIndex < 0 Or nodeIndex > m_NumOfNodes - 1 Then
        Debug.Print "Node index out of range - try again with a valid node index!"
        getNodeInfo = False
        Exit Function
    End If
    
    'Fill the supplied variables with the corresponding data for this node
    With m_NodeDirectory(nodeIndex)
        dstNodeName = .NodeName
        dstNodeID = .NodeID
        dstOptionalNodeType = .OptionalNodeType
    End With
    
    getNodeInfo = True

End Function

'Shortcut function for extracting a node's contents directly to file.  This function requires the data to have been added via the dedicated
' autoAddNodeFromFile() function, above, as that function automatically populates all necessary fields for file extraction.
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
Public Function autoExtractSingleFile(ByVal dstPath As String, ByVal srcFilename As String, Optional ByVal dstPathIncludesFilename As Boolean = False, Optional ByVal OptionalNodeType As Long = 0, Optional ByVal ignoreRelativePathIfPresent As Boolean = False) As Boolean

    'Node headers use a fixed-length 32 character string for storing names.  As such, we pad the file node type identifier to 32 chars.
    Dim nodeName32_1 As String * 32, nodeName32_2 As String * 32
    nodeName32_1 = PDP_DATA_FILE
    nodeName32_2 = PDP_DATA_FILEANDPATH
    
    'If this node contains a relative path, we'll detect it automatically
    Dim thisNodeType As PDP_SPECIAL_NODE_TYPE
    
    'Cast the source filename to an array.  This allows us to make direct byte array comparisons without coercing the byte arrays into strings.
    Dim filenameArray() As Byte
    ReDim filenameArray(0 To Len(srcFilename) * 2 - 1) As Byte
    CopyMemory VarPtr(filenameArray(0)), StrPtr(srcFilename), UBound(filenameArray) + 1
    
    'Start searching the node array for matches
    Dim retBytes() As Byte, retString As String
    Dim i As Long, nodeIndex As Long, nodeFound As Boolean
    nodeIndex = -1
    nodeFound = False
    
    For i = 0 To UBound(m_NodeDirectory)
    
        'Look for two types of node names: bare filenames, and filenames with relative paths attached.
        ' (pdPackage handles both types automatically)
        If StrComp(nodeName32_1, m_NodeDirectory(i).NodeName, vbBinaryCompare) = 0 Then
            nodeFound = True
            thisNodeType = PDPSNT_FILENAME
        ElseIf StrComp(nodeName32_2, m_NodeDirectory(i).NodeName, vbBinaryCompare) = 0 Then
            nodeFound = True
            thisNodeType = PDPSNT_FILENAME_AND_RELATIVE_PATH
        Else
            nodeFound = False
        End If
        
        If nodeFound Then
        
            'This node appears to be a file node.  As a failsafe, check for the special file-type node flag.
            If getBitFlag(PDP_FLAG_FILE_NODE, m_NodeDirectory(i).NodeFlags(0)) Then
            
                'This is a valid file-type node.  If the user supplied an optionalNodeType, check it now.
                nodeFound = True
                
                If OptionalNodeType <> 0 Then
                    If m_NodeDirectory(i).OptionalNodeType = OptionalNodeType Then
                        nodeFound = True
                    Else
                        nodeFound = False
                    End If
                End If
                
                'If the optional node type check succeeded, the last thing we need to do is check the actual filename of this node.
                If nodeFound Then
                    
                    'Extract the header array, which will contain the filename of this file
                    If getNodeDataByIndex(i, True, retBytes, True) Then
                    
                        'If this byte array equals the source filename, we have a match (DING DING DING).
                        If checkArrayEquality(filenameArray, retBytes) Then
                        
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
            dstPath = m_File.EnforcePathSlash(dstPath)
            
            Select Case thisNodeType
            
                'This path is just a bare filename.  Append it to the destination path and carry on
                Case PDPSNT_FILENAME
                    fullDestinationPath = dstPath & srcFilename
                
                'This path includes a relative path.  Additional work is required.
                Case PDPSNT_FILENAME_AND_RELATIVE_PATH
                
                    'The user wants us to ignore the relative path and just write the damn file - easy enough!
                    If ignoreRelativePathIfPresent Then
                        fullDestinationPath = dstPath & m_File.GetFilename(srcFilename)
                    
                    'Relative path(s) must be preserved.
                    Else
                    
                        'Start by assembling the desired final filename, with all relative folders attached.
                        
                        'First, strip any preceding slashes in the relative path
                        If StrComp(Left$(srcFilename, 1), "\", vbBinaryCompare) = 0 Then srcFilename = Right$(srcFilename, Len(srcFilename) - 1)
                        
                        'Concatenate the two paths
                        Dim dstPathComplete As String
                        dstPathComplete = dstPath & m_File.GetPathOnly(srcFilename)
                        
                        'Attempt to create the folder.  (If the folder already exists, this will return TRUE, so no harm.)
                        If m_File.CreateFolder(dstPathComplete, True) Then
                        
                            'The full target path exists.  Append the original filename to it, then proceed with the extraction
                            fullDestinationPath = m_File.EnforcePathSlash(dstPathComplete) & m_File.GetFilename(srcFilename)
                        
                        'The path could not be constructed.  Set the full destination path to a blank string, which will cause the function
                        ' to terminate prematurely.
                        Else
                            fullDestinationPath = ""
                        End If
                    
                    End If
                    
            End Select
            
        End If
        
        'Make sure the destination path was successfully constructed, including any relative path data
        If Len(fullDestinationPath) > 0 Then
        
            'Retrieve this node's contents
            If getNodeDataByIndex(nodeIndex, False, retBytes()) Then
            
                'Erase the target file, if it exists
                If m_File.FileExist(fullDestinationPath) Then Kill fullDestinationPath
                
                'Dump the node's contents to file
                Dim fileNum As Integer
                fileNum = FreeFile
                
                Open fullDestinationPath For Binary Access Write As #fileNum
                    Put #fileNum, , retBytes
                Close #fileNum
                
                autoExtractSingleFile = True
            
            Else
                autoExtractSingleFile = False
            End If
            
        Else
            Debug.Print "WARNING! pdPackage.autoExtractSingleFile() was unable to construct the output path for " & srcFilename & "."
            Debug.Print "WARNING! As such, the file was not written.  (But one or more folders may have been constructed as part of the extraction.)"
            autoExtractSingleFile = False
        End If
        
    Else
        autoExtractSingleFile = False
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
Public Function autoExtractAllFiles(ByVal dstPath As String, Optional ByVal OptionalNodeType As Long = 0, Optional ByVal ignoreRelativePaths As Boolean = False) As Boolean
    
    'Enforce a trailing slash on dstPath, as we'll potentially be appending a lot of filenames to it
    dstPath = m_File.EnforcePathSlash(dstPath)
    
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
        If StrComp(nodeName32_1, m_NodeDirectory(i).NodeName, vbBinaryCompare) = 0 Then
            nodeFound = True
            thisNodeType = PDPSNT_FILENAME
        ElseIf StrComp(nodeName32_2, m_NodeDirectory(i).NodeName, vbBinaryCompare) = 0 Then
            nodeFound = True
            thisNodeType = PDPSNT_FILENAME_AND_RELATIVE_PATH
        Else
            nodeFound = False
        End If
        
        If nodeFound Then
        
            'This node appears to be a file node.  As a failsafe, check for the special file-type node flag.
            If getBitFlag(PDP_FLAG_FILE_NODE, m_NodeDirectory(i).NodeFlags(0)) Then
            
                'This is a valid file-type node.  If the user supplied an optionalNodeType, check it now.
                nodeFound = True
                
                If OptionalNodeType <> 0 Then
                    If m_NodeDirectory(i).OptionalNodeType = OptionalNodeType Then
                        nodeFound = True
                    Else
                        nodeFound = False
                    End If
                End If
                
                'If the optional node type check succeeded, this is a valid file.  Prepare to extract it.
                If nodeFound Then
                    
                    'Extract the header array, which contain the file's name and any relative path data
                    If getNodeDataByIndex(i, True, retBytes, True) Then
                        
                        'Cast the returned bytes directly into a string; this preserves Unicode characters.
                        strFilename = Space$((UBound(retBytes) + 1) \ 2)
                        CopyMemory StrPtr(strFilename), VarPtr(retBytes(0)), UBound(retBytes) + 1
                        Debug.Print "First: " & strFilename
                        'Now we need to construct a full destination path + filename.  This can vary by several criteria.
                        Dim fullDestinationPath As String
                        
                        'If the caller wants relative paths ignored, extract just the filename from the original path, and append it to the destination path.
                        If ignoreRelativePaths Then
                            fullDestinationPath = dstPath & m_File.GetFilename(strFilename)
                        
                        'The caller wants to make use of relative paths.
                        Else
                            
                            Select Case thisNodeType
                            
                                'This path is just a bare filename.  Append it to the destination path and carry on.
                                Case PDPSNT_FILENAME
                                    fullDestinationPath = dstPath & strFilename
                                
                                'This path includes a relative path.  Additional work is required to make sure the destination folder hierarchy exists.
                                Case PDPSNT_FILENAME_AND_RELATIVE_PATH
                                
                                    'Start by assembling the desired final filename, with all relative folders attached.
                                    
                                    'First, strip any preceding slashes in the relative path
                                    If StrComp(Left$(strFilename, 1), "\", vbBinaryCompare) = 0 Then strFilename = Right$(strFilename, Len(strFilename) - 1)
                                    
                                    'Concatenate the two paths
                                    Dim dstPathComplete As String
                                    dstPathComplete = dstPath & m_File.GetPathOnly(strFilename)
                                    
                                    'Attempt to create the folder hierarchy.  (If the folder already exists, this will return TRUE, so no harm.)
                                    If m_File.CreateFolder(dstPathComplete, True) Then
                                    
                                        'The full target path exists.  Append the original filename to it, then proceed with the extraction
                                        fullDestinationPath = m_File.EnforcePathSlash(dstPathComplete) & m_File.GetFilename(strFilename)
                                        
                                    'The path could not be constructed.  Set the full destination path to a blank string, which will cause the function
                                    ' to terminate prematurely.
                                    Else
                                        fullDestinationPath = ""
                                    End If
                                    
                            End Select
                        
                        End If
                        
                        'Make sure the destination path was successfully constructed, including any relative path data.
                        ' (If constructing the relative path failed, the destination path string will have been erased.)
                        If Len(fullDestinationPath) > 0 Then
                        
                            'Retrieve this node's contents
                            If getNodeDataByIndex(i, False, retBytes()) Then
                            
                                'Erase the target file, if it exists
                                If m_File.FileExist(fullDestinationPath) Then Kill fullDestinationPath
                                
                                'Dump the node's contents to file
                                Dim fileNum As Integer
                                fileNum = FreeFile
                                
                                Open fullDestinationPath For Binary Access Write As #fileNum
                                    Put #fileNum, , retBytes
                                Close #fileNum
                                
                            Else
                                Debug.Print "WARNING! pdPackage.autoExtractAllFiles() could not extract file contents from an otherwise valid node.  Extraction abandoned."
                                allNodesSuccessful = False
                            End If
                            
                        Else
                            Debug.Print "WARNING! pdPackage.autoExtractAllFiles() was unable to construct the output path for " & strFilename & "."
                            Debug.Print "WARNING! As such, the file was not written.  (But one or more folders may have been constructed as part of the extraction.)"
                            allNodesSuccessful = False
                        End If
                        
                    Else
                        allNodesSuccessful = False
                        Debug.Print "WARNING! pdPackage.autoExtractAllFiles() could not extract a filename from an otherwise valid node.  Extraction abandoned."
                    End If
                    
                'This set of 3 EndIf's looks confusing, but they simply terminate all file node detection criteria blocks.
                ' No action needs to be taken if any node detection checks fail; that just means this node isn't part of
                ' the current extraction request.
                End If
            End If
        End If
    
    Next i
    
    'Report success.  allNodesSuccessful will be TRUE IFF there were no failures or errors during extraction.
    autoExtractAllFiles = allNodesSuccessful

End Function

'Are two byte arrays equal?
Private Function checkArrayEquality(ByRef srcArray1() As Byte, ByRef srcArray2() As Byte) As Boolean

    'Check bounds first
    If UBound(srcArray1) <> UBound(srcArray2) Then
        checkArrayEquality = False
    
    'Array bounds are equal, so an actual per-element test is required
    Else
    
        Dim i As Long
        For i = 0 To UBound(srcArray1)
        
            If srcArray1(i) <> srcArray2(i) Then
                checkArrayEquality = False
                Exit Function
            End If
        
        Next i
        
        'If we made it all the way here, the arrays are equal
        checkArrayEquality = True
    
    End If

End Function

'Node data access functions.  These functions fill a byte array with the requested data, and they will return TRUE if successful.
' (FALSE generally only occurs if the requested node cannot be found in the source data.)
'
'If useHeaderBuffer is TRUE, the node's header buffer will be retrieved.  Otherwise, the node's data buffer will be used.
'
' Checksum validation and decompression are handled automatically, depending on the contents of the source file.  If performance
' is at a premium, you can forcibly disable checksum validation, even if the file supplies checksum data.
Public Function getNodeDataByName(ByVal targetNodeName As String, ByVal useHeaderBuffer As Boolean, ByRef dstArray() As Byte, Optional ByVal disableChecksumValidation As Boolean = False) As Boolean

    'Node headers use a fixed-length 32 character string for storing names.  As such, pad the supplied node name to
    ' 32 chars as necessary.
    Dim nodeName32 As String * 32
    nodeName32 = targetNodeName
    
    'Search the node array for a matching name
    Dim i As Long, nodeIndex As Long
    nodeIndex = -1
    
    For i = 0 To UBound(m_NodeDirectory)
    
        If StrComp(nodeName32, m_NodeDirectory(i).NodeName, vbTextCompare) = 0 Then
        
            'This node is the one!
            nodeIndex = i
            Exit For
        
        End If
    
    Next i
    
    'If a matching node was found, use getNodeDataByIndex() to retrieve its data
    If nodeIndex > 0 Then
        getNodeDataByName = getNodeDataByIndex(nodeIndex, useHeaderBuffer, dstArray(), disableChecksumValidation)
    Else
        getNodeDataByName = False
    End If

End Function

Public Function getNodeDataByID(ByVal targetNodeID As Long, ByVal useHeaderBuffer As Boolean, ByRef dstArray() As Byte, Optional ByVal disableChecksumValidation As Boolean = False) As Boolean

    'Search the node array for a matching ID
    Dim i As Long, nodeIndex As Long
    nodeIndex = -1
    
    If m_NumOfNodes > 0 Then
    
        For i = 0 To UBound(m_NodeDirectory)
        
            If targetNodeID = m_NodeDirectory(i).NodeID Then
            
                'This node is the one!
                nodeIndex = i
                Exit For
            
            End If
        
        Next i
        
        'If a matching node was found, use getNodeDataByIndex() to retrieve its data
        If nodeIndex >= 0 Then
            getNodeDataByID = getNodeDataByIndex(nodeIndex, useHeaderBuffer, dstArray(), disableChecksumValidation)
        Else
            getNodeDataByID = False
        End If
        
    Else
        getNodeDataByID = False
    End If

End Function

Public Function getNodeDataByID_UnsafeDstPointer(ByVal targetNodeID As Long, ByVal useHeaderBuffer As Boolean, ByVal dstPointer As Long, Optional ByVal disableChecksumValidation As Boolean = False) As Boolean

    'Search the node array for a matching ID
    Dim i As Long, nodeIndex As Long
    nodeIndex = -1
    
    If m_NumOfNodes > 0 Then
    
        For i = 0 To UBound(m_NodeDirectory)
        
            If targetNodeID = m_NodeDirectory(i).NodeID Then
            
                'This node is the one!
                nodeIndex = i
                Exit For
            
            End If
        
        Next i
        
        'If a matching node was found, use getNodeDataByIndex() to retrieve its data
        If nodeIndex >= 0 Then
            getNodeDataByID_UnsafeDstPointer = getNodeDataByIndex_UnsafeDstPointer(nodeIndex, useHeaderBuffer, dstPointer, disableChecksumValidation)
        Else
            getNodeDataByID_UnsafeDstPointer = False
        End If
        
    Else
        getNodeDataByID_UnsafeDstPointer = False
    End If
        
End Function

'getNodeHeaderByName and getNodeHeaderByID both wrap this function, getNodeHeaderByIndex.  (Those functions are simply used
' to retrieve the relevant index for a node.)
'
'Given a node index, and a target location (header buffer or data buffer), fill a destination array with the contents of that
' data buffer.  Decompression and checksum validation is handled automatically, depending on the contents of the node.
Public Function getNodeDataByIndex(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByRef dstArray() As Byte, Optional ByVal disableChecksumValidation As Boolean = False) As Boolean

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
    m_DataBuffer.Position = dataOffset
    
    'If the original size and packed size are different, assume the data is compressed.
    If (dataPackedSize <> dataOriginalSize) Then
        
        'Make sure zLib is available.  If it isn't, the user is screwed.
        If m_ZLibAvailable Then
        
            'Resize the destination buffer to the original size of the data
            ReDim dstArray(0 To dataOriginalSize - 1) As Byte
            
            'Use zLib to decompress the data
            Dim zLibResult As Long
            zLibResult = uncompress_safe(VarPtr(dstArray(0)), VarPtr(dataOriginalSize), m_DataBuffer.ReadBytes_PointerOnly(dataPackedSize), dataPackedSize)
            
            If zLibResult <> 0 Then
                
                raiseInternalMessage "File node is compressed, but zLib failed to decompress it.  Error code returned: %1", zLibResult
                getNodeDataByIndex = False
                Exit Function
            
            End If
            
        Else
        
            raiseInternalMessage "File node is compressed, but zLib is missing.  Decompression impossible; abandoning load."
            getNodeDataByIndex = False
            Exit Function
        
        End If
        
    'If the data is uncompressed, make a direct copy of the corresponding chunk of the master buffer
    Else
        dstArray = m_DataBuffer.ReadBytes(dataPackedSize)
    End If
    
    'Retrieve the stored checksum value for this node (if any).
    Dim checkSumOriginal As Long
    
    If useHeaderBuffer Then
        checkSumOriginal = m_NodeDirectory(nodeIndex).NodeHeaderAdler32
    Else
        checkSumOriginal = m_NodeDirectory(nodeIndex).NodeDataAdler32
    End If
    
    'Apply checksum validation now, unless one of three criteria is met:
    ' 1) The node does not contain a checksum to validate against
    ' 2) The user has forcibly disabled checksumming this node
    ' 3) zLib is missing
    If (checkSumOriginal <> 0) And (Not disableChecksumValidation) And m_ZLibAvailable Then
    
        'Like CRC32 functions, Adler checksums accept a previous value as their initial input.  If you don't want to supply
        ' this, you can supply a null buffer to get the library's recommended initial value.  (That's what we do here.)
        Dim zLibAdlerSum As Long
        zLibAdlerSum = checkSumArbitraryArray(dstArray)
        
        'If the checksums do not match, fail the function and exit
        If checkSumOriginal <> zLibAdlerSum Then
            raiseInternalMessage "Checksum failed for node #" & nodeIndex & ".  Expected value: " & checkSumOriginal & ", calculated value: " & zLibAdlerSum & ". Load abandoned."
            getNodeDataByIndex = False
            Exit Function
        Else
            #If DEBUGMODE = 1 Then
                If useHeaderBuffer Then
                    pdDebug.LogAction "Checksum successfully verified for node #" & nodeIndex & " (header)."
                Else
                    pdDebug.LogAction "Checksum successfully verified for node #" & nodeIndex & " (data)."
                End If
            #End If
        End If
        
    End If
    
    'If we made it all the way here, the node has been successfully loaded!
    getNodeDataByIndex = True

End Function

'WARNING!  This function is for expert users only.  Do not use it unless you fully understand the repercussions described below.
'
'Given a node index and a target buffer region (header or data), blindly copy the contents of that data buffer to a pointer supplied by the user.
' As with the normal getNodeDataByIndex() function, decompression and checksum validation is handled automatically, depending on the contents of the node.
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
Public Function getNodeDataByIndex_UnsafeDstPointer(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByVal dstPointer As Long, Optional ByVal disableChecksumValidation As Boolean = False) As Boolean

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
    m_DataBuffer.Position = dataOffset
    
    'If the original size and packed size are different, assume the data is compressed.
    If (dataPackedSize <> dataOriginalSize) Then
        
        'Make sure zLib is available.  If it isn't, the user is screwed.
        If m_ZLibAvailable Then
            
            'Use zLib to decompress the data directly into the pointer we've been given
            Dim zLibResult As Long
            zLibResult = uncompress_safe(dstPointer, VarPtr(dataOriginalSize), m_DataBuffer.ReadBytes_PointerOnly(dataPackedSize), dataPackedSize)
            
            If zLibResult <> 0 Then
                
                raiseInternalMessage "File node is compressed, but zLib failed to decompress it.  Error code returned: %1", zLibResult
                getNodeDataByIndex_UnsafeDstPointer = False
                Exit Function
            
            End If
            
        Else
        
            raiseInternalMessage "File node is compressed, but zLib is missing.  Decompression impossible; abandoning load."
            getNodeDataByIndex_UnsafeDstPointer = False
            Exit Function
        
        End If
        
    'If the data is uncompressed, make a direct copy of the corresponding chunk of the master buffer
    Else
        CopyMemory dstPointer, m_DataBuffer.ReadBytes_PointerOnly(dataPackedSize), dataPackedSize
    End If
    
    'Retrieve the stored checksum value for this node (if any).
    Dim checkSumOriginal As Long
    
    If useHeaderBuffer Then
        checkSumOriginal = m_NodeDirectory(nodeIndex).NodeHeaderAdler32
    Else
        checkSumOriginal = m_NodeDirectory(nodeIndex).NodeDataAdler32
    End If
    
    'Apply checksum validation now, unless one of three criteria is met:
    ' 1) The node does not contain a checksum to validate against
    ' 2) The user has forcibly disabled checksumming this node
    ' 3) zLib is missing
    If (checkSumOriginal <> 0) And (Not disableChecksumValidation) And m_ZLibAvailable Then
    
        'Like CRC32 functions, Adler checksums accept a previous value as their initial input.  If you don't want to supply
        ' this, you can supply a null buffer to get the library's recommended initial value.  (That's what we do here.)
        Dim zLibAdlerSum As Long
        zLibAdlerSum = checkSumArbitraryData(dstPointer, dataOriginalSize)
        
        'If the checksums do not match, fail the function and exit
        If checkSumOriginal <> zLibAdlerSum Then
            raiseInternalMessage "Checksum failed for node #" & nodeIndex & ".  Expected value: " & checkSumOriginal & ", calculated value: " & zLibAdlerSum & ". Load abandoned."
            getNodeDataByIndex_UnsafeDstPointer = False
            Exit Function
        Else
            If useHeaderBuffer Then
                Debug.Print "Checksum successfully verified for node #" & nodeIndex & " (header)."
            Else
                Debug.Print "Checksum successfully verified for node #" & nodeIndex & " (data)."
            End If
        End If
        
    End If
    
    'If we made it all the way here, the node has been successfully loaded!
    getNodeDataByIndex_UnsafeDstPointer = True

End Function

'If you want to use compression functions, you must provide the class with a path to a STDCALL (WAPI) copy of zLib, including
' "/zlibwapi.dll" at the end of the path.  This class assumes a file named "zlibwapi.dll"; if you are using one called zLib.dll,
' you must rewrite the zLib API function declarations to match.
'
'Optionally, if the caller wants to handle LoadLibrary/FreeLibrary by itself, it is welcome to do so.  (If you plan on using
' multiple pdPackage instances within a single application, it makes more sense to have the caller handle this just once.)
' Simply pass TRUE as the second parameter and this class will assume you've already used LoadLibrary successfully.
' Additionally, the third parameter can be set as FALSE to notify this class that the caller tried to initiate zLib, but
' failed to do so successfully.
Public Function init_ZLib(ByVal pathToZLibDLL As String, Optional ByVal leaveInitToCaller As Boolean = False, Optional callerInitiationSuccess As Boolean = True) As Boolean

    If leaveInitToCaller Then
        m_ZLibAvailable = callerInitiationSuccess
        m_ZLibToken = 0
        
    'The caller wants us to handle zLib initialization.  Attempt to load the zLib library path we were passed.
    Else
    
        If Len(pathToZLibDLL) <> 0 Then
        
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

'Set an arbitrary 1-bit position [0-31] in a Long-type value to either 1 or 0.
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

'Retrieve an arbitrary 1-bit position [0-31] in a Long-type value.
' Inputs:
'  1) position of the flag, which must be in the range [0, 31]
'  2) The Long from which you want the bit retrieved
Private Function getBitFlag(ByVal flagPosition As Long, ByVal srcLong As Long) As Boolean

    If (flagPosition >= 0) And (flagPosition <= 31) Then
    
        'Create a bitmask, using flagPosition to determine bit offset
        Dim longMask As Long
        If flagPosition < 31 Then longMask = 2 ^ flagPosition Else longMask = &H80000000
        
        getBitFlag = ((srcLong And longMask) <> 0)
        
    Else
        Debug.Print "C'mon - you know Longs only have 32 bits.  Flag positions must be on the range [0, 31]!"
    End If

End Function

'Once a package has been loaded from file, you can use this function to check various flags in the file
Public Function getPackageFlag(ByVal flagToCheck As PDP_FLAGS, Optional ByVal checkLocation As PDP_FLAG_LOCATION = PDP_LOCATION_ANY, Optional ByVal nodeIndex As Long = 0) As Boolean

    'Check the relevant flag location
    Select Case checkLocation
    
        'Any location
        Case PDP_LOCATION_ANY
        
            'Check directory and data flags first
            getPackageFlag = getBitFlag(flagToCheck, m_FileHeader.DirectoryFlags(0)) Or getBitFlag(flagToCheck, m_FileHeader.DataFlags(0))
            
            'If the flag is still false, scan individual nodes, looking for a hit
            If Not getPackageFlag Then
            
                Dim i As Long
                For i = 0 To m_NumOfNodes - 1
                    getPackageFlag = getPackageFlag Or getBitFlag(flagToCheck, m_NodeDirectory(i).NodeFlags(0))
                    If getPackageFlag Then Exit For
                Next i
            
            End If
        
        'Directory chunk flags
        Case PDP_LOCATION_DIRECTORY
            getPackageFlag = getBitFlag(flagToCheck, m_FileHeader.DirectoryFlags(0))
        
        'Data chunk flags
        Case PDP_LOCATION_DATA
            getPackageFlag = getBitFlag(flagToCheck, m_FileHeader.DataFlags(0))
        
        'Specific node flags; if used, we'd better hope the user supplied the right node index!
        Case PDP_LOCATION_INDIVIDUALNODE
            getPackageFlag = getBitFlag(flagToCheck, m_NodeDirectory(nodeIndex).NodeFlags(0))
    
    End Select

End Function

'Since this class makes extensive use of checksums, it's a convenient place to expose an arbitrary checksum function.  This function uses the
' same checksum strategy as pdPackage itself, so checksums by this function can be used as a failsafe check against internal pdPackage checksums.
Public Function checkSumArbitraryData(ByVal dataPointer As Long, ByVal dataLength As Long) As Long
    
    'Like CRC32 functions, Adler checksums accept a previous value as their initial input.  If you don't want to supply
    ' this, you can supply a null buffer to get the library's recommended initial value.  (That's what we do here.)
    checkSumArbitraryData = adler32(adler32(0, ByVal 0&, 0), ByVal dataPointer, dataLength)
    
End Function

'Thin wrapper to checkSumArbitraryData, above, as a convenience when working with VB arrays
Public Function checkSumArbitraryArray(ByRef array0Based() As Byte) As Long
    checkSumArbitraryArray = checkSumArbitraryData(VarPtr(array0Based(0)), UBound(array0Based) + 1)
End Function

'Thin wrapper to checkSumArbitraryData, above, as a convenience when working with files
Public Function checkSumArbitraryFile(ByRef pathToFile As String) As Long

    'Attempt to load the file into a byte array
    If m_File.FileExist(pathToFile) Then
    
        Dim fileNum As Integer
        fileNum = FreeFile
        
        Dim fileContents() As Byte
        
        Open pathToFile For Binary Access Read As #fileNum
            If LOF(fileNum) > 0 Then
                ReDim fileContents(0 To LOF(fileNum) - 1)
                Get fileNum, 1, fileContents
            End If
        Close #fileNum
        
        'Make sure the file loaded successfully
        If UBound(fileContents) >= LBound(fileContents) Then
            checkSumArbitraryFile = checkSumArbitraryArray(fileContents)
        End If
        
    End If

End Function

Private Sub Class_Initialize()

    'Reset all module-level variables
    m_ZLibAvailable = False
    m_ZLibToken = 0
    
    'Prepare a module-level pdFSO instance, which will make file operations much easier
    Set m_File = New pdFSO
    
End Sub

Private Sub Class_Terminate()

    'If we loaded zLib, free it now
    If m_ZLibToken <> 0 Then FreeLibrary m_ZLibToken
    
    'Release our data stream buffer
    If Not (m_DataBuffer Is Nothing) Then
        m_DataBuffer.CloseBuf
        Set m_DataBuffer = Nothing
    End If
    
    'Release our pdFSO instance
    Set m_File = Nothing

End Sub

'If something unexpected occurs, and the user needs to know about it (e.g. it's not just debugging data), pdPackage will raise a message
' using this function.  PhotoDemon uses a central "Message" function to handle messages like this, but you can change it to Debug.Print instead.
Private Sub raiseInternalMessage(ByVal msgString As String, ParamArray ExtraText() As Variant)
    
    If UBound(ExtraText) >= LBound(ExtraText) Then
        Message msgString, ExtraText
    Else
        Message msgString
    End If
    
End Sub
