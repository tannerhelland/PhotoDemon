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
'Copyright ©2013-2014 by Tanner Helland
'Created: 05/April/14
'Last updated: 24/October/14
'Last update: add support for second-pass directory chunk and data chunk compression
'Dependencies: clsBasicBuffer class by dilettante (used to read/write byte array buffers more easily)
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
' - The lowest expected value for this constant is 64, representative of PhotoDemon 6.4, when the format was first implemented.
' - The current expected value is 65.  65 slightly modifies the format of the directory chunk, by using a straight byte array
'   instead of the actual UDT struct used in previous versions (which VB custom-formats).  This was a necessary change prior to
'   enabling second-pass compression of the directory and data chunks.
' - Both v64 and v65 packages are supported by this class, but it will only write v65 packages.
Private Const THIS_PDPACKAGE_VERSION As Long = 65
Private Const OLD_PDPACKAGE_VERSION_SUMMER_2014 As Long = 64

'The first four bytes of a valid pdPackage file must always be &H4B504450 ("PDPK" in ASCII, if viewed in a text editor
' - note the little-endian declaration here).
Private Const PDP_UNIVERSAL_IDENTIFIER As Long = &H4B504450

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
                                            ' the data chnk prior to copying it into the buffer class.
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
End Enum

#If False Then
    Private Const PDP_FLAG_ZLIB_REQUIRED = 0
#End If

'Immediately following the PDP_HEADER is the directory chunk, which is comprised of FilerHeader.NodeCount individual
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
    
        'Like CRC32 functions, Adler checksums accept a previous value as their initial input.  If you don't want to supply
        ' this, you can supply a null buffer to get the library's recommended initial value.  (That's what we do here.)
        Dim zLibAdlerSum As Long
        zLibAdlerSum = adler32(adler32(0, ByVal 0&, 0), ByVal dataPointer, dataLength)
        
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

    'Prepare a temporary array to hold the string
    Dim tmpStringCopy() As Byte
    tmpStringCopy = StrConv(srcDataString, vbFromUnicode)
    
    'Forward the byte array to the actual addNodeData function, and return its result.
    addNodeDataFromString = addNodeDataFromPointer(nodeIndex, useHeaderBuffer, VarPtr(tmpStringCopy(0)), UBound(tmpStringCopy) + 1, compressData, compressionLevel, requestChecksum)

End Function

'When all nodes have been successfully added, the user can finally write out their data to file.  This function will return TRUE
' if successful, FALSE if unsuccessful.  (Note that it assumes the caller has done some basic validation on the file path, like
' obtaining permission from the user to overwrite an existing file.)
Public Function writePackageToFile(ByVal dstFilename As String, Optional ByVal secondPassDirectoryCompression As Boolean = False, Optional ByVal secondPassDataCompression As Boolean = False) As Boolean
    
    'Start by updating the file header.  Most of this will have been done when the class was initialized, but some values
    ' can't be set until all nodes have been added.
    With m_FileHeader
    
        'Update the final node count
        .NodeCount = m_NumOfNodes
                
        'Update the size of the directory chunk
        .DirectoryChunkSize = .NodeStructSize * m_NumOfNodes
        
        'Trim the data buffer, and update the size of the data chunk using the exact size of the final buffer.
        m_DataBuffer.ShrinkBuf
        .DataChunkSize = m_DataBuffer.Size
        
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
    If FileExist(dstFilename) Then Kill dstFilename
    
    'Open the target file
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open dstFilename For Binary As #fileNum
    
        'Writing the actual data to file is incredibly easy!
        
        'Header
        Put #fileNum, 1, m_FileHeader
        
        'Directory chunk
        Put #fileNum, , rawDirectoryBuffer
        
        'Data chunk
        Put #fileNum, , rawDataBuffer
        
    Close #fileNum
    
    'Easy as pie.  Return TRUE and exit.
    writePackageToFile = True
    
    'TODO!  Add basic error handling.

End Function

'Copy one array into another, with variable compression.  This is used in many places throughout this class; basically any action that
' potentially supports compression should use this function to transfer data between arrays, even when compression is not in use.
' (This function will abstract away all the messy CopyMemory bits, keeping the rest of the class clean.)
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
        dstSize = srcSize + (srcSize * 0.01) + 12
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

    Dim fileNum As Integer
    Dim rawBuffer() As Byte, tmpCompressionBuffer() As Byte
    Dim zLibResult As Long
    
    'Before doing anything else, make sure the file exists
    If Not FileExist(srcFilename) Then
        Debug.Print "Requested pdPackage file doesn't exist.  Validate all paths before sending them to the pdPackage class!"
        GoTo StopPackageFileRead
    End If
    
    'Open the file
    fileNum = FreeFile
    Open srcFilename For Binary As #fileNum
    
        'Validate the file signature.  Valid pdPackage files must have their first 4 bytes set to "PDPK" in ASCII.
        Dim pdpCheck As Long
        Get #fileNum, 1, pdpCheck
        
        If (pdpCheck <> PDP_UNIVERSAL_IDENTIFIER) Then
            Message "File doesn't have valid pdPackage header.  Abandoning load."
            GoTo StopPackageFileRead
        End If
        
        'If the caller requested a sub-type validation, perform that now
        If subTypeValidator <> 0 Then
        
            Get #fileNum, 5, pdpCheck
        
            If (pdpCheck <> subTypeValidator) Then
                Message "File doesn't match requested sub-type value.  Abandoning load."
                GoTo StopPackageFileRead
            End If
        
        End If
        
        'If both the file signature and the sub-type check out, load the full file header now.
        Get #fileNum, 1, m_FileHeader
        
        'Check the PDP version.  If it's older than the current version, loading will be handled by one of the legacy load functions.
        If m_FileHeader.PDP_Version < THIS_PDPACKAGE_VERSION Then
            
            pdpCheck = loadLegacyPDPackage(fileNum)
            If Not pdpCheck Then
                Message "Legacy file format could not be decoded."
                GoTo StopPackageFileRead
            End If
        
        'This is the current PDP version.  Continue the normal loading process.
        Else
            
            'Make sure the struct size is correct; if it isn't, all subsequent parsing will fail.
            Dim tmpNode As PDP_NODE
            If m_FileHeader.NodeStructSize <> LenB(tmpNode) Then
                Message "Node struct size in header is invalid.  This file looks to be corrupt!  Abandoning load."
                GoTo StopPackageFileRead
            End If
            
            'Next, it's time to retrieve the directory.  The size of the directory varies, based on whether it received second-pass compression.
            If m_FileHeader.DirectoryChunkSizeCompressed = 0 Then
            
                'No compression applied, so we can retrieve the node directory as-is.
                ReDim rawBuffer(0 To m_FileHeader.DirectoryChunkSize - 1) As Byte
                Get #fileNum, , rawBuffer
                
            Else
            
                'Second-pass compression was applied.  Make sure zLib is available.  (If it isn't, the user is screwed.)
                If m_ZLibAvailable Then
                
                    'Resize the destination buffer to the original size of the data
                    ReDim rawBuffer(0 To m_FileHeader.DirectoryChunkSize - 1) As Byte
                    
                    'Resize the source buffer to the compressed directory size, and retrieve the compressed directory.
                    ReDim tmpCompressionBuffer(0 To m_FileHeader.DirectoryChunkSizeCompressed - 1) As Byte
                    Get #fileNum, , tmpCompressionBuffer
                    
                    'Use zLib to decompress the data
                    zLibResult = uncompress(rawBuffer(0), m_FileHeader.DirectoryChunkSize, tmpCompressionBuffer(0), m_FileHeader.DirectoryChunkSizeCompressed)
                    
                    If zLibResult <> 0 Then
                        
                        Message "Directory is compressed, but zLib failed to decompress it.  Error code returned: %1", zLibResult
                        GoTo StopPackageFileRead
                    
                    End If
                    
                Else
                
                    Message "Directory chunk is compressed, but zLib is missing.  Decompression impossible; abandoning load."
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
                Get #fileNum, , rawBuffer
                
            Else
            
                'Second-pass compression was applied.  Make sure zLib is available.  (If it isn't, the user is screwed.)
                If m_ZLibAvailable Then
                
                    'Resize the destination buffer to the original size of the data
                    ReDim rawBuffer(0 To m_FileHeader.DataChunkSize - 1) As Byte
                    
                    'Resize the source buffer to the compressed directory size, and retrieve the compressed directory.
                    ReDim tmpCompressionBuffer(0 To m_FileHeader.DataChunkSizeCompressed - 1) As Byte
                    Get #fileNum, , tmpCompressionBuffer
                    
                    'Use zLib to decompress the data
                    zLibResult = uncompress(rawBuffer(0), m_FileHeader.DataChunkSize, tmpCompressionBuffer(0), m_FileHeader.DataChunkSizeCompressed)
                    
                    If zLibResult <> 0 Then
                        
                        Message "Data chunk is compressed, but zLib failed to decompress it.  Error code returned: %1", zLibResult
                        GoTo StopPackageFileRead
                    
                    End If
                    
                Else
                
                    Message "Data chunk is compressed, but zLib is missing.  Decompression impossible; abandoning load."
                    GoTo StopPackageFileRead
                
                End If
            
            End If
            
            'Pass the data chunk to the m_DataBuffer class (which we'll use for further parsing).
            Set m_DataBuffer = New clsBasicBuffer
            m_DataBuffer.OpenBuf
            m_DataBuffer.WriteBytes rawBuffer
            
        End If
        
    'And that's all, folks!  Close the file and exit.
    Close #fileNum
    
    readPackageFromFile = True
    Exit Function
    
StopPackageFileRead:
    
    Debug.Print "An error occurred in the pdPackager.readPackageFromFile function.  Additional data should have been supplied by the Message() function."
    If fileNum <> 0 Then Close #fileNum
    readPackageFromFile = False
    Exit Function

End Function

'Any PDPackage versions prior to the current version (THIS_PDPACKAGE_VERSION) are loaded via this function, after the file header
' has been parsed.  That means that the filePointer passed to this function already points to the byte IMMEDIATELY FOLLOWING the header.
Private Function loadLegacyPDPackage(ByRef fileNum As Integer) As Boolean

    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Legacy pdPackage format detected.  Attempting decoding now..."
    #End If

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
                
                'We do not close the file here; that is handled by the caller, which may have additional validation steps to perform.
                loadLegacyPDPackage = True
                
            'Struct size mismatch - file is probably not PDI type.
            Else
                Message "Node struct size in header is invalid.  This file looks to be corrupt!  Abandoning load."
                loadLegacyPDPackage = False
            End If
            
        Case Else
            loadLegacyPDPackage = False
        
    End Select

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
    
    For i = 0 To UBound(m_NodeDirectory)
    
        If targetNodeID = m_NodeDirectory(i).NodeID Then
        
            'This node is the one!
            nodeIndex = i
            Exit For
        
        End If
    
    Next i
    
    'If a matching node was found, use getNodeDataByIndex() to retrieve its data
    If nodeIndex > 0 Then
        getNodeDataByID = getNodeDataByIndex(nodeIndex, useHeaderBuffer, dstArray(), disableChecksumValidation)
    Else
        getNodeDataByID = False
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
    
    'Move the data buffer position to the relevant offset, then retrieve the data arary
    Dim tmpDataArray() As Byte
    m_DataBuffer.Position = dataOffset
    tmpDataArray = m_DataBuffer.ReadBytes(dataPackedSize)
    
    'If the original size and packed size are different, assume the data is compressed.
    If (dataPackedSize <> dataOriginalSize) Then
        
        'Make sure zLib is available.  If it isn't, the user is screwed.
        If m_ZLibAvailable Then
        
            'Resize the temporary buffer to the original size of the data
            ReDim dstArray(0 To dataOriginalSize - 1) As Byte
            
            'Use zLib to decompress the data
            Dim zLibResult As Long
            zLibResult = uncompress(dstArray(0), dataOriginalSize, tmpDataArray(0), UBound(tmpDataArray) + 1)
            
            If zLibResult <> 0 Then
                
                Message "File node is compressed, but zLib failed to decompress it.  Error code returned: %1", zLibResult
                getNodeDataByIndex = False
                Exit Function
            
            End If
            
        Else
        
            Message "File node is compressed, but zLib is missing.  Decompression impossible; abandoning load."
            getNodeDataByIndex = False
            Exit Function
        
        End If
        
    'The data is uncompressed.  No further processing is necessary.
    Else
        dstArray = tmpDataArray
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
        zLibAdlerSum = adler32(adler32(0, ByVal 0&, 0), dstArray(0), UBound(dstArray) + 1)
        
        'If the checksums do not match, fail the function and exit
        If checkSumOriginal <> zLibAdlerSum Then
            Message "Checksum failed for node #" & nodeIndex & ".  Expected value: " & checkSumOriginal & ", calculated value: " & zLibAdlerSum & ". Load abandoned."
            getNodeDataByIndex = False
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
    getNodeDataByIndex = True

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
        
        'Write  a 0 flag
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

Private Sub Class_Initialize()

    'Reset all module-level variables
    m_ZLibAvailable = False
    m_ZLibToken = 0
    
End Sub

Private Sub Class_Terminate()

    'If we loaded zLib, free it now
    If m_ZLibToken > 0 Then FreeLibrary m_ZLibToken
    
    'Release our data stream buffer
    If Not (m_DataBuffer Is Nothing) Then m_DataBuffer.CloseBuf
    Set m_DataBuffer = Nothing

End Sub
