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
'Created: 05/April/13
'Last updated: 09/April/14
'Last update: general code clean-up now that everything seems to be working.  Lots of stress testing will happen in
'             the coming weeks, but for now I think this class is working pretty damn well.
'Dependencies: clsBasicBuffer by dilettante (used to read/write byte array buffers more easily)
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
' 5) Support for checksum (Adler32) validation of each individual node.
' 6) Support for two data entries per node, typically a header byte array and an actual data byte array.  These two structs
'     don't need to be used (one or the other or neither is just fine), but I find it very helpful to be store two pieces of
'     data per node (typically so I can read key data from a node without extracting its full contents).
' 7) Per-node compression options, meaning you can extract a single file without having to decompress the entire
'     archive, and you can customize compression for each node in the package.
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
'     performance reasons, this pdPackager class creates the full archive in RAM before writing it to file.  This makes it
'     very fast, but ill-suited to extraordinarily large archive sizes.
'
' 4) When reading pdPackage files, the full file contents will be cached in memory.  Do not reload the pdPackage file unless
'     you absolutely have to!  Once it is cached, you can access individual nodes without reloading the entire file.
'
' 5) This class has clsBasicBuffer as a dependency.  Thank you to dilettante for this minimal buffer class, which greatly
'     simplifies the process of reading/writing byte streams.
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
' necessary to prepare a node directory array.
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

    NodeName As String * 32                 '32 double-byte characters are allotted for each node name
    NodeID As Long                          'Alternatively, calling functions can specify an optional 4-byte numerical ID.  Nodes can be read by 32-char name or 4-byte ID.
    OptionalNodeType As Long                'Calling functions can also assign each node a TYPE if they want; this value has no meaning to this pdPackager class.
    NodeFlags(0 To 3) As Long               '16 bytes of node-specific flags are allowed.  At present, these are unused.
    
    'One of the unique features of pdPackages is that each node is allotted two entries in the data chunk.  These entries don't have
    ' to be used; in fact, neither has to be used, but they can be helpful for reading node-specific information without having to
    ' decode the entire node contents.
    ' (Also, if I ever get around to implementing a file/folder wrapper for this class, the header chunk will be used to store
    '  relative path locations for each file in the package!)
    NodeHeaderOffset As Long                'Absolute offset of this node's header in the data chunk, STARTING FROM THE START OF THE DATA CHUNK, not the start of the file!
    NodeHeaderPackedSize As Long            'Packed size of this node's header.  (This is the size the node's header array occupies in the pdPackage data chunk.)
    NodeHeaderOriginalSize As Long          'Original size of this node's header.  (If this value is the same as NodeHeaderCompSize, the node header was stored uncompressed.)
    NodeHeaderAdler32 As Long               'Adler32 checksum of the UNCOMPRESSED header bytes.  This can be used to verify the correctness of the data post-decompression.
    
    NodeDataOffset As Long                  'Absolute offset of this node's data in the data chunk, STARTING FROM THE START OF THE DATA CHUNK, not the start of the file!
    NodeDataPackedSize As Long              'Packed size of this node's data.  (This is the size the node's data array occupies in the pdPackage data chunk.)
    NodeDataOriginalSize As Long            'Original size of this node's data.  (If this value is the same as NodeHeaderCompSize, the node data was stored uncompressed.)
    NodeDataAdler32 As Long                 'Adler32 checksum of the UNCOMPRESSED data bytes.  This can be used to verify the correctness of the data post-decompression.
    
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

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
' have a zLibToken value that must be freed at termination time.
Private m_ZLibAvailable As Boolean, m_ZLibToken As Long

'When writing new pdPackage files, these variables will hold the package contents as they are being assembled in memory.
Private m_FileHeader As PDP_HEADER
Private m_NodeDirectory() As PDP_NODE
Private m_NumOfNodes As Long

'The actual data chunk of the pdPackage will be assembled using a clsBasicBuffer instance.  It greatly simplifies the process of
' assembling a 1D byte array from discrete individual chunks.
Private m_DataBuffer As clsBasicBuffer

'pdPackage operations are roughly divided into two groups:
'
' - GET operations, for retrieving data from existing pdPackage files.  GET operations include:
'        readPackageFromFile, getNodeInfo, getNodeDataByID/Index/Name
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
Public Sub prepareNewPackage(Optional ByVal numOfDataNodes As Long = 0, Optional ByVal optPackageID As Long = 0)

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
        .NodeStructSize = Len(tmpNode)
        
        'DirectoryChunkSize can be assumed from the numOfDataNodes, but note that it will be verified again before the data is
        ' actually written out to file.
        .DirectoryChunkSize = .NodeStructSize * numOfDataNodes
        
        'DirectoryFlags() and DataFlags() are set by separate functions.  For now, assume a value of 0 for all flags.
        Dim i As Long
        For i = 0 To 3
            .DirectoryFlags(i) = 0
            .DataFlags(i) = 0
        Next i
        
        'DataChunkSize will remain unknown until all nodes have been added.
        .DataChunkSize = 0
        
        '4 bytes are reserved at the end as a "just in case".  For now, they should always be 0.
        .Reserved = 0
        
    End With
    
    'Resize the directory array to the number of supplied nodes; note that this step is optional; if no node count is supplied,
    ' the addNode() function will automatically increment itself as necessary.
    m_NumOfNodes = 0
    ReDim m_NodeDirectory(0 To numOfDataNodes) As PDP_NODE
        
    'Prepare the data buffer.  Note that chunk size is largely irrelevant for our purposes; when handed a byte array that exceeds
    ' the size of the default chunk size, the buffer class is smart enough to extend the buffer by the size of that byte array.
    Set m_DataBuffer = New clsBasicBuffer
    m_DataBuffer.ChunkSize = 32768
    m_DataBuffer.OpenBuf

End Sub

'Add a new node to this pdPackage instance.  Note that this function DOES NOT actually add the node's data arrays to the main
' data buffer - those are done in a subsequent step, by the user, as necessary.  (It's a little pesky to separate node additions
' into multiple steps, but this allows for much more fine-grained control over node addition, without overwhelming this function
' with a monstrous list of parameters.
'
'This function returns the index of the added node, which external functions can then use to supply this node's actual data arrays.
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

'Add data for a given node.  Three required params, three optional params.  The function will return TRUE if successful.
'
' The three required params are:
' 1) Index of the node for which data is being supplied
' 2) Destination array for this data.  Remember that each node in a pdPackage supports TWO data arrays, which can be
'    used however the caller pleases.  (Typically, one is used for header data, while the other is used for binary data.)
' 3) Reference to a byte array containing the data the caller wants written.  This class doesn't care how the byte array is
'    created or what it contains, with the exception that it CANNOT BE FIXED WIDTH.
'
' The three optional params are:
' 1) Whether to compress the data before writing it.  If zLib is unavailable, this param has no meaning, as data will always
'    be written without compression.
' 2) Compresson level.  Per the zLib spec, this is a value from 0 (uncompressed) to 9 (best possible compression).  Note that
'    a value of -1 means "use default compression level", and a value of 1 means "compress the data, but do it as quickly as
'    possible".  I don't recommend requesting compression, then using compression level 0 - that would be stupid!
' 3) Whether to calculate a checksum for the uncompressed data bytes.  This is completely optional.  If a checksum is found in
'    the file, pdPackage will automatically process it at load-time, and reject the file if checksums do not match.  Note that
'    checksumming the data increases write time non-trivially.
Public Function addNodeData(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByRef dataBytes() As Byte, Optional ByVal compressData As Boolean = True, Optional ByVal compressionLevel As Long = -1, Optional ByVal requestChecksum As Boolean = False) As Boolean

    'Start by validating the node index we were passed.  If it's out of range, exit immediately.
    If nodeIndex < 0 Or nodeIndex > m_NumOfNodes - 1 Then
        Debug.Print "Node index out of range - try again with a valid node index!"
        addNodeData = False
        Exit Function
    End If
    
    'Update the pre-compression values for this data chunk.
    If useHeaderBuffer Then
        m_NodeDirectory(nodeIndex).NodeHeaderOriginalSize = UBound(dataBytes) + 1
    Else
        m_NodeDirectory(nodeIndex).NodeDataOriginalSize = UBound(dataBytes) + 1
    End If
    
    'Mark the offset for this data chunk.
    If useHeaderBuffer Then
        m_NodeDirectory(nodeIndex).NodeHeaderOffset = m_DataBuffer.Position
    Else
        m_NodeDirectory(nodeIndex).NodeDataOffset = m_DataBuffer.Position
    End If
    
    'If we are not using compression, we can write the node data as-is.  Note that header or data doesn't matter here,
    ' because the data is simply added to the pdPackages data chunk.  Order is irrelevant.
    If (Not compressData) Or (Not m_ZLibAvailable) Then
    
        m_DataBuffer.WriteBytes dataBytes
        
        If useHeaderBuffer Then
            m_NodeDirectory(nodeIndex).NodeHeaderPackedSize = UBound(dataBytes) + 1
        Else
            m_NodeDirectory(nodeIndex).NodeDataPackedSize = UBound(dataBytes) + 1
        End If
        
    'Data compression was requested.  We now have a bit of extra work to do.
    Else
    
        'Start by preparing an array to receive the zLib-compressed data.  The zLib creators recommend a buffer 1% larger than
        ' the input buffer, plus 12 extra bytes for a header.  This is the "worst possible scenario" for zLib compression.
        Dim compressedData() As Byte
        Dim originalSize As Long, compressedSize As Long
        originalSize = UBound(dataBytes) + 1
        compressedSize = originalSize + (originalSize * 0.01) + 12
        ReDim compressedData(0 To compressedSize - 1) As Byte
        
        'Request compression from zLib.  It'll take care of the rest.
        Dim zLibResult As Long
        If compressionLevel >= 0 Then
            zLibResult = compress2(compressedData(0), compressedSize, dataBytes(0), UBound(dataBytes) + 1, compressionLevel)
        Else
            zLibResult = compress(compressedData(0), compressedSize, dataBytes(0), UBound(dataBytes) + 1)
        End If
        
        'zLib obviously has no way to update our VB-style SafeArray header, so we must manually resize the buffer now.
        'TODO: update the data buffer to provide an optional Length parameter, which takes n bytes from a buffer.
        '      That would save us this costly ReDim Preserve statement.
        ReDim Preserve compressedData(0 To compressedSize - 1) As Byte
        
        'Make sure zLib compressed the data successfully.
        If zLibResult = 0 Then
        
            'zLib was successful.  Add the compressed data stream to the data buffer.
            m_DataBuffer.WriteBytes compressedData
        
            If useHeaderBuffer Then
                m_NodeDirectory(nodeIndex).NodeHeaderPackedSize = UBound(compressedData) + 1
            Else
                m_NodeDirectory(nodeIndex).NodeDataPackedSize = UBound(compressedData) + 1
            End If
            
            'For fun, supply some debug info on the compression.
            Debug.Print "Node data compressed successfully; compressed data is " & Format(CStr(100 - (100 * (CDbl(compressedSize) / CDbl(originalSize)))), "#0.00") & "% smaller (" & originalSize & " to " & compressedSize & " bytes)."
            
        'zLib failed.  Write the original, uncompressed data to file instead.
        Else
            
            m_DataBuffer.WriteBytes dataBytes
        
            If useHeaderBuffer Then
                m_NodeDirectory(nodeIndex).NodeHeaderPackedSize = UBound(dataBytes) + 1
            Else
                m_NodeDirectory(nodeIndex).NodeDataPackedSize = UBound(dataBytes) + 1
            End If
            
            'Something went horribly wrong.  Write the uncompressed original data to the buffer, and provide a debug warning.
            Debug.Print "zLib compression failed; returned error code was #" & zLibResult & ".  Writing original uncompressed data to buffer instead..."
        
        End If
        
    End If
    
    'Next, checksum the file (if requested).
    If requestChecksum And m_ZLibAvailable Then
    
        'Like CRC32 functions, Adler checksums accept a previous value as their initial input.  If you don't want to supply
        ' this, you can supply a null buffer to get the library's recommended initial value.  (That's what we do here.)
        Dim zLibAdlerSum As Long
        zLibAdlerSum = adler32(adler32(0, ByVal 0&, 0), dataBytes(0), UBound(dataBytes) + 1)
        
        If useHeaderBuffer Then
            m_NodeDirectory(nodeIndex).NodeHeaderAdler32 = zLibAdlerSum
        Else
            m_NodeDirectory(nodeIndex).NodeDataAdler32 = zLibAdlerSum
        End If
    
    End If
    
    'This chunk was added successfully!  Return TRUE and exit.
    addNodeData = True
    
End Function

'Thin wrapper for addNodeData, above, but allows the user to supply a string.
Public Function addNodeDataFromString(ByVal nodeIndex As Long, ByVal useHeaderBuffer As Boolean, ByRef srcDataString As String, Optional ByVal compressData As Boolean = True, Optional ByVal compressionLevel As Long = -1, Optional ByVal requestChecksum As Boolean = False) As Boolean

    'Prepare a temporary array to hold the string
    Dim tmpStringCopy() As Byte
    tmpStringCopy = StrConv(srcDataString, vbFromUnicode)
    
    'Forward the byte array to the actual addNodeData function, and return its result.
    addNodeDataFromString = addNodeData(nodeIndex, useHeaderBuffer, tmpStringCopy, compressData, compressionLevel, requestChecksum)

End Function

'When all nodes have been successfully added, the user can finally write out their data to file.  This function will return TRUE
' if successful, FALSE if unsuccessful.  (Note that it assumes the caller has done some basic validation on the file path, like
' obtaining permission from the user to overwrite an existing file.)
Public Function writePackageToFile(ByVal dstFilename As String) As Boolean
    
    'Start by updating the file header.  Most of this will have been done when the class was initialized, but some values
    ' wouldn't have been known until all nodes were added.
    With m_FileHeader
    
        'Update the final node count
        .NodeCount = m_NumOfNodes
                
        'Update the size of the directory chunk
        .DirectoryChunkSize = .NodeStructSize * m_NumOfNodes
        
        'Trim the data buffer, and update the size of the data chunk using the exact size of the final buffer.
        m_DataBuffer.ShrinkBuf
        .DataChunkSize = m_DataBuffer.Size
        
    End With
    
    'If the node directory was being dynamically resized during the write process, it may contain blank entries at the tail end.
    ' Resize it to its exact final size now.
    If UBound(m_NodeDirectory) > m_NumOfNodes - 1 Then ReDim Preserve m_NodeDirectory(0 To m_NumOfNodes - 1) As PDP_NODE
    
    'Kill the destination file if it already exists
    If FileExist(dstFilename) Then Kill dstFilename
    
    'Open the target file
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open dstFilename For Binary As #fileNum
    
        'Writing the actual data to file is incredibly easy!
        
        'Header
        Put #fileNum, 1, m_FileHeader
        
        'Node directory
        Put #fileNum, , m_NodeDirectory
        
        'Data buffer
        Put #fileNum, , m_DataBuffer.CurrentBytes
        
    Close #fileNum
    
    'Easy as pie.  Return TRUE and exit.
    writePackageToFile = True
    
    'TODO!  Add basic error handling.

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
        
        'As yet another failsafe, make sure the header's node struct size is correct.  If it isn't, exit now.
        Dim tmpNode As PDP_NODE
        If m_FileHeader.NodeStructSize <> Len(tmpNode) Then
            Message "Node struct size in header is invalid.  This file looks to be corrupt!  Abandoning load."
            GoTo StopPackageFileRead
        End If
        
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
            
            Debug.Print "compressed size: " & dataPackedSize & ", uncompressed size: " & dataOriginalSize
            
            If zLibResult <> 0 Then
                
                Message "File node is compressed, but zLib failed to decompress it.  Error code returned: %1", zLibResult
                getNodeDataByIndex = False
                Exit Function
            
            End If
            
        Else
        
            Message "File node is compressed, but zLib dll is missing.  Decompression is impossible, abandoning load."
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
            Debug.Print "Checksum successfully verified for node #" & nodeIndex & "."
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

Private Sub Class_Initialize()

    'Reset all module-level variables
    m_ZLibAvailable = False
    m_ZLibToken = 0
    
End Sub

Private Sub Class_Terminate()

    'If we loaded zLib, free it now
    If m_ZLibToken > 0 Then FreeLibrary m_ZLibToken
    
    'Release our data stream buffer
    m_DataBuffer.CloseBuf
    Set m_DataBuffer = Nothing

End Sub
