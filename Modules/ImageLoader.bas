Attribute VB_Name = "ImageImporter"
'***************************************************************************
'Low-level image import interfaces
'Copyright 2001-2026 by Tanner Helland
'Created: 4/15/01
'Last updated: 12/November/25
'Last update: clean up final work on JPEG-2000 via OpenJPEG import path
'
'This module provides low-level "import" functionality for importing image files into PD.
' You will not generally want to interface with this module directly; instead, rely on the
' high-level functions in the "Loading" module. They will intelligently drop into this module
' as necessary, sparing you the messy work of handling format-specific import details.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PhotoDemon now provides many of its own image format parsers.  You can disable individual
' formats for testing purposes, but note that fallback methods like internal Windows libraries
' *CANNOT* read most (if any) of these formats.  If you encounter problems with a specific
' image format, PLEASE FILE AN ISSUE ON GITHUB.
Private Const USE_INTERNAL_PARSER_CBZ As Boolean = True
Private Const USE_INTERNAL_PARSER_HDR As Boolean = True
Private Const USE_INTERNAL_PARSER_HGT As Boolean = True
Private Const USE_INTERNAL_PARSER_ICO As Boolean = True
Private Const USE_INTERNAL_PARSER_MBM As Boolean = True
Private Const USE_INTERNAL_PARSER_ORA As Boolean = True
Private Const USE_INTERNAL_PARSER_PCX As Boolean = True
Private Const USE_INTERNAL_PARSER_PNG As Boolean = True
Private Const USE_INTERNAL_PARSER_PSD As Boolean = True
Private Const USE_INTERNAL_PARSER_PSP As Boolean = True
Private Const USE_INTERNAL_PARSER_QOI As Boolean = True
Private Const USE_INTERNAL_PARSER_WBMP As Boolean = True
Private Const USE_INTERNAL_PARSER_XBM As Boolean = True
Private Const USE_INTERNAL_PARSER_XCF As Boolean = True

'PNGs get some special preference due to their ubiquity; a persistent class enables better caching
Private m_PNG As pdPNG

'Similarly, JPEG auto-rotate behavior is persistently cached
Private m_JpegObeyEXIFOrientation As PD_BOOL

'Some user preferences control how image importing behaves.  Because these preferences are accessed frequently, we cache them
' locally improve performance.  External functions should use our wrappers instead of accessing the preferences directly.
' Also, changes to these preferences obviously require a re-cache; use the reset function, below, for that.
Public Sub ResetImageImportPreferenceCache()
    m_JpegObeyEXIFOrientation = PD_BOOL_UNKNOWN
End Sub

Public Function GetImportPref_JPEGOrientation() As Boolean
    If (m_JpegObeyEXIFOrientation = PD_BOOL_UNKNOWN) Then
        If UserPrefs.GetPref_Boolean("Loading", "ExifAutoRotate", True) Then m_JpegObeyEXIFOrientation = PD_BOOL_TRUE Else m_JpegObeyEXIFOrientation = PD_BOOL_FALSE
    End If
    GetImportPref_JPEGOrientation = (m_JpegObeyEXIFOrientation = PD_BOOL_TRUE)
End Function

'PDI loading.  PDI is PhotoDemon's native format, e.g. PDI is to PhotoDemon what PSD is to PhotoShop,
' or XCF to GIMP.
'
'Note the unique "sourceIsUndoFile" parameter for this load function.  PDI files are used
' to store undo/redo data, and when a PDI file is loaded as part of an Undo/Redo action,
' we deliberately ignore certain segments in the file (e.g. settings like "LastSaveFormat"
' which we do not want to Undo/Redo).  This parameter is passed to the pdImage initializer,
' and it tells it to ignore certain settings.
Public Function LoadPDI_Normal(ByRef pdiPath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByVal sourceIsUndoFile As Boolean = False) As Boolean
    
    PDDebug.LogAction "PDI file identified.  Starting pdPackage decompression..."
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    On Error GoTo LoadPDIFail
    
    'PDI files require a parent pdImage container
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    
    'First things first: create a pdPackage instance.
    Dim pdiReader As pdPackageChunky
    Set pdiReader = New pdPackageChunky
    
    'Load the file.  If this step fails, it means the file is not a modern PDI file.
    ' We want to drop back to legacy loaders and avoid further handling.
    If (Not pdiReader.OpenPackage_File(pdiPath, "PDIF")) Then
        PDDebug.LogAction "Legacy PDI file encountered; dropping back to pdPackage v2 functions..."
        Set pdiReader = Nothing
        LoadPDI_Normal = LoadPDI_LegacyV2(pdiPath, dstDIB, dstImage, sourceIsUndoFile)
        Exit Function
    End If
    
    'Still here?  The file validated successfully.
    PDDebug.LogAction "pdPackage successfully read and initialized.  Starting package parsing..."
    
    'The first chunk in the file must *always* be an image header chunk ("IHDR").  Validate it before continuing.
    If (pdiReader.GetChunkName() <> "IHDR") Then
        PDDebug.LogAction "First chunk in file is not IHDR; attempting legacy loader..."
        Set pdiReader = Nothing
        LoadPDI_Normal = LoadPDI_LegacyV2(pdiPath, dstDIB, dstImage, sourceIsUndoFile)
        Exit Function
    End If
    
    'Retrieve the image header and initialize a pdImage object against it
    Dim chunkName As String, chunkLength As Long, chunkData As pdStream
    If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
    
        If (Not dstImage.SetHeaderFromXML(chunkData.ReadString_UTF8(chunkLength), sourceIsUndoFile)) Then
            PDDebug.LogAction "pdImage failed to interpret IHDR string; abandoning import."
            LoadPDI_Normal = False
            Exit Function
        End If
    
    Else
        PDDebug.LogAction "Failed to read IHDR chunk; abandoning import."
        LoadPDI_Normal = False
        Exit Function
    End If
    
    'If we're still here, the base pdImage object initialized successfully.  Load all remaining chunks,
    ' skipping ones that we don't know how to interpret.
    PDDebug.LogAction "pdImage created successfully.  Moving on to individual layers..."
    
    'Raster layers will be decompressed directly into their primary buffer.
    Dim tmpDIBPointer As Long, tmpDIBLength As Long
    
    Dim curLayerIndex As Long, layerInitializedOK As Boolean
    curLayerIndex = 0
    layerInitializedOK = False
    
    Do While pdiReader.ChunksRemain()
            
        chunkName = UCase$(pdiReader.GetChunkName())
        chunkLength = pdiReader.GetChunkDataSize()
        
        Select Case chunkName
        
            'Layer header
            Case "LHDR"
            
                'Initialize the target layer against the chunk data
                If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
                    layerInitializedOK = dstImage.GetLayerByIndex(curLayerIndex).CreateNewLayerFromXML(chunkData.ReadString_UTF8(chunkLength))
                    If (Not layerInitializedOK) Then PDDebug.LogAction "WARNING! Layer #" & curLayerIndex & " failed to initialize header."
                Else
                    PDDebug.LogAction "WARNING! Layer #" & curLayerIndex & " failed to read header."
                End If
            
            'Layer data (raster or vector)
            Case "LDAT"
            
                'Hopefully this layer was already initialized successfully!
                If layerInitializedOK Then
                    
                    layerInitializedOK = False
                    
                    'Raster vs vector layers are initialized differently.
                    If dstImage.GetLayerByIndex(curLayerIndex).IsLayerRaster Then
                    
                        'Decompress directly into a DIB buffer
                        If (Not dstImage.GetLayerByIndex(curLayerIndex).GetLayerDIB Is Nothing) Then
                            dstImage.GetLayerByIndex(curLayerIndex).GetLayerDIB.SetInitialAlphaPremultiplicationState True
                            dstImage.GetLayerByIndex(curLayerIndex).GetLayerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                            If (tmpDIBLength <> 0) Then
                                layerInitializedOK = pdiReader.GetNextChunk(chunkName, chunkLength, Nothing, tmpDIBPointer, tmpDIBLength)
                                If (Not layerInitializedOK) Then PDDebug.LogAction "WARNING!  Layer bitmap wasn't retrieved because GetNextChunk failed!"
                            Else
                                PDDebug.LogAction "WARNING!  Layer bitmap wasn't retrieved because target pointer was null!"
                            End If
                        Else
                            PDDebug.LogAction "WARNING!  Layer bitmap wasn't retrieved because target buffer wasn't initialized!"
                        End If
                        
                    ElseIf dstImage.GetLayerByIndex(curLayerIndex).IsLayerVector Then
                    
                        'Vector layers are stored as lightweight XML.  Retrieve the string now.
                        If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
                            layerInitializedOK = dstImage.GetLayerByIndex(curLayerIndex).SetVectorDataFromXML(chunkData.ReadString_UTF8(chunkLength))
                            If (Not layerInitializedOK) Then PDDebug.LogAction "WARNING!  Vector layer couldn't parse XML packet."
                        Else
                            PDDebug.LogAction "WARNING! Layer #" & curLayerIndex & " failed to read vector string."
                        End If
                        
                    'Other layer types are not currently supported
                    Else
                        PDDebug.LogAction "WARNING!  Unknown layer type found??"
                    End If
                    
                    'Reset the "layer has been initialized" flag and advance the layer index tracker
                    If layerInitializedOK Then
                        curLayerIndex = curLayerIndex + 1
                        layerInitializedOK = False
                    Else
                        PDDebug.LogAction "WARNING!  Layer # " & curLayerIndex & " did not load its data chunk."
                    End If
                    
                'Encountering a layer data chunk when no corresponding layer header has been initialized
                ' is (obviously) a failure state.
                Else
                    PDDebug.LogAction "WARNING!  Layer data chunk found but no layer header was encountered first?"
                End If
            
            'Metadata chunks are messy; we just offload them to the metadata engine at present
            'ExifTool metadata chunk (a bare XML packet as received directly from ExifTool)
            Case "MDET"
                If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
                    If (Not dstImage.ImgMetadata.LoadAllMetadata(chunkData.ReadString_UTF8(chunkLength), dstImage.imageID, sourceIsUndoFile)) Then
                        PDDebug.LogAction "WARNING: MDET metadata chunk rejected by metadata parser."
                    End If
                Else
                    PDDebug.LogAction "WARNING!  Failed to retrieve MDET metadata chunk."
                End If
            
            'Serialized, post-parsing metadata struct
            Case "MDPD"
                If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
                    dstImage.ImgMetadata.RecreateFromSerializedXMLData chunkData.ReadString_UTF8(chunkLength)
                Else
                    PDDebug.LogAction "WARNING!  Failed to retrieve MDPD metadata chunk."
                End If
                
            'Any other chunks are unknown; just skip 'em
            Case Else
                pdiReader.SkipToNextChunk
            
        End Select
        
    'Continue parsing chunks until none remain
    Loop
    
    PDDebug.LogAction "PDI parsing complete.  Returning control to main image loader..."
    PDDebug.LogAction "(Time required to load PDI file: " & VBHacks.GetTimeDiffNowAsString(startTime) & ")"
    
    'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
    ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateBlank 16, 16, 32, 0
    
    'That's all there is to it!  Mark the load as successful and carry on.
    LoadPDI_Normal = True
    
    Exit Function
    
LoadPDIFail:
    
    PDDebug.LogAction "WARNING!  LoadPDIFail error routine reached.  Checking for known error states..."
    
    'Before falling back to a generic error message, check for a couple known problem states.
    Message "An error has occurred (#%1 - %2).  PDI load abandoned.", Err.Number, Err.Description
    
    LoadPDI_Normal = False
    Exit Function

End Function

'Load just the layer stack from a standard PDI file, and non-destructively align our current layer stack to match.
' At present, this function is only used internally by the Undo/Redo engine.
Public Function LoadPDI_HeadersOnly(ByRef pdiPath As String, ByRef dstImage As pdImage) As Boolean
    
    On Error GoTo LoadPDIHeaderFail
    
    'First things first: create a pdPackage instance.
    Dim pdiReader As pdPackageChunky
    Set pdiReader = New pdPackageChunky
    
    'Load the file.  If this step fails, it means the file is not a modern PDI file.
    ' We want to drop back to legacy loaders and avoid further handling.
    If (Not pdiReader.OpenPackage_File(pdiPath, "PDIF")) Then
        PDDebug.LogAction "LoadPDI_HeadersOnly failed - target file is not PDI?"
        Exit Function
    End If
    
    'Still here?  The file validated successfully.
    PDDebug.LogAction "pdPackage successfully read and initialized.  Starting package parsing..."
    
    'The first chunk in the file must *always* be an image header chunk ("IHDR").  Validate it before continuing.
    If (pdiReader.GetChunkName() <> "IHDR") Then
        PDDebug.LogAction "First chunk in file is not IHDR; abandoning load..."
        Exit Function
    End If
    
    'Retrieve the image header and NON-DESTRUCTIVELY initialize a pdImage object against it.
    ' (Non-destructively means no layers are destroyed in the process - we need them around because
    ' we're gonna be reusing them!)
    Dim chunkName As String, chunkLength As Long, chunkData As pdStream
    If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
    
        If (Not dstImage.SetHeaderFromXML(chunkData.ReadString_UTF8(chunkLength), True, True)) Then
            PDDebug.LogAction "pdImage failed to interpret IHDR string; abandoning import."
            LoadPDI_HeadersOnly = False
            Exit Function
        End If
    
    Else
        PDDebug.LogAction "Failed to read IHDR chunk; abandoning import."
        LoadPDI_HeadersOnly = False
        Exit Function
    End If
    
    'With the main pdImage now assembled, the next task is to populate all layer headers.  This is a bit more
    ' confusing than a regular PDI load, because we have to maintain existing layer DIB data (ugh!).
    
    'In a nutshell, we need to:
    ' 1) Extract each layer header from file, in turn
    ' 2) Compare each layer in the current image against the layer header data found in the file.  If header
    '    values are inconsistent (e.g. a layer is in the wrong z-order), we need to non-destructively move it
    '    to the index specified by the PDI file.
    ' 3) After moving the layer into place, we need to ask it to non-destructively overwrite its header with
    '    the header from the PDI file (e.g. don't touch they layer's pixel or vector data - only the header!).
    '
    'Note also that header-only files may include image metadata; this is also grabbed while here, as metadata
    ' changes fall under the "header-only undo required" banner.
    Dim layerHeaders As pdStringStack
    Set layerHeaders = New pdStringStack
    
    Dim mdFound As Boolean
    mdFound = False
    
    Do While pdiReader.ChunksRemain()
            
        chunkName = UCase$(pdiReader.GetChunkName())
        chunkLength = pdiReader.GetChunkDataSize()
        
        Select Case chunkName
    
            'Layer header
            Case "LHDR"
            
                'Retrieve the XML but don't do anything with it just yet; instead, just cache it locally.
                If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
                    layerHeaders.AddString chunkData.ReadString_UTF8(chunkLength)
                Else
                    PDDebug.LogAction "WARNING! Layer #" & layerHeaders.GetNumOfStrings() & " failed to read header."
                End If
                
            'Metadata chunks come in two varieties
            
            'ExifTool metadata chunk (a bare XML packet as received directly from ExifTool)
            Case "MDET"
                If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
                    mdFound = dstImage.ImgMetadata.LoadAllMetadata(chunkData.ReadString_UTF8(chunkLength), dstImage.imageID, True)
                    If (Not mdFound) Then
                        PDDebug.LogAction "WARNING: MDET metadata chunk rejected by metadata parser."
                    End If
                Else
                    PDDebug.LogAction "WARNING!  Failed to retrieve MDET metadata chunk."
                End If
            
            'Serialized, post-parsing metadata struct.  This only exists if the user has edited metadata
            ' manually during the current session.
            Case "MDPD"
                If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
                    dstImage.ImgMetadata.RecreateFromSerializedXMLData chunkData.ReadString_UTF8(chunkLength)
                    mdFound = True
                Else
                    'Metadata editing is actually pretty rare
                    'PDDebug.LogAction "WARNING!  Failed to retrieve MDPD metadata chunk."
                End If
                
            'Any other chunks are unknown; just skip 'em
            Case Else
                pdiReader.SkipToNextChunk
    
        End Select
        
    'Keep iterating chunks
    Loop
    
    'We now have a copy of all layer header data as it appears in the file (yay?).
    
    'We now need to iterate through the collection and retrieve a corresponding layer ID from each XML string.
    ' This allows us to match headers from the file against headers in the current pdImage object, and detect
    ' any mismatches in z-order.
    Dim xmlReader As pdSerialize
    Set xmlReader = New pdSerialize
    
    Dim layerIDs As pdStack
    Set layerIDs = New pdStack
    
    Dim i As Long
    For i = 0 To layerHeaders.GetNumOfStrings - 1
        xmlReader.SetParamString layerHeaders.GetString(i)
        layerIDs.AddInt xmlReader.GetLong("layer-id", , True)
    Next i
    
    'We now have a collection of all layer headers, and their corresponding layer IDs.
    
    'Our last job is to compare each discovered ID (and its corresponding index) against the current
    ' layer collection.  Mismatches must be manually resolved.
    Dim curLayerID As Long
    
    For i = 0 To dstImage.GetNumOfLayers - 1
        
        'Retrieve the ID for this index
        curLayerID = layerIDs.GetInt(i)
        
        'Ensure the layer is in its correct position
        If (dstImage.GetLayerIndexFromID(curLayerID) <> i) Then dstImage.SwapTwoLayers dstImage.GetLayerIndexFromID(curLayerID), i
        
        'Forcibly overwrite the layer's header data with whatever we retrieved from file
        If (Not dstImage.GetLayerByIndex(i).CreateNewLayerFromXML(layerHeaders.GetString(i), , True)) Then
            PDDebug.LogAction "WARNING! Layer #" & i & " failed to initialize header."
        End If
        
    'Repeat for all layers
    Next i
    
    'Finally, cover the case of metadata modifications.  If metadata has been removed from the image,
    ' no metadata chunks will appear in the file.  Erase any metadata now.
    If (Not mdFound) Then dstImage.ImgMetadata.Reset
    
    LoadPDI_HeadersOnly = True
    Exit Function
    
LoadPDIHeaderFail:
    PDDebug.LogAction "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI header-only load abandoned."
    LoadPDI_HeadersOnly = False
    
End Function

'Load a single layer from a standard PDI file.
'
'This function is only used internally by the Undo/Redo engine.  If the nearest diff to a layer-specific
' change is a full pdImage stack, this function is used to extract only the relevant layer (or layer header)
' from the target undo/redo file.
Public Function LoadPDI_SingleLayer(ByRef pdiPath As String, ByRef dstLayer As pdLayer, ByVal targetLayerID As Long, Optional ByVal loadHeaderOnly As Boolean = False) As Boolean
    
    On Error GoTo LoadLayerFromPDIFail
    LoadPDI_SingleLayer = False
    
    'Before doing anything else, load and validate the target file.
    Dim pdiReader As pdPackageChunky
    Set pdiReader = New pdPackageChunky
    
    If (Not pdiReader.OpenPackage_File(pdiPath, "PDIF")) Then
        PDDebug.LogAction "LoadPDI_SingleLayer() failed - target file isn't in PDI format?"
        Set pdiReader = Nothing
        Exit Function
    End If
    
    'The first chunk in the file must *always* be an image header chunk ("IHDR").  Validate it before continuing.
    If (pdiReader.GetChunkName() <> "IHDR") Then
        PDDebug.LogAction "First chunk in file is not IHDR; abandoning load..."
        Exit Function
    End If
    
    'Still here?  The file appears valid.
    Dim chunkName As String, chunkLength As Long, chunkData As pdStream, chunkLoaded As Boolean
    chunkLoaded = False
    
    Dim tmpString As String, xmlReader As pdSerialize
    Set xmlReader = New pdSerialize
    
    'Start iterating chunks in the file, looking for layer header chunks specifically.
    Do While pdiReader.ChunksRemain()
        
        chunkName = UCase$(pdiReader.GetChunkName())
        
        'Is this a layer header chunk?
        If (chunkName = "LHDR") Then
        
            'Load the underlying XML directly into a parser
            If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
                
                tmpString = chunkData.ReadString_UTF8(chunkLength)
                xmlReader.SetParamString tmpString
                
                'Look for the layer ID we were passed
                If (xmlReader.GetLong("layer-id", -1, True) = targetLayerID) Then
                    
                    'This header belongs to the layer in question.  Initialize the layer against it.
                    If dstLayer.CreateNewLayerFromXML(tmpString, , loadHeaderOnly) Then
                    
                        'The target layer has now been rebuilt using the data from the target PDI file.
                        
                        'If this is not a header-only load, we now need to iterate chunks until we encounter
                        ' this layer's data chunk (LDAT).  The spec requires the next LDAT in the file to
                        ' be the one belonging to this header.
                        If loadHeaderOnly Then
                            chunkLoaded = True
                        Else
                            
                            chunkLoaded = False
                            Do While pdiReader.ChunksRemain()
                                
                                chunkName = UCase$(pdiReader.GetChunkName())
                                If (chunkName = "LDAT") Then
                                
                                    'This is the data chunk we want!
                                    chunkLoaded = True
                                    
                                    'Raster vs vector layers are initialized differently.
                                    If dstLayer.IsLayerRaster Then
                                    
                                        'Decompress directly into a DIB buffer
                                        If (Not dstLayer.GetLayerDIB Is Nothing) Then
                                            
                                            Dim tmpDIBPointer As Long, tmpDIBLength As Long
                                            dstLayer.GetLayerDIB.SetInitialAlphaPremultiplicationState True
                                            dstLayer.GetLayerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                                            
                                            If (tmpDIBPointer <> 0) Then
                                                chunkLoaded = pdiReader.GetNextChunk(chunkName, chunkLength, Nothing, tmpDIBPointer, tmpDIBLength)
                                                If (Not chunkLoaded) Then PDDebug.LogAction "WARNING!  Layer bitmap wasn't retrieved because GetNextChunk failed!"
                                            Else
                                                PDDebug.LogAction "WARNING!  Layer bitmap wasn't retrieved because target pointer was null!"
                                            End If
                                            
                                        Else
                                            PDDebug.LogAction "WARNING!  Layer bitmap wasn't retrieved because target buffer wasn't initialized!"
                                        End If
                                        
                                    ElseIf dstLayer.IsLayerVector Then
                                    
                                        'Vector layers are stored as lightweight XML.  Retrieve the string now.
                                        If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
                                            chunkLoaded = dstLayer.SetVectorDataFromXML(chunkData.ReadString_UTF8(chunkLength))
                                            If (Not chunkLoaded) Then PDDebug.LogAction "WARNING!  Vector layer couldn't parse XML packet."
                                        Else
                                            PDDebug.LogAction "WARNING! Layer failed to read vector string."
                                        End If
                                        
                                    'Other layer types are not currently supported
                                    Else
                                        PDDebug.LogAction "WARNING!  Unknown layer type found??"
                                    End If
                                    
                                    'Regardless of success/failure, exit this loop
                                    Exit Do
                                    
                                End If
            
                            Loop
                            
                            'If we finished loading the chunk, exit the loop completely
                            If chunkLoaded Then
                                dstLayer.NotifyOfDestructiveChanges
                                Exit Do
                            End If
                        
                        End If
                        
                    'Failsafe only
                    Else
                        PDDebug.LogAction "WARNING!  LoadPDI_SingleLayer() couldn't initialize layer header."
                    End If
                
                'No "else" branch required.  (If the layer ID doesn't match, we just want to keep iterating chunks.)
                End If
                
            Else
                PDDebug.LogAction "WARNING! Bad layer header found in " & pdiPath
            End If
            
        'If this isn't a layer header, skip ahead to the next one
        Else
            pdiReader.SkipToNextChunk
        End If
        
    'Keep searching for the header we want
    Loop
    
    'That's all there is to it!  Mark the load as successful and carry on.
    LoadPDI_SingleLayer = True
    
    Exit Function
    
LoadLayerFromPDIFail:
    Message "An error has occurred (#%1 - %2).  PDI load abandoned.", Err.Number, Err.Description
    LoadPDI_SingleLayer = False
    Exit Function

End Function

'Load a single PhotoDemon layer from a standalone pdLayer file (which is really just a modified PDI file).
' This function is only used internally by the Undo/Redo engine.  Its counterpart is SavePDI_SingleLayer in
' the Saving module; any changes there must be mirrored here.
Private Function LoadPDLayer(ByVal pdiPath As String, ByRef dstLayer As pdLayer, Optional ByVal loadHeaderOnly As Boolean = False) As Boolean
    
    On Error GoTo LoadPDLayerFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackageChunky
    Set pdiReader = New pdPackageChunky
    
    'Load the file.  The reader uses memory-mapped file I/O, so do not modify the file until the
    ' read process completes.  (Note that this step will also validate the incoming file.)
    If pdiReader.OpenPackage_File(pdiPath, "UNDO") Then
    
        Dim chunkName As String, chunkLength As Long, chunkData As pdStream, chunkLoaded As Boolean
        Dim layerHeaderFound As Boolean
        
        'Iterate chunks, looking for a layer header
        Do While pdiReader.ChunksRemain()
            
            chunkLoaded = False
            chunkName = pdiReader.GetChunkName()
            chunkLength = pdiReader.GetChunkDataSize()
            
            'Layer header.  Note that we'll pull the chunk data into a dedicated stream before converting
            ' it from UTF-8; this is simply a convenience.
            If (chunkName = "LHDR") Then
                If pdiReader.GetNextChunk(chunkName, chunkLength, chunkData) Then
                    layerHeaderFound = True
                    dstLayer.CreateNewLayerFromXML chunkData.ReadString_UTF8(chunkLength), , loadHeaderOnly
                    chunkLoaded = True
                End If
            End If
            
            'Layer raster/vector data (only if "loadHeaderOnly" is NOT set).
            If (Not loadHeaderOnly) And layerHeaderFound And (chunkName = "LDAT") Then
                
                'How we extract this data varies by layer type.  Raster layers can skip the need for a temporary buffer, because we've
                ' already created a DIB with a built-in buffer for the pixel data.
                '
                'Other layer types (e.g. vector layers) are tiny so a temporary buffer does not matter; also, unlike raster buffers, we cannot
                ' easily predict the buffer size in advance, so we rely on pdPackage to do it for us
                Dim nodeLoadedSuccessfully As Boolean
                nodeLoadedSuccessfully = False
                
                If dstLayer.IsLayerRaster Then
                
                    'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                    Dim tmpDIBPointer As Long, tmpDIBLength As Long
                    dstLayer.GetLayerDIB.SetInitialAlphaPremultiplicationState True
                    dstLayer.GetLayerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                    
                    'Because we already know the decompressed size of the pixel data, we don't need to
                    ' double-allocate it - instead, decompress it directly from its (memory-mapped)
                    ' source into the already-allocated pixel container.
                    nodeLoadedSuccessfully = pdiReader.GetNextChunk(chunkName, chunkLength, , tmpDIBPointer, tmpDIBLength)
                
                Else
                    nodeLoadedSuccessfully = pdiReader.GetNextChunk(chunkName, chunkLength, chunkData)
                    nodeLoadedSuccessfully = dstLayer.SetVectorDataFromXML(chunkData.ReadString_UTF8(chunkLength))
                End If
                
                'If the load was successful, notify the target layer that its DIB data has been changed; the layer will use this to
                ' regenerate various internal caches.
                If nodeLoadedSuccessfully Then
                    dstLayer.NotifyOfDestructiveChanges
                    
                'Failure means package bytes could not be read, or alternately, checksums didn't match.  (Note that checksums are currently
                ' disabled for this function, for performance reasons, but I'm leaving this check in case we someday decide to re-enable them.)
                Else
                    PDDebug.LogAction "LoadPDLayer: node was not loaded successfully."
                End If
                
                chunkLoaded = True
                
            End If
            
            'Ensure we moved forward at least one chunk
            If (Not chunkLoaded) Then pdiReader.SkipToNextChunk
            
        Loop
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPDLayer = True
    
    'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
    Else
        PDDebug.LogAction "LoadPDLayer: file didn't pass validation."
        LoadPDLayer = False
    End If
    
    Exit Function
    
LoadPDLayerFail:
    PDDebug.LogAction "LoadPDLayer: VB error #" & Err.Number & ": " & Err.Description
    LoadPDLayer = False
    Exit Function

End Function

'Use GDI+ to load an image.  This does very minimal error checking (which is a no-no with GDI+) but because it's only a
' fallback when FreeImage can't be found, I'm postponing further debugging for now.
'Used for PNG and TIFF files if FreeImage cannot be located.
Public Function LoadGDIPlusImage(ByVal imagePath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByRef numOfPages As Long = 1, Optional ByVal nonInteractiveMode As Boolean = False, Optional ByVal overrideParameters As String = vbNullString) As Boolean
    
    LoadGDIPlusImage = False
    
    If GDI_Plus.GDIPlusLoadPicture(imagePath, dstDIB, dstImage, numOfPages, nonInteractiveMode, overrideParameters) Then
        If (Not dstDIB Is Nothing) Then LoadGDIPlusImage = ((dstDIB.GetDIBWidth <> 0) And (dstDIB.GetDIBHeight <> 0))
    End If
    
End Function

'SVG support is primarily handled by the 3rd-party resvg library
Public Function LoadSVG(ByVal imagePath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByVal overrideParameters As String = vbNullString, Optional ByRef userCanceledImportDialog As Boolean = False) As Boolean
    
    On Error GoTo LoadSVGFail
    
    'For now, we rely on proper file extensions before handing data off to resvg
    If Plugin_resvg.IsResvgEnabled() Then
        
        LoadSVG = Plugin_resvg.LoadSVG_FromFile(imagePath, dstImage, dstDIB, False, overrideParameters, userCanceledImportDialog)
        
        'If successful, set format-specific flags in the parent pdImage object
        If LoadSVG Then
        
            dstImage.SetOriginalFileFormat PDIF_SVG
            dstImage.NotifyImageChanged UNDO_Everything
            dstImage.SetOriginalColorDepth 32
            dstImage.SetOriginalGrayscale False
            If (dstImage.GetDPI <= 0) Then dstImage.SetDPI 96, 96
            
            'Assume alpha is present on 32-bpp images; assume it is *not* present if the SVG fills
            ' the entire visible area.
            dstImage.SetOriginalAlpha DIBs.IsDIBTransparent(dstDIB)
            
            'SVG files don't currently support color management
            dstDIB.SetColorManagementState cms_ProfileConverted
            
        Else
            PDDebug.LogAction "resvg failed; SVG load abandoned"
        End If
        
    Else
        LoadSVG = False
    End If
    
    Exit Function
    
LoadSVGFail:
    PDDebug.LogAction "WARNING!  SVG parsing failed with error #" & Err.Number & ": " & Err.Description
    LoadSVG = False
    
End Function

'Load data from a PD-generated Undo file.  This function is fairly complex, on account of PD's new diff-based Undo engine.
' Note that two types of Undo data must be specified: the Undo type of the file requested (because this function has no
' knowledge of that, by design), and what type of Undo data the caller wants extracted from the file.
'
'New as of 11 July '14 is the ability to specify a custom layer destination, for layer-relevant load operations.
' If this value is NOTHING, the function will automatically load the data to the relevant layer in the parent pdImage object.
' If a pdLayer object is supplied, however, it will be used instead.
Public Sub LoadUndo(ByVal undoFile As String, ByVal undoTypeOfFile As Long, ByVal undoTypeOfAction As PD_UndoType, Optional ByVal targetLayerID As Long = -1, Optional ByVal suspendRedraw As Boolean = False, Optional ByRef customLayerDestination As pdLayer = Nothing)
    
    'Certain load functions require access to a DIB, so declare a generic one in advance
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'If selection data was loaded as part of this diff, this value will be set to TRUE.  We check it at the end of
    ' the load function, and activate various selection-related items as necessary.
    Dim selectionDataLoaded As Boolean
    selectionDataLoaded = False
    
    'Regardless of outcome, notify the parent image of this change
    PDImages.GetActiveImage.NotifyImageChanged undoTypeOfAction, PDImages.GetActiveImage.GetLayerIndexFromID(targetLayerID)
    
    'Depending on the Undo data requested, we may end up loading one or more diff files at this location
    Select Case undoTypeOfAction
    
        'UNDO_EVERYTHING: a full copy of both the pdImage stack and all selection data is wanted
        Case UNDO_Everything
            ImageImporter.LoadPDI_Normal undoFile, tmpDIB, PDImages.GetActiveImage(), True
            PDImages.GetActiveImage.MainSelection.ReadSelectionFromFile undoFile & ".selection"
            Tools.NotifyImageSizeChanged
            selectionDataLoaded = True
            
        'UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE: a full copy of the pdImage stack is wanted
        '             Because the underlying file data must be of type UNDO_EVERYTHING or UNDO_IMAGE/_VECTORSAFE, we
        '             don't have to do any special processing to the file - just load the whole damn thing.
        Case UNDO_Image, UNDO_Image_VectorSafe
            ImageImporter.LoadPDI_Normal undoFile, tmpDIB, PDImages.GetActiveImage(), True
            Tools.NotifyImageSizeChanged
            
            'Once the full image has been loaded, we now know that at least the *existence* of all layers is correct.
            ' Unfortunately, subsequent changes to the pdImage header (or individual layers/layer headers) still need
            ' to be manually reconstructed, because they may have changed between the last full pdImage write and the
            ' current image state.  This step is handled by the Undo/Redo engine, which will call this LoadUndo function
            ' as many times as necessary to reconstruct each individual layer against its most recent diff.
        
        'UNDO_IMAGEHEADER: a full copy of the pdImage stack is wanted, but with all DIB data ignored (if present)
        '             For UNDO_IMAGEHEADER requests, we know the underlying file data is a PDI file.  We don't actually
        '             care if it has DIB data or not, because we'll just ignore it - but a special load function is
        '             required, due to the messy business of non-destructively aligning the current layer stack with
        '             the layer stack described by the file.
        Case UNDO_ImageHeader
            ImageImporter.LoadPDI_HeadersOnly undoFile, PDImages.GetActiveImage()
            Tools.NotifyImageSizeChanged
            
            'Once the full image has been loaded, we now know that at least the *existence* of all layers is correct.
            ' Unfortunately, subsequent changes to the pdImage header (or individual layers/layer headers) still need
            ' to be manually reconstructed, because they may have changed between the last full pdImage write and the
            ' current image state.  This step is handled by the Undo/Redo engine, which will call this LoadUndo function
            ' as many times as necessary to reconstruct each individual layer against its most recent diff.
        
        'UNDO_LAYER, UNDO_LAYER_VECTORSAFE: a full copy of the saved layer data at this position.
        '             Because the underlying file data can be different types (layer data can be loaded from standalone layer saves,
        '             or from a full pdImage stack save), we must check the undo type of the saved file, and modify our load
        '             behavior accordingly.
        Case UNDO_Layer, UNDO_Layer_VectorSafe
            
            'New as of 11 July '14 is the ability for the caller to supply their own destination layer for layer-specific Undo data.
            ' Check this optional parameter, and if it is NOT supplied, point it at the relevant layer in the parent pdImage object.
            If (customLayerDestination Is Nothing) Then Set customLayerDestination = PDImages.GetActiveImage.GetLayerByID(targetLayerID)
            
            'Layer data can appear in multiple types of Undo files
            Select Case undoTypeOfFile
            
                'The underlying save file is a standalone layer entry.  Simply overwrite the target layer with the data from the file.
                Case UNDO_Layer, UNDO_Layer_VectorSafe
                    ImageImporter.LoadPDLayer undoFile & ".layer", customLayerDestination, False
            
                'The underlying save file is a full pdImage stack.  Extract only the relevant layer data from the stack.
                Case UNDO_Everything, UNDO_Image, UNDO_Image_VectorSafe
                    ImageImporter.LoadPDI_SingleLayer undoFile, customLayerDestination, targetLayerID, False
                
            End Select
        
        'UNDO_LAYERHEADER: a full copy of the saved layer header data at this position.  Layer DIB data is ignored.
        '             Because the underlying file data can be many different types (layer data header can be loaded from
        '             standalone layer header saves, or full layer saves, or even a full pdImage stack), we must check the
        '             undo type of the saved file, and modify our load behavior accordingly.
        Case UNDO_LayerHeader
            
            'Layer header data can appear in multiple types of Undo files
            Select Case undoTypeOfFile
            
                'The underlying save file is a standalone layer entry.  Simply overwrite the target layer header with the
                ' header data from this file.
                Case UNDO_Layer, UNDO_Layer_VectorSafe, UNDO_LayerHeader
                    ImageImporter.LoadPDLayer undoFile & ".layer", PDImages.GetActiveImage.GetLayerByID(targetLayerID), True
            
                'The underlying save file is a full pdImage stack.  Extract only the relevant layer data from the stack.
                Case UNDO_Everything, UNDO_Image, UNDO_Image_VectorSafe, UNDO_ImageHeader
                    ImageImporter.LoadPDI_SingleLayer undoFile, PDImages.GetActiveImage.GetLayerByID(targetLayerID), targetLayerID, True
                
            End Select
        
        'UNDO_SELECTION: a full copy of the saved selection data is wanted
        '                 Because the underlying file data must be of type UNDO_EVERYTHING or UNDO_SELECTION, we don't have to do
        '                 any special processing.
        Case UNDO_Selection
            PDImages.GetActiveImage.MainSelection.ReadSelectionFromFile undoFile & ".selection"
            selectionDataLoaded = True
            
            
        'For now, any unhandled Undo types result in a request for the full pdImage stack.  This line can be removed when
        ' all Undo types finally have their own custom handling implemented.
        Case Else
            ImageImporter.LoadPDI_Normal undoFile, tmpDIB, PDImages.GetActiveImage(), True
            
        
    End Select
    
    'If a selection was loaded, activate all selection-related stuff now
    If selectionDataLoaded Then
    
        'Activate the selection as necessary
        PDImages.GetActiveImage.SetSelectionActive PDImages.GetActiveImage.MainSelection.IsLockedIn
        
        'Synchronize the text boxes as necessary
        SelectionUI.SyncTextToCurrentSelection PDImages.GetActiveImageID()
    
    End If
    
    'If a selection is active, request a redraw of the selection mask before rendering the image to the screen.  (If we are
    ' "undoing" an action that changed the image's size, the selection mask will be out of date.  Thus we need to re-render
    ' it before rendering the image or OOB errors may occur.)
    If PDImages.GetActiveImage.IsSelectionActive Then PDImages.GetActiveImage.MainSelection.RequestNewMask
    
    'Render the image to the screen, if requested
    If (Not suspendRedraw) Then Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

'Load a raw pdDIB file dump into the destination image and DIB.  (Note that pdDIB is likely compressed,
' so it is possible for this function to fail if the required compression library goes missing.)
Public Function LoadRawImageBuffer(ByVal imagePath As String, ByRef dstDIB As pdDIB, Optional ByRef dstImage As pdImage = Nothing) As Boolean

    On Error GoTo LoadRawImageBufferFail
    
    'Ask the destination DIB to create itself using the raw image buffer data
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    LoadRawImageBuffer = dstDIB.CreateFromFile(imagePath)
    If Not (dstImage Is Nothing) Then
        dstImage.Width = dstDIB.GetDIBWidth
        dstImage.Height = dstDIB.GetDIBHeight
    End If
    
    Exit Function
    
LoadRawImageBufferFail:
    
    Debug.Print "ERROR ENCOUNTERED IN ImageImporter.LoadRawImageBuffer: " & Err.Number & ", " & Err.Description
    LoadRawImageBuffer = False
    Exit Function

End Function

'Test an incoming image file against every supported decoder engine.  This ensures the greatest likelihood of loading
' a problematic file.
Public Function CascadeLoadGenericImage(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef freeImage_Return As PD_OPERATION_OUTCOME, ByRef decoderUsed As PD_ImageDecoder, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long, Optional ByVal overrideParameters As String = vbNullString, Optional ByRef userCanceledImportDialog As Boolean = False, Optional ByRef suspendWarnings As VbMsgBoxResult = vbNo) As Boolean
    
    CascadeLoadGenericImage = False
    
    'Before jumping out to a 3rd-party library, check for any image formats that we must decode using internal plugins.
    
    'PD's internal PNG/APNG parser is preferred for all PNG images.  For backwards compatibility reasons,
    ' it does *not* rely on the .png extension.  (Instead, it will manually verify the PNG signature,
    ' then work from there.)
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_PNG Then
        CascadeLoadGenericImage = LoadPNGOurselves(srcFile, dstImage, dstDIB, imageHasMultiplePages, numOfPages)
        If CascadeLoadGenericImage Then
            decoderUsed = id_PNGParser
            dstImage.SetOriginalFileFormat PDIF_PNG
        End If
    End If
    
    'OpenRaster support was added in v8.0.  OpenRaster is similar to ODF, basically a .zip wrapper
    ' around an XML file and a bunch of PNGs - easy enough to support!
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_ORA Then
        CascadeLoadGenericImage = LoadOpenRaster(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_ORAParser
            dstImage.SetOriginalFileFormat PDIF_ORA
        End If
    End If
    
    'PSD support was added in v8.0.
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_PSD Then
        CascadeLoadGenericImage = LoadPSD(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_PSDParser
            dstImage.SetOriginalFileFormat PDIF_PSD
        End If
    End If
    
    'PSP support was added in v9.0.
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_PSP Then
        CascadeLoadGenericImage = LoadPSP(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_PSPParser
            dstImage.SetOriginalFileFormat PDIF_PSP
        End If
    End If
    
    'JPEG XL support was added in v10.0
    If (Not CascadeLoadGenericImage) Then
        CascadeLoadGenericImage = LoadJXL(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_libjxl
            dstImage.SetOriginalFileFormat PDIF_JXL
        End If
    End If
    
    'HEIF/HEIC support (import only) was first added in v8.0, but loading required Win 10 and possible
    ' extra downloads from the MS Store.  As of PD 2024.8, libheif itself is used.
    If (Not CascadeLoadGenericImage) Then
        CascadeLoadGenericImage = LoadHEIF(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_libheif
            dstImage.SetOriginalFileFormat PDIF_HEIF
        End If
    End If
    
    'AVIF support was added in v9.0.
    If (Not CascadeLoadGenericImage) Then
        CascadeLoadGenericImage = LoadAVIF(srcFile, dstImage, dstDIB, imageHasMultiplePages, numOfPages, (suspendWarnings <> vbYes))
        If CascadeLoadGenericImage Then
            decoderUsed = id_libavif
            dstImage.SetOriginalFileFormat PDIF_AVIF
        End If
    End If
    
    'A custom ICO parser was added in v8.0.
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_ICO Then
        CascadeLoadGenericImage = LoadICO(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_ICOParser
            dstImage.SetOriginalFileFormat PDIF_ICO
        End If
    End If
    
    'A custom MBM parser was added in v9.0
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_MBM Then
        CascadeLoadGenericImage = LoadMBM(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_MBMParser
            dstImage.SetOriginalFileFormat PDIF_MBM
        End If
    End If
    
    'A custom CBZ parser was added in v9.0
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_CBZ Then
        CascadeLoadGenericImage = LoadCBZ(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_CBZParser
            dstImage.SetOriginalFileFormat PDIF_CBZ
        End If
    End If
    
    'JPEG-LS (via the CharLS library) support was added in v9.0
    If (Not CascadeLoadGenericImage) Then
        CascadeLoadGenericImage = LoadJLS(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_CharLS
            dstImage.SetOriginalFileFormat PDIF_JLS
        End If
    End If
    
    'WebP was originally handled by FreeImage, but in v9.0 I switched to using libwebp directly
    If (Not CascadeLoadGenericImage) And Plugin_WebP.IsWebPEnabled() Then
        If Plugin_WebP.IsWebP(srcFile) Then
            CascadeLoadGenericImage = LoadWebP(srcFile, dstImage, dstDIB)
            If CascadeLoadGenericImage Then
                decoderUsed = id_libwebp
                dstImage.SetOriginalFileFormat PDIF_WEBP
            End If
        End If
    End If
    
    'QOI support was added in v9.0
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_QOI Then
        CascadeLoadGenericImage = LoadQOI(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_QOIParser
            dstImage.SetOriginalFileFormat PDIF_QOI
        End If
    End If
        
    'SVG/Z support was added in v9.0
    If (Not CascadeLoadGenericImage) And Plugin_resvg.IsFileSVGCandidate(srcFile) Then
        CascadeLoadGenericImage = LoadSVG(srcFile, dstDIB, dstImage, overrideParameters, userCanceledImportDialog)
        If CascadeLoadGenericImage Then
            decoderUsed = id_resvg
            dstImage.SetOriginalFileFormat PDIF_SVG
        End If
    End If
        
    'GIMP XCF support was added in v9.0
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_XCF Then
        CascadeLoadGenericImage = LoadXCF(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_XCFParser
            dstImage.SetOriginalFileFormat PDIF_XCF
        End If
    End If
    
    'Shuttle Radar Topography Mission (SRTM) HGT format was added in v10.0
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_HGT Then
        CascadeLoadGenericImage = LoadHGT(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_HGTParser
            dstImage.SetOriginalFileFormat PDIF_HGT
        End If
    End If
    
    'PDF support was added in v10.0
    If (Not CascadeLoadGenericImage) Then
        CascadeLoadGenericImage = LoadPDF(srcFile, dstImage, dstDIB, False, False, userCanceledImportDialog)
        If CascadeLoadGenericImage Then
            decoderUsed = id_pdfium
            dstImage.SetOriginalFileFormat PDIF_PDF
        End If
    End If
    
    'A custom XBM parser was added in v10.0
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_XBM Then
        CascadeLoadGenericImage = LoadXBM(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_XBMParser
            dstImage.SetOriginalFileFormat PDIF_XBM
        End If
    End If
    
    'A custom WBMP parser was added in v10.0
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_WBMP Then
        CascadeLoadGenericImage = LoadWBMP(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_WBMPParser
            dstImage.SetOriginalFileFormat PDIF_WBMP
        End If
    End If
    
    'A custom PCX parser was added in v2025.4
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_PCX Then
        CascadeLoadGenericImage = LoadPCX(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_PCXParser
            dstImage.SetOriginalFileFormat PDIF_PCX
        End If
    End If
    
    'FreeImage claims to work with DDS, but it only parses a tiny subset of ancient DDS files.
    ' In v2025.5+, I added a workaround (on Win7+) using DirectXTex.
    If (Not CascadeLoadGenericImage) And Plugin_DDS.IsDirectXTexAvailable() Then
        CascadeLoadGenericImage = LoadDDS(srcFile, dstImage, dstDIB, imageHasMultiplePages, numOfPages, (suspendWarnings <> vbYes))
        If CascadeLoadGenericImage Then
            decoderUsed = id_DirectXTex
            dstImage.SetOriginalFileFormat PDIF_DDS
        End If
    End If
    
    'FreeImage claims to work with JPEG-2000, but its support is nearly a decade out-of-date
    ' and it fails to cover many edge-cases correctly.
    ' In v2025.9+, I added support for direct integration with OpenJPEG.
    If (Not CascadeLoadGenericImage) And Plugin_OpenJPEG.IsOpenJPEGEnabled() Then
        CascadeLoadGenericImage = LoadJPEG2000(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_OpenJPEG
            dstImage.SetOriginalFileFormat PDIF_JP2
        End If
    End If
    
    'Seeing a pattern?  FreeImage's HDR coverage is also problematic.  In v2026.1+ I wrote my own HDR parser.
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_HDR Then
        CascadeLoadGenericImage = LoadHDR(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_HDRParser
            dstImage.SetOriginalFileFormat PDIF_HDR
        End If
    End If
    
    'If our various internal engines passed on the image, we now want to attempt either FreeImage or GDI+.
    ' (Pre v8.0, we *always* tried FreeImage first, but as time goes by, I realize the library is prone to
    ' a number of esoteric bugs.  It also suffers performance-wise compared to GDI+.  As such, I am now
    ' more selective about which library gets used first.)
    If (Not CascadeLoadGenericImage) Then
    
        'FreeImage's TIFF support (via libTIFF?) is wonky.  It's prone to bad crashes and inexplicable
        ' memory issues (including allocation failures on normal-sized images), so for TIFFs we want to
        ' try GDI+ before trying FreeImage.  (PD's GDI+ image loader was heavily restructured in v8.0 to
        ' support things like multi-page import, so this strategy wasn't viable until then.)
        Dim tryGDIPlusFirst As Boolean
        tryGDIPlusFirst = Strings.StringsEqual(Files.FileGetExtension(srcFile), "tif", True) Or Strings.StringsEqual(Files.FileGetExtension(srcFile), "tiff", True)
        
        'On modern Windows builds (8+) FreeImage is markedly slower than GDI+ at loading JPEG images,
        ' so let's also default to GDI+ for JPEGs.
        If (Not tryGDIPlusFirst) Then
            If OS.IsWin7OrLater() Then tryGDIPlusFirst = Strings.StringsEqual(Files.FileGetExtension(srcFile), "jpg", True) Or Strings.StringsEqual(Files.FileGetExtension(srcFile), "jpeg", True)
        End If
        
        'GIFs are much faster via GDI+, but there are some known bugs parsing animated GIFs on XP.
        ' For now, I'm not really willing to write an XP-specific workaround; hopefully animated GIFs
        ' on XP is a rare use-case.
        If (Not tryGDIPlusFirst) Then tryGDIPlusFirst = Strings.StringsEqual(Files.FileGetExtension(srcFile), "gif", True)
        
        '32-bit BMP files give FreeImage trouble.  Unfortunately, BMP handling via GDI+ varies by OS version.
        ' (It's pretty good on Win 7+, but mediocre on XP.)
        ' There's no good answer here, but for now we'll rely on GDI+, mostly because FreeImage will "successfully"
        ' load some BMPs but only because it loads transparent surfaces at the target size, which makes it very
        ' hard to detect failures.
        If (Not tryGDIPlusFirst) Then tryGDIPlusFirst = Strings.StringsEqual(Files.FileGetExtension(srcFile), "bmp", True)
        
        'To disable these various heuristics and force either GDI+ import (TRUE) or FreeImage import (FALSE),
        ' use this override:
        'tryGDIPlusFirst = False
        
        If tryGDIPlusFirst Then
            CascadeLoadGenericImage = AttemptGDIPlusLoad(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages, overrideParameters)
            freeImage_Return = PD_FAILURE_GENERIC
            If (Not CascadeLoadGenericImage) And ImageFormats.IsFreeImageEnabled() Then CascadeLoadGenericImage = AttemptFreeImageLoad(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
            
        'For other formats, let FreeImage have a go at it, and we'll try GDI+ if it fails
        Else
            freeImage_Return = PD_FAILURE_GENERIC
            If (Not CascadeLoadGenericImage) And ImageFormats.IsFreeImageEnabled() Then
                
                'Disallow some formats where FreeImage has known bugs.
                Dim doNotAllowFI As Boolean
                doNotAllowFI = False
                
                'JPEG-2000 files crash FreeImage with some regularity
                If (Not doNotAllowFI) Then doNotAllowFI = Plugin_OpenJPEG.IsFileJP2(srcFile)
                
                'Some PCX and DCX images crash FreeImage too
                If (Not doNotAllowFI) Then doNotAllowFI = Strings.StringsEqualAny(Files.FileGetExtension(srcFile), True, "pcx", "dcx")
                
                'If this file looks safe, allow FreeImage to try it
                If (Not doNotAllowFI) Then CascadeLoadGenericImage = AttemptFreeImageLoad(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
                
            End If
            If (Not CascadeLoadGenericImage) And (freeImage_Return <> PD_FAILURE_USER_CANCELED) Then CascadeLoadGenericImage = AttemptGDIPlusLoad(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages, overrideParameters)
        End If
        
    End If
    
End Function

Private Function AttemptFreeImageLoad(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef freeImage_Return As PD_OPERATION_OUTCOME, ByRef decoderUsed As PD_ImageDecoder, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long) As Boolean
    
    AttemptFreeImageLoad = False
    
    'FreeImage has known crashes on certain image formats.  Do NOT attempt to load these!
    If Strings.StringsEqual(Files.FileGetExtension(srcFile), "xbm", True) Then Exit Function
    
    PDDebug.LogAction "Attempting to load via FreeImage..."
    
    'Start by seeing if the image file contains multiple pages.
    ' If it does, we will load each page as a separate layer.
    numOfPages = Plugin_FreeImage.IsMultiImage(srcFile)
    imageHasMultiplePages = (numOfPages > 1)
    freeImage_Return = FI_LoadImage_V5(srcFile, dstDIB, , , dstImage)
    AttemptFreeImageLoad = (freeImage_Return = PD_SUCCESS)
    
    'FreeImage worked!  Copy any relevant information from the DIB to the parent pdImage object (such as file format),
    ' then continue with the load process.
    If AttemptFreeImageLoad Then
        decoderUsed = id_FreeImage
        dstImage.SetOriginalFileFormat dstDIB.GetOriginalFormat
        dstImage.SetDPI dstDIB.GetDPI, dstDIB.GetDPI
        dstImage.SetOriginalColorDepth dstDIB.GetOriginalColorDepth
    End If
        
End Function

Private Function AttemptGDIPlusLoad(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef freeImage_Return As PD_OPERATION_OUTCOME, ByRef decoderUsed As PD_ImageDecoder, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long, Optional ByVal overrideParameters As String = vbNullString) As Boolean

    PDDebug.LogAction "Attempting to load via GDI+..."
    AttemptGDIPlusLoad = LoadGDIPlusImage(srcFile, dstDIB, dstImage, numOfPages, False, overrideParameters)
    
    If AttemptGDIPlusLoad Then
        decoderUsed = id_GDIPlus
        dstImage.SetOriginalFileFormat dstDIB.GetOriginalFormat
        dstImage.SetDPI dstDIB.GetDPI, dstDIB.GetDPI
        dstImage.SetOriginalColorDepth dstDIB.GetOriginalColorDepth
        imageHasMultiplePages = (numOfPages > 1)
    End If
        
End Function

'Test an incoming image file against PD's internal decoder engines.  This function is much faster than
' CascadeLoadGenericImage(), above, and it should be preferentially used for image files generated by PD itself.
Public Function CascadeLoadInternalImage(ByVal internalFormatID As Long, ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef freeImage_Return As PD_OPERATION_OUTCOME, ByRef decoderUsed As PD_ImageDecoder, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long) As Boolean
    
    Select Case internalFormatID
        
        Case PDIF_PDI
        
            'PDI images require various compression plugins to be present, and are only loaded via a custom routine
            ' (obviously, since they are PhotoDemon's native format)
            CascadeLoadInternalImage = LoadPDI_Normal(srcFile, dstDIB, dstImage)
            
            dstImage.SetOriginalFileFormat PDIF_PDI
            dstImage.SetOriginalColorDepth 32
            dstImage.NotifyImageChanged UNDO_Everything
            decoderUsed = id_PDIParser
            
        'TMPDIB files are raw pdDIB objects dumped directly to file.  In some cases, this is faster and easier for PD than wrapping
        ' the pdDIB object inside a pdPackage layer (e.g. during clipboard interactions, since we start with a raw pdDIB object
        ' after selections and such are applied to the base layer/image, so we may as well just use the raw pdDIB data we've cached).
        Case PDIF_RAWBUFFER
            
            'These raw pdDIB objects may require 3rd-party compression libraries for parsing
            ' (compression is optional), so it is possible for the load function to fail if
            ' zstd/lz4/libdeflate goes missing.
            CascadeLoadInternalImage = LoadRawImageBuffer(srcFile, dstDIB, dstImage)
            
            dstImage.SetOriginalFileFormat PDIF_UNKNOWN
            dstImage.SetOriginalColorDepth 32
            dstImage.NotifyImageChanged UNDO_Everything
            decoderUsed = id_PDIParser
            
        'Straight TMP files are internal files (BMP, typically) used by PhotoDemon.
        ' As ridiculous as it sounds, we must default to the generic load engine list,
        ' as the format of a TMP file is not guaranteed in advance.  Because of this,
        ' we can rely on the generic load engine to properly set things like
        ' "original color depth".
        '
        '(TODO: settle on a single tmp file format, so we don't have to play this game??)
        Case PDIF_TMPFILE
            CascadeLoadInternalImage = ImageImporter.CascadeLoadGenericImage(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
            dstImage.SetOriginalFileFormat PDIF_UNKNOWN
            
    End Select
    
End Function

Private Function LoadAVIF(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long, Optional ByVal allowErrorPopups As Boolean = False) As Boolean
    
    LoadAVIF = False
    On Error GoTo LibAVIFDidntWork
    
    'AVIF support was provisionally added in v9.0.  Loading requires 64-bit Windows and a copy of the official
    ' exe binaries (for example, https://github.com/AOMediaCodec/libavif/releases/tag/v0.9.0) inside the
    ' /App/PhotoDemon/Plugins subfolder.  PhotoDemon will offer to automatically download and configure a
    ' portable copy if the user interacts with the AVIF format in some way (import/export).
    Dim potentialAVIF As Boolean
    potentialAVIF = Strings.StringsEqualAny(Files.FileGetExtension(srcFile), True, "avci", "avcs", "avif", "avifs", "heic")
    
    'In 2025.3, I also added a quick check for AVIF file headers so we can reliably detect AVIF files with
    ' bad extensions (which would be missed by the check above).
    If (Not potentialAVIF) Then potentialAVIF = Plugin_AVIF.IsFilePotentiallyAVIF(srcFile)
    If potentialAVIF Then
        
        'If this system is 64-bit capable but libavif doesn't exist, ask if we can download a copy
        If OS.OSSupports64bitExe Then
            
            If (Not Plugin_AVIF.IsAVIFImportAvailable()) Then
                If (Not Plugin_AVIF.PromptForLibraryDownload_AVIF()) Then Exit Function
            
            'If the library *does* exist, check for updates
            Else
                Plugin_AVIF.CheckAVIFVersionAndOfferUpdates True
            End If
            
        End If
        
        If Plugin_AVIF.IsAVIFImportAvailable() Then
        
            'It's an ugly workaround, but necessary; convert the AVIF to a temporary image file
            ' in a format we can directly process (currently PNG).
            Dim tmpFile As String
            LoadAVIF = Plugin_AVIF.ConvertAVIFtoStandardImage(srcFile, tmpFile, allowErrorPopups)
            
            'If that worked, load the intermediary image (PNG format) using the relevant decoder
            If LoadAVIF Then LoadAVIF = LoadPNGOurselves(tmpFile, dstImage, dstDIB, imageHasMultiplePages, numOfPages)
            
            'Regardless of outcome, kill the temp file
            Files.FileDeleteIfExists tmpFile
            
            'If succcessful, flag the image format and return
            If LoadAVIF Then dstImage.SetOriginalFileFormat PDIF_AVIF
            
        End If
        
    End If
    
    Exit Function
    
LibAVIFDidntWork:
    LoadAVIF = False
    
End Function

Private Function LoadCBZ(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadCBZ = False
    
    'pdCBZhandles all the dirty work for us
    Dim cCBZ As pdCBZ
    Set cCBZ = New pdCBZ
    
    'Validate the potential comic book archive
    LoadCBZ = cCBZ.IsFileCBZ(srcFile, True)
    
    'If validation passes, attempt a full load
    If LoadCBZ Then
        PDDebug.LogAction "CBZ format found; loading pages..."
        LoadCBZ = cCBZ.LoadCBZ(srcFile, dstImage)
    End If
    
    'Perform some PD-specific object initialization before exiting
    If LoadCBZ Then
        
        dstImage.SetOriginalFileFormat PDIF_CBZ
        dstImage.NotifyImageChanged UNDO_Everything
        dstImage.SetOriginalColorDepth 32
        dstImage.SetOriginalGrayscale False
        dstImage.SetOriginalAlpha True
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank 16, 16, 32, 0
        dstDIB.SetColorManagementState cms_ProfileConverted
        
    End If
    
End Function

Public Function LoadDDS(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long, Optional ByVal allowErrorPopups As Boolean = False) As Boolean
    
    LoadDDS = False
    On Error GoTo LoadFailed
    
    'FreeImage claims to support DDS, but it only supports a tiny subset of ancient DDS files.
    ' In v2025.5+ I added DirectXTex as an alternate load path for DDS files.  This requires a copy
    ' of texconv.exe in the /App/PhotoDemon/Plugins subfolder.
    Dim potentialDDS As Boolean
    potentialDDS = Strings.StringsEqualAny(Files.FileGetExtension(srcFile), True, "dds")
    
    'Let's also do a quick check for DDS file headers so we can reliably detect DDS files with
    ' bad extensions (which would be missed by the check above).
    If (Not potentialDDS) Then potentialDDS = Plugin_DDS.IsFilePotentiallyDDS(srcFile)
    If potentialDDS Then
        
        If Plugin_DDS.IsDirectXTexAvailable() Then
        
            'It's an ugly workaround, but necessary; convert the DDS to a temporary image file
            ' in a format we can directly process (currently PNG).
            Dim tmpFile As String
            LoadDDS = Plugin_DDS.ConvertDDStoStandardImage(srcFile, tmpFile, allowErrorPopups)
            
            'If that worked, load the intermediary image (PNG format) using the relevant decoder
            If LoadDDS Then LoadDDS = LoadPNGOurselves(tmpFile, dstImage, dstDIB, imageHasMultiplePages, numOfPages, True)
            
            'Regardless of outcome, kill the temp file
            Files.FileDeleteIfExists tmpFile
            
            'If succcessful, flag the image format and return
            If LoadDDS Then dstImage.SetOriginalFileFormat PDIF_DDS
            
        End If
        
    End If
    
    Exit Function
    
LoadFailed:
    LoadDDS = False
    
End Function

Private Function LoadHDR(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadHDR = False
    
    'pdHDR handles all the dirty work for us
    Dim cReader As pdHDR
    Set cReader = New pdHDR
    
    'Validate and (potentially) load the file in one fell swoop
    LoadHDR = cReader.LoadHDR_FromFile(srcFile, dstImage, dstDIB)
    
    'Perform some PD-specific object initialization before exiting
    If LoadHDR Then

        'Set format flags and reset internal image caches
        dstImage.SetOriginalFileFormat PDIF_HDR
        dstImage.NotifyImageChanged UNDO_Everything

        'HDR files are always 96-bpp color (but we may "spruce them up a bit" during loading, for fun)
        dstImage.SetOriginalColorDepth 24
        dstImage.SetOriginalGrayscale False

        'HDR files never contain alpha
        dstImage.SetOriginalAlpha False

        'HDR files contain their own embedded color-space data, which is always parsed and applied during
        ' the mapping to 24-bpp RGB
        dstDIB.SetColorManagementState cms_ProfileConverted

    End If
    
End Function

Private Function LoadHGT(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadHGT = False
    
    'pdHGT handles all the dirty work for us
    Dim cReader As pdHGT
    Set cReader = New pdHGT
    
    'Validate and (potentially) load the file in one fell swoop
    LoadHGT = cReader.LoadHGT_FromFile(srcFile, dstImage, dstDIB)
    
    'Perform some PD-specific object initialization before exiting
    If LoadHGT Then
        
        'Set format flags and reset internal image caches
        dstImage.SetOriginalFileFormat PDIF_HGT
        dstImage.NotifyImageChanged UNDO_Everything
        
        'HGT files are always 24-bpp grayscale (but we may "spruce them up a bit" during loading, for fun)
        dstImage.SetOriginalColorDepth 24
        dstImage.SetOriginalGrayscale True  'Will need to vary in the future based on user import settings...
        
        'HGT files never contain alpha
        dstImage.SetOriginalAlpha False
        
        'HGT files are raw linear files; as such we don't need to color-manage them
        dstDIB.SetColorManagementState cms_ProfileConverted
        
    End If
    
End Function

Private Function LoadICO(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadICO = False
    
    'pdICO handles all the dirty work for us
    Dim cIconReader As pdICO
    Set cIconReader = New pdICO
    
    'Validate the potential icon file
    LoadICO = cIconReader.IsFileICO(srcFile, True)
    
    'If validation passes, attempt a full load
    If LoadICO Then LoadICO = (cIconReader.LoadICO(srcFile, dstImage, dstDIB) < ico_Failure)
    
    'Perform some PD-specific object initialization before exiting
    If LoadICO Then
        
        dstImage.SetOriginalFileFormat PDIF_ICO
        dstImage.NotifyImageChanged UNDO_Everything
        dstImage.SetOriginalColorDepth 32
        dstImage.SetOriginalGrayscale False
        dstImage.SetOriginalAlpha True
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank 16, 16, 32, 0
        dstDIB.SetColorManagementState cms_ProfileConverted
        
    End If
    
End Function

Private Function LoadJLS(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadJLS = False
    
    'Ensure the CharLS library is available
    If (Not Plugin_CharLS.IsCharLSEnabled()) Then Exit Function
    
    'For now, we perform basic validation against the file extension; this is primarily for performance
    ' reasons, as CharLS does not provide a "validation" function so we'd need to load the full source
    ' file into memory and we want to avoid that unless absolutely necessary.
    If Strings.StringsNotEqual(Files.FileGetExtension(srcFile), "jls", True) Then Exit Function
    
    'CharLS handles everything for us
    LoadJLS = Plugin_CharLS.LoadJLS(srcFile, dstImage, dstDIB)
    
    'Perform some PD-specific object initialization before exiting
    If LoadJLS Then
        dstImage.SetOriginalFileFormat PDIF_JLS
        dstImage.NotifyImageChanged UNDO_Everything
        dstImage.SetOriginalColorDepth 32       'TODO: retrieve this from file?
        dstImage.SetOriginalGrayscale False     'Same here?
        dstImage.SetOriginalAlpha True          'Same here?
    End If
    
End Function

Private Function LoadJPEG2000(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadJPEG2000 = False
    
    'OpenJPEG handles all the dirty work for us
    If (Not Plugin_OpenJPEG.IsFileJP2(srcFile)) Then Exit Function
    
    'If validation passes, attempt a full load
    LoadJPEG2000 = Plugin_OpenJPEG.LoadJP2(srcFile, dstImage, dstDIB)
    
    'Perform some PD-specific object initialization before exiting
    If LoadJPEG2000 Then
        
        dstImage.SetOriginalFileFormat PDIF_JP2
        dstImage.NotifyImageChanged UNDO_Everything
        
        'JP2 files can have arbitrary numbers of color components.  PD only cares about gray vs RGB vs RGBA
        Dim channelCount As Long
        channelCount = Plugin_OpenJPEG.GetComponentCountOfLastImage()
        If (channelCount <= 0) Then channelCount = 3
        If (channelCount = 2) Then channelCount = 1
        If (channelCount > 4) Then channelCount = 4
        dstImage.SetOriginalGrayscale (channelCount < 3)
        dstImage.SetOriginalAlpha (channelCount >= 4)
        
        'JP2 files can have arbitrary "precision" (bits-per-component).  We only want to distinguish between SDR and HDR.
        Dim finalPrec As Long
        finalPrec = Plugin_OpenJPEG.GetPrecisionOfLastImage()
        
        If (finalPrec <= 8) Then
            dstImage.SetOriginalColorDepth 8 * channelCount
        Else
            dstImage.SetOriginalColorDepth 16 * channelCount
        End If
        
    End If
    
End Function

Private Function LoadJXL(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadJXL = False
    
    'Ensure libjxl is available and functioning correctly (e.g. we are on Vista or later)
    If (Not Plugin_jxl.IsJXLImportAvailable()) Then
        
        'If the target file is a likely JPEG-XL candidate (by file extension),
        ' ask the user if they want to download libjxl.
        Dim fileMightBeJXL As Boolean
        fileMightBeJXL = Strings.StringsEqual(Files.FileGetExtension(srcFile), "jxl", True)
        If (Not fileMightBeJXL) Then fileMightBeJXL = Plugin_jxl.IsFileJXL_NoExternalLibrary(srcFile)
        
        If fileMightBeJXL Then
            
            'Prompt for download (if allowed) and attempt to continue
            If (Not Plugin_jxl.PromptForLibraryDownload_JXL(True)) Then Exit Function
            If (Not Plugin_jxl.IsJXLImportAvailable()) Then Exit Function
            
        'If this image isn't (likely) JPEG-XL, do not prompt
        Else
            Exit Function
        End If

    'If libjxl *is* available, check for plugin updates before proceeding
    Else
        
        'Before offering updates, see if the target file is even a JPEG XL image.
        ' (If it isn't, we don't want to spam the user with an irrelevant update notification.)
        If Plugin_jxl.IsFileJXL(srcFile) Then
            Plugin_jxl.CheckJXLVersionAndOfferUpdates True
        Else
            LoadJXL = False
            Exit Function
        End If
        
    End If
    
    'Offload the remainder of the job to the libjxl interface
    LoadJXL = Plugin_jxl.LoadJXL(srcFile, dstImage, dstDIB)
    
    'Perform some PD-specific object initialization before exiting
    If LoadJXL And (Not dstImage Is Nothing) Then
        
        'Note the correct image format and undo state (some of this is also handled by the upstream function,
        ' but I add it here for clarity because JPEG-XL files currently use PNG as an interop format).
        dstImage.SetOriginalFileFormat PDIF_JXL
        dstImage.NotifyImageChanged UNDO_Everything
        
        'Note that other load-time checks - like color-management - have already been handled by the LoadJXL function.
        
    End If
    
End Function

Private Function LoadMBM(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadMBM = False
    
    'pdMBM handles all the dirty work for us
    Dim cReader As pdMBM
    Set cReader = New pdMBM
    
    'Validate and (potentially) load the file in one fell swoop
    LoadMBM = cReader.LoadMBM_FromFile(srcFile, dstImage, dstDIB)
    
    'Perform some PD-specific object initialization before exiting
    If LoadMBM Then
        
        dstImage.SetOriginalFileFormat PDIF_MBM
        dstImage.NotifyImageChanged UNDO_Everything
        
        'Retrieve alpha, grayscale, and color-depth data from the pdMBM object
        If (cReader.GetColorDepth > 0) Then
            dstImage.SetOriginalColorDepth cReader.GetColorDepth
        Else
            dstImage.SetOriginalColorDepth 32
        End If
        
        dstImage.SetOriginalGrayscale cReader.IsMBMGrayscale()
        
        'Assume alpha is present on 32-bpp images; assume it is *not* present on lower bit-depths
        dstImage.SetOriginalAlpha (dstImage.GetOriginalColorDepth = 32)
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns
        ' a width/height of zero, the upstream load function will think the load process failed.
        ' Because of that, we must initialize the DIB to *something*.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank 16, 16, 32, 0
        
        'MBM files don't support color management
        dstDIB.SetColorManagementState cms_ProfileConverted
        
    End If
    
End Function

Private Function LoadOpenRaster(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadOpenRaster = False
    
    'pdOpenRaster handles all the dirty work for us
    Dim cOpenRaster As pdOpenRaster
    Set cOpenRaster = New pdOpenRaster
    
    'Validate the potential OpenRaster file
    LoadOpenRaster = cOpenRaster.IsFileORA(srcFile, True)
    
    'If validation passes, attempt a full load
    If LoadOpenRaster Then LoadOpenRaster = cOpenRaster.LoadORA(srcFile, dstImage)
    
    'Perform some PD-specific object initialization before exiting
    If LoadOpenRaster Then
        
        dstImage.SetOriginalFileFormat PDIF_ORA
        dstImage.NotifyImageChanged UNDO_Everything
        dstImage.SetOriginalColorDepth 32
        dstImage.SetOriginalGrayscale False
        dstImage.SetOriginalAlpha True
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank 16, 16, 32, 0
        dstDIB.SetColorManagementState cms_ProfileConverted
        
    End If
    
End Function

'Load a PCX image.  Originally this was handled by FreeImage, but FreeImage doesn't handle all PCXs correctly.
' So I wrote my own loader in Jan 2025.
Private Function LoadPCX(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadPCX = False
    
    'pdPCX handles all the dirty work for us
    Dim cReader As pdPCX
    Set cReader = New pdPCX
    
    'Validate the file (note: this will return TRUE for both PCX and DCX images)
    If cReader.IsFilePCX(srcFile) Then
    
        'Load the file
        LoadPCX = cReader.LoadPCX_FromFile(srcFile, dstImage, dstDIB)
        
        'Perform some PD-specific object initialization before exiting
        If LoadPCX Then
            
            dstImage.SetOriginalFileFormat PDIF_PCX
            dstImage.NotifyImageChanged UNDO_Everything
            dstImage.SetOriginalGrayscale cReader.HasGrayscale()
            dstImage.SetOriginalAlpha cReader.HasAlpha()
            dstImage.SetOriginalColorDepth cReader.EquivalentColorDepth
            dstImage.UpdateSize
            
            'DPI is not always reliable in PCX files, but we attempt to recover it anyway
            Dim srcXDPI As Single, srcYDPI As Single
            If cReader.GetDPI(srcXDPI, srcYDPI) Then dstImage.SetDPI srcXDPI, srcYDPI
            
            'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
            ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
            If (dstDIB Is Nothing) Then
                Set dstDIB = New pdDIB
                dstDIB.CreateBlank 16, 16, 32, 0, 0
                dstDIB.SetInitialAlphaPremultiplicationState True
            End If
            
        End If
    
    '/File is not PCX format
    End If
        
End Function

'Load a PDF file as a multi-layer image (typically one page per layer).  This function can also be used to load the first
' page from a PDF (for preview purposes) - to enable this, set "previewOnly" to TRUE.
'
'This function may need to raise a UI to ask the user for PDF import settings (like page resolution).  To skip this
' (during a batch process, for example), set the "noUI" parameter to TRUE.  When "previewOnly" is set to TRUE,
' the "noUI" parameter will automatically be enabled as well.
Public Function LoadPDF(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal previewOnly As Boolean = False, Optional ByVal noUI As Boolean = False, Optional ByRef userCanceledImportDialog As Boolean = False) As Boolean

    LoadPDF = False
    If (Not OS.IsVistaOrLater) Then Exit Function
    
    'Quickly validate a PDF-specific magic number before continuing
    If (Not Plugin_PDF.IsFileLikelyPDF(srcFile)) Then Exit Function
    
    'If the user requests "preview only" mode, set the "noUI" mode to match
    If previewOnly Then noUI = True
    
    'pdPDF handles all the dirty work for us, including prompting the user for a password as necessary.
    Dim cPDF As pdPDF
    Set cPDF = New pdPDF
    LoadPDF = cPDF.LoadPDFFromFile(srcFile)
    
    'In the future, more complex validation could be performed here, but for now,
    ' let's just double-confirm that the PDF object is happy with the loaded PDF.
    If LoadPDF Then LoadPDF = cPDF.HasPDF()
    
    'If we don't have a valid, loaded PDF, exit immediately.
    If (Not LoadPDF) Then Exit Function
    
    'Get the page count and validate it as > 0
    Dim numPages As Long
    numPages = cPDF.GetPageCount
    If (numPages <= 0) Then Exit Function
    If (Not cPDF.LoadPage(0)) Then Exit Function
    
    'If we're still here, a valid PDF exists and is ready for further processing.
    
    'If a UI is allowed, prompt the user for import settings
    Dim userAnswer As VbMsgBoxResult, userSettings As String
    If (Not noUI) And (Not previewOnly) Then
        
        userAnswer = Dialogs.PromptImportPDF(cPDF, userSettings)
        
        'The user can cancel the import dialog; abandon the entire load process if this happens
        If (userAnswer <> vbOK) Then
            userCanceledImportDialog = (userAnswer = vbCancel)
            LoadPDF = False
            Exit Function
        End If
        
    End If
    
    'Prep a parser for any user settings set via the import dialog
    Dim cSettings As pdSerialize
    Set cSettings = New pdSerialize
    cSettings.SetParamString userSettings
    
    'Retrieve user width/height and DPI settings
    Dim userWidthPx As Long, userHeightPx As Long, userDPI As Single
    userWidthPx = cSettings.GetLong("final-width-px", 0, True)
    userHeightPx = cSettings.GetLong("final-height-px", 0, True)
    userDPI = cSettings.GetSingle("final-dpi", 96!, True)
    
    'Failsafe against insane DPI values
    If (userDPI <= 1!) Then userDPI = 1!
    
    'If the retrieval of user settings failed, use a default DPI and retrieve the embedded PDF dimensions
    'As a failsafe, retrieve the dimensions of the first page IN POINTS.  (We can use this to generate
    ' a default image size if retrieving user settings failed for some reason.)
    Dim baseImageWidth As Single, baseImageHeight As Single
    baseImageWidth = cPDF.GetPageWidthInPoints()
    baseImageHeight = cPDF.GetPageHeightInPoints()
    
    'Use either the user settings, or the failsafe backup to calculate a page size IN PIXELS
    Dim defaultWidthInPixels As Long, defaultHeightInPixels As Long
    defaultWidthInPixels = Int(Units.ConvertOtherUnitToPixels(mu_Points, baseImageWidth, userDPI))
    defaultHeightInPixels = Int(Units.ConvertOtherUnitToPixels(mu_Points, baseImageHeight, userDPI))
    
    Dim baseWidthInPixels As Long, baseHeightInPixels As Long
    If (userWidthPx <= 0) Then
        baseWidthInPixels = defaultWidthInPixels
    Else
        baseWidthInPixels = userWidthPx
    End If
    If (userHeightPx <= 0) Then
        baseHeightInPixels = defaultHeightInPixels
    Else
        baseHeightInPixels = userHeightPx
    End If
    
    'We now want to determine a ratio between the page size the user selected, and the underlying
    ' page dimensions.  How we calculate this ratio doesn't really matter, it just needs to be consistent
    ' for *all* imported pages - because, for example, the user could choose to arbitrarily scale pages
    ' to half their width, but their original height - and we need to honor this for pages that are
    ' rotated or at a different size than the base page, too!
    Dim hRatio As Single, vRatio As Single
    hRatio = CDbl(baseWidthInPixels) / CDbl(defaultWidthInPixels)
    vRatio = CDbl(baseHeightInPixels) / CDbl(defaultHeightInPixels)
    
    'Initialize the target image with basic properties
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    dstImage.Width = baseWidthInPixels
    dstImage.Height = baseHeightInPixels
    dstImage.SetDPI userDPI, userDPI
    
    Dim i As Long
    
    'Figure out which pages the user actually wants loaded
    Dim listOfPages As pdStack
    Set listOfPages = New pdStack
    
    If previewOnly Then
        listOfPages.AddInt 0
    Else
        
        'Retrieve the list of pages from the incoming param string
        Dim strPagesToImport As String
        strPagesToImport = cSettings.GetString("import-pages", "all", True)
        
        'Note that the pages to be imported are loaded into a stack, and the stack is popped in
        ' *REVERSE* order - so we deliberately add pages to the stack in reverse order, to ensure
        ' correct order when popped.
        
        'Add all pages
        If Strings.StringsEqual(strPagesToImport, "all", True) Then
            For i = cPDF.GetPageCount() - 1 To 0 Step -1
                listOfPages.AddInt i
            Next i
        
        'Add only the first pages
        ElseIf Strings.StringsEqual(strPagesToImport, "first", True) Then
            listOfPages.AddInt 0
        
        'Custom page range
        ElseIf Strings.StringsEqual(strPagesToImport, "custom", True) Then
            
            'Retrieve the list of pages from a separate custom element, and ask the converter to change
            ' numbers from base-1 to base-0
            If TextSupport.ConvertPageRangeToStack(cSettings.GetString("page-list"), listOfPages, -1) Then
                
                'Reverse order so that the *lowest* page number comes *last*
                listOfPages.ReverseStack
                
            'List is bad; default to first page only
            Else
                listOfPages.ResetStack
                listOfPages.AddInt 0
            End If
            
        'Bad param string; default to first page only
        Else
            listOfPages.AddInt 0
        End If
        
        'Ensure at least one valid page was added to the import list
        If (listOfPages.GetNumOfInts <= 0) Then
            userCanceledImportDialog = vbCancel
            LoadPDF = False
            Exit Function
        End If
        
    End If
    
    'Before continuing, ensure the list of pages is sorted from most to least (because we pop pages off
    ' the stack in reverse order), but note that the user can override this in the import UI
    If cSettings.GetBool("reverse-pages", False, True) Then
        listOfPages.SortStackByValue True
    Else
        listOfPages.SortStackByValue False
    End If
    
    'Cache background color and/or transparency settings in advance (so we don't have to dip into the
    ' settings object in the inner loop)
    Dim bkColor As Long, bkOpaque As Boolean, bkOpacity As Long
    bkOpaque = cSettings.GetBool("background-solid", True, True)
    bkColor = cSettings.GetLong("background-color", RGB(255, 255, 255), True)
    
    'Other render settings are supplied to pdfium as flags
    Dim renderFlags As PDFium_RenderOptions: renderFlags = 0
    
    'Antialias for displays, printer, or none
    Select Case cSettings.GetLong("antialiasing", 0, True)
        Case 0
            renderFlags = renderFlags = FPDF_LCD_TEXT
        Case 1
            renderFlags = FPDF_PRINTING
        Case Else
            renderFlags = FPDF_RENDER_NO_SMOOTHIMAGE Or FPDF_RENDER_NO_SMOOTHPATH Or FPDF_RENDER_NO_SMOOTHTEXT
    End Select
    
    If cSettings.GetBool("annotations", False, True) Then renderFlags = renderFlags Or FPDF_ANNOT
    
    'If this is *not* a preview (or batch process), prep some UI bits
    Dim updateUI As Boolean
    updateUI = (Not previewOnly) And (Macros.GetMacroStatus() <> MacroBATCH) And (Macros.GetMacroStatus() <> MacroPLAYBACK)
    
    Dim numPagesTotal As Long
    numPagesTotal = listOfPages.GetNumOfInts()
    
    If updateUI Then
        ProgressBars.SetProgBarMax numPagesTotal
        ProgressBars.SetProgBarVal 0
    End If
    
    'Time to start iterating layers!
    Dim idxPage As Long, numPagesProcessed As Long
    numPagesProcessed = 0
    
    Do While listOfPages.PopInt(idxPage)
        
        numPagesProcessed = numPagesProcessed + 1
        
        If updateUI Then
            ProgressBars.SetProgBarVal numPagesProcessed
            Message "Loading page %1 of %2...", CStr(numPagesProcessed), numPagesTotal, "DONOTLOG"
        End If
        
        'Size can vary by page and bounding box.  Calculate a size for *this* page.
        Dim thisPageWidthPts As Single, thisPageHeightPts As Single
        If cPDF.LoadPage(idxPage) Then
            
            thisPageWidthPts = cPDF.GetPageWidthInPoints()
            thisPageHeightPts = cPDF.GetPageHeightInPoints()
            
            'Convert the default page dimensions (in points) to pixels
            Dim thisPageWidthPx As Long, thisPageHeightPx As Long
            thisPageWidthPx = Int(Units.ConvertOtherUnitToPixels(mu_Points, thisPageWidthPts, userDPI))
            thisPageHeightPx = Int(Units.ConvertOtherUnitToPixels(mu_Points, thisPageHeightPts, userDPI))
            
            'Multiply the final size (in pixels) by the ratio calculated from the user's import settings
            thisPageWidthPx = Int(thisPageWidthPx * hRatio + 0.5)
            thisPageHeightPx = Int(thisPageHeightPx * vRatio + 0.5)
            
            'Initialize a backing surface at the target size
            Dim tmpDIB As pdDIB
            If (tmpDIB Is Nothing) Then Set tmpDIB = New pdDIB
            
            'In preview mode, we want to render the page against a white background.  (Many PDFs only store black
            ' outlines against a transparent background, to improve printer behavior - and the previews look funny
            ' if rendered against a checkerboard!)
            If previewOnly Then
                bkColor = RGB(255, 255, 255)
                bkOpacity = 255
            Else
                If bkOpaque Then
                    'Color is set from the settings object, above
                    bkOpacity = 255
                Else
                    bkColor = RGB(0, 0, 0)
                    bkOpacity = 0
                End If
            End If
            tmpDIB.CreateBlank thisPageWidthPx, thisPageHeightPx, 32, bkColor, bkOpacity
            tmpDIB.SetInitialAlphaPremultiplicationState True
            
            'Render the page contents onto the target DIB
            ' (TODO: expose rendering options to the user)
            ' (TODO: calculate the way bounding boxes affect this)
            cPDF.RenderCurrentPageToPDDib tmpDIB, 0, 0, thisPageWidthPx, thisPageHeightPx, FPDF_Normal, renderFlags
            If (Not bkOpacity) Then tmpDIB.SetAlphaPremultiplication True, True
            
            'Prep a new layer object and initialize it
            Dim newLayerID As Long, tmpLayer As pdLayer
            newLayerID = dstImage.CreateBlankLayer()
            Set tmpLayer = dstImage.GetLayerByID(newLayerID)
            
            'We need a base layer name for each page.  For now, this is just "page".
            Dim baseLayerName As String
            baseLayerName = g_Language.TranslateMessage("Page %1", idxPage + 1)
            tmpLayer.InitializeNewLayer PDL_Image, baseLayerName, tmpDIB, True
            
            'Make the base layer visible, but no others.
            tmpLayer.SetLayerVisibility (numPagesProcessed = 1)
            tmpLayer.SetLayerBlendMode BM_Normal
            
        End If
        
    Loop
    
    'Set the base layer as the active one
    ' (TODO: may need to revisit this if page order is reversed?  Compare other software...)
    dstImage.SetActiveLayerByIndex 0
    
    'Perform some PD-specific object initialization before exiting
    If LoadPDF Then
        
        dstImage.SetOriginalFileFormat PDIF_PDF
        dstImage.NotifyImageChanged UNDO_Everything
        dstImage.SetOriginalColorDepth 24
        dstImage.SetOriginalGrayscale False
        dstImage.SetOriginalAlpha False
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank 16, 16, 32, 0
        dstDIB.SetColorManagementState cms_ProfileConverted
        
    End If
    
    'Unload any changes made to the primary app UI
    If updateUI Then ProgressBars.ReleaseProgressBar
    
End Function

Private Function LoadPNGOurselves(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long, Optional ByVal disableEmbeddedColorData As Boolean = False) As Boolean
    
    LoadPNGOurselves = False
    
    'pdPNG handles all the dirty work for us
    Set m_PNG = New pdPNG
    If disableEmbeddedColorData Then m_PNG.SetIgnoreColorData True
    LoadPNGOurselves = (m_PNG.LoadPNG_Simple(srcFile, dstImage, dstDIB, False) <= png_Warning)
    
    If LoadPNGOurselves Then
    
        'If we've experienced one or more warnings during the load process, dump them out to the debug file.
        If (m_PNG.Warnings_GetCount() > 0) Then m_PNG.Warnings_DumpToDebugger
        
        'Relay any useful state information to the destination image object; this information may be useful
        ' if/when the user saves the image.
        dstImage.SetOriginalFileFormat PDIF_PNG
        dstImage.NotifyImageChanged UNDO_Everything
        
        dstImage.SetOriginalColorDepth m_PNG.GetBitsPerPixel()
        dstImage.SetOriginalGrayscale (m_PNG.GetColorType = png_Greyscale) Or (m_PNG.GetColorType = png_GreyscaleAlpha)
        dstImage.SetOriginalAlpha m_PNG.HasAlpha()
        
        'Now, some PNG-specific info that can be helpful if the user wants to "preserve original settings" later
        dstImage.ImgStorage.AddEntry "png-color-type", m_PNG.GetColorType()
        If m_PNG.HasChunk("bKGD") Then dstImage.ImgStorage.AddEntry "png-background-color", m_PNG.GetBackgroundColor()
        If m_PNG.HasChunk("tRNS") Then
            Dim trnsColor As Long
            If m_PNG.GetTransparentColor(trnsColor) Then dstImage.ImgStorage.AddEntry "png-transparent-color", trnsColor
        End If
        
        'Because color-management has already been handled (if applicable), this is a great time to premultiply alpha
        If (Not dstDIB.GetAlphaPremultiplication()) Then dstDIB.SetAlphaPremultiplication True
        
        'If this is not an animated PNG, free all associated memory now.
        ' (Animated PNGs will be freed at a later stage.)
        If (Not m_PNG.IsAnimated()) Then
            Set m_PNG = Nothing
        Else
            numOfPages = m_PNG.NumAnimationFrames()
            imageHasMultiplePages = (numOfPages > 1)
            If (Not imageHasMultiplePages) Then Set m_PNG = Nothing
        End If
        
    End If

End Function

Public Function LoadRemainingPNGFrames(ByRef dstImage As pdImage) As Boolean
    LoadRemainingPNGFrames = (m_PNG.ImportStage7_LoadRemainingFrames(dstImage) < png_Failure)
    Set m_PNG = Nothing
    dstImage.NotifyImageChanged UNDO_Image
End Function

'Use PD's internal PSD parser to attempt to load a target PSD file.
Private Function LoadPSD(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadPSD = False
    
    'pdPSD handles all the dirty work for us
    Dim cPSD As pdPSD
    Set cPSD = New pdPSD
    
    'Validate the potential psd file
    LoadPSD = cPSD.IsFilePSD(srcFile, True)
    
    'If validation passes, attempt a full load
    If LoadPSD Then LoadPSD = (cPSD.LoadPSD(srcFile, dstImage, dstDIB) < psd_Failure)
    
    'Perform some PD-specific object initialization before exiting
    If LoadPSD And (Not dstImage Is Nothing) Then
        
        dstImage.SetOriginalFileFormat PDIF_PSD
        dstImage.NotifyImageChanged UNDO_Everything
        dstImage.SetOriginalColorDepth cPSD.GetBitsPerPixel()
        dstImage.SetOriginalGrayscale cPSD.IsGrayscaleColorMode()
        dstImage.SetOriginalAlpha cPSD.HasAlpha()
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank 16, 16, 32, 0
        dstDIB.SetColorManagementState cms_ProfileConverted
        
        'Before exiting, ensure all color management data has been added to PD's central cache
        Dim profHash As String
        If cPSD.HasICCProfile() Then
            profHash = ColorManagement.AddProfileToCache(cPSD.GetICCProfile(), True, False, False)
            dstImage.SetColorProfile_Original profHash
            
            'IMPORTANT NOTE: at present, the destination image - by the time we're done with it - will have been
            ' hard-converted to sRGB, so we don't want to associate the destination DIB with its source profile.
            ' Instead, note that it is currently sRGB.
            profHash = ColorManagement.GetSRGBProfileHash()
            dstDIB.SetColorProfileHash profHash
            dstDIB.SetColorManagementState cms_ProfileConverted
            
        End If
        
    End If
    
End Function

'Use PD's internal PSP parser to attempt to load a target PaintShop Pro file.
Private Function LoadPSP(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadPSP = False
    
    'pdPSP handles all the dirty work for us
    Dim cPSP As pdPSP
    Set cPSP = New pdPSP
    
    'Validate (and if valid, load) the potential PSP file in one fell swoop
    LoadPSP = (cPSP.LoadPSP(srcFile, dstImage, dstDIB) < psp_Failure)
    
    'Perform some PD-specific object initialization before exiting
    If LoadPSP And (Not dstImage Is Nothing) Then
        
        dstImage.SetOriginalFileFormat PDIF_PSP
        dstImage.NotifyImageChanged UNDO_Everything
        dstImage.SetOriginalColorDepth cPSP.GetOriginalColorDepth()
        dstImage.SetOriginalGrayscale cPSP.IsPSPGrayscale()
        dstImage.SetOriginalAlpha cPSP.HasAlpha()
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank 16, 16, 32, 0
        dstDIB.SetColorManagementState cms_ProfileConverted

        'Before exiting, ensure all color management data has been added to PD's central cache
        Dim profHash As String
        If cPSP.HasICCProfile() Then
            profHash = ColorManagement.AddProfileToCache(cPSP.GetICCProfile(), True, False, False)
            dstImage.SetColorProfile_Original profHash

            'IMPORTANT NOTE: at present, the destination image - by the time we're done with it -
            ' will have been hard-converted to sRGB, so we don't want to associate the destination
            ' DIB with its source profile. Instead, note that it is currently sRGB to prevent the
            ' central color-manager from attempting to correct it on its own.
            profHash = ColorManagement.GetSRGBProfileHash()
            dstDIB.SetColorProfileHash profHash
            dstDIB.SetColorManagementState cms_ProfileConverted

        End If

    End If
    
End Function

Private Function LoadQOI(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadQOI = False
    
    'pdQOI handles all the dirty work for us
    Dim cReader As pdQOI
    Set cReader = New pdQOI
    
    'Validate and (potentially) load the file in one fell swoop
    LoadQOI = cReader.LoadQOI_FromFile(srcFile, dstImage, dstDIB)
    
    'Perform some PD-specific object initialization before exiting
    If LoadQOI Then
        
        'Set format flags and reset internal image caches
        dstImage.SetOriginalFileFormat PDIF_QOI
        dstImage.NotifyImageChanged UNDO_Everything
        
        'QOI files are always 24- or 32-bpp
        If (cReader.GetOriginalChannelCount = 4) Then
            dstImage.SetOriginalColorDepth 32
        Else
            dstImage.SetOriginalColorDepth 24
        End If
        
        dstImage.SetOriginalGrayscale False
        
        'Use channel count to determine original alpha state
        dstImage.SetOriginalAlpha (cReader.GetOriginalChannelCount = 4)
        
        'QOI files don't really support color management, but can be flagged as sRGB vs linear;
        ' handling this is TODO pending relevant test files
        dstDIB.SetColorManagementState cms_ProfileConverted
        
    End If
    
End Function

Public Function LoadHEIF(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal previewOnly As Boolean = False) As Boolean
        
    LoadHEIF = False
    
    'Ensure support libraries exist
    If (Not Plugin_Heif.IsLibheifEnabled()) Then Exit Function
    
    'Validate file
    If Plugin_Heif.IsFileHeif(srcFile) Then
        
        'Attempt load via external library
        LoadHEIF = Plugin_Heif.LoadHeifImage(srcFile, dstImage, dstDIB, True, previewOnly)
        
        'If the load was successful, populate some default properties
        If LoadHEIF And (Not dstImage Is Nothing) Then
            dstImage.SetOriginalFileFormat PDIF_HEIF
            dstImage.NotifyImageChanged UNDO_Everything
        End If
        
    End If

End Function

'Load a WAP Bitmap (WBMP).  Originally this was handled by FreeImage, but as part of a broader movement way from
' relying on FreeImage (the library is effectively abandoned) I wrote my own small WBMP decoder in 2024.
Private Function LoadWBMP(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadWBMP = False
    
    'pdWBMP handles all the dirty work for us
    Dim cReader As pdWBMP
    Set cReader = New pdWBMP
    
    'Validate the file
    If cReader.IsFileWBMP(srcFile) Then
    
        'Load the file
        LoadWBMP = cReader.LoadWBMP_FromFile(srcFile, dstImage, dstDIB)
        
        'Perform some PD-specific object initialization before exiting
        If LoadWBMP Then
            
            dstImage.SetOriginalFileFormat PDIF_WBMP
            dstImage.NotifyImageChanged UNDO_Everything
            dstImage.SetOriginalGrayscale True
            dstImage.SetOriginalAlpha False
            
            'WBMP images are always 1-bit
            dstImage.SetOriginalColorDepth 1
            
            'As simple 1-bit icons, WBMP files obviously don't support color management!
            dstDIB.SetColorManagementState cms_ProfileConverted
            
        End If
    
    '/File is not WBMP format
    End If
        
End Function

'Use libwebp to parse a WebP file
Private Function LoadWebP(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadWebP = False
    
    'Perform a failsafe format check before continuing.  (In PD, this check is always performed
    ' by the calling function, but in case of code refactoring in the future, we maintain a check
    ' here too.)
    If (Not Plugin_WebP.IsWebP(srcFile)) Then Exit Function
    
    'libwebp provides no direct file-reading possibilities, so we need to pre-load the file into memory.
    ' This is generally not a problem as WebP files tend to be (per the name?) web-oriented sizes.
    ' Also - note that libwebp *does* provide an incremental parser for reading pixel data, but it's not
    ' helpful for determining things like file properties (e.g. isAnimated), so if we have to pre-load
    ' the file for one thing, we may as well pre-load it for everything.
    Dim fullFile() As Byte
    If Files.FileLoadAsByteArray(srcFile, fullFile) Then
        
        'libwebp handles all subsequent parsing duties for us.
        PDDebug.LogAction "WebP file found.  Handing parsing duties over to pdWebP..."
        
        Dim cWebP As pdWebP
        Set cWebP = New pdWebP
        LoadWebP = cWebP.LoadWebP_FromMemory(srcFile, VarPtr(fullFile(0)), UBound(fullFile) + 1, dstImage, dstDIB)
        
        'Perform some PD-specific object initialization before exiting
        If LoadWebP And (Not dstImage Is Nothing) Then
            
            dstImage.SetOriginalFileFormat PDIF_WEBP
            dstImage.NotifyImageChanged UNDO_Everything
            
            dstImage.SetOriginalColorDepth 32
            dstImage.SetOriginalGrayscale False
            dstImage.SetOriginalAlpha cWebP.HasAlpha()
            dstImage.SetAnimated cWebP.IsAnimated()
            
            'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
            ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
            Set dstDIB = New pdDIB
            dstDIB.CreateBlank 16, 16, 32, 0
            dstDIB.SetColorManagementState cms_ProfileConverted
    
    'TODO: testing of WebP files and color management; this has changed since migrating to direct libwebp interfacing
    '        'Before exiting, ensure all color management data has been added to PD's central cache
    '        Dim profHash As String
    '        If cPSP.HasICCProfile() Then
    '            profHash = ColorManagement.AddProfileToCache(cPSP.GetICCProfile(), True, False, False)
    '            dstImage.SetColorProfile_Original profHash
    '
    '            'IMPORTANT NOTE: at present, the destination image - by the time we're done with it -
    '            ' will have been hard-converted to sRGB, so we don't want to associate the destination
    '            ' DIB with its source profile. Instead, note that it is currently sRGB to prevent the
    '            ' central color-manager from attempting to correct it on its own.
    '            profHash = ColorManagement.GetSRGBProfileHash()
    '            dstDIB.SetColorProfileHash profHash
    '            dstDIB.SetColorManagementState cms_ProfileConverted
    '
    '        End If
        
        End If

    Else
        PDDebug.LogAction "WARNING! LoadWebP() couldn't load source file."
    End If
    
End Function

'Load an X Bitmap (XBM) image.  Originally this was handled by FreeImage, but FreeImage crashes on every XBM load attempt.
' So I wrote my own loader in March 2024.
Private Function LoadXBM(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadXBM = False
    
    'pdXBM handles all the dirty work for us
    Dim cReader As pdXBM
    Set cReader = New pdXBM
    
    'Validate the file
    If cReader.IsFileXBM(srcFile) Then
    
        'Load the file
        LoadXBM = cReader.LoadXBM_FromFile(srcFile, dstImage, dstDIB)
        
        'Perform some PD-specific object initialization before exiting
        If LoadXBM Then
            
            dstImage.SetOriginalFileFormat PDIF_XBM
            dstImage.NotifyImageChanged UNDO_Everything
            dstImage.SetOriginalGrayscale True
            dstImage.SetOriginalAlpha False
            
            'XBM images are always 1-bit
            dstImage.SetOriginalColorDepth 1
            
            'As simple 1-bit icons, XBM files obviously don't support color management!
            dstDIB.SetColorManagementState cms_ProfileConverted
            
        End If
    
    '/File is not XBM format
    End If
        
End Function

Private Function LoadXCF(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadXCF = False
    
    'pdXCF handles all the dirty work for us
    Dim cReader As pdXCF
    Set cReader = New pdXCF
    
    'Validate and (potentially) load the file in one fell swoop
    LoadXCF = cReader.LoadXCF_FromFile(srcFile, dstImage, dstDIB)
    
    'Perform some PD-specific object initialization before exiting
    If LoadXCF Then
        
        'Set format flags and reset internal image caches
        dstImage.SetOriginalFileFormat PDIF_XCF
        dstImage.SetDPI cReader.GetOriginalDPI, cReader.GetOriginalDPI
        dstImage.NotifyImageChanged UNDO_Everything
        
        'Mark any other image-level properties
        dstImage.SetOriginalColorDepth cReader.GetOriginalColorDepth()
        dstImage.SetOriginalGrayscale cReader.HasGrayscale()
        dstImage.SetOriginalAlpha cReader.GetOriginalAlphaState()
        
        'Before exiting, ensure all color management data has been added to PD's central cache
        Dim profHash As String
        If (Not cReader.GetICCProfile() Is Nothing) Then
            profHash = ColorManagement.AddProfileToCache(cReader.GetICCProfile(), True, False, False)
            dstImage.SetColorProfile_Original profHash
            
            'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
            ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
            If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
            dstDIB.CreateBlank 16, 16, 32, 0
            dstDIB.SetColorManagementState cms_ProfileConverted
            
            'IMPORTANT NOTE: at present, the destination image - by the time we're done with it - will have been
            ' hard-converted to sRGB, so we don't want to associate the destination DIB with its source profile.
            ' Instead, note that it is currently sRGB.
            profHash = ColorManagement.GetSRGBProfileHash()
            dstDIB.SetColorProfileHash profHash
            dstDIB.SetColorManagementState cms_ProfileConverted
            
        End If
        
    End If
    
End Function

'Most portions of PD operate exclusively in 32-bpp mode.  (This greatly simplifies the compositing pipeline.)
'Returns: TRUE if changes were made to the target DIB
Public Function ForceTo32bppMode(ByRef targetDIB As pdDIB) As Boolean
    
    ForceTo32bppMode = False
    
    If (targetDIB.GetDIBColorDepth <> 32) Then
        If Drawing2D.IsRenderingEngineActive(P2_GDIPlusBackend) Then GDI_Plus.GDIPlusConvertDIB24to32 targetDIB Else targetDIB.ConvertTo32bpp
        ForceTo32bppMode = True
    End If

End Function

'Autosave images are tricky to restore, as we have to try and reconstruct the image using whatever data we find in the Temp folder.
' IMPORTANT NOTE: the passed srcFile string *may be modified by this function, by design.*  Plan accordingly.
Public Function SyncRecoveredAutosaveImage(ByRef srcFile As String, ByRef srcImage As pdImage) As Boolean
    
    PDDebug.LogAction "SyncRecoveredAutosaveImage invoked; attempting to recover usable data from the Autosave database..."
    srcImage.ImgStorage.AddEntry "CurrentLocationOnDisk", srcFile
            
    'Ask the AutoSave engine to synchronize this image's data against whatever it can recover from the Autosave database
    Autosaves.AlignLoadedImageWithAutosave srcImage
            
    'This is a bit wacky, but - the Autosave engine will automatically update the "locationOnDisk" attribute based on
    ' information inside the Autosave recovery database.  We thus want to overwrite the original srcFile value (which points
    ' at a temp file copy of whatever we're attempting to recover), with the new, recovered srcFile value.
    srcFile = srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)
    
    SyncRecoveredAutosaveImage = True
    
End Function

'After loading an image file, you can call this function to set up any post-load pdImage attributes (like name and save state)
Public Function GenerateExtraPDImageAttributes(ByRef srcFile As String, ByRef targetImage As pdImage, ByRef suggestedFilename As String) As Boolean
    
    'If PD explicitly requested a custom image name, we can safely assume the calling routine is NOT loading a generic image file
    ' from disk - instead, this image came from a scanner, or screen capture, or some format that doesn't automatically yield a
    ' usable filename.
    
    'Therefore, our job is to coordinate between the image's suggested name (which will be suggested at first-save), the actual
    ' location on disk (which we treat as "non-existent", even though we're loading from a temp file of some sort), and the image's
    ' save state (which we forcibly set to FALSE to ensure the user is prompted to save before closing the image).
    If (LenB(suggestedFilename) = 0) Then
    
        'The calling routine didn't specify a custom image name, so we can assume this is a normal image file.
        'Prep all default attributes using the filename itself.
        targetImage.ImgStorage.AddEntry "CurrentLocationOnDisk", srcFile
        targetImage.ImgStorage.AddEntry "OriginalFileName", Files.FileGetName(srcFile, True)
        targetImage.ImgStorage.AddEntry "OriginalFileExtension", Files.FileGetExtension(srcFile)
        
        'Note the image's save state; PDI files are specially marked as having been "saved losslessly".
        If (targetImage.GetCurrentFileFormat = PDIF_PDI) Then
            targetImage.SetSaveState True, pdSE_SavePDI
        Else
            targetImage.SetSaveState True, pdSE_SaveFlat
        End If
        
    Else
    
        'The calling routine has specified a file name.  Assume this is a special case, and force a Save As...
        ' dialog in the future by not specifying a location on disk
        targetImage.ImgStorage.AddEntry "CurrentLocationOnDisk", vbNullString
        targetImage.ImgStorage.AddEntry "OriginalFileName", suggestedFilename
        targetImage.ImgStorage.AddEntry "OriginalFileExtension", vbNullString
        
        'For this special case, mark the image as being totally unsaved; this forces us to eventually show a save prompt
        targetImage.SetSaveState False, pdSE_AnySave
        
    End If
    
End Function

'After successfully loading an image, you can call this helper function to automatically apply needed UI changes.
Public Sub ApplyPostLoadUIChanges(ByRef srcFile As String, ByRef srcImage As pdImage, Optional ByVal addToRecentFiles As Boolean = True)

    'Just to be safe, update the color management profile of the current monitor
    ' TODO: find a better place to handle this
    CheckParentMonitor True
    
    'Reset the main viewport's scroll bars to (0, 0)
    Viewport.DisableRendering
    FormMain.MainCanvas(0).SetScrollVisibility pdo_Both, True
    FormMain.MainCanvas(0).SetScrollValue pdo_Both, 0
    
    'Reflow any image-window-specific display elements on the actual image form (status bar, rulers, etc).
    ' Importantly, this also shows/hides the image tabstrip that's available when multiple images are loaded.
    FormMain.MainCanvas(0).AlignCanvasView
    
    'Fit the current image to the active viewport
    FitImageToViewport
    Viewport.EnableRendering
    
    'Notify the UI manager that it now has one more image to deal with
    If (Macros.GetMacroStatus <> MacroBATCH) Then Interface.NotifyImageAdded srcImage.imageID
    
    'Add this file to the MRU list (unless specifically told not to)
    If addToRecentFiles And (Macros.GetMacroStatus <> MacroBATCH) Then g_RecentFiles.AddFileToList srcFile, srcImage
    
End Sub
