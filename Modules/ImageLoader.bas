Attribute VB_Name = "ImageImporter"
'***************************************************************************
'Low-level image import interfaces
'Copyright 2001-2019 by Tanner Helland
'Created: 4/15/01
'Last updated: 20/January/19
'Last update: integrate our now homebrew PSD decoder
'
'This module provides low-level "import" functionality for importing image files into PD.  You will not generally want
' to interface with this module directly; instead, rely on the high-level functions in the "Loading" module.
' They will intelligently drop into this module as necessary, sparing you the messy work of having to handle
' format-specific details (which are many).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'As of v7.2, PD includes its own custom-built PNG parser.  This offers a number of performance and
' feature enhancements relative to the 3rd-party libraries.  I know of no reason why it would need
' to be disabled, but if you want to fall back to the old FreeImage and GDI+ interface, you can set
' this to FALSE.
Private Const USE_INTERNAL_PARSER_PNG As Boolean = True

'OpenRaster support was added in the 7.2 release cycle.  I know of no reason to disable it at present,
' but you can use this constant to deactivate parsing support if necessary.
Private Const USE_INTERNAL_PARSER_ORA As Boolean = True

'PD's internal PSD parser is still under heavy construction.  If it fails, PD will fall back to
' FreeImage's rudimentary PSD support (e.g. no layer, just a composite image).  If you try to
' load a PSD and it doesn't load correctly, PLEASE FILE AN ISSUE ON GITHUB.  I don't have a modern
' copy of Photoshop for testing, so outside help is necessary for fixing esoteric PSD bugs!
Private Const USE_INTERNAL_PARSER_PSD As Boolean = True

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

'PDI loading.  "PhotoDemon Image" files are the only format PD supports for saving layered images.  PDI to PhotoDemon is like
' PSD to PhotoShop, or XCF to Gimp.
'
'Note the unique "sourceIsUndoFile" parameter for this load function.  PDI files are used to store undo/redo data, and when one of their
' kind is loaded as part of an Undo/Redo action, we must ignore certain elements stored in the file (e.g. settings like "LastSaveFormat"
' which we do not want to Undo/Redo).  This parameter is passed to the pdImage initializer, and it tells it to ignore certain settings.
Public Function LoadPhotoDemonImage(ByVal pdiPath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByVal sourceIsUndoFile As Boolean = False) As Boolean
    
    PDDebug.LogAction "PDI file identified.  Starting pdPackage decompression..."
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    On Error GoTo LoadPDIFail
    
    'PDI files require a parent pdImage container
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    
    'Load the file into the pdPackager instance.  Note that this step will also validate the incoming file.
    ' (Also, prior to v7.0, PD would copy the entire source file into memory, then load the PDI from there.  This no longer occurs;
    '  instead, the file is left on-disk, and data is only loaded on a per-node basis.  This greatly reduces memory load.)
    ' (Also, because PDI files store data roughly sequentially, we can use OptimizeSequentialAccess for a small perf boost.)
    If pdiReader.ReadPackageFromFile(pdiPath, PD_IMAGE_IDENTIFIER, PD_SM_MemoryBacked, PD_SA_ReadOnly, OptimizeSequentialAccess) Then
    
        PDDebug.LogAction "pdPackage successfully read and initialized.  Starting package parsing..."
        
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String, retSize As Long
        
        If pdiReader.GetNodeDataByIndex(0, True, retBytes, False, retSize) Then
            
            PDDebug.LogAction "Initial PDI node retrieved.  Initializing corresponding pdImage object..."
            
            'Copy the received bytes into a string
            retString = Space$(retSize \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.ReadExternalData retString, True, sourceIsUndoFile
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        PDDebug.LogAction "pdImage created successfully.  Moving on to individual layers..."
        
        'With the main pdImage now assembled, the next task is to populate all layers with two pieces of information:
        ' 1) The layer header, which contains stuff like layer name, opacity, blend mode, etc
        ' 2) Layer-specific information, which varies by layer type.  For DIBs, this will be a raw stream of bytes
        '    containing the layer DIB's raster data.  For text or other vector layers, this is an XML stream containing
        '    whatever information is necessary to construct the layer from scratch.
        
        Dim i As Long
        For i = 0 To dstImage.GetNumOfLayers - 1
        
            PDDebug.LogAction "Retrieving layer header " & i & "..."
            
            'First, retrieve the layer's header
            If pdiReader.GetNodeDataByIndex(i + 1, True, retBytes, False, retSize) Then
            
                'Copy the received bytes into a string
                retString = Space$(retSize \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
                
                'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                If Not dstImage.GetLayerByIndex(i).CreateNewLayerFromXML(retString) Then
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
                End If
                
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
            'How we extract the rest of the layer's data varies by layer type.  Raster layers can skip the need for a temporary buffer,
            ' because we've already created a DIB with a built-in buffer for the pixel data.
            '
            'Other layer types (e.g. vector layers) are tiny so a temporary buffer does not matter; also, unlike raster buffers, we cannot
            ' easily predict the buffer size in advance, so we rely on pdPackage to do it for us
            Dim nodeLoadedSuccessfully As Boolean
            nodeLoadedSuccessfully = False
            
            'Image (raster) layers
            If dstImage.GetLayerByIndex(i).IsLayerRaster Then
                
                PDDebug.LogAction "Raster layer identified.  Retrieving pixel bits..."
                
                'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                Dim tmpDIBPointer As Long, tmpDIBLength As Long
                dstImage.GetLayerByIndex(i).layerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                'At present, all pdPackage layers will contain premultiplied alpha, so force the corresponding state now
                dstImage.GetLayerByIndex(i).layerDIB.SetInitialAlphaPremultiplicationState True
                
                nodeLoadedSuccessfully = pdiReader.GetNodeDataByIndex_UnsafeDstPointer(i + 1, False, tmpDIBPointer)
            
            'Text and other vector layers
            ElseIf dstImage.GetLayerByIndex(i).IsLayerVector Then
                
                PDDebug.LogAction "Vector layer identified.  Retrieving layer XML..."
                
                If pdiReader.GetNodeDataByIndex(i + 1, False, retBytes, False, retSize) Then
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after pdPackager gained full Unicode compatibility.
                    retString = Space$(retSize \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
                    
                    'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                    If dstImage.GetLayerByIndex(i).CreateVectorDataFromXML(retString) Then
                        nodeLoadedSuccessfully = True
                    Else
                        Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                    End If
                    
                Else
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                End If
                    
            'In the future, additional layer types can be handled here
            Else
                Debug.Print "WARNING! Unknown layer type exists in this PDI file: " & dstImage.GetLayerByIndex(i).GetLayerType
            End If
            
            'If successful, notify the parent of the change
            If nodeLoadedSuccessfully Then
                dstImage.NotifyImageChanged UNDO_Layer, i
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
        
        Next i
        
        PDDebug.LogAction "All layers loaded.  Looking for remaining non-essential PDI data..."
        
        Dim nonEssentialParseTime As Currency
        VBHacks.GetHighResTime nonEssentialParseTime
        
        'Finally, check to see if the PDI image has a metadata entry.  If it does, load that data now.
        If pdiReader.GetNodeDataByName("pdMetadata_Raw", True, retBytes, False, retSize) Then
        
            PDDebug.LogAction "Raw metadata chunk found.  Retrieving now..."
            
            'Copy the received bytes into a string
            retString = Space$(retSize \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
            
            'Pass the string to the parent image's metadata handler, which will parse the XML data and prepare a matching
            ' internal metadata struct.
            If Not dstImage.ImgMetadata.LoadAllMetadata(retString, dstImage.imageID, sourceIsUndoFile) Then
                
                'For invalid metadata, do not reject the rest of the PDI file.  Instead, just warn the user and carry on.
                PDDebug.LogAction "WARNING: PDI Metadata Node rejected by metadata parser."
                
            End If
        
        End If
        
        '(As of v7.0, a serialized copy of the image's metadata is also stored.  This copy contains all user edits
        ' and other changes.)
        If pdiReader.GetNodeDataByName("pdMetadata_Raw", False, retBytes, False, retSize) Then
        
            PDDebug.LogAction "Serialized metadata chunk found.  Retrieving now..."
            
            'Copy the received bytes into a string
            retString = Space$(retSize \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
            
            'Pass the string to the parent image's metadata handler, which will parse the XML data and prepare a matching
            ' internal metadata struct.
            dstImage.ImgMetadata.RecreateFromSerializedXMLData retString
        
        End If
        
        PDDebug.LogAction "PDI parsing complete.  Returning control to main image loader..."
        PDDebug.LogAction "(Time required to load PDI file: " & VBHacks.GetTimeDiffNowAsString(startTime) & ", non-essential components took " & VBHacks.GetTimeDiffNowAsString(nonEssentialParseTime) & ")"
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank 16, 16, 32, 0
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonImage = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed.  This may be a legacy PDI file -- try that function next.
        PDDebug.LogAction "Legacy PDI file encountered; dropping back to pdPackage v1 functions..."
        LoadPhotoDemonImage = LoadPDI_Legacy(pdiPath, dstDIB, dstImage, sourceIsUndoFile)
    
    End If
    
    Exit Function
    
LoadPDIFail:
    
    PDDebug.LogAction "WARNING!  LoadPDIFail error routine reached.  Checking for known error states..."
    
    'Before falling back to a generic error message, check for a couple known problem states.
    
    'Case 1: this file is compressed (using one or more libraries), and the user has somehow messed up their PD plugin situation
    Dim cmpMissing As Boolean
    cmpMissing = pdiReader.GetPackageFlag(PDP_HF2_ZlibRequired, PDP_LOCATION_ANY) And (Not PluginManager.IsPluginCurrentlyEnabled(CCP_zLib))
    cmpMissing = cmpMissing Or pdiReader.GetPackageFlag(PDP_HF2_ZstdRequired, PDP_LOCATION_ANY) And (Not PluginManager.IsPluginCurrentlyEnabled(CCP_zstd))
    cmpMissing = cmpMissing Or pdiReader.GetPackageFlag(PDP_HF2_Lz4Required, PDP_LOCATION_ANY) And (Not PluginManager.IsPluginCurrentlyEnabled(CCP_lz4))
    
    If cmpMissing Then
        PDMsgBox "The PDI file ""%1"" contains compressed data, but the required plugin is missing or disabled.", vbCritical Or vbOKOnly, "Compression plugin missing", Files.FileGetName(pdiPath)
        Exit Function
    End If

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else
            Message "An error has occurred (#%1 - %2).  PDI load abandoned.", Err.Number, Err.Description
        
    End Select
    
    LoadPhotoDemonImage = False
    Exit Function

End Function

'Load just the layer stack from a standard PDI file, and non-destructively align our current layer stack to match.
' At present, this function is only used internally by the Undo/Redo engine.
Public Function LoadPhotoDemonImageHeaderOnly(ByVal pdiPath As String, ByRef dstImage As pdImage) As Boolean
    
    On Error GoTo LoadPDIHeaderFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.ReadPackageFromFile(pdiPath, PD_IMAGE_IDENTIFIER) Then
    
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String, retSize As Long
        If pdiReader.GetNodeDataByIndex(0, True, retBytes, False, retSize) Then
        
            'Copy the received bytes into a string
            retString = Space$(retSize \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.ReadExternalData retString, True, True, True
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'With the main pdImage now assembled, the next task is to populate all layer headers.  This is a bit more
        ' confusing than a regular PDI load, because we have to maintain existing layer DIB data (ugh!).
        ' So basically, we must:
        ' 1) Extract each layer header from file, in turn
        ' 2) See if the current pdImage copy of this layer is in the proper position in the layer stack; if it isn't,
        '    move it into the location specified by the PDI file.
        ' 3) Ask the layer to non-destructively overwrite its header with the header from the PDI file (e.g. don't
        '    touch its DIB or vector-specific contents).
        '
        'Note also that header data may include image metadata; this is handled separately.
        Dim layerNodeName As String, layerNodeID As Long, layerNodeType As Long
        
        Dim i As Long
        For i = 0 To dstImage.GetNumOfLayers - 1
        
            'Before doing anything else, retrieve the ID of the node at this position.  (Retrieve the rest of the node
            ' header too, although we don't actually have a use for those values at present.)
            pdiReader.GetNodeInfo i + 1, layerNodeName, layerNodeID, layerNodeType
            
            'We now know what layer ID is supposed to appear at this position in the layer stack.  If that layer ID
            ' is *not* in its proper position, move it now.
            If (dstImage.GetLayerIndexFromID(layerNodeID) <> i) Then dstImage.SwapTwoLayers dstImage.GetLayerIndexFromID(layerNodeID), i
            
            'Now that the node is in place, we can retrieve its header.
            If pdiReader.GetNodeDataByIndex(i + 1, True, retBytes, False, retSize) Then
            
                'Copy the received bytes into a string
                retString = Space$(retSize \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
                
                'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                If Not dstImage.GetLayerByIndex(i).CreateNewLayerFromXML(retString, , True) Then
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
                End If
                
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
            'Normally we would load the layer's DIB data here, but we don't care about that when loading just the headers!
            ' Continue to the next layer.
        
        Next i
        
        'Finally, check to see if the PDI image has a metadata entry.  If it does, load that data now.
        If pdiReader.GetNodeDataByName("pdMetadata_Raw", True, retBytes, False, retSize) Then
            
            'Copy the received bytes into a string, then pass that string to the parent image's metadata handler
            retString = Space$(retSize \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
            If dstImage.ImgMetadata.LoadAllMetadata(retString, dstImage.imageID, True) Then
                
                '(As of v7.0, a serialized copy of the image's metadata is also stored.  This copy contains all user edits
                ' and other changes.)
                If pdiReader.GetNodeDataByName("pdMetadata_Raw", False, retBytes, False, retSize) Then
                    retString = Space$(retSize \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
                    dstImage.ImgMetadata.RecreateFromSerializedXMLData retString
                End If
                
            Else
                PDDebug.LogAction "WARNING!  ImageImporter.LoadPhotoDemonImageHeaderOnly() failed to retrieve to parse this image's metadata chunk."
            End If
        
        Else
            Debug.Print "FYI, this PDI file does not contain metadata information."
            
            'As a failsafe (because the target image may already exist, and this operation can be triggered by
            ' something like "Redo: Remove all metadata", meaning the target image already has a full metadata
            ' manager), erase the target image's metadata collection, if any.
            dstImage.ImgMetadata.Reset
            
        End If
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonImageHeaderOnly = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file may be a legacy PDI format.
        PDDebug.LogAction "PDI v2 validation failed.  Attempting v1 load engine..."
        LoadPhotoDemonImageHeaderOnly = LoadPhotoDemonImageHeaderOnly_Legacy(pdiPath, dstImage)
    
    End If
    
    Exit Function
    
LoadPDIHeaderFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadPhotoDemonImageHeaderOnly = False
    Exit Function

End Function

'Load a single layer from a standard PDI file.
' At present, this function is only used internally by the Undo/Redo engine.  If the nearest diff to a layer-specific change is a
' full pdImage stack, this function is used to extract only the relevant layer (or layer header) from the PDI file.
Public Function LoadSingleLayerFromPDI(ByVal pdiPath As String, ByRef dstLayer As pdLayer, ByVal targetLayerID As Long, Optional ByVal loadHeaderOnly As Boolean = False) As Boolean
    
    On Error GoTo LoadLayerFromPDIFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.ReadPackageFromFile(pdiPath, PD_IMAGE_IDENTIFIER) Then
    
        'PDI files all follow a standard format: a pdImage node at the top, which contains the full pdImage header,
        ' followed by individual nodes for each layer.  Layers are stored in stack order, which makes it very fast and easy
        ' to reconstruct the layer stack.
        
        'Unfortunately, stack order is not helpful in this function, because the target layer's position may have changed
        ' since the time this pdImage file was created.  To work around that, we must located the layer using its cardinal
        ' ID value, which is helpfully stored as the node ID parameter for a given layer node.
        
        Dim retBytes() As Byte, retString As String, retSize As Long
        
        If pdiReader.GetNodeDataByID(targetLayerID, True, retBytes, False, retSize) Then
        
            'Copy the received bytes into a string
            retString = Space$(retSize \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
            
            'Pass the string to the target layer, which will read the XML data and initialize itself accordingly.
            ' Note that we also pass along the loadHeaderOnly flag, which will instruct the layer to erase its current
            ' DIB as necessary.
            If (Not dstLayer.CreateNewLayerFromXML(retString, , loadHeaderOnly)) Then
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'If this is not a header-only operation, repeat the above steps, but for the layer DIB this time
        If (Not loadHeaderOnly) Then
        
            'How we extract this data varies by layer type.  Raster layers can skip the need for a temporary buffer, because we've
            ' already created a DIB with a built-in buffer for the pixel data.
            '
            'Other layer types (e.g. vector layers) are tiny so a temporary buffer does not matter; also, unlike raster buffers, we cannot
            ' easily predict the buffer size in advance, so we rely on pdPackage to do it for us
            Dim nodeLoadedSuccessfully As Boolean
            nodeLoadedSuccessfully = False
            
            'Image (raster) layers
            If dstLayer.IsLayerRaster Then
                
                'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                Dim tmpDIBPointer As Long, tmpDIBLength As Long
                dstLayer.layerDIB.SetInitialAlphaPremultiplicationState True
                dstLayer.layerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                nodeLoadedSuccessfully = pdiReader.GetNodeDataByID_UnsafeDstPointer(targetLayerID, False, tmpDIBPointer)
                    
            'Text and other vector layers
            ElseIf dstLayer.IsLayerVector Then
                
                If pdiReader.GetNodeDataByID(targetLayerID, False, retBytes, False, retSize) Then
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after pdPackager received Unicode compatibility.
                    retString = Space$(retSize \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
                    
                    'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                    If dstLayer.CreateVectorDataFromXML(retString) Then
                        nodeLoadedSuccessfully = True
                    Else
                        Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                    End If
                
                Else
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                End If
            
            'In the future, additional layer types can be handled here
            Else
                Debug.Print "WARNING! Unknown layer type exists in this PDI file: " & dstLayer.GetLayerType
            
            End If
                
            'If successful, notify the target layer that its DIB data has been changed; the layer will use this to regenerate various internal caches
            If nodeLoadedSuccessfully Then
                dstLayer.NotifyOfDestructiveChanges
                
            'Bytes could not be read, or alternately, checksums didn't match for the first node.
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
                
        End If
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadSingleLayerFromPDI = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file may be a legacy PDI format.
        PDDebug.LogAction "PDI v2 validation failed.  Attempting v1 load engine..."
        LoadSingleLayerFromPDI = LoadSingleLayerFromPDI_Legacy(pdiPath, dstLayer, targetLayerID, loadHeaderOnly)
    
    End If
    
    Exit Function
    
LoadLayerFromPDIFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadSingleLayerFromPDI = False
    Exit Function

End Function

'Load a single PhotoDemon layer from a standalone pdLayer file (which is really just a modified PDI file).
' At present, this function is only used internally by the Undo/Redo engine.  Its counterpart is SavePhotoDemonLayer in
' the Saving module; any changes there should be mirrored here.
Public Function LoadPhotoDemonLayer(ByVal pdiPath As String, ByRef dstLayer As pdLayer, Optional ByVal loadHeaderOnly As Boolean = False) As Boolean
    
    On Error GoTo LoadPDLayerFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.ReadPackageFromFile(pdiPath, PD_LAYER_IDENTIFIER) Then
    
        'Layer variants of PDI files contain a single node.  The layer's header is stored to the node's header chunk
        ' (in XML format, as expected).  The layer's DIB data is stored to the node's data chunk (in binary format, as expected).
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String, retSize As Long
        
        If pdiReader.GetNodeDataByIndex(0, True, retBytes, False, retSize) Then
        
            'Copy the received bytes into a string
            retString = Space$(retSize \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
            
            'Pass the string to the target layer, which will read the XML data and initialize itself accordingly.
            ' Note that we pass the loadHeaderOnly request to this function; if this is a header-only load, the target
            ' layer must retain its current DIB.  This functionality is used by PD's Undo/Redo engine.
            dstLayer.CreateNewLayerFromXML retString, , loadHeaderOnly
            
        'Bytes could not be read, or alternately, checksums didn't match.  (Note that checksums are currently disabled
        ' for this function, for performance reasons, but I'm leaving this check in case we someday decide to re-enable them.)
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'Unless a header-only load was requested, we will now repeat the steps above, but for layer-specific data
        ' (a raw DIB stream for raster layers, or an XML string for vector/text layers)
        If (Not loadHeaderOnly) Then
        
            'How we extract this data varies by layer type.  Raster layers can skip the need for a temporary buffer, because we've
            ' already created a DIB with a built-in buffer for the pixel data.
            '
            'Other layer types (e.g. vector layers) are tiny so a temporary buffer does not matter; also, unlike raster buffers, we cannot
            ' easily predict the buffer size in advance, so we rely on pdPackage to do it for us
            Dim nodeLoadedSuccessfully As Boolean
            nodeLoadedSuccessfully = False
            
            'Image (raster) layers
            If dstLayer.IsLayerRaster Then
                
                'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                Dim tmpDIBPointer As Long, tmpDIBLength As Long
                dstLayer.layerDIB.SetInitialAlphaPremultiplicationState True
                dstLayer.layerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                nodeLoadedSuccessfully = pdiReader.GetNodeDataByIndex_UnsafeDstPointer(0, False, tmpDIBPointer)
                
            'Text and other vector layers
            ElseIf dstLayer.IsLayerVector Then
                
                If pdiReader.GetNodeDataByIndex(0, False, retBytes, False, retSize) Then
                
                    'Convert the byte array to a Unicode string.
                    retString = Space$(retSize \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
                    
                    'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                    If dstLayer.CreateVectorDataFromXML(retString) Then
                        nodeLoadedSuccessfully = True
                    Else
                        Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                    End If
                    
                Else
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                End If
            
            'In the future, additional layer types can be handled here
            Else
                Debug.Print "WARNING! Unknown layer type exists in this PDI file: " & dstLayer.GetLayerType
            
            End If
                
            'If the load was successful, notify the target layer that its DIB data has been changed; the layer will use this to
            ' regenerate various internal caches.
            If nodeLoadedSuccessfully Then
                dstLayer.NotifyOfDestructiveChanges
                
            'Failure means package bytes could not be read, or alternately, checksums didn't match.  (Note that checksums are currently
            ' disabled for this function, for performance reasons, but I'm leaving this check in case we someday decide to re-enable them.)
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
        End If
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonLayer = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadPhotoDemonLayer = False
    
    End If
    
    Exit Function
    
LoadPDLayerFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadPhotoDemonLayer = False
    Exit Function

End Function

'Use GDI+ to load an image.  This does very minimal error checking (which is a no-no with GDI+) but because it's only a
' fallback when FreeImage can't be found, I'm postponing further debugging for now.
'Used for PNG and TIFF files if FreeImage cannot be located.
Public Function LoadGDIPlusImage(ByVal imagePath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByRef numOfPages As Long = 1) As Boolean
    
    LoadGDIPlusImage = False
    
    If GDI_Plus.GDIPlusLoadPicture(imagePath, dstDIB, dstImage, numOfPages) Then
        If (Not dstDIB Is Nothing) Then LoadGDIPlusImage = ((dstDIB.GetDIBWidth <> 0) And (dstDIB.GetDIBHeight <> 0))
    End If
    
End Function

Public Function IsFileSVGCandidate(ByRef imagePath As String) As Boolean
    IsFileSVGCandidate = Strings.StringsEqual(Right$(imagePath, 3), "svg", True)
    
    'Compressed SVG files are not currently supported.  (For them to work, we'd need to decompress to a temp file, which causes
    ' some messy interaction details with ExifTool - we'll deal with this in the future.)
    'If (Not IsFileSVGCandidate) Then IsFileSVGCandidate = CBool(StrComp(LCase$(Right$(imagePath, 4)), "svgz", vbBinaryCompare) = 0)
End Function

'SVG support is *experimental only*!  This function should not be enabled in production builds.
Public Function LoadSVG(ByVal imagePath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As Boolean
    
    On Error GoTo LoadSVGFail
    
    'In the future, we'll add meaningful heuristics, but for now, don't even attempt a load unless the file extension matches.
    If IsFileSVGCandidate(imagePath) Then
        
        PDDebug.LogAction "Waiting for SVG parsing to complete..."
        
        'Hang out while we wait for ExifTool to finish processing this image's metadata
        Do While (Not dstImage.ImgMetadata.HasMetadata)
            DoEvents
            If ExifTool.IsMetadataFinished Then
                dstImage.ImgMetadata.LoadAllMetadata ExifTool.RetrieveMetadataString, dstImage.imageID
            End If
        Loop
        
        'Retrieve the target SVG's width and height
        Dim svgWidth As String, svgHeight As String
        svgWidth = dstImage.ImgMetadata.GetTagValue("SVG:ImageWidth", vbBinaryCompare, True)
        svgHeight = dstImage.ImgMetadata.GetTagValue("SVG:ImageHeight", vbBinaryCompare, True)
        
        'If there's a viewbox, grab it too
        Dim svgHasViewbox As Boolean
        svgHasViewbox = dstImage.ImgMetadata.DoesTagExistFullName("SVG:ViewBox", , vbBinaryCompare)
        
        Dim svgWidthL As Long, svgHeightL As Long
        If (Len(svgWidth) <> 0) Then
            
            'Check for sizes defined as percentages (possible when a view box is specified)
            If (InStr(1, svgWidth, "%", vbBinaryCompare) <> 0) Then
                'TODO: grab width/height from viewbox
                svgWidthL = 100
            Else
                Debug.Print "HERE 2", InStr(1, svgWidth, "%", vbBinaryCompare)
                svgWidthL = CLng(svgWidth)
            End If
            
        Else
            svgWidthL = 100
        End If
        
        If (Len(svgHeight) <> 0) Then
            If (InStr(1, svgHeight, "%", vbBinaryCompare) <> 0) Then
                svgHeightL = 100    'TODO: grab height/height from viewbox
            Else
                svgHeightL = CLng(svgHeight)
            End If
            
        Else
            svgHeightL = 100
        End If
        
        LoadSVG = True
        dstDIB.CreateBlank svgWidthL, svgHeightL, 32, vbWhite, 255
        dstDIB.SetInitialAlphaPremultiplicationState True
        
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
'New as of 11 July '14 is the ability to specify a custom layer destination, for layer-relevant load operations.  If this value is NOTHING,
' the function will automatically load the data to the relevant layer in the parent pdImage object.  If this layer is supplied, however,
' the supplied layer reference will be used instead.
Public Sub LoadUndo(ByVal undoFile As String, ByVal undoTypeOfFile As Long, ByVal undoTypeOfAction As PD_UndoType, Optional ByVal targetLayerID As Long = -1, Optional ByVal suspendRedraw As Boolean = False, Optional ByRef customLayerDestination As pdLayer = Nothing)
    
    'Certain load functions require access to a DIB, so declare a generic one in advance
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'If selection data was loaded as part of this diff, this value will be set to TRUE.  We check it at the end of
    ' the load function, and activate various selection-related items as necessary.
    Dim selectionDataLoaded As Boolean
    selectionDataLoaded = False
    
    'Regardless of outcome, notify the parent image of this change
    PDImages.GetActiveImage.NotifyImageChanged undoTypeOfAction, targetLayerID
    
    'Depending on the Undo data requested, we may end up loading one or more diff files at this location
    Select Case undoTypeOfAction
    
        'UNDO_EVERYTHING: a full copy of both the pdImage stack and all selection data is wanted
        Case UNDO_Everything
            ImageImporter.LoadPhotoDemonImage undoFile, tmpDIB, PDImages.GetActiveImage(), True
            PDImages.GetActiveImage.MainSelection.ReadSelectionFromFile undoFile & ".selection"
            selectionDataLoaded = True
            
        'UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE: a full copy of the pdImage stack is wanted
        '             Because the underlying file data must be of type UNDO_EVERYTHING or UNDO_IMAGE/_VECTORSAFE, we
        '             don't have to do any special processing to the file - just load the whole damn thing.
        Case UNDO_Image, UNDO_Image_VectorSafe
            ImageImporter.LoadPhotoDemonImage undoFile, tmpDIB, PDImages.GetActiveImage(), True
            
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
            ImageImporter.LoadPhotoDemonImageHeaderOnly undoFile, PDImages.GetActiveImage()
            
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
                    ImageImporter.LoadPhotoDemonLayer undoFile & ".layer", customLayerDestination, False
            
                'The underlying save file is a full pdImage stack.  Extract only the relevant layer data from the stack.
                Case UNDO_Everything, UNDO_Image, UNDO_Image_VectorSafe
                    ImageImporter.LoadSingleLayerFromPDI undoFile, customLayerDestination, targetLayerID, False
                
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
                    ImageImporter.LoadPhotoDemonLayer undoFile & ".layer", PDImages.GetActiveImage.GetLayerByID(targetLayerID), True
            
                'The underlying save file is a full pdImage stack.  Extract only the relevant layer data from the stack.
                Case UNDO_Everything, UNDO_Image, UNDO_Image_VectorSafe, UNDO_ImageHeader
                    ImageImporter.LoadSingleLayerFromPDI undoFile, PDImages.GetActiveImage.GetLayerByID(targetLayerID), targetLayerID, True
                
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
            ImageImporter.LoadPhotoDemonImage undoFile, tmpDIB, PDImages.GetActiveImage(), True
            
        
    End Select
    
    'If a selection was loaded, activate all selection-related stuff now
    If selectionDataLoaded Then
    
        'Activate the selection as necessary
        PDImages.GetActiveImage.SetSelectionActive PDImages.GetActiveImage.MainSelection.IsLockedIn
        
        'Synchronize the text boxes as necessary
        Selections.SyncTextToCurrentSelection PDImages.GetActiveImageID()
    
    End If
    
    'If a selection is active, request a redraw of the selection mask before rendering the image to the screen.  (If we are
    ' "undoing" an action that changed the image's size, the selection mask will be out of date.  Thus we need to re-render
    ' it before rendering the image or OOB errors may occur.)
    If PDImages.GetActiveImage.IsSelectionActive Then PDImages.GetActiveImage.MainSelection.RequestNewMask
        
    'Render the image to the screen, if requested
    If (Not suspendRedraw) Then ViewportEngine.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

'Load a raw pdDIB file dump into the destination image and DIB.  (Note that pdDIB may have applied zLib compression during the save,
' depending on the parameters it was passed, so it is possible for this function to fail if zLib goes missing.)
Public Function LoadRawImageBuffer(ByVal imagePath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As Boolean

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
Public Function CascadeLoadGenericImage(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef freeImage_Return As PD_OPERATION_OUTCOME, ByRef decoderUsed As PD_ImageDecoder, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long) As Boolean
    
    CascadeLoadGenericImage = False
    
    'Before jumping out to a 3rd-party library, check for any image formats that we must decode using internal plugins.
    
    'SVG support is just experimental at present!
    CascadeLoadGenericImage = ImageImporter.LoadSVG(srcFile, dstDIB, dstImage)
    If CascadeLoadGenericImage Then
        decoderUsed = id_SVGParser
        dstImage.SetOriginalFileFormat PDIF_SVG
        dstImage.SetDPI 96, 96
        dstImage.SetOriginalColorDepth 32
        dstImage.SetOriginalGrayscale False
        dstImage.SetOriginalAlpha True
    End If
    
    'PD's internal PNG parser is preferred for all PNG images.  For backwards compatibility reasons, it does *not* rely
    ' on the .png extension.  (Instead, it will manually verify the PNG signature, then work from there.)
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_PNG Then
        CascadeLoadGenericImage = LoadPNGOurselves(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_PNGParser
            dstImage.SetOriginalFileFormat PDIF_PNG
        End If
    End If
    
    'OpenRaster support was added in v7.2.  OpenRaster is similar to ODF, basically a .zip wrapper around an XML file
    ' and a bunch of PNGs - easy enough to support!
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_ORA Then
        CascadeLoadGenericImage = LoadOpenRaster(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_ORAParser
            dstImage.SetOriginalFileFormat PDIF_ORA
        End If
    End If
    
    'PD's internal PSD decoder is experimental as of v7.2.  If it fails (likely), we still fall back
    ' to FreeImage's generic PSD support in a subsequent step.
    If (Not CascadeLoadGenericImage) And USE_INTERNAL_PARSER_PSD Then
        CascadeLoadGenericImage = LoadPSD(srcFile, dstImage, dstDIB)
        If CascadeLoadGenericImage Then
            decoderUsed = id_PSDParser
            dstImage.SetOriginalFileFormat PDIF_PSD
        End If
    End If
    
    'If our various internal engines passed on the image, we now want to attempt either FreeImage or GDI+.
    ' (Pre v7.2, we *always* tried FreeImage first, but as time goes by, I realize the library is prone to a
    ' lot of bugs.  It also suffers performance-wise compared to GDI+.  As such, I am now more selective about
    ' which library gets used first.)
    If (Not CascadeLoadGenericImage) Then
    
        'FreeImage's TIFF support (via libTIFF?) is wonky.  It's prone to bad crashes and inexplicable memory
        ' issues (including allocation failures on normal-sized images), so for TIFFs we want to try GDI+ before
        ' trying FreeImage.  (PD's GDI+ image loader was heavily restructured in v7.2 to support things like
        ' multi-page import, so this strategy wasn't viable until then.)
        Dim tryGDIPlusFirst As Boolean
        tryGDIPlusFirst = Strings.StringsEqual(Files.FileGetExtension(srcFile), "tif", True) Or Strings.StringsEqual(Files.FileGetExtension(srcFile), "tiff", True)
        
        'On modern Windows builds (8+) FreeImage is markedly slower than GDI+ at loading JPEG images, so let's also default
        ' to GDI+ for JPEGs.
        If (Not tryGDIPlusFirst) Then tryGDIPlusFirst = Strings.StringsEqual(Files.FileGetExtension(srcFile), "jpg", True) Or Strings.StringsEqual(Files.FileGetExtension(srcFile), "jpeg", True)
        
        If tryGDIPlusFirst Then
            CascadeLoadGenericImage = AttemptGDIPlusLoad(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
            freeImage_Return = PD_FAILURE_GENERIC
            If (Not CascadeLoadGenericImage) And ImageFormats.IsFreeImageEnabled() Then CascadeLoadGenericImage = AttemptFreeImageLoad(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
            
        'For other formats, let FreeImage have a go at it, and we'll try GDI+ if it fails
        Else
            freeImage_Return = PD_FAILURE_GENERIC
            If (Not CascadeLoadGenericImage) And ImageFormats.IsFreeImageEnabled() Then CascadeLoadGenericImage = AttemptFreeImageLoad(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
            If (Not CascadeLoadGenericImage) And (freeImage_Return <> PD_FAILURE_USER_CANCELED) Then CascadeLoadGenericImage = AttemptGDIPlusLoad(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
        End If
        
    End If
    
End Function

Private Function AttemptFreeImageLoad(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef freeImage_Return As PD_OPERATION_OUTCOME, ByRef decoderUsed As PD_ImageDecoder, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long) As Boolean
    
    PDDebug.LogAction "Attempting to load via FreeImage..."
    
    'Start by seeing if the image file contains multiple pages.  If it does, we will load each page as a separate layer.
    ' TODO: preferences or prompt for how to handle such files??
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
        If (dstDIB.GetOriginalFormat = PDIF_PNG) And (dstDIB.GetBackgroundColor <> -1) Then dstImage.ImgStorage.AddEntry "pngBackgroundColor", dstDIB.GetBackgroundColor
        
    End If
        
End Function

Private Function AttemptGDIPlusLoad(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef freeImage_Return As PD_OPERATION_OUTCOME, ByRef decoderUsed As PD_ImageDecoder, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long) As Boolean

    PDDebug.LogAction "Attempting to load via GDI+..."
    AttemptGDIPlusLoad = LoadGDIPlusImage(srcFile, dstDIB, dstImage, numOfPages)
    
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
            CascadeLoadInternalImage = LoadPhotoDemonImage(srcFile, dstDIB, dstImage)
            
            dstImage.SetOriginalFileFormat PDIF_PDI
            dstImage.SetOriginalColorDepth 32
            dstImage.NotifyImageChanged UNDO_Everything
            decoderUsed = id_Internal
            
        'TMPDIB files are raw pdDIB objects dumped directly to file.  In some cases, this is faster and easier for PD than wrapping
        ' the pdDIB object inside a pdPackage layer (e.g. during clipboard interactions, since we start with a raw pdDIB object
        ' after selections and such are applied to the base layer/image, so we may as well just use the raw pdDIB data we've cached).
        Case PDIF_RAWBUFFER
            
            'These raw pdDIB objects may require zLib for parsing (compression is optional), so it is possible for the load function
            ' to fail if zLib goes missing.
            CascadeLoadInternalImage = LoadRawImageBuffer(srcFile, dstDIB, dstImage)
            
            dstImage.SetOriginalFileFormat PDIF_UNKNOWN
            dstImage.SetOriginalColorDepth 32
            dstImage.NotifyImageChanged UNDO_Everything
            decoderUsed = id_Internal
            
        'Straight TMP files are internal files (BMP, typically) used by PhotoDemon.  As ridiculous as it sounds, we must
        ' default to the generic load engine list, as the format of a TMP file is not guaranteed in advance.  Because of this,
        ' we can rely on the generic load engine to properly set things like "original color depth".
        '
        '(TODO: settle on a single tmp file format, so we don't have to play this game??)
        Case PDIF_TMPFILE
            CascadeLoadInternalImage = ImageImporter.CascadeLoadGenericImage(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
            dstImage.SetOriginalFileFormat PDIF_UNKNOWN
            
    End Select
    
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

Private Function LoadPNGOurselves(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    LoadPNGOurselves = False
    
    'pdPNG handles all the dirty work for us
    Dim cPNG As pdPNG
    Set cPNG = New pdPNG
    LoadPNGOurselves = (cPNG.LoadPNG_Simple(srcFile, dstImage, dstDIB, False) <= png_Warning)
    
    If LoadPNGOurselves Then
    
        'If we've experienced one or more warnings during the load process, dump them out to the debug file.
        If (cPNG.Warnings_GetCount() > 0) Then cPNG.Warnings_DumpToDebugger
        
        'Relay any useful state information to the destination image object; this information may be useful
        ' if/when the user saves the image.
        dstImage.SetOriginalFileFormat PDIF_PNG
        dstImage.NotifyImageChanged UNDO_Everything
        
        dstImage.SetOriginalColorDepth cPNG.GetBytesPerPixel()
        dstImage.SetOriginalGrayscale (cPNG.GetColorType = png_Greyscale) Or (cPNG.GetColorType = png_GreyscaleAlpha)
        dstImage.SetOriginalAlpha cPNG.HasAlpha()
        If cPNG.HasChunk("bKGD") Then dstImage.ImgStorage.AddEntry "pngBackgroundColor", cPNG.GetBackgroundColor()
        
        'Because color-management has already been handled (if applicable), this is a great time to premultiply alpha
        dstDIB.SetAlphaPremultiplication True
        
    End If

End Function

'Use PD's internal PSD parser to attempt to load a target PSD file.
Private Function LoadPSD(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    LoadPSD = False
    
    'pdpsd handles all the dirty work for us
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
        dstImage.SetOriginalColorDepth cPSD.GetBytesPerPixel()
        dstImage.SetOriginalGrayscale cPSD.IsGrayscale()
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
    ViewportEngine.DisableRendering
    FormMain.MainCanvas(0).SetScrollVisibility pdo_Both, True
    FormMain.MainCanvas(0).SetScrollValue pdo_Both, 0
    
    'Reflow any image-window-specific display elements on the actual image form (status bar, rulers, etc).
    ' Importantly, this also shows/hides the image tabstrip that's available when multiple images are loaded.
    FormMain.MainCanvas(0).AlignCanvasView
    
    'Fit the current image to the active viewport
    FitImageToViewport
    ViewportEngine.EnableRendering
    
    'Notify the UI manager that it now has one more image to deal with
    If (Macros.GetMacroStatus <> MacroBATCH) Then Interface.NotifyImageAdded srcImage.imageID
                            
    'Add this file to the MRU list (unless specifically told not to)
    If addToRecentFiles And (Macros.GetMacroStatus <> MacroBATCH) Then g_RecentFiles.AddFileToList srcFile, srcImage
    
End Sub

'Legacy import functions for old PDI versions are found below.  These functions are no longer maintained; use at your own risk.
Private Function LoadPDI_Legacy(ByVal pdiPath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByVal sourceIsUndoFile As Boolean = False) As Boolean

    PDDebug.LogAction "Legacy PDI file identified.  Starting pdPackage decompression..."
    
    On Error GoTo LoadPDIFail
    
    'PDI files require a parent pdImage container
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    
    'First things first: create a legacy pdPackage instance.  It will handle all the messy business of extracting individual
    ' data bits from the source file.
    Dim pdiReader As pdPackagerLegacy
    Set pdiReader = New pdPackagerLegacy
    pdiReader.Init_ZLib vbNullString, True, PluginManager.IsPluginCurrentlyEnabled(CCP_zLib)
    
    'Load the file into the pdPackagerLegacy instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.ReadPackageFromFile(pdiPath, PD_IMAGE_IDENTIFIER) Then
    
        PDDebug.LogAction "pdPackage successfully read and initialized.  Starting package parsing..."
        
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.GetNodeDataByIndex(0, True, retBytes, sourceIsUndoFile) Then
            
            PDDebug.LogAction "Initial PDI node retrieved.  Initializing corresponding pdImage object..."
            
            'Copy the received bytes into a string
            If pdiReader.GetPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                retString = Space$((UBound(retBytes) + 1) \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            Else
                retString = StrConv(retBytes, vbUnicode)
            End If
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.ReadExternalData retString, True, sourceIsUndoFile
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        PDDebug.LogAction "pdImage created successfully.  Moving on to individual layers..."
        
        'With the main pdImage now assembled, the next task is to populate all layers with two pieces of information:
        ' 1) The layer header, which contains stuff like layer name, opacity, blend mode, etc
        ' 2) Layer-specific information, which varies by layer type.  For DIBs, this will be a raw stream of bytes
        '    containing the layer DIB's raster data.  For text or other vector layers, this is an XML stream containing
        '    whatever information is necessary to construct the layer from scratch.
        
        Dim i As Long
        For i = 0 To dstImage.GetNumOfLayers - 1
        
            PDDebug.LogAction "Retrieving layer header " & i & "..."
            
            'First, retrieve the layer's header
            If pdiReader.GetNodeDataByIndex(i + 1, True, retBytes, sourceIsUndoFile) Then
            
                'Copy the received bytes into a string
                If pdiReader.GetPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                    retString = Space$((UBound(retBytes) + 1) \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
                Else
                    retString = StrConv(retBytes, vbUnicode)
                End If
                
                'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                If Not dstImage.GetLayerByIndex(i).CreateNewLayerFromXML(retString) Then
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
                End If
                
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
            'How we extract the rest of the layer's data varies by layer type.  Raster layers can skip the need for a temporary buffer,
            ' because we've already created a DIB with a built-in buffer for the pixel data.
            '
            'Other layer types (e.g. vector layers) are tiny so a temporary buffer does not matter; also, unlike raster buffers, we cannot
            ' easily predict the buffer size in advance, so we rely on pdPackage to do it for us
            Dim nodeLoadedSuccessfully As Boolean
            nodeLoadedSuccessfully = False
            
            'Image (raster) layers
            If dstImage.GetLayerByIndex(i).IsLayerRaster Then
                
                PDDebug.LogAction "Raster layer identified.  Retrieving pixel bits..."
                
                'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                Dim tmpDIBPointer As Long, tmpDIBLength As Long
                dstImage.GetLayerByIndex(i).layerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                'At present, all pdPackage layers will contain premultiplied alpha, so force the corresponding state now
                dstImage.GetLayerByIndex(i).layerDIB.SetInitialAlphaPremultiplicationState True
                
                nodeLoadedSuccessfully = pdiReader.GetNodeDataByIndex_UnsafeDstPointer(i + 1, False, tmpDIBPointer, sourceIsUndoFile)
            
            'Text and other vector layers
            ElseIf dstImage.GetLayerByIndex(i).IsLayerVector Then
                
                PDDebug.LogAction "Vector layer identified.  Retrieving layer XML..."
                
                If pdiReader.GetNodeDataByIndex(i + 1, False, retBytes, sourceIsUndoFile) Then
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after pdPackagerLegacy gained full Unicode compatibility.
                    retString = Space$((UBound(retBytes) + 1) \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
                    
                    'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                    If dstImage.GetLayerByIndex(i).CreateVectorDataFromXML(retString) Then
                        nodeLoadedSuccessfully = True
                    Else
                        Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                    End If
                    
                Else
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                End If
                    
            'In the future, additional layer types can be handled here
            Else
                Debug.Print "WARNING! Unknown layer type exists in this PDI file: " & dstImage.GetLayerByIndex(i).GetLayerType
            
            End If
            
            'If successful, notify the parent of the change
            If nodeLoadedSuccessfully Then
                dstImage.NotifyImageChanged UNDO_Layer, i
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
        
        Next i
        
        PDDebug.LogAction "All layers loaded.  Looking for remaining non-essential PDI data..."
        
        'Finally, check to see if the PDI image has a metadata entry.  If it does, load that data now.
        If pdiReader.GetNodeDataByName("pdMetadata_Raw", True, retBytes, sourceIsUndoFile) Then
        
            PDDebug.LogAction "Raw metadata chunk found.  Retrieving now..."
            
            'Copy the received bytes into a string
            If pdiReader.GetPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                retString = Space$((UBound(retBytes) + 1) \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            Else
                retString = StrConv(retBytes, vbUnicode)
            End If
            
            'Pass the string to the parent image's metadata handler, which will parse the XML data and prepare a matching
            ' internal metadata struct.
            If Not dstImage.ImgMetadata.LoadAllMetadata(retString, dstImage.imageID) Then
                
                'For invalid metadata, do not reject the rest of the PDI file.  Instead, just warn the user and carry on.
                Debug.Print "PDI Metadata Node rejected by metadata parser."
                
            End If
        
        End If
        
        '(As of v7.0, a serialized copy of the image's metadata is also stored.  This copy contains all user edits
        ' and other changes.)
        If pdiReader.GetNodeDataByName("pdMetadata_Raw", False, retBytes, sourceIsUndoFile) Then
        
            PDDebug.LogAction "Serialized metadata chunk found.  Retrieving now..."
            
            'Copy the received bytes into a string
            If pdiReader.GetPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                retString = Space$((UBound(retBytes) + 1) \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            Else
                retString = StrConv(retBytes, vbUnicode)
            End If
            
            'Pass the string to the parent image's metadata handler, which will parse the XML data and prepare a matching
            ' internal metadata struct.
            dstImage.ImgMetadata.RecreateFromSerializedXMLData retString
        
        End If
        
        PDDebug.LogAction "PDI parsing complete.  Returning control to main image loader..."
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank 16, 16, 32, 0
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPDI_Legacy = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadPDI_Legacy = False
    
    End If
    
    Exit Function
    
LoadPDIFail:
    
    PDDebug.LogAction "WARNING!  LoadPDIFail error routine reached.  Checking for known error states..."
    
    'Before falling back to a generic error message, check for a couple known problem states.
    
    'Case 1: zLib is required for this file, but the user doesn't have the zLib plugin
    Dim cmpMissing As Boolean
    cmpMissing = pdiReader.GetPackageFlag(PDP_HF2_ZlibRequired, PDP_LOCATION_ANY) And (Not PluginManager.IsPluginCurrentlyEnabled(CCP_zLib))
    
    If cmpMissing Then
        PDMsgBox "The PDI file ""%1"" contains compressed data, but the required plugin is missing or disabled.", vbCritical Or vbOKOnly, "Compression plugin missing", Files.FileGetName(pdiPath)
        Exit Function
    End If
    
    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else
            Message "An error has occurred (#%1 - %2).  PDI load abandoned.", Err.Number, Err.Description
        
    End Select
    
    LoadPDI_Legacy = False
    Exit Function

End Function

'Load just the layer stack from a standard PDI file, and non-destructively align our current layer stack to match.
' At present, this function is only used internally by the Undo/Redo engine.
Private Function LoadPhotoDemonImageHeaderOnly_Legacy(ByVal pdiPath As String, ByRef dstImage As pdImage) As Boolean
    
    On Error GoTo LoadPDIHeaderFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackagerLegacy
    Set pdiReader = New pdPackagerLegacy
    pdiReader.Init_ZLib vbNullString, True, PluginManager.IsPluginCurrentlyEnabled(CCP_zLib)
    
    'Load the file into the pdPackagerLegacy instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.ReadPackageFromFile(pdiPath, PD_IMAGE_IDENTIFIER) Then
    
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.GetNodeDataByIndex(0, True, retBytes, True) Then
        
            'Copy the received bytes into a string
            If pdiReader.GetPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                retString = Space$((UBound(retBytes) + 1) \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            Else
                retString = StrConv(retBytes, vbUnicode)
            End If
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.ReadExternalData retString, True, True, True
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'With the main pdImage now assembled, the next task is to populate all layer headers.  This is a bit more
        ' confusing than a regular PDI load, because we have to maintain existing layer DIB data (ugh!).
        ' So basically, we must:
        ' 1) Extract each layer header from file, in turn
        ' 2) See if the current pdImage copy of this layer is in the proper position in the layer stack; if it isn't,
        '    move it into the location specified by the PDI file.
        ' 3) Ask the layer to non-destructively overwrite its header with the header from the PDI file (e.g. don't
        '    touch its DIB or vector-specific contents).
        
        Dim layerNodeName As String, layerNodeID As Long, layerNodeType As Long
        
        Dim i As Long
        For i = 0 To dstImage.GetNumOfLayers - 1
        
            'Before doing anything else, retrieve the ID of the node at this position.  (Retrieve the rest of the node
            ' header too, although we don't actually have a use for those values at present.)
            pdiReader.GetNodeInfo i + 1, layerNodeName, layerNodeID, layerNodeType
            
            'We now know what layer ID is supposed to appear at this position in the layer stack.  If that layer ID
            ' is *not* in its proper position, move it now.
            If dstImage.GetLayerIndexFromID(layerNodeID) <> i Then dstImage.SwapTwoLayers dstImage.GetLayerIndexFromID(layerNodeID), i
            
            'Now that the node is in place, we can retrieve its header.
            If pdiReader.GetNodeDataByIndex(i + 1, True, retBytes, True) Then
            
                'Copy the received bytes into a string
                If pdiReader.GetPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                    retString = Space$((UBound(retBytes) + 1) \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
                Else
                    retString = StrConv(retBytes, vbUnicode)
                End If
                
                'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                If Not dstImage.GetLayerByIndex(i).CreateNewLayerFromXML(retString, , True) Then
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
                End If
                
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
            'Normally we would load the layer's DIB data here, but we don't care about that when loading just the headers!
            ' Continue to the next layer.
        
        Next i
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonImageHeaderOnly_Legacy = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadPhotoDemonImageHeaderOnly_Legacy = False
    
    End If
    
    Exit Function
    
LoadPDIHeaderFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadPhotoDemonImageHeaderOnly_Legacy = False
    Exit Function

End Function

'Load a single layer from a standard PDI file.
' At present, this function is only used internally by the Undo/Redo engine.  If the nearest diff to a layer-specific change is a
' full pdImage stack, this function is used to extract only the relevant layer (or layer header) from the PDI file.
Private Function LoadSingleLayerFromPDI_Legacy(ByVal pdiPath As String, ByRef dstLayer As pdLayer, ByVal targetLayerID As Long, Optional ByVal loadHeaderOnly As Boolean = False) As Boolean
    
    On Error GoTo LoadLayerFromPDIFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackagerLegacy
    Set pdiReader = New pdPackagerLegacy
    pdiReader.Init_ZLib vbNullString, True, PluginManager.IsPluginCurrentlyEnabled(CCP_zLib)
    
    'Load the file into the pdPackagerLegacy instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.ReadPackageFromFile(pdiPath, PD_IMAGE_IDENTIFIER) Then
    
        'PDI files all follow a standard format: a pdImage node at the top, which contains the full pdImage header,
        ' followed by individual nodes for each layer.  Layers are stored in stack order, which makes it very fast and easy
        ' to reconstruct the layer stack.
        
        'Unfortunately, stack order is not helpful in this function, because the target layer's position may have changed
        ' since the time this pdImage file was created.  To work around that, we must located the layer using its cardinal
        ' ID value, which is helpfully stored as the node ID parameter for a given layer node.
        
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.GetNodeDataByID(targetLayerID, True, retBytes, True) Then
        
            'Copy the received bytes into a string
            retString = Space$((UBound(retBytes) + 1) \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            
            'Pass the string to the target layer, which will read the XML data and initialize itself accordingly.
            ' Note that we also pass along the loadHeaderOnly flag, which will instruct the layer to erase its current
            ' DIB as necessary.
            If Not dstLayer.CreateNewLayerFromXML(retString, , loadHeaderOnly) Then
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'If this is not a header-only operation, repeat the above steps, but for the layer DIB this time
        If Not loadHeaderOnly Then
        
            'How we extract this data varies by layer type.  Raster layers can skip the need for a temporary buffer, because we've
            ' already created a DIB with a built-in buffer for the pixel data.
            '
            'Other layer types (e.g. vector layers) are tiny so a temporary buffer does not matter; also, unlike raster buffers, we cannot
            ' easily predict the buffer size in advance, so we rely on pdPackage to do it for us
            Dim nodeLoadedSuccessfully As Boolean
            nodeLoadedSuccessfully = False
            
            'Image (raster) layers
            If dstLayer.IsLayerRaster Then
                
                'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                Dim tmpDIBPointer As Long, tmpDIBLength As Long
                dstLayer.layerDIB.SetInitialAlphaPremultiplicationState True
                dstLayer.layerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                nodeLoadedSuccessfully = pdiReader.GetNodeDataByID_UnsafeDstPointer(targetLayerID, False, tmpDIBPointer, True)
                    
            'Text and other vector layers
            ElseIf dstLayer.IsLayerVector Then
                
                If pdiReader.GetNodeDataByID(targetLayerID, False, retBytes, True) Then
                
                    'Convert the byte array to a Unicode string
                    retString = Space$((UBound(retBytes) + 1) \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
                    
                    'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                    If dstLayer.CreateVectorDataFromXML(retString) Then
                        nodeLoadedSuccessfully = True
                    Else
                        Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                    End If
                
                Else
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                End If
            
            'In the future, additional layer types can be handled here
            Else
                Debug.Print "WARNING! Unknown layer type exists in this PDI file: " & dstLayer.GetLayerType
            
            End If
                
            'If successful, notify the target layer that its DIB data has been changed; the layer will use this to regenerate various internal caches
            If nodeLoadedSuccessfully Then
                dstLayer.NotifyOfDestructiveChanges
                
            'Bytes could not be read, or alternately, checksums didn't match for the first node.
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
                
        End If
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadSingleLayerFromPDI_Legacy = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadSingleLayerFromPDI_Legacy = False
    
    End If
    
    Exit Function
    
LoadLayerFromPDIFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadSingleLayerFromPDI_Legacy = False
    Exit Function

End Function
